require 'ox'
require 'find'
require 'matrix'
require 'fileutils'
require 'zip'

class String
  # Returns true if the given String represents a numeric value 
  def numeric?
    Float(self) != nil rescue false
  end
end

class Matrix
  # Set the cell i,j to x, where i=row, and j=column.
  def []=(i, j, x)
    @rows[i][j]=x
  end
end

class Fixnum
  # Returns the column name corresponding to the given number. e.g: 1.col_name => 'B'
  def col_name
    val=self/26
    (val > 0 ? (val - 1).col_name : "") + (self % 26 + 65).chr
  end
end

# The namespace for all classes and modules included on Oxcelix.
module Oxcelix
  # The Cellhelper module defines some methods useful to manipulate Cell objects
  module Cellhelper
    # Set the excel cell name (eg: 'A2')
    def r(val); @xlcoords = val; end;
    # Set cell type    
    def t(val); @type = val; end;
    # Set cell value
    def v(val); @value = val; end;
    # Set cell style (number format and style) 
    def s(val); @style = val; end;
    
    # When called without parameters, returns the x coordinate of the calling cell object based on the value of #@xlcoords
    # If a parameter is given, #x will return the x coordinate corresponding to the parameter
    # E.g.: <tt>('B3') # => 1</tt>
    def x(coord=nil)
      if coord.nil?
        coord = @xlcoords
      end
      ('A'..(coord.scan(/\p{Alpha}+|\p{Digit}+/u)[0])).to_a.length-1
    end

    # When called without parameters, returns the y coordinate of the calling cell object based on the value of #@xlcoords
    # If a parameter is given, #y will return the y coordinate corresponding to the parameter
    # E.g.: <tt>('B3') # => 2</tt>
    def y(coord=nil)
      if coord.nil?
        coord = @xlcoords
      end
      coord.scan(/\p{Alpha}+|\p{Digit}+/u)[1].to_i-1
    end
  end

  # Simple class representing an excel cell.
  class Cell
    attr_accessor :xlcoords, :type, :value, :comment, :style
    include Cellhelper
  end

  ##
  # The Xlsheet class is a SAX parser based on the Ox library. It parses a
  # SpreadsheetML (AKA Office Open XML) formatted XML file and returns an array
  # of Cell objects (#cellarray).
  #
  # Xlsheet will omit the following:
  # * empty cells
  # * cells containing formulas
  #
  # Only non-empty cells of merged groupsbe added to #cellarray. A separate array
  # (#mergedcells) is reserved for merging.
  class Xlsheet < ::Ox::Sax
    attr_accessor :xmlstack, :lastopennode, :mergedcells, :cellarray, :cell
    def initialize()
      @xmlstack = []
      @lastopennode = nil
      @mergedcells = []
      @cellarray = []
      @cell = Cell.new
    end

    # Save SAX state-machine state to #xmlstack if and only if the processed
    # element is a :c (column) or a :mergeCell (merged cell)
    def start_element(name)
      if name == :c || name == :mergeCell
        @xmlstack << name
      end
    end
    
    # Step back in the stack (#xmlstack.pop), clear actual cell information
    def end_element(name)
      @xmlstack.pop
      if name == :c || name == :mergeCell
        @cell=Cell.new
      end
    end
    
    # Set cell value, style, etc. This will only happen if the cell has an
    # actual value AND the parser's state is :c.
    # If the state is :mergeCell AND the actual attribute name is :ref the
    # attribute will be added to the merged cells array
    def attr(name, str)
      if @xmlstack.last == :c # && name != :s
        @cell.send name, str if @cell.respond_to?(name)
      elsif xmlstack.last == :mergeCell && name == :ref
        @mergedcells << str
      end
    end

    # String type cell content is parsed here. String interpolation using the
    # sharedStrings.xml file is done in the #Sharedstrings class
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric?
          @cell.v str
          @cellarray << @cell
        end
        cell=Cell.new
      end
    end
  end

  # Ox based SAX parser which pushes shared strings (taken from the sharedString.xml file) to an array 
  # These strings will replace the references in the cells (interpolation). 
  class Sharedstrings < ::Ox::Sax
    attr_accessor :stringarray
    def initialize
      @stringarray=[]
    end

    # Push the comment string into @stringarray
    def text(str)
      @stringarray << str
    end
  end

  # The Comments class is a parser which builds an array of comments
  class Comments < ::Ox::Sax
    attr_accessor :commarray, :comment
    def initialize
      @commarray=[]
      @comment={}
    end
    
    # Push Cell comment hash (comment + reference) to @commarray
    def text(str)
      @comment[:comment]=str.gsub('&#10;', '')
      @commarray << @comment
      @comment = Hash.new
    end
    
    # Returns reference
    def attr(name, str)
      if name == :ref
        @comment[:ref]=str
      end
    end
  end

  # The Sheet class represents an excel sheet.
  class Sheet
    attr_accessor :name, :sheetId, :relationId, :data
    def initialize; @data=[]; end
  end

  # Helper methods for the Workbook class
  module Workbookhelper
    # returns a sheet based on its name
    def [] (sheetname=String)
      @sheets.select{|s| s.name==sheetname}[0]
    end
  end

  # The Workbook class will open the excel file, and convert it to a collection of
  # Matrix objects
  class Workbook
    include Cellhelper
    include Workbookhelper
    attr_accessor :sheetbase, :sheets, :sharedstrings
    ##
    # Create a new Workbook object.
    #
    # filename is the name of the Excel 2007/2010 file to be opened (xlsx)
    #
    # options is a collection of options that can be passed to Workbook.
    # Options may include:
    # * :copymerge (=> true/false) - Copy and repeat the content of the merged cells into the whole group, e.g. 
    # the group of three merged cells <tt>|   a   |</tt> will become <tt>|a|a|a|</tt>
    # * :include (Ary) - an array of sheet names to be included
    # * :exclude (Ary) - an array of sheet names not to be processed
    #
    # The excel file is first getting unzipped, then the workbook.xml file gets
    # processed. This file stores sheet metadata, which will be filtered (by including
    # and excluding sheets from further processing)
    #
    # The next stage is building sheets.
    # This includes:
    # * Parsing the XML files representing the sheets
    # * Interpolation of the shared strings
    # * adding comments to the cells
    # * Converting each sheet to a Matrix object
    def initialize(filename, options={})
      @destination = Dir.pwd+'/tmp'
      FileUtils.mkdir(@destination)
      Zip::File.open(filename){ |zip_file|
        zip_file.each{ |f| 
          f_path=File.join(@destination, f.name)
          FileUtils.mkdir_p(File.dirname(f_path))
          zip_file.extract(f, f_path) unless File.exists?(f_path)
        }
      }
      @sheets=[]
      @sheetbase={}
      @a=Ox::load_file(@destination+'/xl/workbook.xml')
      
      sheetdata(options); commentsrel; shstrings;
      @sheets.each do |x|
        sname="sheet#{x[:sheetId]}"
        @sheet = Xlsheet.new()
        File.open(@destination+"/xl/worksheets/#{sname}.xml", 'r') do |f|
          Ox.sax_parse(@sheet, f)
        end
        comments=
        mkcomments(x[:comments])
        @sheet.cellarray.each do |sh|
          if sh.type=="s"
            sh.value = @sharedstrings[sh.value.to_i]
          end
          if !comments.nil?
            comm=comments.select {|c| c[:ref]==(sh.xlcoords)}
            if comm.size > 0
              sh.comment=comm
            end
            comments.delete_if{|c| c[:ref]==(sh.xlcoords)}
          end
        end
        x[:cells] = @sheet.cellarray
        x[:mergedcells] = @sheet.mergedcells
      end
      FileUtils.remove_dir(@destination, true)
      matrixto(options[:copymerge])
    end
    
    private
    # @private
    # Given the data found in workbook.xml, create a hash and push it to the sheets
    # array.
    #
    # The hash will not be pushed into the array if the sheet name is blacklisted
    # (it appears in the *excluded_sheets* array) or does not appear in the list of
    # included sheets.
    #
    # If *included_sheets* (the array of whitelisted sheets) is *nil*, the hash is added.
    def sheetdata options={}
      @a.locate("workbook/sheets/*").each do |x|
        @sheetbase[:name] = x[:name]
        @sheetbase[:sheetId] = x[:sheetId]
        @sheetbase[:relationId] = x[:"r:id"]
        @sheets << @sheetbase
        @sheetbase=Hash.new
      end
      sheetarr=@sheets.map{|i| i[:name]}
      if options[:include].nil?; options[:include]=[]; end
      if options[:include].to_a.size>0
        sheetarr.keep_if{|item| options[:include].to_a.detect{|d| d==item}}
      end
      sheetarr=sheetarr-options[:exclude].to_a
      @sheets.keep_if{|item| sheetarr.detect{|d| d==item[:name]}}
      @sheets.uniq!
    end
    
    # Build the relationship between sheets and the XML files storing the comments
    # to the actual sheet.
    def commentsrel #!!!MI VAN HA NINCS KOMMENT???????
     unless Dir[@destination + '/xl/worksheets/_rels'].empty?
      Find.find(@destination + '/xl/worksheets/_rels') do |path|
        if File.basename(path).split(".").last=='rels'
          f=Ox.load_file(path)
          f.locate("Relationships/*").each do |x|
            if x[:Target].include?"comments"
              @sheets.each do |s|
                if File.basename(path,".rels")=="sheet"+s[:sheetId]+".xml"
                  s[:comments]=x[:Target]
                end
              end
            end
          end
        end
      end
     else
       @sheets.each do |s|
         s[:comments]=nil
       end
     end
    end

    # Invokes the Sharedstrings helper class
    def shstrings
      strings = Sharedstrings.new()
      File.open(@destination + '/xl/sharedStrings.xml', 'r') do |f|
        Ox.sax_parse(strings, f)
      end
      @sharedstrings=strings.stringarray
    end
    
    # Parses the comments related to the actual sheet.  
    def mkcomments(commentfile)
      unless commentfile.nil?
        comms = Comments.new()
        File.open(@destination + '/xl/'+commentfile.gsub('../', ''), 'r') do |f|
          Ox.sax_parse(comms, f)
        end
        return comms.commarray
      end
    end

    # Returns an array of Matrix objects.
    # For each sheet, matrixto first checks the address (xlcoords) of the
    # last cell in the cellarray, then builds a *nil*-filled Matrix object of
    # size *xlcoords.x, xlcoords.y*.
    #
    # The matrix will then be filled with Cell objects according to their coordinates.
    #
    # If the *copymerge* parameter is *true*, it creates a submatrix (minor)
    # of every mergegroup (based on the mergedcells array relative to the actual
    # sheet), and after the only meaningful cell of the minor is found, it is
    # copied back to the remaining cells of the group. The coordinates (xlcoords)
    # of each copied cell is changed to reflect the actual Excel coordinate.
    # 
    # The matrix will replace the array of cells in the actual sheet.
    def matrixto(copymerge)
      @sheets.each_with_index do |sheet, i|
        m=Matrix.build(sheet[:cells].last.y+1, sheet[:cells].last.x+1) {nil}
        sheet[:cells].each do |c|
          m[c.y, c.x]=c
        end
        if copymerge==true
          sheet[:mergedcells].each do |mc|
            a = mc.split(':')
            x1=x(a[0])
            y1=y(a[0])
            x2=x(a[1])
            y2=y(a[1])
            mrange=m.minor(y1..y2, x1..x2)
            valuecell=mrange.to_a.flatten.compact[0]
            (x1..x2).each do |col|
              (y1..y2).each do |row|
                if valuecell != nil
                  valuecell.xlcoords=(col.col_name)+(row+1).to_s
                  m[col, row]=valuecell
                else
                  valuecell=Cell.new
                  valuecell.xlcoords=(col.col_name)+(row+1).to_s
                  m[col, row]=valuecell
                end
              end
            end
          end
        end
        s=Sheet.new
        s.name=@sheets[i][:name]; s.sheetId=@sheets[i][:sheetId]; s.relationId=@sheets[i][:relationId]
        s.data=m
        @sheets[i]=s
      end
    end
  end
end
