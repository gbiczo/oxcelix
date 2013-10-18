# The namespace for all classes and modules included on Oxcelix.
module Oxcelix
  # Helper methods for the Workbook class
  module Workbookhelper
    # returns a sheet based on its name

    # @example Select a sheet
    #   w = Workbook.new('Example.xlsx')
    #   sheet = w["Examplesheet"]
    def [] (sheetname=String)
      @sheets.select{|s| s.name==sheetname}[0]
    end
  end

  # The Workbook class will open the excel file, and convert it to a collection of
  # Matrix objects
  # @!attribute [rw] sheets
  #   @return [Array] a collection of {Sheet} objects 
  class Workbook
    include Cellhelper
    include Workbookhelper
    attr_accessor :sheets
    ##
    # Create a new {Workbook} object.
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
      @sharedstrings=[]
      
      f=IO.read(@destination+'/xl/workbook.xml')
      @a=Ox::load(f)
      
      sheetdata(options); commentsrel; shstrings;

      @sheets.each do |x|

        @sheet = Xlsheet.new()

        File.open(@destination+"/xl/#{x[:filename]}", 'r') do |f|
          Ox.sax_parse(@sheet, f)
        end
        comments = mkcomments(x[:comments])
        @sheet.cellarray.each do |sh|
          if sh.type=="s"
            sh.value = @sharedstrings[sh.value.to_i]
          end
          if !comments.nil?
            comm=comments.select {|c| c[:ref]==(sh.xlcoords)}
            if comm.size > 0
              sh.comment=comm[0][:comment]
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

        relationshipfile=nil
        fname=nil
        unless Dir[@destination + '/xl/_rels'].empty?
          Find.find(@destination + '/xl/_rels') do |path|
            if File.basename(path).split(".").last=='rels'
              g=IO.read(path)
              relationshipfile=Ox::load(g)
            end
          end
        end
        relationshipfile.locate("Relationships/*").each do |rship|
          if rship[:Id] == x[:"r:id"]
            @sheetbase[:filename]=rship[:Target]
          end
        end


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
    def commentsrel
     unless Dir[@destination + '/xl/worksheets/_rels'].empty?
      Find.find(@destination + '/xl/worksheets/_rels') do |path|
        if File.basename(path).split(".").last=='rels'
          a=IO.read(path)
          f=Ox::load(a)
          f.locate("Relationships/*").each do |x|
            if x[:Target].include?"comments"
              @sheets.each do |s|
                if "worksheets/" + File.basename(path,".rels")==s[:filename]
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
    # @param [String] commentfile
    # @return [Array] a collection of comments relative to the Excel sheet currently processed
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
    # @param [Bool ] copymerge 
    # @return [Matrix] a Matrix object that stores the cell values, and, depending on the copymerge parameter, will copy the merged value
    #  into every merged cell
    def matrixto(copymerge)
      @sheets.each_with_index do |sheet, i|
        m=Sheet.build(sheet[:cells].last.y+1, sheet[:cells].last.x+1) {nil}
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
                  m[row, col]=valuecell
                else
                  valuecell=Cell.new
                  valuecell.xlcoords=(col.col_name)+(row+1).to_s
                  m[row, col]=valuecell
                end
              end
            end
          end
        end
        m.name=@sheets[i][:name]; m.sheetId=@sheets[i][:sheetId]; m.relationId=@sheets[i][:relationId]
        @sheets[i]=m
      end
    end
  end
end
