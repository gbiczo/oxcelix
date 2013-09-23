require 'ox'
require 'find'
require 'matrix'
require 'fileutils'

require 'debugger'
debugger
class String
  def numeric?
    Float(self) != nil rescue false
  end
end

class Fixnum
  def col_name
    val=self/26
    (val > 0 ? (val - 1).col_name : "") + (self % 26 + 65).chr
  end
end

class Matrix
  def []=(i, j, x)
    @rows[i][j]=x
  end
end

module Cellhelper
  def r(val); @xlcoords = val; end;
  def t(val); @type = val; end;
  def v(val); @value = val; end;
  def s(val); @style = val; end;
  
  def x(coord=nil)
    if coord.nil?
      coord = @xlcoords
    end
    ('A'..(coord.scan(/\p{Alpha}+|\p{Digit}+/u)[0])).to_a.length

  end

  def y(coord=nil)
    if coord.nil?
      coord = @xlcoords
    end
    coord.scan(/\p{Alpha}+|\p{Digit}+/u)[1].to_i
  end
end

class Cell
  attr_accessor :xlcoords, :type, :value, :comment
  include Cellhelper
end

class Xlsheet < ::Ox::Sax
  attr_accessor :xmlstack, :lastopennode, :mergedcells, :cellarray, :cell
  def initialize()
    @xmlstack = []
    @lastopennode = nil
    @mergedcells = []
    @cellarray = []
    @cell = Cell.new
  end
  
  def start_element(name)
    if name == :c || name == :mergeCell
      @xmlstack << name
    end
  end
  
  def end_element(name)
    @xmlstack.pop
    if name == :c || name == :mergeCell
      @cell=Cell.new
    end
  end
  
  def attr(name, str)
    if @xmlstack.last == :c # && name != :s
      @cell.send name, str if @cell.respond_to?(name)
    elsif xmlstack.last == :mergeCell && name == :ref
      @mergedcells << str
    end
  end
  
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

class Sharedstrings < ::Ox::Sax
  attr_accessor :stringarray
  def initialize
    @stringarray=[]
  end
  
  def text(str)
    @stringarray << str
  end
end

class Comments < ::Ox::Sax
  attr_accessor :commarray, :comment
  def initialize
    @commarray=[]
    @comment={}
  end
  
  def text(str)
    @comment[:comment]=str.gsub('&#10;', '')
    @commarray << @comment
    @comment = Hash.new
  end
  
  def attr(name, str)
    if name == :ref
      @comment[:ref]=str
    end
  end
end

class Sheet
  attr_accessor :name, :sheetId, :relationId, :data
  def initialize; @data=[]; end
end

class Workbook
  include Cellhelper
  attr_accessor :sheetbase, :sheets, :sharedstrings
  def initialize(filename)
    require 'zip'
    #mkdir tmp!!!
    @destination = 'tmp'
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
    sheetdata; commentsrel; shstrings;
    @sheets.each do |x|
      sname="sheet#{x[:sheetId]}"
      @sheet = Xlsheet.new()
      File.open(@destination+"/xl/worksheets/#{sname}.xml", 'r') do |f|
        Ox.sax_parse(@sheet, f)
      end
      comments=mkcomments(x[:comments])
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
    matrixto(args)
  end
  
  private
  def sheetdata
    @a.locate("workbook/sheets/*").each do |x|
      @sheetbase[:name] = x[:name]
      @sheetbase[:sheetId] = x[:sheetId]
      @sheetbase[:relationId] = x[:"r:id"]
      @sheets << sheetbase
      @sheetbase=Hash.new
    end
  end
  
  def commentsrel
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
  end
  
  def shstrings
    strings = Sharedstrings.new()
    File.open(@destination + '/xl/sharedStrings.xml', 'r') do |f|
      Ox.sax_parse(strings, f)
    end
    @sharedstrings=strings.stringarray
  end
  
  def mkcomments(commentfile)
    unless commentfile.nil?
      comms = Comments.new()
      File.open(@destination + '/xl/'+commentfile.gsub('../', ''), 'r') do |f|
        Ox.sax_parse(comms, f)
      end
      return comms.commarray
    end
  end

  def matrixto(*args)
    @sheets.each_with_index do |sheet, i|
      m=Matrix.build(sheet[:cells].last.y, sheet[:cells].last.x) {nil}#!!!!! each!!! SHEETS or SHEET?
      sheet[:cells].each do |c|
        m[c.y-1, c.x-1]=c#.value ##value? c? comment?#
      end
      if args.include? 'copymerged' #ez majd egy hash legyen
        sheet[:mergedcells].each do |mc|
          a = mc.split(':')
          #return matrix cell addresses
          x1=x(a[0])-1 #-1 ?
          y1=y(a[0])-1
          x2=x(a[1])-1 #-1 ?
          y2=y(a[1])-1
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
      @sheets[i][:cells]=m
    end
  end
end

#sheet1.cellarray.each do |x|
##    puts "old values: #{x.inspect}"
#    a=x.value.to_i
#    if x.type=="s"
#   x.value = strings.stringarray[a]
#    end
##    puts "new values: #{x.inspect}"
#end
w=Workbook.new('Schedules3.xlsx')
puts "sheet0: #{w.sheets[0]}"
#puts "Matrix: #{w.to_m('copymerged').inspect}"

#puts "sheetbase: #{w.sheetbase}"
#puts "sheets: #{w.sheets}"
#puts w.sheets[2]
#puts w.sheets[0].mergedcells
