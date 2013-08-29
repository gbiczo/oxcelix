require 'ox'
require 'find'
require 'matrix'

class String
  def numeric?
    Float(self) != nil rescue false
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
  def x; ('A'..(xlcoords.scan(/\p{Alpha}+|\p{Digit}+/u)[0])).to_a.length; end
  def y; xlcoords.scan(/\p{Alpha}+|\p{Digit}+/u)[1].to_i; end
end

class Cell
  attr_accessor :xlcoords, :type, :value, :comment#, :x, :y
  include Cellhelper
end

class Xlsheet < ::Ox::Sax
  attr_accessor :xmlpath, :lastopennode, :mergedcells, :cellarray, :cell
  def initialize()
    @xmlstack = []
    @lastopennode = nil
    @mergedcells = []
    @cellarray = [] #XML.new, vagy valami?
    @cell = Cell.new
  end
  
  def start_element(name)
    if name == :c || name == :mergeCell
      xmlstack << name
    end
  end
  
  def end_element(name)
    xmlstack.pop
  end
  
  def attr(name, str)
    if xmlstack.last == :c && name != :s
      cell.send name, str if cell.respond_to?(name)
    elsif xmlstack.last == :mergeCell && name == :ref
      mergedcells << str
    end
  end
  
  def text(str)
    if xmlstack.last == :c
      if cell.type != "shared" && cell.type != "e" && str.numeric?
        cell.v str
        cellarray << cell
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
    stringarray << str
  end
end

class Comments < ::Ox::Sax
  attr_accessor :commarray, :comment
  def initialize
    @commarray=[]
    @comment={}
  end
  
  def text(str)
    comment[:comment]=str.gsub('&#10;', '')
    commarray << comment
    comment = Hash.new
  end
  
  def attr(name, str)
    if name == :ref
      comment[:ref]=str
    end
  end
end

class Sheet
  attr_accessor :name, :sheetId, :relationId, :data
  def initialize; data=[]; end
end

class Workbook
  attr_accessor :sheetsmeta, :sheetbase, :sheets, :sharedstrings
  def initialize #FILENAME!
    @sheets=[]
    @sheetbase={}
    @a=Ox::load_file('xl/workbook.xml')
    sheetdata; commentsrel; shstrings;
    sheets.each do |x|
      sname="sheet#{x[:sheetId]}"
      sheet = Xlsheet.new()
      File.open("xl/worksheets/#{sname}.xml", 'r') do |f|
        Ox.sax_parse(sheet, f)
      end
      comments=mkcomments(x[:comments])
      sheet.cellarray.each do |sh|
        if sh.type=="s"
          sh.value = sharedstrings[sh.value.to_i]
        end
        if !comments.nil?
          comm=comments.select {|c| c[:ref]==(sh.xlcoords)}
          if comm.size > 0
            sh.comment=comm
          end
          comments.delete_if{|c| c[:ref]==(sh.xlcoords)}
        end
      end
      x[:cells] = sheet.cellarray
      x[:mergedcells] = sheet.mergedcells
    end
  end
  
  def to_m(*args) #self? module?
    #args: merged cells, excel-style columns, value or comment?
    # Cell exclusion here?
    #elso-utolso cella koordinatai
    #    matrices=[]
    m=Matrix.build(sheets[0][:cells].last.x, sheets[0][:cells].last.y) {nil}#!!!!! each!!!
    sheets[0][:cells].each do |c|
      m[c.x-1, c.y-1]=c.value ##value? c? comment?#
    end
    return m
    #puts @sheets[0][:cells].last.y
    #excel-style?
    #matrix felepitese
    #cellankent x, y
    #merge-unk van?
    #matrix visszaad
  end
  
  private
  def sheetdata
    @a.locate("workbook/sheets/*").each do |x|
      sheetbase[:name] = x[:name]
      sheetbase[:sheetId] = x[:sheetId]
      sheetbase[:relationId] = x[:"r:id"]
      sheets << sheetbase
      sheetbase={}
    end
  end
  
  def commentsrel
    Find.find('xl/worksheets/_rels') do |path|
      if File.basename(path).split(".").last=='rels'
        f=Ox.load_file(path)
        f.locate("Relationships/*").each do |x|
          if x[:Target].include?"comments"
            sheets.each do |s|
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
    File.open('xl/sharedStrings.xml', 'r') do |f|
      Ox.sax_parse(strings, f)
    end
    sharedstrings=strings.stringarray
  end
  
  def mkcomments(commentfile)
    unless commentfile.nil?
      comms = Comments.new()
      File.open('xl/'+commentfile.gsub('../', ''), 'r') do |f|
        Ox.sax_parse(comms, f)
      end
      return comms.commarray
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
w=Workbook.new
puts w.to_m
#puts w.sheets[2]
#puts w.sheets[0].mergedcells
