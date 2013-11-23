
module Oxcelix
  # The Cellhelper module defines some methods useful to manipulate Cell objects
  module Numformats
    attr_accessor :dtmap
    def initialize
      @dtmap = {'hh.mm'=>'%H.%M', 'hh.m'=>'%H.%-M', 'H.mm'=>'%-k.%M', 'H.m'=>'%-k.%-M', 'mm.ss'=>'%M.%-S', 'm.ss'=>'%-M.%-S', \
          'mm.s'=>'%M.%S', 'm.s'=>'%-M.%S', 'H'=>'%-k', 'hh'=>'%H', 'm{5}'=>'%b', 'm{4}'=>'%B', 'm{3}'=>'%b', 'mm'=>'%m', \
          'm'=>'%-m', 'd{4}'=>'%A', 'd{3}}'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'y{4}'=>'%Y', 'yy'=>'%y', 'AM/PM'=>'%p', \
          'A/P'=>'%p', '.0'=>'', 's'=>'%S', 'ss'=>'%-S'}
      @dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', 'ss'=>'%-S', 's'=>'%S', 'm{5}'=>'%b', 'm{4}'=>'%B',
                'm{3}'=>'%b', 'mm'=>'%m', 'm'=>'%-m', 'd{4}'=>'%A', 'd{3}'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y',
                'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 's'=>'%S', 'ss'=>'%-S'}
      @dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', \
                    'ss'=>'%-S', 's'=>'%S', 'mmmmm'=>'%b', 'mmmm'=>'%B', 'mmmm'=>'%b', 'mm'=>'%m', \
                    'm'=>'%-m', 'dddd'=>'%A', 'ddd'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y', \
                    'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 'ss'=>'%-S', 's'=>'%S'}
    ##?????
    end

    'hh.mm'=>'%H.%M', 'hh.m'=>'%H.%-M', 'H.mm'=>'%-k.%M', 'H.m'=>'%-k.%-M', 'mm.ss'=>'%M.%-S', 'm.ss'=>'%-M.%-S', 'mm.s'=>'%M.%S', 'm.s'=>'%-M.%S', 'H'=>'%-k', 'hh'=>'%H', 'm{5}'=>'%b', 'm{4}'=>'%B', 'm{3}'=>'%b', 'mm'=>'%m', 'm'=>'%-m', 'd{4}'=>'%A', 'd{3}}'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'y{4}'=>'%Y', 'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 's'=>'%S', 'ss'=>'%-S'

#    fhash={:rvalue=>nil, :dvalue=>nil}
SPRINTF vagy FORMAT!!! (Kernel::)
1) 4 csoportra osztas:
        1) Format for positive numbers
        2) Format for negative numbers
        3)  Format for zeros
        4)  Format for text
2) csoportonkent parse
  a) ha szam (pl: #, 0 ÉS nem más ami nem []ben van)
    i) ha van benne . (float?) ->
    ii) ha van benne % () ->
    iii) ha a # és 0 között, után, előtt van elválasztó karakter (, szóköz, egyéb) ->
    iv) ha van benne -, ha van benne / (racionalis szam)
  b) ha ido (van benne d, m, y, h, s DE NEM []ben)->
    i) ha 1 vagy több karakterből áll a substring ->
    i) ha tartalmaz barmilyen elvalasztot a karakterek VAGY stringek között->
  c) konyveloi cucc?
 
if fmt =~ /[#0%]/
  numeric (value, fmt)
elsif fmt.downcase =~ /[dmysh]/
  datetime (value, fmt)
else
  return value, fmt
end

def numeric value, fmt
  ostring = '%'
  #grp=0
  hashes = 0
  zeroes = 0

  arr = fmt.split(".")
  arr.each_with_index do |f,j|
    f.chars.each_with_index do |ch, i|
      if ch == '#'
        hashes += 1
      elsif ch == '0'
        zeroes += 1
      elsif ch == '.'
        ostring = mkstr(ostring, hashes, zeroes, ch)
        hashes = 0
        zeroes = 0
      elsif i == 0
        ostring.prepend ch + ' '
      elsif
        i == f.chars.size - 1
        ostring += ch
      end
    end
    if j==0 || arr.size > 1
      ostring += "."
    end
    ostring = mkstr(ostring, hashes, zeroes, '')
  end
end
  #if grp <= 1
  #formatted_n = sprintf('%2.2f $', n).to_s.reverse.gsub(/\d#{grp}(?=\d)/, '\&').reverse ####ATIRANDUSZ

  def mkstr (ostring, hashes, zeroes, ch)
    if hashes > 0 && !ostring.include?(".")
      ostring += '#' + hashes.to_s  #nem jo! ketszer fogja a hashmarkot beletenni 
    end
    if zeroes > 0
      ostring += zeroes.to_s
    end
    if ostring.include?(".") && !ostring.include?("e")
      ostring += "f"
    elsif ostring.include?(".") && ostring.include?("e")
      ostring += "e"
    else
      ostring += "d"
    end
    ostring += ch
  end



  fmtarr = fmt.split('.')
  t = ''
  if fmtarr.size > 1
    t = 'f'
  else
    t = 'd'
  end
  hmark = 0
  zeroes = 0
  format = '%'
  fmtarr.eack do |x|
    x.scan(/#/) { hmark += 1}
    x.scan(/0/) { zeroes += 1}
  if zeroes <= 0
    zeroes = ''
  end
  format += zeroes.to_s + hmark.to_s
  #e, %, $, vesszo stb.

end

def datetime value, fmt
  deminutified = fmt.gsub(/(?<hrs>H|h)(?<div>.)m/, '\k<hrs>\k<div>i')
                     .gsub(/im/, 'ii')
                     .gsub(/m(?<div>.)(?<secs>s)/, 'i\k<div>\k<secs>')
                     .gsub(/mi/, 'ii')
  rubified = deminutified.gsub /[yMmDdHhSsi]/, @dtmap
end

 
    FARY=[
    #    ID  Format Code
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value} }, #0   General 
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value.to_i} }, #1   0 
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{@value.to_f} m[c.y, c.x]=c}, #2   0.00 
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse} },#3   #,##0
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{@value.to_f.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse} }, #4   #,##0.00
    nil, nil, nil, nil, #5, 6, 7, 8
    { :rvalue => Proc.new{@value.to_i}, :dvalue => Proc.new{("%2.2i\%" % @value)} }, #9   0% ##*100?
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{("%2.2f\%" % @value)} }, #10  0.00%
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{("%E" % @value)} }, #11  0.00E+00
12  # ?/? RATIONAL?
13  # ??/?? RATIONAL?
    { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%d/%m/%Y")} }, #14  d/m/yyyy
    { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%d-%b-%y")} }, #15  d-mmm-yy
    { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%d-%b")} }, #16  d-mmm
17  { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%b-%y")} }, #mmm-yy
18  h:mm tt
19  h:mm:ss tt
20  H:mm
21  H:mm:ss
22  m/d/yyyy H  #%m/,##0 ;(#,##0)
nil,nil,nil,nil,nil,nil,nil,nil,nil,nil,nil,nil,nil,nil,nil, #23-37
38  #,##0 ;[Red](#,##0)
39  #,##0.00;(#,##0.00)
40  #,##0.00;[Red](#,##0.00)
45  mm:ss
46  [h]:mm:ss
47  mmss.0
48  ##0.0E+0
49  @ ]

  end
  module Cellvalues
    # Set the excel cell name (eg: 'A2')
    # @param [String] val Excel cell address name
    def r(val); @xlcoords = val; end;
    # Set cell type    
    def t(val); @type = val; end;
    # Set cell value
    def v(val); @value = val; end;
    # Set cell style (number format and style) 
    def s(val); @style = val; end;
  end
  
  module Cellhelper
    # When called without parameters, returns the x coordinate of the calling cell object based on the value of #@xlcoords
    # If a parameter is given, #x will return the x coordinate corresponding to the parameter
    # @example find x coordinate (column number) of a cell
    #   c = Cell.new
    #   c.xlcoords = ('B3')
    #   c.x #=> 1
    # @param [String] coord Optional parameter used when method is not called from a Cell object
    # @return [Integer] x coordinate
    def x(coord=nil)
      if coord.nil?
        coord = @xlcoords
      end
      ('A'..(coord.scan(/\p{Alpha}+|\p{Digit}+/u)[0])).to_a.length-1
    end

    # When called without parameters, returns the y coordinate of the calling cell object based on the value of #@xlcoords
    # If a parameter is given, #y will return the y coordinate corresponding to the parameter
    # @example find y coordinate (row number) of a cell
    #   c = Cell.new
    #   c.xlcoords = ('B3')
    #   c.y #=> 2
    # @param [String] coord Optional parameter used when method is not called from a Cell object
    # @return [Integer] x coordinate
    def y(coord=nil)
      if coord.nil?
        coord = @xlcoords
      end
      coord.scan(/\p{Alpha}+|\p{Digit}+/u)[1].to_i-1
    end
    def to_ru; fary[@style].call :rvalue; end;
    def to_d; fary[@style].call :dvalue; end;
  end
end