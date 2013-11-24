module Oxcelix
  # The Cellhelper module defines some methods useful to manipulate Cell objects
  module Numformats
    attr_accessor :dtmap
    def initialize
      @dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', \
                    'ss'=>'%-S', 's'=>'%S', 'mmmmm'=>'%b', 'mmmm'=>'%B', 'mmmm'=>'%b', 'mm'=>'%m', \
                    'm'=>'%-m', 'dddd'=>'%A', 'ddd'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y', \
                    'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 'ss'=>'%-S', 's'=>'%S'}
    end

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
  ostring = "%"
  strippedfmt = fmt.gsub /\?/, '0'
  puts "fmt: #{fmt}, strippedfmt: #{strippedfmt}"
  prefix, decimals, sep, floats, expo, postfix=/(^[^\#0e].?)?([\#0]*)?(.)?([\#0]*)?(e.?)?(.?[^\#0e]$)?/i.match(strippedfmt).captures
  ostring.prepend prefix.to_s
  puts "prefix: #{prefix}, decimals: #{decimals}, sep:#{sep}, floats: #{floats}, expo: #{expo}, postfix: #{postfix}"
  if !decimals.nil? && decimals.size != 0
    if (eval decimals) == nil
      ostring += "##{decimals.size}"
    elsif (eval decimals) == 0
      ostring += decimals.size.to_s
    end
  else
    ostring += decimals
  end
  ostring += sep.to_s
  if !floats.nil? && floats.size != 0
    ostring += ((floats.size.to_s) +"f")
  end
  if sep.nil? && floats.nil? || floats.size == 0
    ostring += "d"
  end
  ostring += (expo.to_s + postfix.to_s)
  puts "ostring: #{ostring}"
  return ostring
end

def datetime value, fmt
  deminutified = fmt.gsub(/(?<hrs>H|h)(?<div>.)m/, '\k<hrs>\k<div>i')
                     .gsub(/im/, 'ii')
                     .gsub(/m(?<div>.)(?<secs>s)/, 'i\k<div>\k<secs>')
                     .gsub(/mi/, 'ii')
  rubified = deminutified.gsub /[yMmDdHhSsi]/, @dtmap
end


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