
module Oxcelix
  # The Cellhelper module defines some methods useful to manipulate Cell objects
  module Numformats
#    fhash={:rvalue=>nil, :dvalue=>nil}
    fary=[
    #    ID  Format Code
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value} }, #0   General 
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value.to_i} }, #1   0 
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{@value.to_f} }, #2   0.00 
    { :rvalue => Proc.new{@value}, :dvalue => Proc.new{@value.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse} },#3   #,##0
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{@value.to_f.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse} }, #4   #,##0.00
    { :rvalue => Proc.new{@value.to_i}, :dvalue => Proc.new{("%2.2i\%" % @value)} }, #9   0% ##*100?
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{("%2.2f\%" % @value)} }, #10  0.00%
    { :rvalue => Proc.new{@value.to_f}, :dvalue => Proc.new{("%E" % @value)} }, #11  0.00E+00
12  # ?/? RATIONAL?
13  # ??/?? RATIONAL?
    { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%d/%m/%Y")} }, #14  d/m/yyyy
    { :rvalue => Proc.new{DateTime.new(1899,12,30)+@value}, :dvalue => Proc.new{(DateTime.new(1899,12,30)+@value).strftime("%d/%b/%y")} }, #15  d-mmm-yy
16  d-mmm
17  mmm-yy
18  h:mm tt
19  h:mm:ss tt
20  H:mm
21  H:mm:ss
22  m/d/yyyy H  #%m/,##0 ;(#,##0)
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
    def to_r; fary[@style].call :rvalue; end;
    def to_d; fary[@style].call :dvalue; end;
  end
end