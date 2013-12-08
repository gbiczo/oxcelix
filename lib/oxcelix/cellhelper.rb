module Oxcelix

  # The Cellvalues module provides methods for setting cell values. They are named after the relevant XML entitiesd and 
  # called directly by the Xlsheet SAX parser.
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
  
  # The Cellhelper module defines some methods useful to manipulate Cell objects
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
  end
end