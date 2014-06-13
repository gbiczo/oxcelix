
module Oxcelix
  ##
  # The Xlsheet class is a SAX parser based on the Ox library. It parses a
  # SpreadsheetML (AKA Office Open XML) formatted XML file and returns an array
  # of Cell objects {#cellarray} and an array of merged cells {#mergedcells}.
  #
  # Xlsheet will omit the following:
  # * empty cells
  # * cells containing formulas
  #
  # Only non-empty cells of merged groups will be added to {#cellarray}. A separate array
  # {#mergedcells} is reserved for merging.
  class Xlsheet < ::Ox::Sax
  # @!attribute [rw] xmlstack
  #   @return [Array] Stores the state machine's actual state
  # @!attribute [rw] mergedcells
  #   @return [Array] the array of merged cells
  # @!attribute [rw] cellarray
  #   @return [Array] the array of non-empty (meaningful) cells of the current sheet
  # @!attribute [rw] cell
  #   @return [Cell] the cell currently being processed.
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    def initialize()
      @xmlstack = []
      @mergedcells = []
      @cellarray = []
      @cell = Cell.new
    end

    # Save SAX state-machine state to {#xmlstack} if and only if the processed
    # element is a :c (column) or a :mergeCell (merged cell)
    # @param [String] name Start element
    def start_element(name)
      case name
      when :c
        @xmlstack << name
      when :mergeCell
        @xmlstack << name
      end
    end
    
    # Step back in the stack ({#xmlstack}.pop), clear actual cell information
    # @param [String] name Element ends
    def end_element(name)
      @xmlstack.pop
      case name
      when :c
        @cell=Cell.new
      when :mergeCell
        @cell=Cell.new
      end
    end
    
    # Set cell value, style, etc. This will only happen if the cell has an
    # actual value AND the parser's state is :c.
    # If the state is :mergeCell AND the actual attribute name is :ref the
    # attribute will be added to the merged cells array.
    # The attribute name is tested against the Cell object: if the cell
    # has a method named the same way, that method is called with the str parameter.
    # @param [String] name of the attribute.
    # @param [String] str Content of the attribute 
    def attr(name, str)
      case @xmlstack.last
      when :c
        @cell.send name, str if @cell.respond_to?(name)
      when :mergeCell
        @mergedcells << str if name == :ref
      end
    end

    # Cell content is parsed here. For cells containing strings, interpolation using the
    # sharedStrings.xml file is done in the #Sharedstrings class.
    # The numformat attribute gets a value here based on the styles variable, to preserve the numeric formatting (thus the type) of values.
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric?
          @cell.v str
          @cellarray << @cell
        end
        @cell=Cell.new
      end
    end
  end

  class Sheetpage < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell

    class << @cellarray
      def << value
        super(value)
        yield @cellarray
      end
    end
        
    def initialize(per_page, pageno)
      @PER_PAGE=per_page
      @PAGENO=pageno
      super()
    end
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric? && ((@PER_PAGE * (@PAGENO-1)..(@PER_PAGE*@PAGENO-1)).include?@cell.y)
          @cell.v str
          @cellarray << @cell
        end
        @cell=Cell.new
      end
    end
  end

  class Sheetrange < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    
    class << @cellarray
      def << value
        super(value)
        yield @cellarray
      end
    end

    def initialize(range)
      @cell=Cell.new
      @RANGE_START=range.begin
      @RANGE_END=range.end
      super()
    end
   
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric? 
          if (((@cell.x(@RANGE_START)..@cell.x(@RANGE_END)).include? @cell.x) && ((@cell.y(@RANGE_START)..@cell.y(@RANGE_END)).include? @cell.y))
            @cell.v str
            @cellarray << @cell
          end
        end
        @cell=Cell.new
      end
    end

  end
end