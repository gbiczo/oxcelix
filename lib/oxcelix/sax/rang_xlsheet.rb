module Oxcelix
  class Sheetrange < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    
    def initialize(range_start, range_end)
      @RANGE_START=range_start
      @RANGE_END=range_end
      super
    end
    
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric? && ((@RANGE_START..@RANGE_END).include? @cell.xlcoords))
          @cell.v str
          @cellarray << @cell
        end
        @cell=Cell.new
      end
    end

  end
end
