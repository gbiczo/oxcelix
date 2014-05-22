module Oxcelix
  class Paginatedsheet < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    
    def initialize(per_page, pageno)
      @PER_PAGE=per_page
      @PAGENO=pageno
      super
    end
    
    def text(str)
      if @xmlstack.last == :c
        if @cell.type != "shared" && @cell.type != "e" && str.numeric? && (@cell.y.between? (@PER_PAGE*(@PAGENO-1), @PER_PAGE*@PAGENO-1))
          @cell.v str
          @cellarray << @cell
        end
        @cell=Cell.new
      end
    end

  end
end
