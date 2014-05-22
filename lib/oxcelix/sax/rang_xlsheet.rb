module Oxcelix
  class Paginatedsheet < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    
    def initialize(per_page, pageno)
      @PER_PAGE=per_page
      @PAGENO=pageno
      super
    end
    
    def attr(name, str)
      # if PER_PAGE?
      if @xmlstack.last == :c
        @cell.send name, str if @cell.respond_to?(name)
      elsif xmlstack.last == :mergeCell && name == :ref
        @mergedcells << str
      end
    end
    
  end
end
