module Oxcelix
  class Paginatedsheet < Xlsheet
    attr_accessor :xmlstack, :mergedcells, :cellarray, :cell
    def initialize(per_page, pageno)
      @PER_PAGE=per_page
      @PAGENO=pageno
      super
    end
  end
end
