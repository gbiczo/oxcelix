
module Oxcelix
  # The Sheet class represents an excel sheet.
  class Sheet < Matrix
    include Cellhelper
  # @!attribute [rw] name
  #   @return [String] Sheet name
  # @!attribute [rw] sheetId
  #   @return [String] returns the sheetId SheetML internal attribute
  # @!attribute [rw] relationId
  #   @return [String] Internal reference key. relationID is used internally by Excel 2010 to e.g. build up the relationship between worksheets and comments
    attr_accessor :name, :sheetId, :relationId

    # The [] method overrides the standard Matrix::[]. It will now accept Excel-style cell coordinates.
    # @param [String] i
    # @return [Cell] the object denoted with the Excel column-row name.
    # @example Select a cell in a sheet
    #   w = Oxcelix::Workbook.new('Example.xlsx')
    #   w.sheets[0][3,1] #=> #<Oxcelix::Cell:0x00000001e00fa0 @xlcoords="B4", @style="0", @type="n", @value="3">
    #   w.sheets[0]['B4'] #=> #<Oxcelix::Cell:0x00000001e00fa0 @xlcoords="B4", @style="0", @type="n", @value="3">
    def [](i, *j)
      if i.is_a? String
        super(y(i),x(i))
      else
        super(i,j[0])
      end
    end

    def to_m(*attrs)
      m=Matrix.build(self.col(0).length, self.row(0).length){nil}
      self.each do |x, row, col|
        m[row, col]=x
      end
      return m
    end
  end
end
