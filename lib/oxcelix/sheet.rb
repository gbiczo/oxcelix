
module Oxcelix
  # The Sheet class represents an excel sheet.
  class Sheet < Matrix
    include Cellhelper
    include Numberhelper
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
    
    #The to_m method returns a simple Matrix object filled with the raw values of the original Sheet object.
    # @return [Matrix] a collection of string values (the former #Cell::value)
    def to_m(*attrs)
      m=Matrix.build(self.row_size, self.column_size){nil}
      self.each_with_index do |x, row, col|
        if attrs.size == 0 || attrs.nil?
          m[row, col] = x.value
        end
      end
      return m
    end
    
    # The to_ru method returns a Matrix of "rubified" values. It basically builds a new Matrix
    # and puts the result of the #Cell::to_ru method of every cell of the original sheet in
    # the corresponding Matrix cell.
    # @return [Matrix] a collection of ruby objects (#Integers, #Floats, #DateTimes, #Rationals, #Strings)
    def to_ru
      m=Matrix.build(self.row_size, self.column_size){nil}
      self.each_with_index do |x, row, col|
        if x.nil?
          m[row, col] = nil
        else
          m[row, col] = x.to_ru
        end
      end
      return m
    end
    
    # The to_fmt method returns a Matrix of "formatted" values. It basically builds a new Matrix
    # and puts the result of the #Cell::to_fmt method of every cell of the original sheet in
    # the corresponding Matrix cell. The #Cell::to_fmt will pass the original values to to_ru, and then
    # depending on the value, will run strftime on DateTime objects and sprintf on numeric types.
    # @return [Matrix] a collection of Strings
    def to_fmt
      m=Matrix.build(self.row_size, self.column_size){nil}
      self.each_with_index do |x, row, col|
        if x.nil?
          m[row, col] = nil
        else
          m[row, col] = x.to_fmt
        end
     end
     return m
   end
 end
end
