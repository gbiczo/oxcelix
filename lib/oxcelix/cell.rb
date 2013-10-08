module Oxcelix
  # Simple class representing an excel cell.
  # @!attribute [rw] xlcoords
  #   @return [String] The Excel-style coordinates of the cell object
  # @!attribute [rw] type
  #   @return [String] Cell content type
  # @!attribute [rw] value
  #   @return [String] the type of the cell
  # @!attribute [rw] comment
  #   @return [String] Comment text
  # @!attribute [rw] style
  #   @return [String] Excel style attribute
  class Cell
    attr_accessor :xlcoords, :type, :value, :comment, :style
    include Cellhelper
    include Cellvalues
  end
end
