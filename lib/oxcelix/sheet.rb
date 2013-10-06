
module Oxcelix
  # The Sheet class represents an excel sheet.
  # @!attribute [rw] name
  #   @return [String] Sheet name
  # @!attribute [rw] sheetId
  #   @return [String] returns the sheetId SheetML internal attribute
  # @!attribute [rw] relationId
  #   @return [String] returns the relation key used to reference the comments file related to a certain sheet.
  # @!attribute [rw] data
  #   @return [Array] stores the collection of cells of a certain sheet. This is the copy of the sum of cellarray and mergedcells.
  #   This is laterconverted to a matrix
  class Sheet
    attr_accessor :name, :sheetId, :relationId, :data
    def initialize; @data=[]; end
  end
end
