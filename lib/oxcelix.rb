require 'ox'
require 'find'
require 'matrix'
require 'fileutils'
require 'tmpdir'
require 'zip'
require_relative './oxcelix/cellhelper'
require_relative './oxcelix/numformats'
require_relative './oxcelix/nf'
require_relative './oxcelix/cell'
require_relative './oxcelix/sheet'
require_relative './oxcelix/workbook'
require_relative './oxcelix/sax/sharedstrings'
require_relative './oxcelix/sax/comments'
require_relative './oxcelix/sax/xlsheet'
require_relative './oxcelix/sax/styles'

class String
  # Returns true if the given String represents a numeric value 
  # @return [Bool] true if the string represents a numeric value, false if it doesn't.
  def numeric?
    Float(self) != nil rescue false
  end
end

class Matrix
  # Set the cell value i,j to x, where i=row, and j=column.
  # @param i [Integer] Row
  # @param j [Integer] Column
  # @param x [Object] New cell value
  def []=(i, j, x)
    @rows[i][j]=x
  end
end

class Fixnum
  # Returns the column name corresponding to the given number. e.g: 1.col_name => 'B'
  # @return [String]
  def col_name
    val=self/26
    (val > 0 ? (val - 1).col_name : "") + (self % 26 + 65).chr
  end
end

# The namespace for all classes and modules included in Oxcelix.
module Oxcelix

end
