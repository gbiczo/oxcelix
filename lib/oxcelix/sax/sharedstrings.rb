module Oxcelix
  # Ox based SAX parser which pushes shared strings (taken from the sharedString.xml file) to an array 
  # These strings will replace the references in the cells (interpolation). 
  class Sharedstrings < ::Ox::Sax
    # @!attribute [rw] stringarray
    #   @return [Array] the array of all the strings found in sharedStrings.xml
    attr_accessor :stringarray
    def initialize
      @stringarray=[]
    end

    # Push the comment string into @stringarray
    def text(str)
      @stringarray << str
    end
  end
end
