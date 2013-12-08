require 'ox'
module Oxcelix

  # Ox based SAX parser which pushes the number formats (taken from the styles.xml file) to an array 
  # The reference taken from the cell's 's' attribute points to an element of the
  # style array, which in turn points to a number format (numFmt) that can be
  # either built-in (@formats) or defined in the styles.xml itself.
  class Styles < ::Ox::Sax
    attr_accessor :styleary, :xmlstack, :temparray
    def initialize
      @temparray=[]
      @styleary=[]
      @xmlstack = []
      @numform={}
    end

    def nf key, value
      @numform[key]=value
      if @numform.size == 2
        @temparray << @numform
        @numform = {}
      end
    end

    def numFmtId str
      if @xmlstack[-2] == :cellXfs
        @styleary << str
      elsif @xmlstack[-2] == :numFmts
        nf :numFmtId, str
      end
    end

    def formatCode str
      nf :formatCode, str
    end
    
    def start_element(name)
      @xmlstack << name
    end

    def end_element(name)
      @xmlstack.pop
    end

    def attr(name, str)
      self.send name, str if self.respond_to?(name)
    end
  end
end