require 'ox'
module Oxcelix

  # Ox based SAX parser which pushes the number formats (taken from the styles.xml file) to an array 
  # The reference taken from the cell's 's' attribute points to an element of the
  # style array, which in turn points to a number format (numFmt) that can be
  # either built-in (@formats) or defined in the styles.xml itself.
  class Styles < ::Ox::Sax
    attr_accessor :formats, :defined_formats, :styleary, :xmlstack, :temparray
    def initialize
      @formats=[
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00", nil, nil, nil, nil,
        "0%",
        "0.00%",
        "0.00E+00",
        "# ?/?",
        "# ??/??",
        "d/m/yyyy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "h:mm tt",
        "h:mm:ss tt",
        "H:mm",
        "H:mm:ss",
        "m/d/yyyy H:mm", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil,
        "#,##0 ;(#,##0)",
        "#,##0 ;[Red](#,##0)",
        "#,##0.00;(#,##0.00)",
        "#,##0.00;[Red](#,##0.00)", nil, nil, nil, nil,
        "mm:ss",
        "[h]:mm:ss",
        "mmss.0",
        "##0.0E+0",
        "@,"
      ]
      defined_formats=Array.new(114)
      @formats.push *defined_formats
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
