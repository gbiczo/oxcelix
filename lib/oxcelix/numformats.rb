module Oxcelix
# The Numformats module provides helper methods that either return the Cell object's raw @value as a ruby value
# (e.g. Numeric, DateTime, String) or formats it according to the excel _numformat_ string (#Cell.numformat).
    module Numformats
      @dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', \
                    'ss'=>'%-S', 's'=>'%S', 'mmmmm'=>'%b', 'mmmm'=>'%B', 'mmmm'=>'%b', 'mm'=>'%m', \
                    'm'=>'%-m', 'dddd'=>'%A', 'ddd'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y', \
                    'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 'ss'=>'%-S', 's'=>'%S'}

    # Get the cell's value and excel format string and return a string, a ruby Numeric or a DateTime object accordingly
    def to_ru
      if @numformat == nil || @numformat.downcase == "generic"
        return @value
      end
      if @numformat =~ /[#0%\?\/]/
        return eval @value
      elsif @numformat.downcase =~ /[dmysh]/
        if (0.0..1.0).include? @value
          return DateTime.new(1900, 01, 01) + (eval @value)
        else
          return DateTime.new(1899, 12, 31) + (eval @value)
        end
      else
        eval @value rescue @value
      end
    end

    # Get the cell's value, convert it with to_ru and finally, format it based on the value's type.
    def to_fmt
    v = to_ru
    if v.is_a? DateTime
      datetime v
    elsif v.is_a? Numeric
      numeric v
    else 
      return @value.to_s
    end
  end

  private
  # Convert the excel-style number format to a ruby #Kernel::Format string and return a String using sprintf.
  # the conversion is internally done by regexp'ing 7 groups: prefix, decimals, separator, floats, exponential (E+)
  # and postfix. Rational numbers ar not handled yet.
  # @parameter
  def numeric val
    ostring = "%"
    strippedfmt = @numformat.gsub /\?/, '0'
    prefix, decimals, sep, floats, expo, postfix=/(^[^\#0e].?)?([\#0]*)?(.)?([\#0]*)?(e.?)?(.?[^\#0e]$)?/i.match(strippedfmt).captures
    ostring.prepend prefix.to_s
    if !decimals.nil? && decimals.size != 0
      if (eval decimals) == nil
        ostring += "##{decimals.size}"
      elsif (eval decimals) == 0
        ostring += decimals.size.to_s
      end
    else
      ostring += decimals
    end
    ostring += sep.to_s
    if !floats.nil? && floats.size != 0
      ostring += ((floats.size.to_s) +"f")
    end
    if sep.nil? && floats.nil? || floats.size == 0
      ostring += "d"
    end
    ostring += (expo.to_s + postfix.to_s)
    return sprintf ostring, val
  end

  # Return
  def datetime val
    deminutified = @numformat.gsub(/(?<hrs>H|h)(?<div>.)m/, '\k<hrs>\k<div>i')
                     .gsub(/im/, 'ii')
                     .gsub(/m(?<div>.)(?<secs>s)/, 'i\k<div>\k<secs>')
                     .gsub(/mi/, 'ii')
    return val.strftime(deminutified.gsub(/[yMmDdHhSsi]/, @dtmap))
  end
end