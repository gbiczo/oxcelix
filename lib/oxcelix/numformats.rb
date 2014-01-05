module Oxcelix
# The Numformats module provides helper methods that either return the Cell object's raw @value as a ruby value
# (e.g. Numeric, DateTime, String) or formats it according to the excel _numformat_ string (#Cell.numformat).
    module Numformats
      Dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', 'h'=>'%-k',\
                    'ss'=>'%-S', 's'=>'%S', 'mmmmm'=>'%b', 'mmmm'=>'%B', 'mmm'=>'%b', 'mm'=>'%m', \
                    'm'=>'%-m', 'dddd'=>'%A', 'ddd'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y', \
                    'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 'ss'=>'%-S', 's'=>'%S'}

      # Convert the temporary format array (the collection of non-default number formatting strings defined in the excel sheet in use)
      # to a series of hashes containing an id, an excel format string, a converted format string and an object class the format is
      # interpreted on.
      def add_custom_formats fmtary
        fmtary.each do |x|
          if x[:formatCode] =~ /[#0%\?]/
            ostring = numeric x[:formatCode]
            if x[:formatCode] =~ /\//
              cls = 'rational'
            else
              cls = 'numeric'
            end
          elsif x[:formatCode].downcase =~ /[dmysh]/
            ostring = datetime x[:formatCode]
            cls = 'date'
          elsif x[:formatCode].downcase == "general"
            ostring = nil
            cls = 'string'
          end
          Formatarray << {:id => x[:numFmtId].to_s, :xl => x[:formatCode].to_s, :ostring => ostring, :cls => cls}
        end
      end
      
      # Convert the excel-style number format to a ruby #Kernel::Format string and return that String.
      # The conversion is internally done by regexp'ing 7 groups: prefix, decimals, separator, floats, exponential (E+)
      # and postfix. Rational numbers ar not handled yet.
      # @param [String] val an Excel number format string.
      # @return [String] a rubyish Kernel::Format string. 
      def numeric val
        ostring = "%"
        strippedfmt = val.gsub(/\?/, '0').gsub(',','')
        prefix, decimals, sep, floats, expo, postfix=/(^[^\#0e].?)?([\#0]*)?(\.)?([\#0]*)?(e.?)?(.?[^\#0e]$)?/i.match(strippedfmt).captures
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
        if !floats.nil? && floats.size != 0 # expo!!!
          ostring += ((floats.size.to_s) +"f")
        end
        if sep.nil? && floats.nil? || floats.size == 0
          ostring += "d"
        end
        ostring += (expo.to_s + postfix.to_s) #postfix '+' ?
        return ostring
      end

      # Convert excel-style date formats into ruby DateTime strftime format strings
      # @param [String] val an Excel number format string.
      # @return [String] a DateTime::strftime format string.
      def datetime formatcode
        deminutified = formatcode.downcase.gsub(/(?<hrs>H|h)(?<div>.)m/, '\k<hrs>\k<div>i')
                         .gsub(/im/, 'ii')
                         .gsub(/m(?<div>.)(?<secs>s)/, 'i\k<div>\k<secs>')
                         .gsub(/mi/, 'ii')
        return deminutified.gsub(/[yMmDdHhSsi]*/, Dtmap)
      end
    end
    
    # The Numberhelper module implements methods that return the formatted value or the value converted into a Ruby type (DateTime, Numeric, etc)
    module Numberhelper
      include Numformats
      # Get the cell's value and excel format string and return a string, a ruby Numeric or a DateTime object accordingly
      # @return [Object] A ruby object that holds and represents the value stored in the cell. Conversion is based on cell formatting.
      # @example Get the value of a cell:
      #   c = w.sheets[0]["B3"] # => <Oxcelix::Cell:0x00000002a5b368 @xlcoords="A3", @style="84", @type="n", @value="41155", @numformat=14>
      #   c.to_ru # => <DateTime: 2012-09-03T00:00:00+00:00 ((2456174j,0s,0n),+0s,2299161j)>
      #
      def to_ru
        if !@value.numeric? || Numformats::Formatarray[@numformat.to_i][:xl] == nil || Numformats::Formatarray[@numformat.to_i][:xl].downcase == "general"
          return @value
        end
        if Numformats::Formatarray[@numformat.to_i][:cls] == 'date'
          return DateTime.new(1899, 12, 30) + (eval @value)
        else Numformats::Formatarray[@numformat.to_i][:cls] == 'numeric' || Numformats::Formatarray[@numformat.to_i][:cls] == 'rational'
            return eval @value rescue @value
        end
      end

      # Get the cell's value, convert it with to_ru and finally, format it based on the value's type.
      # @return [String] Value gets formatted depending on its class. If it is a DateTime, the #DateTime.strftime method is used,
      # if it holds a number, the #Kernel::sprintf is run.
      # @example Get the formatted value of a cell:
      #   c = w.sheets[0]["B3"] # => <Oxcelix::Cell:0x00000002a5b368 @xlcoords="A3", @style="84", @type="n", @value="41155", @numformat=14>
      #   c.to_fmt # => "3/9/2012"
      #
      def to_fmt
        begin
          if Numformats::Formatarray[@numformat][:cls] == 'date'
              self.to_ru.strftime(Numformats::Formatarray[@numformat][:ostring]) rescue @value
          elsif Numformats::Formatarray[@numformat.to_i][:cls] == 'numeric' || Numformats::Formatarray[@numformat.to_i][:cls] == 'rational'
              sprintf(Numformats::Formatarray[@numformat][:ostring], self.to_ru) rescue @value
          else
            return @value
          end
        end
      end
    end
  end
