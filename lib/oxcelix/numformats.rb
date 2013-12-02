module Oxcelix
# The Numformats module provides helper methods that either return the Cell object's raw @value as a ruby value
# (e.g. Numeric, DateTime, String) or formats it according to the excel _numformat_ string (#Cell.numformat).
    module Numformats
      Dtmap = {'hh'=>'%H', 'ii'=>'%M', 'i'=>'%-M', 'H'=>'%-k', 'h'=>'%-k',\
                    'ss'=>'%-S', 's'=>'%S', 'mmmmm'=>'%b', 'mmmm'=>'%B', 'mmm'=>'%b', 'mm'=>'%m', \
                    'm'=>'%-m', 'dddd'=>'%A', 'ddd'=>'%a', 'dd'=>'%d', 'd'=>'%-d', 'yyyy'=>'%Y', \
                    'yy'=>'%y', 'AM/PM'=>'%p', 'A/P'=>'%p', '.0'=>'', 'ss'=>'%-S', 's'=>'%S'}

      # Convert the temporary format array to a series of hashes containing an id, an excel format string, a converted
      # format string and an object class the format is interpreted on.
      def add fmtary
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
          Numformats::Formatarray << {:id => x[:numFmtId].to_s, :xl => x[:formatCode].to_s, :ostring => ostring, :cls => cls}
        end
      end
      
      # Convert the excel-style number format to a ruby #Kernel::Format string and return that String.
      # The conversion is internally done by regexp'ing 7 groups: prefix, decimals, separator, floats, exponential (E+)
      # and postfix. Rational numbers ar not handled yet.
      # @param [String] val an Excel number format string.
      # @return [String] a rubyish Kernel::Format string. 
      def numeric val
        ostring = "%"
        strippedfmt = @numformat.gsub(/\?/, '0').gsub(',','')
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
        return deminutified.gsub(/[yMmDdHhSsi]*/, Numformats::Dtmap)
      end
    end
    
    module Numberhelper
      include Numformats
      # Get the cell's value and excel format string and return a string, a ruby Numeric or a DateTime object accordingly
      def to_ru cell
        if @numformats[cell.numformat.to_i][:xl] == nil || @numformats[cell.numformat.to_i][:xl].downcase == "general"
          return cell.value
        end
        if @numformats[cell.numformat.to_i][:cls] == 'numeric' || @numformats[cell.numformat.to_i][:cls] == 'rational'
          return eval cell.value
        elsif @numformats[cell.numformat.to_i][:cls] == 'date'
          if (0.0..1.0).include? cell.value
            return DateTime.new(1900, 01, 01) + (eval cell.value)
          else
            return DateTime.new(1899, 12, 31) + (eval cell.value)
          end
        else
          eval cell.value rescue cell.value
        end
      end

      # Get the cell's value, convert it with to_ru and finally, format it based on the value's type.
      def to_fmt
        if @numformats[@numberformat.to_i][:cls] == 'date'
          datetime to_ru
        elsif @numformats[@numberformat.to_i][:cls] == 'numeric' || @numformats[@numberformat.to_i][:cls] == 'rational'
          numeric to_ru
        else 
          return @value.to_s
        end
      end

      private
    end
  end