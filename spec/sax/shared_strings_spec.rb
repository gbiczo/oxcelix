require './spec/spec_helper'
describe Oxcelix::Sharedstrings do
  describe 'number of elements in shared_strings' do

    def values_from_oga
      require 'oga'
      doc            = Oga.parse_xml(File.read('spec/fixtures/shared_strings.xml'))
      correct_values = doc.xpath('*[local-name() = "sst"]/*[local-name() = "si"]/*[local-name() = "t"]').map{|x| x.text}
    end

    it "counts also quasi-empty elements" do
      strings = Oxcelix::Sharedstrings.new
      File.open('spec/fixtures/shared_strings.xml', 'r') do |f|
        Ox.sax_parse(strings, f, {})
      end

      strings.stringarray.size.should == values_from_oga.size
    end
  end
end
