require './spec/spec_helper.rb'

describe "Oxcelix module" do
  describe 'Sheet' do
    context 'format' do
      it "should not raise a syntax error on a Text field with a ZIP code" do
        file = 'spec/fixtures/test.xlsx'
        w=Oxcelix::Workbook.new(file)
        lambda { w.sheets[0].to_fmt }.should_not raise_error
      end
    end
  end
end
