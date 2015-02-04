require './spec/spec_helper.rb'

describe "Oxcelix module" do
  describe 'Workbook' do
    context 'normal' do
      it "should open the excel file and return a Workbook object" do
        file = 'spec/fixtures/test.xlsx'
        w=Oxcelix::Workbook.new(file)
        w.sheets.size.should    == 2
        w.sheets[0].name.should == "Testsheet1"
        w.sheets[1].name.should == "Testsheet2"
      end
    end
    context 'with excluded sheets' do
      it "should open the sheets not excluded of the excel file" do
        file = 'spec/fixtures/test.xlsx'
        w=Oxcelix::Workbook.new(file, {:exclude=>['Testsheet2']})
        w.sheets.size.should    == 1
        w.sheets[0].name.should == "Testsheet1"
        end
    end
    context 'with included sheets' do
      it "should open only the included sheets of the excel file" do
        file = 'spec/fixtures/test.xlsx'
        w    = Oxcelix::Workbook.new(file, {:include=>['Testsheet2']})
        w.sheets.size.should    == 1
        w.sheets[0].name.should == "Testsheet2"
      end
    end
  end
end
