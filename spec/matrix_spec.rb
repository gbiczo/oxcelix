require "rspec"
require_relative '../lib/oxcelix.rb'

 describe "Matrix object" do
  describe '#[]=' do
    it "should set a cell to a new value" do
     m_obj=Matrix.build(4, 4){nil}
     m_obj[3,3]='foo'
     m_obj[3,3].should == 'foo'
   end
 end
end
