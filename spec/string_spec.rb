
require './spec/spec_helper'

describe "String object" do
  describe 'numeric?' do
    context "with numbers" do
     it "should return true" do
        (1..100).each do |x|
          x.to_s.numeric?.should == true
        end
      end
    end
    context "with strings" do
      it "should return false" do
        ('a'..'zz').each do |x|
          x.numeric?.should == false
        end
      end
    end
  end
end
