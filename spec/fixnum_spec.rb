require "rspec"
require_relative './spec_helper.rb'
require_relative '../lib/oxcelix.rb'

 describe "Fixnum object" do
  describe '#col_name' do
    it "returns a string representing an excel column name" do
      (0..25).each do |x|
        x.col_name.should == ('A'..'Z').to_a[x]
      end
    end
  end
end
