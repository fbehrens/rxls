require 'spec_helper'
module Rxls
  describe Base do

    before(:all) do
      @b = Base.new(:x)
      @w = @b.worksheet 1,'a c'
    end
    
    it 'converts exel colum names to array indexes' do
      @w.to_row_numbers('a').should == [0]
      @w.to_row_numbers('a b').should == [0,1]
      @w.to_row_numbers('aa').should == [26]
    end
    
    it 'iterates over worksheet' do
      @w.map{ |row,number,nachname| nachname}.should == ["Behresn", "Schnell"]
    end
    
    after(:all) do
      @b.quit
    end
    
  end
  
end
