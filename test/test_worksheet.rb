require 'helper'
module Rxls
  describe Worksheet do

    before(:each) do
      @b = Base.new(:x)
      @w = @b.worksheet 1,'a c'
    end
    
    it 'converts exel colum names to array indexes' do
      assert_equal [0]  , @w.to_row_numbers('a')
      assert_equal [0,1], @w.to_row_numbers('a b')
      assert_equal [26] , @w.to_row_numbers('aa')
    end
    
    it 'iterates over worksheet' do
      assert_equal %w(Behresn Schnell), @w.map{ |row,number,nachname| nachname}
    end
    
#    after(:all) {      @b.quit   }
    
  end
  
end
