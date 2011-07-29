module Rxls
  
  class Base
    
    def initialize file
      @obj = begin 
        WIN32OLE.connect('excel.application')
      rescue
        WIN32OLE.new('excel.application')
      end
      file = File.expand_path "../../../test/fixtures/#{file}.xls", __FILE__ if file.kind_of? Symbol
      @obj.workbooks.open file
    end
    
    def quit
      @obj.quit
    end
    
    def worksheet(n,mapping=nil)
      Worksheet.new(@obj,n,mapping)
    end
    
  end
    
end
