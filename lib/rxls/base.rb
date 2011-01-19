require 'win32ole'
module Rxls
  class Base
    
    def initialize file
      @obj = begin 
        WIN32OLE.connect('excel.application')
      rescue
        WIN32OLE.new('excel.application')
      end
      file = File.expand_path "../../../spec/fixtures/#{file}.xls", __FILE__ if file.kind_of? Symbol
      @obj.workbooks.open file
    end
    
    def quit
      @obj.quit
    end
    
    def worksheet(n,mapping=nil)
      Worksheet.new(@obj,n,mapping)
    end
    
  end
  
  class Worksheet
    include Enumerable
    
    # 'a'   => [0]
    # 'a b' => [0,1]
    def to_row_numbers(s)
      s.split(' ').map do |r|
        r = 96.chr + r if r.length == 1
        s = nil
        r.each_byte do |b|
          s = s ? s + (b-97) : (b-96) * 26
        end
        s
      end
    end
    
    def initialize(o,n,mapping)
      @columns = to_row_numbers mapping
      @values = o.Worksheets(n).UsedRange.Value
      headers = @values.shift # skip 1st row
      @headers = @columns.map{|c| headers[c]} 
    end
    
    def each
      @values.each_with_index do |row,column|
        values = @columns ? @columns.map{|r| row[r]} : row 
        yield values.unshift(column) 
      end
    end
    
    def dump
      vcs = @columns.map{|column| ValueCounter.new}
      each do |row_with_number|
        row_with_number[1..-1].each_with_index do |value,index|
          vcs[index] << dump_value(value)
        end
      end 
      vcs.each_with_index{|vc,index| puts "#{@headers[index]}:#{vc}"}
    end
    
    def dump_value(v)
      case v
      when String
        v == '' ? "String(empty)" : "String"
      else
        v.class
      end
    end
    
  end
  
end
