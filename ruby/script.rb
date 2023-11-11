require 'roo'
require 'spreadsheet'

class ExcelTable

  def initialize(path)
    @xls = Roo::Spreadsheet.open(path)
    @headers = @xls.row(1)
  end

  def to_2d_array
    data = []
    data << @headers
    (2..@xls.last_row).each do |row_index|
      data << @xls.row(row_index)
    end
    data
  end

  def row(index)
    data = @xls.row(index + 1)
    Hash[@headers.zip(data)]
  end

  def each
    (2..@xls.last_row).each do |row_index|
      data = @xls.row(row_index)
      yield Hash[@headers.zip(data)]
    end
  end

  def [](key)
    if key.is_a? String
      col_index = @headers.index(key)
      return nil if col_index.nil?
      @xls.column(col_index + 1)[1..-1]
    else
      raise ArgumentError, "Invalid key type. Use column name as a String."
    end
  end

  def []=(key, index, value)
    if key.is_a? String
      col_index = @headers.index(key)
      if col_index
        @xls.set(index + 1, col_index + 1, value)
      else
        raise ArgumentError, "Column '#{key}' not found in headers."
      end
    else
      raise ArgumentError, "Invalid key type. Use column name as a String."
    end
  end

  def method_missing(method_name, *args)
    col_name = method_name.to_s
    if @headers.include?(col_name)
      col_index = @headers.index(col_name)
      @xls.column(col_index + 1)[1..-1]
    else
      super
    end
  end

  def respond_to_missing?(method_name, include_private = false)
    col_name = method_name.to_s
    @headers.include?(col_name) || super
  end

  def exclude_total_rows
    @xls.each_with_index do |row, index|
      if row.join.downcase.include?('total') || row.join.downcase.include?('subtotal')
        @xls.delete_row(index + 1)
      end
    end
  end

  def self.union(table1, table2)
    raise ArgumentError, "Header mismatch" unless table1.headers == table2.headers

    new_table = table1.clone
    new_table.instance_variable_set(:@xls, table1.xls + table2.xls[1..-1])
    new_table
  end

  def self.subtract(table1, table2)
    raise ArgumentError, "Header mismatch" unless table1.headers == table2.headers

    new_table = table1.clone
    table2.each do |row|
      new_table.xls.delete_if { |r| r == row.values }
    end
    new_table
  end

  def map(&block)
    result = []
    each do |row|
      result << yield(row)
    end
    result
  end

  def select(&block)
    result = []
    each do |row|
      result << row if yield(row)
    end
    result
  end

  def reduce(initial_value, &block)
    accumulator = initial_value
    each do |row|
      accumulator = yield(accumulator, row)
    end
    accumulator
  end

end

class Array
  def avg
    sum / size.to_f
  end
end