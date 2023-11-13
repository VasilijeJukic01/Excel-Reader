require 'roo'
require 'spreadsheet'

$VERBOSE = nil

class ExcelTable
  attr_reader :headers, :xls

  def initialize(path)
    @xls = Roo::Spreadsheet.open(path)
    @headers = @xls.row(1)
  end

  def to_2d_array
    data = []
    data << @headers
    (2..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      data << row_data unless contains_total_or_subtotal?(row_data)
    end
    data
  end

  def row(index)
    data = @xls.row(index + 1)
    return nil if contains_total_or_subtotal?(data)
    Hash[@headers.zip(data)]
  end

  def each
    (2..@xls.last_row).each do |row_index|
      data = @xls.row(row_index)
      yield Hash[@headers.zip(data)] unless contains_total_or_subtotal?(data)
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

  def contains_total_or_subtotal?(row_data)
    row_data.any? { |cell_value| cell_value.to_s.downcase.include?("total") || cell_value.to_s.downcase.include?("subtotal") }
  end

  def +(other_table)
    unless other_table.headers == @headers
      puts "Headers of the tables are not the same. Tables cannot be unioned."
      return
    end
    new_table_data = []
    (2..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      new_table_data << row_data unless contains_total_or_subtotal?(row_data)
    end
    (2..other_table.xls.last_row).each do |row_index|
      row_data = other_table.xls.row(row_index)
      new_table_data << row_data unless contains_total_or_subtotal?(row_data)
    end
    new_table_data
  end

  def -(other_table)
    unless other_table.headers == @headers
      puts "Headers of the tables are not the same. Tables cannot be subtracted."
      return
    end
    new_table_data = []
    (2..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      next if contains_total_or_subtotal?(row_data) || other_table.to_2d_array[row_index] === row_data
      new_table_data << row_data
    end
    new_table_data
  end

  def map
    result = []
    each do |row|
      result << yield(row)
    end
    result
  end

  def select
    result = []
    each do |row|
      result << row if yield(row)
    end
    result
  end

  def reduce(initial_value)
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

  def method_missing(method_name, *args)
    data = method_name.to_s
    if self.include?(data)
      self.index(data)
    else
      super
    end
  end
end