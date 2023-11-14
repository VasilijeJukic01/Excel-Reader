require 'roo'
require 'spreadsheet'

$VERBOSE = nil

class ExcelTable
  attr_reader :headers, :xls

  def initialize(path)
    @xls = Roo::Spreadsheet.open(path)
    @headers = find_headers
  end

  def find_headers
    headers_row = 1
    headers_row += 1 while @xls.row(headers_row).compact.empty?
    @xls.row(headers_row)
  end

  def to_2d_array
    data = []
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      data << row_data unless contains_total_or_subtotal?(row_data)
    end
    data
  end

  def find_first_non_empty_row
    row_index = 1
    row_index += 1 while @xls.row(row_index).compact.empty?
    row_index
  end

  def first_non_empty_column
    col_index = 1
    while @xls.column(col_index).compact.empty?
      col_index += 1
    end
    col_index
  end

  def row(index)
    data = @xls.row(find_first_non_empty_row + index)
    return nil if contains_total_or_subtotal?(data)
    Hash[@headers.zip(data)]
  end

  def each
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      data = @xls.row(row_index)
      yield Hash[@headers.zip(data)] unless contains_total_or_subtotal?(data)
    end
  end

  def [](column_name)
    column_index = @headers.index(column_name)
    return nil if column_index.nil?

    data = (find_first_non_empty_row..@xls.last_row).map do |row_index|
      @xls.row(row_index)[column_index]
    end

    { column_name => data }
  end

  def []=(column_name, row_index, value)
    column_index = @headers.index(column_name)
    return nil if column_index.nil?

    row_start = find_first_non_empty_row
    target_row_index = row_start + row_index

    @xls.set(target_row_index, column_index + first_non_empty_column, value)

    self
  end

  def method_missing(method_name, *args)
    col_name = method_name.to_s
    if @headers.include?(col_name)
      col_index = @headers.index(col_name)
      @xls.column(col_index + first_non_empty_column)[1..-1]
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
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      new_table_data << row_data unless contains_total_or_subtotal?(row_data)
    end
    ((other_table.find_first_non_empty_row+1)..other_table.xls.last_row).each do |row_index|
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
    other_table_index = other_table.find_first_non_empty_row
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      next if contains_total_or_subtotal?(row_data) || other_table.to_2d_array[other_table_index+1] === row_data
      other_table_index += 1
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