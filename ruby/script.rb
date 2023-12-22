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

  def find_first_non_empty_row
    row_index = 1
    row_index += 1 while @xls.row(row_index).compact.empty?
    row_index
  end

  def first_non_empty_column
    col_index = 1
    col_index += 1 while @xls.column(col_index).compact.empty?
    col_index
  end

  def is_empty_row?(row_data)
    row_data.compact.empty?
  end

  def to_2d_array
    data = []
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      next if is_empty_row?(row_data)
      data << row_data unless contains_total_or_subtotal?(row_data)
    end
    data
  end

  def row(index)
    data = @xls.row(find_first_non_empty_row + index)
    return nil if contains_total_or_subtotal?(data) || is_empty_row?(data)
    Hash[@headers.zip(data)]
  end

  def each
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      data = @xls.row(row_index + 1)
      next if is_empty_row?(data)
      yield Hash[@headers.zip(data)] unless contains_total_or_subtotal?(data)
    end
  end

  def [](column_name)
    column_index = @headers.index(column_name)
    return nil if column_index.nil?
    ((find_first_non_empty_row + 1)..@xls.last_row).map do |row_index|
      row_data = @xls.row(row_index)
      next if is_empty_row?(row_data)
      row_data[column_index]
    end
  end

  def []=(column_name, row_index, value)
    column_index = @headers.index(column_name)
    return nil if column_index.nil?
    @xls.set(row_index + find_first_non_empty_row, column_index + first_non_empty_column, value)
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
    return puts "Headers of the tables are not the same. Tables cannot be added." unless other_table.headers == @headers
    data = []
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      data << row_data unless contains_total_or_subtotal?(row_data)
    end
    ((other_table.find_first_non_empty_row+1)..other_table.xls.last_row).each do |row_index|
      row_data = other_table.xls.row(row_index)
      data << row_data unless contains_total_or_subtotal?(row_data)
    end
    data
  end

  def -(other_table)
    return puts "Headers of the tables are not the same. Tables cannot be subtracted." unless other_table.headers == @headers
    data = []
    other_table_index = other_table.find_first_non_empty_row
    (find_first_non_empty_row..@xls.last_row).each do |row_index|
      row_data = @xls.row(row_index)
      next if contains_total_or_subtotal?(row_data) || other_table.to_2d_array[other_table_index+1] === row_data
      other_table_index += 1
      data << row_data
    end
    data
  end

end

class Array

  def avg
    sum / size.to_f
  end

  def method_missing(method_name, *args)
    data = method_name.to_s
    self.include?(data) ? self.index(data) : super
  end

end