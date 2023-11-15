require_relative './script.rb'

table1 = ExcelTable.new('res/test1.xlsx')
table2 = ExcelTable.new('res/test2.xlsx')
table3 = ExcelTable.new('res/test3.xlsx')

# 2D Array of table data
# p table3.to_2d_array

# Table row by index
# p table3.row(1)

# Each method for table
# table3.each { |row| puts row }

# Direct access to table column
# p table3['Student']
# p table3['Student'][2]

# Set value to table cell
# table3['Student', 2] = 'Dimitri'
# p table3.to_2d_array

# Access to table column by method
# p table3.Student

# Average and Sum of table column
# p table1.X.avg
# p table1.X.sum

# Get row by column and value
# p table3.Student.Ana

# Map, Select and Reduce methods
p table1.X.map { |x| x * 2}
p table1.X.select { |x| x > 2 }
p table1.X.reduce(0) { |sum, x| sum + x }

# Table union
# arr = table1 + table2
# p arr

# Table subtraction
# arr = table1 - table2
# p arr