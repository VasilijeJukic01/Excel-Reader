require_relative './script.rb'

table = ExcelTable.new('test.xlsx')

# p table.to_2d_array

# table['Y', 2]= 230

# table.each { |row| puts row }

p table.Y.map { |x| x * 2}


