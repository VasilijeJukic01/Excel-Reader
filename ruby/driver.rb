require_relative './script.rb'

table1 = ExcelTable.new('res/test.xlsx')
table2 = ExcelTable.new('res/test2.xlsx')
table3 = ExcelTable.new('res/test3.xlsx')

# table['Y', 2]= 230

# table.each { |row| puts row }

# p table.Y.map { |x| x * 2}

arr = table1 - table2
p arr
# p table1.to_2d_array
# p table3.Student.Ana
