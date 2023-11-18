# Ruby Excel Reader

## Introduction
This Ruby script is designed to work with Excel documents, focusing on metaprogramming principles. The library allows users to interact with tables, assuming each table has a header row and an optional last row that can serve as a sum row.

## Features
1. Retrieving Table Values <br>
```table.to_2d_array```
2. Accessing Rows <br>
```table.row(1)```
3. Implemented an each function to iterate over all cells in the table from left to right. <br>
```table.each { |row| puts row }```
4. Handling Merged Fields <br>
Merged fields are automatically handled by the library.
5. Enhanced Syntax for Accessing and Setting Values <br>
```table['Name']``` <br>
```table['Name'][2]``` <br>
```table['Name', 2] = 'Value'```  <br>
6. Direct Column Access <br>
```table.Name```
7. Subtotal/Average Calculation <br>
```table.column.sum``` <br>
```table.column.avg```  <br>
8. Retrieving Row By Cell <br>
```table.Name.value```
9. Column Functions (map, select, reduce) <br>
```table.column.map { |x| x * 2}```
<br>```table.column.select { |x| x > 2 }```
<br>```table.column.reduce(0) { |sum, x| sum + x }```  <br>
10. Table Addition and Subtraction <br>
```table1 + table2``` <br>
```table1 - table2``` <br>
11. Ignoring Rows with Keywords <br>
Rows containing "total" or "subtotal" are automatically ignored.
