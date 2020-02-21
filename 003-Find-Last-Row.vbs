'locating the last row that has data in a column A

'using the range method
lastRow = range("A" & Rows.Count).End(xlUp).Row


'using the cells method
lastRow = cells(Rows.Count, 1).End(xlUp).Row



