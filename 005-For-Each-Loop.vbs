'create a loop structure to loop through all sheets in a workbook
'the next ws command instructs the program to move to the next sheet after the code for the sheet has executed

'declaration not strictly required - as the loop construct will implicitly declare _
 but considered a good practice 
dim ws as excel.worksheet


for each ws in thisworkbook.worksheets
  
  'code for changes you want to make on each sheet goes here


next ws 


