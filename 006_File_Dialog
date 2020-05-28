'Use FileDialog Object to reference a set of files to be processed

<pre><code>
Dim fd As FileDialog
 
 'Create a FileDialog object as a File Picker dialog box.
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
 'Declare a variable to contain the path
 'of each selected item.
 'must be a variant since the For Each structure operates on variants only
 Dim vrtSelectedItem As Variant
 
 'Use a With...End With block to reference the FileDialog object.
 With fd
    
    .InitialFileName = "string/for/path/to/file(s)/for/processing/goes/here"
    
 
 'Use the Show method to display the File Picker dialog box and return the user's action.
 'The user pressed the button.
 If .Show = -1 Then
 
 'Step through each file referenced by the string and process accordingly
 For Each vrtSelectedItem In .SelectedItems
 
    Set wbk = Application.Workbooks.Open(vrtSelectedItem)
        
    
    MsgBox "Now processing: " & wbk.Name
    
    Set wks = wbk.ActiveSheet
    
    MsgBox "now processing sheet: " & wks.Name
    
    Set rng = wks.Range("A1")
       
    
    'find the last column in the dataset
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Cells(2, lastCol + 1).Value = "TableReferenceId"
    
    Cells(2, lastCol + 2).Value = "CensusNotes "
    
    
    'find the last row in the dataset
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    sCensusTableID = Left(Cells(1, lastCol).Value, InStr(1, Cells(1, lastCol).Value, "_", vbTextCompare) - 1)
    
    MsgBox "Census Table ID is: " & sCensusTableID
    
    
    lngCol = lastCol
    
    While lngCol >= 1
        If InStr(1, Cells(2, lngCol).Value, "Error!!", vbTextCompare) > 0 Then
            wks.Cells(2, lngCol).EntireColumn.Delete
        
        End If
        
        lngCol = lngCol - 1
        
    Wend
    
    Stop
    
    Set wks = Nothing
    
    wbk.Close
     
 Next vrtSelectedItem
 'The user pressed Cancel.
 Else
 End If
 End With
</code></pre>
