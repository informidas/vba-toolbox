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
   
    
    'find the last row in the dataset
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        
    Set wks = Nothing
    
    wbk.Close
     
 Next vrtSelectedItem
 
 Else
 'The user pressed Cancel.
 
 End If
 End With
</code></pre>
