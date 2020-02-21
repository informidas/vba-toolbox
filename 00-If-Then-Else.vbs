' using the If-Then-Else statement for conditional testing
' once a test resolves to True in VBA, then none of the other tests will run
' i.e. it only runs that block of code that resolves to True


If (first test) Then

  'do these things if test is True

ElseIf (second test) Then

  'do these things if test is True

ElseIf (next test...) Then

  'do these things if test is True

Else: 

  'do these things if none of teh above test were True

End If


' Example:

Sub DeclareGender(sGender as string):

    If sGender = 'M' Then 
      MsgBox "Male"
    
    ElseIf sGender = "F" Then 
      MsgBox "Female"

    ElseIf sGender = "T" Then 
      MsgBox "Transgender"

    Else: 
      MsgBox "Other"

    End If


End Sub
