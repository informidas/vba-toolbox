'This provides the construct for an Input Box that prompts and accepts input from the user
'Input boxes are used with variables to capture input from a user and store in a variable for future use.

Sub SayHello():
  dim first_name as string

  first_name = inputbox("Please enter your first name")

  msgbox("Hello " + first_name)

End Sub
