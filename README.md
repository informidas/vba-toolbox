# vba-toolbox
This repo provide users with the essential tools needed to effectively work with VBA.

## Objectives
The sources files are intended to provide code snippest and context for some common use cases when working with VBA to automate tasks. The key objectives are: <br>

* Offer some *Best Practice* ideas to cultivate good programming habits

* Provide some code snippest that can be reused from project to project

* Methods for Troubleshooting some frequent errors

* Built-in Functions and constructs that are helpful during development of your program

### Fundamental Concepts

#### **Variable** 
A variable is a container in memory used to store values for manipulation, comparison and display. Variables can be of different types such as: <br>
* *string*
* *number*
* *date*
* *boolean*
* *array*
* *variant* - which is a catch all variable type that can be of any type. More on this later

These are just a few of the more frequently used variables, However vba comes with the ability to define many other object types.

**When creating variable here a few things to keep in mind**: <br>

* Variable names are not case sensitive (so FIRSTNAME and firstname are the same).

* It is good practice to keep your variables in a consistent case / format (i.e. choose a convention and stick to it!).

* A popular convention for variable names is camel case (i.e. first word lower case and additional words in the name are title case): <br>
e.g. of camel case: firstName, lastName, phoneNumber, graduationDate

* Another popular convention for variable names is snake case (i.e. each word is joined by an underscore) <br>
e.g. of snake case: first_name, last_name, phone_number, graduation_date

* Variable names should be descriptive without being too long. For example it will be easier for you to understand what the variable first_name is used for than a variable defined as f

**Declaring Variables in VBA** - using *Dim* keyword <br>
*Dim* is the vba reserved word used to declare a variable. Dim is actually an abbreviation for the word *dimension*. In VBA we declare a variable as follows: <br>

>
> Dim variable_name as variable_type
>

Here are some examples of variable declarations <br>
>
> Dim first_name as string <br>
> Dim age as integer <br>
> Dim is_adult as boolean <br>
> Dim dob as date <br>
>

Variables are typically declared inside of a Procedure or Function

#### **Comments**
Comments are used in VBA to give hints, guidance and documentation for what our code is doing at different steps along the way. Comments are extremely useful when other users are trying to read our code, as well as in helping the author revisit their code some time in the future. To create a comment in VBA, we type a single apostrophe ('), followed by the actual comment. For example: <br>

> 'This is my first comment
> 'This is my second comment

Anything we type to the right of the *" ' "* symbol, the VBA interpreter will treat as a comment

##### **Printing Messages** - using the *MsgBox* keyword
In VBA, we use the *MsgBox* built-in function to print / output a message to the screen. It takes one mandatory input parameter for a message string, as well as a number of other optional parameters 



#### **Sub**
A subroutine (or sub) is a procedure or package of vba code consisting of one or more statements that get executed when the sub is run / called. The format for a subroutine is as follows: <br>

> 
> *Sub Name-of-the-Sub():*
>
> *End Sub*

In between the *Sub* and the *End Sub* lines, we write the vba statements to be executed.

#### **Function**
A Function is similar to a subroutine in structure. Typically a function is used when we would like to return a value as the direct output of the program, which can then be used by another subroutine or function. The format for declaring a Function is:

>
> *Function Name-of-Function() as return type* 
>
> *End Function* 

In between the *Function* and the *End Function* lines, we write the vba statements to be executed.

 Here, *return type* means the type returned by the Function. Let's look at a concrete example: <br>
 
#### **Arguments**
Arguments are parameters that are passed into Subroutines and Functions as variables. This helps to make code dynamic, since we can change the values we pass in as inputs. 

 **Example**
 Lets say we wanted a subroutine that will accept one input from a user: <br>
 1. first name
 
 We would like the subroutine to print a hello message using the person's first name. <br>
 So if I type *John* for the first name the program should print: <br>
 *Hello John*. How would we do this? Here is a possible solution using a subroutine<br>
 
 >Sub PrintName(first_name as string)
 >
 >'print the hello message <br>
 >msgbox "Hello " + first_name
 >
 >End Sub

 If we wanted to use a Function to solve the same challenge we could do the following: <br>
 
 > Function PrintName(first_name as string) as string
 >    PrintName = "Hello " + first_name
 >  return PrintName
 > End Function 


### Looping Structures
Frequently when we provide a solution VBA, we need to iterate over a range of  cells, rows and columns. VBA provides Looping structures to help when we need to process a collection of cells, rows, columns and worksheets.

#### For Each Next
The for each next structure can be used when working with a collection / group of worksheets in a workbook. For example, let us say I have 5 sheets in a workbook labeled *sheet1, sheet2, sheet3, sheet4 abd sheet5*. If we needed to some kind of processing to all these sheets we could use a loop structure like this: <br>

> For Each ws in Worksheets
>  'do the processing here
> Next ws 

The *Next ws* statement is teh instruction to move to the next sheet



### Best Practices

* Use comments to confirm the logic of your program 

* Use Msgbox to test the value of variables that you have defined and assigned values

* Use a consistent case and convention for variable declarations

* Use descriptive names for variables

* Initialize numeric variables (i.e. those used for counters and loops

#### Naming Variables and Objects
