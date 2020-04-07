# vba-toolbox
This is an introductory guide to Visual Basics for Applications (VBA). At its core, VBA is a scripting languauge that provides users with the ability to control the Microsoft Office envirnment programmatically.

Like most programming languages, VBA has a number of programming constructs that help to to extend its power and flexibility. Additionally it offers intelli-sense through the built-in editor within the office suite (invoked by *alt + F11* on windows and *fn + cmd + F11* on Mac). This aides the learning curve for beginners using VBA as a first language foray into the world of programming.

This guide will offer a gentle introduction to VBA and its constructs. It is intended to provide readers with a solid foundation upon which to explore and extend their skills.This repo provide users with the essential tools needed to effectively work with VBA.

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

<pre><code>
$ Dim variable_name as variable_type
$
</code></pre>
Here are some examples of variable declarations <br>
<pre><code>
 Dim first_name as string <br>
 Dim age as integer <br>
 Dim is_adult as boolean <br>
 Dim dob as date <br>
</code></pre>

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

### Invoking the VBA Editor
The Microsoft Office suite provides a built-in code editor that allows users to code instructions relevant to the application. There are several ways to invoke your VBA Editor, depending on your operating system.<br>
For Windows
- Option 1: While holding the **Alt** key, press the function **F11** key (**Alt + F11**)
- Option 2: Select the Developer tab, then press the Visual Basic button to activate the VBA Editor

For MAC
- Option 1: While holding the **fn** + **Alt** keys, press the function **F11** key (**Alt + F11**)
- Option 2: Select the Developer tab, then press the Visual Basic button to activate the VBA Editor

hint: The Developer Menu is not viosible by default. if you need to activate it, do the following:
  From the menu choices choose *File* - *Options* - *Customize Ribbon* , then check Developer Option from the column on right. 

Below is a sample of the VBA Editor window that appears when we use one of the above Options to launch the editor.
![VBA IDE](https://github.com/informidas/vba-basic-documentation/blob/master/VBA_IDE.PNG "sample VBA Editor screen")


---

Here a few key considerations before we begin our sample coding:
- To begin coding, we will enter our instructions in the blank area below general.
- These instructions that we enter are called statements or commands.
- Each new statement / command is placed on a separate line
- to end and instruction we simply press the Enter / return key
  
### VBA Constructs
VBA provides some the most useful programming constructs, many of which can be found in other popular programming language. 

##### Declaring a Variable
In order to use a variable in VBA we define it as follows:
*Dim variableName as variableType*
Some of the most frequently used Variable Types are:

* string
* integer
* long
* double
* single
* boolean
* array
* date
* decimal
* byte
* currency

Using some of the popular datatypes you could define variables as follows: <br>

> Dim fullname as string <br>
> Dim age as integer <br>
> Dim salary as currency <br>
> Dim DOB as date <br>
> Dim hasDegree as boolean <br>
> 

##### Generating Comments
It is a good coding practice to include comments in your code. Comments provide a way for others reviewing your code to understand the intent of each statement in particular and your program in general. We declare a line of comment using a single apostrophe (')

Here is how you can declare a comment:
>
> ' This is a comment
>
> ' This is a second comment
>

#### Printing Messages to the screen
An important part of programming is printing messages to the screen to interact with users. In VBA, we print messages to the screen using message boxes. To generate a message box, type the following:

>
> msgbox("your message goes here between the quotes")
>
>

#### Objects, Methods and Properties
Another important concept to remember when programming in VBA is that everything is based on a hierarchy of objects. The hierarchy for Microsoft Excel is as follows:
Excel *Application -> Workbook > worksheet > columns and rows > cells and ranges*


#### Cells and Ranges
When using VBA to add data to a sheet, we use the range or cell objects to manipulate rows and columns on the Excel spreasheet.
Ranges are defined by a the keyword **range** followed by an open parenthesis, followed by a cell reference of a letter and a number, followed by a closing parenthesis. <br>
Cells on the other hand, use a row and column reference in the form cells(row number, column number). <br>
For example if we needed to reference the cell C4 we would type: <br>
*cells(4, 3)* <br>
since we want the 4th row of the 3rd column.<br> 

Below are examples of using the range and cell options for adding a heading **Product** in cell A1 we type:

>
> *range("A1").value = "Product"* <br>
> *cells(1,1).value = "Product"*
>

#### Loops and Iterators
Loops are useful VBA constructs when we need to iterate through a list or collection of items. While there are a number of Loop constructs we will focus on using the For Loop construct.<br>
 
##### Using a For Loop
The structure of a For Loop is as follows: <br>
>
> For i = x to y 
>   *do some step 1*  <br>
>   *do step 2*  <br>
>   *do step 3 etc*  
> Next i
- *i* is considered an iterator <br>
- *x* is considered the lower boundary or where the loop will begin from  <br>
- *y* is considered the upper boundary or where the loop will stop <br>

Let's use an example to illustrate. <br>

Suppose we needed to add the state abbreviation in column C, based on the data in column B how could we do this using a loop? <br>

![State_Abbreviation](https://github.com/informidas/vba-basic-documentation/blob/master/State_Abbreviation.PNG "table used in For Loop example") <br>
>
> Sub AddStateAbbrev()
>    Dim i As Integer
>    
>    For i = 2 To 8
>    
>        If Cells(i, 2).Value = "New Jersey" Then
>            Cells(i, 3).Value = "NJ"
>            
>        ElseIf Cells(i, 2).Value = "New York" Then
>            Cells(i, 3).Value = "NY"
>            
>        ElseIf Cells(i, 2).Value = "Connecticut" Then
>            Cells(i, 3).Value = "CT"
>            
>        End If
>    
>    Next i
>    
> End Sub
>


#### Arrays

- An array is a collection of elements. 
- Elements in the array can be of similar or varying types
- Each element in an array can be access using an index (I.e. each element is associated with an index number)
- In VBA, Arrays are zero based index - meaning the numbering for elements in an index begin with 0.

**Example**
Let's say we wanted to use an Array to store the days of a week, we would do the following:

>
> ' declare the array
>
> Dim DaysOfWeek(6) as string
> 
> ' Assign the days of week to each element in the array:
>
> DaysOfWeek(0) = "Mon"
> DaysOfWeek(1) = "Tue"
> DaysOfWeek(2) =  "wed"
> DaysOfWeek(3) = "Thu"
> DaysOfWeek(4) = "Fri"
> DaysOfWeek(5) = "Sat"
> DaysOfWeek(6) = "Sun"

Alternately, all assignments could be performed in a single statement.

>
> DaysOfWeek = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
>

Once an array has been declared and assigned values we can now reference in the remainder of Our subroutine or program.
We reference elements as follows:
>
> Msgbox("The first day of the week is: " + DaysOfWeek(0) )
>

#### Subroutines
A Subroutine is a block of code (i.e. series of vba statements or commands). This subroutine when executed will run all statements in the block.

Creating a subroutine begins with the keyword *Sub* and ends with the keywords *End Sub* . Below is an example of a subroutine.

>
##### Declaring a Subroutine
> Sub HelloWorld()
>    msgbox "Hello World!"
> End Sub

The real power of a subroutine is in its ability to take input parameter (known as an *argument*) and output a value or message (known as a *return value*) to the user. Using the HelloWord() subroutine example, we could modify the subroutine as follows: <br /><br />

>   ' In this example we modify the **HelloWorld()** subroutine to accept an argument labeled as *first_name*. <br /> We defined the *first_name* parameter as a string
>
>   Sub HelloWorld(first_name as string) <br />
>      msgbox "Hello " + sName + "!"  <br/>
>   End Sub <br />
>
