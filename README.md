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

#### **Comments**
Comments are used in VBA to give hints, guidance and documentation for what our code is doing at different steps along the way. Comments are extremely useful when other users are trying to read our code, as well as in helping the author revisit their code some time in the future. To create a comment in VBA, we type a single apostrophe ('), followed by the actual comment. For example: <br>

> 'This is my first comment
> 'This is my second comment

Anything we type to the right of the *" ' "* symbol, the VBA interpreter will treat as a comment

#### **Dim** *Statement*
A Dim statement is the vba reserved word used to declare a variable for use. Dim is actually an abbreviation for the work dimension. In VBA we declare a variable as follows: <br>

>
> Dim variable_name as variable_type
>

Here are some examples of variable declarations: <br>

> Dim first_name as string
> Dim age as integer
> Dim is_adult as boolean
> Dim dob as date


Variables are typically declared inside of a Procedure or Function

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

In between the *Function* and the *End Function* lines, we write the vba statemenst to be executed.

 Here, *return type* means the type returned by the Function. Let's look at a concrete example: <br>
  Lets say we wanted 

 

### Best Practices

#### Naming Variables and Objects
