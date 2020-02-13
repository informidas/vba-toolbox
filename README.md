# vba-toolbox
This repo provide users with the essential tools needed to effectively work with VBA.

## Objectives
The sources files are intended to provide code snippest and context for some common use cases when working with VBA to automate tasks. The key objectives are: <br>

* Offer some Best Practices ideas to cultivate good programming habits

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

These are just a few of the more frequently used variables, However vba comes with the ability to define a ton of object types.



#### **Dim** *Statement*
A Dim statement is the vba reserved word used to declare a variable for use. Dim is actually an abbreviation for the work dimension. In VBA we declare a variable as follows: <br>
>
> Dim variable_name as variable_type
>
>

#### **Sub**
A subroutine (or sub) is a procedure or package of vba code consisting of one or more statements that get executed when the sub is run / called. The format for a subroutine is as follows: <br>

> 
> Sub *Name-of-the-Sub*():
>
> End Sub

In between the Sub and the End Sub lines, we write the vba statemenst to be executed. 

### Best Practices


#### Naming Variables and Objects
