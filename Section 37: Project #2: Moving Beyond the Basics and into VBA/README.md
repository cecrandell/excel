# Section 37: Project #2: Moving Beyond the Basics and into VBA

<!-- ## Introduction to Project #2: Interacting with the User

## Beyond the Basics Sort Project Exercise Files (DOWNLOAD) -->

## Project #2: Introduction to the Excel VBA Range.Sort Method

- [SORT function](https://support.microsoft.com/en-au/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c)

## Creating the Excel VBA Sort Procedures for the Project

- `:=` is the assignment operator

<img src="Images/1.png" width="800" />

- Copy and paste the procedures and customize for division, category, and total

<img src="Images/2.png" width="800" />

## Project #2: Prompting the User for Information

- [InputBox function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function)

- `userInput = InputBox("What is your favorite color?", "Favorite Color")` - asks the user for input

## Continue Excel VBA InputBox

- `vbcrlf` is carriage return-linefeed combination

- [Miscellaneous constants](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/miscellaneous-constants)

<img src="Images/3.png" width="800" />
<img src="Images/4.png" width="800" />

## Project #2: Building Logic into Your Macros

<img src="Images/5.png" width="800" />

## Project #2: Alerting the User of Errors

- [MsgBox return values](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-constants)

- `vbYes - 6 - Yes button pressed`

<img src="Images/6.png" width="800" />
<img src="Images/7.png" width="800" />

## Using Excel VBA Error Control Statements

- What if the user does not input anything?

- `On Error GoTo errHandler`

- [](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)

- Example:

```
Sub InitializeMatrix(Var1, Var2, Var3, Var4)
 On Error GoTo ErrorHandler
 . . .
 Exit Sub
ErrorHandler:
 . . .
 Resume Next
End Sub
```

<img src="Images/8.png" width="800" />

## Create a Button to Run the Sort Procedure and Save

<img src="Images/9.png" width="800" />

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
