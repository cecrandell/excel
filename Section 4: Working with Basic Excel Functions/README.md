# Section 4: Working with Basic Excel Functions

## Excel Function

- "A predefined formula that performs a calculation."

## Building Block of Excel Functions

- 3 Parts of an Excel Function:

`= FUNCTION NAME (ARGUMENTS)` (e.g., `=SUM(B4:B8)`)

- The colon character in Excel indicates "through", like in a range (e.g., `=SUM(B4:B8)`).

## Function Arguments Window

<img src="Images/argumentswindow.png" width="400" />

- To open the Function Arguments Window, you can click on the "fx" icon in the formula bar. This assists with defining the function you're using and what arguments are needed/what they mean.

- On Windows, on the bottom left of the Function Arguments Window, you can click "Help on this function", and it will take you to a Mirosoft webpage with more information.

## Excel Functions

- 461+ functions as of 2013 in Excel.

- Excel functions can be found in the "Formulas" tab.

- Often if Excel removes a function, they'll replace it with a different one.

## SUM()

<img src="Images/sum.png" width="400" />

- To find the `SUM` function, go to the "Formulas" tab and click on the drop down on "Math & Trig" and find `SUM`.

<img src="Images/arguments.png" width="400" />

- Excel will detect which cells you want to `SUM` and predict then populate the range in the "Function Arguments Window". You can cutomize as necessary.

## MIN() and MAX()

<img src="Images/min.png" width="400" />
<img src="Images/max.png" width="400" />

- To find `MIN()` and `MAX()` formulas, you can find them under the "Formulas" tab, then "More Fucntions", then "Statistical".

<img src="Images/minsearch.png" width="400" />
<img src="Images/maxsearch.png" width="400" />

- Or, you can click on the "fx" icon on the formula bar and search for the function you're looking for.

## AVERAGE()

<img src="Images/typeaverage.png" width="400" />

- If you start typing a function in the formula entry section of the formula bar, it will show a drop down menu for you to chose from.

<img src="Images/average.png" width="400" />

- Above is an example of the `AVERAGE()` function in action.

## COUNT()

- The `COUNT()` function only counts numeric values - it does not include empty/blank cells or text values.

<img src="Images/count.png" width="400" />

- Above is an example of the `COUNT()` function in action.

## Adjacent Cells Errors

<img src="Images/adjacent.png" width="400" />

- Sometimes Excel tries to be helpful, and it detects a numeric value (e.g., a date value) above a range of cells (adjacent to the range), and as such, Excel throws an error.

<img src="Images/update.png" width="400" />

- Excel will suggest that you should "Update Formula to Include Cells", thus including the date value in the range.

<img src="Images/inaccurate.png" width="400" />

- This would produce inaccurate results, so we can fix this by clicking on the error and ignoring the error (see below).

<img src="Images/ignore.png" width="400" />

## AutoSum

<img src="Images/autosum.png" width="400" />

- To AutoSum, select a range of cells, and then click `AutoSum` under the "Formulas" tab.

- If you select a range of cells, the shortcut key for `AutoSum` is Alt + = for Windows, and Command + Shift + T for Mac.

- The default for AutoSum is a column range, so if you have a cell selected and click `AutoSum`, it will sum the column above it, which is inaccurate if you are trying to sum numeric values across the row. This is something to check when validating a dataset and checking for inaccuracies.

## Other Auto Functions

<img src="Images/auto.png" width="400" />

- You can also use the auto drop down to automate the `MAX()` and `MIN()` functions, among others.

## AutoFill

<img src="Images/autofill.png" width="400" />

- Excel has a powerful tool called `AutoFill`, which is a little green square in the bottom right corner of a selected cell.

<img src="Images/drag.png" width="400" />

- If you drag that square vertically or horizontally, it will extend the relative formula across cells.

## Quiz

<img src="Images/quiz1.png" width="400" />

<img src="Images/quiz2.png" width="400" />

<img src="Images/quiz3.png" width="400" />

<img src="Images/quiz4.png" width="400" />

<img src="Images/quiz.png" width="400" />

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
