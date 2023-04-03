# Section 14: Working with an Excel List

## Understanding the Excel List Structure

- The column headers are identifers, and they are there for us (e.g., to identify what the data is below it) and for Excel (e.g., for pivot tables, sorting, etc.).

- When you first receive a list, make sure there are no empty rows or columns.

- If there is an empty row, Excel will identify this as two separate lists, because they are non-contiguous.

## Sorting a List Using Single Level Sort

- The purpose of a list is to store data so that later on yu can find it. Sorting makes data easy to find.

<img src="Images/columnsort.png" width="400" />
<img src="Images/alphabetical.png" width="400" />

- An easy way to sort a column is to select a cell in the column you want to sort by, then select the "Data" tab in the ribbon, then select the ascending or descending icon.

## Sorting a List Using Multi-Level Sorts

<img src="Images/sort.png" width="400" />

- To sort by more than one column, select a cell in one of the columns of data you want to sort by, then select the "Data" tab in the ribbon, and click the "Sort" icon.

<img src="Images/addlevel.png" width="400" />
<img src="Images/sorted.png" width="400" />

- Then select "Add Level" in the top left corner of the window (+ icon for Mac) and select the additional column you want to sort by in the drop down next to "Then by". You can see the list is sorted by `Last Name` first, then `Emp ID`.

## Using Custom Sorts in an Excel List

<img src="Images/rightclicklist.png" width="400" />

- In the bottom left of the workbook, if you right-click on the arrows, you can see a list of the worksheets in your workbook.

<img src="Images/monthaz.png" width="400" />

- If you were to sort a column by month, Excel sorts this alphabetically by default, which we often do not want.

<img src="Images/customlist.png" width="400" />
<img src="Images/customlists.png" width="400" />

- If you select a cell in the month column, then select the "Data" tab in the ribbon, then click the "Sort" icon, you can select "Custom List..." from the drop down under "Order".

- Excel has four pre-made custom lists: days of the week, abbreviated days of the week, months of the year, and abbreviated months of the year.

<img src="Images/monthsort.png" width="400" />

- Select the months of the year and then Excel will sort accordingly.

## Filter an Excel List Using the AutoFilter Tool

- To filter your list, select a cell in one of the columns of data you want to filter by, then select the "Data" tab in the ribbon, and click the "Filter" icon. Doing so inserts drop down icons next to each of our column headers.

<img src="Images/selectall.png" width="400" />
<img src="Images/july.png" width="400" />

- You can then click a drop down for your chosen column, and untick "(Select All)" to then tick the data value you do want to filter on.

- Once you click out of the filter window, you can see a funnel icon next to the drop down of the column you just filtered, indicating that this column is filtered.

<img src="Images/funnel.png" width="400" />

- You can then go back into the filter window and select other data values you want to include in the filter.

<img src="Images/secondcolumn.png" width="400" />

- You can also filter multiple columns.

<img src="Images/clear.png" width="400" />

- To clear all of your filters, you can click "Clear" under the "Data" tab in the ribbon.

## Creating Subtotals in a List

<img src="Images/unsortedgroup.png" width="400" />
<img src="Images/sortedgroup.png" width="400" />

- Before we can run the built-in Excel subtotal tool, you need to sort the appropriate column that you want to subtotal by (e.g., if I want to subtotal the sales by product, I need to first sort the product column).

<img src="Images/insertrowsum.png" width="400" />

- The manual way of subtotalling (which you should never do, especially if you have thousands of rows of data) is inserting a blank row between each group and manually summing the sales.

<img src="Images/subtotal.png" width="400" />

- The (much) better way of subtotalling is to select a cell in one of the columns of data, then select the "Data" tab in the ribbon, and click the "Subtotal" icon.

<img src="Images/product.png" width="400" />

- Then select the column you want to subtotal the sales for (in our example: "Product").

<img src="Images/sumsubtotal.png" width="400" />

- Then select the function you want to use (for subtotal: "Sum").

<img src="Images/subtotalscomplete.png" width="400" />

- Now you can see the subtotals of sales for products!

<img src="Images/1.png" width="400" />

- A nifty tool that Excel provides when you subtotal: if you select "1" from the far left panel, Excel displays just the "Grand Total".

<img src="Images/2.png" width="400" />

- If you select "2" from the far left panel, Excel displays just the group subtotals and the "Grand Total".

<img src="Images/3.png" width="400" />

- If you select "3" from the far left panel, Excel displays everything.

## Format a List as a Table

- "Format as Table" was introduced to Excel in 2007.

- This tool prevents you from formatting a table manually, and it also is dynamic, so it updates when you filter or sort your list.

<img src="Images/formatastable.png" width="400" />

- To format your list, select a cell in one of the columns of data, then select the "Home" tab in the ribbon, and click the "Format as Table" drop down.

<img src="Images/formatastablerange.png" width="400" />

- You can then select which style you would like for your table and select the data range.

<img src="Images/formattedtable.png" width="400" />

- After clicking "OK", your table is styled, filtered, and there's a new "Table" tab ("Design" tab in Windows) in the ribbon.

<img src="Images/totalrow.png" width="400" />

- From the "Table" tab ("Design" tab in Windows) in the ribbon, you can tick "Total Row", it will by default put a count at the bottom of the outer column.

<img src="Images/counttable.png" width="400" />

- You can choose a different function by selecting the drop down next to that count.

<img src="Images/cornericon.png" width="400" />
<img src="Images/newrows.png" width="400" />

- If you drag down the little corner icon in the bottom of this cell, you can add blank rows to the bottom of your table.

- When you populate these blank rows, the "Total Row" functionality is dynamic and will update automatically.

<img src="Images/newfunction.png" width="400" />

- You can also add another function to the bottom of a column by selecting the cell under the last row of that column and then selecting the drop down to choose a function.

## Duplicate

- To identify duplicate rows, find a unique column, select the top cell of data, then on Windows: Control + Shift + Down Arrow (on Mac Command + Shift + Down Arrow) to select all cells with data in that column.

<img src="Images/duplicatevalues.png" width="400" />

- Then in the "Home" tab, select "Conditional Formatting" then "Highlight Cells Rules" then "Duplicate Values...".

<img src="Images/duplicatevaluesok.png" width="400" />

- A new window will open, and you can choose what to look for and how to style cells and then click "OK". In our case, we are looking for duplicates.

<img src="Images/duplicatevaluesstyled.png" width="400" />

- You can now see that the duplicates have been highlighted.

## Removing Duplicates

- There are two ways to remove duplicates:

<img src="Images/datatabduplicates.png" width="400" />

1. If you haven't formatted your table, you can select "Remove Duplicates" from the "Data" tab in the ribbon.

<img src="Images/tabletabduplicates.png" width="400" />

2. If you have formatted your table, you can select "Remove Duplicates" from the "Table" tab ("Design" tab in Windows) in the ribbon.

<img src="Images/windowduplicates.png" width="400" />

- A new window will appear, and it wants you to identify what you consider to be duplicates.

- If you want to remove **exact** matches, keep everything in the window selected. Sometimes, you might want to unselect first name (e.g., if one row says "Rob" and another says "Robert").

<img src="Images/windowduplicatesempid.png" width="400" />

- But since employee IDs should be unique, we only want to select `Emp ID`.

<img src="Images/duplicatesremoved.png" width="400" />

- When we click "OK", we can see the duplicates have been removed.

## Quiz

<img src="Images/quiz1.png" width="400" />

<img src="Images/quiz2.png" width="400" />

<img src="Images/quiz3.png" width="400" />

<img src="Images/quiz4.png" width="400" />

<img src="Images/quiz.png" width="400" />

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
