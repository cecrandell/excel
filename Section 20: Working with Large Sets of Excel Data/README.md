# Section 20: Working with Large Sets of Excel Data

## Using the Freeze Panes Tool

<img src="Images/freeze.png" width="400" />

- Usually when we have hundreds or thousands of rows in a list, we want to be able to see the column headers as we scroll. To achieve this, we select the cell A2, then navigate to the "View" tab in the ribbon, then select "Freeze Panes".

<img src="Images/unfreeze.png" width="400" />

- To disable this functionality, we can toggle "Unfreeze Panes" in the "View" tab in the ribbon.

## Grouping Data (Columns and/or Rows)

<img src="Images/groupcolumn.png" width="400" />
<img src="Images/plusgroupcolumn.png" width="400" />

- In [Section 5](/Section%2005%3A%20Modifying%20an%20Excel%20Worksheet/README.md), we learned about hiding and unhiding rows and columns, and in this section, we will learn a similar functionality: groups. To group a column, we select a column, then navigate to the "Data" tab in the ribbon, and select "Group" drop down menu and select "Group...", and a minus "-" icon will appear. We can then collapse this column, and the minus will toggle into a plus "+" icon. This enables us to expand the column again.

<img src="Images/multiplegroupcolumn.png" width="400" />

<img src="Images/multiplegroupcolumns.png" width="400" />

<img src="Images/multiplegrouprows.png" width="400" />

- We can also group multiple columns and multiple rows at a time.

- To ungroup columns or rows, you can select the "Ungroup" drop down menu and select "Ungroup..." for one group or "Clear Outline" to clear all groups in the worksheet you are in.

## Print Options for Large Sets of Data

<img src="Images/toomanypages.png" width="400" />

- Often when we are working with large sets of data, printing can be painful, with rows being cut off and expanding into too many pages.

<img src="Images/printtitles.png" width="400" />

- To make print large sets of data more readable and efficient, one thing we can do is make sure the column headers repeat on each page. To do this, we navigate to the "Page Layout" tab in the ribbon, then select "Print Titles" which opens the "Page Setup" window.

<img src="Images/pagesetupprinttitles.png" width="400" />

- If you put your cursor in the entry box next to "Rows to repeat at top:", you can then select the column header row and click "OK".

<img src="Images/printpreviewprinttitles.png" width="400" />

- Now our column headers will appear on each page.

<img src="Images/pageorder.png" width="400" />

- Another thing we can do to make our data more readable when we print is adjusting the page order. If we navigate to the "View" tab in the ribbon and select "Page Break Preview", we can see that the data on the left will print ten pages before printing the data on the right on the eleventh page.

<img src="Images/pageorderover.png" width="400" />

- To fix this, we navigate to the "Page Layout" tab in the ribbon, then select "Print Titles" which opens the "Page Setup" window. Then we change the "Page Order" section from "Down, then over" to "Over, then down".

<img src="Images/pageorderfixed.png" width="400" />

- Now the data on the left will be printed on Page 1, and the data on the right will be printed on Page 2, and so on.

## Linking Worksheets

<img src="Images/summary.png" width="400" />

- Sometimes we want data from one worksheet to appear in another worksheet. For example, if we have one worksheet, "SUMMARY", that needs to incorporate the data of three other worksheets: "2013", "2014", and "2015".

- One way we can do this is through 3D Formulas (a formula that goes to other worksheets).

- A 3D formula (or a "3D reference") performs a mathematical operation that references cells in other worksheets.

<img src="Images/summaryequals.png" width="400" />

- To create a 3D formula, first we select the cell in the worksheet where we want to put the result of the 3D formula, then we type equals "=" to begin a function.

<img src="Images/2013.png" width="400" />
<img src="Images/2013b4.png" width="400" />
<img src="Images/2015b4.png" width="400" />

- Next we select one of the three worksheets we want to sum, then we select the cell within that worksheet with the value we want to sum, then we type plus "+", and select the next worksheet. Repeat this until you've included all of the values you want, and type enter or return.

<img src="Images/3dformula.png" width="400" />

- Now we can see the sum of each value from three different worksheets in one cell on one worksheet. We can also manually type a 3D formula, using this format:

`='[worksheet name]'![cell name]+'[worksheet name]'![cell name]+'[worksheet name]'![cell name]`

- Replace worksheet name with the actual worksheet name (e.g., 2013) and cell name with the actual cell name (e.g., B4).

- Fun fact: 3D formulas can work across different workbooks too.

## Consolidating Data from Multiple Worksheets

- Another way to get data from one worksheet to appear in another worksheet is to use Excel's "Consolidate" functionality.

- First, we select the cell from the top most left corner of where we want the results of our consolidation to appear.

<img src="Images/consolidate.png" width="400" />

- Then we navigate to the "Data" tab in the ribbon and select "Consolidate".

<img src="Images/consolidatewindowfunction.png" width="400" />

- A new window will appear, and we can select our function from the drop down menu.

<img src="Images/consolidatewindowadd2013.png" width="400" />

<img src="Images/consolidatewindowadd2014.png" width="400" />

<img src="Images/consolidatewindowadd2015.png" width="400" />

- We can then manually type our the references of our worksheets and cells, or we can select them by putting the cursor in the entry box beside "Reference:" and select the worksheet then the cell range that contains the cell values we want.

- Then we click the plus "+" button to add each worksheet individually.

<img src="Images/consolidatewindowleftcolumn.png" width="400" />

- Then we tick "Left column" under "Use labels in:" so that Excel knows where to get our labels in case rows are added or deleted within our cell range and can update dynamically.

<img src="Images/consolidatewindowlinks.png" width="400" />

- We can also tick "Create links to source data" so that when cell values update, the consolidations update dynamically too.

<img src="Images/consolidateresult.png" width="400" />

- When we click "OK", we can see the results of our consolidation.

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
