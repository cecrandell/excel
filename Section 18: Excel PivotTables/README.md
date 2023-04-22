# Section 18: Excel PivotTables

## Understanding Excel PivotTables

- Excel PivotTables allow you to quickly summarize data from a list (e.g., hundreds or thousands of rows). Sometimes you want to see how many sales happened in a month or a year, which product sold the best, or which employees are performing best.

- For example, you might have a column of months, a column of product types, and a column of sales. A PivotTable can show a new table that shows the sales for each product type by month.

- You can quickly edit the layout of a PivotTable, what it summarizes, and how it summarizes.

## Creating an Excel PivotTable

- The most time-consuming part of building a PivotTable is deciding what it is you want to summarize within that PivotTable.

- There's two ways to reference the data that you want to include in a PivotTable: you can reference the range of cells (which can be problematic if you add/delete rows) or you can format your list as a table and use the table name inside the PivotTable reference.

- To format as a table, click anywhere in your list, then go to the "Home" tab in the ribbon and select a style from the "Format as Table" drop down and then indicate the cell range.

<img src="Images/tablename.png" width="800" />

- You can then rename the table name in the "Table Design" tab in the ribbon on the far left (no spaces).

- Once your table is formatted, you can now create a PivotTable by navigating to the "Insert" tab in the ribbon, then select "PivotTable" on the far left.

<img src="Images/createpivottable.png" width="800" />

- A new window will appear and you can select which table you want to use and where you want to put the PivotTable.

<img src="Images/newworksheet.png" width="800" />

- When you click "OK", a new worksheet will be created, the "PivotTable Fields" section will open up, and you will get two new tabs: "PivotTable Analyze" and "Design".

- The column headers from your table will show up in your "Field Name" list. A PivotTable is made up of four sections: "Filters", "Columns", "Rows", and "Values". To start building a PivotTable in your new worksheet, starting dragging values from the "Field Name" list into any of the four sections.

<img src="Images/newpivottable.png" width="800" />

- What you drag into the "Rows" section will appear on the left side of the PivotTable. What you drag into the "Columns" section will appear on the top of the PivotTable. What you drag into the "Values" table will be what populates the rows of your PivotTable.

## Grouping PivotTable Data

<img src="Images/groupselection.png" width="800" />
<img src="Images/newgroupselection.png" width="800" />

- If you would like to group your data in your PivotTable (e.g., grouping months into quarters), you can do so by selecting the cells you want to group, then navigating to the "PivotTable Analyze" tab in the ribbon, then select the "Group Selection" drop down, and click "Group Selection". It will then group that cell range you select and put all remaining cells in that column into their own groups.

<img src="Images/morenewgroupselections.png" width="800" />

- You can repeat the step above with each group you want to make.

<img src="Images/renamegroupselection.png" width="800" />

- Feel free to rename the groups by clicking into the cell.

<img src="Images/printgroup.png" width="800" />

- When you collapse or expand the groups, "Print Preview" will update this and will print however you have the groups configured.

<img src="Images/month2.png" width="800" />
<img src="Images/removegroups.png" width="800" />

- Once you create a group selection, the "PivotTable Fields" section will add a default field into the "Rows" section. To remove groups, you can delete that default field.

<img src="Images/region.png" width="800" />
<img src="Images/salesperson.png" width="800" />

- Another way to make a group is to add multiple values from the "Field Name" list in the "PivotTable Fields" section into the "Rows" section. The order matters (e.g., if you have "Region" above "Salesperson", it will group by "Region". If you have "Salesperson" above "Region", it will group by "Salesperson").

<img src="Images/multiplegroups.png" width="800" />

- You can add more than two values into the "Rows" section, and the rows in the PivotTable will be grouped in the order that they're shown in the "Rows" section.

## Formatting PivotTable Data

- If we need to format the PivotTable (e.g., give currency formatting to numeric cells), we can do so the traditional (manual) way by going to the "Home" tab in the ribbon and clicking the currency icon. However, if we were to add new rows to the PivotTable, Excel is not consistently intuitive in adding the same formatting.

<img src="Images/fieldsettings.png" width="800" />
<img src="Images/fieldsettingswindows.png" width="800" />
<img src="Images/fieldsettingsapplied.png" width="800" />

- To add dynamic formatting to the PivotTable, we right-click on the field we want to format and select the settings we want.

## Modifying PivotTable Calculations

- Sometimes the calculations that the PivotTable applies by default (e.g., summing numeric values and counting alpha values) are not exactly what we need.

<img src="Images/averageofsales.png" width="800" />

- If we wanted a different calculation (e.g., average or minimum values), we can make changes by right-clicking the field, selecting "Field Settings", and choosing a new calculation. You can rename the "Field Name:" accordingly.

<img src="Images/sumandaverage.png" width="800" />

- We can also add a field from the "Field Name" list more than once to have multiple columns of calculations for one field (e.g., "Sum of Sales" and "Average of Sales").

<img src="Images/percentagedifference.png" width="800" />

- If we wanted to show the percentage of how much sales went up or down from month to month, we can do so by adding the "Sales" field again to the "Values" section, right-clicking to select "Field Settings"

<img src="Images/percentagedifferencefrom.png" width="800" />
<img src="Images/percentagedifferencebases.png" width="800" />

- Then you can rename your field, then select "Show data as", then select "% Difference From" from the drop down, then select "Month" as our "Base field:" and "(previous)" as our "Base item:".

<img src="Images/percentagedifferenceapplied.png" width="800" />

- You can now see the percentage of how much sales went up or down from month to month in a second column.

## Drilling Down into PivotTable Data

- If you need to get background information for a cell value, Excel has a feature called "Drilldown".

<img src="Images/drilldown.png" width="800" />
<img src="Images/drilldownworksheet.png" width="800" />

- To drill down a cell value, all you need to do is double-click the cell in the PivotTable, and Excel will create a brand new worksheet will all of the data pertaining to that cell value.

- This feature is especially useful if you need to provide data for a particular salesperson or employee. You can just double-click the cell value, and it will make a brand new worksheet that you've drilled down into, and you can send this new worksheet to that person.

## Creating PivotCharts

<img src="Images/pivotchart.png" width="800" />

- To create a PivotChart based on the data in the PivotTable, you select a cell anywhere in the PivotTable, then select the "PivotTable Analyze" tab in the ribbon, then select "PivotChart".

<img src="Images/columnchart.png" width="800" />

- The default chart is a column chart, but you can change this by selecting the PivotChart, then navigating to the new "Design" tab in the ribbon, and then select the "Change Chart Type" drop down and select your PivotChart of choice.

- The data is dynamic, so if you update values in the PivotTable, the PivotChart will update too (and vice versa). For example, if you remove a value (e.g., "Monthly Percentage"), it will also be removed from the PivotChart.

- Combining a PivotTable with a PivotChart is the beginning of a dashboard.

## Filtering PivotTable Data

- If you filter the PivotTable, the PivotChart will be filtered too.

<img src="Images/filterpivottable.png" width="800" />

- To filter your PivotTable, you can drag a value (e.g., "Year") from the "PivotTable Field" list into the "Filter" section, and then the filter can be seen and changed at the top of your PivotTable and PivotChart.

## Filtering with the Slicer Tool

- Slicers are interactive kinds of dashboard-like elements to filter your PivotTable results.

<img src="Images/insertslicer.png" width="800" />
<img src="Images/slicer.png" width="800" />

- To add a slicer, click anywhere in your PivotTable, then navigate to the "PivotTable Analyze" tab in your ribbon and select "Insert Slicer" and select which value you'd like to filter by.

<img src="Images/multipleslicers.png" width="800" />

- You can also add multiple slicers by repeating the step above.

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
