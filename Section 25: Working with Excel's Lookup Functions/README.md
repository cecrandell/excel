# Section 25: Working with Excel's Lookup Functions

## Microsoft Excel VLOOKUP() Function

- [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)

- VLOOKUP() means vertical lookup

<img src="Images/1.png" width="800" />

1. First argument: what value do I need to look for to help me find the last name (the unique match to join on)(lock the column, not the row `$B3`)
2. Second argument: the entire table/list (very important that the first column of this table/array is the unique value to join on `ID`)
3. Third argument: the column number of the value you want to return back
4. Fourth argument: TRUE if you do not mind a non-exact match (e.g, for ID, it will join on the next best ID if it cannot find the exact ID match), otherwise, FALSE

<img src="Images/2.png" width="800" />

## Microsoft Excel HLOOKUP() Function

- [HLOOKUP function](https://support.microsoft.com/en-au/office/hlookup-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f)

- HLOOKUP() means horizontal lookup

1. First argument: what value do I need to look for to help me find the inventory (the unique match to join on)(lock the column, not the row `$B$3`)
2. Second argument: the entire table/list 
3. Third argument: the number of the column of the unique value you want to join on
4. Fourth argument: TRUE if you do not mind a non-exact match (e.g, for ID, it will join on the next best ID if it cannot find the exact ID match), otherwise, FALSE

<img src="Images/3.png" width="800" />

**Developer**

- Caroline Crandell - cecrandell - cecrandell19@gmail.com - [LinkedIn](https://www.linkedin.com/in/carolinecrandell/)
