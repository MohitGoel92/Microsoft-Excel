# Microsoft-Excel

This repository contains the necessary tools for learning Excel. This will take a complete Excel novice to a level of business proficiency in reporting and analytics.

Before we get started with the technicalities of Excel, let's begin with some basic facts.

- Since the 2007 release of Excel, the standard Excel file extension has been changed to .xlsx.
- An Excel workbook can contain upto 255 worksheets.
- The Quick Access Toolbar that is located at the top left corner of the application (the green bar) provides access to some of the basic Excel commands.
- We may select any function that we commonly use and pin it to the Quick Access Toolbar for quick utilisation anytime. We only need to simply right click on any command found on the ribbon and select "Add to Quick Access Toolbar".

## Entering and Editing Text and Formulas

We are free to simply click on any cell and enter any numerical or alphabetical data (or a combination of both) into the cell.

If we produce a table with headings and numerical data, we notice that the numerical data is aligned to the right of the cell and the alphabetical headings are aligned to the left of the cell. This is the case with the majority of softwares and there are two reasons for this:

- This makes it clearer to differentiate the difference between the alphabetical and numerical data, avoiding confusion and therefore being practical.
- As the numerical data is aligned to the right, this makes all decimals perfectly aligned regardless of the differences in the number of decimal places. If the numbers were left aligned, we may see the decimal point fluctuate like a wave as we go down the column which is confusing to the eye and therefore prone to giving a false impression.

**Note:** If Excel shows a green triangle at the top left corner of the cell, that means Excel is unsure whether to classify the cell contents as a number or string. This can be resolved by clicking the triangle and setting it to numerical or string.

**Note:** Dates are treated as numerical data, therefore they will be right aligned.

### Date Formatting

If we wish to change the date setting of a cell(s), we simply highlight the cells, select the home tab and select the drop down arrow at the Number section. In the dropdown menu, we can either select a pre-set option or select More Number Formats if we wish to specify a setting that is not showing. We now either select an option from the Date Category or select Customer Category and manually specify the setting we desire. An example where this is useful is if we wanted the date format to be mmm-yyyy as it is not available as a pre-set option.

### Zooming In/Out of the Excel Worksheet

We can either hold down the CTRL key and use the mouse wheel or we can use the zoom bar which is at the bottom right of the worksheet.

### Relative Vs Absolute Cell References

There are two types of cell references: relative and absolute. Relative and absolute references behave differently when copied and filled to other cells. Relative references change when a formula is copied to another cell. Absolute references, on the other hand, remain constant no matter where they are copied.

**Note:** A cell can be made absolute by using the dollar sign ($) in the formula.

### Things to Remember

- The standard way of starting an Excel calculation is with the equals (=) sign.
- Any reference to a specific cell must by referenced by the column first and then the row. For example, cell A1 (column = A and row = 1).
- An absolute cell reference will not change if a formula is copied to another cell in the worksheet.
- In order to make a specific cell absolute, we add a dollar sign ($) before the column and/or the row reference. For example, if we wish to make the cell A1 absolute we use $A$1.

## Working with Basic Excel Functions

**Excel Function:** A predefined formula that performs a calculation.

An example of an Excel Function is the SUM function. For example, if we wished to sum up a column of values, the function may look something like:

```
=SUM(A2:A10)
```

**Note:** A function has three main parts, they are the equals (=), the function name (SUM, MIN, AVERAGE ... etc) and the arguments/parameters. In the above example, the parameters were the two cells A2 and A10.

**Note:** A list of some common functions can be found by clicking the *fx* button on the left of the formula bar. However, a more extensive list is found when the Formula tab is selected where we will find the functions categorised.

**Note:** After a function has been applied, we may see a green triangle at the top left corner of the cell. This is due to Excel trying to assist us as Excel has no context, as it only differentiates between a string or numerical data. Once we have checked the current function and it is correct, it is good practice to remove the warning so the viewers see a clean Excel sheet.

