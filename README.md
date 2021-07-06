# Microsoft-Excel

This repository contains the necessary tools for learning Excel. This will take a complete Excel novice to a level of business proficiency in reporting and analytics.

Before we get started with the technicalities of Excel, let's begin with some basic facts.

- Since the 2007 release of Excel, the standard Excel file extension has been changed to .xlsx.
- An Excel workbook can contain upto 255 worksheets.
- The *Quick Access Toolbar* that is located at the top left corner of the application (the green bar) provides access to some of the basic Excel commands.
- We may select any function that we commonly use and pin it to the Quick Access Toolbar for quick utilisation anytime. For this, we only require to right click on any command found on the ribbon and select *Add to Quick Access Toolbar*.

## Entering and Editing Text and Formulas

We are free to simply click on any cell and enter any numerical or alphabetical data (or a combination of both) into the cell.

If we produce a table with headings and numerical data, we notice that the numerical data is aligned to the right of the cell and the alphabetical headings are aligned to the left of the cell. This is the case with the majority of softwares and there are two reasons for this:

- This makes it clearer to differentiate between the alphabetical and numerical data, therefore avoiding confusion and being practical.
- As the numerical data is aligned to the right, this makes all decimals perfectly aligned regardless of the differences in the number of decimal places. If the numbers were left aligned, we may see the decimal point fluctuate like a wave as we go down the column which is confusing to the eye and prone to giving a false impression.

**Note:** If Excel shows a green triangle at the top left corner of the cell, that means Excel is unsure whether to classify the cell contents as a number or string. This can be resolved by clicking the triangle and setting it to numerical or string.

**Note:** Dates are treated as numerical data, therefore they will be right aligned.

### Date Formatting

If we wish to change the date setting of a cell(s), we simply highlight the cells, select the home tab and select the drop down arrow in the *Number* section. In the dropdown menu, we can either select a pre-set option or select More Number Formats if we wish to specify a setting that is not showing. We now either select an option from the Date Category or select Customer Category and manually specify the setting we desire. An example where this is useful is if we wanted the date format to be mmm-yyyy as it is not available as a pre-set option.

### Zooming In/Out of the Excel Worksheet

We can either hold down the *CTRL* key and use the mouse wheel or we can use the zoom bar which is at the bottom right of the worksheet.

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

An example of an Excel Function is the *SUM* function. For example, if we wished to sum up a column of values, the function may look something like:

```
=SUM(A2:A10)
```

**Note:** A function has three main parts, they are the equals (=), the function name (SUM, MIN, AVERAGE ... etc) and the arguments/parameters. In the above example, the parameters were the two cells A2 and A10.

**Note:** A list of some common functions can be found by clicking the *fx* button on the left of the formula bar. However, a more extensive list is found by selecting the *Formula* tab; here is where we will find the functions categorised.

**Note:** After a function has been applied, we may see a green triangle at the top left corner of the cell. This is due to Excel trying to assist us. Excel has no context, as it only differentiates between a string or numerical data. Once we have checked the current function and it is correct, it is good practice to remove the warning so the viewers see a clean Excel sheet.

**Note:** The small square at the bottom right of the cell is *Auto Fill*. We can drag this to replicate formulas in nearby cells.

**Note: AutoSum** In the Formula tab, the most common statistical functions are categorised in the AutoSum category. The keyboard shortcut for AutoSum is:

```
ALT + =
```

## Modifying an Excel Worksheet

We can move the entire data table by highlighting the area which contains the data, placing the cursor on the edge of the highlighted area and moveing the data to the new desired location. This can also be achieved by adding/removing columns/rows in the worksheet or using cut/copy and paste. To insert rows/columns we can either go onto the *Home* tab and use the functions under the Cells category, or use the keyboard shortcuts:

```
SHIFT + SPACE (highlight row)
CTRL + SPACE (highlight column)
CTRL + SHIFT + + (add the row/column)
CTRL + - (remove the row/column)
```

To make a column or row the appropriate size, we can double click between the two letters (columns) or numbers (rows) where the spacing issue lies, or drag until appropriate. In addition, we can click the triangle at the top-left corner of the worksheet which will highlight the whole worksheet, and double click between any two numbers (rows) or letters (columns). This will ensure appropriate spacing between all the rows and/or columns across the whole worksheet. Similarly, we can make all the columns or rows the same size by selecting the triangle at the top-left of the worksheet and manually altering the size of the columns or rows. 

We can also hide a column(s) or row(s) by highlighting them, right clicking and selecting hide in the menu. To unhide we highlight the columns/rows, right click and select unhide.

To rename or delete a worksheet, simply right click the sheet tab at the bottom and select rename/delete. Alternatively, if we wish to rename the sheet we can double click the sheet name and edit.

**Note:** When you delete a worksheet in Excel, you cannot undo this step. In the event a worksheet has been regrettably deleted, we can close Excel without saving to prevent that worksheet being deleted, however, this will result in losing all the changes we have made after the latest save.

To reorder the sheets, we simply drag and drop the sheet to the place we desire. And if we wish to copy a sheet, we simply hold CTRL, select the sheet and drag it to the right and drop. In addition, if we wish to copy or move a sheet from our current workbook to another workbook, we simply right click the worksheet, select move or copy and send it to either an existing workbook or create a new workbook there and then.

## Formatting Data in an Excel Worksheet

### Font Formatting and Themes

In the *Home* tab we see a category called *Font*. Here we can change the *Font Theme, Colour, Size, Bold, Italic, Underline*, add a border around a cell(s) ... etc. We are also able to change the theme of the workbook entirely. This is achieved by selecting the *Page Layout* tab and selecting *Themes, Colours, Fonts* or *Effects* from the *Themes* category.

In the *Home* tab, we see a *Number* category which allows us to alter the number of decimal points, percentage, commas or currency. Other functions such as date and time conversions also exist here.

The *Format Painter* is located in the *Home* tab in the *Clipboard* category. This allows us to copy the format of another cell and apply the same format to however many cells we desire. If we double click the *Format Painter*, we may continuously click cells in order to replicate the same format. To turn it off we are required to either select *Format Painter* again or press the *ESC* key.

*Cell Styles* is saved formatting which is reused for formatting whenever it is required in the future. This is located in the *Home* tab in the *Styles* category.

**Note:** *Conditional Formatting* can also be performed using the option here. This is where cells are highlighted given a criteria, for example, showing cells red if the value is greater than £500. We can change these rules at any time by selecting *Conditional Formatting* and *Managing Rules*.

To *Merge* cells or *Wrap Text*, we go to the *Home* tab and are able to select these options from the *Alignment* category.

## Inserting Images and Shapes

To insert an image, we go onto the *Insert* tab and click on the *Illustrations* drop down menu. To insert an image we select *Pictures* and to insert a shape we select *Shapes*.

For *SmartArt*, we go onto the *Insert* tab and click on the *Illustrations* drop down menu. To insert *SmartArt* we select *SmartArt* and explore the entire selection.

**Note:** If we select the image or shape we have imported, a new tab called *SmartArt Design* or *Shape Format* reveals itself. This tab is used to modify our current image or shape.

## Creating Charts

To produce charts we highlight the desired data we wish to chart, select the *Insert* tab and go onto the *Charts* category. After we have selected our desired chart, it will be produced in the worksheet. Selecting the chart will result in two new tabs being produced, they are the *Chart Design* and *Format* tab. These tabs are used to refine the current visual to our specific requirements.

**Note:** If we wish to alter the data selected by the chart, we select the chart, go onto the *Chart Design* tab and select *Select Data* in the *Data* section. If we unselect data from the menu or manually alter the selected data using *CTRL*, the axis may still show the names of the previous headings, despite there being no data. To remove this we right click the axis, select *Format Axis* and under *Axis Type* select *Text* axis. This will remove any headings where data has been unselected from.

The chart(s) we have produced can also be sent/moved to another worksheet. This is performed by selecting the chart, going onto the *Chart Design* tab and selecting *Move Chart* in the *Location* category. This is a great way of clearly displaying a chart on full screen and not cluttering a worksheet too much.

## Printing an Excel Worksheet

To print an Excel worksheet, select the *File* tab and go onto *Print*. The options here will be useful when printing in the best possible format using the settings.

In the print settings, if we alternate from Landscape to Portrait we will notice horizontal and vertical dotted lines. These lines represent page breaks that coincide with the page sizes. Therefore, any visualisation or data that crosses this line will be split up and printed onto different pages.

We can use the scaling setting to enlarge certain data or visualisations for clarity and neatness. When scaled, the horizontal and vertical dotted page break lines will also correspond by increasing or decreasing the space between them.

Another option we can utilise is the margins. Decreasing margins may allow more space for additional visuals to be included per page.

To have a birds eye view of the page breaks i.e. how the visuals look per page, we can select the *View* tab and in the *Workbooks Views* we select *Page Layout*.

**Note:** Inside the *Page Layout* view, we can work with the *Header* and *Footer* of the Excel document. If the *Header* and *Footer* is not visible, double click the top of the page and it will appear. When a cell of the *Header* or *Footer* is selected, a tab called *Header & Footer* will appear with options such as *Page Numbers, File Paths, Dates* and *Pictures* which can be used to paste a company logo for instance.

In the *Page Layout* tab within the *Page Setup* category, we have an option called *Print Area*. Setting the print area will print out only that area of a worksheet. This can also be performed by highlighting the cells and selecting the *Print Active Sheets* options within the print settings.

## Working with Excel Templates

Templates allow us to reuse structures (tables, charts, formulas .. etc) which enables us to efficiently perform preset or standardised tasks. A realistic application of this is sending across reports or dashboards to clients or within the company, which involves updating periodical reports with new data. For a data analyst, this is considered standard practice in many companies across industries.

Excel standard templates can be found by selecting *File* -> *New*, and either selecting from the templates presented or searching from the search bar. Once selected, we simply click on the *Create* button.

**Note:** If we are to create our own template, when saving the workbook save it as an *Excel Template*. This option can be found when selecting *Save As* -> *Browse* -> *Save as type*.

# Section II

## Working with an Excel List

**List:** A list is a rectangular range of cells on a worksheet. It has one or more adjacent columns and two or more rows.

When we have a table of data, Excel treats the first row as headings when any kind of formatting has been applied to it. For instance, making it bold, a different colour or altering the font style. Having the first row as headers is useful when performing actions like sorting, filtering, pivot tables and calculations. It is imperative to ensure there are no missing rows or columns, as a missing row or column will separate the dataset into two. When performing Excel operations, only the first part of the dataset will therefore be impacted, giving us incorrect findings.

### Sorting a List Using Single-Level Sort

When sorting a column from a data table, for instance, by last name, we simply select any cell within that column, select the *Data* tab and use the first controls from the *Sort & Filter* section.

### Sorting a List Using Multi-Level Sort

A multi-level sort is required when ordering two or more columns that are within the same data table. This is achieved by selecting any cell within the table -> *Data* tab -> *Sort & Filter* section -> selecting the *Sort* option. We will be required to manually add the levels (columns names) of the criteria we wish to sort.

**Note:** The first level we use to filter (sort by) will be the column that will be sorted independently. If another column is sorted after, the first column will act like a grouping. For example, if we filter last name and then first name, the last names will be sorted but the first names will only correspond to the last names. Therefore the first names are grouped by the last names which will be sorted.

### Using Custom Sorts in an Excel List

Custom sorts are highly useful in situations where we are required to order our list using a specific criteria. For example, if we have the month names in our list and we wish to order them, alphabetically this will yield April, August, February ... etc. However, this is not what we desire as we wish to acquire a list with the months ordered by date. Therefore, selecting a cell in the data table -> *Data* tab -> *Sort & Filter* section -> selecting the *Sort* option and ordering using a *Custom List* will give us the option of ordering by months by timeline.

### Filtering an Excel List Using the AutoFilter Tool

The Autofilter Tool is useful when we are required to select only a proportion of the data. For instance, if we wish to observe data from January and July only, the Autofilter will come in handy here. To use the Autofilter, select a cell in the data table -> *Data* tab -> *Sort & Filter* section -> select the *Filter* option. We now notice that the headings of the data table have a downward pointing arrow which will allow us to select the desired data. To clear the filter, simply click on the *Clear* button.

**Note:** The keyboard shortcut for this is

```
CTRL + SHIFT + L
```

### Creating Subtotals in a List

With a list of products, we may wish to get the subtotals for our analysis. The manual way of doing this would be to order the products, create a new row between each group of products and summing up the above; this being highly tedious and impractical. The efficient way of doing this would be to order the column we wish to subtotal, then use the *Subtotal* function which is found by selecting the *Data* tab -> *Outline* section -> *Subtotal*.

**Note:** This tool gives us bars at the left of the worksheet to alter the granularity of our view. We may collapse certain or all groups for easy analysis which is a great tool for reporting.

### Format a List as a Table

Formatting a list as a table will allow us to create a certain format for our data table which will remain consistent throughout any table operations we perform. In addition, we will have options such as having a filter button and a totals row. This can be performed by selecting the *Home* tab -> *Styles* section -> *Format as Table* where we are able to pick whichever style we prefer. When selecting the table, we observe a new tab called *Table Design*. In the *Table Style Options* section, we can add a *Total Row*, which will sum up the above, regardless of whether there is filter.

**Note:** This approach is always better than any manual formatting or totalling as table operations will yield undesirable results.

### Using Conditional Formatting to Find and Remove Duplicates

To highlight the duplicates in a list, we simply highlight the column we wish to highlight the duplicates of -> select the *Home* tab -> select *Conditional Formatting* in the *Styles* section -> *Highlight Cells Rules* -> *Duplicate Values ...*. This will highlight all the duplicates we have.

To remove the duplicates, we have two options, we can either:

- Select the *Data* tab -> *Data Tools* section -> *Remove Duplicates* option, or, if we have set our list to be formatted as a table
- Select the *Table Design* tab -> *Tools* section -> *Remove Duplicates* option.

**Note:** When removing duplicates, it is imperative to ensure that we choose a primary key of the table. For example, when searching for duplicates we should search for an ID number, and not employee names as it is possible for two people in a company to have the same first and second name.
