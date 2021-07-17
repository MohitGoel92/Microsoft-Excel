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
- Any reference to a specific cell must be referenced by the column first and then the row. For example, cell A1 (column = A and row = 1).
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

We can move the entire data table by highlighting the area which contains the data, placing the cursor on the edge of the highlighted area and moving the data to the new desired location. This can also be achieved by adding/removing columns/rows in the worksheet or using cut/copy and paste. To insert rows/columns we can either go onto the *Home* tab and use the functions under the Cells category, or use the keyboard shortcuts:

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

Formatting a list as a table will allow us to create a certain format for our data table which will remain consistent throughout any table operations we perform. In addition, we will have options such as having a filter button and a totals row. This can be performed by selecting the *Home* tab -> *Styles* section -> *Format as Table* where we are able to pick whichever style we prefer. When selecting the table, we observe a new tab called *Table Design*. In the *Table Style Options* section, we can add a *Total Row*, which will sum up the above, regardless of whether there is a filter.

**Note:** This approach is always better than any manual formatting or totalling as table operations will yield undesirable results.

### Using Conditional Formatting to Find and Remove Duplicates

To highlight the duplicates in a list, we simply highlight the column we wish to highlight the duplicates of -> select the *Home* tab -> select *Conditional Formatting* in the *Styles* section -> *Highlight Cells Rules* -> *Duplicate Values ...*. This will highlight all the duplicates we have.

To remove the duplicates, we have two options, we can either:

- Select the *Data* tab -> *Data Tools* section -> *Remove Duplicates* option, or, if we have set our list to be formatted as a table
- Select the *Table Design* tab -> *Tools* section -> *Remove Duplicates* option.

**Note:** When removing duplicates, it is imperative to ensure that we choose a primary key of the table. For example, when searching for duplicates we should search for an ID number, and not employee names as it is possible for two people in a company to have the same first and second name.

## Excel List Functions

The list functions we will be discussing are:

- DSUM()
- DAVERAGE()
- DCOUNT()
- SUBTOTAL()

**DSUM() with a Single Criteria:**

This function is used when wanting to find the total for a certain group, in a list that contains many groups. This function will require three arguments, they are the *Database*, *Field* and *Criteria*.

**Note:** When typing a formula into a cell, guidance is given by pressing the function button (*fx*) with regards to the arguments required.

**DSUM() with OR Criteria:**

The OR criteria takes effect when we select at least 2 categories we wish to combine the total(s) for. This is achieved by selecting a range of categories in the third argument that is called *Criteria*.

**DSUM() with AND Criteria:**

The AND criteria takes effect when we select at least one category and one subcategory. This is achieved by selecting two cells that are horizontally next to each other, along with their headings in the third argument of the formula.

**DAVERAGE():**

DSUM() and DAVERAGE() work exactly the same way, with the only difference being DAVERAGE() provides an average per category instead of the total summation of that category.

**DCOUNT():**

**Note:** With DCOUNT() we have two types, DCOUNT() and DCOUNTA(). The only difference between the two is that DCOUNTA() will take all types of cells into consideration, regardless whether they are a number or a string, however, DCOUNT() will only take numeric cells into consideration.

DCOUNT() works exactly the same way as the above and provides the count per category.

**SUBTOTAL():**

SUBTOTAL() takes at least two arguments, they are the *Function_num* (the number 1 to 11 that specifies the summary function for the subtotal), and the cell reference. The numbers 1 to 11 specify a function such as DSUM(), DCOUNT(), DAVERAGE() ... etc. We use the SUBTOTAL() in situations where the user of the worksheet may filter the rows of the table. As the number of rows displayed changes, the SUBTOTAL() summary function also adapts to only encorporate the rows displayed. The other functions such as DSUM() ... etc will not adapt and will show their results for the whole table.

## Excel Data Validation

In Excel, the data validation feature helps us control what can be entered in our worksheet. For example, we can:

- Create a drop down list of items in a cell
- Restrict entries, such as a date range or whole numbers only
- Create custom rules for what can be entered. This ensures that the data we receive is usable for our analysis.

### Creating an Excel Data Validation List

This will allow us to have complete control over what data may be inputted into the workbook so when the data is returned to us, we can continue with our analysis smoothly. To do this, select the cells we wish to apply the rule for -> Go onto the *Data* tab -> *Data Tools* section -> *Data Validation* button -> Input our preferences. This will give the user a drop down to select from.

### Decimal Data Validation

This data validation specifies a numeric range within a column. To do this, similar to the above we select the cells we wish to apply the rule for -> Go onto the *Data* tab -> *Data Tools* section -> *Data Validation* button -> *Settings* -> *Allow: Decimals* -> Input our preferences.

### Adding a Custom Excel Data Validation Error

When a value is disallowed and excel gives a validation error, the error may confuse the user of the worksheet. This therefore implies we must tell excel to give out a custom message to avoid confusion. To do this, similar to the above we select the cells we wish to customise the error message for -> Go onto the *Data* tab -> *Data Tools* section -> *Data Validation* button -> *Error alert* -> Input our preferences.

### Dynamic Formulas by Using Data Validation Techniques

If we have an interactive table where we input information into a cell with corresponding values, we may use data validation to ensure the user of the workbook is aware of any typos or non-existing values if one should be inputted. This is exactly the same as creating an excel data validation list but instead of manually typing in a list of options, we refer the list of options as a range of cells that contain the options for the list. To do this, select the cell(s) we wish to apply the rule for -> Go onto the *Data* tab -> *Data Tools* section -> *Data Validation* button -> Input our preferences. This will give the user a drop down to select from.

## Importing and Exporting Data

In this section, we will be learning how to import data from an external source. The external source may be a small text file or a large Microsoft Access Database.

### Importing Data from Text Files

To import data into excel from a text file, we simply select the *Data* tab -> *Get & Transform Data* section -> Select the *Get Data* button -> *From File* -> *From Text/CSV* -> Select the desired text file.

**Note:** Before we load the data onto our worksheet, we have the option to transform our data.

### Importing Data from a Database

We will be importing data from a Microsoft Access Database. To do this, we simply select the *Data* tab -> *Get & Transform Data* section -> Select the *Get Data* button -> *From Database* -> *From Microsoft Access Database* -> Select the desired file.

### Exporting Data to a Text File

There are several reasons why one may wish to export their excel worksheet to a text file. For instance, if we wanted to move our data into another system such as a database or another file application which may not communicate with Excel. In addition, text files have a smaller file size which is useful for email as there is no clutter from Excel such as formulas or formatting. A lot of the time databases just require raw data, without any of the Excel influence.

To do this we have two options, we either select *File* -> *Export* -> Choose file type or *File* -> *Save As* -> Change file type.

## Excel PivotTables

A pivot table is a table of grouped values that aggregates the individual items of a more extensive table (such as from a database, spreadsheet, or business intelligence program) within one or more discrete categories. This summary might include sums, averages, or other statistics, which the pivot table groups together using a chosen aggregation function applied to the grouped values.

Pivot tables are a technique in data processing. They arrange and rearrange (or "pivot") statistics in order to draw attention to useful information. This leads to finding figures and facts quickly making them integral to data analysis. This ultimately leads to helping businesses or individuals make educated decisions.

### Creating a PivotTable

When creating a pivot table, we are required to select the data on the worksheet before pivoting. However, if there are any changes to the original table, this will not be picked up by the pivot table. Therefore, to avoid this we should format the data as a table and in the *Table Design* tab, name the table. This table name will be used when telling excel which data we wish to pivot.

To pivot the table, click on any cell within the table -> *Insert* tab -> *Tables* section -> Select *PivotTable* -> Enter table name and desired location of the pivot table.

### Grouping Data

Sometimes, we will want to group certain categories (row labels), for instance, grouping months into quarters. To do this within the pivot table, we select the cells we wish to group -> *PivotTable Analyze* tab -> *Group* section -> Select *Group Selection*. We will notice that the other rows now have their own category each. To group the other cells, simply highlight the next group and follow the same procedure.

### Data Formatting

Data formatting is crucial as it gives clarity due to the ease of understanding. For instance if we are observing sales data, the numbers should be a currency and perhaps whole numbers, instead of random decimal numbers everywhere in the data table. To perform data formatting in the pivot table, select the field name we wish to format -> *Value Field Settings ...* -> *Number Format* -> Select *Category*, *Symbol* and *Decimal places*.

### Modifying PivotTable Calculations

Modifying calculations within our pivot table can provide useful insights to the viewer of the workbook. For instance, alongside the total sum of sales per month we may provide the sales difference as a % from last month. This can be achieved by selecting the field -> *Value Field Settings ...* -> *Show values as* tab -> Select desired options.

**Note:** Double clicking a cell in a pivot table will open a seperate worksheet that contains the relevant data for that value. This is like a drill down function for that value.

### Creating PivotCharts

Pivot charts are very powerful as they are quick and easy to alter and provide powerful insights into the data. Pivot charts are quick to produce after a pivot table has been made. This is performed by selecting any cell within the pivot table -> *PivotTable Analyze* tab -> *Tools* section -> *PivotChart* option.

**Note:** Any changes to the pivot table will be reflected by the pivot chart.

**Note:** Dragging a field into the filter section will also add a filter onto the pivot chart. This enables us to chart data with a specific criteria and analyse data further.

### Slicers

Slicers are visual filters. Using a slicer, we may filter the data (or pivot table, pivot chart) by selecting the type of data we desire. For example, let's say we're looking at a dashboard which contains multiple years and regions. If we wanted to select a particular region(s) or year(s), we simply look at the controls to see what options are available and select at least one of them. This will filter the data to our specific requirements.

To add a slicer, click a cell within the pivot table/chart -> *Insert* tab -> *Filters* section -> *Slicer* option.

## PowerPivot Tools

Power Pivot is an Excel add-in you can use to perform powerful data analysis and create sophisticated data models. With Power Pivot, you can combine large volumes of data from various sources, perform information analysis rapidly, and share insights easily.

In both Excel and in Power Pivot, you can create a Data Model, a collection of tables with relationships. The data model you see in a workbook in Excel is the same data model you see in the Power Pivot window. Any data you import into Excel is available in Power Pivot, and vice versa.

### Activating the PowerPivot AddIn

The power pivot add-in can be enabled by selecting the *File* tab -> *Options* -> *Add-ins* -> *Manage:* -> *COM Add-ins* -> *Go* -> Tick *Microsoft Power Pivot for Excel* box. We will now see a new tab called Power Pivot in our excel workbook.

### Creating a Data Model with PowerPivot

To start creating a model with power pivot, we simple click a cell within a data table -> Select the *PowerPivot* tab -> *Tables* section -> *Add to Data Model*. We observe a new window opening up with the power pivot. If we close this window and repeat the process with another data table within the workbook, we will notice two data tabs being present in the power pivot screen; these are two data tables that we can now combine in our power pivot model.

**Note:** In the power pivot window we observe a section called *Get External Data*. It is here where we can combine data from multiple sources in order to produce our desired model.

### PowerPivot Data Model Relationships

To create relationships between tables (two data tabs) within our power pivot view, we may either select *Diagram View* from the *View* tab and then drag and drop the primary key of the parent table onto the corresponding primary key of the child table, or we can right click the heading of the primary key of the parent table in the normal *Data View*, select *Create Relationship* and manually select the two tables we wish to create a relationship between.

**Note:** This is similar to Power BI, see: https://github.com/MohitGoel92/Power-BI

### Creating PivotTables based on Data Models

To perform this, select a cell within the table of the power pivot view and select the *PivotTable* option in the *Home* tab. This will open a pivot table within our original worksheet from which the power pivot window stemmed from. However, upon close observation we see that our pivot table fields are actually tables that contain all the columns of that table. This now allows us to combine fields (columns) from all the tables that share the same primary key from the parent table.

### PowerPivot KPIs

To manage or edit the data model, simply go onto the *Power Pivot* tab of the excel worksheet and select *Manage*; this will bring us to the power pivot window. In the *Home* tab under the *Calculations* section, we observe two options, *AutoSum* and *Create KPI*. Selecting a cell within a column will allow us to create a calculation at the bottom of that column. For example, we can calculate the total sum, average, count ... etc. Once this step has been completed, we are able to *Create KPI*. This option allows us to manually set how our data should be interpreted. For instance, if we are dealing with profit margin, a profit margin below 10% may be deemed bad (red zone), between 10%-30% reasonable (amber zone) and over 30% ideal (green zone). Once this has been set, we will notice a new field within our pivot table on our original worksheet that has the same name as our calculation. This can be dragged onto the pivot table, just like a normal field. This will now indicate where each row of the data stands in terms of its KPI.

## Working with Large Data Sets

### Freeze Panes

When working with data tables, we may need to scroll down the table but as we scroll, we may notice that the headings disappear. This results in us requiring to scroll up to remind us of what the column means. However, we can use *Freeze Panes* to lock a certain row(s) and/or column(s). To do this, we either select a cell, highlight a column(s) and/or row(s) -> *View* tab -> *Window* section -> *Freeze Panes* option -> *Freeze Panes*.

### Grouping Data (Columns and/or Rows)

In an excel sheet, we can hide columns or rows by highlighting them, right clicking and selecting hide; we can unhide them later. However, we can also set a function at the side of the worksheet which enables us to group the rows or columns by clicking a button for easy user functionality. To do this, we highlight the row(s) or column(s) -> *Data* tab -> *Outline* button -> *Group* -> *Group*.

**Note:** This can be performed on the same column each time with an additional column/row. This is useful if we desire to give the user of the worksheet the option to group the data once column at a time per click or numerous.

### Printing Options for Large Data Sets

When printing large datasets, we may encounter formatting issues such as the column headings not being present in following pages and columns being cut off the right side due to space shortages being printed midway through printing. This however is easily remedied by going onto the *Page Layout* tab -> *Page Setup* section -> *Print Titles* -> In the *Rows to repeat at the top* section we highlight the row with the column titles on the worksheet -> *Page Order* option -> Select *Over, then down*.

