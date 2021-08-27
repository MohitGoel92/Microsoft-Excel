# Microsoft Excel

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

### Linking Worksheets (3D Formulas)

These are also known as cubed formulas. A reference that refers to the same cell or range on multiple worksheets is called a 3D reference. Using 3D formulas allows us to calculate data throughout a workbook using multiple worksheets. Below is an example of how a 3D formula may look:

```
='2013'!B4+'2014'!B4+'2015'!B4
```

The above formula is summing up three cells *(B4)* from three different sheets *(2013, 2014, 2015)* onto a summary sheet.

### Consolidating Data from Multiple Worksheets

To summarise and report results from separate worksheets, we can consolidate data from each sheet into a master worksheet. The sheets can be in the same workbook as the master worksheet, or in other workbooks. When we consolidate data, we assemble data so that we can more easily update and aggregate as necessary.

For example, if we have an expense worksheet for each of our regional offices, we might use consolidation to roll these figures into a master corporate expense worksheet. This master worksheet might also contain sales totals and averages, current inventory levels, and highest selling products for the entire enterprise.

To do this, we click on a cell of an empty table we wish to populate from the top-left corner rightwards -> *Data* tab -> *Data Tools* section -> *Consolidate* -> Highlight a table we wish to include and select *Add* -> Repeat this process for all the tables we wish to add -> Tick the *Use labels in: Left column* box.

**Note:** The *create links to source data* option once ticked will update the consolidated table whenever any changes occur in a linked table.

# Section III

## Conditional Functions

### Working with Name Ranges

A named range is where one or more cells are grouped together and have been given a name. Using named ranges can make formulas easier to read and understand. They also provide simple navigation via the *Name Box*. For example, if we wanted to find the sum total of a range of cells, the function may look like the below:

```
=SUM(C5:C9)
```

However, we may highlight the same range of cells but give them a name by writing a name in the top left *Name Box* within the *Home* tab. Finding the sum total may now look like the below:

```
=SUM(Name_of_cell_range)
```

**Note:** The name of the cell range cannot contain any spaces.

**Advantages**

- Giving a cell(s) a name provides more context/definition. For example, naming a range of cells Week_1 notifies the user of the worksheet that we are using data from week 1.
- Name ranges can not only be used in all formulas, we can name large sets of data by a single name which makes formulas more readable as the user will have context.
- All of the names of the named ranges are all saved within the name box. Therefore, regardless of which worksheet we are viewing, we can easily navigate to the named cell range by clicking on the drop down of the *Name Box* within the *Home* tab, and selecting the desired data we wish to view.
- We are able to print just the named data range.

**Disadvantages**

- Name ranges are absolute formulas. Therefore, we cannot simply drag the formula across to, for instance, sum the columns to the right.

### Editing Name Ranges

To edit or delete name ranges, we go onto the *Formulas* tab -> *Defined Names* section -> Select *Name Manager* -> Select the name range and click *Edit*.

### The IF() Function

The *IF* function is one of the most popular functions in Excel, and it allows you to make logical comparisons between a value and what you expect. So an *IF* statement can have two results. The first result is if your comparison is *TRUE*, the second if your comparison is *FALSE*. An example of the *IF* function is given below:

```
=IF(A1>B1,"Yes","No")
```

The above function will produce a "Yes" if the cell A1 is greater than B2, otherwise it will produce a "No".

### The IF() Function with a Name Range

In the above *IF* statement, if we drag the function down the reference cell B1 will change to B2. If we wish to compare the A column to a number specifically located at cell B1, we must make the cell absolute by using dollar signs. The above function will look like the below:

```
=IF(A1>$B$1,"Yes","No")
```

As Name Range cell(s) is also an absolute cell, therefore, if the target value in cell B1 is named "Target_KPI" for instance, the above *IF* statement will look like the below:

```
=IF(A1>Target_KPI,"Yes","No")
```

### Nesting Functions

A nested function is tucked inside another Excel function as one of its arguments. Nesting functions let you return results you would have a hard time getting otherwise. Two examples of nested functions are given below, they are both producing the same result:

```
=AND(H5="Yes",B5>=8000,C5>=8000,D5>=8000,E5>=8000)
=AND(H5="Yes",MIN(B5:E5)>=8000)
```

### Nesting AND() Function within the IF() Function

A nested function is tucked inside another Excel function as one of its arguments. Nesting functions let you return results you would have a hard time getting otherwise. When a function is nested inside another, the inner function is calculated first. Then that result is used as an argument for the outer function.

Nesting an *AND* function within an *IF* function allows us to test multiple conditions. A function of this sort will look something like the below:

```
=IF(AND(H5="Yes",MIN(B5:E5)>=8000),"Bonus","No Bonus")
```

**Note:** Before the parenthesis (open bracket), we may select the *fx* button and the pop up will guide us to the appropriate arguments required.

### COUNTIF() Function

*COUNTIF* is an Excel function to count cells in a range that meet a single condition. *COUNTIF* can be used to count cells that contain dates, numbers, and text. For instance, a *COUNTIF* function may look like the below:

```
=COUNTIF(H5:H9,"Yes")
```

The above will produce a count of the cells that contain a "Yes".

### SUMIF() Function

The Excel *SUMIF* function returns the sum of the cells that meet a single condition. Criteria can be applied to dates, numbers, and text.

For instance, a *SUMIF* function may look like the below:

```
=SUMIF(C:C,G5,E:E)
```

Where column C is the column which contains a list of the different types of groups, G5 contains the criteria we desire and column E contains the data we wish to sum (sales data).

### IFERROR() Function

The *IFERROR* function is used to catch errors and return a more friendly result or message when an error is detected. When a formula returns a normal result, the *IFERROR* function returns that result. When a formula returns an error, *IFERROR* returns an alternative result. *IFERROR* is an elegant way to trap and manage errors. The *IFERROR* function is a modern alternative to the *ISERROR* function.

Use the *IFERROR* function to trap and handle errors produced by other formulas or functions. *IFERROR* checks for the following errors: *#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?*, or *#NULL!*.

For instance, an *IFERROR* function may look like the below:

```
=IFERROR(VLOOKUP($B11,'Master Emp List'!$A$1:$I$38,3,FALSE),"EMP ID NOT FOUND")
```

In the above function, if the *VLOOKUP* returns an error, instead of an error message it will return the message "EMP ID NOT FOUND". This makes the reasoning of the error clear to the user of the worksheet, therefore making it easier to remedy.

## LOOKUP Functions

The *LOOKUP* function performs an approximate match lookup in a one-column or one-row range, and returns the corresponding value from another one-column or one-row range. *LOOKUP*'s default behaviour makes it useful for solving certain problems in Excel.

### VLOOKUP()

*VLOOKUP* stands for 'Vertical Lookup'. It is a function that makes Excel search for a certain value in a column (the so-called 'table array'), in order to return a value from a different column in the same row. A *VLOOKUP()* function will take a similar form as the below:

```
=VLOOKUP($B3,'Master Emp List'!$A$1:$I$38,2,FALSE)
```

**Note:** The *VLOOKUP* has 4 arguments. The last argument is *FALSE* because, if we had *TRUE* instead, the function will produce the closest match. However, to be precise in our work, we should not seek the closest match but have an output of the exact match or error.

### HLOOKUP()

*HLOOKUP* in Excel stands for ‘Horizontal Lookup’. It is a function that makes Excel search for a certain value in a row (the so-called ‘table array’), in order to return a value from a different row in the same column. Similar to a *VLOOKUP*, a *HLOOKUP* function will take a similar form as the below:

```
=HLOOKUP($B$3,'Master Inventory List'!$A$2:$G$5,2,FALSE)
```

### INDEX() and MATCH() Function

The *VLOOKUP* and *HLOOKUP* are both powerful tools however, come with a limitation. The limitation is that they both start with the cell/column of the identification key to the left. Therefore, any data to the left of that cell/column is not able to be used. To eliminate this, we could cut and paste the identification column to the left of the table, but that means we have altered the data structure which is never a good idea as we may have applications/other worksheets using the current data structure. In addition, *VLOOKUP* is very slow. Therefore, if we have many Lookups in a workbook, it may take time to load up or save. *INDEX* and *MATCH* functions merge together to overcome these limitations.

The Excel *INDEX* function returns the value at a given location in a range or array. You can use *INDEX* to retrieve individual values, or entire rows and columns. The *MATCH* function is often used together with *INDEX* to provide row and column numbers. An *INDEX* function will look like the below:

```
=INDEX('INDEX MATCH Master Emp List'!$A$1:$I$38,10,3)
```

The value that will be displayed in the cell of which the function exists is the value of the 10th row and 3rd column, from the table specified in the first argument.

A *MATCH* function will look like the below:

```
=MATCH($B4,'INDEX MATCH Master Emp List'!$A$2:$A$38,0)
```

The value that will be displayed in the cell of which the function exists, is the row number of the corresponding column of the specified table in the second argument.

Combining *INDEX* and *MATCH* will result in a function that takes the below shape, for ease of reading it is generic:

```
=INDEX(Desired_lookup_table, 
       MATCH(Lookup_value_of_column, Column_of_master_table_where_lookup_value_exists, 0)
       MATCH(Lookup_value_of_row_of_columns, Row_of_column_names_of_master_table, 0))
```

**Note:** In the above function, the first argument establishes the master data table, the second argument finds the correct row number and the third argument finds the correct column number. Therefore, although this function may seem tedious, it is just a small extension of the original/basic *INDEX* function, with two nested *MATCH* functions.

## Working with Text Based Functions

### LEFT(), RIGHT() and MID() Functions

The functions *LEFT*, *MID*, and *RIGHT* allow us to extract substrings from an existing string of data within a given cell. These functions take the generic form given below:

```
=LEFT(Cell, Number_Of_Characters)
=RIGHT(Cell, Number_Of_Characters)
=MID(Cell, Position_Of_Character, Number_Of_Characters)
```

### LEN() Function

*LEN* function is a text function in excel that returns the length of a string/ text. *LEN* Function in Excel can be used to count the number of characters in a text string and able to count letters, numbers, special characters, non-printable characters, and all spaces from an Excel cell. In simple words, *LEN* Function is used to calculate the length of a text in an Excel cell. This function takes the generic form given below:

```
=LEN(Cell)
```

### SEARCH() Function

The Excel *SEARCH* function returns the location of one text string inside another. *SEARCH* returns the position of the first character of find_text inside within_text. Unlike *FIND*, *SEARCH* allows wildcards, and is not case-sensitive.

For example, if we wanted to extract the first name of people from a column that contains both the first name and second, we could use the below function:

```
=LEFT(A2,SEARCH(" ",A2)-1)
```

For the above function, we have used *SEARCH* to find the space that separates the first and second names of a person minus 1. This will produce a number which will only include the characters for the first name. Functions like this can be dragged down and used for the entire worksheet, making them efficient.

### CONCATENATE() Function

We use *CONCATENATE* to join two or more text strings into one string. This function takes the generic form given below:

```
=CONCATENATE(Cell_1," ",Cell_2)
```

This will result in our current cell having an output of both cells being combined and separated with a space.

## Auditing a Worksheet

### Tracing Precedents in Formulas

*Trace Precedents* are cells or a group of cells that affect the value of the active cell. Microsoft Excel provides users with the flexibility of doing complex calculations using formulas such as average, sum, count, etc.

However, the formulas may sometimes return wrong values or even give an error message that you must resolve in order to get the correct values. Excel makes it easier to resolve calculation errors by providing easy-to-use tools such as *Trace Precedents* that one can use to audit the calculations.

To access *Trace Precedents*, select the cell where our formula exists -> Go onto the *Formulas* tab -> *Formula Auditing* section -> *Trace Precedents*.

**Note:** To remove the arrows, simply select the option *Remove Arrows* within the *Formula Auditing* section.

### Tracing Dependents in Formulas

*Trace Dependents* is an Excel auditing tool that shows the cells that are affected by an active cell by displaying arrows linking the related cells to the active cell. When the cells are located on the same worksheet, it is relatively straightforward since Excel will link from the related cells to the active cell using blue arrows.

If the cells are linked to another sheet, instead of an arrow linking the cells, there will be a dotted black arrow pointing to a small icon in the worksheet. To view the details of the cells related to the active cell, double-click the dotted line, and it will open a dialog box with a list of the related cells. Click on any of the listed cells to view the details of the cell.

To access *Trace Dependents*, select the cell where our formula exists -> Go onto the *Formulas* tab -> *Formula Auditing* section -> *Trace Dependents*.

**Note:** To remove the arrows, simply select the option *Remove Arrows* within the *Formula Auditing* section.

### Working with the Watch Window

The Excel *Watch Window* is a tool to help you monitor certain cells that you select to watch. Once you specify the cells you want to watch, a new window will pop up displaying real-time changes happening to the cell or cells you're watching.

To access the *Watch Window*, simple select the cell we wish to monitor -> *Formulas* tab -> *Formula Auditing* section -> *Watch Window*.

**Note:** This window will now stay afloat. If we perform any impacting changes, we will notice that the value of our monitored cell has also changed, therefore making us aware.

### Showing Formulas

If you are working on a spreadsheet where a lot of formulas are being used, it may become challenging to comprehend how all those formulas relate to other cells and whether they are indeed correct. Showing formulas in Excel instead of their results can help you track the data used in each calculation and quickly check your formulas for errors. Microsoft Excel provides a really simple and quick way to show formulas in cells.

To perform this, select the *Formulas* tab -> *Formula Auditing* section -> *Show Formulas*.

**Note:** To unshow formulas, simply click the *Show Formulas* button again.

## Protecting Worksheets and Workbooks

To prevent other users from accidentally or deliberately changing, moving, or deleting data in a worksheet, you can lock the cells on your Excel worksheet and then protect the sheet with a password. Say you own the team status report worksheet, where you want team members to add data in specific cells only and not be able to modify anything else. With worksheet protection, you can make only certain parts of the sheet editable and users will not be able to modify data in any other region in the sheet.

### Protecting Specific Cells in a Workbook

Every cell in an Excel workbook has a property called locked, of which is always turned on. To not allow anyone to make any changes to a cell(s), worksheet or workbook, we only need to lock the cell(s), workbook or worksheet. However, there are two steps to this process of which are given below:

**Step 1:** Highlight the cell(s) we wish to allow the user of the worksheet to update -> *Home* tab -> *Font* section -> *Font Setting* button at the bottom right corner of the section -> *Protection* tab -> Untick the *Locked* box. The cell(s) is now unlocked, which removes the lock from the cell(s).

**Step 2:** We are now going to protect the sheet. To do this, we go onto the *Review* tab -> *Protect* section -> *Protect Sheet* option -> Create a password (optional) -> Click *OK*.

### Protecting the Structure of a Workbook

We are able to protect the structure of the workbook. This will not allow the user of the workbook to perform actions such as creating a new tab, deleting an existing tab, renaming a tab or reordering tabs. 

To perform this we navigate to the *Review* tab -> *Protect* section -> *Protect Workbook* -> Create a new password (optional) -> Select *OK*.

### Adding a Workbook Password

Adding a password to a workbook will restrict the user base of that workbook as only those with the password will be able to open/access it. To create a password for the workbook, simply navigate to the *File* tab -> *Info* -> *Protect Workbook* -> *Encrypt with Password* -> Set password -> *OK*.

**Note:** To remove the password, follow the same procedure but instead of setting a new password, simply leave it blank.

## What If Tools

### Goal Seek Tool

If you know the result that you want from a formula, but are not sure what input value the formula needs to get that result, use the Goal Seek feature. For example, suppose that you need to borrow some money. You know how much money you want, how long you want to take to pay off the loan, and how much you can afford to pay each month. You can use Goal Seek to determine what interest rate you will need to secure in order to meet your loan goal. In this case, we are required to take two steps. Firstly we will use the *PMT* function, and then the *What-If Analysis*.

#### The Payment Function (PMT)

The *PMT Function* is a financial function which calculates the payment for a loan based on constant payments and a constant interest rate. This function takes the generic form given below:

```
= PMT(Rate, Nper, Pv)
```

where:

- Rate: The interest rate for the loan, this is usually the (Interest Rate/12).
- Nper: The total number of payments for the loan.
- Pv: The present value, or the total amount that a series of future payments is worth now; also known as the principal.

#### What-If Analysis (Goal Seek Tool)

To utilise the *Goal Seek* Tool, navigate to the *Data* tab -> *Forecast* section -> *What-If Analysis* -> *Goal Seek*. 

In my experiences, this tool may be potentially confusing to use at first so for further information please use the below resource:

https://support.microsoft.com/en-us/office/use-goal-seek-to-find-the-result-you-want-by-adjusting-an-input-value-320cb99e-f4a4-417f-b1c3-4f369d6e66c7

### The Solver Tool

*Solver* is a Microsoft Excel add-in program we can use for what-if analysis. We use *Solver* to find an optimal (maximum or minimum) value for a formula in one cell — called the objective cell — subject to constraints, or limits, on the values of other formula cells on a worksheet. Solver works with a group of cells, called decision variables or simply variable cells that are used in computing the formulas in the objective and constraint cells. Solver adjusts the values in the decision variable cells to satisfy the limits on constraint cells and produce the result you want for the objective cell.

Put simply, we can use *Solver* to determine the maximum or minimum value of one cell by changing other cells. For example, you can change the amount of our projected advertising budget and see the effect on our projected profit amount.

The *Solver* tool however is not a standard tool. We are therefore required to manually activate it ourself. To do this we navigate to *File* -> *Options* -> *Add-ins* -> Hit *Go...* next to *Manage: Excel Add-ins* -> Tick the box *Solver Add-in*. This new add-in called *Solver* will appear in the *Analyze* section of the *Data* tab.

Please see a screenshot of an example below where this tool has been used. The example is from a worksheet that is included within the Section 3 folder attached.

<p align="center"> <img width="1200" src= "/Pics/Solver.PNG"> </p>

### Building Data Tables with What-If Analysis

We are able to build data tables using the *What-If Analysis*. For instance, using the *PMT* function as previously observed, we can calculate the *PMT* using a specific interest rate but also create a table of values that display the *PMT* for different interest rates. For clarity, let's observe the screenshot below from the workbook included in Section 3:

<p align="center"> <img width="1200" src= "/Pics/Data_Table.PNG"> </p>

From the above, we have computed cell C7 using the *PMT* function by using cells C2, C3 and C4. However, if we wanted to compute the Monthly Payment for cells C8 to C14 using the interest rates from Column B, we are only required to first highlight the two columns -> *Data* tab -> Select *What-If Analysis* from the *Forecast* section -> *Data Table ...* -> Input cell C3 within the *Column Input Cell:* box -> *OK*.

We will observe all the cells being populated with the Monthly Payment reflecting the change in interest rates in Column B.

### Creating Scenarios

A *Scenario* is a set of values that Excel saves and can substitute automatically on our worksheet. We can create and save different groups of values as scenarios and then switch between these scenarios to view the different results. If several people have specific information that you want to use in scenarios, you can collect the information in separate workbooks, and then merge the scenarios from the different workbooks into one. After you have all the scenarios you need, you can create a scenario summary report that incorporates information from all the scenarios. Scenarios are managed with the *Scenario Manager* wizard from the *What-If Analysis* section on the *Data* tab.

For clarity, let's observe the screenshot below from the workbook included in Section 3:

<p align="center"> <img width="1200" src= "/Pics/Scenarios.PNG"> </p>

In the above example we have three scenarios, they are "Default", "Best Case Scenario" and "Worst Case Scenario". Selecting any one of these in the *Scenarios* box and selecting *Show* will alter the numbers in the bottom table to preset values which will impact the growth numbers in the above table. This enables us to show how the numbers will be impacted with each scenario. This is a great visual representation of scenario analysis.

## Automating Repetitive Tasks with Macros

Excel *Macro* is a record and playback tool that simply records your Excel steps and the macro will play it back as many times as you want. VBA Macros save time as they automate repetitive tasks. It is a piece of programming code that runs in an Excel environment but you don’t need to be a coder to program macros. Though, you need basic knowledge of VBA to make advanced modifications in the macro.

*VBA* is the acronym for Visual Basic for Applications. It is a programming language that Excel uses to record your steps as you perform routine tasks. You do not need to be a programmer or a very technical person to enjoy the benefits of macros in Excel. Excel has features that automatically generate the source code for you.

### Activating the Developer Tab

The *Developer* tab, which is a built-in tab in Excel, provides the features needed to use Visual Basic for Applications (VBA) and perform a macro operation. The tab allows users to create VBA applications, design forms, create macros, import and export XML data, etc. The tab is disabled by default. It must be enabled first in the *Options* section on the *File* menu to make it visible on the toolbar at the top of the Excel window. 

To activate the *Developer* tab, we navigate and right click on any tab name -> Select *Customize the Ribbon* -> Select the *Customize Ribbon* option on the left of the window -> Tick the *Developer* box -> *OK*. The *Developer* tab is now active and will appear at the top like all the other tabs.

### Creating a Macro with the Macro Recorder

To create a macro, we simply navigate to the *Developer* tab -> *Code* section -> Select *Record Macro* -> Name the macro -> Set the shortcut key -> Select where to store the macro -> Write a description of the macro -> *OK*.

As soon as we click *OK*, the macro starts recording. Anything we do now is captured by Excel. When we have completed the desired steps, we simply click *Stop Recording* in the *Code* section of the *Developer* tab.

**Note:** When setting a keyboard shortcut for the macro, it is imperative to consider that the shortcut will override any existing shortcut. Therefore, try to avoid setting the macro shortcut to any pre-existing shortcut.

### Editing a Macro with VBA

Suppose during the recording of the macro we make an error. This does not mean we are now required to re-record all the steps again as we can simply change the steps where we have made the error using VBA. To access this, navigate to the *Developer* tab -> *Code* section -> Select *Visual Basic* -> In the project window on the top left scroll down to *Modules* (macros are stored as modules) -> Select the module from the bottom left window ->  The code will now appear on the window on the right.

In our example from the Excel workbook from section 3, the VBA code will appear as below:

```
Sub FormatTable()
'
' FormatTable Macro
' This macros places headers on the table and formats the data.
'
' Keyboard Shortcut: Ctrl+j
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "EMP ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Last Name"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "First Name"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Dept"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Email"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Ext"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Location"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Hire Date"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Pay"
    Cells.Select
    Cells.EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A1:I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Range("A1:I1").Select
    Selection.Font.Size = 11
    Selection.Font.Bold = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-39
    Selection.NumberFormat = _
        "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "dd/mmm/yy"
    Selection.NumberFormat = "d-mmm-yy"
    Cells.Select
    Selection.ColumnWidth = 13.09
    Range("A1").Select
    Selection.AutoFilter
    Selection.ColumnWidth = 15
    Columns("B:B").AutoFit
    Columns("C:C").AutoFit
End Sub
```

Without getting too into VBA at this stage, using simple logic we are able to edit the code and tailor it to our desired final output.

### Creating Buttons to Run Macros

In order to run our macro, we simply navigate to the *Developer* tab -> *Code* section -> *Macros* -> Select the desired macro and hit *Run*. Alternatively, we can produce a button at the side of the worksheet that also runs our desired macro. To do this we simply navigate to the *Controls* section within the *Developer* tab -> Select *Insert* -> Select the first option called *Button*. We simply set up the button by making our own box, right clicking it to make visual changes, clicking somewhere within the worksheet to get out of edit mode and selecting it using the left click to run. Both of these approaches are valid when trying to run macros.

## Macros and VBA Practice Projects

In this section, we will discuss projects that build upon each other to deepen our understanding of VBA and automation. In the seven projects we will discuss:

- **Automation**: The objective is to hit a single button and our report or final product should be ready.
- **Macro Recorder**: There are a few different ways where we can automate our experience through macros. The first way we will discuss is the *Macro Recorder*. This is where Excel records our moves by writing them out in the background using code.
- **VBA Concepts**: Here will discuss *Looping Structures* which is useful if we want something to happen multiple times. We create logical statements to achieve this. In addition, we will go through variables.
- **Dynamic Code**: *Dynamic* formulas allow us to insert Excel's formulas into cells even when the formula must reference rows that will be inserted later during the merge process. They can repeat for each inserted row or use only the cell where the field is placed. In other words, *Dynamic Code* accommodates for differences in the data, for instance, if the amount of data being received is not consistent.
- **Reusable Code**: *Reusable Code* is code that can be reused without having to re-write it everywhere.
- **Variables**: Declaring a variable helps us store information such as single or multiple values, where it can easily be retrieved. For instance, an input from a user by asking them a question.
- **Logical Statements**: Logical functions are used to test whether a situation is true or false. Depending on the result of that test, we can then elect to do one thing or another. These decisions can be used to display information, perform different calculations, or to perform further tests.
- **Loops**: We will learn how to get blocks of code or entire procedures of code to run multiple times. With a loop we can perform tasks that are to be repeated for all worksheets simultaneously.

### VBA Modules

When we record a Macro, it will show up within the *VBA* window as a module. If we wish to write our own *VBA* code, we will have to first create a new module by navigating onto the options bar at the top of the *VBA* window -> *Insert* -> *Module*.

### Creating a VBA Procedure

Now that we've got our module, we can begin writing our own code. However, in a new module we do not simply start typing code using our keyboard, as there are very specific key words and structures that *VBA* expects us to adhere to as we write our code. When recording a Macro using the *Record Macro* button, excel is writing down all the code in the background in what's called a *Procedure*. 

**Note:** A *Procedure* is a *Macro* and vice versa.

Now that we've gotten our *Module*, we must insert a *Procedure*, name the *Procedure* and then put our code inside that *Procedure*. To do this we navigate to the options bar at the top of the *VBA* window -> Select our current *Module* -> *Insert* -> *Procedure ...*. We now observe a new window called *Add Procedure*. We firstly insert our desired name at the top, and then observe three options called *Sub*, *Function* and *Property*.

**Sub (or Sub Procedure)**

Sub is a set of codes to run within the keywords *Sub* and *End Sub*. You can think of it as a list of instructions of what you want VBA to do in order. For example, if we want to automate a bunch of actions such as adding formatted headers to our dataset, then we write those down for *VBA* to execute.

Note the below key differences between Sub and Function:

1) Sub can take parameters from a user.

2) Sub cannot return a value.

**Function (or Function Procedure)**

Consider the following formula:

```
y = 2x+3, or

y = f(x)
```

In the above formula, the value of y depends on x. In mathematics we call y the “dependent variable” and x as “independent variable”, because y depends on x. We also call this formula “y is a function of x”, as the formula 2x+3 or f(x) returns value y. In Excel we have used Functions before, such as SUM, COUNT, TEXT and VLOOKUP. The returned value depends on the Function parameter (or argument).

Below is a summary of Function.

1) Function usually requires one or more parameters (arguments) from the user.

2) Function returns a value.

3) Function can be used independently, which means the use of a formula is not bound to be used only by specific Cells or Worksheets.

Similar to a Worksheet Function (those you use everyday in Excel), VBA also offers a lot of additional Functions but they come from a completely different library. You may still find the Excel VBA library quite similar to Worksheet library but they are different in terms of syntax and you fail to find many Worksheet Function in Excel VBA library.

**Property**

Property is the “attribute” or “characteristics” of an Object. For example, Range has attributes such as colour and border colour. You can set a Property to a Range, or you can retrieve the Property value from a Range.

Property has the ability to set and retrieve value because Property is declared as a pair, one declares using Let, one declares using Get, but there are Property that do not have the retrieval part declared.

Below is an example to set a formula in Range A1:

```
Range("A1").Formula = "=A4+A10"
```

Or you can retrieve a formula from A1:

```
val = Range("A1").Formula
```

**Method**

Method can only be used in VBA, but not in an Excel worksheet. In terms of declaration, *Method* is the same as *Function* except that *Method* only works for a particular Object.

In Excel, “Object” mostly refers to a Range, Worksheet or Workbook. Each Object has its own set of Methods, for example, Range has its exclusive Method called “Autofill” used to autofill formulas in the cells below.

```
Set sourceRange = Worksheets("Sheet1").Range("A1:A2") 
Set fillRange = Worksheets("Sheet1").Range("A1:A20") 
sourceRange.AutoFill Destination:=fillRange
```

**Note:** Different Objects may use the same Method name, but they can be entirely different things.

Below are some attributes of Method:
 
1) Method usually requires one or more parameters (arguments) from the user, but allows no argument.
 
2) Method always returns something, for instance an Object, Number or String.
 
3) Method requires an Object to act on, it cannot be used independently.
 
4) Method has its own library. Therefore it is imperative to not mix up the Function library with Method Library.

**Scope: Public and Private**

Below the *Type* options we have two *Scope* options, *Public* and *Private*. The basic concept is that *Public* variables, subs or functions can be seen and used by all modules in the workbook while *Private* variables, subs and functions can only be used by code within the same module.

For our learning purposes, we will name the procedure FirstProcedure, select *Sub Type* and *Public Scope*. When we hit the *OK*, we observe the following within the VBA window:

```
Public Sub FirstProcedure()

End Sub
```

**Note:** As we get more experienced with *VBA*, we will not be able to type this out without following the above procedure.

### Adding Code to a VBA Procedure

Let's observe the simple VBA code below:

```
Public Sub FirstProcedure()
    ActiveCell.Value = "Excel VBA"
End Sub
```

If we go back to our Excel worksheet and search for the our Macros, the above Macro called "FirstProcedure" will be present. If we select it and run it, the string "Excel VBA" will be produced within that cell.

### VBA Variables

A *Variable* is a location within our code (or memory) where we can store content. A typical example where a variable is used is when the *Procedure* will recieve an input from a user and use it later in some way. For instance, let's observe the VBA code below:

```
Public Sub FirstProcedure()
    Dim UserInput As String
    
    UserInput = "Hello World!"
    
    ActiveCell.Value = UserInput
End Sub
```

where:

- *Dim* stands for dimension which creates some space (plots out some space or the "dimension") of which we can put something inside of. *Dim* creates a standard variable.
- UserInput is the variable name.
- *As String* declares our variable type as a string or text value. Other types include Integer, Boolean ... etc.

**Note:** If we click onto a cell within our worksheet, we can simply click the *Run* button and the *Procedure* will execute and perform the stated actions.

### Building Logical IF Statements

Let's observe the below IF statement:

```
Public Sub FirstProcedure()
    Dim UserInput As String
    
    UserInput = "21"
    
    If UserInput > 20 Then
        ActiveCell.Value = "Access Granted"
    Else
        ActiveCell.Value = UserInput
    End If
End Sub
```

We observe that the user input is now a numeric value and the IF statement is wedged between *If* and *End If*. As the conditions state, if the UserInput is greater than 20, the cell we have clicked onto will have the output "Access Granted", otherwise it will generate the number inputted. In our case, 21 is greater than 20, therefore the *Procedure* will produce the output "Access Granted".

### Working with Loops to Repeat Blocks of Code

We will be discussing a *Do While Loop*, a very common loop. The Excel *Do While Loop* function is used to loop through a set of defined instructions/code while a specific condition is true. Let's observe the below *VBA* Code which is designed to turn a cell red if the number within that cell is greater than 10:

```
Public Sub FirstLoop()

    Dim i As Integer
    
    i = 1
    
    Do While i <= 10
        
        If ActiveCell.Value > 10 Then
            ActiveCell.Interior.Color = RGB(255, 0, 0)
        End If
        
        ActiveCell.Offset(1, 0).Select
        
        i = i + 1
    
    Loop
End Sub
```

where: 

- *Dim* stands for dimension which creates some space (plots out some space or the "dimension") of which we can put something inside of. *Dim* creates a standard variable.
- *i* is the variable name, of which we have arbitrarily set to 1.
- *Do While i <= 10* ensures that the loop runs until *i* hits a value of 10. In other words, the loop should run 10 times (from *i* = 1 to *i* = 10 inclusively with an incremental step of 1, this is established by the line *i = i + 1*).
- *As Integer* declares our variable type as an integer. Other types include String, Boolean ... etc.
- *If* and *End If* are the keywords enclosing the loop.
- *ActiveCell.Value > 10* is our condition that will decide to either colour the background or not. In this case, if the cell contains a value greater than 10, the background will be coloured.
- *ActiveCell.**Interior.Color** = RGB(255, 0, 0)* means we wish to have red at 100% level (255 means 100% of that color) with zero green and zero blue.
- *ActiveCell.**Offset(1, 0)**.Select* moves the active cell on which the code will execute down by one row and zero columns each time (each loop).

### Project 1

In this project, we will get acquainted with the Excel Macro recorder. For this project, we will simply record us adding column headers and formatting them to our data. This will then be used to perform the same task in the following tab using the shortcut *CTRL + j*, and using a button on the following tab.

### Project 2

In this project, we will be writing and analysing VBA code which sorts data depending on the user input, therefore being interactive with the user. This will be accomplished by creating variables within a procedure, using logic through *IF* statements, using an *Input Box* object, and a *MSG* box (or message) object.

**Note:** We notice that the project is saved as *.xlsm* which is due to this being a *Macros* enabled document as there is already VBA code contained within the file.

We will be using sorting procedures that only run a single action, command or selection. Let's observe one of the three sorting procedures within the project:

```
Sub DivisionSort()
'
' Sort List by Division Ascending
'

'
    Selection.Sort Key1:=Range("A4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub
```

where: 

- A4 is the first data cell of the column called "Division" of which will be sorted if the user wishes.

To produce a message box in order to be interactive with the user, we will be using the below code:

```
Public Sub UserSortInput()

    Dim userInput As String
    Dim promptMSG As String
    Dim tryAgain As Integer
    
    promptMSG = "Enter a numeric value to sort ..." & vbCrLf & _
        "1 --- Sort by Division" & vbCrLf & _
        "2 --- Sort by Category" & vbCrLf & _
        "3 --- Sort by Total"
        
    userInput = InputBox(promptMSG)
    
    If userInput = "1" Then
        DivisionSort
    ElseIf userInput = "2" Then
        CategorySort
    ElseIf userInput = "3" Then
        TotalSort
    Else
        tryAgain = MsgBox("Invalid Value! Please Try Again", vbYesNo)
        If tryAgain = 6 Then
            UserSortInput
        End If
    End If

End Sub
```

where:

- *&* is the ampersand which is used to concatenate or combine.
- *vbCrLf* is the *Character Return Line Feed* which drops down a line so we can give the user the choices.
- *& vbCrLf & _* therefore combines the current line with the next line.
- The variable *userInput* will save the input from the user from the line *userInput = InputBox(promptMSG)*, which will then be used going forward to sort the desired column values.
- The *IF* statements are between the *If* and the *End If*. Depending on what the user inputs, if it's "1" then the *Procedure* will execute a *Sort* of the "Division" column, if it's "2" then the *Procedure* will execute a *Sort* of the "Category" column, and if it's a "3" the *Procedure* will execute a *Sort* of the "Total" column.
- *MsgBox* is a built in Excel function that displays our desired message in the event of an error. In our example, if the user inputs a value of anything other than 1, 2 or 3, we wish to catch it, direct them by explaining what went wrong and how this can be remedied. This also enables us to give the user additional information and be descriptive about it, in case the user has received incorrect information and the remedial procedure or implications of the current situation require action.

**Note:** "TryAgain = 6" refers to the value obtained from the *MsgBox* if the user selects "Yes". The *MsgBox* function returns a number value based on which button the user presses. Use the below table for clarification:

<p align="center"> <img width="200" src= "/Pics/EMB.png"> </p>

On our *MsgBox*, we used a Yes/No button combination. There are other buttons that you can use other than Yes/No.

### Project 3

In this project, we will be creating three *Procedures*. Two will be recorded using the *Macro* recorder and we will use *VBA* to create our own procedure that makes a call on the two recorded procedures. For example, if we had to clean up data that was on dozens or hundreds of worksheets in order to create a daily report, this process can be automated and completed with the click of one button.

**Note:** Before using macros, ensure a duplicate copy of a worksheet is created as another tab within the workbook to test our code on. In addition, save the document before running any macros as macros cannot be undone. If something goes wrong now, we simply close the document without saving and reopen it.

The two recorded *Macros* we produce in the project have the below *VBA* code:

```
Sub AddHeaders()
'
' AddHeaders Macro
' This will add headers the columns of data
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub FormatHeaders()
'
' FormatHeaders Macro
' This will format the headers
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("G7").Select
End Sub
```

The *Procedure* we code ourself in the project that executes the *Macros* on each tab is given below:

```
Public Sub CleanUpData()

    Dim i As Integer
    
    i = 1
    
    Do While i <= Worksheets.Count
        
        Worksheets(i).Select
        AddHeaders
        FormatHeaders
        
        i = i + 1
    
    Loop
End Sub
```

where:

- Each worksheet is selected using the code "Worksheets(i).Select".
- "AddHeaders" and "FormatHeaders" are the names of the two *Macros* that will run one after the other.

**Note:** Running the macro "CleanUpData" will apply both the procedures to each tab in a flash.

### Project 4

In this project, we will observe the automation of a function inside of excel. In our workbook we have four worksheets, and we will automate the sum function by summing the "Total Expense" column on each worksheet with a click of a single button. The emphasis here will be to explore the technique of utilising variables to help us navigate to the desired cell of which the summation should occur. We will perform this using two steps. We will firstly automate the *SUM* function within excel using *VBA*, and secondly we will automate the *SUM* function to happen across multiple worksheets.

**Note:** Upon observing the worksheets, we notice that although the "Total Expense" column is consistently in column "F", the summation cell for that column is not consistent. We must bear this in mind to accommodate this inconsistency.

The *VBA* code is what we will produce in the project:

```
Public Sub Automate_Sum()
    
    'selects the F2 cell of the active sheet
    Range("F2").Select
    
    'selects the last cell in the column
    Selection.End(xlDown).Select
    
    lastCell = ActiveCell.Address(False, False)
    
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "=SUM(F2:" & lastCell & ")" 
    
End Sub
```

where:

- "Range("F2").Select" selects the F2 cell of the active sheet.
- "Selection.End(xlDown).Select" is similar to selecting *column F + CTRL + Down Arrow*.
- The arguments (False, False) in the parenthesis of "lastCell = ActiveCell.Address(False, False)" is to not allow the dollar signs within the formula or cell reference itself (for the row and column).
- "ActiveCell.Offset(1, 0).Select" moves to a cell one below and no change to the column, or in other words, the cell below the last populated cell of the column to enable a summation of the column.
- "ActiveCell.Value = "=SUM(F2:" & lastCell & ")" " is the summation formula.

The first step is now complete and we can successfully run the *VBA* code on our first weeksheet.

**Note:** In our *VBA* view, we are able to run each line of code independently and observe excel doing its work. To do this, we simply hit the *F8* key and the code will run line by line, every time we click *F8*.

The code below is the finished product which loops the automated summation through all worksheets:

```
Public Sub Automate_Sum()
    Dim lastCell As String
    Dim i As Integer
    
    i = 1
    
    Do While i <= Worksheets.Count
        
        Worksheets(i).Select
            
        'selects the F2 cell of the active sheet
        Range("F2").Select
    
        'selects the last cell in the column
        Selection.End(xlDown).Select
    
        lastCell = ActiveCell.Address(False, False)
    
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "=SUM(F2:" & lastCell & ")"
    
        i = i + 1
    Loop
End Sub
```

### Project 5

In this project, we will be combining what we have previously learnt. This includes loops, variables, copying and pasting content, making calls to other procedures that we've already created, formatting the data and automating the processes such that they run through all worksheets and populate the last tab with our desired report. Like the previous project, we will be breaking this down into two steps. Firstly, we will be creating a loop for the procedures to loop across all the desired tabs, and secondly we will be creating a procedure that copies and pastes data that creates our report on the final tab.

The *VBA* code used for step one is given below:

```
Public Sub FinalReportLoop()

    Dim i As Integer
    
    i = 1
    
    Do While i <= Worksheets.Count - 1
        Worksheets(i).Select
        
        AddHeaders
        FormatData
        AutoSum
        
        ' Copying the current data
        range 
        
        
        i = i + 1
    Loop
End Sub
```

where:

- "AddHeaders", "FormatData" and "AutoSum" are procedures we have already made. For the whole code, please see project 5.

The second step required us to copy the data from four tabs and paste it onto the fifth tab for the purpose of creating a report. The *VBA* code for this after the second step has been completed looks like the below:

```
Public Sub FinalReportLoop()

    Dim i As Integer
    
    i = 1
    
    Do While i <= Worksheets.Count - 1
        Worksheets(i).Select
        
        AddHeaders
        FormatData
        AutoSum
        
        ' Copying the current data
        Range("A1").Select
        
        ' Current region is like hitting CTRL + a and all the cells are selected
        Selection.CurrentRegion.Select
            
        Selection.Copy
        
        ' Select the final report worksheet called Yearly Report
        Worksheets("Yearly Report").Select
        
        ' Find the empty cells
        Range("A30000").Select
        
        ' We are going up until we hit a populated cell, i.e. a cell with data
        Selection.End(xlUp).Select
        
        ActiveCell.Offset(3, 0).Select
                
        ' Paste the new data into the worksheet
        ActiveSheet.Paste
                    
        i = i + 1
        
        ' Natural fit columns width per worksheet
        Columns("A:F").EntireColumn.AutoFit
    
    Loop
    
    ' Natural fit columns of report worksheet
    Columns("A:F").EntireColumn.AutoFit
End Sub
```

### Project 6

In this project, we will be talking about User Forms and how we can interact with the user with a drop down menu with buttons the users can click on. The code we created in project 5 will also be incorporated as one of our tools as part of the functionality of the user form. The first step of the project will be to create a user interface of the user form. The user form will interact with all the worksheets, add worksheets and run the final report. Let's start building the user form inside of our *VBA* window.

In the *VBA* window, when we wanted to create a procedure we navigated to the *toolbar* at the top -> *Insert* -> *Module* -> Create our procedure within the *VBA* window. To create a form, we navigate to the *toolbar* -> *Insert* -> *UserForm*. We now observe a new folder called *Forms*.

The picture below shows us the above and the *Toolbox* button used for creating features on the user from box.

<p align="center"> <img width="200" src= "/Pics/userform.png"> </p>
