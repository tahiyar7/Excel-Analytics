# Excel-Analytics
This repository contains Excel-based analytics projects showcasing various data analysis techniques, including data cleaning, visualization, and statistical analysis. It serves as a resource for learning and demonstrating practical applications of Excel in data-driven decision-making.

## Conditional Formatting 
<br />
This cheat sheet is one of my favorites short cuts for the excel daily tips. 
<br />

![image](https://github.com/user-attachments/assets/478f3ce0-06ec-4cd4-a8da-90ae83912bbe)



## Lookup and Data Cleaning Functions in Excel
Lookup Functions
Lookup functions are essential for searching and retrieving data within Excel.<br />

:blue_square:VLOOKUP: Searches for a value in the first column of a table and returns a value in the same row from a specified column. <br />
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])<br />
lookup_value: The value to search for.<br />
table_array: The range of cells that contains the data.<br />
col_index_num: The column number in the table from which to retrieve the value.<br />
range_lookup: (Optional) TRUE for approximate match, FALSE for exact match.<br />

:blue_square:HLOOKUP: Searches for a value in the top row of a table and returns a value in the same column from a specified row.<br />
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])<br />
lookup_value: The value to search for.<br />
table_array: The range of cells that contains the data.<br />
row_index_num: The row number in the table from which to retrieve the value.<br />
range_lookup: (Optional) TRUE for approximate match, FALSE for exact match.<br />

:blue_square: XLOOKUP: Searches a range or an array and returns an item corresponding to the first match it finds. If a match doesn’t exist, then XLOOKUP can return the closest (approximate) match.<br />
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])<br />

:blue_square:lookup_value: The value to search for.<br />
lookup_array: The array or range to search.<br />
return_array: The array or range to return.<br />
if_not_found: (Optional) The value to return if no match is found.<br />
match_mode: (Optional) 0 - Exact match, 1 - Exact match or next larger item, -1 - Exact match or next smaller item, 2 - Wildcard character match.<br />
search_mode: (Optional) 1 - Search first to last, -1 - Search last to first, 2 - Binary search ascending order, -2 - Binary search descending order.<br />

:blue_square:MATCH: Searches for a specified item in a range of cells and returns the relative position of that item.<br />
MATCH(lookup_value, lookup_array, [match_type])<br />
lookup_value: The value to search for.<br />
lookup_array: The range of cells to search.<br />
match_type: (Optional) 1 - Less than, 0 - Exact match, -1 - Greater than.<br />
:blue_square:INDEX: Returns the value of an element in a table or array, selected by the row and column number indexes.<br />


INDEX(array, row_num, [column_num])<br />
array: The range of cells or array constant.<br />
row_num: The row number in the array.<br />
column_num: (Optional) The column number in the array.<br />
:blue_square:CHOOSE: Returns a value from a list of values, based on an index number.<br />


CHOOSE(index_num, value1, [value2], ...)<br />
index_num: Specifies which value to return.<br />
value1, value2, ...: The values from which to choose.<br />
:blue_square:OFFSET: Returns a reference to a range that is a specified number of rows and columns from a cell or range of cells.<br />


OFFSET(reference, rows, cols, [height], [width])<br />
reference: The reference from which you want to base the offset.<br />
rows: The number of rows, up or down, that you want the upper-left cell to refer to.<br />
cols: The number of columns, to the left or right, that you want the upper-left cell to refer to.<br />
height: (Optional) The height, in number of rows, that you want the returned reference to be.<br />
width: (Optional) The width, in number of columns, that you want the returned reference to be.<br />
## Data Cleaning Functions<br />
Data cleaning functions help in preparing and cleaning data for analysis.<br />
<br />

:blue_square:TRIM: Removes all spaces from a text string except for single spaces between words.<br />
TRIM(text)<br />
text: The text from which you want spaces removed.<br />

:blue_square:CLEAN: Removes all nonprintable characters from text.<br />
CLEAN(text)<br />
text: The text from which you want to remove nonprintable characters.<br />

:blue_square:SUBSTITUTE: Substitutes new text for old text in a text string.<br />
SUBSTITUTE(text, old_text, new_text, [instance_num])<br />
text: The text or the reference to a cell containing text.<br />
old_text: The text you want to replace.<br />
new_text: The text you want to replace old_text with.<br />
instance_num: (Optional) Specifies which occurrence of old_text you want to replace. If omitted, every occurrence of old_text is replaced.<br />

:blue_square:REPLACE: Replaces part of a text string, based on the number of characters you specify, with a different text string.<br />
REPLACE(old_text, start_num, num_chars, new_text)
old_text: The text string containing characters you want to replace.<br />
start_num: The position of the first character you want to replace.
num_chars: The number of characters in old_text that you want REPLACE to replace with new_text.
new_text: The text you want to replace old_text with.<br />

:blue_square:TEXT: Converts a value to text in a specific number format.<br />
TEXT(value, format_text)
value: The value to be converted to text.
format_text: The number format that you want to apply.<br />

:blue_square:VALUE: Converts a text string that represents a number to a number.<br />
VALUE(text)
text: The text enclosed in quotation marks or a reference to a cell containing the text you want to convert.<br />

:blue_square:UPPER: Converts text to uppercase.<br />
UPPER(text)
text: The text you want to convert to uppercase.<br />

:blue_square:LOWER: Converts text to lowercase.<br />
LOWER(text)
text: The text you want to convert to lowercase.<br />

:blue_square:PROPER: Converts text to proper case; the first letter in each word in uppercase, and all other letters in lowercase.
PROPER(text)<br />
text: The text you want to convert to proper case.<br />

:blue_square:LEFT: Returns the specified number of characters from the start of a text string.<br />
LEFT(text, [num_chars])
text: The text string containing the characters you want to extract.
num_chars: (Optional) Specifies the number of characters you want LEFT to extract.<br />

:blue_square:RIGHT: Returns the specified number of characters from the end of a text string.<br />
RIGHT(text, [num_chars])
text: The text string containing the characters you want to extract.<br />
num_chars: (Optional) Specifies the number of characters you want RIGHT to extract.<br />

:blue_square:MID: Returns a specific number of characters from a text string, starting at the position you specify.<br />
MID(text, start_num, num_chars)
text: The text string containing the characters you want to extract.<br />
start_num: The position of the first character you want to extract.<br />
num_chars: The number of characters you want MID to return.<br />

:blue_square:FIND: Finds one text value within another (case-sensitive).<br />
FIND(find_text, within_text, [start_num])
find_text: The text you want to find.
within_text: The text containing the text you want to find.
start_num: (Optional) The character number in within_text at which to start the search.<br />

:blue_square:SEARCH: Finds one text value within another (not case-sensitive).<br />
SEARCH(find_text, within_text, [start_num])
find_text: The text you want to find.
within_text: The text containing the text you want to find.
start_num: (Optional) The character number in within_text at which to start the search.<br />


#Comprehensive Guide to Excel Functions for Data Analysts and Data Scientists<br />
Excel is a powerful tool that offers a wide range of functions catering to different needs in data analysis, financial modeling, accounting, and more. Below, we categorize these functions to help you leverage Excel for your data projects.

1. Financial Functions
Excel provides numerous functions to handle financial calculations, which are essential for data analysts and financial modelers.

FV (Future Value): Calculates the future value of an investment.
PV (Present Value): Computes the present value of an investment.
NPV (Net Present Value): Calculates the net present value of an investment based on a series of periodic cash flows and a discount rate.
IRR (Internal Rate of Return): Determines the internal rate of return for a series of cash flows.
PMT (Payment): Calculates the payment for a loan based on constant payments and a constant interest rate.
RATE: Returns the interest rate per period of an annuity.<br />
2. Statistical Functions
These functions help in statistical analysis, allowing data analysts and scientists to perform various statistical calculations.

AVERAGE: Returns the average of a set of numbers.
MEDIAN: Finds the median of a set of numbers.
MODE: Determines the most frequently occurring number in a dataset.
STDEV.P and STDEV.S: Calculate the standard deviation for a population and a sample, respectively.
VAR.P and VAR.S: Compute the variance for a population and a sample, respectively.<br />
CORREL: Returns the correlation coefficient between two data sets.
LINEST: Returns the parameters of a linear trend.
3. Data Analysis Functions
Excel functions for data analysis help in managing, manipulating, and analyzing data sets.

SORT: Sorts the contents of a range or array.
FILTER: Filters a range of data based on criteria you define.
UNIQUE: Returns a list of unique values in a list or range.
XLOOKUP: Searches a range or an array and returns an item corresponding to the first match it finds.
SUMIF and SUMIFS: Sum cells based on a single or multiple criteria.
COUNTIF and COUNTIFS: Count cells based on single or multiple criteria.
4. Mathematical Functions
Essential mathematical functions for performing calculations and analyses.

SUM: Adds all the numbers in a range of cells.
PRODUCT: Multiplies all the numbers in a range of cells.
ABS: Returns the absolute value of a number.
ROUND, ROUNDUP, ROUNDDOWN: Rounds a number to a specified number of digits.
INT: Rounds a number down to the nearest integer.<br />
RAND and RANDBETWEEN: Generate random numbers.
5. Logical Functions
Logical functions are used to perform logical operations.

IF: Returns one value if a condition is true and another value if it is false.
AND: Returns TRUE if all arguments are TRUE.
OR: Returns TRUE if any argument is TRUE.
NOT: Reverses the logic of its argument.
IFERROR: Returns a value you specify if a formula evaluates to an error; otherwise, it returns the result of the formula.<br />
6. Text Functions
Text functions are used for string manipulation and text analysis.

CONCATENATE / CONCAT / TEXTJOIN: Combine multiple strings into one string.
LEFT, RIGHT, MID: Extract a specified number of characters from a string.
LEN: Returns the number of characters in a string.
TRIM: Removes all spaces from a text string except for single spaces between words.
UPPER, LOWER, PROPER: Convert text to uppercase, lowercase, or proper case.
7. Lookup and Reference Functions
These functions are used to search and reference data within Excel.

VLOOKUP: Looks for a value in the first column of a table and returns a value in the same row from a specified column.<br />
HLOOKUP: Searches for a value in the top row of a table and returns a value in the same column from a specified row.
MATCH: Searches for a specified item in a range of cells and returns the relative position of that item.
INDEX: Returns the value of an element in a table or array, selected by the row and column number indexes.
8. Date and Time Functions
Date and time functions are used to manipulate and perform calculations on dates and times.

TODAY: Returns the current date.
NOW: Returns the current date and time.
DATE: Returns the date given a year, month, and day.
DATEDIF: Calculates the difference between two dates.
EDATE: Returns the date that is the indicated number of months before or after the start date.<br />
EOMONTH: Returns the last day of the month, n months in the future or past.
9. Engineering Functions
Functions useful for engineering and complex calculations.<br />

COMPLEX: Converts real and imaginary coefficients into a complex number.<br />
IMSUM, IMPRODUCT: Perform arithmetic operations with complex numbers.
CONVERT: Converts a number from one measurement system to another.







Advanced Excel Functions for Data Analysts and Data Scientists
1. Financial Functions
Excel offers powerful financial functions to perform complex financial calculations.

XNPV: Calculates the net present value for a schedule of cash flows at specific dates.
XIRR: Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.<br />
CUMIPMT: Returns the cumulative interest paid on a loan between two periods.
CUMPRINC: Returns the cumulative principal paid on a loan between two periods.<br />
DB: Returns the depreciation of an asset for a specified period using the fixed-declining balance method.<br />
DDB: Returns the depreciation of an asset for a specified period using the double-declining balance method.<br />
SYD: Returns the sum-of-years' digits depreciation of an asset for a specified period.
VDB: Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or another method you specify.
ACCRINT: Returns the accrued interest for a security that pays periodic interest.
ACCRINTM: Returns the accrued interest for a security that pays interest at maturity.
2. Statistical Functions
Advanced statistical functions help in performing more complex statistical analyses.

FORECAST.ETS: Returns a future value based on existing (historical) values using the AAA version of the Exponential Smoothing (ETS) algorithm.
FORECAST.ETS.SEASONALITY: Returns the length of the repetitive pattern Excel detects for the specified time series.
FORECAST.ETS.CONFINT: Returns a confidence interval for the forecast value at a specified target date.<br />
FORECAST.ETS.STAT: Returns a statistical value as a result of time series forecasting.<br />
PERCENTILE.EXC: Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.<br />
PERCENTILE.INC: Returns the k-th percentile of values in a range.<br />
QUARTILE.EXC: Returns the quartile of the data set, based on percentile values from 0..1, exclusive.<br />
QUARTILE.INC: Returns the quartile of a data set.<br />
RANK.EQ: Returns the rank of a number in a list of numbers.<br />
RANK.AVG: Returns the rank of a number in a list of numbers.<br />
WEIBULL.DIST: Returns the Weibull distribution.<br />
GAMMA.DIST: Returns the gamma distribution.<br />
GAMMA.INV: Returns the inverse of the gamma cumulative distribution.<br />
LOGNORM.DIST: Returns the cumulative log-normal distribution.<br />
LOGNORM.INV: Returns the inverse of the log-normal cumulative distribution.<br />
NORM.DIST: Returns the normal distribution for a specified mean and standard deviation.<br />
NORM.INV: Returns the inverse of the normal cumulative distribution.<br />
T.DIST: Returns the Student's t-distribution.<br />
T.INV: Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom.<br />
F.DIST: Returns the F probability distribution.<br />
F.INV: Returns the inverse of the F probability distribution.<br />
3. Data Analysis Functions<br />
Advanced data analysis functions for handling complex datasets and performing sophisticated analyses.<br />

GETPIVOTDATA: Extracts data stored in a PivotTable report.<br />
CUBESET: Defines a calculated set of members or tuples by sending a set expression to the cube on the server.<br />
CUBEVALUE: Returns an aggregated value from a cube.<br />
CUBEMEMBER: Returns a member or tuple from the cube.<br />
CUBESETCOUNT: Returns the number of items in a set.<br />
CUBERANKEDMEMBER: Returns the nth, or ranked, member in a set.<br />
CUBEKPIMEMBER: Returns a key performance indicator (KPI) property and displays the KPI name in the cell.<br />
4. Mathematical and Trigonometric Functions<br />
Advanced mathematical and trigonometric functions for complex calculations.<br />

MROUND: Rounds a number to the nearest multiple of a specified value.<br />
CEILING.MATH: Rounds a number up, to the nearest integer or to the nearest multiple of significance.<br />
FLOOR.MATH: Rounds a number down, to the nearest integer or to the nearest multiple of significance.<br />
ERF.PRECISE: Returns the error function.<br />
ERFC.PRECISE: Returns the complementary ERF function.<br />
FACTDOUBLE: Returns the double factorial of a number.<br />
GCD: Returns the greatest common divisor.<br />
LCM: Returns the least common multiple.<br />
MDETERM: Returns the matrix determinant of an array.<br />
MINVERSE: Returns the matrix inverse of an array.<br />
MMULT: Returns the matrix product of two arrays.<br />
MULTINOMIAL: Returns the multinomial of a set of numbers.<br />
QUOTIENT: Returns the integer portion of a division.<br />
SQRTPI: Returns the square root of (number * pi).<br />
SUMPRODUCT: Returns the sum of the products of corresponding array components.<br />
SUMSQ: Returns the sum of the squares of the arguments.<br />
TEXT: Converts a value to text in a specific number format.<br />
UNICODE: Returns the number (code point) corresponding to the first character of the text.<br />
5. Logical Functions<br />
Advanced logical functions to perform more complex logical operations.<br />

IFNA: Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression.<br />
IFS: Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.<br />
SWITCH: Evaluates an expression against a list of values and returns the result corresponding to the first matching value.<br />
6. Text Functions<br />
Advanced text functions for handling and manipulating text strings.<br />

TEXTJOIN: Combines text from multiple ranges and/or strings, and includes a delimiter you specify between each text value that will be combined.<br />
UNICHAR: Returns the Unicode character that is referenced by the given numeric value.<br />
CONCAT: Combines multiple ranges and/or strings into one string.<br />
7. Lookup and Reference Functions<br />
Advanced lookup and reference functions for more sophisticated data retrieval.<br />

MATCH: Searches for a specified item in a range of cells and returns the relative position of that item.<br />
INDEX: Returns the value of an element in a table or array, selected by the row and column number indexes.<br />
CHOOSE: Returns a value from a list of values, based on an index number.<br />
HLOOKUP: Searches for a value in the top row of a table and returns a value in the same column from a specified row.<br />
VLOOKUP: Searches for a value in the first column of a table and returns a value in the same row from a specified column.<br />
XLOOKUP: Searches a range or an array, and returns an item corresponding to the first match it finds. If a match doesn’t exist, then XLOOKUP can return the closest (approximate) match.<br />
8. Date and Time Functions<br />
Advanced date and time functions for sophisticated date and time manipulations.<br />

NETWORKDAYS.INTL: Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days.<br />
WORKDAY.INTL: Returns the serial number of the date before or after a specified number of workdays with custom weekends.<br />
ISO.WEEKNUM: Returns the number of the ISO week number of the year for a given date.<br />
YEARFRAC: Returns the year fraction representing the number of whole days between start_date and end_date.<br />
9. Engineering Functions<br />
Advanced engineering functions for complex engineering calculations.<br />

IMDIV: Returns the quotient of two complex numbers.<br />
IMPOWER: Returns a complex number raised to an integer power.<br />
IMSQRT: Returns the square root of a complex number.<br />
10. Information Functions<br />
Advanced information functions for more complex data handling.<br />

CELL: Returns information about the formatting, location, or contents of a cell.<br />
INFO: Returns information about the current operating environment.<br />
TYPE: Returns a number indicating the data type<br />

