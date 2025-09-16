# Lesson 1. Formulas and Functions. References and Arrays

## Mathematical Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `ABS` | Returns the absolute value of a number | `=ABS(-5)` → 5 |
| `PI` | Returns the value of π (3.14159) | `=PI()` → 3.141592654 |
| `DEGREES` | Converts radians to degrees | `=DEGREES(PI())` → 180 |
| `RADIANS` | Converts degrees to radians | `=RADIANS(180)` → 3.141592654 |
| `SIN` | Returns the sine of an angle (in radians) | `=SIN(RADIANS(90))` → 1 |
| `COS` | Returns the cosine of an angle (in radians) | `=COS(RADIANS(0))` → 1 |
| `TAN` | Returns the tangent of an angle (in radians) | `=TAN(RADIANS(45))` → 1 |
| `COT` | Returns the cotangent of an angle (in radians) | `=COT(RADIANS(45))` → 1 |
| `ASIN` | Returns the arcsine of a number, in radians | `=DEGREES(ASIN(1))` → 90 |
| `ACOS` | Returns the arccosine of a number, in radians | `=DEGREES(ACOS(0))` → 90 |
| `ATAN` | Returns the arctangent of a number, in radians | `=DEGREES(ATAN(1))` → 45 |
| `EXP` | Returns *e* raised to the power of a number | `=EXP(1)` → 2.718281828 |
| `LN` | Returns the natural logarithm of a number | `=LN(EXP(1))` → 1 |
| `LOG` | Returns the logarithm of a number with a specified base | `=LOG(8,2)` → 3 |
| `LOG10` | Returns the base-10 logarithm of a number | `=LOG10(1000)` → 3 |
| `SIGN` | Returns the sign of a number (1, -1, or 0) | `=SIGN(-25)` → -1 |
| `SQRT` | Returns the square root of a number | `=SQRT(16)` → 4 |
| `POWER` | Returns a number raised to a power | `=POWER(2,3)` → 8 |
| `INT` | Rounds a number down to the nearest integer | `=INT(5.9)` → 5 |
| `QUOTIENT` | Returns the integer portion of a division | `=QUOTIENT(10,3)` → 3 |

## Statistical Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `AVERAGE` | Returns the arithmetic mean | `=AVERAGE(10,20,30)` → 20 |
| `AVERAGEIF` | Returns the average of cells that meet a condition | `=AVERAGEIF(B2:B5,"<23000")` |
| `COUNT` | Counts numbers in a range | `=COUNT(A1:A5)` |
| `COUNTIF` | Counts cells that meet a condition | `=COUNTIF(A1:A5,">10")` |
| `MAX` | Returns the maximum value | `=MAX(1,3,7)` → 7 |
| `MIN` | Returns the minimum value | `=MIN(1,3,7)` → 1 |
| `MEDIAN` | Returns the median | `=MEDIAN(1,5,100)` → 5 |
| `MODE.SNGL` | Returns the most frequent value | `=MODE.SNGL(1,2,2,3)` → 2 |
| `CORREL` | Correlation coefficient | `=CORREL(A1:A12,B1:B12)` |

## Logical Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `IF` | Returns value depending on condition | `=IF(A1>10,"Yes","No")` |
| `AND` | Returns TRUE if all conditions are met | `=AND(A1>0,B1<5)` |
| `OR` | Returns TRUE if any condition is met | `=OR(A1>0,B1<5)` |
| `NOT` | Reverses the logical value | `=NOT(A1>100)` |
| `TRUE` | Returns TRUE | `=TRUE()` |
| `FALSE` | Returns FALSE | `=FALSE()` |
| `IFERROR` | Returns custom value if error | `=IFERROR(A1/B1,"Error")` |

## Text Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `CHAR` | Returns character by code | `=CHAR(65)` → "A" |
| `CODE` | Returns code of first character | `=CODE("A")` → 65 |
| `CONCATENATE` | Joins text strings | `=CONCATENATE("Hello"," ","World")` |
| `EXACT` | Checks if two texts are identical | `=EXACT("a","A")` → FALSE |
| `FIND` | Finds substring (case-sensitive) | `=FIND("cat","Concatenate")` → 4 |
| `SEARCH` | Finds substring (not case-sensitive) | `=SEARCH("CAT","Concatenate")` → 4 |
| `LEFT` | Returns leftmost characters | `=LEFT("Excel",2)` → "Ex" |
| `RIGHT` | Returns rightmost characters | `=RIGHT("Excel",2)` → "el" |
| `LEN` | Returns length of text | `=LEN("Excel")` → 5 |
| `LOWER` | Converts to lowercase | `=LOWER("Excel")` → "excel" |
| `UPPER` | Converts to uppercase | `=UPPER("Excel")` → "EXCEL" |
| `PROPER` | Capitalizes first letter of each word | `=PROPER("excel formulas")` → "Excel Formulas" |
| `MID` | Returns substring | `=MID("Excel",2,3)` → "xce" |
| `REPLACE` | Replaces part of text | `=REPLACE("Excel",2,3,"123")` → "E123l" |
| `SUBSTITUTE` | Replaces text | `=SUBSTITUTE("Excel 2025","2025","365")` → "Excel 365" |
| `REPT` | Repeats text | `=REPT("A",3)` → "AAA" |
| `TRIM` | Removes extra spaces | `=TRIM(" A B ")` → "A B" |
| `TEXT` | Formats numbers | `=TEXT(1234.5,"#,##0.00 ₽")` → "1 234,50 ₽" |

## Date and Time Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `TODAY` | Current date | `=TODAY()` → 16.09.2025 |
| `NOW` |  Current date and time | `=NOW()` → 16.09.2025 10:25 |
| `DATE` | Creates a date | `=DATE(2025,9,16)` |
| `TIME` | Creates a time value | `=TIME(10,30,0)` → 10:30 |
| `YEAR` | Returns year | `=YEAR(TODAY())` → 2025 |
| `MONTH` | Returns month | `=MONTH(TODAY())` → 9 |
| `DAY` | Returns day | `=DAY(TODAY())` → 16 |
| `HOUR` | Returns hour | `=HOUR(NOW())` |
| `MINUTE` | Returns minutes | `=MINUTE(NOW())` |
| `SECOND` | Returns seconds | `=SECOND(NOW())` |
| `WEEKNUM` | Returns week number | `=WEEKNUM(TODAY())` |
| `DAYS` | Days between dates | `=DAYS("31-Dec-2025","01-Jan-2025")` → 364 |

## Lookup & Reference Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `VLOOKUP` | Vertical lookup | `=VLOOKUP(25,A2:C10,2,FALSE)` |
| `HLOOKUP` | Horizontal lookup | `=HLOOKUP(100,B1:F3,2,FALSE)` |
| `ROW` | Returns row number | `=ROW(D11)` → 11 |
| `COLUMN` | Returns column number | `=COLUMN(D11)` → 4 |
| `ADDRESS` | Returns cell address | `=ADDRESS(2,3)` → $C$2 |
| `INDIRECT` | Reference from text | `=INDIRECT("B2")` → value in B2 |
| `ROWS` | Number of rows in array | `=ROWS(A1:C5)` → 5 |
| `COLUMNS` | Number of columns in array | `=COLUMNS(A1:C5)` → 3 |
| `OFFSET` | Returns a shifted reference | `=OFFSET(A1,1,2)` → cell C2 |
| `MATCH` | Returns position in array | `=MATCH(25,A1:A3,0)` → 2 |
| `TRANSPOSE` | Transposes array | `=TRANSPOSE(A1:B2)` |

## Matrix Functions

| Function | Description | Example |
|---------------|-------------|---------|
| `MMULT` | Matrix multiplication | `=MMULT({1,2;3,4},{5;6})` → {17;39} |
| `MINVERSE` | Inverse of a square matrix | `=MINVERSE({1,2;3,4})` |
| `MDETERM` | Determinant of a matrix | `=MDETERM({1,2;3,4})` → -2 |

