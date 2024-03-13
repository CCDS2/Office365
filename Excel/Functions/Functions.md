1. **SUM**: Adds up numbers.
   ```excel
   =SUM(A1:A5)
   ```
   Example: If cells A1 to A5 contain the numbers 10, 20, 30, 40, and 50 respectively, the formula will return the sum of these numbers, which is 150.

2. **AVERAGE**: Calculates the average of numbers.
   ```excel
   =AVERAGE(B2:B10)
   ```
   Example: If cells B2 to B10 contain the numbers 15, 20, 25, 30, 35, 40, 45, 50, and 55 respectively, the formula will return the average of these numbers, which is 35.

3. **MAX**: Returns the largest number in a range.
   ```excel
   =MAX(C2:C20)
   ```
   Example: If cells C2 to C20 contain various numbers, the formula will return the largest number among them.

4. **MIN**: Returns the smallest number in a range.
   ```excel
   =MIN(D1:D15)
   ```
   Example: If cells D1 to D15 contain various numbers, the formula will return the smallest number among them.

5. **COUNT**: Counts the number of cells that contain numbers.
   ```excel
   =COUNT(E1:E100)
   ```
   Example: If cells E1 to E100 contain a mix of numbers and empty cells, the formula will return the count of cells containing numbers.

6. **IF**: Performs a conditional test and returns one value if the condition is true and another value if the condition is false.
   ```excel
   =IF(F2>10, "Pass", "Fail")
   ```
   Example: If the value in cell F2 is greater than 10, the formula will return "Pass"; otherwise, it will return "Fail".

7. **VLOOKUP**: Searches for a value in the first column of a table array and returns a value in the same row from another column.
   ```excel
   =VLOOKUP(G2, A2:B10, 2, FALSE)
   ```
   Example: If cell G2 contains a value to search for, the formula will search for that value in column A of the range A2:B10. If found, it will return the corresponding value from column B of the same row.

8. **HLOOKUP**: Searches for a value in the first row of a table array and returns a value in the same column from another row.
   ```excel
   =HLOOKUP(H2, A1:D5, 3, FALSE)
   ```
   Example: If cell H2 contains a value to search for, the formula will search for that value in row 1 of the range A1:D5. If found, it will return the corresponding value from column 3 of the same column.

9. **INDEX**: Returns the value of a cell in a specified row and column of a table array.
   ```excel
   =INDEX(A1:E10, 3, 2)
   ```
   Example: The formula will return the value in the cell at the intersection of the 3rd row and 2nd column of the range A1:E10.

10. **MATCH**: Searches for a specified value in a range and returns the relative position of that item.
    ```excel
    =MATCH(I2, F2:F20, 0)
    ```
    Example: If cell I2 contains a value to search for, the formula will search for that value in the range F2:F20. If found, it will return the relative position of the item in the range.

11. **CONCATENATE**: Joins several text strings into one string.
    ```excel
    =CONCATENATE("Hello", " ", "World")
    ```
    Example: The formula will concatenate the text strings "Hello", a space, and "World" to form the text string "Hello World".

12. **LEFT**: Returns the leftmost characters from a text string.
    ```excel
    =LEFT(J2, 5)
    ```
    Example: If cell J2 contains the text string "Excel is great", the formula will return the leftmost 5 characters of the string, which is "Excel".

13. **RIGHT**: Returns the rightmost characters from a text string.
    ```excel
    =RIGHT(K2, 3)
    ```
    Example: If cell K2 contains the text string "Excel is great", the formula will return the rightmost 3 characters of the string, which is "eat".

14. **MID**: Returns a specific number of characters from a text string, starting at the position you specify.
    ```excel
    =MID(L2, 3, 2)
    ```
    Example: If cell L2 contains the text string "Excel is great", the formula will return 2 characters starting from the 3rd position, which is "ce".

15. **LEN**: Returns the number of characters in a text string.
    ```excel
    =LEN(M2)
    ```
    Example: If cell M2 contains the text string "Excel is great", the formula will return the number 13, which is the total number of characters in the string.

16. **DATE**: Returns the serial number of a particular date.
    ```excel
    =DATE(2024, 3, 13)
    ```
    Example: The formula will return the serial number representing the date March 13, 2024.

17. **TODAY**: Returns the current date.


    ```excel
    =TODAY()
    ```
    Example: The formula will return the current date.

18. **NOW**: Returns the current date and time.
    ```excel
    =NOW()
    ```
    Example: The formula will return the current date and time.

19. **ROUND**: Rounds a number to a specified number of digits.
    ```excel
    =ROUND(N2, 2)
    ```
    Example: If cell N2 contains a number, the formula will round that number to 2 decimal places.

20. **ROUNDUP**: Rounds a number up, away from zero, to the nearest multiple of significance.
    ```excel
    =ROUNDUP(O2, -1)
    ```
    Example: If cell O2 contains a number, the formula will round that number up to the nearest multiple of 10.

21. **ROUNDDOWN**: Rounds a number down, toward zero, to the nearest multiple of significance.
    ```excel
    =ROUNDDOWN(P2, -2)
    ```
    Example: If cell P2 contains a number, the formula will round that number down to the nearest multiple of 100.

22. **TRIM**: Removes extra spaces from text, except for single spaces between words.
    ```excel
    =TRIM(Q2)
    ```
    Example: If cell Q2 contains the text string "   Excel   is   great   ", the formula will return "Excel is great" by removing the extra spaces.

23. **UPPER**: Converts text to uppercase.
    ```excel
    =UPPER(R2)
    ```
    Example: If cell R2 contains the text string "excel", the formula will return "EXCEL".

24. **LOWER**: Converts text to lowercase.
    ```excel
    =LOWER(S2)
    ```
    Example: If cell S2 contains the text string "EXCEL", the formula will return "excel".

25. **PROPER**: Capitalizes the first letter of each word in a text string.
    ```excel
    =PROPER(T2)
    ```
    Example: If cell T2 contains the text string "excel is great", the formula will return "Excel Is Great".

26. **ABS**: Returns the absolute value of a number.
    ```excel
    =ABS(U2)
    ```
    Example: If cell U2 contains the number -10, the formula will return 10.

27. **SQRT**: Returns the square root of a number.
    ```excel
    =SQRT(V2)
    ```
    Example: If cell V2 contains the number 16, the formula will return 4.

28. **LOG**: Returns the logarithm of a number to the base you specify.
    ```excel
    =LOG(W2, 10)
    ```
    Example: If cell W2 contains the number 100, the formula will return 2 since log base 10 of 100 is 2.

29. **EXP**: Returns e raised to the power of a given number.
    ```excel
    =EXP(X2)
    ```
    Example: If cell X2 contains the number 2, the formula will return e^2, which is approximately 7.389.

30. **SIN**: Returns the sine of an angle.
    ```excel
    =SIN(Y2)
    ```
    Example: If cell Y2 contains the angle in radians, the formula will return the sine of that angle.

31. **COS**: Returns the cosine of an angle.
    ```excel
    =COS(Z2)
    ```
    Example: If cell Z2 contains the angle in radians, the formula will return the cosine of that angle.

32. **TAN**: Returns the tangent of an angle.
    ```excel
    =TAN(AA2)
    ```
    Example: If cell AA2 contains the angle in radians, the formula will return the tangent of that angle.

33. **PI**: Returns the value of pi (π).
    ```excel
    =PI()
    ```
    Example: The formula will return the value of pi, which is approximately 3.14159265358979.

34. **RAND**: Returns a random number between 0 and 1.
    ```excel
    =RAND()
    ```
    Example: The formula will return a random number between 0 and 1.

35. **RANDBETWEEN**: Returns a random integer between the numbers you specify.
    ```excel
    =RANDBETWEEN(1, 100)
    ```
    Example: The formula will return a random integer between 1 and 100.

36. **AND**: Returns TRUE if all of its arguments are TRUE.
    ```excel
    =AND(AB2=5, AC2="Yes")
    ```
    Example: The formula will return TRUE if cell AB2 contains the value 5 and cell AC2 contains the text "Yes".

37. **OR**: Returns TRUE if any argument is TRUE.
    ```excel
    =OR(AD2="Pass", AD2="Fail")
    ```
    Example: The formula will return TRUE if cell AD2 contains either the text "Pass" or "Fail".

38. **NOT**: Reverses the logical value of its argument.
    ```excel
    =NOT(AE2="Error")
    ```
    Example: The formula will return TRUE if cell AE2 does not contain the text "Error".

39. **IFERROR**: Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.
    ```excel
    =IFERROR(AF2/AH2, "Division by zero")
    ```
    Example: The formula will return the result of the division of cell AF2 by AH2, unless an error occurs, in which case it will return the text "Division by zero".

40. **SUMIF**: Adds the cells specified by a given criteria.
    ```excel
    =SUMIF(AI2:AI100, ">10")
    ```
    Example: The formula will sum the values in the range AI2:AI100 that are greater than 10.

41. **COUNTIF**: Counts the number of cells specified by a given criteria.
    ```excel
    =COUNTIF(AJ2:AJ100, "<>0")
    ```
    Example: The formula will count the number of non-zero values in the range AJ2:AJ100.

42. **AVERAGEIF**: Calculates the average of cells specified by a given criteria.
    ```excel
    =AVERAGEIF(AK2:AK100, ">=50")
    ```
    Example: The formula will calculate the average of the values in the range AK2:AK100 that are greater than or equal to 50.

43. **SUMIFS**: Adds the cells in a range that meet multiple criteria.
    ```excel
    =SUMIFS(AL2:AL100, AM2:AM100, ">200", AN2:AN100, "<500")
    ```
    Example: The formula will sum the values in the range AL2:AL100 that are greater than 200 and less than 500, based on the corresponding conditions in the ranges AM2:AM100 and AN2:AN100.

44. **COUNTIFS**: Counts the number of cells in a range that meet multiple criteria.
    ```excel
    =COUNTIFS(AO2

:AO100, ">=50", AP2:AP100, "<=100")
    ```
    Example: The formula will count the number of cells in the range AO2:AO100 that contain values greater than or equal to 50 and less than or equal to 100.

45. **AVERAGEIFS**: Calculates the average of cells in a range that meet multiple criteria.
    ```excel
    =AVERAGEIFS(AQ2:AQ100, AR2:AR100, ">50", AS2:AS100, "<=100")
    ```
    Example: The formula will calculate the average of the values in the range AQ2:AQ100 that are greater than 50 and less than or equal to 100, based on the corresponding conditions in the ranges AR2:AR100 and AS2:AS100.

46. **SUBTOTAL**: Returns a subtotal in a list or database.
    ```excel
    =SUBTOTAL(109, AT2:AT100)
    ```
    Example: The formula will return a subtotal of the values in the range AT2:AT100 using function number 109, which represents the sum function.

47. **IFNA**: Returns a value you specify if a formula results in the #N/A error; otherwise, returns the result of the formula.
   ```excel
   =IFNA(AU2/AW2, "N/A")
   ```
   Example: The formula will return the result of the division of cell AU2 by AW2 unless it results in the #N/A error, in which case it will return "N/A".

48. **INDIRECT**: Returns the reference specified by a text string.
   ```excel
   =INDIRECT("B2")
   ```
   Example: The formula will return the value of the cell specified by the text string "B2".

49. **OFFSET**: Returns a reference offset from a starting cell or range of cells.
   ```excel
   =OFFSET(A1, 3, 2)
   ```
   Example: The formula will return a reference to the cell that is 3 rows down and 2 columns to the right of cell A1.

50. **ROW**: Returns the row number of a reference.
   ```excel
   =ROW(A2)
   ```
   Example: The formula will return the row number of cell A2.

51. **COLUMN**: Returns the column number of a reference.
   ```excel
   =COLUMN(B3)
   ```
   Example: The formula will return the column number of cell B3.

52. **INDEX-MATCH (Combined)**: Combining INDEX and MATCH functions for flexible lookups.
   ```excel
   =INDEX(B2:B10, MATCH(F2, A2:A10, 0))
   ```
   Example: The formula will return the value from column B corresponding to the value in cell F2 by searching for it in column A.

53. **TEXT**: Formats a number and converts it to text.
   ```excel
   =TEXT(C2, "yyyy/mm/dd")
   ```
   Example: The formula will convert the value in cell C2 to text format with the specified date format.

54. **NOW()**: Returns the current date and time.
   ```excel
   =NOW()
   ```
   Example: The formula will return the current date and time.

55. **TODAY()**: Returns the current date.
   ```excel
   =TODAY()
   ```
   Example: The formula will return the current date.

56. **RANDARRAY**: Generates an array of random numbers between 0 and 1.
   ```excel
   =RANDARRAY(5, 5)
   ```
   Example: The formula will generate a 5x5 array of random numbers between 0 and 1.

57. **SEQUENCE**: Generates a sequence of numbers in an array.
   ```excel
   =SEQUENCE(5, 1, 0, 2)
   ```
   Example: The formula will generate a sequence of 5 numbers starting from 0 with an increment of 2.

58. **TRANSPOSE**: Transposes the rows and columns of a range or array.
   ```excel
   =TRANSPOSE(A1:D4)
   ```
   Example: The formula will transpose the range A1:D4, converting rows into columns and vice versa.

59. **FILTER**: Filters a range of data based on criteria you define.
   ```excel
   =FILTER(A2:A100, B2:B100>50)
   ```
   Example: The formula will filter the range A2:A100 based on the condition that values in the corresponding cells in range B2:B100 are greater than 50.

60. **SORT**: Sorts the contents of a range or array.
   ```excel
   =SORT(A2:A100, 1, TRUE)
   ```
   Example: The formula will sort the range A2:A100 in ascending order based on the values in the first column.

61. **TEXTJOIN**: Joins multiple text strings into one string, with a specified delimiter.
   ```excel
   =TEXTJOIN(", ", TRUE, C2:C10)
   ```
   Example: The formula will join the text strings in the range C2:C10 with commas as delimiters, ignoring empty cells.

62. **UNIQUE**: Returns a list of unique values from a range or array.
   ```excel
   =UNIQUE(D2:D100)
   ```
   Example: The formula will return a list of unique values from the range D2:D100.

63. **SPLIT**: Splits text into separate cells based on a delimiter.
   ```excel
   =SPLIT(E2, ",")
   ```
   Example: The formula will split the text in cell E2 into separate cells using comma as the delimiter.

64. **SUBSTITUTE**: Substitutes new text for old text in a text string.
   ```excel
   =SUBSTITUTE(F2, "old", "new")
   ```
   Example: The formula will replace occurrences of the text "old" with "new" in the text string in cell F2.

65. **REPT**: Repeats text a given number of times.
   ```excel
   =REPT(G2, 3)
   ```
   Example: The formula will repeat the text in cell G2 three times.

66. **SEARCH**: Finds one text

 string within another (case-sensitive).
   ```excel
   =SEARCH(H2, "Excel is great")
   ```
   Example: The formula will search for the text in cell H2 within the string "Excel is great" and return the starting position of the text.

67. **FIND**: Finds one text string within another (case-sensitive).
   ```excel
   =FIND(I2, "Excel is great")
   ```
   Example: The formula will find the position of the text in cell I2 within the string "Excel is great".

68. **LEFT (with FIND)**: Extracts a specified number of characters from the start of a text string.
   ```excel
   =LEFT(J2, FIND(" ", J2)-1)
   ```
   Example: The formula will extract characters from the start of the text in cell J2 until the first space.

69. **RIGHT (with LEN and FIND)**: Extracts a specified number of characters from the end of a text string.
   ```excel
   =RIGHT(K2, LEN(K2)-FIND(" ", K2))
   ```
   Example: The formula will extract characters from the end of the text in cell K2 starting from the position after the first space.

70. **MID (with FIND)**: Extracts a specific number of characters from a text string, starting at a specified position.
   ```excel
   =MID(L2, FIND(" ", L2)+1, FIND(" ", L2, FIND(" ", L2)+1)-FIND(" ", L2)-1)
   ```
   Example: The formula will extract characters from the text in cell L2 starting after the first space and ending before the second space.

71. **LEN (with SUBSTITUTE)**: Returns the length of a text string after replacing a specific character.
   ```excel
   =LEN(SUBSTITUTE(M2, " ", ""))
   ```
   Example: The formula will return the length of the text string in cell M2 after removing all spaces.

72. **CONCAT**: Joins several text strings into one string.
   ```excel
   =CONCAT(N2, " ", O2)
   ```
   Example: The formula will concatenate the text in cells N2 and O2 with a space between them.

73. **CONCATENATE (with IF)**: Joins several text strings into one string conditionally.
   ```excel
   =CONCATENATE(IF(P2="Yes", "Pass", "Fail"), " - ", Q2)
   ```
   Example: The formula will concatenate "Pass" or "Fail" based on the value in cell P2 with a hyphen and the text in cell Q2.

74. **CONCATENATE (with IF and TEXT)**: Joins several text strings into one string conditionally.
   ```excel
   =CONCATENATE(IF(R2>0, "Profit: $", "Loss: $"), TEXT(R2, "#,##0.00"))
   ```
   Example: The formula will concatenate "Profit: $" or "Loss: $" based on the value in cell R2, followed by the formatted value of R2.

75. **INDEX-MATCH-MATCH (Nested)**: Combining INDEX and MATCH functions for two-dimensional lookups.
   ```excel
   =INDEX(B2:F10, MATCH(H2, A2:A10, 0), MATCH(I2, B1:F1, 0))
   ```
   Example: The formula will return the value from a 2D range (B2:F10) based on matching row and column headers.

76. **INDIRECT (with TEXT)**: Returns the reference specified by a text string that is formatted as a cell reference.
   ```excel
   =INDIRECT(TEXT(S2, "R0C00"), FALSE)
   ```
   Example: The formula will return the value of the cell referenced by the row number in S2.

77. **CELL (with ROW and COLUMN)**: Returns information about a cell, such as its file name, sheet name, row number, and column number.
   ```excel
   =CELL("address", T2)
   ```
   Example: The formula will return the address of the cell in T2.

78. **GETPIVOTDATA**: Returns data stored in a PivotTable report.
   ```excel
   =GETPIVOTDATA("Sales", PivotTable1, "Region", "East", "Product", "Widget")
   ```
   Example: The formula will return the sales value from a PivotTable where the region is "East" and the product is "Widget".

79. **HYPERLINK**: Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the internet.
   ```excel
   =HYPERLINK("http://www.example.com", "Click here")
   ```
   Example: The formula will create a hyperlink to the URL "http://www.example.com" with the display text "Click here".Certainly! Here are bigger examples/explanations for each of the 100 Excel functions:



82. **WORKDAY**: Returns a date that is a specified number of working days before or after a given date.
   ```excel
   =WORKDAY(Y2, 5)
   ```
   Example: The formula will return a date that is 5 working days after the date in cell Y2.

83. **WEEKNUM**: Returns the week number in the year for a given date.
   ```excel
   =WEEKNUM(Z2)
   ```
   Example: The formula will return the week number for the date in cell Z2.

84. **EOMONTH**: Returns the serial number of the last day of the month before or after a specified number of months.
   ```excel
   =EOMONTH(AA2, 3)
   ```
   Example: The formula will return the serial number of the last day of the month, 3 months after the date in cell AA2.

85. **YEARFRAC**: Returns the fraction of the year represented by the number of whole days between two dates.
   ```excel
   =YEARFRAC(AB2, AC2)
   ```
   Example: The formula will return the fraction of the year between the dates in cells AB2 and AC2.

86. **RATE**: Calculates the interest rate per period of an annuity.
   ```excel
   =RATE(10, -1000, 5000)
   ```
   Example: The formula will calculate the interest rate per period for an annuity with 10 periods, a present value of -$1,000, and a future value of $5,000.

87. **NPER**: Calculates the number of periods for an investment based on periodic, constant payments and a constant interest rate.
   ```excel
   =NPER(0.06/12, -200, 10000)
   ```
   Example

: The formula will calculate the number of periods required for an investment with a monthly payment of $200, a present value of -$10,000, and a monthly interest rate of 6%.

88. **PV**: Calculates the present value of an investment based on a constant interest rate and periodic, constant payments.
   ```excel
   =PV(0.05, 5, 1000, , 1)
   ```
   Example: The formula will calculate the present value of an investment that pays $1,000 at the end of each period for 5 periods with a 5% annual interest rate.

89. **FV**: Calculates the future value of an investment based on a constant interest rate and periodic, constant payments.
   ```excel
   =FV(0.06, 5, -100, , 1)
   ```
   Example: The formula will calculate the future value of an investment with $100 payments at the beginning of each period for 5 periods with a 6% annual interest rate.

90. **NPV**: Calculates the net present value of an investment based on a series of periodic cash flows and a discount rate.
   ```excel
   =NPV(0.05, D2:D10) + C2
   ```
   Example: The formula will calculate the net present value of a series of cash flows (in cells D2:D10) using a discount rate of 5% and then add the initial investment amount in cell C2.

91. **IRR**: Calculates the internal rate of return for a series of cash flows that occur at regular intervals.
   ```excel
   =IRR(E2:E10)
   ```
   Example: The formula will calculate the internal rate of return for a series of cash flows in cells E2:E10.

92. **MIRR**: Calculates the modified internal rate of return for a series of cash flows, considering both the cost of the investment and the interest earned on reinvestment of cash.
   ```excel
   =MIRR(F2:F10, 0.05, 0.07)
   ```
   Example: The formula will calculate the modified internal rate of return for a series of cash flows in cells F2:F10, with a financing rate of 5% and a reinvestment rate of 7%.

93. **XNPV**: Calculates the net present value of an investment with irregular cash flows, using specific dates and a discount rate.
   ```excel
   =XNPV(0.1, G2:G10, H2:H10)
   ```
   Example: The formula will calculate the net present value of a series of cash flows in cells G2:G10, with corresponding dates in cells H2:H10, using a discount rate of 10%.

94. **XIRR**: Calculates the internal rate of return for a series of cash flows that occur at irregular intervals.
   ```excel
   =XIRR(I2:I10, J2:J10)
   ```
   Example: The formula will calculate the internal rate of return for a series of cash flows in cells I2:I10, with corresponding dates in cells J2:J10.

95. **DURATION**: Calculates the Macaulay duration of a security with an assumed par value of $100.
   ```excel
   =DURATION(0.05, 0.06, 0.08, 0.07, 2)
   ```
   Example: The formula will calculate the Macaulay duration of a security with a 5% yield, a 6% annual coupon rate, an 8% annual yield to maturity, a 7% annual reinvestment rate, and 2 years to maturity.

96. **MDURATION**: Calculates the modified Macaulay duration of a security with an assumed par value of $100.
   ```excel
   =MDURATION(0.05, 0.06, 0.08, 0.07, 2)
   ```
   Example: The formula will calculate the modified Macaulay duration of a security with a 5% yield, a 6% annual coupon rate, an 8% annual yield to maturity, a 7% annual reinvestment rate, and 2 years to maturity.

97. **COUPDAYS**: Calculates the number of days from the beginning of the coupon period to the settlement date.
   ```excel
   =COUPDAYS("1/1/2024", "6/30/2024", 2, TRUE)
   ```
   Example: The formula will calculate the number of days from January 1, 2024, to June 30, 2024, with semiannual coupons and U.S. (NASD) basis.

98. **COUPDAYSNC**: Calculates the number of days from the settlement date to the next coupon date.
   ```excel
   =COUPDAYSNC("1/1/2024", "6/30/2024", 2, TRUE)
   ```
   Example: The formula will calculate the number of days from January 1, 2024, to the next coupon date, June 30, 2024, with semiannual coupons and U.S. (NASD) basis.

99. **COUPNUM**: Calculates the number of coupons payable between the settlement date and maturity date, inclusive.
   ```excel
   =COUPNUM("1/1/2024", "12/31/2024", 2, TRUE)
   ```
   Example: The formula will calculate the number of coupons payable between January 1, 2024, and December 31, 2024, with semiannual coupons and U.S. (NASD) basis.

100. **COUPPCD**: Calculates the previous coupon date before the settlement date.
    ```excel
    =COUPPCD("1/1/2024", "6/30/2024", 2, TRUE)
    ```
    Example: The formula will calculate the previous coupon date before June 30, 2024, with semiannual coupons and U.S. (NASD) basis.

These Excel functions and formulas provide powerful tools for data analysis, financial calculations, text manipulation, and more. By understanding and mastering these functions, you can significantly improve your efficiency and effectiveness in Excel.

80. **DATEDIF**: Calculates the difference between two dates.
   ```excel
   =DATEDIF(U2, V2, "D")
   ```
   Example: The formula will calculate the difference in days between the dates in cells U2 and V2.

81. **NETWORKDAYS**: Returns the number of whole working days between two dates.
   ```excel
   =NETWORKDAYS(W2, X2)
   ```
   Example: The formula will return the number of working days between the dates in cells W2 and X2.




    
