

1. **SUM**: Adds up numbers.
   ```
   =SUM(A1:A5)
   ```

2. **AVERAGE**: Calculates the average of numbers.
   ```
   =AVERAGE(B2:B10)
   ```

3. **MAX**: Returns the largest number in a range.
   ```
   =MAX(C2:C20)
   ```

4. **MIN**: Returns the smallest number in a range.
   ```
   =MIN(D1:D15)
   ```

5. **COUNT**: Counts the number of cells that contain numbers.
   ```
   =COUNT(E1:E100)
   ```

6. **IF**: Performs a conditional test and returns one value if the condition is true and another value if the condition is false.
   ```
   =IF(F2>10, "Pass", "Fail")
   ```

7. **VLOOKUP**: Searches for a value in the first column of a table array and returns a value in the same row from another column.
   ```
   =VLOOKUP(G2, A2:B10, 2, FALSE)
   ```

8. **HLOOKUP**: Searches for a value in the first row of a table array and returns a value in the same column from another row.
   ```
   =HLOOKUP(H2, A1:D5, 3, FALSE)
   ```

9. **INDEX**: Returns the value of a cell in a specified row and column of a table array.
   ```
   =INDEX(A1:E10, 3, 2)
   ```

10. **MATCH**: Searches for a specified value in a range and returns the relative position of that item.
    ```
    =MATCH(I2, F2:F20, 0)
    ```

11. **CONCATENATE**: Joins several text strings into one string.
    ```
    =CONCATENATE("Hello", " ", "World")
    ```

12. **LEFT**: Returns the leftmost characters from a text string.
    ```
    =LEFT(J2, 5)
    ```

13. **RIGHT**: Returns the rightmost characters from a text string.
    ```
    =RIGHT(K2, 3)
    ```

14. **MID**: Returns a specific number of characters from a text string, starting at the position you specify.
    ```
    =MID(L2, 3, 2)
    ```

15. **LEN**: Returns the number of characters in a text string.
    ```
    =LEN(M2)
    ```

16. **DATE**: Returns the serial number of a particular date.
    ```
    =DATE(2024, 3, 13)
    ```

17. **TODAY**: Returns the current date.
    ```
    =TODAY()
    ```

18. **NOW**: Returns the current date and time.
    ```
    =NOW()
    ```

19. **ROUND**: Rounds a number to a specified number of digits.
    ```
    =ROUND(N2, 2)
    ```

20. **ROUNDUP**: Rounds a number up, away from zero, to the nearest multiple of significance.
    ```
    =ROUNDUP(O2, -1)
    ```

21. **ROUNDDOWN**: Rounds a number down, toward zero, to the nearest multiple of significance.
    ```
    =ROUNDDOWN(P2, -2)
    ```

22. **TRIM**: Removes extra spaces from text, except for single spaces between words.
    ```
    =TRIM(Q2)
    ```

23. **UPPER**: Converts text to uppercase.
    ```
    =UPPER(R2)
    ```

24. **LOWER**: Converts text to lowercase.
    ```
    =LOWER(S2)
    ```

25. **PROPER**: Capitalizes the first letter of each word in a text string.
    ```
    =PROPER(T2)
    ```

26. **ABS**: Returns the absolute value of a number.
    ```
    =ABS(U2)
    ```

27. **SQRT**: Returns the square root of a number.
    ```
    =SQRT(V2)
    ```

28. **LOG**: Returns the logarithm of a number to the base you specify.
    ```
    =LOG(W2, 10)
    ```

29. **EXP**: Returns e raised to the power of a given number.
    ```
    =EXP(X2)
    ```

30. **SIN**: Returns the sine of an angle.
    ```
    =SIN(Y2)
    ```

31. **COS**: Returns the cosine of an angle.
    ```
    =COS(Z2)
    ```

32. **TAN**: Returns the tangent of an angle.
    ```
    =TAN(AA2)
    ```

33. **PI**: Returns the value of pi (π).
    ```
    =PI()
    ```

34. **RAND**: Returns a random number between 0 and 1.
    ```
    =RAND()
    ```

35. **RANDBETWEEN**: Returns a random integer between the numbers you specify.
    ```
    =RANDBETWEEN(1, 100)
    ```

36. **AND**: Returns TRUE if all of its arguments are TRUE.
    ```
    =AND(AB2=5, AC2="Yes")
    ```

37. **OR**: Returns TRUE if any argument is TRUE.
    ```
    =OR(AD2="Pass", AD2="Fail")
    ```

38. **NOT**: Reverses the logical value of its argument.
    ```
    =NOT(AE2="Error")
    ```

39. **IFERROR**: Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.
    ```
    =IFERROR(AF2/AH2, "Division by zero")
    ```

40. **SUMIF**: Adds the cells specified by a given criteria.
    ```
    =SUMIF(AI2:AI100, ">10")
    ```

41. **COUNTIF**: Counts the number of cells specified by a given criteria.
    ```
    =COUNTIF(AJ2:AJ100, "<>0")
    ```

42. **AVERAGEIF**: Calculates the average of cells specified by a given criteria.
    ```
    =AVERAGEIF(AK2:AK100, ">=50")
    ```

43. **SUMIFS**: Adds the cells in a range that meet multiple criteria.
    ```
    =SUMIFS(AL2:AL100, AM2:AM100, ">200", AN2:AN100, "<500")
    ```

44. **COUNTIFS**: Counts the number of cells in a range that meet multiple criteria.
    ```
    =COUNTIFS(AO2:AO100, ">=50", AP2:AP100, "<=100")
    ```

45. **AVERAGEIFS**: Calculates the average of cells in a range that meet multiple criteria.
    ```
    =AVERAGEIFS(AQ2:AQ100, AR2:AR100, ">50", AS2:AS100, "<=100")
    ```

46. **SUBTOTAL**: Returns a subtotal in a list or database.
    ```
    =SUBTOTAL(109, AT2:AT100
    '''
Certainly! Here are 54 more Excel functions:

47. **IFNA**: Returns a value you specify if a formula results in the #N/A error; otherwise, returns the result of the formula.
   ```
   =IFNA(AU2/AW2, "N/A")
   ```

48. **INDIRECT**: Returns the reference specified by a text string.
   ```
   =INDIRECT("B2")
   ```

49. **OFFSET**: Returns a reference offset from a starting cell or range of cells.
   ```
   =OFFSET(A1, 3, 2)
   ```

50. **ROW**: Returns the row number of a reference.
   ```
   =ROW(A2)
   ```

51. **COLUMN**: Returns the column number of a reference.
   ```
   =COLUMN(B3)
   ```

52. **INDEX-MATCH (Combined)**: Combining INDEX and MATCH functions for flexible lookups.
   ```
   =INDEX(B2:B10, MATCH(F2, A2:A10, 0))
   ```

53. **TEXT**: Formats a number and converts it to text.
   ```
   =TEXT(C2, "yyyy/mm/dd")
   ```

54. **NOW()**: Returns the current date and time.
   ```
   =NOW()
   ```

55. **TODAY()**: Returns the current date.
   ```
   =TODAY()
   ```

56. **RANDARRAY**: Generates an array of random numbers between 0 and 1.
   ```
   =RANDARRAY(5, 5)
   ```

57. **SEQUENCE**: Generates a sequence of numbers in an array.
   ```
   =SEQUENCE(5, 1, 0, 2)
   ```

58. **TRANSPOSE**: Transposes the rows and columns of a range or array.
   ```
   =TRANSPOSE(A1:D4)
   ```

59. **FILTER**: Filters a range of data based on criteria you define.
   ```
   =FILTER(A2:A100, B2:B100>50)
   ```

60. **SORT**: Sorts the contents of a range or array.
   ```
   =SORT(A2:A100, 1, TRUE)
   ```

61. **TEXTJOIN**: Joins multiple text strings into one string, with a specified delimiter.
   ```
   =TEXTJOIN(", ", TRUE, C2:C10)
   ```

62. **UNIQUE**: Returns a list of unique values from a range or array.
   ```
   =UNIQUE(D2:D100)
   ```

63. **SPLIT**: Splits text into separate cells based on a delimiter.
   ```
   =SPLIT(E2, ",")
   ```

64. **SUBSTITUTE**: Substitutes new text for old text in a text string.
   ```
   =SUBSTITUTE(F2, "old", "new")
   ```

65. **REPT**: Repeats text a given number of times.
   ```
   =REPT(G2, 3)
   ```

66. **SEARCH**: Finds one text string within another (case-sensitive).
   ```
   =SEARCH("find", H2)
   ```

67. **FIND**: Finds one text string within another (case-insensitive).
   ```
   =FIND("find", I2)
   ```

68. **LENB**: Returns the number of bytes used to represent a text string.
   ```
   =LENB(J2)
   ```

69. **REPLACE**: Replaces characters within a text string, starting at the specified position.
   ```
   =REPLACE(K2, 2, 3, "new")
   ```

70. **CONCAT**: Concatenates a range of cells.
   ```
   =CONCAT(L2:L10)
   ```

71. **HYPERLINK**: Creates a clickable hyperlink.
   ```
   =HYPERLINK("http://www.example.com", "Click here")
   ```

72. **DAYS**: Calculates the number of days between two dates.
   ```
   =DAYS(M2, N2)
   ```

73. **NETWORKDAYS**: Calculates the number of working days between two dates.
   ```
   =NETWORKDAYS(O2, P2)
   ```

74. **EDATE**: Returns a date that is a specified number of months before or after another date.
   ```
   =EDATE(Q2, 3)
   ```

75. **EOMONTH**: Returns the last day of the month, n months before or after a given date.
   ```
   =EOMONTH(R2, -2)
   ```

76. **YEAR**: Returns the year of a date.
   ```
   =YEAR(S2)
   ```

77. **MONTH**: Returns the month of a date.
   ```
   =MONTH(T2)
   ```

78. **DAY**: Returns the day of the month.
   ```
   =DAY(U2)
   ```

79. **WEEKDAY**: Returns the day of the week as a number.
   ```
   =WEEKDAY(V2)
   ```

80. **ISNUMBER**: Checks if a value is a number.
   ```
   =ISNUMBER(W2)
   ```

81. **ISTEXT**: Checks if a value is text.
   ```
   =ISTEXT(X2)
   ```

82. **ISBLANK**: Checks if a cell is empty.
   ```
   =ISBLANK(Y2)
   ```

83. **ISERROR**: Checks if a cell contains an error.
   ```
   =ISERROR(Z2)
   ```

84. **ISODD**: Checks if a number is odd.
   ```
   =ISODD(AA2)
   ```

85. **ISEVEN**: Checks if a number is even.
   ```
   =ISEVEN(AB2)
   ```

86. **ISLOGICAL**: Checks if a value is TRUE or FALSE.
   ```
   =ISLOGICAL(AC2)
   ```

87. **ISNONTEXT**: Checks if a value is not text.
   ```
   =ISNONTEXT(AD2)
   ```

88. **ISNUMBER**: Checks if a value is a number.
   ```
   =ISNUMBER(AE2)
   ```

89. **ISREF**: Checks if a value is a reference.
   ```
   =ISREF(AF2)
   ```

90. **ISTEXT**: Checks if a value is text.
   ```
   =ISTEXT(AG2)
   ```

91. **AVERAGEA**: Calculates the average of values, including text and logical values.
   ```
   =AVERAGEA(AH2:AH100)
   ```

92. **COUNTA**: Counts the number of cells that are not empty.
   ```
   =COUNTA(AI2:AI100)
   ```

93. **MAXA**: Returns the largest value in a set of values, including text and logical values.
   ```
   =MAXA(AJ2:AJ100)
   ```

94. **MINA**: Returns the smallest value in a set of values, including text and logical values.
   ```
   =MINA(AK2:AK100)
   ```

95. **PRODUCT**: Multiplies numbers together.
   ```
   =PRODUCT(AL2:AL10)
   ```

Of course! Here are 4 more Excel functions:

97. **MEDIAN**: Returns the median of the given numbers.
   ```
   =MEDIAN(AM2:AM10)
   ```

98. **VAR.P**: Estimates the variance based on a population.
   ```
   =VAR.P(AN2:AN100)
   ```

99. **VAR.S**: Estimates the variance based on a sample.
   ```
   =VAR.S(AO2:AO100)
   ```

100. **STDEV.P**: Estimates the standard deviation based on a population.
   ```
   =STDEV.P(AP2:AP100)
   ```


    
