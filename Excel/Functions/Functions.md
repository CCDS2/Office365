

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
