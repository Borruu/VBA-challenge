# Stock market data analysis with VBA

### Objective
Create a script that loops through all the stocks for one year and outputs the following information:
- The ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.
- Use conditional formatting that will highlight positive change in green and negative change in red.
![image](https://github.com/Borruu/VBA-challenge/assets/112932520/696d850a-1c73-4d46-abb4-1aee19626fb4)
- **Bonus challenge:** add functionality to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
![image](https://github.com/Borruu/VBA-challenge/assets/112932520/3ce647ca-29db-4a09-9b63-2ffcdf7981d0)

### Solution
Subroutine `ticker_03` in file titled `Mod02_VBA_Code` loops through all the stocks for one year to address the components below.

1. Columns I to L are populated with the values for ticker, yearly change, percent change and total stock volume respectively. 
2. Column J (yearly change) has conditional formatting applied to reflect positive or negative change through interior colour change. Cells with no change will have no interior colour.
3. Number formatting is applied to columns K and L. Headings are added to columns I to L and all columns then have width adjusted.
4. Values for the bonus challenge are found and stored in variables.
5. Bonus challenge values are populated in columns P and Q.
6. Number formatting is applied to relevant cells in column Q.
7. Headings are added to columns O, P and Q.
8. Formatting (bold text and autofit width) are applied to the results.
9. Steps 1 to 8 are then repeated on all worksheets.

### References
Dataset created by Trilogy Education Services, a 2U, Inc. brand.

