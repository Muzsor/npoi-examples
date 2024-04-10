# Copy sample row and recalculate workbook

This is a simple demo that use NPOI to copy a sample row and evaluate all formula in workbook.
## Test data
The original sheet has a header and one row of data. The Item Price and Amount is use ```ROUNDUP(RAND(),00)``` to generate random number. Total is ```B2*C2``` 
![original.png](images%2Foriginal.png)

The result will have 99 rows of the data that clone from 2nd row but make Total relative to the row.
![result.png](images%2Fresult.png)