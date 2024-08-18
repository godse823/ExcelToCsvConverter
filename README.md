# ExcelToCsvConverter

Excel file include XLSX and XLS both. XLSX newer format i.e. post 2003 and XLS is older i.e. prev to 2003.
xls can have max 65536 rows per sheet while xlsx can have around ~1 million.
Naive way to process both files can be using workbook.

Disadvantage of using Workbook:
1. Workbook loads whole file in memory which leads to out of memory
2. Workbooks also fails to load files having larger size(even larger than 10 MB)


Resolution
-The parser used are works in streaming fashion, hence whole file will not get loaded in memory. Hence, its a memory efficient
-Can parse files more than 50 MB.
