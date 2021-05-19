# VBA-chalenge
The process of the project:

The first thing that I did was mapping the project in Excel. It helped me to see the big picture of the project and started thinking which information I was looking for. For example, how to calculate the total volume traded for the tickers and to ensure I had the correct opening and closing prices. Having got a better understanding of the data I was ready to begin figuring out the code to do the report.
I decided to concentrate on getting the report working on one sheet (A) in the testing workbook as follows:
1.	Initially I worked on getting the key data correct by scanning the rows to find ticker codes, opening, and closing pricing and to get the sum of the volume traded.
2.	Once I had this working for one ticker, I figured out how to process all the tickers in the sheet using a Do While Loop. I then formatted the output into specific columns. 
3.	I then worked on accumulating the summary data for highest increase, lowest decrease and maximum volume traded. 
4.	Finally, I formatted the colours and numbers correctly.
Once I had this working for one sheet it was simply a matter of applying a For Each Loop to process each sheet in the workbook. Having got this working with the test data, I then ran against the multi-year data.
In Visual Studio I used the debug features (watch, run, add breakpoints) and stepped through line by line to check the execution and fix bugs in the code as crashes and errors occurred. I had several issues with looping including incorrect condition checks which kept the loop going forever and incorrect sums when new tickers were found. 
