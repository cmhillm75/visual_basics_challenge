Dr Bogue in referencing formatting numbers in cells, stated to google in office hours on 10/7.

In formatting for the color indexes, utilized the Sub_Checkerboard as reference.

For determining how to calculate to LastRow & columns, creating titles that remain in spreadsheets 
& how to write a for loop function that will iterate accross multiple worksheets utilzied the 
inclass work for census_data_2016-2019 part 1.

On how to save a vbs script completed a google search and after several failed attempts found website that showed how it works outside of Excel.
https://www.geeksforgeeks.org/how-to-make-save-and-run-a-simple-vbscript-program/

For summary table utilized the stock change as a variable menu for Dim items and again on 10/7 office hours Dr Bogue referenced corrections on 
how to obtain last row in a column and need to list ws.Cells.

In determining how to get the percentage change and quarterly change in the stock_change sub and the summary table sub for Greatest increase/decrease % I debugged/utilzed 
https://support.microsoft.com/en-us/office/formatnumber-function-91030eab-2887-43d4-9c17-311ab6ebf43b & Microsoft AI Copilot
to change from a standard +/- integer and to be formatted in the 0.00 & 00.00% formats. 

Again in summary table I was stuck on how to set a value for greatest increases and decreases. per Xpert Learning assistant https://bootcampspot.instructure.com/courses/6344/external_tools/313
confirmed to set As Double and it provided a value of 0 and -1. Since there would be no limit to how high a number could go on positive side, I set the floor as negative and used -1000

In determining how to get the loop thru all worksheets to manage going thru each ticker and how to handle that, determined that I needed to create a dictionary to hold 
the processed tickers as object. https://www.mrexcel.com/board/threads/vba-creating-objects-in-a-loop.770533/

In determining what to call the worksheets I needed to loop thru i referenced https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
along with class materials from 

Variable names/items to Dim for and how to label to process the loop code, stack over flow https://stackoverflow.com/questions/64394196/for-loops-in-vba-initializing-variable 
which showed me how to take my list of titles from row 1 and what to set initial values as ie. ticker = "" since we were using the ws.Cells(i, 1).Value for the ticker, volume, % change, total change to go thru all the tickers 
as counting each ticker as an individual variable would be excessive work. 