# stock_analysis


## Overview
### Purpose
The goal for this project is to create a Microsoft Excel VBA code to determine whether the stocks from 2017 and 2018 is worth investing or not by collecting information from the given data. Specifacly for this challenge, the assignment was to refactor the code we created during the module to increase the efficiency.

### Data
The given excel file was consisted from 2 sheets, 2017 and 2018. Both sheets include the ticker name (Ticker), open date (Date), highest and lowest price (High, Low), closing date (Close), adjusted closing price (Adj. Close), and the volume of the stock (Volume). They both have 12 different stocks with 3013 rows and 8 columns.


## Results
### Analysis
Click here to view the Excel file: [VBA Challenge - Stock Analysis]([https://github.com/ninicholasas/stock_analysis/blob/main/VBA_Challenge.xlsm])

The written VBA code is in Module 6. I first started by copying the code for creating the input box, chart headers, ticker array, and the activation for the proper worksheet. Next I have changed how to iterate through the tickers by adding a tickerIndex and a for-loop to loop through the ticker, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
Below is the actual code written in Module 6;

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        ticker = tickers(j)
        tickerVolumes(j) = 0
        tickerStartingPrices(j) = 0
        tickerEndingPrices(j) = 0
    Next j
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            End If
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    'Activate Worksheet
    Worksheets("All Stocks Analysis").Activate
    'Set the header Bold
    Range("A3:C3").Font.FontStyle = "Bold"
    'Draw the bottom line for the header
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Change the total volume format
    Range("B4:B15").NumberFormat = "#,##0"
    'Change the Return number format
    Range("C4:C15").NumberFormat = "0.0%"
    'Change the width of column B
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        'Setting the interior color to green if positive, red if negative
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    
    'Pop up a message box indicatingthe time it needed to run the code
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

From the outcome of the above code we will get the same results from the code before it was refactored;

<img width="242" alt="VBA_Challenge_Result_2017" src="https://user-images.githubusercontent.com/110373282/201551627-3787a248-272f-45b0-bdf7-9cdccf82b3dc.png"><img width="244" alt="VBA_Challenge_Result_2018" src="https://user-images.githubusercontent.com/110373282/201551634-fec408e1-dee1-406c-a046-259eede1eb35.png">

as you can see from the screenshot above, we could tell that **ENPH** and **RUN** would be the ideal stocks to investigate.
By the growth of the return rate I might consider **RUN** more than **ENPH**.

Now comparing the run time between the original code and the refactored code;

***Original***

<img width="251" alt="green_stock_analysis_2018" src="https://user-images.githubusercontent.com/110373282/201551883-b70b6c43-f731-44c2-b956-c60b9ac5b16b.png">

***Refactored***

<img width="255" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110373282/201551895-0c7ddfed-1c28-41e7-85dd-71f26e671483.png">

the refactored code ran faster than 5 times the original code.

## Summary
### Advantages and Disadvantages of Refactoring Code in General
There are many advantages in refactoring a code. Refactoring improves the maintainability of the code i.e. makes the code easier to understand, helps find bugs easier and enhances executing the program. On the other hand if the code is not properly refactored it may end up introducing new bugs and lose the flexibility of the original code.

### Advantages and Disadvantages of Refactoring Code in this Challenge
As shown in the **Analysis**, there was a significant difference in the run time between the original and refactored code.
The clarity of the code seams to be a little lower than the original code, but comparing the run time the outcome is enough to overlook the fact that the code looks more complicated.
