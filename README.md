# Analysis of Green Energy Stocks

## Overview of Project

Steve, a recent finance graduate wanted to perform analysis on green energy stocks in order to help his parents diversify their investment portfolio as initially they were only interested to invest in green energy stocks. Steve's parents so far only invested in DAQO New Energy Corp. For this analysis, I created a Visual Basic Application tool for Steve where he can easily compare the past performance on each of the stocks (ticker symbols) by extracting data on total daily volume and percentage of return. A timer function was also set up to give Steve a better understanding on how long exactly will it take for the system to give results.

### Purpose

The purpose of this analysis in ```VBA_Challenge.xlsm``` is to make an efficient way to track the performance of different stocks using VBA. After being able to run the analysis of 12 different stocks, I could find a more efficient way to work with the given data. In order to make the analysis more efficient, I refactored my code. This project looks to see if my refactoring made the analysis more efficient.

## Results

### 2017 Stocks Analysis ###

The results of the VBA run are shown below. Based on the table we can see that almost all of the stocks were in green for 2017 with only "TERP" showing a net loss of -7.21%. The top-performing stocks for the year were "DQ", "SEDG", "ENPH" and "FSLR" respectively, with DQ almost doubling in price by the end of the year. It seems that this year was generally good one for sustainable energy stocks and anybody invested in the stocks listed below except for "TERP" would have seen a positive return on investment. Initially when we wrote our VBA script, it had a total run time of 1.015625 seconds for analysis on all Stocks in 2017. After we refactored the code the run time was reduced to 0.203125.


![ Allstocks2017.png]( https://github.com/smj452/stocks-analysis/blob/main/Resources/Allstocks2017.png)


**Run Time Initial**

![2017_originalcode]( https://github.com/smj452/stocks-analysis/blob/main/Resources/2017_originalcode.png)


**Run Time Refactored**

![VBA_Challenge_2017.png]( https://github.com/smj452/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)


### 2018 Stocks Analysis
Our analysis in 2018 shows that most of the green stock’s performance was low. Out of the 12 stocks we analyzed we could only find two stocks ENPH and RUN that had a positive net return in 2018. DQ’s performance was low in 2018 with a negative return of -62.60%. Our initial code took 1.054688 seconds to run. After the code was refactored and committed, it only took 0.234375 seconds to execute and give us the results making our analysis faster.


![ Allstocks2018.png]( https://github.com/smj452/stocks-analysis/blob/main/Resources/Allstocks2018.png)


**Run Time Initial**

![2018_originalcode]( https://github.com/smj452/stocks-analysis/blob/main/Resources/2018_originalcode.png)

**Run Time Refactored**

![VBA_Challenge_2018.png]( https://github.com/smj452/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)


**Refactored Code**
```
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
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single, tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" & yearValue & ")"
   'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    For i = 0 To 11
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
## Summary

### Code Refactoring ###

Code Refactoring is a way of restructuring and optimizing already written code to improve its flow, readability, and execution.
### Advantages of Refactoring 
- Refactoring the code helps in optimizing the performance of the code and troubleshoot for errors.
- The code becomes easier to understand and follow for others involved in the project.

### Disadvantages of Refactoring ###
- Refactoring can lead to additional bugs if the coder is not familiar with the code in the first place.
- Commenting and editing the code takes additional time and effort.

### Original and Refactored VBA Script
One of the advantages of refactoring code in VBA script is that you can use the original code and make editions side by side with your old code using different modules. The major disadvantage of refactoring code in VBA script is that if you do not have a strong understanding of the existing code, you will struggle to refactor your code.
	




