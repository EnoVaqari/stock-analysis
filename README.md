# Stock Analysis

## Overview - Stock Analysis with VBA

### Background
Steve wanted an analysis on stock market. The analysis was performed using Virtual Basic for Application (VBA) through Excel, which created a user friendly analysis for Steve to understand the stock market.

### Purpose
Steve wanted an analysis of bot 2017 and 2018 dataset for his parents. Our goal is to edit or refactor the code, in order to get the results faster. After editing or refactoring the code, we will compare original and refactored scripts to identify the code that runs faster.

## Results

Below is the original script used to analyze the Stock Market dataset.
```
Sub MacroCheck()

Dim testMessage As String
testMessage = "Hello World!"
MsgBox (testMessage)


End Sub


Sub DQAnalysis()
    
    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQ0 (Ticker: DQ)"
   
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    'setting initial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    'Loop over all rows
    For i = 2 To RowCount
    
    'increase totalVolume if ticker is "DQ"
    If Cells(i, 1).Value = "DQ" Then
        'totalVolume increased by value in the current row
        totalVolume = totalVolume + Cells(i, 8).Value
    End If
    
    If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
        startingPrice = Cells(i, 6).Value
        
    End If
    
    If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
     
     endingPrice = Cells(i, 6).Value
     
     End If
     
    
    
Next i
 'MsgBox (totalVolume)
 
 Worksheets("DQ Analysis").Activate
 Cells(4, 1).Value = 2018
 Cells(4, 2).Value = totalVolume
 Cells(4, 3).Value = (endingPrice / startingPrice) - 1

'formatting
Range("c4").NumberFormat = "0.00%"
Range("b4").NumberFormat = "$#,##0"

 

 
           
End Sub




Sub AllStockAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

Worksheets("All Stocks Analysis").Activate

yearValue = InputBox("what year would you like to run the analysis on?")
startTime = Timer
'title
Range("a1").Value = "All Stocks (" + yearValue + ")"

'Headers
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'Assigned tickers
Dim tickers(11) As String

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

'Initialize variables for starting and ending price
Dim startingPrice As Single
Dim endingPrice As Single

'Data worksheet activated
Worksheets(yearValue).Activate


'numbers of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'loop through tickers
For i = 0 To 11
ticker = tickers(i)
totalVolume = 0


    'loop through rows in the data
Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
    'total volume for current ticker
    If Cells(j, 1).Value = ticker Then
    totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
    'starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    startingPrice = Cells(j, 6).Value
    End If
    
    'ending price for current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    endingPrice = Cells(j, 6).Value
    
    End If
Next j
     
'output data for all stock analysis worksheet
Worksheets("All Stocks Analysis").Activate

'output data for current ticker
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i

endTime = Timer
MsgBox "Code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub formatAllStocksAnalysisTable()

'Formatting
Worksheets("All Stocks Analysis").Activate

'Visual Style Formatting
Range("A3:C3").Font.FontStyle = "Bold Italic"
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "$#,##0.00"
Range("C4:C15").NumberFormat = "0.00%"
Columns("B").AutoFit

'Color Formatting

'Setting up the loop
dataRowStart = 4
dataRowEnd = 15
For i = dataRowStart To dataRowEnd
 
 'Conditions
 
 'Green color for positive outcomes
  If Cells(i, 3) > 0 Then
  Cells(i, 3).Interior.Color = vbGreen
  
  'red color for negative outcomes
  ElseIf Cells(i, 3) < 0 Then
  Cells(i, 3).Interior.Color = vbRed
  
  Else
  'clear cell color if we have 0 as outcome
  Cells(i, 3).Interior.Color = xlNone
  
  End If
  
  Next i
  
End Sub

Sub ClearWorksheet()
Worksheets("All Stocks Analysis").Activate

Cells.Clear

End Sub

Sub ClearDQWorksheet()

'Activate All Stocks Analysis Worksheet
    Worksheets("DQ Analysis").Activate

    Cells.Clear

End Sub

```
Below is the refactored VBA script used to analyze the Stock Market dataset.

```
Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Title Analysis
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
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
    tickerVolumes(tickerIndex) = 0
    

    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End If
            End If
            
            
        '3c) check if the current row is the last row with the selected ticker
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End if
            End If
            
        '3d) Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker, total daily volume and return
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

    
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

### Stock Analysis for 2017 and 2018

For 2017, all stocks had a positive return on investment rate with the exeption of TERP, which had -7.21% ROI rate.
![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/All_Stocks_2017.png)

For 2018, ENPH and RUN had positive outcomes with 81.92% and 83.95% return on investments respectively. The rest of the outcomes were negative. 

![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/All_Stocks_2018.png)

### Original code Vs Refactored Code 

By refactoring the code, there was a visible improvement on the time the code ran from 0.9375 seconds to 0.1875 seconds for 2017, and from 0.8710938 seconds to 0.1679688 seconds for 2018. 


![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/original_2017.png)
![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/refactored_2017.png)
![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/original_2018.png)
![](https://github.com/EnoVaqari/stock-analysis/blob/main/Resources/Refactored_2018.png)

## Summary

### Advantages of refactoring
Based on our analysis, refactoring  using VBA script reduced the time the code ran significantly, which is a major advantage when using large datasets.

### Disadvantages of refactoring 

Refactoring the code exposes the analysis at risk of losing or destroying parts of the original code, which can result in errors and waste of time. However, this disadvantage can be prevented by being careful, saving the original work, and comparing results to ensure accuracy.


