# Stock-Analysis
## Overview of Project

The Purpose of this project was to refactor VBA Code to sort by a group of stock tickers and pull their yearly results for 2017 and 2018. We are trying to determine if these stocks are worth investing in by understanding their Year over Year . We have done this for module 2, but this challenge is specifically going to increase the efficency of the original code. 

## Details
### Data

Our spreadsheet VBA_Challenge2018.xmls originally contained two sheets Each sheet is seperated by Year (2017, 2018)  In each sheet is 6 columns with the column headers reading left to right. Ticker, Date, Open, High, Low, Adj Close, and Volume. There are 3012 total rows.

### Goal

The goal of this challenge is to retrieve our 12 tickers, including their daily volume and return on each stock. 
To run the Marco to compare a stock performance of 2017 vs 2018. As well as share the execution time of the original script vs the refractoring code. Identify pro's and con's of refractoring code. 


## Results

Before refractoring the code, I needed to downloaded and installed VBA_Challenge.vbs, I renamed my greenstocks.xlms to VBA_Challenge.xmls and uploaded the VBA_Challenge.vbs into the macro. This gave me a initial template to follow along the challenge. Here is a copy and paste from VBS

'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickervolumes(12) As Long
    Dim tickerstartingprices(12) As Single
    Dim tickerendingprices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
         tickervolumes(i) = 0
         tickerstartingprices(i) = 0
         tickerendingprices(i) = 0
  
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerstartingprices(tickerIndex) = Cells(i, 6).Value
            
        'End If
         End If
         
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerstartingprices(tickerIndex) = Cells(i, 6).Value
        End If
        
            

            '3d Increase the tickerIndex.
                 If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickervolumes(i)
        Cells(4 + i, 3).Value = tickerendingprices(i) / tickerstartingprices(i) - 1
        
    Next i
    

## Summary
### Pros

Refractoring your code makes it more clean and organized for 
Another engineer to be able to quickly review and easily understand your work, therefore making it less complicated to design, debug, and more simple to add new improvements
Refractoring your code can also reduce run times

### Cons

We cannot refractor the code due to the size of the file. Too Large, Excel has a limit of rows
No proper test case exist

### Advantage of Refractoring Code

The biggest advantage was the decrease macro run time. Originally the analysis in Greenstocks.xlms took 1 second to run. However in the screenshots below, by refractoring our code. we reduced runtimes to less than 20% of a second. Refractoring our code was really beneficial because in the real world, these run times can be hours. Imagine taking 80% off a hour wait. Thats useful. 

### 2018 vs 2017 Run time Screenshots
https://github.com/Jaguo415/Stock-Analysis/blob/main/Screen%20Shot%202021-06-20%20at%2012.08.10%20PM.png
https://github.com/Jaguo415/Stock-Analysis/blob/main/Screen%20Shot%202021-06-20%20at%2012.15.33%20PM.png

