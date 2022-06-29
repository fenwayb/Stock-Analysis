<h1 Align="Center">
  
  Stock Analysis

  # Overview
  
  ### Purpose
  
  <p></p>

  ### Background
  
  <p></p>
  
  # Analysis
  <p>![VBA_Old_2017](https://user-images.githubusercontent.com/106105597/176467199-5f358133-18f2-4f2a-a7cf-666e65656292.png)

  ![VBA_Challenge_2017](https://user-images.githubusercontent.com/106105597/176467212-a41b4aad-c92f-49f0-8f67-f8bae440de4a.png)

  ![VBA_Old_2018](https://user-images.githubusercontent.com/106105597/176467292-95dec2da-c85a-4730-8be3-fb73545dd10e.png)
    
  ![VBA_Challenge_2018](https://user-images.githubusercontent.com/106105597/176467307-53a04995-d3a8-49e2-a41f-44e830ae5495.png)

  </p>
    
  # Summary
  
  <p></p>
  
  ### Why Refactor?
   
  <p></p> 
  
  ### Why did this refactor work?
  
  '''
  '1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         '3d Increase the tickerIndex.
         
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
  '''
  
  <p></p>
