<h1 Align="Center">
  
  Stock Analysis

  # Overview
  
  ### Purpose
  
  <p></p>

  ### Background
  
  <p></p>
  
  # Analysis
  <p>
  
  text before
    
  ![VBA_Old_2017](https://user-images.githubusercontent.com/106105597/176469561-78469bd4-d85c-406b-a159-c8fe0e27376b.png)
   #### original code 2017
    
  ![VBA_Challenge_2017](https://user-images.githubusercontent.com/106105597/176467212-a41b4aad-c92f-49f0-8f67-f8bae440de4a.png)
  #### refactored code 2017
    

  ![VBA_Old_2018](https://user-images.githubusercontent.com/106105597/176479678-0bce1b87-21a0-4547-9818-8c579d009da4.png)
  #### original code 2018
    
  ![VBA_Challenge_2018](https://user-images.githubusercontent.com/106105597/176479785-251f7c76-d2fe-4a31-a561-76d3b6eb8de4.png)
  #### refactored code 2018
    
  </p>
    
  # Summary
  
  <p></p>
  
  ### Why Refactor?
   
  <p></p> 
  
  ### Why did this refactor work?
  ```
  For i = 0 To 11
    
        ticker = tickers(i)
        
        totalVolume = 0
        
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount
            
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
            
            End If
            
        Next j

        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        
        Cells(4 + i, 2).Value = totalVolume
        
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
  ```
  
  ```diff
  +1a) Create a ticker Index
    
    tickerIndex = 0
    
  +1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
+2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
  +2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
 +3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
 +3b) Check if the current row is the first row with the selected tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
        End If
        
 +3c) check if the current row is the last row with the selected ticker
 +If the next row’s ticker doesn’t match, increase the tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
  +3d) Increase the tickerIndex.
         
         tickerIndex = tickerIndex + 1
         
        End If
    
     Next i
           
  +4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
   ```
  I wanted to make the comments clearer so to make this code run you will need to replace the + before each comment with a '
  
  <p></p>
