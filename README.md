<h1 Align="Center">
  
  Stock Analysis

  # Overview
  
  ### Purpose and Background
  
  <p>In this project we used Visual Basic for Applications, or VBA to analyze stocks of green energy companies for our friend Steve who wants to help his parents find the best stocks to invest in. We made the program easy for someone not experienced with VBA to run with a button that will run our code. The analysis here will be about refactoring that code to see if we can make it run faster</p>
  
  # Analysis
  <p>
  
  Based on the following screenshots our code refactor was a massive success. It decreased the runtime of the code for both the 2017 and 2018 datasets by nearly ten-fold. While these data sets are small enough that the difference in speed does not make a huge difference as the datasets get larger it becomes drastically more important for it to run fast
    
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
   
  <p>The reason to refactor code is to potentially find a way to make the code more optimized - either to make it run faster or potentially use less memory. This is common because there are many different ways you can accomplish the same goal in code and so there often might be a faster way than the first way you tried</p> 
  
  ### Why did this refactor work?
  The key to this refactor making the code run faster is that we un-nested the loops from the original code. In the first code block, we were using three variables, and at the beginning of every loop through the spreadsheet we were reinitializing those variables back to 0 so they can be used again. In the refactored code however, we turned those 3 variables in to 36 variables, 1 copy of each variable for each ticker, and this allows us to run through the spreadsheet all at once without having to overwrite any of the variables back to 0. After doing a bit of research on the topic it seems that one of the operations that takes the most amount of time in VBA is writing to a spreadsheet. So even though we are using signficantly more variables in the refactored code, it will run much faster because it only needs to write our data to the spreadsheet once at the very end instead of having to do it at the end of each loop
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
