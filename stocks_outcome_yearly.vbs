VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-------------------------------------
'VBA Homework - The VBA of Wall Street
'-------------------------------------

Sub stocks_outcome_yearly()

Dim wsheet As Worksheet

    ' --------------------------------------------
    ' loop through all sheets
    ' --------------------------------------------
    For Each wsheet In Worksheets
    
          '------------------------------------------------------------------------------------------------
          ' Sort the data by the columns "ticker" and "date" in ascending order in case data was disturbed
          '------------------------------------------------------------------------------------------------
         
          ' Get the final row
          finalRow = wsheet.Cells(Rows.Count, 1).End(xlUp).Row
          
          ' Sort by ticker and date
          With ActiveSheet.Sort
             .SortFields.Add Key:=Range("A1:A" & finalRow), Order:=xlAscending
             .SortFields.Add Key:=Range("B1:B" & finalRow), Order:=xlAscending
             .SetRange Range("A1:G" & finalRow)
             .Header = xlYes
             .Apply
          End With
    
          ' --------------------------------------------
          ' INSERT TITLES
          ' --------------------------------------------
             
          ' Insert title for totals per ticker
          wsheet.Range("I1").Value = "Ticker"
          wsheet.Range("J1").Value = "Yearly Change"
          wsheet.Range("K1").Value = "Percent Change"
          wsheet.Range("L1").Value = "Total Stock Volume"
          
          ' Formatting title
          wsheet.Range("I1:L1").Interior.ColorIndex = 34
          wsheet.Range("I1:L1").Borders.ColorIndex = 1
          wsheet.Range("I1:L1").Font.Bold = True
            
          ' Insert title for greatest values
          wsheet.Range("O1").Value = "Ticker"
          wsheet.Range("P1").Value = "Value"
          wsheet.Range("N2").Value = "Greatest % Increase"
          wsheet.Range("N3").Value = "Greatest % Decrease"
          wsheet.Range("N4").Value = "Greatest Total Volume"
           
          ' Formatting title
          wsheet.Range("O1:P1").Interior.ColorIndex = 34
          wsheet.Range("N4:P4").Borders.ColorIndex = 1
          wsheet.Range("O1:P1").Font.Bold = True
          wsheet.Range("N2:N4").Font.Bold = True
          
          
          ' ---------------------------------------------------------------------------------------------------------------
          ' DEFINITION OF VARIABLES - PART 1 - Summary Table (Ticker, Yearly Change, Percent Change and Total Stock Volume)
          ' ---------------------------------------------------------------------------------------------------------------
          
          ' Created a variable to hold the beggining price of the stock of a given year
          Dim beginningPrice As Double
          
          ' Create a variable to hold the closing price of the stock at the end of the year
          Dim closingPrice As Double
          
          ' Created a variable to hold the change in price per stock at the end of the year (closingPrice - beginningPrice)
          Dim yearlyChange As Double
          
          ' Create a variable to hold percent change from opening price at the beginning of a given year to the closing price at the end of that year
          Dim percentChange As Double
          
          ' Create a variable to hold the total Stock Volume
          Dim totalStockVolume As Double
          
          ' Create a variable to hold the number of rows of the summary table
          Dim vRow As Long
          
          ' Create a variable to hold the total of row per each ticker
          Dim totalRowTickers As Long
                      
                      
          ' ---------------------------------------------------------------------------------------
          ' DEFINITION OF VARIABLES PART 2 - STOCK WITH GREATEST VALUES
          ' ---------------------------------------------------------------------------------------
                      
          ' Create variable to hold the Ticker of the stock with the greatest % Decrease
          Dim greatestDecreaseTicker As String
          
          ' Create variable for the Value of the stock with the greatest % Decrease
          Dim greatestPercentDec As Double
          
          ' Create variable to hold the Ticker of the stock with the greatest % Increase
          Dim greatestIncreaseTicker As String

          ' Create variable for stock with greatest % Increase
          Dim greatestPercentIn As Double
          
          ' Create variable to hold the Ticker of the Greatest Total Volume
          Dim greatestVolumeTicker As String
            
          ' Create variable to hold the Greatest Total Volume
          Dim greatestTotalVolume As Double
    
                   
          ' Initialized variables
          percentChange = 0
          yearlyChange = 0
          vRow = 2
          totalStockVolume = 0
          totalRowTickers = 0


          ' --------------------------------------------------------------------------------------
          ' CALCULATE DATA - PART 1 (Ticker, Yearly Change, Percent Change and Total Stock Volume)
          ' --------------------------------------------------------------------------------------
          
          ' Loop through all stocks of the current worksheet
          For i = 2 To finalRow
                          
              ' Check if the name of tickers are the same, in case are diferent....
              If wsheet.Cells(i + 1, 1).Value <> wsheet.Cells(i, 1).Value Then
                 
                 ' Assign the name of the current ticker to the Summary Table
                 wsheet.Range("I" & vRow).Value = wsheet.Cells(i, 1).Value
              
                 ' Get the Beginning Price of the current ticker
                 beginningPrice = wsheet.Cells(i - totalRowTickers, 3).Value
                 
                 ' Get the Closing Price of the current ticker
                 closingPrice = wsheet.Cells(i, 6).Value
                 
                 ' Calculate the Yearly Change
                 yearlyChange = closingPrice - beginningPrice
                 
                 ' Assign the yearly change of the current ticker to the Summary Table
                 wsheet.Range("J" & vRow).Value = yearlyChange
                 
                 ' Calcultate the Percent Change
                 If beginningPrice = 0 Then
                    percentChange = 0
                 Else
                    percentChange = yearlyChange / beginningPrice
                 End If
                 
                 ' Assign the percent change to the Summary Table
                 wsheet.Range("K" & vRow).Value = Format(percentChange, "Percent")
                 
                 ' Add to the Total Stock Volume
                 totalStockVolume = totalStockVolume + wsheet.Cells(i, 7).Value
                 
                 ' Assign the Total Stock Volume to the Summary Table
                 wsheet.Range("L" & vRow).Value = totalStockVolume
                 
                 ' Highlight positive changes in green and negative changes in red on the values of yearlyChange on the Summary Table
                 If wsheet.Range("J" & vRow).Value < 0 Then
                    wsheet.Range("J" & vRow).Interior.ColorIndex = 3
                 ElseIf wsheet.Range("J" & vRow).Value > 0 Then
                    wsheet.Range("J" & vRow).Interior.ColorIndex = 4
                 End If
                 
                 ' Add 1 row to the Summary Table
                 vRow = vRow + 1
                 
                 ' Reset the values of the totalStockVolume and totalRowTickers for the next ticker
                 totalStockVolume = 0
                 totalRowTickers = 0
               
              ' If the next cell value the same ticker...
              Else
              
                 ' Add to the tickers total
                 totalRowTickers = totalRowTickers + 1
                  
                 ' Add to the Total Stock Volume
                 totalStockVolume = totalStockVolume + wsheet.Cells(i, 7).Value
                 
                 ' Highlight positive changes in green and negative changes in red on the values of yearlyChange on the summary table
                 If wsheet.Range("J" & vRow).Value < 0 Then
                    wsheet.Range("J" & vRow).Interior.ColorIndex = 3
                 ElseIf wsheet.Range("J" & vRow).Value > 0 Then
                    wsheet.Range("J" & vRow).Interior.ColorIndex = 4
                 End If
                  
              End If
              
              ' ------------------------------------------------------
              ' SET GREATEST DATA - PART 2
              ' ------------------------------------------------------
                
              ' Set the greatest pencent increase
              If greatestPercentIn < percentChange Then
                 greatestPercentIn = percentChange
                 greatestIncreaseTicker = wsheet.Cells(i, 1).Value
              End If
              
              ' Set the greatest pencent decrease
              If greatestPercentDec > percentChange Then
                 greatestPercentDec = percentChange
                 greatestDecreaseTicker = wsheet.Cells(i, 1).Value
              End If
                      
              ' Set the greatest total stock volume
              If greatestTotalVolume < totalStockVolume Then
                 greatestTotalVolume = totalStockVolume
                 greatestVolumeTicker = wsheet.Cells(i, 1).Value
              End If
                         
         Next i
                    
          ' ------------------------------------------------------
          ' ASSIGN DATA - PART 2 - STOCK WITH GREATEST VALUES
          ' ------------------------------------------------------
          
          ' Assign the name of the stock with the greatest percentage Increase
          wsheet.Range("O2").Value = greatestIncreaseTicker
          
          ' Assign the value of the stock with the greatest pencentage Increase
          wsheet.Range("P2").Value = Format(greatestPercentIn, "Percent")
          
          ' Assign the name of the stock with the greatest percentage Decrease
          wsheet.Range("O3").Value = greatestDecreaseTicker
          
          ' Assign the value of the stock with the greatest percentage Decrease
          wsheet.Range("P3").Value = Format(greatestPercentDec, "Percent")
          
          ' Assign the Ticket of the stock with the Greatest Total Volume
          wsheet.Range("O4").Value = greatestVolumeTicker
          
          ' Assign the value of the stock with the Greatest Total Volume
          wsheet.Range("P4").Value = greatestTotalVolume
          
          ' Reset values of the variables for Greatest Values for the next Worksheet
          greatestIncreaseTicker = ""
          greatestDecreaseTicker = ""
          greatestVolumeTicker = ""
          greatestPercentIn = 0
          greatestPercentDec = 0
          greatestTotalVolume = 0
      
    Next wsheet
    
End Sub
