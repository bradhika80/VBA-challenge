
'***********************************************
'
' Stock Market Analysis - The VBA of Wall Street
'
' An application to perform stock market analysis with real time stock data
'
' Student Name : Radhika Balasubramaniam
'
'*********************************************

Sub StockMarketAnalysis()

    'Declare variables
    
    Dim lastRowCount As Long
    Dim VolumeAnalysisDisplayRowCount As Long
    
    'variables for analytical calculations
    
     Dim ticker As String
     Dim openingValue As Double
     Dim closingValue As Double
     Dim yearlyChange As Double
     Dim percentageChange As Double
     Dim totalStockVolume As Single
    
     'Bonus calculations variables
     
     Dim greatestIncreaseTicker As String
     Dim greatestIncrease As Double
     Dim greatestDecreaseTicker As String
     Dim greatestDecrease As Double
     Dim greatestStockTicker As String
     Dim greatestStockValue As Single
     
     
    '*** loop through the sheets ***
    
     For Each ws In Worksheets
    
        '**** Format the header row for the Volume Analysis Display Results
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
          
     
        
         'get the last row number from the worksheet
         lastRowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         'pre initialize variables before the data processing
         ticker = ws.Cells(2, 1).Value
         openingValue = CDbl(ws.Cells(2, 3).Value)
         closingValue = 0
         yearlyChange = 0
         percentageChange = 0
         totalStockVolume = 0
         VolumeAnalysisDisplayRowCount = 2
         
         'Bonus variables initialize
         greatestIncreaseTicker = Empty
         greatestIncrease = 0
         
         greatestDecreaseTicker = Empty
         greatestDecrease = 0
         
         greatestStockTicker = Empty
         greatestStockValue = 0
         
         ' ***** logic for processing the records in the worksheet begins here *****
         'iterate through the rows in worksheet
         
         For i = 2 To lastRowCount
            
            ' check if the ticker value is different from the current cell.
            ' If different perform calculations, print values in sheet, reset the variable, else add the total stock and assign the closing value.
            
            If ticker <> ws.Cells(i, 1).Value Then
                
                'calculate the yearly change and percentage change
                yearlyChange = closingValue - openingValue
                
                'check if openingvalue is zero, if zero then percentage change is 100%
                If openingValue = 0 Then
                    percentageChange = 1
                Else
                    percentageChange = yearlyChange / openingValue
                End If
                
                'print the value in the cells
                ws.Range("I" & VolumeAnalysisDisplayRowCount).Value = ticker
                ws.Range("J" & VolumeAnalysisDisplayRowCount).Value = yearlyChange
                
                ' color the cell based on value
                If yearlyChange < 0 Then
                   ws.Range("J" & VolumeAnalysisDisplayRowCount).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & VolumeAnalysisDisplayRowCount).Interior.ColorIndex = 4
                End If
                
                ws.Range("K" & VolumeAnalysisDisplayRowCount).Value = Format(percentageChange, "Percent")
                ws.Range("L" & VolumeAnalysisDisplayRowCount).Value = totalStockVolume
                
                'add 1 to move to next row
                VolumeAnalysisDisplayRowCount = VolumeAnalysisDisplayRowCount + 1
                
                'bonus calculations
                If greatestIncrease < percentageChange Then
                    greatestIncreaseTicker = ticker
                    greatestIncrease = percentageChange
                End If
                
                If greatestDecrease > percentageChange Or greatestDecreaseTicker = Empty Then
                    greatestDecreaseTicker = ticker
                    greatestDecrease = percentageChange
                End If
                
                If greatestStockValue < totalStockVolume Then
                    greatestStockTicker = ticker
                    greatestStockValue = totalStockVolume
                End If
                
                
                'Reset variables
                ticker = ws.Cells(i, 1).Value
                openingValue = CDbl(ws.Cells(i, 3).Value)
                closingValue = 0
                yearlyChange = 0
                percentageChange = 0
                totalStockVolume = 0
            
            Else
                'store the closing value
                closingValue = CDbl(ws.Cells(i, 6).Value)
                'calculate total stock volume
                totalStockVolume = totalStockVolume + CLng(ws.Cells(i, 7).Value)
                
            End If
         
         Next i 'end - loop for record processing
         
         '****print bonus calculations
         
         ws.Range("O1").Value = Empty
         ws.Range("P1").Value = "Ticker"
         ws.Range("Q1").Value = "Value"
         
         ws.Range("O2").Value = "Greatest % increase"
         ws.Range("P2").Value = greatestIncreaseTicker
         ws.Range("Q2").Value = Format(greatestIncrease, "Percent")
         
         ws.Range("O3").Value = "Greatest % decrease"
         ws.Range("P3").Value = greatestDecreaseTicker
         ws.Range("Q3").Value = Format(greatestDecrease, "Percent")
    
    
         ws.Range("O4").Value = "Greatest total volume"
         ws.Range("P4").Value = greatestStockTicker
         ws.Range("Q4").Value = greatestStockValue
    
    
    Next 'end - loop for iterate worksheet

 MsgBox ("Completed")


End Sub


