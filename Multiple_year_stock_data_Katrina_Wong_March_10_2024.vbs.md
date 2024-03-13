'Creating macro that runs across all worksheets at once
Sub FormatAllWorksheets()
    Dim ws As Worksheet
    
    'Declare k to increment worksheets
    Dim k As Long
    
    'Declare i to be used for source row and j to be use for destination row
    Dim i As Long
    Dim j As Long
    
    'Declare earliest date and latest date for the year
    'Declare the earliest date price and the latest date price
    Dim EarliestDate As Long
    Dim LatestDate As Long
    Dim EarliestDatePrice As Currency
    Dim LatestDatePrice As Currency
    
    'Declare TotalShares for Total Stock Volume Column
    Dim TotalShares As Double
    
    'Declare cell for color coding
    Dim cell As Range
    
    Dim lastrow As Long
    'lastrow = 5
    'lastrow = 753001
    
    'Initialize k outside the loop to retain its value
    k = 2
    
    'Begin For Loop to go through each worksheet
    For Each ws In ThisWorkbook.Worksheets
 
    'Begin logic for worksheet
  
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Initialize variables
    j = 2
    EarliestDate = 99999999
    LatestDate = 0
    TotalShares = 0
    
    'Begin For Loop for first part of the assignment
    For i = 2 To lastrow
    
    'To calculate Total Stock Volume Column
        TotalShares = ws.Cells(i, 7).Value + TotalShares
   
    'To update Yearly Change, find earliest date and price and latest date and price
    If ws.Cells(i, 2).Value < EarliestDate Then
            EarliestDate = ws.Cells(i, 2).Value
            EarliestDatePrice = ws.Cells(i, 3).Value
   
         End If
         
    If ws.Cells(i, 2).Value > LatestDate Then
            LatestDate = ws.Cells(i, 2).Value
            LatestDatePrice = ws.Cells(i, 6).Value
            
         End If
        
    'To update Ticker column
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
           ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
           
    'To update Yearly Change column
           ws.Cells(j, 10).Value = LatestDatePrice - EarliestDatePrice
           
     'To update the color by positive and negative values
            Set cell = ws.Cells(j, 10) ' Set the cell range
           
            If ws.Cells(j, 10).Value < 0 Then
            ' Apply red fill color to the cell for negative values
            cell.Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(j, 10).Value > 0 Then
            ' Apply green fill color to the cell for positive values
            cell.Interior.Color = RGB(0, 255, 0)
            End If
           
      'To update Percentage Change column
            ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / EarliestDatePrice
            
            'To update the color by positive and negative values
            Set cell2 = ws.Cells(j, 11) ' Set the cell range
           
            If ws.Cells(j, 11).Value < 0 Then
            ' Apply red fill color to the cell for negative values
            cell2.Interior.Color = RGB(255, 0, 0)
            ElseIf ws.Cells(j, 11).Value > 0 Then
            ' Apply green fill color to the cell for positive values
            cell2.Interior.Color = RGB(0, 255, 0)
            End If
            
    'To update Total Stock Volume Column
            ws.Cells(j, 12).Value = TotalShares
            
    'Reset TotalShares for the next ticker symbol
            TotalShares = 0
                
             j = j + 1
             EarliestDate = 99999999
             LatestDate = 0
    
        End If
        
    Next i

    'For the second part of the assignment, using a second for loop
    
    'Declare Greatest % Increase and Decrease
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalStockVolume As Double
    
    'Initialize variables
    GreatestTotalStockVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    
   'Begin Second For Loop
    For i = 2 To lastrow
         
         'Initialize variables
         j = 2
        
         'To update Greatest % Increase
            If ws.Cells(i, 11).Value > GreatestIncrease Then
                 GreatestIncrease = ws.Cells(i, 11).Value
                 ws.Cells(j, 16).Value = ws.Cells(i, 9).Value
                 ws.Cells(j, 17).Value = ws.Cells(i, 11).Value
             
             End If
             
         'To update Greatest % Decrease
            If ws.Cells(i, 11).Value < GreatestDecrease Then
                 GreatestDecrease = ws.Cells(i, 11).Value
                 ws.Cells(j + 1, 16).Value = ws.Cells(i, 9).Value
                 ws.Cells(j + 1, 17).Value = ws.Cells(i, 11).Value
             
             End If
             
             'To update Greatest Total Stock Volume
               If ws.Cells(i, 12).Value > GreatestTotalStockVolume Then
                 GreatestTotalStockVolume = ws.Cells(i, 12).Value
                 ws.Cells(j + 2, 16).Value = ws.Cells(i, 9).Value
                 ws.Cells(j + 2, 17).Value = ws.Cells(i, 12).Value
           
               End If
        
          Next i
          
       ' Declare variables for the next worksheet
        EarliestDate = 99999999
        LatestDate = 0
        TotalShares = 0
        GreatestTotalStockVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
    
        ' Reset variables for the next worksheet
         k = k + 1
    
    Next ws
   
End Sub



