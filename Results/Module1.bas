Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

 Dim Rng As Range
    Dim RowCount, FRowCount As Integer
    Dim Summary_Table_Row As Integer
    Dim yrStart, yrEnd, yrChange As Double
    Dim vol, bigVol As String
    Dim opening As Boolean
    Dim bigInc, bigDec As Double
    
For Each ws In Worksheets
    
    'Counts number of rows in column A
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Set Rng = ws.Range("I1")
    
    'using advance filter to copy unique values to a new column
    ws.Range("A1:A" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True
    
    'Formatting Ticker Column
    ws.Cells(1, 9).Value = "Ticker"
    ws.Range("I1").ColumnWidth = 9
    
    'Formatting Yearly Change Column
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Range("J1").ColumnWidth = 12
    
    'Formatting Percent Change Column
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Range("K1").ColumnWidth = 15
    ws.Range("K2:K" & RowCount).NumberFormat = "0.00%"
    
    'Formatting Total Stock Volume Column
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Range("L1").ColumnWidth = 17
    
    'counts how many filtered rows there are
    FRowCount = ws.Cells(Rows.Count, "K").End(xlUp).Row
    'formats other summary tables cells to be percents
    
    'formats our value column header
    ws.Cells(1, 16).Value = "Value"
    
    'formats the value of biggest increase/decrease cells as percents
    'how to use range?
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
    'formats 2nd summary table Ticker header
    ws.Cells(1, 15).Value = "Ticker"
    
    'formats labels and cell width for the labels of our 2nd summary table
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Range("N1").ColumnWidth = 19
    
    'initializes our starting row,volume and opening
      Summary_Table_Row = 2
      vol = 0
      opening = False
      
        For i = 2 To RowCount
            'checks if it is the first opening price
            
            If opening = False Then
                yrStart = ws.Cells(i, 3).Value
                
                'will not go through this loop again until we have different tickers in next loop
                opening = True
            
            End If
            
            'when the ticker value changes
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                'adds next volume value
                vol = vol + ws.Cells(i, 7).Value
                
                'gets the closing value
                yrEnd = ws.Cells(i, 6).Value
                
                'calculates the yrChange
                yrChange = yrEnd - yrStart
                
                'inputs our volume value
                ws.Range("L" & Summary_Table_Row).Value = vol
                
                'inputs our year change value
                ws.Range("J" & Summary_Table_Row).Value = yrChange
                
                'calculates and inputs our percent change value
                'I dont multiply by 100 because i formatted column K to be percents
                ws.Range("K" & Summary_Table_Row).Value = (yrChange / yrStart)
                    
                    'changes the color of the Year Change column, white is no change
                     If (ws.Cells(Summary_Table_Row, 10).Value > 0) Then
                     
                         'green for positive
                         ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                     ElseIf (ws.Cells(Summary_Table_Row, 10).Value < 0) Then
                     
                        'red for negative
                         ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
                    End If
                    
                    'go to next row in Summary Table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'resets our volume
                vol = 0
                
                'resets our opening value so we can store a new opening value
                opening = False
                
                Else
                    'adds next volume value
                    vol = vol + ws.Cells(i, 7).Value
                    
                End If
          Next i
          
    'counts the rows in our filtered ticker value column
    FRowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Uses the min max function to fin the largest/smallest value
    bigInc = Application.WorksheetFunction.Max(ws.Range("K1:K" & FRowCount))
    bigDec = Application.WorksheetFunction.Min(ws.Range("K1:K" & FRowCount))
    bigVol = Application.WorksheetFunction.Max(ws.Range("L1:L" & FRowCount))
    
    'places the variables we found above in the desired cells
    ws.Cells(2, 16).Value = bigInc
    ws.Cells(3, 16).Value = bigDec
    ws.Cells(4, 16).Value = bigVol
    
    'scans through our filtered ticker column to corresponding ticker values
            For j = 2 To FRowCount
            
            'finds ticker value for bigInc
            If bigInc = ws.Cells(j, 11).Value Then
                ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
                End If
                
            'finds ticker value for bigDec
            If bigDec = ws.Cells(j, 11).Value Then
                ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
                End If
            
            'finds ticker value for largest total volume
            If bigVol = ws.Cells(j, 12).Value Then
                ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
            End If
         
        Next j
          
    Next ws

End Sub
