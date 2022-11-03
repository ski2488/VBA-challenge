Sub StockSummary()
    
    'Initialize variables
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim DateOpen As Long
    Dim DateClose As Long
    Dim SummaryRow As Integer

    'Initialize Values
    Volume = 0
    OpenPrice = 0
    ClosePrice = 0
    SummaryRow = 2

    ' Loop through all sheets
    For Each ws In Worksheets
    'Insert Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Annual Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Find Last Row #
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Date Open and Close unique to each sheet
    DateOpen = ws.Cells(2, 2).Value
    DateClose = ws.Range("B" & lastRow).Value

    'Loop thru all stocks
    For i = 2 To lastRow

    'Check if we are still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set Ticker
        Ticker = ws.Cells(i, 1).Value
        
        'Set Volume
        Volume = Volume + ws.Cells(i, 7).Value
        'Set ClosePrice
        ClosePrice = ws.Cells(i, 6).Value
        
        'Print the Ticker in the Summary Table
        ws.Range("I" & SummaryRow).Value = Ticker
        
        'Print the Yearly Change in the Summary Table
        ws.Range("J" & SummaryRow).Value = ClosePrice - OpenPrice
        
        'Conditional Format
        If ws.Range("J" & SummaryRow).Value > 0 Then
        ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
        Else
        ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
        End If
        
        'Print the Percent Change in the Summary Table
        ws.Range("K" & SummaryRow).Value = (ClosePrice - OpenPrice) / OpenPrice
        
        'Print the Volume in the Summary Table
        ws.Range("L" & SummaryRow).Value = Volume
        
        'Add one to the summary table row
        SummaryRow = SummaryRow + 1
      
        'Reset the Volume
        Volume = 0
    
    ElseIf ws.Cells(i, 2).Value = DateOpen Then
           'Set Open Price
           OpenPrice = ws.Cells(i, 3).Value
           'Add to the Volume
           Volume = Volume + ws.Cells(i, 7).Value
            
    Else
    
        'Add to the Volume
        Volume = Volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i


'Reset SummaryRow for each new sheet
SummaryRow = 2

Next ws

End Sub
