Attribute VB_Name = "Module1"
Sub VBA_Challenge()

For Each ws In Worksheets
Dim Worksheet As String
Worksheet = ws.Name

Dim Ticker As String
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim Summary_Table As Double
Summary_Table = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Year_Open = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value
            
            Year_Close = ws.Cells(i, 6).Value
            Yearly_Change = Year_Close - Year_Open
            
            If Year_Open <> 0 Then
            Percent_Change = Yearly_Change / Year_Open

            Else
            
            Year_Open = 0
            
            End If
            
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & Summary_Table).Value = Ticker
            ws.Range("J" & Summary_Table).Value = Yearly_Change
            ws.Range("K" & Summary_Table).Value = Format(Percent_Change, "0.00%")
            ws.Range("L" & Summary_Table).Value = Total_Stock_Volume
            Summary_Table = Summary_Table + 1
            Total_Stock_Volume = 0
    
            Year_Open = ws.Cells(i + 1, 3).Value
            
        Else

            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

        End If

    Next i
    
lastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
 
    For x = 2 To lastrow
    
        If ws.Cells(x, 10).Value < 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 3
        
        Else
            ws.Cells(x, 10).Interior.ColorIndex = 4
        
        End If
    
    Next x

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

    For j = 2 To lastrow
    
    If ws.Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
    ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
    ws.Cells(2, 17).Value = Format(ws.Cells(j, 11).Value, "0.00%")
    
    ElseIf ws.Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
    ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
    ws.Cells(3, 17).Value = Format(ws.Cells(j, 11).Value, "0.00%")
      
    ElseIf ws.Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
    ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
    
    End If
    
    Next j

 Next ws
     
End Sub
