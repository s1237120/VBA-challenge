Attribute VB_Name = "Module1"
Sub StockAnalysis():
    For Each ws In Worksheets
    
        Dim LastRow As Long
        Dim LastColumn As Long
        Dim Ticker As String
        Dim Percent_Change As Double
        Dim Total_Stock As Long
        Dim Summary_Table_Row As Integer
        Dim Yearly_Change As Double
        
        Summary_Table_Row = 2
        Total_Stock = 0
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
      
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(2, 3).Value
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            Percent_Change = (Yearly_Change / ws.Cells(2, 3).Value)
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change * 100
            
            
            Total_Stock = Total_Stock + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock
            Total_Stock = 0
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            End If
        
        Next i
        
        
        
        For i = 2 To LastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
         For i = 2 To LastRow
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 11).Value < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
        
        
        
            
        
      Next ws

End Sub

