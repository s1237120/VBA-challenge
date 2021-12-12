Attribute VB_Name = "Module1"
Sub StockAnalysis():

    Dim LastRow As Long
    Dim LastColumn As Long
    Dim Ticker As String
    Dim Percent_Change As Long
    Dim Total_Stock As Long
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("B:B").NumberFormat = "m/d/yyyy"
    
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        
        Summary_Table_Row = Summary_Table_Row + 1
        
    
        End If
    
    Next i
    
        
    
    
 
    'Range("B") = Format(Date, "yyyy-mm-dd")
    
    
    

End Sub
