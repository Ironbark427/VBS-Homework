Sub StockChecker1()
    Dim Last_Row As Long
    Dim First_Date As Long
    Dim Last_Date As Long
    Dim ShareTotal As Variant
    Dim First_Val As Double
    Dim Last_Val As Double
    Dim PerChng As Double
    Dim AmtChng As Double
    Dim StockRow As String
    Dim ws As Integer

    Application.ScreenUpdating = False
    
'    For ws = 1 To Application.Sheets.Count
    
    Worksheets(3).Activate

        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        ShareTotal = 0
        c = 1
        
        Cells(1, 12) = "Ticker"
        Cells(1, 13) = "Yearly Change"
        Cells(1, 14) = "Percent Change"
        Cells(1, 15) = "Total Stock Volume"
        
        Range("G:G").NumberFormat = "0"
        
        For i = 2 To Last_Row
            ShareTotal = ShareTotal + Cells(i, 7)
            If Not Cells(i, 1).Value = Cells(i + 1, 1) Then
                c = c + 1
                Cells(c, 12).Value = Cells(i, 1)
                Cells(c, 15).Value = ShareTotal
                ShareTotal = 0
            End If
        Next i
        
        t = 2
        
        For i = 2 To c
            For p = t To Last_Row
                If Cells(i, 12) = Cells(p, 1) And First_Val = 0 Then
                    First_Val = Cells(p, 3)
                ElseIf Cells(i, 12) = Cells(p, 1) And First_Val > 0 And Not Cells(i, 12) = Cells(p + 1, 1) Then
                    Last_Val = Cells(p, 6)
                End If
                t = p
                If First_Val > 0 And Last_Val > 0 Then p = Last_Row
            Next p
            
                AmtChng = Last_Val - First_Val
                PerChng = (Last_Val - First_Val) / Last_Val
        
                Cells(i, 13) = AmtChng
                If AmtChng <= 0 Then
                    Cells(i, 13).Interior.ColorIndex = 3
                Else
                    Cells(i, 13).Interior.ColorIndex = 4
                End If
                Cells(i, 14) = PerChng
                Range("N:N").NumberFormat = "0.00%"
                First_Val = 0
                Last_Val = 0
        Next i
        
    
  '  Next ws

    Application.ScreenUpdating = True
    
End Sub