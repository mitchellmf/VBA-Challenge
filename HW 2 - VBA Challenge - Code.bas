Attribute VB_Name = "RibbonX_Code"
Sub Stocks()
    Dim DataCount As Integer
    Dim DataCountKeep As Integer
    Dim TRC As Integer
    Dim StockVol As Long
    Dim StockVolKeep As Long
    Dim i As Long
    Dim j As Long
    Dim Ticker As String
DataCount = 0
StockVol = 0
TRC = 1
For i = 2 To 70926
    For j = 2 To 2
        If Cells(i, j).Value > 0 Then
            DataCount = DataCount + 1
            ' StockVol = StockVol + Cells(i, 7)
                If Cells(i + 1, j - 1).Value <> Cells(i, j - 1).Value Then
                    Ticker = Cells(i, j - 1).Value
                    DataCountKeep = DataCount
                    StockVolKeep = StockVol
                    TRC = TRC + 1
                    Cells(TRC, 9).Value = Ticker
                    Cells(TRC, 20).Value = DataCountKeep
                    Cells(TRC, 21).Value = Cells(i - DataCountKeep + 1, 3).Value
                    Cells(TRC, 22).Value = Cells(i, 6).Value
                    Cells(TRC, 23).Value = Cells(i - DataCountKeep + 1, 2).Value
                    Cells(TRC, 24).Value = Cells(i, 2).Value
                    Yrly_Chng = Cells(TRC, 22).Value - Cells(TRC, 21).Value
                    Yrly_PctChng = (Cells(TRC, 22).Value - Cells(TRC, 21).Value) / Cells(TRC, 21).Value
                    Cells(TRC, 10).Value = Yrly_Chng
                    Cells(TRC, 11).Value = Yrly_PctChng
                    Cells(TRC, 11).Style = "Percent"
                    If Yrly_Chng < 0 Then
                        Cells(TRC, 10).Interior.ColorIndex = 3
                        ElseIf Yrly_Chng > 0 Then Cells(TRC, 10).Interior.ColorIndex = 4
                    End If
                    Cells(TRC, 12).Value = StockVolKeep
                    DataCount = 0
                    StockVol = 0
                End If
        End If
     Next j
    Next i
max_inc = WorksheetFunction.Max(Range("K2:K501"))
min_inc = WorksheetFunction.Min(Range("K2:K501"))
max_vol = WorksheetFunction.Min(Range("L2:L501"))
Cells(2, 15).Value = max_inc
Cells(3, 15).Value = min_inc
Cells(2, 15).Style = "Percent"
Cells(3, 15).Style = "Percent"
Cells(4, 15).Value = max_vol
' Ticker
' Yearly Change
' Percent Change
' Total Stock Volume
' Ticker
' Value
' Greatest % Increase
' Greatest % Decrease
' Greatest Total Volume
End Sub



