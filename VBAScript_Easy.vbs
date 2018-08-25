Sub Stock_Volume():
  Dim Ticker_Name As String
  Dim Stock_Total As Double
  Stock_Total = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2  
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Total Stock Volume"
  For i = 2 To 705713
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker_Name = Cells(i, 1).Value
      Stock_Total = Stock_Total + Cells(i, 7).Value
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      Range("J" & Summary_Table_Row).Value = Stock_Total
      Summary_Table_Row = Summary_Table_Row + 1
      Stock_Total = 0
    Else
      Stock_Total = Stock_Total + Cells(i, 7).Value
    End If
  Next i
End Sub
