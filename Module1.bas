Attribute VB_Name = "Module1"
Sub stockTicker()

For Each ws In Worksheets

    Dim ticker As String
    Dim vol As Double
    vol = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim year_open As Double
    Dim year_close As Double
    Dim year_Percent As Double
    year_Percent = 0
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'create headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly_change"
    Range("K1").Value = "Yearly_percentage"
    Range("L1").Value = "Total Stock Volume"
'loop through tickers
    For i = 2 To LastRow
    If year_open = 0 Then
          year_open = Cells(i, 3).Value
      End If

      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          yearly_change = year_close - year_open
          year_Percent = (yearly_change / year_open)

          ticker = Cells(i, 1).Value

          vol = vol + Cells(i, 7).Value

          Range("j" & Summary_Table_Row).Value = yearly_change
          Range("I" & Summary_Table_Row).Value = ticker

          Range("K" & Summary_Table_Row).Value = year_Percent
          Range("K" & Summary_Table_Row).Style = "Percent"
          Range("L" & Summary_Table_Row).Value = vol
          Summary_Table_Row = Summary_Table_Row + 1
          'Reset
          vol = 0
      Else
          vol = vol + Cells(i, 7).Value
      End If


    Next i
    
Next ws

End Sub
