Attribute VB_Name = "Module1"
Sub ticker()

  ' Initial variable for Ticker Name
  Dim Ticker_Name As String

  ' Initial variable for ticker Volume
  Dim Ticker_Volume As Double
  Ticker_Volume = 0

  ' Location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Column names of summary table
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Total Stock Volume"

  ' Loop through all ticker Names
  For i = 2 To 760192

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker Volume
      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Total Volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Volume
      Ticker_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Ticker Volume
      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

    End If

  Next i

End Sub



