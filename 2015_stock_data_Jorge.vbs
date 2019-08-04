Sub Tickerfunction()

' the ticker is a string,

    Dim Ticker As String
' double data is the best data to represent the total volume
    Dim vol As Double
    vol = 0

    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Total Stock Vol"


    Summary_Table_Row = 2

    For i = 2 To Sub Tickerfunction()

' the ticker is a string,

    Dim Ticker As String
' double data is the best data to represent the total volume
    Dim vol As Double
    vol = 0

    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Total Stock Vol"


    Summary_Table_Row = 2

    For i = 2 To 760192

      If year_open = 0 Then

          year_open = Cells(i, 3).Value
      End If

      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          yearly_change = year_close - year_open


          Ticker = Cells(i, 1).Value


          vol = vol + Cells(i, 7).Value

'Output of value

          Range("I" & Summary_Table_Row).Value = Ticker

          Range("J" & Summary_Table_Row).Value = vol


'Move down the summary table

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0


      Else

          vol = vol + Cells(i, 7).Value


      End If


    Next i

End Sub

      If year_open = 0 Then

          year_open = Cells(i, 3).Value
      End If

      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          yearly_change = year_close - year_open


          Ticker = Cells(i, 1).Value


          vol = vol + Cells(i, 7).Value

'Output of value

          Range("I" & Summary_Table_Row).Value = Ticker

          Range("J" & Summary_Table_Row).Value = vol


'Move down the summary table

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0


      Else

          vol = vol + Cells(i, 7).Value


      End If


    Next i

End Sub