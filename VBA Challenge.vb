Sub MultiYrStock()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim A As Worksheet
Dim ticker As String
Dim i As Integer
Dim lastrow As Long
Dim totalvolume As Double
Dim qtrlychg As Double
Dim pctchg As Double
Dim open_date As Date




totalvolume = 0


lastrow = Cells(Rows.Count, 10).End(xlUp).Row

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
      totalvolume = totalvolume + Cells(i, 7).Value
      Range("I" & Summary_Table_Row).Value = ticker
      
      Range("L" & Summary_Table_Row).Value = totalvolume
      
    Summary_Table_Row = Summary_Table_Row + 1
      totalvolume = 0

    Else

      totalvolume = totalvolume + Cells(i, 7).Value

    End If



    If open_date = #1/2/2022# Then
        ElseIf close_date = #3/31/2022# Then
       qtrlychg = Cells(i, 3).Value - Cells(i, 6).Value
       Range("J" & Summary_Table_Row).Value = qtrlychg
        
    
    End If

    


  Next i

For i = 2 To lastrow
        If i < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 4
        End If
        
Next i

End Sub

