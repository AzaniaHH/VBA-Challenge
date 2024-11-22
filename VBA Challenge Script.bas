Attribute VB_Name = "Module1"
Sub AnalyzeStocksWithFormatting()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim SummaryRow As Integer
    Dim CalculatedRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Initialize variables
        SummaryRow = 2
        TotalVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0

        ' Set column headers for summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Find the last row
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Loop through rows
        For i = 2 To LastRow
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ' Check if the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculate OpenPrice safely
                If ws.Cells(i, 7).Value > 0 Then
                    CalculatedRow = Int(i - (TotalVolume / ws.Cells(i, 7).Value) + 1)
                Else
                    CalculatedRow = i ' Default to the current row if division is invalid
                End If

                ' Ensure CalculatedRow is valid
                If CalculatedRow < 1 Or CalculatedRow > LastRow Then
                    OpenPrice = 0 ' Default to 0 if row is invalid
                ElseIf IsEmpty(ws.Cells(CalculatedRow, 3).Value) Then
                    OpenPrice = 0 ' Default to 0 if the cell is empty
                Else
                    OpenPrice = ws.Cells(CalculatedRow, 3).Value
                End If

                ' Calculate quarterly change and percentage change
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If

                ' Write summary data
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Update greatest values
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If

                ' Reset for next ticker
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
            End If
        Next i

        ' Apply conditional formatting to "Quarterly Change" column
        Dim ChangeRange As Range
        Set ChangeRange = ws.Range(ws.Cells(2, 10), ws.Cells(SummaryRow - 1, 10))
        
        ' Clear existing conditional formatting
        ChangeRange.FormatConditions.Delete
        
        ' Apply green for positive change
        With ChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(0, 255, 0) ' Green
        End With

        ' Apply red for negative change
        With ChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .Interior.Color = RGB(255, 0, 0) ' Red
        End With

        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
    Next ws

    MsgBox "Analysis Complete with Conditional Formatting!"
End Sub

