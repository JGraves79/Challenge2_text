Attribute VB_Name = "Module1"
Sub Delete_columns()

Dim ws As Worksheet

For Each ws In Worksheets
        ' Activate the worksheet
        ws.Activate

Range("J:Q").Delete

Next ws

End Sub


Sub Format_Date()

Dim wks As Worksheet
Dim lastrow As String
Dim Ldate As Date
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim j As Integer

For Each wks In Worksheets
 
lastrow = wks.Cells(Rows.Count, "B").End(xlUp).Row

For j = 2 To lastrow

a = Int(Left(wks.Range("B" & j), 4))
b = Int(Right(wks.Range("B" & j), 2))
c = Int(Mid(wks.Range("B" & j), 5, 2))

Ldate = DateSerial(a, c, b)

Range("B" & j) = Ldate

Next j

Next wks


End Sub

Sub TickerLoop()
    Dim ws As Worksheet
    Dim Ticker_Name As String
    Dim Summary_row As Integer
    Dim total_vol As Double
    Dim open_val As Double
    Dim close_val As Double
    Dim qtr_chg As Double
    Dim per_chg As Double
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String

    ' Initialize variables
    total_vol = 0
    Summary_row = 1
    max_increase = -9999
    max_decrease = 9999
    max_volume = 0

    ' Loop through each worksheet
    For Each ws In Worksheets
        ' Activate the worksheet
        ws.Activate

        ' Name header rows
        If ws.Range("J1").Value = "" Then ws.Range("J1").Value = "Ticker"
        If ws.Range("K1").Value = "" Then ws.Range("K1").Value = "Quarterly Change"
        If ws.Range("L1").Value = "" Then ws.Range("L1").Value = "Percent Change"
        If ws.Range("M1").Value = "" Then ws.Range("M1").Value = "Total Stock Volume"

        ' Reset variables for each worksheet
        total_vol = 0
        Summary_row = 1
        max_increase = -9999
        max_decrease = 9999
        max_volume = 0

        ' Iterate through each row in the worksheet
        For i = 2 To ws.Cells(Rows.Count, "B").End(xlUp).Row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                total_vol = total_vol + ws.Cells(i, 7).Value

                ' Get closing value of the last quarter
                close_val = ws.Cells(i, 6).Value

                ' Add total volume together
                Summary_row = Summary_row + 1

                ws.Cells(Summary_row, 10).Value = Ticker_Name

                ' Calculate quarterly change
                qtr_chg = close_val - open_val
                ws.Cells(Summary_row, 11).Value = qtr_chg

                ' Apply conditional formatting
                If qtr_chg > 0 Then
                    ws.Cells(Summary_row, 11).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf qtr_chg < 0 Then
                    ws.Cells(Summary_row, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                ' Calculate percent change
                If open_val <> 0 Then
                    per_chg = qtr_chg / open_val
                Else
                    per_chg = 0
                End If
                ws.Cells(Summary_row, 12).Value = per_chg
                ws.Cells(Summary_row, 12).NumberFormat = "0.00%"

                ws.Cells(Summary_row, 13).Value = total_vol

                ' Check for greatest % increase, % decrease, and total volume
                If per_chg > max_increase Then
                    max_increase = per_chg
                    max_increase_ticker = Ticker_Name
                End If

                If per_chg < max_decrease Then
                    max_decrease = per_chg
                    max_decrease_ticker = Ticker_Name
                End If

                If total_vol > max_volume Then
                    max_volume = total_vol
                    max_volume_ticker = Ticker_Name
                End If

                ' Reset values for next ticker
                total_vol = 0
                open_val = 0
                close_val = 0
            Else
                ' Add total volume together
                total_vol = total_vol + ws.Cells(i, 7).Value

                ' Get opening value of the first quarter
                If open_val = 0 Then
                    open_val = ws.Cells(i, 3).Value
                End If
            End If
        Next i

        ' Output the results for greatest % increase, % decrease, and total volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"

        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("Q3").Value = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
    Next ws

    MsgBox ("Worksheets Updated")
End Sub

