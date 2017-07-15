Sub StockSolution()
'Declare Variables and Data Types'
Dim Stock As String, Volume As Double, First As Long, EndRow As Single, change As Double, total As Single
Dim dailyChange As Single, avgChange As Single, days As Integer, j As Long, k As Long, l As Long, h As Long, c As Long
    'Sort Sheet'
    Sheets("Stock_data_2016").Activate
    Columns("A:G").Sort key1:=Range("A:A"), _
     order1:=xlAscending, Header:=xlYes
     'declare start values for markers'
     days = 0
     dailyChange = 0
     avgChange = 0
     Volume = 0
     j = 2
     k = 2
     c = 2
     l = 2
     h = 2
     First = 2
     'Loop through, get values. This process takes about 30 seconds.'
     For i = 2 To 797711

         If Cells(i + 1, 1).Value = Cells(i, 1) Then
             Volume = Volume + Cells(i, 7).Value
             days = days + 1
             dailyChange = dailyChange + (Cells(i, 4) - Cells(i, 5))
         Else
             Volume = Volume + Cells(i, 7).Value
             ' Range("K" & Rows.Count).End(xlUp).Offset(1).Value = Volume
             Cells(k, 11).Value = Volume
             k = k + 1
             Volume = 0
             days = days + 1
             dailyChange = dailyChange + (Cells(i, 4) - Cells(i, 5))
             avgChange = dailyChange / days
             days = 0
             dailyChange = 0
             ' Range("J" & Rows.Count).End(xlUp).Offset(1).Value = avgChange
             Cells(j, 10).Value = avgChange
             j = j + 1
             avgChange = 0
             EndRow = i
             total = (Cells(EndRow, 6) - Cells(First, 3))
             Cells(h, 8).Value = total
             h = h + 1
             change = Round((total / Cells(First, 3) * 100), 2)
             Cells(c, 9).Value = change
             c = c + 1
             Cells(l, 12).Value = Cells(i, 1).Value
             l = l + 1
             First = i + 1
         End If

     Next i
     'Copy and paste data over.'
     Range("L2:L797711").Cut
     Sheets.Add.Name = "StockTransfer"
     Sheets("StockTransfer").Activate
     Range("A2:A797711").Select
     ActiveSheet.Paste
     Application.CutCopyMode = False
     Sheets("Stock_data_2016").Activate
     Range("H2:H797711").Cut
     Sheets("StockTransfer").Activate
     Range("B2:B797711").Select
     ActiveSheet.Paste
     Sheets("Stock_data_2016").Activate
     Range("I2:I797711").Cut
     Sheets("StockTransfer").Activate
     Range("C2:C797711").Select
     ActiveSheet.Paste
     Cells(1, 1).Value = "Ticker"
     Cells(1, 2).Value = "Total Change"
     Cells(1, 3).Value = "% Change"
     Cells(1, 4).Value = "Avg. Daily Change"
     Cells(1, 5).Value = "Volume"
     Sheets("Stock_data_2016").Activate
     Range("K2:K797711").Cut
     Sheets("StockTransfer").Activate
     Range("E2:E797711").Select
     ActiveSheet.Paste
     Sheets("Stock_data_2016").Activate
     Range("J2:J797711").Cut
     Sheets("StockTransfer").Activate
     Range("D2:D797711").Select
     ActiveSheet.Paste
    'Declare Headers for greatest values.'
    Cells(2, 7).Value = "Greatest Volume"
    Cells(5, 7).Value = "Greatest % Increase"
    Cells(8, 7).Value = "Greatest % Decrease"
    Cells(11, 7).Value = "Greatest Avg. Change"

    Dim grtVol, grtIncr, grtDecr, grtAvgChg As Double
    Dim grtVolRng, grtIncrRng, grtDecrRng, grtAvgChgRng As Range
    Dim sht As Worksheet
    'Process to figure out greatest values.'
    Set sht = Worksheets("StockTransfer")

    Set grtVolRng = Worksheets("StockTransfer").Range("E2:E3161")
    grtVol = Application.WorksheetFunction.Max(grtVolRng)

    Cells(2, 8).Value = grtVol

    Set grtIncrRng = Worksheets("StockTransfer").Range("C2:C3161")
    grtIncr = Application.WorksheetFunction.Max(grtIncrRng)

    Cells(5, 8).Value = grtIncr / 100
    Cells(5, 8).NumberFormat = "0%"


    Set grtDecrRng = Worksheets("StockTransfer").Range("C2:C3161")
    grtDecr = Application.WorksheetFunction.Min(grtDecrRng)

    Cells(8, 8).Value = grtDecr / 100
    Cells(8, 8).NumberFormat = "0.00%"


    Set grtAvgChgRng = Worksheets("StockTransfer").Range("D2:D3161")
    grtAvgChg = Application.WorksheetFunction.Max(grtAvgChgRng)

    Cells(11, 8).Value = grtAvgChg
    'Loop through, change colors to red or green in total depending on whether it lost money or gained.'
    'Loop through, find the ticker value that corresponds to the greatest Value.'
    'Loop through, change Column C to reflect percentage.'
    For i = 2 To 3161
        If Cells(i, 2).Value < 0 Then
            Cells(i, 2).Interior.ColorIndex = 3
        Else
            Cells(i, 2).Interior.ColorIndex = 4
        End If
        
        If Cells(i, 3).Value = grtDecr Then
            Cells(8, 9).Value = Cells(i, 1).Value
        End If

        If Cells(i, 3).Value = grtIncr Then
            Cells(5, 9).Value = Cells(i, 1).Value
        End If
        
        If Cells(i, 4).Value = grtAvgChg Then
            Cells(11, 9).Value = Cells(i, 1).Value
        End If

        If Cells(i, 5).Value = grtVol Then
            Cells(2, 9).Value = Cells(i, 1).Value
        End If

        Cells(i, 3).Value = Cells(i, 3).Value / 100
        Cells(i, 3).NumberFormat = "0.00%"
    Next i

    'Format'
    sht.Cells.EntireColumn.AutoFit

End Sub