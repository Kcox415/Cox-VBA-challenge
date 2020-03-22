Attribute VB_Name = "Module1"
Sub stock_loop()
    On Error Resume Next
' Declare variables
    Dim ws_count As Integer
    Dim i As Integer
    Dim j As Long
    Dim k As Integer
    Dim l As Integer
    Dim total_vol As Double
    Dim open_price_beginning As Double
    Dim close_price_end As Double
    k = 2
    ws_count = ActiveWorkbook.Worksheets.Count
' Loop through all sheets
    For i = 1 To ws_count
' Create headers for new data
        ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
' Set initial stock volume equal to zero
        total_vol = 0
' Loop through the entire sheet
        For j = 2 To ActiveWorkbook.Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell).Row
            If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j + 1, 1).Value Then
' Loop through each sheet to find yearly and percentage changes in stock prices
                close_price_end = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                ActiveWorkbook.Worksheets(i).Cells(k, 10).Value = close_price_end - open_price_beginning
                ActiveWorkbook.Worksheets(i).Cells(k, 11).Value = (close_price_end - open_price_beginning) / open_price_beginning
                open_price_beginning = 0
                close_price_end = 0
' Loop through each sheet and add the volume for each stock
                ActiveWorkbook.Worksheets(i).Cells(k, 9).Value = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                ActiveWorkbook.Worksheets(i).Cells(k, 12).Value = total_vol
                k = k + 1
                total_vol = 0
            ElseIf ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j, 1).Value Then
                open_price_beginning = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            Else
                total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            End If
        Next j
        
' Highlighting positive changes green and negative changes red
        For l = 2 To ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count
            ActiveWorkbook.Worksheets(i).Cells(l, 11).Style = "Percent"
            If ActiveWorkbook.Worksheets(i).Cells(l, 11).Value > 0 Then
                With ActiveWorkbook.Worksheets(i).Cells(l, 11).Interior
                    .ColorIndex = 4
                    .TintAndShade = 0.6
                End With
            Else
                With ActiveWorkbook.Worksheets(i).Cells(l, 11).Interior
                    .ColorIndex = 3
                    .TintAndShade = 0.6
                End With
            End If
        Next l
        
' Start loop again on the next sheet
        k = 2
    Next i
End Sub







