sub stock_data()

    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count

    total_volume = 0
    summary_row = 2
    ticker = ""

    for r = 2 to NumRows
        If Cells (r,1).Value <> Cells(r+1, 1).Value Then
        ticker = Cells(r,1).Value
        total_volume = total_volume + Cells(r,3).Value
        Range("H" & summary_row).Value = ticker
        Range("I" & summary_row).Value = total_volume
        summary_row = summary_row + 1
        total_volume = 0
    Else
        total_volume = total_volume + Cells(r,3).Value
    End If
    Next r

End sub