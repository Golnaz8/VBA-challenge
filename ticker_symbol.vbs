Sub quarter_pt1()
    ' Create variables
    Dim i As Integer
    Dim ws as Worksheet
    Dim Ticker_Symbol as String
    Dim Table_Row as Integer
    Table_Row = 2

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"



        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        for i = 1 to LastRow

            if ws.cells(i + 1, 1).Value <> ws.cells(i , 1).Value Then

            ' Set the Ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value

            ' Print the Ticker symbol in each Table
            ws.Range("I" & Table_row).Value = Ticker_Symbol

            ' Add one to the table row
            Table_row = Table_row + 1
        Next i

        Table_row = 2

    Next ws

End Sub