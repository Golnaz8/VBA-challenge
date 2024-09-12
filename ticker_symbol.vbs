Sub quarter_pt1()
    ' Create variables
    Dim i As Long
    Dim ws As Worksheet
    Dim Ticker_Symbol As String
    Dim Table_Row As Long
    Table_Row = 2

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Determine the Last Row
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the Ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value

            ' Print the Ticker symbol in each Table
            ws.Range("I" & Table_Row).Value = Ticker_Symbol

            ' Add one to the table row
            Table_Row = Table_Row + 1

            End If

        Next i

        Table_Row = 2

    Next ws

End Sub