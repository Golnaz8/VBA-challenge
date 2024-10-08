Sub quarter_pt1()
    ' Create variables
    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker_Symbol As String
    Dim Table_Row As Long
    Dim Q_Open_Price As Double
    Dim Q_Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Counter As Long
    Dim Total_Stock_Volume As Double
    Table_Row = 2
    Counter = 0
    Total_Stock_Volume = 0

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' Set the Opening Price
                Q_Open_Price = ws.Cells(i - Counter, 3).Value
                
                ' Set the Closing Price
                Q_Close_Price = ws.Cells(i, 6).Value

                ' Calculate the Quarterly Change
                Quarterly_Change = Q_Close_Price - Q_Open_Price

                ' Calculate the TOtal Stuck Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                ' Print the Ticker symbol in each Table
                ws.Range("I" & Table_Row).Value = Ticker_Symbol

                ' Print the Quarterly Change in each Table
                ws.Range("J" & Table_Row).Value = Quarterly_Change
                ws.Range("J" & Table_Row).NumberFormat = "0.00"

                ' Add color formatting to Quarterly Change column
                If Quarterly_Change > 0 Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                ElseIf Quarterly_Change <0 Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                Else 
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 0
                End If

                ' Print the Percent Change in each Table
                If Q_Open_Price <> 0 Then
                    ws.Range("K" & Table_Row).Value = (Quarterly_Change / Q_Open_Price)
                    ws.Range("K" & Table_Row).NumberFormat = "0.00%"

                Else
                    ws.Range("K" & Table_Row).Value = 0
                End If

                ' Print the Total Stock Volume in each Table
                ws.Range("L" & Table_Row).Value = Total_Stock_Volume
                
                ' Add one to the table row
                Table_Row = Table_Row + 1

                ' Reset the Counter
                Counter = 0

                ' Reset the Total Stock Volume
                Total_Stock_Volume = 0

            Else
                Counter = Counter + 1
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If

        Next i

        Table_Row = 2

        ' Autofit to display data
        ws.Columns("A:L").AutoFit

    Next ws

End Sub