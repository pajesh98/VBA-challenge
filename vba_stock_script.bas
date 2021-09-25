Attribute VB_Name = "Module1"
Sub stocks()

'Macro will run for each worksheets
For Each ws In Worksheets

'Column Headers Label
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 ws.Range("O2").Value = "Greatest % Increase"
 ws.Range("O3").Value = "Greatest % Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"

'Define variable for each dimension.
Dim TickerName As String
Dim LastRow As Long
Dim Total_Stocks As Double

Dim SummaryTableRow As Long
SummaryTableRow = 2
Dim Yearly_Open As Double
Dim Yearly_Close As Double
Dim Yearly_Change As Double
Dim PreviousAmount As Long
PreviousAmount = 2
Dim PercentChange As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim LastRowValue As Long
Dim GreatestTotalVolume As Double
Total_Stocks = 0
GreatestTotalVolume = 0
GreatestIncrease = 0
GreatestDecrease = 0

 ' Below function finds the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' Add To Ticker Total Volume
            Total_Stocks = Total_Stocks + ws.Cells(i, 7).Value
            ' Check if we are still within the same ticker name if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then



                ' Set Ticker Name
                TickerName = ws.Cells(i, 1).Value
                ' Display the ticker name in column I
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ' subtotal of each ticker and number format for the totals
                ws.Range("L" & SummaryTableRow).NumberFormat = "#,##0"
                ws.Range("L" & SummaryTableRow).Value = Total_Stocks
                Total_Stocks = 0


                ' Set Open, Close and Yearly Change Name
                Yearly_Open = ws.Range("C" & PreviousAmount)
                Yearly_Close = ws.Range("F" & i)
                Yearly_Change = Yearly_Close - Yearly_Open
                ws.Range("J" & SummaryTableRow).Value = Yearly_Change

                ' Determine Percent Change
                If Yearly_Open = 0 Then
                    PercentChange = 0
                Else
                    Yearly_Open = ws.Range("C" & PreviousAmount)
                    PercentChange = Yearly_Change / Yearly_Open
                End If
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Formatting
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add One To The Summary Table Row
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i
' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start Loop For Final Results
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "#,##0"
            

 
 Next ws
 
 End Sub

