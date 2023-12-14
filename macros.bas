VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StocksSort():
'Creating loop
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
'Summary tables

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"    
    
    'Define variables
    Dim tick As String
    Dim summary_row As Integer
    summary_row = 2
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim open_price, close_price As Double
    open_price = ws.Cells(2, "C").Value
    close_price = 0
    Dim yearly_change, change_per As Double 
    Dim total_volume As LongLong 
    total_volume = 0

    For i = 2 To lastrow
        total_volume = total_volume + ws.Cells(i, "G")
        If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
            ws.Cells(summary_row, "I").Value = ws.Cells(i, "A")
            ws.Cells(summary_row, "J").Value = ws.Cells(i, "F") - open_price
            ws.Cells(summary_row, "K").Value = (ws.Cells(i, "F") - open_price) / open_price

            If ws.Range("J" & sumrow).Value > 0 Then
                ws.Range("J" & sumrow).Interior.ColorIndex = 4

            Else
                ws.Range("J" & sumrow).Interior.ColorIndex = 3
            End If
 
            ws.Cells(sumrow, "K").NumberFormat = "0.00%"
            ws.Cells(sumrow, "L").Value = total_volume
            

            change_per = (ws.Cells(i, "F") - open_price) / open_price
            

            summary_row = sumrow + 1

            open_price = ws.Cells(i + 1, "C")

            tick = ws.Cells(i, "A")
            If change_per > max_increase Then
                max_increase = change_per
                max_increase_tick = tick
            ElseIf change_per < max_decrease Then
                max_decrease = change_per
                max_decrease_tick = tick
            End If
            
            If total_volume > max_volume Then
            max_volume = total_volume
            max_volume_tick = tick
            End If
            total_volume = 0
            ' Other than ending all ifs and setting up nexts, last thing I did was print where I want everything to go in Summary Table 2
            ws.Cells(2, "P").Value = max_increase_tick
            ws.Cells(2, "Q").Value = max_increase
            ws.Cells(2, "Q").NumberFormat = "0.00%"
            ws.Cells(3, "P").Value = max_decrease_tick
            ws.Cells(3, "Q").Value = max_decrease
            ws.Cells(3, "Q").NumberFormat = "0.00%"
            ws.Cells(4, "P").Value = max_volume_tick
            ws.Cells(4, "Q").Value = max_volume
            
        End If
    Next i
Next ws
                
End Sub
