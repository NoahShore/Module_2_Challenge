Attribute VB_Name = "Module1"
Sub TotalCharges():

'This section of code relating to copying across worksheets is from the link github in the ReadME File
    Dim ws As Worksheet

    For Each ws In Sheets
        Worksheets(ws.Name).Activate

    Dim Ticker As String

    Dim Volume_Total As Double
        Volume_Total = 0

    Dim summary_table_row As Integer
        summary_table_row = 2

    Dim Yearly_Change As Integer
        Yearly_Change = 0

    Dim Percent_Change As Double
        Percent_Change = 0

    Dim year_end As Double

    Dim year_start As Double




For i = 2 To 22771

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
    
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        year_end = Cells(i, 6).Value
        
        Range("I" & summary_table_row).Value = Ticker
    
        Range("L" & summary_table_row).Value = Volume_Total
    
        summary_table_row = summary_table_row + 1
    
        Volume_Total = 0
    
    Else
        
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        year_start = Cells(i, 3).Value
        
    End If
    
Next i

summary_table_row = 2

For i = 2 To 22771

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        year_end = Cells(i, 6).Value
    
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
        year_start = Cells(i, 3).Value
        
    End If
    
    If year_start > 0 And year_end > 0 Then
    
        increase = year_end - year_start
    
        percent_increase = increase / year_start
    
        Range("J" & summary_table_row).Value = increase
    
        Range("K" & summary_table_row).Value = FormatPercent(percent_increase)
    
        year_end = 0
    
        year_start = 0
        
        summary_table_row = summary_table_row + 1
    End If

Next i

' This portion of code relating to greatest changes is from the linked GitHub in the ReadME File
max_increase = WorksheetFunction.Max(ActiveSheet.Columns("k"))
max_decrease = WorksheetFunction.Min(ActiveSheet.Columns("k"))
max_volume_change = WorksheetFunction.Max(ActiveSheet.Columns("l"))

Range("p2").Value = FormatPercent(max_increase)
Range("p3").Value = FormatPercent(max_decrease)
Range("p4").Value = FormatPercent(max_volume_change)

For i = 2 To 22771
    If max_increase = Cells(i, 11).Value Then
        Range("O2").Value = Cells(i, 9).Value
    ElseIf max_decrease = Cells(i, 11) Then
        Range("O3").Value = Cells(i, 9).Value
    ElseIf max_volume_change = Cells(i, 11).Value Then
        Range("O4").Value = Cells(i, 9).Value
    End If
Next i

For i = 2 To 22771
    
    If Cells(i, 10).Value > 0 Then
        
        Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 10).Value < 0 Then
        
        Cells(i, 10).Interior.ColorIndex = 3
    
    End If

Next i

For i = 2 To 22771
    
    If Cells(i, 11).Value > 0 Then
        
        Cells(i, 11).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 11).Value < 0 Then
        
        Cells(i, 11).Interior.ColorIndex = 3
    
    End If

Next i

Next ws
    
End Sub
