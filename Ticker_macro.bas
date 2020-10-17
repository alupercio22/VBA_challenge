Attribute VB_Name = "Module1"
Sub ticker_macro()

'dims needed
Dim ticker As String
Dim final_row As Long
Dim ws As Worksheet
Dim yr_open As Double
Dim yr_close As Double
Dim percent As Double
Dim volume As Variant
Dim change As Double
Dim summary_row As Integer
Dim total_volume As Long


'On Error Resume Next
'set worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    'find final_row
    final_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    summary_row = 2
    'set output headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    yr_open = ws.Cells(2, 3).Value
        'set i
        For i = 2 To final_row
            
            volume = ws.Cells(i, 7).Value + volume

            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yr_close = ws.Cells(i, 6).Value
                change = yr_close - yr_open
                If yr_open = 0 Then
                    percent = 0
                Else
                    percent = change / yr_open
                End If
                
                ticker = ws.Cells(i, 1).Value
                ws.Cells(summary_row, 9).Value = ticker
                ws.Cells(summary_row, 10).Value = change
                ws.Cells(summary_row, 11).Value = percent
                ws.Cells(summary_row, 12).Value = volume
                
                yr_open = ws.Cells(i + 1, 3).Value
                summary_row = summary_row + 1
            
                volume = 0
        
            
            End If
        Next i
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    Dim k As Range
    j = ws.Cells(Rows.Count, "J").End(xlUp).Row
    For i = 2 To j
        If ws.Cells(i, 10).Value >= 0 Then
           ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If

        If ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
            ws.Cells(2, 17).Value = ws.Cells(i, 11)
            ws.Cells(2, 16).Value = ws.Cells(i, 9)
        End If
        
        If ws.Cells(i, 11).Value < ws.Cells(3, 17) Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11)
        ws.Cells(3, 16).Value = ws.Cells(i, 9)
        End If
        
        If ws.Cells(i, 12).Value > ws.Cells(4, 17) Then
        ws.Cells(4, 17).Value = ws.Cells(i, 12)
        ws.Cells(4, 16).Value = ws.Cells(i, 9)
        End If
    Next i
    
    Next ws
End Sub

