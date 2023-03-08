Attribute VB_Name = "Module3"
Sub Hard()
    ' Note: must run moderate exercise first!
    Call Moderate

    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    Dim max, min As Double
    Dim rownum As Integer
    Dim min_row_index, max_row_index, max_total_vol_index As Integer
    Dim max_total_vol As Double

    max = 0
    min = 0
    max_total_vol = 0
    
    For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        ' replace min/max percentage change if we find lower/higher value
        If Cells(i, 11) > max Then
            max = Cells(i, 11)
            max_row_index = i
        End If
        
        If Cells(i, 11) < min Then
            min = Cells(i, 11)
            min_row_index = i
        End If
        
        ' replace max total volume value if higher value found
        If Cells(i, 12) > max_total_vol Then
            max_total_vol = Cells(i, 12)
            max_total_vol_index = i
        End If
    Next i
    
    ' Write out the values to specified cells
    Range("P2") = Cells(max_row_index, 9).Value
    Range("P3") = Cells(min_row_index, 9).Value
    Range("P4") = Cells(max_total_vol_index, 9).Value
    
    Range("Q2") = max
    Range("Q3") = min
    Range("Q4") = max_total_vol
    
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"

    Columns("O").AutoFit
    Columns("P").AutoFit
    Columns("Q").AutoFit
End Sub


Sub ClearHard()
    Call ClearModerate

    Columns("O:Q").ClearContents
    Columns("O:Q").ClearFormats
    Columns("O:Q").UseStandardWidth = True
End Sub

Sub ClearHardChallenge()
    Call ClearModerateChallenge

    For Each ws In Worksheets
        ws.Columns("O:Q").ClearContents
        ws.Columns("O:Q").ClearFormats
        ws.Columns("O:Q").UseStandardWidth = True
    Next ws
End Sub
