Sub ChallengeTable()

'================================================
'           CHALLENGE CODE
'================================================

For Each ws In Worksheets

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2:O4").Font.Bold = True
    ws.Range("N1:S1").Font.Bold = True

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double

    GreatestIncrease = ws.Cells(2, 12).Value
    GreatestDecrease = ws.Cells(2, 12).Value
    GreatestTotalVolume = ws.Cells(2, 13).Value

    'Count how many rows in Summary table
    Dim LastRowSummary As Integer
    LastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        '--GREATEST INCREASE--
        
        'Identify Greatest Increase'
        
        GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowSummary))
    
        'Print & Format Greatest Increase'
        
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
    
            'Find the ticker name with the Greatest Increase, then input ticker name in summary table 2.
            
            For i = 2 To LastRowSummary
                    If ws.Cells(i, 12).Value = GreatestIncrease Then
                        ws.Range("P2").Value = ws.Cells(i, 10).Value
                    End If
    
            Next i


        '--GREATEST DECREASE--
        
        'Identify Greatest Decrease'
        
        GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("L2:L" & LastRowSummary))
        
        'Print & Format Greatest Decrease'
        
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%"

            'Find the ticker name with the Greatest Decrease, then input ticker name in summary table 2.
            
            For i = 2 To LastRowSummary
                If ws.Cells(i, 12).Value = GreatestDecrease Then
                    ws.Range("P3").Value = ws.Cells(i, 10).Value
            
            End If

        Next i

        '--GREATEST TOTAL VOLUME--
        
        'Find out the last number in total volumn and input into summary table

        GreatestTotalVolume = Application.WorksheetFunction.Max(ws.Range("M2:M" & LastRowSummary))

        ws.Range("Q4").Value = GreatestTotalVolume

        'Find the ticker name with the greatest total volumn, then input ticker name in summary table 2.
    
            For i = 2 To LastRowSummary
                If ws.Cells(i, 13).Value = GreatestTotalVolume Then
                ws.Range("P4").Value = ws.Cells(i, 10).Value
            End If
    
        'Column Formatting
         ws.Range("O:S").EntireColumn.AutoFit
         
        Next i

Next ws


End Sub

