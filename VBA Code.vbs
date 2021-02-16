
Sub TickerTable()
  

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

      ' Set an initial variable for holding the ticker code
      Dim Ticker As String
    
      ' Set an initial variable for holding the total per ticker brand
      Dim Ticker_Open As Double
      Dim Ticker_Close As Double
      Dim Ticker_Vol As Variant
      Dim NextTickerStart As Double
      Dim YearlyChange As Double
      Dim PercentChange As Double
      
      
      Ticker_Open = 0
      Ticker_Close = 0
      Ticker_Vol = 0

    
        'Set header rows for table in each worksheet'
        ws.Range("J1").Value = "<ticker>"
        ws.Range("K1").Value = "<Yearly Change>"
        ws.Range("L1").Value = "<Percentage Change>"
        ws.Range("M1").Value = "<Total Stock Volume>"
        ws.Range("J1:T1").Font.Bold = True
        'ws.Range("N1").Value = "<Open Price>" -- DO NOT PRINT
        'ws.Range("O1").Value = "<Close Price>" -- DO NOT PRINT
        
        'Default to each Ticker open value'
        Ticker_Open = ws.Cells(2, 3).Value
 
        'Find total Rows in worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
      ' Loop through all tickers for  volume count
      For i = 2 To LastRow
               
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          ' Set the Ticker name
          Ticker = ws.Cells(i, 1).Value
    
          ' Add to the Ticker Details
          Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
          Ticker_Close = Ticker_Close + ws.Cells(i, 6).Value
              
          ' Print the Ticker in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Ticker
          
          ' Print the Ticker Volume & Open & Close to the Summary Table
          ws.Range("M" & Summary_Table_Row).Value = Ticker_Vol
          'ws.Range("M" & Summary_Table_Row).Value = Ticker_Open -- DO NOT PRINT
          'ws.Range("N" & Summary_Table_Row).Value = Ticker_Close -- DO NOT PRINT
          
          ' Define the Yearly Change
          YearlyChange = Ticker_Close - Ticker_Open
    
            'Calculate the Yearly Change
            ws.Range("K" & Summary_Table_Row).Value = Ticker_Close - Ticker_Open
            '$ formatting
            ws.Range("K" & Summary_Table_Row).Style = "Currency"
                
                'Color formatting
                If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
                ElseIf ws.Range("K" & Summary_Table_Row).Value = 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbYellow
                Else: ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
            
                End If
              
                  'Calculate the Percent Change as 0 to avoid division error
                  If Ticker_Open = 0 Then
                      ws.Range("L" & Summary_Table_Row).Value = 0
                      
                  'If not error then
                  
                  Else:  ws.Range("L" & Summary_Table_Row).Value = (Ticker_Close - Ticker_Open) / Ticker_Open
                                    
                  End If
                
            '% Formatting
            ws.Range("L" & Summary_Table_Row).Style = "Percent"
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                      
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Ticker Total
          Ticker_Vol = 0
          Ticker_Close = 0
          NextTickerStart = i + 1
          Ticker_Open = ws.Cells(NextTickerStart, 3).Value

        ' If the cell immediately following a row is the same brand...
    
       
     Else
        ' Add to the Brand Total
        Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
     
     End If

    Next i
    
    ws.Range("A:S").EntireColumn.AutoFit
       
Next ws

End Sub



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

