
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