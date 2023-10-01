Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
 
      
        Dim Ticker As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        
         
        Dim i As Long
        Dim j As Long
        
        'Worksheet Name
        WorksheetName = ws.Name
        
        'Column headers
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Setting Ticker to first row
        Ticker = 2
        
        'Set start row to 2
        j = 2
        
       'Last cell in column A
        
       LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        
            'Loop through all rows
            For i = 2 To LastRowA
            
               
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                
              'Yearly Change
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

                    
               'Percent Change
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    End If
                    
                'Total Volume
                
                ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
         
                Ticker = Ticker + 1
                
                j = i + 1
               
                End If
            
            Next i
            
        'Last non-blank cell in column I
        
        
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
        

        GreatestVolume = ws.Cells(2, 12).Value
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        
            'Loop
            For i = 2 To LastRowI
            
            'Greatest Volume
                
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
                End If
                
             'Greatest Increase
                
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                End If
                
              'Greatest Decrease
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                End If
                
  
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrase, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next i
            
             'Conditional formating
                    
                    If ws.Cells(Ticker, 10).Value < 0 Then
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                    End If
            
        
            
    Next ws
        
End Sub
