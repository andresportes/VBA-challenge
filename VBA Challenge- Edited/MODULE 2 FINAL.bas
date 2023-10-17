Attribute VB_Name = "Module1"


Sub MultipleSockData()


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = " Total Stock Volume"


Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Double
Dim Summary_Table_Row As Integer
Dim Open_Price As Double
Dim Close_Price As Double


Total_Stock = 0
Summary_Table_Row = 2

Open_Price = ws.Cells(2, 3).Value



       
'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop Through All Ticker Symbols
For i = 2 To LastRow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


'Setting the Ticker Symbol

Ticker = ws.Cells(i, 1).Value


'Total stock volume
Total_Stock = Total_Stock + ws.Cells(i, 7).Value

'Ticker Symbol

ws.Range("I" & Summary_Table_Row).Value = Ticker


ws.Range("L" & Summary_Table_Row).Value = Total_Stock

Close_Price = ws.Cells(i, 6).Value

'Change in price
Yearly_Change = (Close_Price - Open_Price)

Percent_Change = Yearly_Change / Open_Price

'Yearly Change

ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"


Summary_Table_Row = Summary_Table_Row + 1


Open_Price = ws.Cells(i + 1, 3).Value
Percent_Change = 0
Total_Stock = 0


Else

Total_Stock = Total_Stock + ws.Cells(i, 7).Value

End If

Next i

lastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To lastRow_Summary_Table

If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 10

Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i



ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Vaule"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row


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
 

 ws.Columns("I:Q").AutoFit
 
Next ws

End Sub
