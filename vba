Sub Multi_Year_Stock_Data()

'Declare Variables'

Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Total_Volume As Double
Dim Summary_Table_Row As Integer
Dim Opening As Double
Dim Closing As Double

'Loop through worksheets'

For Each ws In Worksheets
    
    Summary_Table_Row = 2
    
    Total_Stock_Volume = 0

    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Opening = ws.Cells(2, 3)
    For i = 2 To LastRow
    
    'Set conditionals'
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Ticker = ws.Cells(i, 1)
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
        
            Closing = ws.Cells(i, 6)
           
            Yearly_Change = Closing - Opening
            
            
            If Opening = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Opening * 100
           End If
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = (Percent_Change & "%")
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
             
            Total_Stock_Volume = 0
       
            Opening = ws.Cells(i + 1, 3)
            
             ' Positive change in green, negative in red
            
             If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
            
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
                End If
            
                Summary_Table_Row = Summary_Table_Row + 1
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
           
        
        
       
        
        End If
        

        
    Next i
    
    'Summary names'
    
    ws.Range("I1") = "Ticker"
    ws.Range("P1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"

    
Next ws


End Sub
