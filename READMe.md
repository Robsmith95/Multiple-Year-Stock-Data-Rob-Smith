Sub Stock_Analysis():
    
    'Loop entire WS
    For Each ws In Worksheets
'---------------------------------------------------------------------------------------------
'Headers Decalred
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly_Change_in_Price"
    Cells(1, 12).Value = "Total_Stock_Volume"
    Cells(1, 11).Value = "Yearly_percentage"
    Cells(1, 3).Value = "Open_Price"
    Cells(1, 6).Value = "Close_Price"
              
'Values Declared
    Dim Ticker As String
    Dim Total_Vol As Double
        Total_Vol = 0
    Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Change_in_Price As Double
    Dim Increase_Row As Long
        Increase_Row = 2
    Dim Percent_Change_in_Price As Double
    
'---------------------------------------------------------------------------------------------
    
'Find the Final Row
    Final_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
' Loop through the rows
    For i = 2 To Final_Row
            Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_Vol
               Total_Vol = 0
    
'---------------------------------------------------------------------------------------------
'Yearly Price change
        Open_Price = ws.Range("C" & Increase_Row)
        Close_Price = ws.Cells(i, 6).Value
        Change_in_Price = Close_Price - Open_Price
                ws.Range("J" & Summary_Table_Row).Value = Change_in_Price
              
                
    If Open_Price = 0 Then
                Percent_Change_in_Price = 0
        Else
    Open_Price = ws.Range("C" & Increase_Row)
            Percent_Change_in_Price = Change_in_Price / Open_Price
                
       End If
                
'---------------------------------------------------------------------------------------------
'Color Formatting
    
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change_in_Price
                
                
        If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
        Summary_Table_Row = Summary_Table_Row + 1
                Increase_Row = i + 1
                
                End If
            
            Next i
    
    Next ws

End Sub


