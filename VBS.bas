Attribute VB_Name = "Module1"
Sub Multiple_year_stock():

For Each ws In Worksheets

    Dim percent_change_Max As Double
    Dim Summary_Table_Row As Integer

    Summary_Table_Row = 2
    greatest_percentage_increase = 0
    greatest_percentage_decrease = 0
    
    
        Dim ticker As String
        Dim yearly_change As Double
        Dim Percent_change As Double
        Dim Total_Stock_Volumn As Double
        Dim Start_value As Double
        Dim Greatest_Total_Stock_Volumn As Double
        
        Greatest_Total_Stock_Volumn = 0
               
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        
        Total_Stock_Volumn = 0
        Start_value = ws.Cells(2, 6).Value
        
        For i = 2 To lastRow
            
            Total_Stock_Volumn = Total_Stock_Volumn + ws.Cells(i, 7).Value
            
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               ticker = ws.Cells(i, 1).Value
               
               
               yearly_change = ws.Cells(i, 6).Value - Start_value
               
               If Start_value = "0" Then
               Percent_change = 0
               Else
               Percent_change = 100 * yearly_change / Start_value
               End If
               
               If Percent_change > greatest_percentage_increase Then
               greatest_percentage_increase = Percent_change
               
               ws.Range("P2").Value = ticker
               End If
               
               If Percent_change < greatest_percentage_decrease Then
               greatest_percentage_decrease = Percent_change
               
               ws.Range("P3").Value = ticker
               End If
               
               If Total_Stock_Volumn > Greatest_Total_Stock_Volumn Then
               Greatest_Total_Stock_Volumn = Total_Stock_Volumn
               
               ws.Range("P4").Value = ticker
               
               End If
               
               Start_value = ws.Cells(i + 1, 6).Value
               
               ws.Range("I" & Summary_Table_Row).Value = ticker
               
               ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volumn
               
               ws.Range("J" & Summary_Table_Row).Value = yearly_change
               
               ws.Range("K" & Summary_Table_Row).Value = CStr(Round(Percent_change, 2)) & "%"
               
               Total_Stock_Volumn = 0
               
               If yearly_change > 0 Then
               ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
               Else
               ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
               End If
               
                              
               Summary_Table_Row = Summary_Table_Row + 1
               
                             
        
           End If
         
        Next i
               
           ws.Range("Q2").Value = CStr(Round(greatest_percentage_increase, 2)) & "%"
           ws.Range("Q3").Value = CStr(Round(greatest_percentage_decrease, 2)) & "%"
           ws.Range("Q4").Value = Greatest_Total_Stock_Volumn
           
           
           
    Next ws
         
    


End Sub
