Sub Stock_Ticker()

For Each ws In Worksheets                                                   'To apply to each worksheet

'Defining variables
Dim Ticker As String
Dim Greatest_Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Min_Percent As Double
Dim Max_Percent As Double
Dim First_Value_Row As Integer

'Initializing variables
First_Value_Row = 2
Total_Stock_Volume = 0
Greatest_Stock_Volume = 0
Greatest_Volume_Ticker = " "
Greatest_Percent_Increase = 0
Greatest_Increase_Ticker = " "
Greatest_Percent_Decrease = 0
Greatest_Decrease_Ticker = " "
Opening_Price = ws.Cells(2, 3).Value

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row                             'Calculates the last data row in each sheet

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    Else
    
        Ticker = ws.Cells(i, 1).Value                                       'Outputs the Ticker name
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value      'Outputs total stock volume - including the changeover row's value
    
    If (Total_Stock_Volume > Greatest_Stock_Volume) Then            'Challenge question - greatest total volume - compares total stock volume with greatest stock volume & captures the greatest stock volume within each iteration
        Greatest_Stock_Volume = Total_Stock_Volume
        Greatest_Volume_Ticker = ws.Cells(i, 1).Value
    End If
        
        
        Closing_Price = ws.Cells(i, 6).Value                                'Captures the closing price value
        Yearly_Change = Closing_Price - Opening_Price                       'Calculates the yearly change
        
    If (Opening_Price = 0) Then                                     'Check for opening price = 0. In the formula used, Opening price cannot be 0, as it generates not divisible by 0 error
        Percent_Change = 0
    End If

    If (Opening_Price <> 0) And (Closing_Price <> 0) Then
        Percent_Change = (Closing_Price - Opening_Price) / Opening_Price * 100
    End If
    
    If (Percent_Change > 0) And (Percent_Change > Greatest_Percent_Increase) Then
        Greatest_Percent_Increase = Percent_Change             'Challenge question - greatest % increase
        Greatest_Increase_Ticker = ws.Cells(i, 1).Value
    ElseIf (Percent_Change < 0) And (Percent_Change < Greatest_Percent_Decrease) Then
        Greatest_Percent_Decrease = Percent_Change             'Challenge question - greatest % decrease
        Greatest_Decrease_Ticker = ws.Cells(i, 1).Value
    
    End If
        
        ws.Range("J" & First_Value_Row).Value = Ticker
        ws.Range("K" & First_Value_Row).Value = Yearly_Change
        
    If (Yearly_Change < 0) Then                                              'Formatting cells green/red for +ve/-ve values
        ws.Range("K" & First_Value_Row).Interior.ColorIndex = 3
    Else
        ws.Range("K" & First_Value_Row).Interior.ColorIndex = 4
    End If
        ws.Range("L" & First_Value_Row).Value = (CStr(Percent_Change) & "%") 'Formatting cells to output value with % sign
        ws.Range("M" & First_Value_Row).Value = Total_Stock_Volume
        
    
        First_Value_Row = First_Value_Row + 1
        Total_Stock_Volume = 0
        Opening_Price = ws.Cells(i + 1, 3).Value
    

    End If
    
Next i
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("J1").Font.Bold = True
        
        ws.Range("K1").Value = "Yearly_Change"
        ws.Range("K1").Font.Bold = True

        ws.Range("L1").Value = "Percent_Change"
        ws.Range("L1").Font.Bold = True

        ws.Range("M1").Value = "Total_Stock_Volume"
        ws.Range("M1").Font.Bold = True

        ws.Range("P1").Value = "Ticker"
        ws.Range("P1").Font.Bold = True

        ws.Range("Q1").Value = "Value"
        ws.Range("Q1").Font.Bold = True
                
        ws.Range("O2").Value = "Greatest %Increase"
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("Q2").Value = (CStr(Greatest_Percent_Increase) & "%")
        
        
        ws.Range("O3").Value = "Greatest %Decrease"
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("Q3").Value = (CStr(Greatest_Percent_Decrease) & "%")
        
        
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P4").Value = Greatest_Volume_Ticker
        ws.Range("Q4").Value = Greatest_Stock_Volume
       
                
                    
Next ws

End Sub
