Attribute VB_Name = "Module"


Sub Ticker()
Dim ws As Worksheet
For Each ws In Worksheets
  'Set variables for holding values
  
  Dim Open_Price As Single
  Open_Price = 0
  
  Dim Close_Price As Single
  Close_Price = 0
  
  Dim Yearly_Change As Single
  Yearly_Change = 0
  
  Dim Delta_Price As Double
  Delta_Price = 0
  
  Dim Percent_Change As Single
  Percent_Change = 0
  
  Dim Ticker As String
  Ticker = " "
  
  Dim Max_Incease As Double
  Max_Increase = 0
  Dim Max_Decrease As Double
  Max_Decrease = 0
  Dim Max_Volume As Double
  Max_Volume = 0

  Dim Summary_Table_Row As Integer
  
  Summary_Table_Row = 2
    
  'Set var for total stock, double is twice Long
  
  Dim Total_Stock As Double
  
  Total_Stock = 0
  
  Dim Lastrow As Long
  Dim i As Long
  
  'Titles for summary rows
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock"
  ws.Range("O1").Value = "Greatest % Increase"
  ws.Range("O2").Value = "Greatest % Decrease"
  ws.Range("O3").Value = "Greatest Total Volume"
  
  
  'Var for last row
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  Open_Price = ws.Cells(2, 3).Value
  
  For i = 2 To Lastrow
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Fill ticker name
        Ticker = ws.Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        
        
        
        'Set Open/Close Price vars
        Close_Price = ws.Cells(i, 6).Value
       
        Delta_Price = Close_Price - Open_Price
        
        If Open_Price <> 0 Then
            Percent_Change = (Delta_Price / Open_Price) * 100
        End If
        
        
        
        'Add total stock for ticker
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                
        'Add values to Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock
        'Percent Change
        ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
        
        'Fill Yearly Change
        ws.Range("J" & Summary_Table_Row).Value = Delta_Price
        If (Delta_Price > 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
         ElseIf (Delta_Price < 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
        End If

            

        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Stock = 0
        Delta_Price = 0
        Percent_Change = 0
        Open_Price = ws.Cells(i + 1, 3).Value
        Else
  
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
    End If
        
        
    
    
    
 Next i
    'Compute Greatest Percent Increase, Decrease and Total Stock Volume
    Max_Increase = WorksheetFunction.Max(Range("K:K"))
    ws.Range("P1").Value = (CStr(Max_Increase * 100) & "%")

    Max_Decrease = WorksheetFunction.Min(Range("K:K"))
    ws.Range("P2").Value = (CStr(Max_Decrease * 100) & "%")

    Max_Volume = WorksheetFunction.Max(Range("L:L"))
    ws.Range("P3").Value = Max_Volume
Next ws

End Sub


