VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AllYearsStockData()
    
    For Each ws In Worksheets
    
        Dim Worksheetname As String
        Dim Ticker_Name As String
        Dim Total_Stock_Volume As Double
        Dim Yearly_Change As Double
        Dim i As Long
        Dim j As Long
        Dim LastRowA As Long
        Dim LastRowB As Long
        Dim Percent_Change As Double
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        
        Worksheetname = ws.Name
        
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    Ticker_Name = 2
    
    j = 2
    
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRowA
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ws.Cells(Ticker_Name, 9).Value = ws.Cells(i, 1).Value
        
        ws.Cells(Ticker_Name, 10).NumberFormat = "#,##0.00"
        
        ws.Cells(Ticker_Name, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
            If ws.Cells(Ticker_Name, 10).Value < 0 Then
            
            ws.Cells(Ticker_Name, 10).Interior.ColorIndex = 3
            
            Else
            
            ws.Cells(Ticker_Name, 10).Interior.ColorIndex = 4
            
            End If
            
            If ws.Cells(j, 3).Value <> 0 Then
            
            Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            
            ws.Cells(Ticker_Name, 11).Value = Format(Percent_Change, "Percent")
            
            Else
            
            ws.Cells(Ticker_Name, 11).Value = Format(0, "Percent")
            
            End If
            
            ws.Cells(Ticker_Name, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            Ticker_Name = Ticker_Name + 1
            
            j = i + 1
            
            End If
            
        Next i
        
    LastRowB = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Greatest_Increase = ws.Cells(2, 11).Value
    Greatest_Decrease = ws.Cells(2, 11).Value
    Greatest_Volume = ws.Cells(2, 12).Value
            
    For i = 2 To LastRowB
    
        If ws.Cells(i, 11).Value > Greatest_Increase Then
        
        Greatest_Increase = ws.Cells(i, 11).Value
        
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        
        Else
        
        Greatest_Increase = Greatest_Increase
        
        End If
        
        If ws.Cells(i, 11).Value < Greatest_Decrease Then
        
        Greatest_Decrease = ws.Cells(i, 11).Value
        
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        
        Else
        
        Greatest_Decrease = Greatest_Decrease
        
        End If
    
        If ws.Cells(i, 12).Value > Greatest_Volume Then
        
        Greatest_Volume = ws.Cells(i, 12).Value
        
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        
        ws.Cells(4, 16).Value = WorksheetFunction.Max(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        Else
        
        Greatest_Volume = Greatest_Volume
        
        End If
            
      ws.Cells(2, 16).Value = Format(Greatest_Increase, "Percent")
      
      ws.Cells(3, 16).Value = Format(Greatest_Decrease, "Percent")
    
        Next i
    
    Next ws
    
End Sub


