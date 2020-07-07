Attribute VB_Name = "Module1"
Sub TestStockData()

' Define variables
Dim Ticker As String
Dim Year_Change As Double
Dim Percent_Change As Double
Dim Stock_Vol As Variant
Dim Open_Price As Double
Dim ws As Worksheet
Dim Sum_Table_Row As Long

' Worksheet Loop
For Each ws In Worksheets
Sum_Table_Row = 1

' column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

' find last row
Dim lastrow As Long
Dim i As Long
Dim j As Integer
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
Open_Price = ws.Cells(2, 3).Value
Stock_Vol = 0
    ' loop on current worksheet to lastrow
    For i = 2 To lastrow
    Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

    'Ticker symbol output
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Sum_Table_Row = Sum_Table_Row + 1
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(Sum_Table_Row, "I").Value = Ticker
    
        'Calculate change in Price
        close_price = ws.Cells(i, 6).Value
        Year_Change = close_price - Open_Price
    
        ' Addressing zeros
        If Open_Price <> 0 Then
        Percent_Change = (Year_Change / Open_Price) * 100
        Else
        Percent_Change = 0
        End If
        
        ' Fill in table with values
        ws.Range("I" & Sum_Table_Row).Value = Ticker
        ws.Range("J" & Sum_Table_Row).Value = Year_Change
        ws.Range("K" & Sum_Table_Row).Value = Percent_Change
        ws.Range("L" & Sum_Table_Row).Value = Stock_Vol
        
        ' Conditional Formatting
        
        If Year_Change >= 0 Then
        'green
        ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 4
        
        Else
        'red
        ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 3
        
        End If
        
        
     ' set new variables for prices and percent changes
        Stock_Vol = 0
        
        Open_Price = ws.Cells(i + 1, 3)
        
        
       End If
    
        
        
        
    Next i
    
    Next ws

End Sub
