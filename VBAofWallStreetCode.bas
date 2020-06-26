Attribute VB_Name = "Module1"


Sub VBA_of_Wall_Street():


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Dim LastRow As Long
Dim Ticker As String
Dim Summary_table_Index As Integer
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

Dim Ticker_Count As Integer
    Ticker_Count = 0
    Summary_table_Index = 2
    Total_Stock_Volume = 0
    
    


LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Open_Price = Cells(2, 3).Value
Ticker = Cells(2, 1).Value
For i = 2 To LastRow
    Close_Price = Cells(i, 6).Value
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
    
    If Ticker <> Cells(i + 1, 1).Value Then
    
    Yearly_Change = Close_Price - Open_Price
    If Open_Price <> 0 Then
    Percent_Change = Yearly_Change / Open_Price * 100
    Else: Percent_Change = 100
    End If
    
        Ticker_Count = Ticker_Count + 1
        Range("I" & Summary_table_Index).Value = Ticker
        Range("J" & Summary_table_Index).Value = Yearly_Change
        Range("K" & Summary_table_Index).Value = Percent_Change
            If Yearly_Change >= 0 Then
                Range("J" & Summary_table_Index).Interior.ColorIndex = (4)
                Else
                Range("J" & Summary_table_Index).Interior.ColorIndex = (3)
                End If
                
        Range("L" & Summary_table_Index).Value = Total_Stock_Volume
        Summary_table_Index = Summary_table_Index + 1
     Open_Price = Cells(i + 1, 3).Value
     Ticker = Cells(i + 1, 1)
     
    
    End If
Next i





Range("I1:L1").Columns.AutoFit



End Sub

