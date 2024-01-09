VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_tester()


Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook

For Each ws In wb.Worksheets
ws.Activate

Range("I1").Value = "Ticker"
Range("J1").Value = "YearlyChange"
Range("K1").Value = "PercentChange"
Range("L1").Value = "TotalStockVolume"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Dim Ticker As String
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Row As Integer
Row = 2
Dim First_Table_Last_Row As LongLong
First_Table_Last_Row = 0
Dim Summary_Last_Row As Integer
Summary_Last_Row = 0
Dim Max As Double
Max = 0
Dim Max_Ticker As String
Dim Min As Double
Min = 0
Dim Min_Ticker As String
Dim Max_Stock_Volume As LongLong
Max_Stock_Volume = 0
Dim Max_Stock_Ticker As String
Dim Ticker_Change_Close As Double
Dim Ticker_Change_Open As Double
Dim Yearly_Change As Double
Dim TotalStockVolume As LongLong
TotalStockVolume = 0

First_Table_Last_Row = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To First_Table_Last_Row
    
    ' Assign Ticker Open value
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        Ticker_Change_Open = Cells(i, 3).Value

    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ' Assign Ticker Value
        Ticker = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        ' Assign Ticker Close value
        Ticker_Change_Close = Cells(i, 6).Value
        Range("J" & Summary_Table_Row).Value = Ticker_Change_Close - Ticker_Change_Open
        Yearly_Change = Range("J" & Summary_Table_Row).Value
        Range("J" & Summary_Table_Row).Style = "Currency"
            If Yearly_Change >= 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf Yearly_Change < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        Range("K" & Summary_Table_Row).Value = FormatPercent(Yearly_Change / Ticker_Change_Open, 2)
            If Range("K" & Summary_Table_Row).Value >= 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf Range("K" & Summary_Table_Row).Value < 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        ' Assign Total Stock Volume
        TotalStockVolume = Cells(i - 1, 7).Value + Cells(i, 7).Value + TotalStockVolume
        Range("L" & Summary_Table_Row) = TotalStockVolume
        TotalStockVolume = 0
        Summary_Table_Row = Summary_Table_Row + 1
        
    ElseIf Cells(i, 1).Value = Cells(i - 1, 1).Value Then
        TotalStockVolume = Cells(i - 1, 7).Value + TotalStockVolume
        
    End If
    
Next i

Summary_Last_Row = Cells(Rows.Count, "J").End(xlUp).Row

        For Summary_Table_Row = 2 To Summary_Last_Row
      
        ' Greatest % Increase
        If Cells(Summary_Table_Row, 11).Value > Cells(Summary_Table_Row + 1, 11).Value Then
            If Cells(Summary_Table_Row, 11).Value > Max Then
            Max_Ticker = Cells(Summary_Table_Row, 9).Value
            Range("P" & 2).Value = Max_Ticker
            Max = Cells(Summary_Table_Row, 11).Value
            Range("Q" & 2).Value = FormatPercent(Max, 2)
            End If

        ElseIf Cells(Summary_Table_Row, 11).Value <= Cells(Summary_Table_Row + 1, 11).Value Then
            If Cells(Summary_Table_Row + 1, 11).Value > Max Then
            Max = Cells(Summary_Table_Row + 1, 11).Value
            Max_Ticker = Cells(Summary_Table_Row + 1, 9).Value
            Range("P" & 2).Value = Max_Ticker
            Range("Q" & 2).Value = FormatPercent(Max, 2)
            End If

        End If

        ' Greatest % Decrease
        If Cells(Summary_Table_Row, 11).Value > Cells(Summary_Table_Row + 1, 11).Value Then
            If Cells(Summary_Table_Row + 1, 11).Value < Min Then
            Min_Ticker = Cells(Summary_Table_Row + 1, 9).Value
            Range("P" & 3).Value = Min_Ticker
            Min = Cells(Summary_Table_Row + 1, 11).Value
            Range("Q" & 3).Value = FormatPercent(Min, 2)
            End If
        ElseIf Cells(Summary_Table_Row, 11).Value <= Cells(Summary_Table_Row + 1, 11).Value Then
            If Cells(Summary_Table_Row, 11).Value < Min Then
            Min = Cells(Summary_Table_Row, 11).Value
            Min_Ticker = Cells(Summary_Table_Row, 9).Value
            Range("P" & 3).Value = Min_Ticker
            Range("Q" & 3).Value = FormatPercent(Min, 2)
            End If
        End If
        

       ' Greatest Total Volume
        If Cells(Summary_Table_Row, 12).Value > Cells((Summary_Table_Row + 1), 12).Value Then
            If Cells(Summary_Table_Row, 12).Value > Max_Stock_Volume Then
            Max_Stock_Volume = Cells(Summary_Table_Row, 12).Value
            Range("Q" & 4).Value = Max_Stock_Volume
            Max_Stock_Ticker = Cells(Summary_Table_Row, 9).Value
            Range("P" & 4).Value = Max_Stock_Ticker
            End If
            
        ElseIf Cells(Summary_Table_Row, 12).Value <= Cells((Summary_Table_Row + 1), 12).Value Then
            If Cells(Summary_Table_Row + 1, 12).Value > Max_Stock_Volume Then
            Max_Stock_Volume = Cells(Summary_Table_Row + 1, 12).Value
            Range("Q" & 4).Value = Max_Stock_Volume
            Max_Stock_Ticker = Cells(Summary_Table_Row + 1, 9).Value
            Range("P" & 4).Value = Max_Stock_Ticker
            End If

        End If

        Next Summary_Table_Row

Next ws

End Sub

