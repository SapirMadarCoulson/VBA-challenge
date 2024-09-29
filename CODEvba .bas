Attribute VB_Name = "Module1"

Sub CombinedQuarterlyAnalysis()
    Dim i As Long
    Dim LastRow As Long
    Dim Summary_Table_Row As Long
    Dim Ticker As String
    Dim Stock_Total As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim PriceDifference As Double
    Dim PercentageChange As Double
    
    Stock_Total = 0
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_Row = 2

    For i = 2 To LastRow
        Ticker = Cells(i, 1).Value


        If i = 2 Or Cells(i - 1, 1).Value <> Ticker Then
            OpeningPrice = Cells(i, 3).Value
        End If

        Stock_Total = Stock_Total + Cells(i, 7).Value

        If i = LastRow Or Cells(i + 1, 1).Value <> Ticker Then
            ClosingPrice = Cells(i, 6).Value
            
            PriceDifference = ClosingPrice - OpeningPrice
            Cells(Summary_Table_Row, 10).Value = PriceDifference
            

            If OpeningPrice <> 0 Then
                PercentageChange = (PriceDifference / OpeningPrice) * 1#
            Else
                PercentageChange = 0
            End If
            
            Range("I1").Value = "Ticker "
            Range("J1").Value = "Quarterly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Volume Stock"
            
            
            Cells(Summary_Table_Row, 11).Value = PercentageChange
            precentagechange = Application.WorksheetFunction.Round(precentagechange, 0)
            Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
            Cells(Summary_Table_Row, 12).Value = Stock_Total
            Cells(Summary_Table_Row, 9).Value = Ticker
            
            Summary_Table_Row = Summary_Table_Row + 1
            Stock_Total = 0
        End If
        
    Next i
End Sub

Sub CheckAndColorColumnJ()
    Dim LastRow As Long
    Dim i As Long
    Dim ColorCell As Range
    

    LastRow = Cells(Rows.Count, "J").End(xlUp).Row

    For i = 2 To LastRow
        Set cell = Cells(i, "J")
        

        If cell.Value > 0 Then
            cell.Interior.Color = RGB(0, 255, 0)
        ElseIf cell.Value < 0 Then
            cell.Interior.Color = RGB(255, 0, 0)
        Else
            cell.Interior.Color = RGB(255, 255, 255)
        End If
    Next i
End Sub

        

Sub FindingMaxAndMinValues()

    Dim High As Range
    Dim Low As Range
    Dim Max_Total As Range
    Dim Ticker As String
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim HighValue As Double
    Dim LastRow As Long
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim HighTicker As String
    Dim i As Long
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    

    Set High = Range("K2:K" & LastRow)
    Set Low = Range("K2:K" & LastRow)
    Set Max_Total = Range("L2:L" & LastRow)
    
    MaxValue = Application.WorksheetFunction.Max(High)
    MinValue = Application.WorksheetFunction.Min(Low)
    HighValue = Application.WorksheetFunction.Max(Max_Total)
    
    For i = 2 To LastRow
        If Cells(i, 11).Value = MaxValue Then
            MaxTicker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value = MinValue Then
            MinTicker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value = HighValue Then
            HighTicker = Cells(i, 9).Value
        End If
    Next i
    
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    

    Range("Q2").Value = MaxValue
    Range("Q3").Value = MinValue
    Range("Q4").Value = HighValue
    
    Range("P2").Value = MaxTicker
    Range("P3").Value = MinTicker
    Range("P4").Value = HighTicker
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Value"

End Sub
