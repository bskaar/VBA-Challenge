Attribute VB_Name = "Module1"

Sub StockPrice_2018()

'For Each ws In Worksheets



'Dim WorksheetName As String
Dim Rowcount As LongLong
Dim Tickercount As LongLong


Rowcount = Cells(Rows.Count, 1).End(xlUp).Row
Tickercount = 2



'WorksheetName = Name

'Stock price loop

For i = 2 To Rowcount

    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    Cells(Tickercount, 9).Value = Cells(i, 1).Value  'Returns stock ticker
    Cells(Tickercount, 10).Value = Cells(i, 6).Value - Cells(i - 250, 3).Value 'Returns yearly stock price value change
    Cells(Tickercount, 11).Value = Format((Cells(i, 6).Value / Cells(i - 250, 3).Value) - 1, "Percent") 'Returns Percent Change
    Cells(Tickercount, 12).Value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(i - 250, 7))) 'Sums Total Volume
    

'Conditional formatting

    CondColor = Cells(Tickercount, 10).Value
    Select Case CondColor
    Case Is > 0
    Cells(Tickercount, 10).Interior.ColorIndex = 4
    Case Is < 0
    Cells(Tickercount, 10).Interior.ColorIndex = 3
    Case Else
    Cells(Tickercount, 10).Interior.ColorIndex = 0
    End Select
    Tickercount = Tickercount + 1
    
End If

Next i

'End of Stock price loop

'Column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
'Summary Data headings
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
'Summary Data Greatest % increase/Year and Volume: Values
    Range("P2") = Format(Application.WorksheetFunction.Max(Range("K:K")), "Percent")
    Range("P3") = Format(Application.WorksheetFunction.Min(Range("K:K")), "Percent")
    Range("P4") = Application.WorksheetFunction.Max(Range("L:L"))
    
'Summary Data Greatest % increase/Year and Volume: Ticker
    MaxIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
    Range("O2") = Cells(MaxIndex, 9)
    MinIndex = WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
    Range("O3") = Cells(MinIndex, 9)
    VolIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
    Range("O4") = Cells(VolIndex, 9)
    Cells(Tickercount, 11).Value = FormatPercent(Cells(Tickercount, 11))

'Next


End Sub


Sub StockPrice_2019()

'For Each ws In Worksheets



'Dim WorksheetName As String
Dim Rowcount As LongLong
Dim Tickercount As LongLong


Rowcount = Cells(Rows.Count, 1).End(xlUp).Row
Tickercount = 2



'WorksheetName = Name

'Stock price loop

For i = 2 To Rowcount

    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    Cells(Tickercount, 9).Value = Cells(i, 1).Value  'Returns stock ticker
    Cells(Tickercount, 10).Value = Cells(i, 6).Value - Cells(i - 251, 3).Value 'Returns yearly stock price value change
    Cells(Tickercount, 11).Value = Format((Cells(i, 6).Value / Cells(i - 251, 3).Value) - 1, "Percent") 'Returns Percent Change
    Cells(Tickercount, 12).Value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(i - 251, 7))) 'Sums Total Volume
    

'Conditional formatting

    CondColor = Cells(Tickercount, 10).Value
    Select Case CondColor
    Case Is > 0
    Cells(Tickercount, 10).Interior.ColorIndex = 4
    Case Is < 0
    Cells(Tickercount, 10).Interior.ColorIndex = 3
    Case Else
    Cells(Tickercount, 10).Interior.ColorIndex = 0
    End Select
    Tickercount = Tickercount + 1
    
End If

Next i

'End of Stock price loop

'Column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
'Summary Data headings
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
'Summary Data Greatest % increase/Year and Volume: Values
    Range("P2") = Format(Application.WorksheetFunction.Max(Range("K:K")), "Percent")
    Range("P3") = Format(Application.WorksheetFunction.Min(Range("K:K")), "Percent")
    Range("P4") = Application.WorksheetFunction.Max(Range("L:L"))
    
'Summary Data Greatest % increase/Year and Volume: Ticker
    MaxIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
    Range("O2") = Cells(MaxIndex, 9)
    MinIndex = WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
    Range("O3") = Cells(MinIndex, 9)
    VolIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
    Range("O4") = Cells(VolIndex, 9)
    Cells(Tickercount, 11).Value = FormatPercent(Cells(Tickercount, 11))

'Next


End Sub




Sub StockPrice_2020()

'For Each ws In Worksheets



'Dim WorksheetName As String
Dim Rowcount As LongLong
Dim Tickercount As LongLong


Rowcount = Cells(Rows.Count, 1).End(xlUp).Row
Tickercount = 2



'WorksheetName = Name

'Stock price loop

For i = 2 To Rowcount

    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    Cells(Tickercount, 9).Value = Cells(i, 1).Value  'Returns stock ticker
    Cells(Tickercount, 10).Value = Cells(i, 6).Value - Cells(i - 252, 3).Value 'Returns yearly stock price value change
    Cells(Tickercount, 11).Value = Format((Cells(i, 6).Value / Cells(i - 252, 3).Value) - 1, "Percent") 'Returns Percent Change
    Cells(Tickercount, 12).Value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(i - 252, 7))) 'Sums Total Volume
    

'Conditional formatting

    CondColor = Cells(Tickercount, 10).Value
    Select Case CondColor
    Case Is > 0
    Cells(Tickercount, 10).Interior.ColorIndex = 4
    Case Is < 0
    Cells(Tickercount, 10).Interior.ColorIndex = 3
    Case Else
    Cells(Tickercount, 10).Interior.ColorIndex = 0
    End Select
    Tickercount = Tickercount + 1
    
End If

Next i

'End of Stock price loop

'Column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
'Summary Data headings
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
'Summary Data Greatest % increase/Year and Volume: Values
    Range("P2") = Format(Application.WorksheetFunction.Max(Range("K:K")), "Percent")
    Range("P3") = Format(Application.WorksheetFunction.Min(Range("K:K")), "Percent")
    Range("P4") = Application.WorksheetFunction.Max(Range("L:L"))
    
'Summary Data Greatest % increase/Year and Volume: Ticker
    MaxIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
    Range("O2") = Cells(MaxIndex, 9)
    MinIndex = WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
    Range("O3") = Cells(MinIndex, 9)
    VolIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
    Range("O4") = Cells(VolIndex, 9)
    Cells(Tickercount, 11).Value = FormatPercent(Cells(Tickercount, 11))

'Next


End Sub

