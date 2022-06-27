Attribute VB_Name = "Module1"
Sub Stocks():
'Variable Declaration
Dim Current As Worksheet
Dim i As Long
Dim Lastrow As Long
Dim Ticker As String
Dim Summarycounter As Long
Dim Firstticker As Long
Dim Topval As String
Dim Bottomval As String
Dim Sigma As Range
Dim Totalvol As Double
Dim Pctchange As Range
Dim Toppct As String
Dim Bottomloc As Integer
Dim Bottompct As String
Dim Condition1 As FormatCondition, Condition2 As FormatCondition
Dim Greatpct As Double
Dim Leastpct As Double
Dim Greatvol As Double
Dim Seeker As String
Dim match As Range
Dim Finder As Double
'Iteratating through each worksheet
    For Each Current In Worksheets
        'Setting/resetting counters and creating headers for desired information table
        Summarycounter = 1
        Lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
        Current.Cells(1, 10).Value = "Ticker"
        Current.Cells(1, 11).Value = "Yearly Change"
        Current.Cells(1, 12).Value = "Percent Change"
        Current.Cells(1, 13).Value = "Total Volume"
        Current.Cells(2, 15).Value = "Greatest % Increase"
        Current.Cells(3, 15).Value = "Greatest % Decrease"
        Current.Cells(4, 15).Value = "Greatest Total Volume"
        Current.Cells(1, 16).Value = "Ticker"
        Current.Cells(1, 17).Value = "Value"
        'Iterating through each row of the raw data
        For i = 2 To Lastrow
            Ticker = Current.Cells(i, 1).Value
            'Conditional to determine if there is a change in the ticker
            If Ticker <> Current.Cells((i - 1), 1).Value Then
                Firstticker = i
                Summarycounter = Summarycounter + 1
                Current.Cells(Summarycounter, 10).Value = Ticker
            'Finding the last value of particular ticker and collecting all of the data
            ElseIf Ticker <> Cells((i + 1), 1).Value Then
                Topval = Current.Cells(Firstticker, 7).Address
                Bottomval = Current.Cells(i, 7).Address
                Set Sigma = Current.Range(Topval, Bottomval)
                Current.Cells(Summarycounter, 11) = Current.Cells(i, 6).Value - Current.Cells(Firstticker, 6).Value
                Current.Cells(Summarycounter, 12) = (Current.Cells(Summarycounter, 11) / Current.Cells(Firstticker, 6).Value) * 100
                Totalvol = WorksheetFunction.Sum(Sigma)
                Current.Cells(Summarycounter, 13).Value = Totalvol
            End If
        Next i
        
        'Bonus Searches
        Greatpct = WorksheetFunction.Max(Current.Range(Current.Cells(2, 11).Address, Current.Cells(Summarycounter, 11).Address))
        Current.Cells(2, 17).Value = Greatpct
        Finder = Current.Cells(2, 17).Value
        Set match = Current.Range("K:K").Find(Finder)
        Seeker = match.Offset(, -1).Value
        Current.Cells(2, 16).Value = Seeker
        Leastpct = WorksheetFunction.Min(Current.Range(Current.Cells(2, 11).Address, Current.Cells(Summarycounter, 11).Address))
        Current.Cells(3, 17).Value = Leastpct
        Finder = Current.Cells(3, 17).Value
        Set match = Current.Range("K:K").Find(Finder)
        Seeker = match.Offset(, -1).Value
        Current.Cells(3, 16).Value = Seeker
        Greatvol = WorksheetFunction.Max(Current.Range(Current.Cells(2, 13).Address, Current.Cells(Summarycounter, 13).Address))
        Current.Cells(4, 17).Value = Greatvol
        Finder = Current.Cells(4, 17).Value
        Set match = Current.Range("M:M").Find(Finder)
        Seeker = match.Offset(, -3).Value
        Current.Cells(4, 16).Value = Seeker
        
        
        
        'Conditional Formatting Code
        Toppct = Current.Cells(2, 11).Address
        Bottomloc = Current.Cells(Rows.Count, 11).End(xlUp).Row
        Bottompct = Current.Cells(Bottomloc, 11).Address
        Set Pctchange = Current.Range(Toppct, Bottompct)
        Pctchange.FormatConditions.Delete
        Set Condition1 = Pctchange.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set Condition2 = Pctchange.FormatConditions.Add(xlCellValue, xlLess, "0")
            With Condition1
                .Interior.Color = vbGreen
            End With
            With Condition2
                .Interior.Color = vbRed
            End With
    Next
End Sub

