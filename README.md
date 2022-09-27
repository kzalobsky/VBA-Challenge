# VBA-Challenge

I set up code for each worksheet (I know this isn't the most efficient way to do this, but it was the best I could do -- lets just say I am happy to be moving on from VBA!)

Each worksheet has the same code:

Sub Worksheet3()

'Determining Column Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Determine Ticker Symbols
ActiveSheet.Range("A2:A753001").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I2"), Unique:=True


'Determine Change between Open and Close Values
Dim Yearly_Change As Long
Dim i As Long
Yearly_Change = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Yearly_Change
Cells(i, 10).Value = Cells(i, 3).Value - Cells(i, 6).Value
Next i

'Convert Yearly Change to Percent
Dim percent_change As Long
Dim p As Long
percent_change = Cells(Rows.Count, 1).End(xlUp).Row

For p = 2 To percent_change
Cells(p, 11) = FormatPercent(Cells(p, 10))
Next p

'Highlighting Percent Change
Dim percentchange As Range
For Each percentchange In Range("K2:K753001")
    If percentchange <= 0 Then
        percentchange.Interior.ColorIndex = 3
        ElseIf percentchange > 0 Then
         percentchange.Interior.ColorIndex = 4
    End If
Next
    
'Finding Total Stock Volume
Dim lRow As Long, num As Long, v As Long
lRow = Range("A753001").End(xlUp).Row
For v = 2 To lRow
    num = Worksheet.Function.Match(Cells(v, 1), Range("G2:G" & lRow), 0)
    If v = num Then
        Cells(i, 7) = WorksheetFunction.SumIf(Range("G2:2" & lastRow), Cells(v, 1), Range("L2:L" & lastRow))
    End If
Next

'Determining Bonus Column Headers
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'Determinig Bonus Values
Cells(2, 16) = Application.WorksheetFunction.Max(Range("K2:K753001"))
Cells(3, 16) = Application.WorksheetFunction.Min(Range("K2:K753001"))
Dim bTicker As Long
Dim greatest As Long
greatest = 20.16
Set myrange = Range("K2:K753001")
bTicker = Application.WorksheetFunction.VLookup(greatest, myrange, 8, False)

End Sub

