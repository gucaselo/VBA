Attribute VB_Name = "Module1"
Option Explicit

Sub VBA_Challenge()

Dim cr As Long, ct As Long, i As Long, t As Integer, s As Variant, e As Double
Range("K2").Select
'Application.ScreenUpdating = False ' This speeds up the calculations and hides calculations.
'Range("A2", Range("A2").End(xlDown)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo
Range(Selection, Selection.End(xlDown)).Select ' Select all cells on active colum
cr = Selection.Rows.Count + 1 ' Assuming no empty cells in between; count all rows in selected column
ct = 1
s = ActiveCell.Row - 1 'Assign location of s on first cell to be later changed
For i = ct To cr
    If s = 1 And ActiveCell.Offset(i - 1, 0).Value = ActiveCell.Offset(i, 0).Value Then
       ct = ct + 1
    
    ElseIf ActiveCell.Offset(i - 1, 0) = ActiveCell.Offset(i, 0) Then
       ' s = e ' Store e value which is ct in variable s in order to compare start and end prices
        ct = ct + 1
    
    Else
        t = t + 1 'Used to determine which row to place findings
       's = ActiveCell.Value ' start of the year price
        e = ct 'end of the year price variable used to stored value of ct before it is zeroed
        
        's = WorksheetFunction.Average((ActiveCell.Offset(ct - 1, 5)), (ActiveCell.Offset(xx, 0)))
        ActiveCell.Offset(t - 1, 10) = ActiveCell.Offset(ct - 1, 0).Value
        'ActiveCell.Offset(t - 1, 11) = ct 'count not used
    
        'Condition to determine how to determine if price increased or dropped
        If ActiveCell.Offset(ct - 1, 1).Value > ActiveCell.Offset(s - 1, 1).Value Then
            ActiveCell.Offset(t - 1, 11) = ActiveCell.Offset(ct - 1, 1).Value - ActiveCell.Offset(s - 1, 1).Value
            s = e ' Store e value which is ct in variable s in order to compare start and end prices
        ElseIf ActiveCell.Offset(ct - 1, 1).Value < ActiveCell.Offset(s - 1, 1).Value Then
        ActiveCell.Offset(t - 1, 11) = ActiveCell.Offset(s - 1, 1).Value - ActiveCell.Offset(ct - 1, 1).Value
            s = e ' Store e value which is ct in variable s in order to compare start and end prices
        
           ' ActiveCell.Offset(t - 1, 11) = WorksheetFunction.Average((ActiveCell.Offset(ct - 1, 1).Value), (ActiveCell.Offset("s" - 1, 2).Value))
        ' Above Change 1 and 2 to their correct colum once we try on real data on this sheet
        'ActiveCell.Offset(t - 1, 11) = ct - ActiveCell.Offset(t - 2, 11)
    'ct = 0
        End If
    End If
   
Next i


End Sub
Sub VBA_Challenge2()

Dim cr As Long, ct As Long, i As Long, t As Long, s As Long, e As Long, WSCount As Double, j As Double, totalrows As Long
Dim k As Long, amountcompleted As Long, istored As Long
Dim CurrentProgress As Double, ProgressPercentage As Double, BarWidth As Long, Bar2Width As Long, TotalProgress As Double, ProgressPercentage2
Dim PercentColumn As Long


'For j = 1 To WSCount
'''Range("A2").Select
'Application.ScreenUpdating = False ' This speeds up the calculations and hides calculations.
'Range("A2", Range("A2").End(xlDown)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo
'''Range(Selection, Selection.End(xlDown)).Select ' Select all cells on active colum
'''cr = Selection.Rows.Count + 1 ' Assuming no empty cells in between; count all rows in selected column
WSCount = ActiveWorkbook.Worksheets.Count ' Count amount of Worksheets in this workbook.
'''ct = 1
'''s = 2  'Assign location of s on first cell to be later changed
'''t = 1 ' Started at 1 since values start at second row

Call InitializeProgressBar
         '   CurrentProgress = (i / cr) + (j / WSCount)
         '   BarWidth = Progress.Border.Width * CurrentProgress
         '   ProgressPercentage = Round(CurrentProgress * 100, 0) ' To get percentage and round it to its near integer.
    
         '   Progress.Bar.Width = BarWidth
         '   Progress.Text.Caption = ProgressPercentage & "% Complete"
    
         '  DoEvents

Do Until i > cr And j > WSCount

    
    For k = 1 To WSCount
    
    Worksheets(k).Activate
    Range("J1:R1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    'Selection.ClearFormats
    
    Next k
 
    
    For j = 1 To WSCount ' Run script on all worksheets
     
        'ActiveWorkbook.Worksheets(j).Select
        Worksheets(j).Activate ' Select each worksheet
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select ' Select all cells on active colum
        'amountcompleted = i + j ' store i variable after each worksheet is completed
        cr = Selection.Rows.Count + 1
        ct = 1
        s = 2  'Assign location of s on first cell to be later changed
        t = 1
        
           'Percentage completed workbook
           TotalProgress = (j - 1) / WSCount
           Bar2Width = Progress.Border2.Width * TotalProgress
         ' ProgressPercentage2 = Round(TotalProgress * 100, 0) ' To get percentage and round it to its near integer.
           Progress.Bar2.Width = Bar2Width
           Progress.text2.Caption = "Worksheet " & j - 1 & " / " & WSCount
           
           DoEvents
                
        For i = 2 To cr
                 
          
           
           'Percentage completed within a worksheet
           CurrentProgress = (i / cr)
           BarWidth = Progress.Border.Width * CurrentProgress
           ProgressPercentage = Round(CurrentProgress * 100, 0) ' To get percentage and round it to its near integer.
           Progress.Bar.Width = BarWidth
           Progress.Text.Caption = ProgressPercentage & "% Complete"
                                
           
           DoEvents
        
        
        
            If i = 2 Then
        
                Range("J1") = "Ticker  "
                Range("K1") = "Yearly Change"
                Range("L1") = "Percentage Change"
                Range("M1") = "Total Stock Volume"
                Range("P2") = "Greatest % Increase"
                Range("P3") = "Greatest % Decrease"
                Range("P4") = "Greatest Total Volume"
                Range("Q1") = "Ticker"
                Range("R1") = "Value      "
                With Range("J1:R1,P2:P4,R4")
                    .Font.Bold = True
                    .EntireColumn.AutoFit
                    .HorizontalAlignment = xlCenter
                End With
                
            End If
        
            If s = 2 And Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                ct = ct + 1
    
            ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            ' s = e ' Store e value which is ct in variable s in order to compare start and end prices
                ct = ct + 1
    
            Else
                t = t + 1 'Used to determine which row to place findings
                e = i + 1 'end of the year price variable used to stored value of ct before it is zeroed
                Cells(t, 10) = Cells(i, 1).Value
                    
                'Condition to determine how to determine if price increased or dropped if divisor different from 0
            If Cells(i, 6).Value > Cells(s, 6).Value And Cells(s, 6).Value <> 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 4
                Cells(t, 12) = Round(((Cells(i, 6).Value - Cells(s, 6).Value) / Cells(s, 6).Value), 2) * 100 & "%"
                ' Above VBA added percentage value on its own without the need to multiplying by 100 or adding a symbol.
                'Cells(t, 12).Interior.ColorIndex = 4
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i)) ' Adding total stock volume
            '    Cells(t, 14) = Application.WorksheetFunction.SumIf(Range("A:A"), "J" & s, Range("G:G")) ' Comparing above
                s = e ' Store e value which is ct in variable s in order to compare start and end prices
            ElseIf Cells(i, 6).Value < Cells(s, 6).Value And Cells(s, 6).Value <> 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 3
                Cells(t, 12) = Round(((((Cells(s, 6).Value - Cells(i, 6).Value) / Cells(s, 6).Value) * 100) * (-1)), 2) & "%"
                ' Above VBA added percentage value on its own without the need to multiplying by 100 or adding a symbol.
                'Cells(t, 12).Interior.ColorIndex = 3
                ' Add color conditional
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i))
                s = e
            'Condition to determine how to determine if price increased or dropped if divisor equal to 0
            ElseIf Cells(i, 6).Value > Cells(s, 6).Value And Cells(s, 6).Value = 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 4
                Cells(t, 12) = Round(0, 2) & "%"
                ' Above VBA added percentage value on its own without the need to multiplying by 100 or adding a symbol.
                'Cells(t, 12).Interior.ColorIndex = 4
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i)) ' Adding total stock volume
                s = e ' Store e value which is ct in variable s in order to compare start and end prices
            ElseIf Cells(i, 6).Value < Cells(s, 6).Value And Cells(s, 6).Value = 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 3
                Cells(t, 12) = Round(0, 2) & "%"
                ' Above VBA added percentage value on its own without the need to multiplying by 100 or adding a symbol.
                'Cells(t, 12).Interior.ColorIndex = 3
                ' Add color conditional
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i)) ' Adding total stock volume
                s = e ' Store e value which is ct in variable s in order to compare start and end prices
        
                
            End If
            End If
   
          Next i
          
          
            Range("L2").Select
            Range(Selection, Selection.End(xlDown)).Select
            PercentColumn = Selection.Rows.Count + 1

            Cells(2, 18) = (Round(Application.WorksheetFunction.Max(Range("L2" & ":L" & PercentColumn)), 2)) * 100
            'Cells(2, 17) = Application.WorksheetFunction.Index(Cells.Offset
            Cells(3, 18) = (Round(Application.WorksheetFunction.Min(Range("L2" & ":L" & PercentColumn)), 2)) * 100
            Cells(4, 18) = Application.WorksheetFunction.Max(Range("M2" & ":M" & PercentColumn))
            
           ' Cells(2, 17) = Application.WorksheetFunction.Index(Range("M2" & ":M" & PercentColumn), Application.WorksheetFunction.Match(Cells(2, 18).Value, rngColumn2, 1), 1)
            
                   
    Next j
'Range("L2").Select
'Range(Selection, Selection.End(xlDown)).Select
'PercentColumn = Selection.Rows.Count + 1

'If Range("L2" & ":L" & PercentColumn).Value > 0 Then
'Cells(2, 18) = Application.WorksheetFunction.Max(Range("L2" & ":L" & PercentColumn))

'ElseIf Range("L2" & ":L" & PercentColumn).Value < 0 Then
'Cells(3, 18) = Application.WorksheetFunction.Min(Range("L2" & ":L" & PercentColumn))
    
    
'End If
    'For j = 2 To PercentColumn
    
    'Next j
    
    Worksheets(1).Activate ' Select each worksheet
    Range("A2").Select
Loop



Unload Progress 'unload user form



End Sub

Sub ProgressBar()

Dim i As Long
Dim totalrows As Long
Dim CurrentProgress As Double
Dim ProgressPercentage As Double
Dim BarWidth As Long

'i = 2
'TotalRows = Application.CountA(Range("A:A"))
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select ' Select all cells on active colum
totalrows = Selection.Rows.Count
i = 2

'Call InitializeProgressBar

Do While i <= totalrows

    'Cells(i, 10).Value = Cells(i, 10).Value * 1.1
    Cells(i, 8) = Cells(i, 1).Value
    
    CurrentProgress = i / totalrows
    BarWidth = Progress.Border.Width * CurrentProgress
    ProgressPercentage = Round(CurrentProgress * 100, 0) ' To get percentage and round it to its near integer.
    
    Progress.Bar.Width = BarWidth
    Progress.Text.Caption = ProgressPercentage & "% Complete"
    
    DoEvents ' ensuring events are still happening during the progress in macros
    
    i = i + 1
    
Loop



Unload Progress 'unload user form


End Sub

Sub InitializeProgressBar()

With Progress

    .Bar.Width = 0
    .Bar2.Width = 0
    .Text.Caption = "0% Complete"
    .text2.Caption = "0/0" ' Amount of worksheets completed
    .Show vbModeless ' modeless allows the user to interact with excel sheet while progress and calculations are going through.
End With

End Sub

Sub testcode()
Dim PercentColumn As Long
'Dim s As Long
's = 2
'i = 264
Range("N2").Select
Range(Selection, Selection.End(xlDown)).Select
PercentColumn = Selection.Rows.Count + 1
Cells(2, 18) = Application.WorksheetFunction.Max(Range("N2" & ":N" & PercentColumn))

End Sub


Sub CountAllRows()

'Count all the rows on all the worksheets for progress bar
    For k = 1 To WSCount
        Worksheets(k).Activate
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select ' Select all cells on active colum
        totalrows = totalrows + Selection.Rows.Count + 1
    Next k

End Sub

Sub GreatestValues()

Dim PercentColumn As Double, m As Long

Range("L2").Select
Range(Selection, Selection.End(xlDown)).Select
PercentColumn = Selection.Rows.Count + 1

'For m = 2 To PercentColumn
'If Range("O2" & ":L" & PercentColumn).Value > 0 Then
Cells(2, 19) = Application.WorksheetFunction.Max(Range("P2" & ":P" & PercentColumn))


'ElseIf Range("O2" & ":O" & PercentColumn).Value < 0 Then
Cells(3, 19) = Application.WorksheetFunction.Min(Range("P2" & ":P" & PercentColumn))
    
    
'End If
    'For j = 2 To PercentColumn
    
    'Next j
'Next m

End Sub
