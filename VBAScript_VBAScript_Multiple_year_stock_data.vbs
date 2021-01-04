Option Explicit

Sub VBA_ChallengeVIP()

Dim cr As Long, ct As Long, i As Long, t As Long, s As Long, e As Long, WSCount As Double, j As Double, totalrows As Long
Dim k As Long, amountcompleted As Long, istored As Long
Dim CurrentProgress As Double, ProgressPercentage As Double, BarWidth As Long, Bar2Width As Long, TotalProgress As Double ', ProgressPercentage2
Dim PercentColumn As Long, m As Long

Application.ScreenUpdating = False ' This helps speed up the process and only show results.
WSCount = ActiveWorkbook.Worksheets.Count ' Count amount of Worksheets in this workbook.
Call InitializeProgressBar

Do Until i > cr And j > WSCount 'Progress Bar loop

    
    For k = 1 To WSCount 'Sums all rows from all worksheets for progress bar % calculation
    
        Worksheets(k).Activate
        Range("J1:R1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Clear
        
    Next k
 
    
    For j = 1 To WSCount 'Run script on all worksheets
     
        Worksheets(j).Activate 'Activates each worksheet
        Range("A2").Select 'Select starting position
        Range(Selection, Selection.End(xlDown)).Select 'Select all cells on active column
        cr = Selection.Rows.Count + 1 'Added one to total row count since we start on row#2
        ct = 1
        s = 2  'Assigned location of "s" on first cell to be later changed
        t = 1
        
           'Percentage completed in the entire workbook.
        TotalProgress = (j - 1) / WSCount
        Bar2Width = Progress.Border2.Width * TotalProgress
        Progress.Bar2.Width = Bar2Width
        Progress.text2.Caption = "Worksheet " & j - 1 & " / " & WSCount
           
        DoEvents
                
        For i = 2 To cr 'Run script on active worksheet
                 
          
           
           'Percentage completed within a worksheet
           CurrentProgress = (i / cr)
           BarWidth = Progress.Border.Width * CurrentProgress
           ProgressPercentage = Round(CurrentProgress * 100, 0)
           Progress.Bar.Width = BarWidth
           Progress.Text.Caption = ProgressPercentage & "% Complete"
                                
           
           DoEvents
        
            If i = 2 Then 'Add column title for values to be calculated at start of sequence.
        
                Range("J1,Q1") = "Ticker    "
                Range("K1") = "Yearly Change"
                Range("L1") = "Percentage Change"
                Range("M1") = "Total Stock Volume"
                Range("P2") = "Greatest % Increase"
                Range("P3") = "Greatest % Decrease"
                Range("P4") = "Greatest Total Volume"
                'Range("Q1") = "Ticker"
                Range("R1") = "Value      "
                With Range("J1:R1,P2:P4,R4") 'Format cells with data titles.
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
                t = t + 1 'Used to determine which row to place findings (ticker).
                e = i + 1 'This variable stores the end price location + 1
                          'which will turn into the start location of next calculation before it is zeroed.
                Cells(t, 10) = Cells(i, 1).Value 'Places ticker value in appropriate location.
                    
                'Conditions to determine if price increased or dropped also considering if divisor different from 0
            If Cells(i, 6).Value > Cells(s, 6).Value And Cells(s, 6).Value <> 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 4 'Cell filled with color Green because value its positive
                Cells(t, 12) = Round(((Cells(i, 6).Value - Cells(s, 6).Value) / Cells(s, 6).Value), 2) * 100 & "%"
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i)) ' Adding total stock volume
                s = e 'Stored previous store value for starting point for next calculation.
            
            ElseIf Cells(i, 6).Value < Cells(s, 6).Value And Cells(s, 6).Value <> 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 3 'Cell filled with color Red because value its negative
                Cells(t, 12) = Round(((((Cells(s, 6).Value - Cells(i, 6).Value) / Cells(s, 6).Value) * 100) * (-1)), 2) & "%"
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i))
                s = e
                
            'Condition to determine if price increased or dropped if divisor equal to 0
            ElseIf Cells(i, 6).Value > Cells(s, 6).Value And Cells(s, 6).Value = 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 4
                Cells(t, 12) = Round(0, 2) & "%" ' To avoid errors in calculation, value is equal to 0%
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i))
                s = e
            ElseIf Cells(i, 6).Value < Cells(s, 6).Value And Cells(s, 6).Value = 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 3
                Cells(t, 12) = Round(0, 2) & "%"
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i))
                s = e
        
                
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

    
    '''''''''''Call Array'''''''''''''''''''
    
    
    'This will unselect all previous selected cells and return to starting point.
    Worksheets(1).Activate 'Activate first worksheet
    Range("A2").Select 'Selects starting point.
Loop



Unload Progress 'Unload user form at the end.



End Sub


Sub InitializeProgressBar() 'Initialize Progress Bar with values shown below.

With Progress
    .Bar.Width = 0
    .Bar2.Width = 0
    .Text.Caption = "0% Complete"
    .text2.Caption = "0/0" ' Amount of worksheets completed starting point.
    .Show vbModeless 'Modeless allows the user to interact with excel sheet while progress and calculations are going through.
End With

End Sub