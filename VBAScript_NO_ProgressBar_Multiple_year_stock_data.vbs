
Option Explicit

Sub VBA_Challenge_NoProgressBar()

Dim cr As Long, i As Long, t As Long, s As Long, e As Long, WSCount As Double, j As Double, k As Long

Application.ScreenUpdating = False 'This helps speed up the process and only show results.
WSCount = ActiveWorkbook.Worksheets.Count 'Count amount of Worksheets in this workbook.

    
For k = 1 To WSCount 'Clear all worksheets (where calculated data to be placed) before starting.
    
    Worksheets(k).Activate
    Range("J1:R1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
        
Next k
 
    
For j = 1 To WSCount 'Run script on all worksheets
     
    Worksheets(j).Activate 'Activates each worksheet
    Range("A2").Select 'Select starting position
    Range(Selection, Selection.End(xlDown)).Select 'Select all cells on active column
    cr = Selection.Rows.Count + 1 'Count rows and added one since we start on row#2
    s = 2  'Variable used to save starting point (row 2) and later value is going to be adjusted as ticker changes.
    t = 1
        
                
    For i = 2 To cr 'Run script on active worksheet
                 
        'Add column titles for values to be calculated at the start of sequence.
        If i = 2 Then
        
            Range("J1,Q1") = "Ticker    "
            Range("K1") = "Yearly Change"
            Range("L1") = "Percentage Change"
            Range("M1") = "Total Stock Volume"
            Range("P2") = "Greatest % Increase"
            Range("P3") = "Greatest % Decrease"
            Range("P4") = "Greatest Total Volume"
            Range("R1") = "Value"
            With Range("J1:R1,P2:P4") 'Format cells with data titles.
                .Font.Bold = True
                .EntireColumn.AutoFit
                .HorizontalAlignment = xlCenter
            End With
                
        End If
        
            'This is the condition used to determine the start and end point of each ticker name
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            t = t + 1 'Used to determine which row to place findings (ticker).
                
            'The start ticker name location was defined as s = 2. This variable will change after by assigning it the value of "e".
            e = i + 1 'This variable stores the end location of a ticker name + 1 which will turn into the start location of next calculation.
                
                          
            Cells(t, 10) = Cells(i, 1).Value 'Places ticker Name in appropriate location.
               
                    
            'Conditions to determine if price increased or dropped also considering if divisor different from 0
            If Cells(i, 6).Value > Cells(s, 6).Value And Cells(s, 6).Value <> 0 Then
                Cells(t, 11) = Round(Cells(i, 6).Value - Cells(s, 6).Value, 2)
                Cells(t, 11).Interior.ColorIndex = 4 'Cell filled with color Green because value its positive
                Cells(t, 12) = Round(((Cells(i, 6).Value - Cells(s, 6).Value) / Cells(s, 6).Value), 2) * 100 & "%"
                Cells(t, 13) = Application.WorksheetFunction.Sum(Range("G" & s & ":G" & i)) ' Adding total stock volume
                s = e 'Starting location for each ticker is now equal to "e".
            
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
          
    'Calling Sub to generate Greatest% Increase, decrease and Total Volume.
    Call GreatestValuesArray

    Range("R4").EntireColumn.AutoFit 'Adjust cell to fit Greatest Total Volume data.
    Range("A2").Select 'Selects starting point. 'For a cleaner look, this will select starting point and not show all the selected rows needed for the calculations.
Next j
    
            
'This will go back to startinig work sheet.
Worksheets(1).Activate 'Activate first worksheet
    

MsgBox "Successfully Completed!"

End Sub

Sub GreatestValuesArray()
Dim GreatestValuesArray()
Dim r As Long, c As Long
Dim TotalRows As Long, TotalColumns As Long


TotalRows = Range("J2", Range("J2").End(xlDown)).Cells.Count ' Count rows
TotalColumns = Range("J2", Range("J2").End(xlToRight)).Cells.Count ' Count Columns

ReDim GreatestValuesArray(1 To TotalRows, 1 To TotalColumns) ' Count columns

'This loop is to assign values to array.
For r = LBound(GreatestValuesArray, 1) To UBound(GreatestValuesArray, 1) ' 1 is the element number (in this case rows) /LowerBound & Upperbound meanings
    
    For c = LBound(GreatestValuesArray, 1) To UBound(GreatestValuesArray, 2) ' 2 is the element number (in this case columns)
        
        GreatestValuesArray(r, c) = Range("J2").Offset(r - 1, c - 1).Value
 
    Next c

Next r


'Re-used same loop variables
'This loop will be used to extract values from array based on conditions.
For r = LBound(GreatestValuesArray, 1) To UBound(GreatestValuesArray, 1) '1 is the element number (in this case rows) /LowerBound & Upperbound
        'Conditions to obtain Greatest % increase
        If r = UBound(GreatestValuesArray, 1) Then 'Exit when reach the last cell of the worksheet.
            Exit For
        End If
        If GreatestValuesArray(r, 3) > GreatestValuesArray(r + 1, 3) And GreatestValuesArray(r, 3) > Cells(2, 18).Value Then
            Cells(2, 17).Value = GreatestValuesArray(r, 1) 'Ticker
            Cells(2, 18).Value = Format((GreatestValuesArray(r, 3)), "Percent") 'Greatest %Increase Value
        ElseIf GreatestValuesArray(r, 3) < GreatestValuesArray(r + 1, 3) And GreatestValuesArray(r + 1, 3) > Cells(2, 18).Value Then
            Cells(2, 17).Value = GreatestValuesArray(r + 1, 1) 'Ticker
            Cells(2, 18).Value = Format((GreatestValuesArray(r + 1, 3)), "Percent") 'Greatest %Increase Value
        End If
        
        'Conditions to obtain Greatest % decrease
        If GreatestValuesArray(r, 3) < GreatestValuesArray(r + 1, 3) And GreatestValuesArray(r, 3) < Cells(3, 18).Value Then
            Cells(3, 17).Value = GreatestValuesArray(r, 1) 'Ticker
            Cells(3, 18).Value = Format((GreatestValuesArray(r, 3)), "Percent") 'Greatest %Decrease Value
        ElseIf GreatestValuesArray(r, 3) > GreatestValuesArray(r + 1, 3) And GreatestValuesArray(r + 1, 3) < Cells(3, 18).Value Then
            Cells(3, 17).Value = GreatestValuesArray(r + 1, 1) 'Ticker
            Cells(3, 18).Value = Format((GreatestValuesArray(r + 1, 3)), "Percent") 'Greatest %Decrease Value
        End If
        
        'Conditions to obtain Greatest Total Volume
        If GreatestValuesArray(r, 4) > GreatestValuesArray(r + 1, 4) And GreatestValuesArray(r, 4) > Cells(4, 18).Value Then
            Cells(4, 17).Value = GreatestValuesArray(r, 1) 'Ticker
            Cells(4, 18).Value = GreatestValuesArray(r, 4) 'Greatest Total Volume
        ElseIf GreatestValuesArray(r, 4) < GreatestValuesArray(r + 1, 4) And GreatestValuesArray(r + 1, 4) > Cells(4, 18).Value Then
            Cells(4, 17).Value = GreatestValuesArray(r + 1, 1) 'Ticker
            Cells(4, 18).Value = GreatestValuesArray(r + 1, 4) 'Greatest Total Volume
        End If

Next r

End Sub

