Option Explicit

'These variables must be stored as a public variable so that it is available across macros

'StoreData saves the contents of the selection itself
Public StoreData As Variant
'StoreRange allows the user to recover what cells were selected even if they clicked on another cell
Public StoreRange As Range

Sub DRB_Round_Count()
    'Perform DRB / FSRDC rounding of unweighted counts
    Dim cell As Object
    Dim unrndval As Double
    Dim absval As Double
    Dim basernd As Double
    Dim base10 As Double
    Dim rndval As Double

    StoreData = Selection.Value
    Set StoreRange = Selection
    
    For Each cell In Selection
        If IsNumeric(cell) Then
            unrndval = Val(cell)
            absval = Abs(unrndval)
            If (absval - Int(absval)) > 0 Or unrndval < 0 Then
                'cell value was either negative or a noninteger
                MsgBox "The cell value of " & unrndval & _
                    " is not consistent with expectations for an unweighted observation count. Perhaps this is an estimate and not a count?"
            Else
                'cell value was a nonnegative integer
                If unrndval = 0 Then
                    'potentially code this with next case, n < 15
                    cell = 0
                ElseIf unrndval < 15 Then
                    'If 1-14 show N < 15
                    cell = "N < 15"
                ElseIf unrndval < 1000000# Then
                    'For values less than 1m, the base multiple varies
                    If unrndval < 100# Then
                        'If 15-99 round to the nearest 10
                        basernd = 10
                    ElseIf unrndval < 1000# Then
                        'If 100-999 round to the nearest 50
                        basernd = 50
                    ElseIf unrndval < 10000# Then
                        'If 1,000-9,999 round to the nearest 100
                        basernd = 100
                    ElseIf unrndval < 100000# Then
                        'If 10,000-99,999 round to the nearest 500
                        basernd = 500
                    Else
                        'If 100,000-999,999 round to nearest 1,000
                        basernd = 1000
                    End If
                    rndval = Round(unrndval / basernd, 0) * basernd
                    cell = rndval
                Else
                    'For values 1m or greater, round to 4 significant figures
                    base10 = 10 ^ Application.WorksheetFunction.Floor(Application.WorksheetFunction.Log10(unrndval), 1)
                    rndval = Round(unrndval / base10, 3) * base10
                    cell = rndval
                End If
            End If
        Else
            MsgBox "The cell value of " & cell & " is not a number."
        End If
    Next cell
End Sub

Sub DRB_Round_Estimate()
    'Perform DRB rounding of estimates to four significant figures
    Dim cell As Object
    Dim unrndval As Double
    Dim absval As Double
    Dim base10 As Double
    Dim rndval As Double
    
    StoreData = Selection.Value
    Set StoreRange = Selection

    'Store a copy of selection contents before overwriting
    For Each cell In Selection
        If IsNumeric(cell) Then
            unrndval = cell.Value
            If unrndval <> 0 Then
                'absval is probably more explicit than necessary but demonstrates fn works for pos and neg values
                absval = Abs(unrndval)
                base10 = 10 ^ Application.WorksheetFunction.Floor(Application.WorksheetFunction.Log10(absval), 1)
                rndval = Round(unrndval / base10, 3) * base10
            ElseIf unrndval = 0 Then
                'If unrndval is 0 then log10 fails so treat special case
                base10 = 1
                rndval = 0
            End If
            cell = rndval
        Else
            MsgBox "The cell value of " & cell & " is not a number."
        End If
    Next cell
End Sub

Sub DRB_Restore_Selection()
    'Restore the values for the selected cells using the stored copy of the data

    'Select the originally selected cells
    StoreRange.Select
    'Restore the data to the selection
    Selection.Value = StoreData
End Sub
