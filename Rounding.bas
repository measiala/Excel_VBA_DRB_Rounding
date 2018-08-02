Sub RoundEstimate()
    'Perform DRB rounding of estimates to four significant figures
    Dim cell As Object

    For Each cell In Selection
        If IsNumeric(cell) Then
            unrndval = Val(cell)
            If unrndval <> 0 Then
                absval = Abs(unrndval)
                base10 = 10 ^ Application.WorksheetFunction.Floor(Application.WorksheetFunction.Log10(absval), 1)
                rndval = Application.WorksheetFunction.Round(unrndval / base10, 3) * base10
            ElseIf unrndval = 0 Then
                base10 = 1
                rndval = 0
            End If
            cell = rndval
        Else
            MsgBox "The cell value of " & cell & " is not a number."
        End If   
    Next cell
End Sub
Sub RoundCount()
    'Perform DRB / FSRDC rounding of unweighted counts
    Dim cell As Object

    For Each cell In Selection
        If IsNumeric(cell) Then
            unrndval = Val(cell)
            If Application.WorksheetFunction.Trunc(unrndval) <> unrndval Or unrndval < 0 Then
                'cell value was either negative or a noninteger
                MsgBox "The cell value of " & unrndval & " is not consistent with expectations for an unweighted observation count."
            Else            
                'cell value was a nonnegative integer
                If unrndval = 0 Then
                    'potentially code this with next case, n < 15
                    cell = 0
                ElseIf unrndval < 15 Then
                    'If 1-14 show N < 15
                    cell = "N < 15"
                ElseIf unrndval < 1e6 Then
                    'For values less than 1m, the base multiple varies
                    If unrndval < 1e2 Then
                        'If 15-99 round to the nearest 10
                        basernd = 10
                    ElseIf unrndval < 1e3 Then
                        'If 100-999 round to the nearest 50
                        basernd = 50
                    ElseIf unrndval < 1e4 Then
                        'If 1,000-9,999 round to the nearest 100
                        basernd = 100
                    ElseIf unrndval < 1e5 Then
                        'If 10,000-99,999 round to the nearest 500
                        basernd = 500
                    Else
                        'If 100,000-999,999 round to nearest 1,000
                        basernd = 1000                    
                    End If
                    rndval = Application.WorksheetFunction.Round(unrndval / basernd, 0) * basernd
                    cell = rndval
                Else
                    'For values 1m or greater, round to 4 significant figures 
                    base10 = 10 ^ Application.WorksheetFunction.Floor(Application.WorksheetFunction.Log10(unrndval), 1)
                    rndval = Application.WorksheetFunction.Round(unrndval / base10, 3) * base10
                    cell = rndval
                End If
            End If
        Else
            MsgBox "The cell value of " & cell & " is not a number."
        End If        
    Next cell
End Sub