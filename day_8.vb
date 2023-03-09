Sub day_8()

'masterbook = [Masterbook.xlsm]MasterSheet

day8 = Left(Now, 10)

For Each element In Range("d2:d119")
    If day8 = element.Text Then
        today_is_a_holiday_or_week_end = True
        element.Select
        Exit For
    End If
Next
Do Until today_is_a_holiday_or_week_end = False
    If day8 = Selection.Text Then
        offset = offset + 1
        day8 = Left(Now - offset, 10)
        Selection.offset(-1).Select
    Else
        today_is_a_holiday_or_week_end = False
    End If
Loop
MsgBox day8
End Sub
