Sub add_new_object(object_name)
    
    last_row = WorksheetFunction.CountA(Workbooks("MasterBook.xlsm").Sheets("MasterSheet").Range("A:A"))
    
    Workbooks("MasterBook.xlsm").Sheets("MasterSheet").Range("a" & last_row & "").Value = object_name

    Workbooks("MasterBook.xlsm").Sheets("MasterSheet").Range("b" & last_row & "").Value = Selection.Value

End Sub
Sub open_personal()
    For Each element In Workbooks
        If element.Name = "Personal.xlsm" Then
            MsgBox "Mastersheet уже открыть"
            Else: Workbooks.Open ("C:\Users\L1s14\AppData\Roaming\Microsoft\Excel\Random_stuf\MasterBook.xlsm")
        End If
    Next
End Sub

