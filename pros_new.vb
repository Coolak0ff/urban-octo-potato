
Sub Табличка_просрочек_2()
    
    Dim DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist As Boolean
    
    Sheets("sheet1").Select
    PosStr = WorksheetFunction.CountA(Range("A:A"))
    Call DeuTable(PosStr)
    Call is_exist(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)
    Call column_with_name
    Call sum
    Call p(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)
    
    o = 0
    
    Range("a1").Select

    For Each element In Range("D12:K12")
    With element
        .FormulaR1C1 = "=WORKDAY(NOW()," & o & ",MasterSheet!R2C2:R118C2)"
        .offset(-4).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW()," & o & ",MasterSheet!R2C2:R118C2)-NOW(),0)+1"
        .offset(-3).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW()," & o & ",MasterSheet!R2C2:R118C2)-NOW(),0)+2"
    End With
        Range("d12").offset(-4).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),0,MasterSheet!R2C2:R118C2)-NOW(),0)"
        Range("d12").offset(-3).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),0,MasterSheet!R2C2:R118C2)-NOW(),0)+1"
        
        Range("k12").offset(-4).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),7,MasterSheet!R2C2:R118C2)-NOW(),0)+1"
        Range("k12").offset(-3).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),10,MasterSheet!R2C2:R118C2)-NOW(),0)+2"
        
        Range("d8:k9").NumberFormat = "General"
        
        i = Range("d8").offset(o).Text
        o = o + 1
        element.offset(, 1).Select
    Next
    Call chtoto(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)
    'Call first_day(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist, d)
    Range("C12:J12").NumberFormat = "dddd dd/mm"
    Call format
    Call white_if_sum_of_critical_is_0
    
End Sub
Sub chtoto(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)

'Range("d13").Select
For Column = 0 To 7
For Each element In Range("d13:d20").offset(, Column)
    i1 = Range("d8").offset(, Column).Text
    i2 = Range("d9").offset(, Column).Text
    Call chtoto_formula(element, i1, i2, d, DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)
    d = d + 1
    element.offset(1).Select
Next
d = 0
Next

End Sub

Sub chtoto_formula(element, i1, i2, d, DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)

    If d = 0 Then
        d1 = """ДЭУ-1"""
        d2 = """ДЭУ-2"""
        If DEU2_exist = False Then d2 = """ДЭУ-3"""
        If DEU3_exist = False Then d2 = """ДЭУ-4"""
        If DEU4_exist = False Then d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 1 Then
        d1 = """ДЭУ-2"""
        d2 = """ДЭУ-3"""
        If DEU3_exist = False Then d2 = """ДЭУ-4"""
        If DEU4_exist = False Then d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 2 Then
        d1 = """ДЭУ-3"""
        d2 = """ДЭУ-4"""
        If DEU4_exist = False Then d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 3 Then
        d1 = """ДЭУ-4"""
        d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 4 Then
        d1 = """ДЭУ-5"""
        d2 = """Зелёнка"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 5 Then
        d1 = """ДРУ"""
        d2 = """ДЭУ-1"""
        If DEU1_exist = False Then d2 = """ДЭУ-2"""
        If DEU2_exist = False Then d2 = """ДЭУ-3"""
        If DEU3_exist = False Then d2 = """ДЭУ-4"""
        If DEU4_exist = False Then d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 6 Then
        d1 = """Зелёнка"""
        d2 = """#Н/Д"""
        If na_exist = False Then d2 = """Общий итог"""
    End If
    If d = 7 Then
        d1 = """#Н/Д"""
        d2 = """Общий итог"""
    End If
    
    element.FormulaR1C1 = _
    "=COUNTIFS(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),"">" & i1 & """,INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),""<" & i2 & """)"
    
    If element.Value = "#N/A" Then element.Value = "0"

End Sub


Sub first_day(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist, d)
    
    With Range("k12")
        .offset(-4).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),7,MasterSheet!R2C2:R118C2)-NOW(),0)+1"
        .offset(-3).FormulaR1C1 = "=ROUNDDOWN(WORKDAY(NOW(),9,MasterSheet!R2C2:R118C2)-NOW(),0)+2"
    End With
    
    Range("k8:k9").NumberFormat = "General"
    
    Range("K13:K20").Select
    For Each element In Selection
    If d = 1 Then
        d1 = """ДЭУ-1"""
        d2 = """ДЭУ-2"""
        n1 = Range("k8").Text
        n2 = Range("k9").Text
        Call first_day_formula(d1, d2, n1, n2)
        If DEU1_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 2 Then
        d1 = """ДЭУ-2"""
        d2 = """ДЭУ-3"""
        Call first_day_formula(d1, d2, n1, n2)
        If DEU2_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 3 Then
        d1 = """ДЭУ-3"""
        d2 = """ДЭУ-4"""
        Call first_day_formula(d1, d2, n1, n2)
        If DEU3_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 4 Then
        d1 = """ДЭУ-4"""
        d2 = """ДЭУ-5"""
        Call first_day_formula(d1, d2, n1, n2)
        If DEU4_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 5 Then
        d1 = """ДЭУ-5"""
        d2 = """Зелёнка"""
        Call first_day_formula(d1, d2, n1, n2)
        If DEU5_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 6 Then
        d1 = """ДРУ"""
        d2 = """ДЭУ-1"""
        If na_exist = False Then d2 = """ДЭУ-2"""
        Call first_day_formula(d1, d2, n1, n2)
        If DRU_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 7 Then
        d1 = """Зелёнка"""
        d2 = """#Н/Д"""
        If na_exist = False Then d2 = """Общий итог"""
        Call first_day_formula(d1, d2, n1, n2)
        If Zel_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 8 Then
        d1 = """#Н/Д"""
        d2 = """Общий итог"""
        Call first_day_formula(d1, d2, n1, n2)
        If na_exist = False Then ActiveCell.Value = "0"
    End If
    d = d + 1
    ActiveCell.offset(1).Select
    Next
    
End Sub
Sub first_day_formula(d1, d2, n1, n2)

    ActiveCell.FormulaR1C1 = _
    "=COUNTIFS(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),"">" & n1 & """,INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),""<" & n2 & """)"

End Sub
Sub format()

    Range("C12:l12").NumberFormat = "dddd dd/mm"
    
    Range("A11:l21").Select
    
    With Selection
        .Borders.LineStyle = xlContinuous
        .ColumnWidth = 17
        .RowHeight = 25
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Range("A11:l12,A11:A21").Select
    With Selection.Interior
        .Color = 12632256 'gray
    End With
    
    With Range("D13:G21,L13:L21").Interior
        .Color = 4210943 'slightly red
    End With
    
    With Range("C13:C21").Interior
        .Color = 2039807 'dark red
    End With
    
    For Each element In Range("c13:g21,l13:l21")
        If element.Value > 0 Then
        element.Font.Bold = True
        End If
    Next
    
    Range("a11:a12,B11:B12,c11:c12,l11:l12").Merge
    
    For Each element In Range("b13:b20")
        If element.Value = 0 Then element.EntireRow.Hidden = True
    Next

End Sub
Sub white_if_sum_of_critical_is_0()

Cells(1, 1).Select
For Each element In Range("d21:l21")
    If element.Value = 0 Then Range(element.offset(-8), element).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
Next

End Sub
Sub Pros(PosStr)

    Range("AJ:AJ").Insert Shift:=xlToRight
    Range("AJ1").FormulaR1C1 = "Просрочка (срок - текущая дата)"
    Range("AJ2:AJ" & PosStr & "").Select
    With Selection
        .FormulaR1C1 = "=ROUNDdown(RC[-1]-NOW(),1)"
        .FillDown
        .NumberFormat = "General"
    End With
    
End Sub

Sub p(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)

    Range("C12").FormulaR1C1 = "Просрочка"
    d = 1
    Range("C13:C20").Select
    For Each elements In Selection
    If d = 1 Then
        d1 = """ДЭУ-1"""
        d2 = """ДЭУ-2"""
        Call pp(d1, d2)
        If DEU1_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 2 Then
        d1 = """ДЭУ-2"""
        d2 = """ДЭУ-3"""
        Call pp(d1, d2)
        If DEU2_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 3 Then
        d1 = """ДЭУ-3"""
        d2 = """ДЭУ-4"""
        Call pp(d1, d2)
        If DEU3_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 4 Then
        d1 = """ДЭУ-4"""
        d2 = """ДЭУ-5"""
        Call pp(d1, d2)
        If DEU4_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 5 Then
        d1 = """ДЭУ-5"""
        d2 = """Зелёнка"""
        Call pp(d1, d2)
        If DEU5_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 6 Then
        d1 = """ДРУ"""
        d2 = """ДЭУ-1"""
        If na_exist = False Then d2 = """ДЭУ-2"""
        Call pp(d1, d2)
        If DRU_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 7 Then
        d1 = """Зелёнка"""
        d2 = """#Н/Д"""
        If na_exist = False Then d2 = """Общий итог"""
        Call pp(d1, d2)
        If Zel_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 8 Then
        d1 = """#Н/Д"""
        d2 = """Общий итог"""
        Call pp(d1, d2)
        If na_exist = False Then ActiveCell.Value = "0"
    End If
    d = d + 1
    ActiveCell.offset(1).Select
    Next
    
End Sub
Sub column_with_name()

    Range("A13").FormulaR1C1 = "ДЭУ-1"
    Range("A14").FormulaR1C1 = "ДЭУ-2"
    Range("A15").FormulaR1C1 = "ДЭУ-3"
    Range("A16").FormulaR1C1 = "ДЭУ-4"
    Range("A17").FormulaR1C1 = "ДЭУ-5"
    Range("A18").FormulaR1C1 = "ДРУ"
    Range("A19").FormulaR1C1 = "Зелёнка"
    Range("A20").FormulaR1C1 = "#Н/Д"
    Range("A21").FormulaR1C1 = "Всего"
    i = 8
    For Each element In Range("D11:K11")
         element.FormulaR1C1 = i & " день"
         i = i - 1
    Next
    

End Sub

Sub pp(d1, d2)

    ActiveCell.FormulaR1C1 = "=COUNTIF(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),""<0"")"

End Sub
Sub sum()
    Range("b12").FormulaR1C1 = "Всего в работе"

    Call in_work
        
    For Each element In Range("b21:L21")
        element.FormulaR1C1 = "=SUM(R[-1]C:R[-8]C)"
    Next
        
    Range("L12").FormulaR1C1 = "Итого срочных"

    For Each element In Range("L13:L20")
        element.FormulaR1C1 = "=SUM(RC[-9]:RC[-5])"
    Next
    
End Sub
Sub in_work()

    Range("b13:b20").Select
    d = 1
    For Each elements In Selection
        If d = 1 Then
            d1 = """ДЭУ-1"""
            d2 = """ДЭУ-2"""
        End If
        If d = 2 Then
            d1 = """ДЭУ-2"""
            d2 = """ДЭУ-3"""
        End If
        If d = 3 Then
            d1 = """ДЭУ-3"""
            d2 = """ДЭУ-4"""
        End If
        If d = 4 Then
            d1 = """ДЭУ-4"""
            d2 = """ДЭУ-5"""
        End If
        If d = 5 Then
            d1 = """ДЭУ-5"""
            d2 = """Зелёнка"""
        End If
        If d = 6 Then
            d1 = """ДРУ"""
            d2 = """ДЭУ-1"""
        End If
        If d = 7 Then
            d1 = """Зелёнка"""
            d2 = """#Н/Д"""
        End If
        If d = 8 Then
            d1 = """#Н/Д"""
            d2 = """Общий итог"""
        End If
    d = d + 1
    Call in_work_formula(d1, d2)
    ActiveCell.offset(1).Select
    Next

End Sub
Sub in_work_formula(d1, d2)

ActiveCell.FormulaR1C1 = "=COUNT(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1))"

End Sub
Sub is_exist(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)

'Dim DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist As Boolean
pos_table = WorksheetFunction.CountA(Range("M:M"))

For Each element In Range("M2:M" & pos_table & "")
    With element
        If .Value = "ДЭУ-1" Then DEU1_exist = True
        If .Value = "ДЭУ-2" Then DEU2_exist = True
        If .Value = "ДЭУ-3" Then DEU3_exist = True
        If .Value = "ДЭУ-4" Then DEU4_exist = True
        If .Value = "ДЭУ-5" Then DEU5_exist = True
        If .Value = "ДРУ" Then DRU_exist = True
        If .Value = "Зелёнка" Then Zel_exist = True
        If .Value = "#Н/Д" Then na_exist = True
    End With
Next
MsgBox DRU_exist
End Sub
Sub days(i, o, DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, DRU_exist, Zel_exist, na_exist)
    Range(Cells(13, 4 + o), Cells(20, 4 + o)).Select
    d = 1
    For Each element In Selection
    
    If d = 1 Then
        d1 = """ДЭУ-1"""
        d2 = """ДЭУ-2"""
        Call days_formula(i, d1, d2)
        If DEU1_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 2 Then
        d1 = """ДЭУ-2"""
        d2 = """ДЭУ-3"""
        Call days_formula(i, d1, d2)
        If DEU2_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 3 Then
        d1 = """ДЭУ-3"""
        d2 = """ДЭУ-4"""
        Call days_formula(i, d1, d2)
        If DEU3_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 4 Then
        d1 = """ДЭУ-4"""
        d2 = """ДЭУ-5"""
        If DEU5_exist = False Then d2 = """Зелёнка"""
        Call days_formula(i, d1, d2)
        If DEU4_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 5 Then
        d1 = """ДЭУ-5"""
        d2 = """Зелёнка"""
        Call days_formula(i, d1, d2)
        If DEU5_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 6 Then
        d1 = """ДРУ"""
        d2 = """ДЭУ-1"""
        Call days_formula(i, d1, d2)
        If DRU_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 7 Then
        d1 = """Зелёнка"""
        d2 = """#Н/Д"""
        If na_exist = False Then d2 = """Общий итог"""
        Call days_formula(i, d1, d2)
        If Zel_exist = False Then ActiveCell.Value = "0"
    End If
    If d = 8 Then
        d1 = """#Н/Д"""
        d2 = """Общий итог"""
        Call days_formula(i, d1, d2)
        If na_exist = False Then ActiveCell.Value = "0"
    End If
    ActiveCell.offset(1).Select
    d = d + 1
    Next
    ActiveCell.offset(-9).Select
    
End Sub

Sub days_formula(i, d1, d2)

    ActiveCell.FormulaR1C1 = _
    "=COUNTIFS(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),"">" & i & """,INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),""<" & i + 1 & """)"

End Sub

Sub DeuTable(PosStr)

    Sheets.Add.Name = "Просрочки2"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R" & PosStr & "C38", Version:=6).CreatePivotTable TableDestination:= _
        "Просрочки2!R1C13", TableName:="abc", DefaultVersion:=6
    Sheets("Просрочки2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("abc")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("abc").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("abc").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("abc")
        With .PivotFields("ID сообщения")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("ДЭУ")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Проблемная тема")
            .Orientation = xlPageField
            .Position = 1
        End With
    End With

    With ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема")
        ActiveSheet.PivotTables("abc").AddDataField ActiveSheet. _
            PivotTables("abc").PivotFields("Просрочка (срок - текущая дата)") _
            , "Сумма по полю Просрочка (срок - текущая дата)", xlSum
        ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема"). _
            ClearAllFilters
        ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема"). _
            CurrentPage = "(All)"
    End With

End Sub
Sub TableParametrs()
    With ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема")
        .Orientation = xlPageField
        .Position = 1
        For i = 1 To .PivotItems.count - 1
        .PivotItems(.PivotItems(i).Name).Visible = False
    Next i
        Call LoopUntilEmpty("Неубранная городская территория")
        Call LoopUntilEmpty("Неубранная остановка общественного транспорта")
        Call LoopUntilEmpty("Неубранная проезжая часть/тротуар")
        Call LoopUntilEmpty("Неубранная территория у станции метро")
        Call LoopUntilEmpty("Подтопление на проезжей части/тротуаре")
        Call LoopUntilEmpty("Снег и гололед в пешеходных переходах")
        Call LoopUntilEmpty("Снег и гололед на остановке")
        Call LoopUntilEmpty("Снег и гололед на проезжей части/тротуаре")
        Call LoopUntilEmpty("Снег и гололед у входа в метро")
    End With
End Sub

Sub LoopUntilEmpty(Name)

Dim x, NumRows As Integer

Application.ScreenUpdating = False
Sheets("Sheet1").Select
NumRows = Range("E1", Range("E1").End(xlDown)).Rows.count
Range("E1").Select
For x = 1 To NumRows
    If ActiveCell = Name Then
    With Sheets("Просрочки").PivotTables("abc").PivotFields("Проблемная тема")
        .PivotItems(Name).Visible = True
    End With
    Exit For
    'MsgBox "kekw"
    End If
ActiveCell.offset(1, 0).Select
Next
Application.ScreenUpdating = True

End Sub


