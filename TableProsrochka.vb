
Sub Табличка_просрочек()
    
    Dim DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist As Boolean
    
    Sheets("sheet1").Select
    PosStr = WorksheetFunction.CountA(Range("A:A"))
    Call DeuTable(PosStr)
    Call is_exist(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)
    Range("D12:K12").Select
    Range("C12:K12").NumberFormat = "dddd"
    
    o = 0
    i = 0
    
    For Each element In Selection
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    If ActiveCell.Text = "суббота" Then
        i = i + 2
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
    If ActiveCell.Text = "воскресенье" Then
        i = i + 1
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
        Call prazdn(i)
        Call days(i, o, DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)
        o = o + 1
        i = i + 1
        ActiveCell.offset(, 1).Select
    Next
    Range("C12:J12").NumberFormat = "dddd dd/mm"
    Call p(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)
    Call n
    Call sum
    Call format
    Range("A12:L21").Select
    Selection.ColumnWidth = 17
    Selection.RowHeight = 25
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub
Sub prazdn(i)

    ActiveCell.NumberFormat = "dd/mm"
    If ActiveCell.Text = "23.02" Then
        i = i + 1
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
    ActiveCell.NumberFormat = "dddd"
    
    ActiveCell.NumberFormat = "dd/mm"
    If ActiveCell.Text = "24.02" Then
        i = i + 1
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
    ActiveCell.NumberFormat = "dddd"

    ActiveCell.NumberFormat = "dd/mm"
    If ActiveCell.Text = "25.02" Then
        i = i + 1
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
    ActiveCell.NumberFormat = "dddd"
    
    ActiveCell.NumberFormat = "dd/mm"
    If ActiveCell.Text = "26.02" Then
        i = i + 1
        ActiveCell.FormulaR1C1 = "=NOW()+ " & i & " "
    End If
    ActiveCell.NumberFormat = "dddd"
  
End Sub
Sub format()

    Range("C12:l12").NumberFormat = "dddd dd/mm"
    Range("A11:l21").Select
    Selection.ColumnWidth = 17
    Selection.RowHeight = 25
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("A11:A21").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("A11:l11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("B12:l12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("A11:l21").Select
    Range("l12").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
    End With
    
    With Range("d13:g21").Interior
        .Color = 4210943
    End With
    
    With Range("C13:C21").Interior
        .Color = 2039807
    End With
    
    With Range("L13:L21").Interior
        .Color = 4210943
    End With
    
    For Each element In Range("c13:g21")
        If element.Value > 0 Then
        element.Font.Bold = True
        End If
    Next
    
    For Each element In Range("l13:l21")
        If element.Value > 0 Then
        element.Font.Bold = True
        End If
    Next
    
    Range("B11:B12").Merge
    Range("a11:a12").Merge
    Range("c11:c12").Merge
    Range("l11:l12").Merge

End Sub
Sub Pros(PosStr)

    Range("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Просрочка (срок - текущая дата)"
    Range("AJ2:AJ" & PosStr & "").Select
    ActiveCell.FormulaR1C1 = "=ROUNDdown(RC[-1]-NOW(),1)"
    Selection.FillDown
    Selection.NumberFormat = "General"
End Sub

Sub p(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)

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
        'If DRU_exist = False Then ActiveCell.Value = "0"
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
Sub n()

    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Всего"
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "ДЭУ-1"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "ДЭУ-2"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "ДЭУ-3"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "ДЭУ-4"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "ДЭУ-5"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "ДРУ"
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Зелёнка"
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "#Н/Д"
    
    i = 8
    For Each element In Range("D11:K11")
         element.FormulaR1C1 = i & " день"
         i = i - 1
    Next


End Sub

Sub pp(d1, d2)

    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-1),""<0"")"

End Sub

Sub sum()
    Range("b12").Select
        ActiveCell.FormulaR1C1 = "Всего в работе"
    Range("b13:b21").Select
        Call in_work
    Range("b21:L21").Select
        For Each elements In Selection
            ActiveCell.FormulaR1C1 = "=SUM(R[-1]C:R[-7]C)"
            ActiveCell.offset(, 1).Select
        Next
    Range("L12").Select
        ActiveCell.FormulaR1C1 = "Итого срочных"
    Range("L13:L20").Select
        For Each element In Selection
            ActiveCell.FormulaR1C1 = "=SUM(RC[-9]:RC[-5])"
            ActiveCell.offset(1).Select
        Next
End Sub

Sub in_work()

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

ActiveCell.FormulaR1C1 = _
        "=COUNT(INDEX(C14,MATCH(" & d1 & ",C13,0)+1):INDEX(C14,MATCH(" & d2 & ",C13,0)-2))"

End Sub
Sub is_exist(DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)

'Dim DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist As Boolean

For Each element In Range("M2:M394") 'заменить на последний номер страки
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
MsgBox DEU1_exist
End Sub
Sub days(i, o, DEU1_exist, DEU2_exist, DEU3_exist, DEU4_exist, DEU5_exist, Zel_exist, na_exist)
    
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
        'If DRU_exist = False Then ActiveCell.Value = "0"
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
    'Call days_formula(i, d1, d2)
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

    Sheets.Add.Name = "Просрочки"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R" & PosStr & "C38", Version:=6).CreatePivotTable TableDestination:= _
        "Просрочки!R1C13", TableName:="abc", DefaultVersion:=6
    Sheets("Просрочки").Select
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
    With ActiveSheet.PivotTables("abc").PivotFields("ID сообщения")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("abc").PivotFields("ДЭУ")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("abc").AddDataField ActiveSheet. _
        PivotTables("abc").PivotFields("Просрочка (срок - текущая дата)") _
        , "Сумма по полю Просрочка (срок - текущая дата)", xlSum
    ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема"). _
        ClearAllFilters
    ActiveSheet.PivotTables("abc").PivotFields("Проблемная тема"). _
        CurrentPage = "(All)"

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
