Sub Сделать_по_красоте()
Application.ScreenUpdating = True
Call ДЭУ
Call Добавить_ДЭУ
Call Pros(PosStr)
Call DRU
Call zelyonka
Call is_PivotItem_exist(DRU_is_exist, DEU1_is_exist, DEU2_is_exist, DEU3_is_exist, DEU4_is_exist, DEU5_is_exist, ZL_is_exist, NA_is_exist)
Call Сделать_табличку(DRU_is_exist, DEU1_is_exist, DEU2_is_exist, DEU3_is_exist, DEU4_is_exist, DEU5_is_exist, ZL_is_exist, NA_is_exist)

Sheets("Таблица").Select
End Sub
Sub Добавить_ДЭУ()

    Sheets("Sheet1").Select
    PosStr = WorksheetFunction.CountA(Range("A:A"))
    For Each element In Range("1:1")
        If element.Text = "Объект" Then
        element.offset(, 1).Select
        End If
    Next
    Selection.Insert Shift:=xlToRight
    With Selection
        .FormulaR1C1 = "ДЭУ"
        .Range("A2:A" & PosStr & "").FormulaR1C1 = "=INDEX(ДЭУ[#Data],MATCH(C6,ДЭУ[Название],0),MATCH(""ДЭУ"",ДЭУ[#Headers],0))"
    End With
    
End Sub
Sub DEU_column()

    For Each element In Range("1:1")
        If element.Text = "ДЭУ" Then
        element.Select
        End If
    Next

End Sub
Sub is_PivotItem_exist(DRU_is_exist, DEU1_is_exist, DEU2_is_exist, DEU3_is_exist, DEU4_is_exist, DEU5_is_exist, ZL_is_exist, NA_is_exist)

PosStr = WorksheetFunction.CountA(Range("A:A"))

Call DEU_column
    
    For Each element In Range(Selection.offset(1), Selection.offset(PosStr - 1))
        If element.Text = "ДРУ" Then DRU_is_exist = True
        If element.Text = "ДЭУ-1" Then DEU1_is_exist = True
        If element.Text = "ДЭУ-2" Then DEU2_is_exist = True
        If element.Text = "ДЭУ-3" Then DEU3_is_exist = True
        If element.Text = "ДЭУ-4" Then DEU4_is_exist = True
        If element.Text = "ДЭУ-5" Then DEU5_is_exist = True
        If element.Text = "Зелёнка" Then ZL_is_exist = True
        If element.Text = "#Н/Д" Then NA_is_exist = True
    Next
    
        
End Sub
Sub Pros(PosStr)

    PosStr = WorksheetFunction.CountA(Range("A:A"))
    
    For Each element In Range("1:1")
        If element.Text = "Регламентный срок подготовки ответа" Then
        element.Select
        End If
    Next
    Selection.Insert Shift:=xlToRight
    Selection.FormulaR1C1 = "Просрочка (срок - текущая дата)"
    With Range(Selection.offset(1), Selection.offset(PosStr - 1))
        .FormulaR1C1 = "=ROUNDdown(RC[-1]-NOW(),1)"
        .FillDown
        .NumberFormat = "General"
    End With
    
    Selection.offset(, 1).Select
    Range(Selection.offset(1), Selection.offset(PosStr - 1)).Insert Shift:=xlToRight
    Selection.FormulaR1C1 = "Просрочено"
    For Each element In Range(Selection.offset(1), Selection.offset(PosStr - 1))
        If element.offset(, -1) < 0 Then
        element.FormulaR1C1 = "Просрочено"
        End If
    Next
    
End Sub

Sub Сделать_табличку(DRU_is_exist, DEU1_is_exist, DEU2_is_exist, DEU3_is_exist, DEU4_is_exist, DEU5_is_exist, ZL_is_exist, NA_is_exist)

PosStr = WorksheetFunction.CountA(Range("A:A"))
excel_version = 6

    Sheets("Sheet1").Select
    Sheets.Add.Name = "Таблица"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R" & PosStr & "C38", Version:=excel_version).CreatePivotTable TableDestination:= _
        "Таблица!R3C1", TableName:="Остосртированная Таблица", DefaultVersion:=excel_version
    Sheets("Таблица").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("Остосртированная Таблица")
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
    
    With ActiveSheet.PivotTables("Остосртированная Таблица")
        .PivotCache.RefreshOnFileOpen = False
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
        .RepeatAllLabels xlRepeatLabels
        .AddDataField ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ID сообщения"), "Количество по полю ID сообщения", xlCount
        .AddDataField ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочено"), "Просрочено ", xlCount
        .PivotFields("Объект").AutoSort xlDescending, "Количество по полю ID сообщения"
    End With

    With ActiveSheet.PivotTables("Остосртированная Таблица")
        With .PivotFields("Просрочка (срок - текущая дата)")
            .Orientation = xlPageField
            .Position = 1
            .EnableMultiplePageItems = True
        End With
        With .PivotFields("Проблемная тема")
            .Orientation = xlPageField
            .Position = 2
        End With
        With .PivotFields("Объект")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("ДЭУ")
            .Orientation = xlRowField
            .Position = 1
            
            If DRU_is_exist = True Then
            .PivotItems("ДРУ").ShowDetail = False
            End If
            
            If DEU1_is_exist Then
            .PivotItems("ДЭУ-1").ShowDetail = False
            End If
            
            If DEU2_is_exist Then
            .PivotItems("ДЭУ-2").ShowDetail = False
            End If
            
            If DEU3_is_exist Then
            .PivotItems("ДЭУ-3").ShowDetail = False
            End If
            
            If DEU4_is_exist Then
            .PivotItems("ДЭУ-4").ShowDetail = False
            End If
            
            If DEU5_is_exist Then
            .PivotItems("ДЭУ-5").ShowDetail = False
            End If
            
            If ZL_is_exist Then
            .PivotItems("Зелёнка").ShowDetail = False
            End If
            
            If NA_is_exist Then
            .PivotItems("#N/A").ShowDetail = False
            End If
        
        End With
            
    End With
         
    Call Buttons
     
End Sub
Sub Buttons()

start_pos = 720

    With ActiveSheet.Buttons
        .Add(start_pos + 80 * 0, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!prosroch"
            .Characters.Text = "Просрочка"
            .Placement = xlFreeFloating
        End With
        .Add(start_pos + 80 * 1, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!all"
            .Characters.Text = "Все"
            .Placement = xlFreeFloating
        End With
        .Add(start_pos + 80 * 2, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!today"
            .Characters.Text = "Сегодня"
            .Placement = xlFreeFloating
        End With
        .Add(start_pos + 80 * 3, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!tomorow"
            .Characters.Text = "Завтра"
            .Placement = xlFreeFloating
        End With
        .Add(start_pos + 80 * 4, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!after_tomorow"
            .Characters.Text = "Послезавтра"
            .Placement = xlFreeFloating
        End With
    End With
    
    Range("a1").Select
    
    checkbox_width = 144

    With ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 100, checkbox_width, 20).Select
        With Selection
            .Name = "checkbox1"
            .Caption = "С просрочкой"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
    With ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 125, checkbox_width, 20).Select
        With Selection
            .Name = "checkbox2"
            .Caption = "С контролем"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
    End With
    End With
    
End Sub
