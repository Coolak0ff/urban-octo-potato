Sub Сделать_по_красоте()
    Application.ScreenUpdating = False
    Call ДЭУ
    Call Добавить_ДЭУ
    Call Pros(PosStr)
    Call DRU
    Call zelyonka
    Call Сделать_табличку
    Call test_massage
    
    Sheets("Таблица").Select
    End Sub
    Sub Добавить_ДЭУ()
    
        Sheets("Sheet1").Select
        PosStr = WorksheetFunction.CountA(Range("A:A"))
        MsgBox PosStr
        Sheets("Sheet1").Select
        Range("G:G").Select
        Selection.Insert Shift:=xlToRight
        ActiveCell.Select
        ActiveCell.FormulaR1C1 = "ДЭУ"
        'ActiveCell.offset(1, 0).Range("A1").Select
        ActiveCell.Range("A2:A" & PosStr & "").Select
        ActiveCell.FormulaR1C1 = "=INDEX(ДЭУ[#Data],MATCH(C6,ДЭУ[Название],0),MATCH(""ДЭУ"",ДЭУ[#Headers],0))"
        Selection.FillDown
        
    End Sub
    Sub Pros(PosStr)
    'Application.ScreenUpdating = False
    
        PosStr = WorksheetFunction.CountA(Range("A:A"))
    
        Range("AJ:AJ").Insert Shift:=xlToRight
        Range("AJ1").FormulaR1C1 = "Просрочка (срок - текущая дата)"
    
    
        Range("AJ2:AJ" & PosStr & "").FormulaR1C1 = "=ROUNDdown(RC[-1]-NOW(),1)"
        Selection.FillDown
        Range("AJ2:AJ" & PosStr & "").NumberFormat = "General"
        
        Columns("AK:AK").Select
        Selection.Insert Shift:=xlToRight
        Range("ak1").FormulaR1C1 = "Просрочено"
    
        For Each element In Range("ak2:ak" & PosStr & "")
        If ActiveCell.offset(, -1) < 0 Then
        ActiveCell.FormulaR1C1 = "Просрочено"
        End If
        ActiveCell.offset(1).Select
        Next
        
    End Sub
    
    Sub Сделать_табличку()
    
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
        
        With ActiveSheet.PivotTables("Остосртированная Таблица").PivotCache
            .RefreshOnFileOpen = False
            .MissingItemsLimit = xlMissingItemsDefault
        End With
        ActiveSheet.PivotTables("Остосртированная Таблица").RepeatAllLabels xlRepeatLabels
        ActiveWorkbook.ShowPivotTableFieldList = True
        
        With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Объект")
            .Orientation = xlRowField
            .Position = 1
        End With
        ActiveSheet.PivotTables("Остосртированная Таблица").AddDataField ActiveSheet. _
            PivotTables("Остосртированная Таблица").PivotFields("ID сообщения"), "Количество по полю ID сообщения", _
            xlCount
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Объект").AutoSort _
            xlDescending, "Количество по полю ID сообщения"
        
        ActiveSheet.PivotTables("Остосртированная Таблица").AddDataField ActiveSheet. _
            PivotTables("Остосртированная Таблица").PivotFields("Просрочено"), _
            "Просрочено ", xlCount
        
        With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ")
           .Orientation = xlRowField
            .Position = 1
        End With
        With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
            .Orientation = xlPageField
            .Position = 1
        End With
            ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields( _
            "Просрочка (срок - текущая дата)").EnableMultiplePageItems = True
        With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Проблемная тема")
            .Orientation = xlPageField
            .Position = 2
        End With
        
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДРУ").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДЭУ-1").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДЭУ-2").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДЭУ-3").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДЭУ-4").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("#N/A").ShowDetail = False
        ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("ДЭУ"). _
            PivotItems("ДЭУ-5").ShowDetail = False
        
        Call Buttons
        Call TableParametrs
         
    End Sub
    Sub Buttons()
    
    start_pos = 920
    
        ActiveSheet.Buttons.Add(start_pos + 80 * 0, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!prosroch"
            .Characters.Text = "Просрочка"
            .Placement = xlFreeFloating
        End With
            ActiveSheet.Buttons.Add(start_pos + 80 * 1, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!all"
            .Characters.Text = "Все"
            .Placement = xlFreeFloating
        End With
            ActiveSheet.Buttons.Add(start_pos + 80 * 2, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!today"
            .Characters.Text = "Сегодня"
            .Placement = xlFreeFloating
        End With
            ActiveSheet.Buttons.Add(start_pos + 80 * 3, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!tomorow"
            .Characters.Text = "Завтра"
            .Placement = xlFreeFloating
        End With
            ActiveSheet.Buttons.Add(start_pos + 80 * 4, 70, 75, 25).Select
        With Selection
            .OnAction = "PERSONAL.XLSB!after_tomorow"
            .Characters.Text = "Послезавтра"
            .Placement = xlFreeFloating
        End With
        
        Range("a1").Select
        
        ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 100, 72, 20).Select
        With Selection
            .Name = "checkbox1"
            .Caption = "С просрочкой"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
        
        ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 125, 72, 20).Select
        With Selection
            .Name = "checkbox2"
            .Caption = "С контролем"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
        
    End Sub
        Sub TableParametrs()
    
                Call LoopUntilEmpty("Некачественное содержание спортивной площадки в парке")
                Call LoopUntilEmpty("Некачественное содержание инфраструктуры в парке")
                Call LoopUntilEmpty("Снег и гололед в парке")
                Call LoopUntilEmpty("Неисправность/некачественное содержание элементов освещения в парке")
                Call LoopUntilEmpty("Некачественное содержание детской площадки в парке")
    
        End Sub
    
    Sub LoopUntilEmpty(Name)
    
    Dim x, NumRows As Integer
    
    Application.ScreenUpdating = False
    Sheets("Sheet1").Select
    NumRows = Range("E1", Range("E1").End(xlDown)).Rows.count
    Range("E1").Select
    For x = 1 To NumRows
        If ActiveCell = Name Then
            With Sheets("Таблица").PivotTables("Остосртированная Таблица").PivotFields("Проблемная тема")
                .PivotItems(Name).Visible = False
            End With
        Exit For
        End If
    ActiveCell.offset(1, 0).Select
    Next
    Application.ScreenUpdating = True
    
    End Sub
    Sub test_massage()
        Sheets("Sheet1").Select
        PosStr = WorksheetFunction.CountA(Range("A:A"))
        d1 = "ДЭУ-1"
        d2 = "ДЭУ-2"
        d3 = "ДЭУ-3"
        d4 = "ДЭУ-4"
        d5 = "ДЭУ-5"
        dr = "ДРУ"
        zl = "Зелёнка"
        'MsgBox PosStr
        With Range("a" & PosStr + 1 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = d1
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
        With Range("a" & PosStr + 2 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = d2
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
        With Range("a" & PosStr + 3 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = d3
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
        With Range("a" & PosStr + 4 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = d4
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
            With Range("a" & PosStr + 5 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = d5
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
        With Range("a" & PosStr + 6 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = dr
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
        With Range("a" & PosStr + 7 & "").Select
            Selection.FormulaR1C1 = "Тестовое сообщение"
            Selection.offset(, 6).FormulaR1C1 = zl
            Selection.offset(, 35).FormulaR1C1 = "999"
        End With
    End Sub
    