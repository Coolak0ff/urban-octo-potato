Sub Сделать_по_красоте()
Application.ScreenUpdating = False
Call ДЭУ
Call Add_DEU
Call Pros(PosStr)
Call DRU
Call zelyonka
Call Сделать_табличку

Sheets("Таблица").Select
End Sub
Sub Add_DEU()
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
Sub Pros(PosStr)

    PosStr = WorksheetFunction.CountA(Range("A:A"))

    Range("AJ:AJ").Insert Shift:=xlToRight
    Range("AJ1").FormulaR1C1 = "Просрочка (срок - текущая дата)"
    Range("AJ2:AJ" & PosStr & "").Select
    With Selection
        .FormulaR1C1 = "=ROUNDdown(RC[-1]-NOW(),1)"
        .FillDown
        .NumberFormat = "General"
    End With
    
    Columns("AK:AK").Insert Shift:=xlToRight
    Range("ak1").FormulaR1C1 = "Просрочено"
    For Each element In Range("ak2:ak" & PosStr & "")
        If element.offset(, -1) < 0 Then
        element.FormulaR1C1 = "Просрочено"
        End If
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
            .PivotItems("ДРУ").ShowDetail = False
            .PivotItems("ДЭУ-1").ShowDetail = False
            .PivotItems("ДЭУ-2").ShowDetail = False
            .PivotItems("ДЭУ-3").ShowDetail = False
            .PivotItems("ДЭУ-4").ShowDetail = False
            .PivotItems("ДЭУ-5").ShowDetail = False
            .PivotItems("Зелёнка").ShowDetail = False
            '.PivotItems("#N/A").ShowDetail = False
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
    
    With ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 100, 72, 20).Select
        With Selection
            .Name = "checkbox1"
            .Caption = "С просрочкой"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
    With ActiveSheet.CheckBoxes.Add(start_pos + 80 * 1, 125, 72, 20).Select
        With Selection
            .Name = "checkbox2"
            .Caption = "С контролем"
            .OnAction = "PERSONAL.XLSB!All"
            .Placement = xlFreeFloating
        End With
    End With
    End With
    
End Sub
