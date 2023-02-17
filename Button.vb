Sub prosroch()
    Application.ScreenUpdating = False
    Call reset
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)") ' после текущего дня
        For i = 1 To .PivotItems.count - 1
            g = g + 1
        Next i
    End With
    
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems ' до текущего дня
        f = f + 1
        If PivotItem.Name > "0" Then Exit For
    Next PivotItem
    
    L = (g + 1) - (f - 1)
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
    For i = 1 To .PivotItems.count - L
        .PivotItems(.PivotItems(i).Name).Visible = True
        .PivotItems(g + 1).Visible = False
    Next i
    End With
    End Sub
    Sub today()
    Call reset
    Application.ScreenUpdating = False
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems
        f = f + 1
        If PivotItem.Name > "0" Then Exit For
    Next PivotItem
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For i = 1 To .PivotItems.count - 1
            g = g + 1
        Next i
    End With
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
    .PivotItems(f).Visible = True
    .PivotItems(g + 1).Visible = False
    End With
    
    End Sub
    Sub tomorow()
    Application.ScreenUpdating = False
    Call reset
    
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems
        f = f + 1
        If PivotItem.Name > "0" Then Exit For
    Next PivotItem
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For i = 1 To .PivotItems.count - 1
            g = g + 1
        Next i
    End With
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
    .PivotItems(f + 1).Visible = True
    .PivotItems(g + 1).Visible = False
    End With
    
    End Sub
    Sub after_tomorow()
    Application.ScreenUpdating = False
    Call reset
    
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems
        f = f + 1
        If PivotItem.Name > "0" Then Exit For
    Next PivotItem
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For i = 1 To .PivotItems.count - 1
            g = g + 1
        Next i
    End With
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
    .PivotItems(f + 2).Visible = True
    .PivotItems(g + 1).Visible = False
    End With
    
    End Sub
    Sub all()
    Application.ScreenUpdating = False
    Call reset
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)") ' после текущего дня
        For i = 1 To .PivotItems.count - 1
            g = g + 1
        Next i
    End With
    
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems
        f = f + 1
        If PivotItem.Value > "0" Then Exit For
    Next PivotItem
    
    For Each PivotItem In ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)").PivotItems
        x = x + 1
        If PivotItem.Value > "14.5" Then t = t + 1
    Next PivotItem
        b = x - (t + 1)
    
    from = f
    up_to = i - b
    
    ActiveSheet.Shapes("checkbox1").Select
    If Selection.Value = 1 Then ' с просрочкой
        from = 1
    End If
    
    ActiveSheet.Shapes("checkbox2").Select
    If Selection.Value = 1 Then ' с контролем
        up_to = 1
    End If
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For w = from To .PivotItems.count - up_to
           .PivotItems(.PivotItems(w).Name).Visible = True
           If up_to = 1 Then .PivotItems(i).Visible = True _
           Else .PivotItems(i).Visible = False
        Next w
    End With
    
    'Application.CutCopyMode = False
    SendKeys "{ESC}"
    
    End Sub
    Sub reset()
    Application.ScreenUpdating = False
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For w = 1 To .PivotItems.count - 0
           .PivotItems(.PivotItems(w).Name).Visible = True
        Next w
    End With
    
    With ActiveSheet.PivotTables("Остосртированная Таблица").PivotFields("Просрочка (срок - текущая дата)")
        For i = 1 To .PivotItems.count - 1
           .PivotItems(.PivotItems(i).Name).Visible = False
           
        Next i
    End With
    End Sub
    