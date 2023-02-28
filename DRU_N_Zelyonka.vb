Sub zelyonka()
'
' Добавляет Зелёнку в столбец ДЭУ по столбцу проблемная тема
'
PosStr = WorksheetFunction.CountA(Range("A:A"))

park_p1 = "Неисправность/некачественное содержание элементов освещения в парке"
park_p2 = "Некачественное содержание улично-дорожных информационных указателей"
park_p3 = "Ненадлежащий уход за зелеными насаждениями в парке"
park_p4 = "Ненадлежащий уход за зелеными насаждениями на проезжей части/тротуаре"
park_p5 = "Снег и гололед в парке"
park_p6 = "Неубранная парковая территория"
park_p7 = "Некачественное содержание площадки для выгула собак в парке"

Call DEU_column

For Each element In Range(Selection.offset(1), Selection.offset(PosStr - 1))
If _
    element.offset(, -2).Value = park_p1 Or _
    element.offset(, -2).Value = park_p2 Or _
    element.offset(, -2).Value = park_p3 Or _
    element.offset(, -2).Value = park_p4 Or _
    element.offset(, -2).Value = park_p5 Or _
    element.offset(, -2).Value = park_p6 Or _
    element.offset(, -2).Value = park_p7 _
    Then element.Value = "Зелёнка"
Next

End Sub
Sub DRU()

'Добавляет ДРУ в столбец ДЭУ по столбцу проблемная тема

PosStr = WorksheetFunction.CountA(Range("A:A"))

DRU_p1 = "Наличие опасно выступающих элементов на проезжей части/тротуаре"
DRU_p2 = "Некачественное содержание МАФ на проезжей части/тротуаре (скамейки, ограждения, урны и др.)"
DRU_p3 = "Некачественная укладка плитки на проезжей части/тротуаре"
DRU_p4 = "Несвоевременное восстановление благоустройства территории после разрытий"
DRU_p5 = "Нечитаемые дорожные знакие"
DRU_p6 = "Отсутствие или повреждение урн"
DRU_p7 = "Повреждение дорожного покрытия на проезжей части/тротуаре (ямы, выбоины, провалы)"
DRU_p8 = "Повреждение дорожного покры во дворе (ямы, выбоины, провалы)"
DRU_p9 = "Повреждение дорожных ограждений"
DRU_p10 = "Повреждение искусственной дорожной неровности"
DRU_p11 = "Повреждение люка/незакрытый люк на проезжей части/тротуаре"
DRU_p12 = "Ямы на трамвайных путях"
DRU_p13 = "Повреждение дорожного покрытия во дворе (ямы, выбоины, провалы)"
DRU_p14 = "Повреждение уличной лестницы"
DRU_p15 = "Повреждение люка/незакрытый люк во дворе"
DRU_p16 = "Повреждение бордюров на проезжей части/тротуаре"
DRU_p17 = "Разрушение/неправильная укладка тактильной плитки"

Call DEU_column

For Each element In Range(Selection.offset(1), Selection.offset(PosStr - 1))
    If _
    element.offset(, -2).Value = DRU_p1 Or _
    element.offset(, -2).Value = DRU_p2 Or _
    element.offset(, -2).Value = DRU_p3 Or _
    element.offset(, -2).Value = DRU_p4 Or _
    element.offset(, -2).Value = DRU_p5 Or _
    element.offset(, -2).Value = DRU_p6 Or _
    element.offset(, -2).Value = DRU_p7 Or _
    element.offset(, -2).Value = DRU_p8 Or _
    element.offset(, -2).Value = DRU_p9 Or _
    element.offset(, -2).Value = DRU_p10 Or _
    element.offset(, -2).Value = DRU_p11 Or _
    element.offset(, -2).Value = DRU_p12 Or _
    element.offset(, -2).Value = DRU_p13 Or _
    element.offset(, -2).Value = DRU_p14 Or _
    element.offset(, -2).Value = DRU_p15 Or _
    element.offset(, -2).Value = DRU_p16 Or _
    element.offset(, -2).Value = DRU_p17 _
    Then element.Value = "ДРУ"
Next

End Sub

