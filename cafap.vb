Sub отфармотировать_цафап()
'
' del_columns Макрос
'

    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    Range("C:E,H:L,O:R,U:U,AA:AC,AE:AE,AG:AL,AN:AP,AR:AT,AV:AX").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("O:O").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "ДЭУ"
    PosStr = WorksheetFunction.CountA(Range("C:C"))
    Range("A2:A" & PosStr & "").FormulaR1C1 = _
        "=INDEX(MasterBook.xlsm!ДЭУ[#Data],MATCH(@C[15],MasterBook.xlsm!ДЭУ[Идентификатор (ID)],0),MATCH(""ДЭУ"",MasterBook.xlsm!ДЭУ[#Headers],0))"
    Call ololo
    Range("B2:B" & PosStr & "").FormulaR1C1 = _
        "=INDEX(MasterBook.xlsm!ДЭУ[#Data],MATCH(@C[14],MasterBook.xlsm!ДЭУ[Идентификатор (ID)],0),MATCH(""Район"",MasterBook.xlsm!ДЭУ[#Headers],0))"
    
    
End Sub
Sub cafap_deu()

PosStr = WorksheetFunction.CountA(Range("C:C"))
    Range("A2:A" & PosStr & "").FormulaR1C1 = _
        "=INDEX(MasterBook.xlsm!ДЭУ[#Data],MATCH(@C[15],MasterBook.xlsm!ДЭУ[Идентификатор (ID)],0),MATCH(""ДЭУ"",MasterBook.xlsm!ДЭУ[#Headers],0))"
    
End Sub

Sub ayes()

Range("Q" & Selection.Row & "").Value = "Выполнено"

End Sub
Sub Макрос10()
'
' Макрос10 Макрос
'

'

    Columns("Q:Q").Select
    Selection.Cut
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("P:P").Select
    Selection.Cut
    Columns("C:C").Select
    ActiveSheet.Paste
End Sub
Sub Макрос23()
'
' Макрос23 Макрос
'

'
    ActiveCell.FormulaR1C1 = _
        "=INDEX(MasterBook.xlsm!ДЭУ[#Data],MATCH(@C[2],MasterBook.xlsm!ДЭУ[Идентификатор (ID)],0),MATCH(""ДЭУ"",MasterBook.xlsm!ДЭУ[#Headers],0))"
End Sub

Sub ololo()
PosStr = WorksheetFunction.CountA(Range("C:C"))
For Each element In Range("P1:P" & PosStr & "")
    element.Value = element.Text
Next

End Sub

Sub open_cafap()
Application.ScreenUpdating = False
For Each element In Selection
Call OpenInBrowser(element)
Next
End Sub

Sub OpenInBrowser(element)
'
' Откыртыть_сообщение_в_браузере Макрос
'
Dim ID, URL, URL1, URL2, Browser As String
URL = element.Text
'Browser = "C:\Program Files\Google\Chrome\Application\Chrome.exe " '
'Shell Browser + URL
ActiveWorkbook.FollowHyperlink Address:=URL
End Sub
Sub Открыть_карту()

For Each element In Selection
    Call maps_code(element, yandex_maps)
    ActiveWorkbook.FollowHyperlink Address:=yandex_maps
Next

End Sub
Sub Выгрузка()

For Each element In Selection
    Call maps_code(element, yandex_maps)
    cafap_pic = Cells(element.Row, 14)
    i = i + 1
    html_file_name = "cafap" & i & ""
    Call html(yandex_maps, cafap_pic, html_file_name)

Next

End Sub
Sub maps_code(element, yandex_maps)

first_cord = Left(Cells(element.Row, 19), InStr(1, Cells(element.Row, 19), ",") - 1)
second_cord = Mid(Cells(element.Row, 19), InStr(1, Cells(element.Row, 19), ",") + 2, 9)

'MsgBox second_cord

Call yandex_url(first_cord, second_cord, yandex_maps)


End Sub
Sub html(yandex_maps, cafap_pic, html_file_name)

'   // Define your variables.
    Dim HTMLFile As String

    HTMLFile = "C:\Users\L1s14\AppData\Roaming\Microsoft\Excel\Random_stuf\" & html_file_name & ".html"
    Close
    'yandex_maps = "https://yandex.ru/maps/213/moscow/?ll=37.708141%2C55.772634&mode=search&sll=37.708141%2C55.772634&text=55.772634%2C37.708141&z=16"
    'cafap_pic = "https://cafap.mos.ru/issue/screenshots?issueNumber=102631257&mode=INITIAL"

'   // Open up the temp HTML file and format the header.
    Open HTMLFile For Output As #1
    
        'Print #1, "<body>"
        Print #1, "<frameset frameborder=0 rows=485,*"; ">"
        Print #1, "<frame src=" & yandex_maps & " name="; topFrame; " scrolling="; no; ">"
        Print #1, "<frameset rows=485,*"; ">"
        Print #1, "<frame src=" & cafap_pic & " name="; mainFrame; ">"
        'Print #1, "</body>"
        
    Close

'ActiveWorkbook.FollowHyperlink Address:=HTMLFile

End Sub
Sub yandex_url(first_cord, second_cord, yandex_maps)

yandex_maps = "https://yandex.ru/maps/213/moscow/?ll=" & second_cord & "%2C" & first_cord & "&mode=search&sll=" & second_cord & "%2C" & first_cord & "&text=" & first_cord & "%2C" & second_cord & "&z=17"

End Sub

Sub for_maps()
MsgBox InStr(1, ActiveCell.offset(, 9), ",")
MsgBox Cells(element.Row, 19).Value
End Sub


