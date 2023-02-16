Sub Открыть_сообщение_в_браузере()
    Application.ScreenUpdating = False
    For Each element In Selection
    Call OpenInBrowser
    ActiveCell.offset(1).Select
    Next
    End Sub
    
    Sub OpenInBrowser()
 
    Dim ID, URL, URL1, URL2, Browser As String
    ID = CStr(ActiveCell)
    URL1 = "https://er.mos.ru/ker/admin/issues/update-oiv?id="
    URL2 = "&section=compliant"
    URL = URL1 + ID + URL2
    'Browser = "C:\Program Files\Google\Chrome\Application\Chrome.exe "     'можно и так, но зачем?
    'Shell Browser + URL                                                    '
    ActiveWorkbook.FollowHyperlink Address:=URL
    End Sub
    