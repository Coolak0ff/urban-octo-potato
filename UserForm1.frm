VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DEU1_Click()
Selection.Value = "ДЭУ-1"
Unload Me
End Sub

Private Sub DEU2_Click()
Selection.Value = "ДЭУ-2"
Unload Me
End Sub

Private Sub DEU3_Click()
Selection.Value = "ДЭУ-3"
Unload Me
End Sub

Private Sub DEU4_Click()
Selection.Value = "ДЭУ-4"
Unload Me
End Sub

Private Sub DEU5_Click()
Selection.Value = "ДЭУ-5"
Unload Me
End Sub

Private Sub Skip_Click()
Unload Me
End Sub

Public Sub UserForm_Initialize()

Me.Label1.Caption = "К какому ДЭУ относиться " & Selection.offset(, -1) & "" & vbCrLf & "" & vbCrLf & "Проблемная тема: " & Selection.offset(, -2).Text & ""

'Me.ComboBox1.List = Array("ДЭУ-1", "ДЭУ-2", "ДЭУ-3", "ДЭУ-4", "ДЭУ-5")

End Sub
