VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Console"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "lang2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strin As String
Public ready As Boolean

Private Sub Form_Load()
ready = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If ready Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    KeyAscii = 0
    ready = True
    Exit Sub
End If
If KeyAscii = 8 Then
    If Len(strin) > 0 Then strin = Left(strin, Len(strin) - 1)
Else
    strin = strin & Chr(KeyAscii)
End If
List1.List(List1.ListCount - 1) = "> " & strin

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii
If Text1.Locked Or ready Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    KeyAscii = 0
    ready = True
    Exit Sub
End If
If KeyAscii = 8 Then
    strin = Left(strin, Len(strin) - 1)
Else
    strin = strin & Chr(KeyAscii)
End If

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
For ww = 0 To List1.ListCount - 1
    List1.Selected(ww) = False
Next
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
For ww = 0 To List1.ListCount - 1
    List1.Selected(ww) = False
Next

End Sub
