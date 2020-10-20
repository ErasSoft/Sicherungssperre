VERSION 5.00
Begin VB.Form frm_Ereignis 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "frm_Bildschirmschoner1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
End
Attribute VB_Name = "frm_Ereignis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Länge As Integer

Private Declare Function LockWorkStation Lib "user32.dll" () As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias _
     "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) _
      As Long

Function UserName() As String
  Dim strName   As String
  Dim nSize     As Long
  Dim lngResult As Long

  nSize = 100
  strName = Space$(100)

  lngResult = GetUserName(strName, nSize)
  If lngResult <> 0 Then
    UserName = Left$(strName, nSize - 1)
  End If
End Function

Private Function sperren()
If (Wahl = 0) Then
    MsgBox Aussage1 & UserName & Aussage2, vbCritical + vbOKOnly, "Computer wird gesperrt", , , Zeit
Else
    MsgBox Aussage1 & Aussage2, vbCritical + vbOKOnly, "Computer wird gesperrt", , , Zeit
End If
    Unload Me
    LockWorkStation
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = A_Code) Then
Unload Me
Else
    Call sperren
End If
'Me.Hide
End Sub


Private Sub Form_Load()
    SetWindowPos frm_Ereignis.hwnd, _
    HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    
    Call Mache_Transparent(Me.hwnd, 1)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call sperren
End Sub


