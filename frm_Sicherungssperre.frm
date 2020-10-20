VERSION 5.00
Begin VB.Form frm_Sicherungssperre 
   BackColor       =   &H0000FF00&
   Caption         =   "Sicherungssperre"
   ClientHeight    =   3135
   ClientLeft      =   3000
   ClientTop       =   2850
   ClientWidth     =   5355
   Icon            =   "frm_Sicherungssperre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5355
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmd_load 
      Caption         =   "Laden"
      Height          =   372
      Left            =   3120
      MouseIcon       =   "frm_Sicherungssperre.frx":030A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "frm_Sicherungssperre.frx":045C
      Height          =   315
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   5040
      Top             =   960
   End
   Begin VB.CheckBox ckb_Auto 
      Caption         =   "Check1"
      Height          =   200
      Left            =   4755
      MouseIcon       =   "frm_Sicherungssperre.frx":0766
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   2
      Top             =   1725
      Width           =   200
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5040
      Top             =   600
   End
   Begin VB.TextBox txt_Auto 
      Alignment       =   1  'Rechts
      Height          =   288
      Left            =   4200
      TabIndex        =   3
      Text            =   "300"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0000FF00&
      Height          =   2535
      Left            =   8400
      TabIndex        =   18
      Top             =   0
      Width           =   2412
      Begin VB.Image img_Eras_Logo 
         Height          =   1065
         Left            =   840
         Picture         =   "frm_Sicherungssperre.frx":08B8
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lbl_copyright_Datum 
         BackStyle       =   0  'Transparent
         Caption         =   "06.11.2011"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl_copyright_Eras 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "alias Eras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl_copyright_Tino 
         BackStyle       =   0  'Transparent
         Caption         =   "Tino Schuldt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl_copyright_by 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lbl_Programm 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Sicherungssperre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lbl_Version 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "v.1.3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Speichern"
      Height          =   372
      Left            =   2160
      MouseIcon       =   "frm_Sicherungssperre.frx":179D
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5040
      Top             =   240
   End
   Begin VB.TextBox txt_Massagebox 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox txt_Zeit 
      Alignment       =   1  'Rechts
      Height          =   288
      Left            =   4200
      TabIndex        =   1
      Text            =   "2"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txt_Aussage2 
      Height          =   288
      Left            =   4200
      TabIndex        =   5
      Text            =   "!"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmd_beenden 
      Caption         =   "Beenden"
      Height          =   372
      Left            =   4320
      MouseIcon       =   "frm_Sicherungssperre.frx":18EF
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmd_neu 
      Caption         =   "Neu"
      Height          =   372
      Left            =   1080
      MouseIcon       =   "frm_Sicherungssperre.frx":1A41
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txt_abbruchzeichen 
      Alignment       =   1  'Rechts
      Height          =   288
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox ckb_Name 
      Caption         =   "Check1"
      Height          =   200
      Left            =   4755
      MouseIcon       =   "frm_Sicherungssperre.frx":1B93
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   12
      Top             =   720
      Value           =   1  'Aktiviert
      Width           =   200
   End
   Begin VB.TextBox txt_Aussage1 
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Text            =   "Du bist nicht "
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmd_ausführen 
      Caption         =   "Ausführen"
      Height          =   372
      Left            =   120
      MouseIcon       =   "frm_Sicherungssperre.frx":1CE5
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "sek."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4965
      TabIndex        =   29
      Top             =   1395
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "sek."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4965
      TabIndex        =   28
      Top             =   1995
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "wenn die Maus nicht mehr bewegt wurde nach:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1980
      Width           =   4095
   End
   Begin VB.Label lbl_Auto 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatisch ausführen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1725
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Angezeigter Text in der Messagebox:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label lbl_Zeit 
      BackStyle       =   0  'Transparent
      Caption         =   "Anzeigedauer der Messagebox:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1365
      Width           =   4095
   End
   Begin VB.Label lbl_Aussage2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dein Text eingeben:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lbl_abbruchzeichen 
      BackStyle       =   0  'Transparent
      Caption         =   "Eine Abbruchtaste eingeben (z.B: a):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1005
      Width           =   4095
   End
   Begin VB.Label lbl_Benutzername 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzername anzeigen?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   4095
   End
   Begin VB.Menu mnu_datei 
      Caption         =   "&Datei"
      Begin VB.Menu mnu_start 
         Caption         =   "Ausführen"
      End
      Begin VB.Menu mnu_ende 
         Caption         =   "Beenden"
      End
   End
End
Attribute VB_Name = "frm_Sicherungssperre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Länge As String

Dim T_Aussage1, T_Aussage2 As String
Dim T_Wahl As Double
Dim Auto_Zeit As Double
Dim Auto_hochzählen As Double
Dim Formular_X, Formular_Y As Double
Public Kanal As Integer
Dim speichername, befehl_header As String

  Dim Prüfsumme As Long
  Dim Daten As String
  Dim lCtr As Long
  Dim DatenASCII() As Byte


Private Declare Function Shell_NotifyIcon Lib "shell32" _
                         Alias "Shell_NotifyIconA" ( _
                         ByVal dwMessage As Long, _
                         ByRef pnid As NOTIFYICONDATA) As Boolean
                         
Private Declare Function SetForegroundWindow Lib "user32" ( _
                         ByVal hwnd As Long) As Long
                         
Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIM_DELETE As Long = &H2&

Private Const NIF_MESSAGE As Long = &H1&
Private Const NIF_ICON As Long = &H2&
Private Const NIF_TIP As Long = &H4&

Private Const WM_MOUSEMOVE As Long = &H200&
Private Const WM_LBUTTONDOWN As Long = &H201&
Private Const WM_LBUTTONUP As Long = &H202&
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_RBUTTONDOWN As Long = &H204&
Private Const WM_RBUTTONUP As Long = &H205&
Private Const WM_RBUTTONDBLCLK As Long = &H206&

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private TIcon As NOTIFYICONDATA



Private Declare Function LockWorkStation Lib "user32.dll" () As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias _
     "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) _
      As Long
      
Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
        
Private Type POINTAPI
  x As Long
  y As Long
End Type

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

Public Function Koordinaten()
Formular_X = frm_Sicherungssperre.Left
Formular_Y = frm_Sicherungssperre.Top
End Function


Public Function Ascii_Code_auslesen()
Prüfsumme = 0
Daten = Code
  DatenASCII = StrConv(Daten, vbFromUnicode)
  For lCtr = LBound(DatenASCII) To UBound(DatenASCII)
    Prüfsumme = (Prüfsumme + DatenASCII(lCtr)) And 255
  Next lCtr
End Function



Private Sub cmd_ausführen_Click()
Call Koordinaten

Aussage1 = txt_Aussage1.Text
Aussage2 = txt_Aussage2.Text
Zeit = txt_Zeit.Text
Code = txt_abbruchzeichen.Text
Länge = Len(Code)
If (Länge > 1) Then
Code = Mid(Code, 1, 1)
End If

Call Ascii_Code_auslesen

A_Code = Prüfsumme
If (ckb_Name = 1) Then
Wahl = 0
Else
Wahl = 1
End If

frm_Ereignis.Show
Me.Hide
End Sub

Private Sub cmd_beenden_Click()
End
End Sub



Private Sub cmd_neu_Click()
txt_Aussage1.Text = "Du bist nicht "
txt_Aussage2.Text = "!"
ckb_Name.Value = 1
txt_abbruchzeichen.Text = ""
txt_Zeit.Text = 2
ckb_Auto.Value = 0
txt_Auto.Enabled = False
txt_Auto.Text = "300"
txt_abbruchzeichen.SetFocus
End Sub

Private Sub cmd_save_Click()
On Error Resume Next

Kanal = FreeFile
Open (speichername) For Output As #Kanal  'hier wird die Datei geöffnet
        
Print #1, "Sicherungssperre Version 1.3.0 - by ErasSoft.de - developer: Tino Schuldt"
Print #1, txt_Aussage1.Text
Print #1, txt_Aussage2.Text
Print #1, ckb_Name.Value
Print #1, txt_abbruchzeichen.Text
Print #1, txt_Zeit.Text
Print #1, ckb_Auto.Value
Print #1, txt_Auto.Text

Close #Kanal
End Sub

Private Sub cmd_load_Click()
On Error GoTo loadfertig
Kanal = FreeFile
Open (speichername) For Input As #Kanal
Input #1, befehl_header

befehl_header = Mid(befehl_header, 1, 16)
If (befehl_header <> "Sicherungssperre") Then
Close #Kanal
Exit Sub
End If

Input #1, befehl_header
txt_Aussage1.Text = befehl_header
Input #1, befehl_header
txt_Aussage2.Text = befehl_header
Input #1, befehl_header
ckb_Name.Value = befehl_header
Input #1, befehl_header
txt_abbruchzeichen.Text = befehl_header
Input #1, befehl_header
txt_Zeit.Text = befehl_header
Input #1, befehl_header
ckb_Auto.Value = befehl_header
Input #1, befehl_header
txt_Auto.Text = befehl_header

loadfertig:
Close #Kanal

End Sub

Private Sub Form_Load()
speichername = "sicherrungssperre.dat"

If Dir(speichername) <> "" Then
cmd_load_Click
End If

    Call Koordinaten

    Me.Hide
    App.TaskVisible = False
    mnu_datei.Visible = False
    
    TIcon.cbSize = Len(TIcon)
    TIcon.hwnd = Picture1.hwnd
    TIcon.uId = 1&
    TIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TIcon.ucallbackMessage = WM_MOUSEMOVE
    TIcon.hIcon = Me.Icon
    'TIcon.szTip = "Was soll ich dazu sagen" & Chr$(0)
    
    ' Hinzufügen des Icons in den Systemtray
    Call Shell_NotifyIcon(NIM_ADD, TIcon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnu_start_Click()
cmd_ausführen_Click
End Sub

Private Sub Timer2_Timer()
T_Aussage1 = txt_Aussage1.Text
T_Aussage2 = txt_Aussage2.Text
If (ckb_Name = 1) Then
T_Wahl = 0
Else
T_Wahl = 1
End If

If (T_Wahl = 0) Then
txt_Massagebox.Text = T_Aussage1 & UserName & T_Aussage2
Else
txt_Massagebox.Text = T_Aussage1 & T_Aussage2
End If

If (ckb_Auto.Value = 0) Then
txt_Auto.Enabled = True
Else
txt_Auto.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If (ckb_Auto.Value = 1) Then
    Auto_Zeit = Val(txt_Auto.Text)
    If (Auto_Zeit = 0) Then
    Else
        If (Auto_hochzählen >= Auto_Zeit) Then
        cmd_ausführen_Click
        End If
    Auto_hochzählen = Auto_hochzählen + 1
    End If
End If
End Sub

Private Sub Timer4_Timer()
  Static done_before As Boolean
  Static CurPosLast As POINTAPI
  Dim CurPosAkt As POINTAPI

    Call GetCursorPos(CurPosAkt)
        
    If (CurPosAkt.x <> CurPosLast.x) Or _
       (CurPosAkt.y <> CurPosLast.y) Then
          Auto_hochzählen = 0                   'Maus wird bewegt
    Else
                                                'Keine Mausbewegung
    End If
        
    CurPosLast = CurPosAkt
End Sub
Private Sub txt_abbruchzeichen_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    'Prüfen ob Enter gedrückt wurde, sonst ignorieren
    cmd_ausführen_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frm_Sicherungssperre.Hide
    
    If UnloadMode = vbAppWindows Or UnloadMode = vbFormCode Then
    
        ' Icon aus dem Systemtray entfernen
        Call Shell_NotifyIcon(NIM_DELETE, TIcon)
    Else
    Call Koordinaten
        Cancel = 1
        
    End If
    
End Sub

Private Sub mnu_ende_Click()

    ' Icon aus dem Systemtray entfernen
    Call Shell_NotifyIcon(NIM_DELETE, TIcon)
    
    Me.Refresh
    End
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As _
    Single, y As Single)
    
    Dim Msg As Long
      
    Msg = x / Screen.TwipsPerPixelX
    
    Select Case Msg
    
    ' Beep
    Case WM_MOUSEMOVE:
    Case WM_LBUTTONDBLCLK: Me.Show
    frm_Sicherungssperre.Left = Formular_X
    frm_Sicherungssperre.Top = Formular_Y
    Case WM_LBUTTONDOWN:
    Case WM_LBUTTONUP:
    Case WM_RBUTTONDBLCLK: Me.Show
    frm_Sicherungssperre.Left = Formular_X
    frm_Sicherungssperre.Top = Formular_Y
    Case WM_RBUTTONDOWN:
    Case WM_RBUTTONUP
    
        ' Diese Funktion muss vor dem anzeigen des
        ' Menüs ausgeführt werden.
        ' weitere Informationen stehen im KB Artikel Q135788 auf
        ' http://support.microsoft.com/kb/q135788/
        Call SetForegroundWindow(Me.hwnd)
        
        ' Menü anzeigen
        Me.PopupMenu mnu_datei
        
        ' bei Verwendung von "TrackPopupMenu" muss noch
        ' die Funktion "PostMessage Me.hwnd, WM_USER, 0&, 0&"
        ' ausgeführt werden
        
    End Select
    
End Sub



