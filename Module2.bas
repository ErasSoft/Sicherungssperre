Attribute VB_Name = "Module2"
Option Explicit

' Benötigte API's für die Timer-Steuerung
Private Declare Function SetTimer Lib "user32" ( _
  ByVal hWnd As Long, _
  ByVal nIDEvent As Long, _
  ByVal uElapse As Long, _
  ByVal lpTimer As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
  ByVal hWnd As Long, _
  ByVal nIDEvent As Long) As Long

Private Const MY_NID = 88
Private Const MY_ELAPSE = 25 ' Wartezeit: 25 MSek.

' Benötigte API's für das Manipulieren der MsgBox
Private Declare Function MessageBox Lib "user32" _
  Alias "MessageBoxA" ( _
  ByVal hWnd As Long, _
  ByVal lpText As String, _
  ByVal lpCaption As String, _
  ByVal wType As Long) As Long

Private Declare Function GetActiveWindow _
  Lib "user32" () As Long

' WindowHandle des aktiven Fensters
Private m_hWnd As Long

' MsgBox OnTop anzeigen
Private m_OnTop As Boolean

' Schließen nach x-Millisekunden
Private m_Time As Long

' Flag für Timer-Ereignis
Private bClose As Boolean

' Benötigte API's für das Anzeigen eines Fenster im Vordergrund
Private Declare Function SetWindowPos Lib "user32" ( _
  ByVal hWnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, _
  ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

' API für Senden eines Fenster-Commands
Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

' Konstante - Fenster schliessen
Private Const WM_CLOSE = &H10

' Unsere neue MsgBox-Funktion
Public Function MsgBox(ByVal sPrompt As String, _
  Optional ByVal nButtons As VbMsgBoxStyle = vbOKOnly, _
  Optional ByVal sTitle As String = "", _
  Optional ByVal sHelpFile As String = "", _
  Optional ByVal nContext As Long = 0, _
  Optional ByVal nTime As Long = 0, _
  Optional ByVal bOnTop As Boolean = False)
  
  Dim nResult As Long
  
  ' Falls MsgBox "OnTop" angezeigt oder nach
  ' x Millisekunden geschlossen werden soll...
  If bOnTop Or nTime > 0 Then
    ' Fensterhandle
    bClose = False
    m_OnTop = True
    m_hWnd = GetActiveWindow()
    m_Time = nTime * 1000

    ' API-Timer starten
    nResult = SetTimer(m_hWnd, MY_NID, MY_ELAPSE, AddressOf MsgBox_TimerEvent)
  
    ' MsgBox anzeigen
    nResult = MessageBox(m_hWnd, sPrompt, sTitle, nButtons)
    
    ' Timer deaktivieren (falls noch aktiviert)
    KillTimer m_hWnd, MY_NID
  Else
    ' andernfalls Standard-MsgBox anzeigen
    nResult = VBA.MsgBox(sPrompt, nButtons, sTitle, sHelpFile, nContext)
  End If

  ' Rückgabewert
  MsgBox = nResult
End Function

' Timer-Event!
Sub MsgBox_TimerEvent()
  Static nWnd As Long
  
  ' API-Timer deaktivieren
  KillTimer m_hWnd, MY_NID
  
  ' MsgBox schließen?
  If bClose Then
    SendMessage nWnd, WM_CLOSE, 0&, 0&
    bClose = False
  Else
  
    ' Fensterhandle der MsgBox
    nWnd = GetActiveWindow()
  
    ' MsgBox On Top setzen
    If m_OnTop Then
      SetWindowPos nWnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE
    End If
  
    ' Timer neu aktivieren
    If m_Time > 0 Then
      bClose = True
      SetTimer m_hWnd, MY_NID, m_Time, AddressOf MsgBox_TimerEvent
    End If
  End If
End Sub



