Attribute VB_Name = "Module4"
'Dieser Sourcecode stammt von http://www.VB-fun.de
'und kann frei verwendet werden. F�r eventuell
'auftretende Sch�den wird keine Haftung �bernommen.
'Bei Fehlern oder Fragen einfach eine Mail an: tipps@VB-fun.de
'Ansonsten viel Spa� und Erfolg mit diesem Sourcecode.

Option Explicit

Declare Function SetWindowPos Lib "user32" ( _
  ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


