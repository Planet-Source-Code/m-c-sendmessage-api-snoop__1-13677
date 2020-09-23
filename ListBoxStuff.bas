Attribute VB_Name = "ListBoxStuff"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETTOPINDEX = &H18E

