Attribute VB_Name = "Operations"
Public pOldProc As Long  ' pointer to Form1's previous window procedure
Public SourceForm As Form  ' pointer to Form1's previous window procedure
Public Sub Magic(hwnd As Long)
'transfer control somewhere else
pOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnMagic(hwnd As Long)
Dim retval As Long  ' return value
' Replace the previous window procedure to prevent crashing.
retval = SetWindowLong(hwnd, GWL_WNDPROC, pOldProc)
End Sub
'The following function acts as Form1's window procedure to process messages.
Public Function WindowProc(ByVal hwnd As Long, ByVal UMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim retval As Long       ' return value
Select Case UMsg

'Text box text entrance*****************************************************************
Case 12 'Public Const WM_SETTEXT = &HC
        GiveMeInfo UMsg, wParam, lParam, "WM_SETTEXT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
'Case 13 'Public Const WM_GETTEXT = &HD
'        GiveMeInfo UMsg, wParam, lParam, "WM_GETTEXT", " "
'        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
'Case 14 'Public Const WM_GETTEXTLENGTH = &HE
'        GiveMeInfo UMsg, wParam, lParam, "WM_GETTEXTLENGTH", " "
'        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 194 'Public Const EM_REPLACESEL = &HC2
        GiveMeInfo UMsg, wParam, lParam, "EM_REPLACESEL", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 256 'Public Const WM_KEYDOWN = &H100
        GiveMeInfo UMsg, wParam, lParam, "WM_KEYDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 257 'Public Const WM_KEYUP = &H101
        GiveMeInfo UMsg, wParam, lParam, "WM_KEYUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 258
        GiveMeInfo UMsg, wParam, lParam, "WM_CHAR", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 770
        GiveMeInfo UMsg, wParam, lParam, "WM_PASTE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 775 'Public Const WM_DESTROYCLIPBOARD = &H307
        GiveMeInfo UMsg, wParam, lParam, "WM_DESTROYCLIPBOARD", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 768 'Public Const WM_CUT = &H300
        GiveMeInfo UMsg, wParam, lParam, "WM_CUT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 769
        GiveMeInfo UMsg, wParam, lParam, "WM_COPY", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
'Text box text entrance*****************************************************************

'dir box *****************************************************************
Case 393 'Public Const LB_GETTEXT = &H189
        GiveMeInfo UMsg, wParam, lParam, "LB_GETTEXT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 384 'Public Const LB_ADDSTRING = &H180
        GiveMeInfo UMsg, wParam, lParam, "LB_ADDSTRING", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
'dir box *****************************************************************

'drive box *****************************************************************
Case 328 'Public Const CB_GETLBTEXT = &H148
        GiveMeInfo UMsg, wParam, lParam, "CB_GETLBTEXT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 336 'Public Const CB_GETITEMDATA = &H150
        GiveMeInfo UMsg, wParam, lParam, "CB_GETITEMDATA", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
'drive box *****************************************************************

Case 278 'before sys menu is displayed
        GiveMeInfo UMsg, wParam, lParam, "WM_INITMENU", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 274
        'item selected in sys menu 'user clicked item in sys menu
        GiveMeInfo UMsg, wParam, lParam, "WM_SYSCOMMAND", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 3
        GiveMeInfo UMsg, wParam, lParam, "WM_MOVE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 5 'end of any form resizing
        GiveMeInfo UMsg, wParam, lParam, "WM_SIZE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 6 'msg box poped up
        GiveMeInfo UMsg, wParam, lParam, "WMSZ_BOTTOM", "Msg box poped up"
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 15 'form resized to max by sizing button in caption bar, WM_PAINT = &HF
        GiveMeInfo UMsg, wParam, lParam, "WM_PAINT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 20
        GiveMeInfo UMsg, wParam, lParam, "WM_ERASEBKGND", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 24
        GiveMeInfo UMsg, wParam, lParam, "WM_SHOWWINDOW", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 32 'mouse moving over vb controls
       'show it only once
       ' If SourceForm.List1.List(SourceForm.List1.ListCount - 3) <> " WM_SETCURSOR = &H20,Mouse moving over vb controls & Form" Then
       ' GiveMeInfo UMsg, wParam, lParam, "WM_SETCURSOR"
       ' End If
       WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 33 'mouse down on Command button
        GiveMeInfo UMsg, wParam, lParam, "WM_MOUSEACTIVATE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 36 'start of manual sizing of form
        GiveMeInfo UMsg, wParam, lParam, "WM_GETMINMAXINFO", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 70 'End of manual sizing of form ?
        GiveMeInfo UMsg, wParam, lParam, "WM_WINDOWPOSCHANGING", "End of manual sizing of form"
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 71 'form resized to normal by sizing button in caption bar
        GiveMeInfo UMsg, wParam, lParam, "WM_WINDOWPOSCHANGED", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 132 'form area mouse presence - countinuous messaging,Mouse present at form area - countinuous
        'show it only once
        'If SourceForm.List1.List(SourceForm.List1.ListCount - 1) <> "WM_NCHITTEST" Then
        'GiveMeInfo UMsg, wParam, lParam, "WM_NCHITTEST", "Mouse present at form area - countinuous "
        'End If
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 133
        GiveMeInfo UMsg, wParam, lParam, "WM_NCPAINT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 134
        GiveMeInfo UMsg, wParam, lParam, "WM_NCACTIVATE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 160 ',Mouse over form caption & menu bar & Border line
         'show it only once
        If SourceForm.List1.List(SourceForm.List1.ListCount - 1) <> "WM_NCMOUSEMOVE" Then
        GiveMeInfo UMsg, wParam, lParam, "WM_NCMOUSEMOVE", "Mouse over form caption & menu bar & Border line"
        End If
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 161 'mouse down on form's caption bar
        GiveMeInfo UMsg, wParam, lParam, "WM_NCLBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 162 'mouse up on form's caption bar
        GiveMeInfo UMsg, wParam, lParam, "WM_NCLBUTTONUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 163
        GiveMeInfo UMsg, wParam, lParam, "WM_NCLBUTTONDBLCLK", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 169
        GiveMeInfo UMsg, wParam, lParam, "WM_NCMBUTTONDBLCLK", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 279 ' WM_INITMENUPOPUP = &H117
        GiveMeInfo UMsg, wParam, lParam, "WM_INITMENUPOPUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 287 'mouse down on form's icon that opens sys menu
        GiveMeInfo UMsg, wParam, lParam, "WM_MENUSELECT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 289 'api menu poped up and beeing poped up countinuous
        'this message is constantly coming in until popup is closed
        'show it only once
        If SourceForm.List1.List(SourceForm.List1.ListCount - 1) <> "WM_ENTERIDLE" Then
        GiveMeInfo UMsg, wParam, lParam, "WM_ENTERIDLE", " "
        End If
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 273 ' WM_COMMAND = &H111
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
        'SourceForm.List1.AddItem " WM_COMMAND = &H111"
Case 307 'each time ..... repainted
        GiveMeInfo UMsg, wParam, lParam, "WM_CTLCOLOREDIT", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 512 ' WM_MOUSEMOVE = &H200mouse moving over vb form
        'GiveMeInfo UMsg, wParam, lParam, "WM_MOUSEMOVE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 513 ' WM_LBUTTONDOWN = &H201mouse down over vb form
        GiveMeInfo UMsg, wParam, lParam, "WM_LBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 514 'mouse up over vb form
        GiveMeInfo UMsg, wParam, lParam, "WM_LBUTTONUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 515 'dbl clickForm Dbl click detected
        GiveMeInfo UMsg, wParam, lParam, "WM_LBUTTONDBLCLK", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 529 'sys menu opened by clicking form icon
        GiveMeInfo UMsg, wParam, lParam, "WM_ENTERMENULOOP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 518 ' WM_RBUTTONDBLCLK = &H206
        GiveMeInfo UMsg, wParam, lParam, "WM_RBUTTONDBLCLK", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 521 ' WM_MBUTTONDBLCLK = &H209
        GiveMeInfo UMsg, wParam, lParam, "WM_MBUTTONDBLCLK", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 530 'sys closed by clicking form icon
        GiveMeInfo UMsg, wParam, lParam, "WM_EXITMENULOOP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 533 ' form area got focus
        GiveMeInfo UMsg, wParam, lParam, "Form area got focus", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 562
        GiveMeInfo UMsg, wParam, lParam, "Unknown", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 309 ' WM_CTLCOLORBTN = &H135
        GiveMeInfo UMsg, wParam, lParam, "WM_CTLCOLORBTN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 519 'middle button down ower form
        GiveMeInfo UMsg, wParam, lParam, "WM_MBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 520 'middle button up ower form
        GiveMeInfo UMsg, wParam, lParam, "WM_MBUTTONUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 516 '"WM_RBUTTONDOWN"
        GiveMeInfo UMsg, wParam, lParam, "WM_RBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 517 ''"WM_RBUTTONUP"
        GiveMeInfo UMsg, wParam, lParam, "WM_RBUTTONUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 131 ' WM_NCCALCSIZE = &H83
        GiveMeInfo UMsg, wParam, lParam, "WM_NCCALCSIZE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 28 ' WM_ACTIVATEAPP = &H1C
        GiveMeInfo UMsg, wParam, lParam, "WM_ACTIVATEAPP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 19 ' WM_QUERYOPEN = &H13
        GiveMeInfo UMsg, wParam, lParam, "WM_QUERYOPEN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 8 ' WM_KILLFOCUS = &H8
        GiveMeInfo UMsg, wParam, lParam, "WM_KILLFOCUS", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 7 ' WM_SETFOCUS = &H7
        GiveMeInfo UMsg, wParam, lParam, "WM_SETFOCUS", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 164 ' WM_NCRBUTTONDOWN = &HA4
        GiveMeInfo UMsg, wParam, lParam, "WM_NCRBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 167 ' WM_NCMBUTTONDOWN = &HA7
        GiveMeInfo UMsg, wParam, lParam, "WM_NCMBUTTONDOWN", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 168 ' WM_NCMBUTTONUP = &HA8
        GiveMeInfo UMsg, wParam, lParam, "WM_NCMBUTTONUP", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 31 ' WM_CANCELMODE = &H1F
        GiveMeInfo UMsg, wParam, lParam, "WM_CANCELMODE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 10 ' WM_ENABLE = &HA
        GiveMeInfo UMsg, wParam, lParam, "WM_ENABLE", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 312 ' WM_CTLCOLORSTATIC = &H138
        GiveMeInfo UMsg, wParam, lParam, "WM_CTLCOLORSTATIC", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 528 ' WM_PARENTNOTIFY = &H210
        GiveMeInfo UMsg, wParam, lParam, "WM_PARENTNOTIFY", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 276 ' WM_HSCROLL = &H114
        GiveMeInfo UMsg, wParam, lParam, "WM_HSCROLL", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 277 ' WM_VSCROLL = &H115
        GiveMeInfo UMsg, wParam, lParam, "WM_VSCROLL", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 43 ' WM_DRAWITEM = &H2B
        GiveMeInfo UMsg, wParam, lParam, "WM_DRAWITEM", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 44 ' WM_MEASUREITEM = &H2C
        GiveMeInfo UMsg, wParam, lParam, "WM_MEASUREITEM", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)
Case 308, 13, 14 'this one causes troubles' constant appearance
WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)

Case Else
'Some unknown
        GiveMeInfo UMsg, wParam, lParam, "Unknown", " "
        WindowProc = CallWindowProc(pOldProc, hwnd, UMsg, wParam, lParam)

End Select
End Function


Public Sub GiveMeInfo(UMsg As Long, wParam As Long, lParam As Long, StrConst As String, Comment As String)
SourceForm.List1.AddItem StrConst
SourceForm.List2.AddItem UMsg 'show it up
SourceForm.List3.AddItem "&H" & Hex(UMsg)
SourceForm.List4.AddItem wParam
SourceForm.List5.AddItem lParam
SourceForm.List6.AddItem Comment

End Sub
