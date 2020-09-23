VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Send message API helper"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Beggining? What to do ?"
      Height          =   615
      Left            =   5760
      TabIndex        =   31
      Top             =   0
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   4680
      Width           =   8175
   End
   Begin VB.ListBox List6 
      Height          =   3765
      Left            =   7200
      TabIndex        =   18
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List5 
      Height          =   3765
      Left            =   5880
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   3765
      Left            =   4800
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   3765
      Left            =   3720
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   3765
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7200
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5400
      Width           =   3735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   8400
      TabIndex        =   8
      Top             =   5160
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Msg box"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "input box"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Enabled         =   0   'False
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear list boxes"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   6000
      Width           =   1230
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   5280
      TabIndex        =   12
      Top             =   6120
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3120
      TabIndex        =   13
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ListBox List7 
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "The controls to play with, check code under command4, I picked it from spying"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   5160
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Comments"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "lParam"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   27
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "wParam"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   26
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "C. Hex"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   25
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "C. Num"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   24
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Constant"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MNC_CLOSE = 1
Private Const MNC_EXECUTE = 2
Private Const MNC_IGNORE = 0
Private Const MNC_SELECT = 3
Private Const WM_MENUCOMMAND = &H126
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
Private Const HTSYSMENU = 3
Private Const WM_UNINITMENUPOPUP = &H125


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Dim hSysMenu1
Private Sub Command1_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
End Sub

Private Sub Command2_Click()
'Unload Form1
End
End Sub

Private Sub Command3_Click()
InputBox "M" & Chr(10) & "C"
End Sub

Private Sub Command4_Click()
retval = SendMessage(Text1.hwnd, &H100, 65, 1966081)
retval = SendMessage(Text1.hwnd, &H102, 97, 1966081)
'retval = SendMessage(Text1.hwnd, &H101, 65, -1071775743)
'Text1.Text = CStr(retval)
End Sub

Private Sub Command5_Click()
MsgBox "M" & Chr(10) & "C", , "M"
End Sub






Private Sub Command7_Click()
MsgBox "Click anywhere, resize form, etc - see the effect in list boxes.Note: list boxes are ment to choose something from them. Firs of course clear them."
End Sub

Private Sub Command8_Click()
List7.RemoveItem List7.ListIndex


End Sub

Private Sub Form_Load()
MsgBox "1.What it does ? Spys any control on your form - what messages is receiving. Gives it out all parameters to enable you to repeat action on any other control of same type from code using SendMessage API. Just click somewhere in listboxes and voila code will be done for you." & Chr(10) & Chr(10) & "2.Unfortunately it spys only controls within this app, which is interesting enough, but I'm sure there are smarter programmers out there that will change this to take a peak in others app. What kind of opurtunities that would rise - I don't need to explain to you, lol." & Chr(10) & Chr(10) & "3.The default thing that this app spys is form itself. To change that find this message in code, there are more explanations." & Chr(10) & Chr(10) & "4.This is the first and last release coz. I got tired of it. Anyway happy programming" & Chr(10) & Chr(10) & "P.S. Whay I made this? I was trying to change sys menu of my form at the point of clicking on form's icon.Did not succed. Any ideas are welcome.Kozlicki@yahoo.com"
'MORE EXPLANATIONS
'example
'change Me.hwnd in next line and under command6_click to text1.hwnd and
'app will spy messages send to that control.
'To test that - run app and type some text into text1, click on it, cut the text,  etc
'you can change this to any other hwnd within your app
Set SourceForm = Me
Magic (Me.hwnd)
End Sub
Private Sub Command6_Click()
UnMagic (Me.hwnd)
Command2.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnMagic (Me.hwnd)
End Sub

Private Sub Form_Terminate()

UnMagic (Me.hwnd)
End Sub

Private Sub List1_Click()
Text2.Text = "CodeRetval= SendMessage ( TargetWindow.hwnd, " & List3.List(List1.ListIndex) & "," & List4.List(List4.ListIndex) & "," & List5.List(List5.ListIndex) & ")"
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
List6.ListIndex = List1.ListIndex
End Sub

Private Sub List1_Scroll()
'get list1.topindex
a = SendMessage(List1.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List5.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List6.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub


Private Sub List2_Click()
Text2.Text = "CodeRetval= SendMessage ( TargetWindow.hwnd, " & List3.List(List1.ListIndex) & "," & List4.List(List4.ListIndex) & "," & List5.List(List5.ListIndex) & ")"
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
List6.ListIndex = List2.ListIndex
End Sub

Private Sub List2_Scroll()
'get list2.topindex
a = SendMessage(List2.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List5.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List6.hwnd, LB_SETTOPINDEX, a, ByVal 0&
End Sub

Private Sub List3_Click()
Text2.Text = "CodeRetval= SendMessage ( TargetWindow.hwnd, " & List3.List(List1.ListIndex) & "," & List4.List(List4.ListIndex) & "," & List5.List(List5.ListIndex) & ")"
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
End Sub

Private Sub List3_Scroll()
'get list2.topindex
a = SendMessage(List3.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List5.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List6.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub

Private Sub List4_Click()
Text2.Text = "CodeRetval= SendMessage ( TargetWindow.hwnd, " & List3.List(List1.ListIndex) & "," & List4.List(List4.ListIndex) & "," & List5.List(List5.ListIndex) & ")"
End Sub

Private Sub List4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
End Sub

Private Sub List4_Scroll()
'get list2.topindex
a = SendMessage(List4.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List5.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List6.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub

Private Sub List5_Click()
Text2.Text = "CodeRetval= SendMessage ( TargetWindow.hwnd, " & List3.List(List1.ListIndex) & "," & List4.List(List4.ListIndex) & "," & List5.List(List5.ListIndex) & ")"
End Sub

Private Sub List5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ListIndex = List5.ListIndex
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
List1.ListIndex = List5.ListIndex
End Sub

Private Sub List5_Scroll()
'get list2.topindex
a = SendMessage(List5.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List6.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub
