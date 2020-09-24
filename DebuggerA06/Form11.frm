VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Windows in Process"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   ScaleHeight     =   7860
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Insert Value"
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   16
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insert Value"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      MaxLength       =   8
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete BP on Hwnd"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   8040
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Double Click To Remove Breakpoint"
      Top             =   4440
      Width           =   5415
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Double Click To Remove Breakpoint"
      Top             =   4440
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   7440
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click To Examine window"
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WM_Command Breakpoint"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   6960
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WM_Breakpoint"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search with Cursor:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WM_COMMAND Breakpoints Value/ClassName/Hwnd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WM_Breakpoints Value/ClassName/Hwnd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Windows: Class Name / Hwnd / In Thread / Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim THei As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub


Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)

Dim CTbDW As Long
CTbDW = CLng("&H" & Text1)
If Err <> 0 Then On Error GoTo 0: MsgBox "Unknown Value Type", vbCritical, "Error": Exit Sub

If CTbDW = WM_COMMAND Then MsgBox "Cannot Use WM_COMMAND in this context!", vbInformation, "Error": Exit Sub

If CheckWndW(CLng("&H" & SelDt(1))) = 0 Then Exit Sub

Dim iRetB As Byte
iRetB = AddBreakWND(BRKW, CLng("&H" & SelDt(1)), CTbDW, 0)
If iRetB = 0 Then MsgBox "Breakpoint allready exist!", vbInformation, "Information": Exit Sub

ReadWMC BRKW, List2
End Sub

Public Function CheckWndW(ByVal hwnd As Long) As Byte
If IsWindow(hwnd) = 0 Then
MsgBox "In The Meantime,that Window becomes invalid! (destroyed)", vbCritical, "Info"
ReadAllS
Else
CheckWndW = 1
End If
End Function


Private Sub Command3_Click()
On Error Resume Next
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub

Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)

Dim CTbDW As Long
CTbDW = CLng("&H" & Text2)
If Err <> 0 Then On Error GoTo 0: MsgBox "Unknown Value Type", vbCritical, "Error": Exit Sub


If CheckWndW(CLng("&H" & SelDt(1))) = 0 Then Exit Sub


Dim iRetB As Byte
iRetB = AddBreakWND(BRKWMCMD, CLng("&H" & SelDt(1)), CTbDW, 0)
If iRetB = 0 Then MsgBox "Breakpoint allready exist!", vbInformation, "Information": Exit Sub

ReadWMC BRKWMCMD, List3
End Sub








Private Sub Command4_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub

Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)
If CheckWndW(CLng("&H" & SelDt(1))) = 0 Then Exit Sub


Form3.TYPEINs = 1
Form3.Show 1
If InsertIsCancel = 0 Then Text1 = InsertVL: Command2_Click
End Sub

Private Sub Command5_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub

Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)
If CheckWndW(CLng("&H" & SelDt(1))) = 0 Then Exit Sub

Form3.TYPEINs = 2
Form3.Show 1
If InsertIsCancel = 0 Then Text2 = InsertVL: Command3_Click

End Sub






Private Sub Command7_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub

End Sub

Private Sub Command6_Click()
ReadAllS
End Sub

Private Sub Command8_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub
Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)
RemoveEntireWND BRKW, CLng("&H" & SelDt(1))
RemoveEntireWND BRKWMCMD, CLng("&H" & SelDt(1))
ReadAllS
End Sub

Private Sub Form_Load()
RemoveX hwnd

IsF11 = True
THei = Height

Call SendMessage(List1.hwnd, &H194, ByVal 6000, ByVal 0&)
Call SendMessage(List2.hwnd, &H194, ByVal 1300, ByVal 0&)
Call SendMessage(List3.hwnd, &H194, ByVal 1300, ByVal 0&)


Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 170
Tabs(1) = 200
Call SendMessage(List1.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
ReadAllS
End Sub

Public Sub ReadAllS()
Dim WDta() As String
Dim wTextE As String
Dim wlen As Long
List1.Clear
LockWindowUpdate List1.hwnd
Dim u As Long
Dim CwWnd As Long
For u = 1 To WINS.count
WDta = WINS.Item(u)
CwWnd = CLng(WDta(1))
wlen = GetWindowTextLength(CwWnd)
If wlen = 0 Then
wTextE = ""
Else
wTextE = Space(wlen + 1)
If wlen > 256 Then wlen = 256
GetWindowText CwWnd, wTextE, wlen + 1
End If


List1.AddItem WDta(0) & vbTab & Hex(WDta(1)) & vbTab & WDta(2) & vbTab & wTextE
Next u

ReadWMC BRKW, List2
ReadWMC BRKWMCMD, List3
LockWindowUpdate 0
End Sub


Private Sub ReadWMC(ByRef COL As Collection, ByRef LB As ListBox)
LB.Clear
Dim u As Long
Dim i As Long
Dim C As Collection
Dim WBdt() As Long
For u = 1 To COL.count
Set C = COL.Item(u)
For i = 1 To C.count
WBdt = C.Item(i)
LB.AddItem Hex(CStr(WBdt(1))) & vbTab & ClassNameEx(WBdt(0)) & vbTab & Hex(WBdt(0))
Next i
Next u

End Sub


Private Sub Form_Unload(Cancel As Integer)
Form16.Visible = True
IsF11 = False
End Sub

Private Sub Label13_Click(Index As Integer)
If Index = 0 Then
Timer1.Enabled = True

Form16.Visible = False

Height = 700
Else
Timer1.Enabled = False

Form16.Visible = True
Height = THei

OnScreen hwnd
End If

End Sub

Private Sub List1_dblClick()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Window first..", vbInformation, "Required": Exit Sub

Dim SelDt() As String
SelDt = Split(List1.List(List1.ListIndex), vbTab)
If CheckWndW(CLng("&H" & SelDt(1))) = 0 Then Exit Sub
Form14.ACChwnd = CLng("&H" & SelDt(1))
Form14.Show 1
End Sub

Private Sub List2_dblClick()
If List2.ListIndex = -1 Or List2.ListCount = 0 Then Exit Sub
Dim SelDt() As String
Dim LxLn As Long
SelDt = Split(List2.List(List2.ListIndex), vbTab)
RemoveBreakWND BRKW, CLng("&H" & SelDt(2)), CLng("&H" & SelDt(0)), 0
ReadWMC BRKW, List2
End Sub

Private Sub List3_dblClick()
If List3.ListIndex = -1 Or List3.ListCount = 0 Then Exit Sub
Dim SelDt() As String
Dim LxLn As Long
SelDt = Split(List3.List(List3.ListIndex), vbTab)

RemoveBreakWND BRKWMCMD, CLng("&H" & SelDt(2)), CLng("&H" & SelDt(0)), 0
ReadWMC BRKWMCMD, List3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub


Private Sub Timer1_Timer()
Dim CheckThreadId As Long
Dim CheckProcessId As Long
Dim X As Long
Dim Y As Long
GetCursor X, Y
Dim hWndx As Long
hWndx = WindowFromPoint(X, Y)
CheckThreadId = GetWindowThreadProcessId(hWndx, CheckProcessId)
If CheckProcessId = ActiveProcess Then
Label6 = "Search with Cursor:" & ClassNameEx(hWndx) & ",hwnd:" & Hex(hWndx)
SniffedHwnd = hWndx
End If
End Sub
