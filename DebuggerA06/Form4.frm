VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Debugger 0.1V by Vanja Fuckar"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   5370
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Config Debugger"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Run Process"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attach"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Processes"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ListBox List4 
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
      Height          =   4830
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PIP As PROCESS_INFORMATION
Dim SC1 As SECURITY_ATTRIBUTES
Dim SC2 As SECURITY_ATTRIBUTES
Dim SI As STARTUPINFO
Dim CFile() As Byte
Dim AplName() As Byte
Dim AppP() As Byte
Private Sub Command1_Click()
If List4.ListCount = 0 Then MsgBox "Get Process first..", vbCritical, "Info": Exit Sub
If List4.ListIndex = -1 Then MsgBox "Select Process in Available Processes List!", vbCritical, "Error": Exit Sub
Dim THRID As Long
Dim THandle As Long
Dim K As String
K = List4.List(List4.ListIndex)
Dim DX() As String
DX = Split(K, vbTab)
NameOfRunned = ""
IsLoadedProcess = 0
ShowCaption = DX(1)
ActiveMemPos = 0
gBegAdr = 0: gLenAdr = 0
THandle = CreateDebuggerMainThread(CLng(DX(0)), THRID, Traffic)

Erase DX
End Sub

Private Sub Command2_Click()
DestroyWindow Traffic
UnregisterClass "TRAFFIC", 0
Unload Me
End
End Sub

Private Sub Command3_Click()
ReadProcessesForDebugger List4
End Sub



Private Sub Command4_Click()
cd1.ShowOpen
If Len(cd1.Filename) = 0 Then Exit Sub
CFile = cd1.Filename
SC1.nLength = Len(SC1)
SC2.nLength = Len(SC2)
SI.CB = Len(sinfo)
SI.dwFlags = 1
SI.wShowWindow = 1
AppP = StrConv(PathFromName(cd1.Filename) & Chr(0), vbFromUnicode)
AplName = StrConv(vbNullString, vbFromUnicode)

IsLoadedProcess = 0
Dim CLine As String
CLine = InputBox("Enter Command Line", "Request")
If Len(CLine) <> 0 Then CLine = " " & CLine
CFile = StrConv(cd1.Filename & CLine & Chr(0), vbFromUnicode)

Call LoadP(ByVal 0&, ByVal VarPtr(CFile(0)), SC1, SC2, 0, &H1&, ByVal 0&, ByVal VarPtr(AppP(0)), SI)
NameOfRunned = cd1.Filename
ShowCaption = NameOfRunned
Dim THRID As Long
Dim THandle As Long
ActiveMemPos = 0
gBegAdr = 0: gLenAdr = 0
THandle = CreateDebuggerMainThread(ByVal 0&, THRID, Traffic)





End Sub







Private Sub Command5_Click()
Form10.Show 1
End Sub





Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List4.hwnd, &H194, ByVal 1000, ByVal 0&)

End Sub

Private Sub Form_Unload(Cancel As Integer)

UninitDBGEvents
UninstallHook
End Sub
