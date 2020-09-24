VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registers/Stack"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   975
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
      Height          =   1710
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Use Up/Down Key As the Navigator"
      Top             =   2880
      Width           =   5415
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
      Height          =   1710
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Use Up/Down Key As the Navigator"
      Top             =   2880
      Width           =   5295
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
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Columns         =   3
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
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto EBP"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto ESP"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stack"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Call Stack"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registers/Flags"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ShowingTHR As Long
Private OCTX As CONTEXT
Public lastSTT As Long





Private Sub Command2_Click()
Visible = False
End Sub



Private Sub Command4_Click()

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot Change Registers!", vbCritical, "Information": Exit Sub

Form7.ShowTH = ShowingTHR
Form7.Caption = "Registers/Flags for Thread:" & ShowingTHR
Form7.Show 1
End Sub
Private Sub ReCaption(ByVal HT As Long, ByVal ThrS As Long)
If lastSTT <> ThrS Then ReadIT
End Sub
Private Sub CLRSLB()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
End Sub
Public Sub sHwC(ByVal HT As Long, ByVal ThrS As Long)
If HT = 0 Then
Caption = "Register/Stack current Display Thread:" & ShowingTHR & " (Thread not valid)": Exit Sub
CLRSLB
ElseIf ThrS = 1 Then
CLRSLB
Caption = "Register/Stack current Display Thread:" & ShowingTHR & " (Running,cannot display)": Exit Sub
Else
Caption = "Register/Stack Current Display Thread:" & ShowingTHR
End If
End Sub


Public Sub ReadIT()
If ShowingTHR = 0 Then Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTHR, ThrS)

lastSTT = ThrS

sHwC HT, ThrS
If ThrS = 1 Then Exit Sub

OCTX = GetContext(ShowingTHR)
ReadContext List1, OCTX
ReadStackFrame List2, OCTX.Eip, OCTX.Ebp
Read9Stack List3, ActiveStackPosition, ActiveStackPosition, "ESP"
Read9Stack List4, ActiveBasePosition, ActiveBasePosition, "EBP"


End Sub



Private Sub Form_Load()

RemoveX hwnd
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Call SendMessage(List2.hwnd, &H194, ByVal 600, ByVal 0&)

Dim Tabs() As Long
ReDim Tabs(1)

Tabs(0) = 20
Tabs(1) = 50
Call SendMessage(List3.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
Call SendMessage(List4.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
End Sub



Private Sub Label12_Click()
If List3.ListCount = 0 Then Exit Sub
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot do that!", vbCritical, "Information": Exit Sub

ActiveStackPosition = OCTX.Esp
Read9Stack List3, ActiveStackPosition, ActiveStackPosition, "ESP"
End Sub

Private Sub Label13_Click()
If List4.ListCount = 0 Then Exit Sub
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot do that!", vbCritical, "Information": Exit Sub
ActiveBasePosition = OCTX.Ebp
Read9Stack List4, ActiveBasePosition, ActiveBasePosition, "EBP"
End Sub





Private Sub List2_dblClick()
If List2.ListCount = 0 Or List2.ListIndex = -1 Then Exit Sub
Dim RAdr() As String
RAdr = Split(List2.List(List2.ListIndex), vbTab)
DISCOUNT = CLng("&H" & RAdr(0)): NextB = 0
Form16.ReleaseShow 1
Unload Me
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If List3.ListCount = 0 Then Exit Sub

Dim ThrS As Long
Dim HT As Long

If KeyCode = 38 Then
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot Examine!", vbCritical, "Information": Exit Sub

ActiveStackPosition = ActiveStackPosition - 4
Read9Stack List3, ActiveStackPosition, OCTX.Esp, "ESP"
ElseIf KeyCode = 40 Then
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot Examine!", vbCritical, "Information": Exit Sub
ActiveStackPosition = ActiveStackPosition + 4
Read9Stack List3, ActiveStackPosition, OCTX.Esp, "ESP"
End If

End Sub
Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
If List4.ListCount = 0 Then Exit Sub
Dim ThrS As Long
Dim HT As Long

If KeyCode = 38 Then
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot Examine!", vbCritical, "Information": Exit Sub
ActiveBasePosition = ActiveBasePosition - 4
Read9Stack List4, ActiveBasePosition, OCTX.Ebp, "EBP"

ElseIf KeyCode = 40 Then
HT = GetHandleOfThread(ShowingTHR, ThrS)
ReCaption HT, ThrS
If HT = 0 Then MsgBox "This thread isn't valid now (termianted in meanwhile)!", vbInformation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread running! Cannot Examine!", vbCritical, "Information": Exit Sub
ActiveBasePosition = ActiveBasePosition + 4
Read9Stack List4, ActiveBasePosition, OCTX.Ebp, "EBP"
End If

End Sub









