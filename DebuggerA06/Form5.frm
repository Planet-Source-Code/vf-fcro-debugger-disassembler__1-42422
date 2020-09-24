VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Breakpoints"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5370
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Breakpoints"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable Breakpoints"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable Breakpoints"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Goto Breakpoint"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
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
      Width           =   10695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List4_dblClick
End Sub

Private Sub Command2_Click()
Set RTRIGGER = Nothing
RestoreAllOriginalBytes
ISBPDisabled = 1
Form16.Label4.BackColor = &HFF&
End Sub

Private Sub Command3_Click()
Set RTRIGGER = Nothing
RestoreAllBreakPoints
ISBPDisabled = 0
Form16.Label4.BackColor = &HAA00&
End Sub

Private Sub Command4_Click()
NextB = 0
Unload Me
End Sub

Private Sub Command5_Click()
Set RTRIGGER = Nothing
RestoreAllOriginalBytes
Set ACTIVEBREAKPOINTS = Nothing
Command4_Click
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List4.hwnd, &H194, ByVal 1000, ByVal 0&)
Dim Tabs() As Long
ReDim Tabs(3)
Tabs(0) = 20
Tabs(1) = 40
Tabs(2) = 220
Tabs(3) = 330
Call SendMessage(List4.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

Dim u As Long
Dim BTP() As Long
Dim TDx() As Byte
Dim MnM As String
Dim FWDD As Byte
Dim ISVV As Byte

Dim DissS As String
Dim TA As Long
Dim TA2 As Long
List4.Clear
For u = 1 To ACTIVEBREAKPOINTS.count
BTP = ACTIVEBREAKPOINTS.Item(u)
GetDataFromMem BTP(0), TDx, 16
TDx(0) = CByte(BTP(1))
MnM = FindInModules(BTP(0), TA, TA2)
DASM.BaseAddress = BTP(0)
DissS = DASM.DisAssemble(TDx, 0, FWDD, 0, 0, ISVV)
List4.AddItem Hex(BTP(0)) & vbTab & DissS & vbTab & "In Module:" & MnM
Next u


End Sub

Private Sub List4_dblClick()
If List4.ListIndex = -1 Then Exit Sub
Dim Xs() As String
Xs = Split(List4.List(List4.ListIndex), vbTab)
ChoosedAdr = CLng("&H" & Xs(0))
Unload Me
End Sub
