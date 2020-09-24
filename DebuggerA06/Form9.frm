VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define Search Area"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3060
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Find Last"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Find Last Valid Address From"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Index           =   1
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      Index           =   0
      Left            =   360
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "From"
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Dalje
Dim cx As Long
Dim Cxl As Long
Dim IsValidRange As Long
cx = CLng("&H" & Text3(0))
Cxl = CLng("&H" & Text3(1))

If cx > Cxl Then MsgBox "Error in Length!", vbCritical, "Error": NullV: Exit Sub
gBegAdr = cx
gLenAdr = SubBy8(Cxl, cx)

If gLenAdr > 39452672 Then NullV: MsgBox "Max Search Area: 32MB", vbExclamation, "Information": Exit Sub
GetDataFromMem gBegAdr, DATAPW, gLenAdr, IsValidRange

If IsValidRange = 0 Then MsgBox "Invalid Memory Data! (Entire or some parts of that Range)", vbExclamation, "Information": NullV: Exit Sub
gSTARTADR = gBegAdr: gLASTADR = AddBy8(gBegAdr, gLenAdr)
gSTARTADR2 = gBegAdr: gLASTADR2 = AddBy8(gBegAdr, gLenAdr)
Unload Me
Exit Sub
Dalje:
On Error GoTo 0
NullV
MsgBox "Unknown Value Type!", vbCritical, "Error!"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub NullV()
gBegAdr = 0
gLenAdr = 0
End Sub


Private Sub Command4_Click()
If Len(Text3(0)) = 0 Then MsgBox "From Address isn't set!", vbInformation, "Information": Exit Sub
Dim LPP As Long
Dim IsValid As Byte
Dim TBuff(127) As Byte
LPP = CLng("&H" & Text3(0))
Text3(1) = ""
Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 128, ByVal 0&)
LPP = LPP + 4
Loop While IsValid <> 0
LPP = LPP - 4
Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 1, ByVal 0&)
LPP = LPP + 1
Loop While IsValid <> 0
LPP = LPP - 1
Text3(1) = Hex(LPP)
Erase TBuff
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub
