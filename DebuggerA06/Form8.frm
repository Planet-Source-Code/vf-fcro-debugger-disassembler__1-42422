VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Memory"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form8"
   ScaleHeight     =   7995
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Hide"
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
      Left            =   8160
      TabIndex        =   19
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4920
      Width           =   9495
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Hex Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   4200
      TabIndex        =   13
      Top             =   6480
      Width           =   3855
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
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Find"
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
         Left            =   960
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Reset"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Find At:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Define Search"
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
      Left            =   8160
      TabIndex        =   9
      Top             =   6960
      Width           =   1335
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
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
   End
   Begin VB.VScrollBar vs1 
      Height          =   4695
      Left            =   9240
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox TextX 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "String Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
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
         Left            =   960
         TabIndex        =   11
         Top             =   960
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Find As Unicode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Find At:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searching Area:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      TabIndex        =   10
      Top             =   6120
      Width           =   8055
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Information:"
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
      TabIndex        =   4
      Top             =   4680
      Width           =   9495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto Address:"
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
      Left            =   8160
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents PX As ProcHex
Attribute PX.VB_VarHelpID = -1


Private Sub Command1_Click()
If Len(Text2) = 0 Then MsgBox "Cannot Search An Empty String!", vbCritical, "Error": Exit Sub
If gBegAdr = 0 And gLenAdr = 0 Then MsgBox "Define Searching Area First!", vbInformation, "Information": Exit Sub
If gSTARTADR >= gBegAdr + gLenAdr Then gSTARTADR = gBegAdr: gLASTADR = gBegAdr + gLenAdr

Dim ret As Long
Dim Pattern As String
Pattern = Text2

If Check1.VALUE = 1 Then
ret = Search2(ByVal VarPtr(DATAPW(gSTARTADR - gBegAdr)), gLASTADR - gSTARTADR, ByVal StrPtr(Pattern), ByVal Len(Pattern) * 2)
Else
ret = Search2(ByVal VarPtr(DATAPW(gSTARTADR - gBegAdr)), gLASTADR - gSTARTADR, ByVal Pattern, ByVal Len(Pattern))
End If

If ret = -1 Then
gSTARTADR = gBegAdr: gLASTADR = gSTARTADR + gLenAdr
Label2 = "Not Found"
Else
ActiveMemPos = gSTARTADR + ret - 1
Label2 = "Find At: " & Hex(ActiveMemPos)
ActiveMemPos = Int((ActiveMemPos) / 16&) * 16&
PrintDump TextX, ActiveMemPos


If Check1.VALUE = 1 Then
gSTARTADR = gSTARTADR + ret + Len(Pattern) * 2 - 1
Else
gSTARTADR = gSTARTADR + ret + Len(Pattern) - 1
End If

End If
End Sub

Private Sub Command2_Click()
Form9.Show 1
Label3 = "Searching Area:" & Hex(gBegAdr) & "-" & Hex(gBegAdr + gLenAdr)

End Sub

Private Sub Command3_Click()
gSTARTADR = gBegAdr: gLASTADR = gBegAdr + gLenAdr
End Sub


Private Sub Command4_Click()
gSTARTADR2 = gBegAdr: gLASTADR2 = gBegAdr + gLenAdr
End Sub

Private Sub Command5_Click()
If Len(Text3) = 0 Then MsgBox "Cannot Search An Empty String!", vbCritical, "Error": Exit Sub

If Asc(Right(Text3, 1)) <> 32 Then MsgBox "Incomplete Hex String!", vbCritical, "Error": Exit Sub

If gBegAdr = 0 And gLenAdr = 0 Then MsgBox "Define Searching Area First!", vbInformation, "Information": Exit Sub
If gSTARTADR2 >= gBegAdr + gLenAdr Then gSTARTADR2 = gBegAdr: gLASTADR2 = gBegAdr + gLenAdr

Dim ret As Long
Dim Pattern As String
Pattern = GetTransferString(Text3)

If Len(Pattern) = 0 Then Exit Sub
ret = Search2(ByVal VarPtr(DATAPW(gSTARTADR2 - gBegAdr)), gLASTADR2 - gSTARTADR2, ByVal Pattern, Len(Pattern))

If ret = -1 Then

gSTARTADR2 = gBegAdr: gLASTADR2 = gSTARTADR2 + gLenAdr
Label4 = "Not Found"
Else
ActiveMemPos = gSTARTADR2 + ret - 1
Label4 = "Find At: " & Hex(ActiveMemPos)
ActiveMemPos = Int((ActiveMemPos) / 16&) * 16&
PrintDump TextX, ActiveMemPos

gSTARTADR2 = gSTARTADR2 + ret + Len(Pattern) - 1
End If

End Sub

Private Sub Command6_Click()
Visible = False
End Sub

Private Sub Form_Load()
TextX.FontName = "FixedSys"
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

vs1.Max = 32767
vs1.Min = 0
vs1.VALUE = 16384
vs1.SmallChange = 1
vs1.LargeChange = 10

PrintDump TextX, ActiveMemPos
Set PX = New ProcHex
PX.MAXLEN = Len(TextX)
Set PX.Text1 = TextX


MEMINF = QueryMem(ActiveMemPos, MEMStr)
Text4 = MEMStr

RemoveX hwnd
'RemoveMx hwnd

Label3 = "Searching Area:" & Hex(gBegAdr) & "-" & Hex(gBegAdr + gLenAdr)
End Sub

Private Sub PX_UpdateAdr(ByVal Address As Long, ByVal DATA As Byte, CancelB As Boolean)
Dim IsValidBP As Byte
GetBreakPoint Address, IsValidBP
If IsValidBP = 1 Then
MsgBox "Access denied by Debugger itself!", vbCritical, "Information": CancelB = True: Exit Sub
End If

Dim ret As Long
ret = WriteProcessMemory(ProcessHandle, ByVal Address, DATA, 1, ByVal 0&)

If ret = 0 Then
CancelB = True
Label9.BackColor = &HFF&
Label9 = "Information:  Unable To Update Memory At:" & Hex(Address) & " ,Data:" & Hex(DATA) & " ,String:" & Chr(DATA)
Else
Label9.BackColor = &HAA00&
Label9 = "Information:  Update Memory At:" & Hex(Address) & " ,Data:" & Hex(DATA) & " ,String:" & Chr(DATA)

If Address >= DISCOUNT Then
NextB = 0
Form16.ReleaseShow 0
End If

If Address >= SubBy8(ActiveStackPosition, 12) And Address <= AddBy8(ActiveStackPosition, 16) Then
Read9Stack Form1.List3, ActiveStackPosition, ActiveStackPosition, "ESP"
End If

If Address >= SubBy8(ActiveBasePosition, 12) And Address <= AddBy8(ActiveBasePosition, 16) Then
Read9Stack Form1.List4, ActiveBasePosition, ActiveBasePosition, "EBP"
End If



End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then
Exit Sub
ElseIf KeyAscii = 13 Then
If Len(Text1) = 0 Then Text1 = "": Exit Sub
ActiveMemPos = CLng("&H" & Text1)
ActiveMemPos = Int((ActiveMemPos) / 16&) * 16&
PrintDump TextX, ActiveMemPos
MEMINF = QueryMem(ActiveMemPos, MEMStr)
Text4 = MEMStr
End If

If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Text3 = "": Exit Sub
If ((Text3.SelStart + 1) Mod 3) = 0 Then KeyAscii = 0: Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Text3.SelLength = 1
Text3.SelText = UCase(Chr(KeyAscii))
KeyAscii = 0
If (Len(Text3) + 1) Mod 3 = 0 Then
Text3.SelStart = Text3.SelStart + 1
Text3.SelText = " "
End If
End Sub












Private Sub vs1_Change()
Static IsS As Byte
If IsS = 1 Then IsS = 0: Exit Sub

If vs1.VALUE = 16383 Then

ActiveMemPos = SubBy8(ActiveMemPos, 16)
ElseIf vs1.VALUE = 16385 Then

ActiveMemPos = AddBy8(ActiveMemPos, 16)
ElseIf vs1.VALUE < 16383 Then

ActiveMemPos = SubBy8(ActiveMemPos, 320)
ElseIf vs1.VALUE > 16385 Then

ActiveMemPos = AddBy8(ActiveMemPos, 320)
End If
OutO:
IsS = 1
vs1.VALUE = 16384
PrintDump TextX, ActiveMemPos
MEMINF = QueryMem(ActiveMemPos, MEMStr)
Text4 = MEMStr
End Sub

Private Sub vs1_Scroll()
ReleaseCapture
vs1.VALUE = 16384
End Sub
