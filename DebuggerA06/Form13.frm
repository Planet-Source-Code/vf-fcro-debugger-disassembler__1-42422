VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form13 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Disassemble on File/Cache References"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Find Last"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Find Last Valid Address From"
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cache"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5160
      Width           =   975
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
      Height          =   4590
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   120
      MaxLength       =   8
      TabIndex        =   1
      Top             =   5160
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
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   10
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose From Module Information (Base of Code+Length)"
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
      Width           =   6855
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
      Left            =   120
      TabIndex        =   3
      Top             =   4920
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
      Left            =   1320
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
On Error GoTo Dalje
Dim Start1 As Long
Dim End1 As Long
Dim SLen1 As Long
Dim IsValidRange As Long
Dim IsSvd As Byte
Start1 = CLng("&H" & Text3(0))
End1 = CLng("&H" & Text3(1))

If Start1 >= End1 Then MsgBox "Error in Length!", vbCritical, "Error":  Exit Sub
SLen1 = SubBy8(End1, Start1)
Dim LCLex() As Byte
GetDataFromMem Start1, LCLex, SLen1, IsValidRange
If IsValidRange = 0 Then MsgBox "Invalid Memory Data! (Entire or some parts of that Range)", vbExclamation, "Information": Exit Sub

Label2(2) = "Status:Saving File...."
DoEvents
IsSvd = DoAsm(LCLex, Start1, End1)

If IsSvd = 1 Then MsgBox "File Saved...", vbInformation, "Information": Unload Me
Erase LCLex
Label2(2) = "Status:Done"
Exit Sub
Dalje:
On Error GoTo 0
Label2(2) = "Status:Error"
MsgBox "Unknown Value Type!", vbCritical, "Error!"
End Sub
Private Function DoAsm(ByRef LCL() As Byte, ByVal StartAdr As Long, ByVal LastAdr As Long) As Byte
On Error GoTo Dalje
cd1.ShowSave
If Len(cd1.Filename) = 0 Then MsgBox "Empty File Name isn't allowed!", vbCritical, "Error": Exit Function
Dim FFL As Long
FFL = FreeFile

If Dir(cd1.Filename, vbHidden Or vbReadOnly) <> "" Then
Dir "": Kill cd1.Filename
Else
Dir ""
End If

Open cd1.Filename For Binary As #FFL

DoAsm = GoAsm(FFL, LCL, StartAdr, LastAdr)

Close #FFL
Exit Function
Dalje:
On Error GoTo 0
MsgBox "File Error!", vbCritical, "Error"
Close #FFL
DoAsm = 0
End Function
Private Function GoAsm(ByVal Fnum As Long, ByRef LCL() As Byte, ByVal StartAdr As Long, ByVal LastAdr As Long) As Byte
Dim CMDS As String
Dim AREF As String
Dim IsError As Byte
Dim CRef As String
Dim ExpSt As String
Dim ORGBYTE As Byte
Dim TFWR As Byte
Dim u As Long
Dim IsValidBP As Byte
Dim counterX As Long
Dim DMY As Long
Dim ExlSt As String
Dim IsJmped As String
Dim IsString As Long
DMY = StartAdr

Do


DASM.BaseAddress = DMY

CMDS = DASM.DisAssemble(LCL, counterX, TFWR, 0, 0, IsError)

If NOTIFYVALG = 1 Then
AREF = IsStringOnAdr(IsString)
If IsString = 1 Then AREF = "  (Possible) String: " & AREF
End If


ExpSt = GetFromExportsSearch(FindInModules(StartAdr), StartAdr)
If Len(ExpSt) <> 0 Then ExpSt = "Export:" & ExpSt

IsJmped = GetFromIndex(INDEXESR, REFSR, StartAdr)
If Len(IsJmped) <> 0 Then
Put #Fnum, , "[" & CStr(Hex(StartAdr)) & "] " & IsJmped & vbCrLf
End If

IsJmped = GetFromIndex(EINDEXESR, EREFSR, StartAdr)
If Len(IsJmped) <> 0 Then
Put #Fnum, , "[" & CStr(Hex(StartAdr)) & "] " & IsJmped & vbCrLf
End If

Put #Fnum, , CStr(Hex(StartAdr))
Put #Fnum, , vbTab
Put #Fnum, , CMDS
Put #Fnum, , vbTab
Put #Fnum, , ExpSt


ExlSt = CheckCALL(VALUES1)

If Len(ExlSt) <> 0 Then
Put #Fnum, , ExlSt
Else
Put #Fnum, , AREF
End If


Put #Fnum, , vbCrLf

StartAdr = StartAdr + TFWR
counterX = counterX + TFWR
AREF = ""
Loop While StartAdr < LastAdr

GoAsm = 1
End Function
Private Sub ReadModBySZ(LB As ListBox)
On Error GoTo Dalje
Dim Ftemp1 As Byte
Dim Ftemp2 As Byte
LB.Clear
Dim Dat() As String
Dim u As Long
For u = 1 To ACTMODULESBYPROCESS.count
Dat = ACTMODULESBYPROCESS.Item(u)
ReadPE2 Dat(1), Ftemp1, Ftemp2, 1
LB.AddItem Dat(0) & vbTab & Hex(AddBy8(Dat(1), NTHEADER.OptionalHeader.BaseOfCode)) & vbTab & Hex(NTHEADER.OptionalHeader.SizeOfCode)
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub




Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
ReDim EREFSR(800000)
Set EINDEXESR = Nothing
ReDim REFSR(800000)
Set INDEXESR = Nothing
ReDim SREFSR(800000)
Set SINDEXESR = Nothing

On Error GoTo Dalje
Dim Start1 As Long
Dim End1 As Long
Dim SLen1 As Long
Dim IsValidRange As Long
Dim IsSvd As Byte
Start1 = CLng("&H" & Text3(0))
End1 = CLng("&H" & Text3(1))

If Start1 >= End1 Then MsgBox "Error in Length!", vbCritical, "Error":  Exit Sub
SLen1 = SubBy8(End1, Start1)
Dim LCLex() As Byte
GetDataFromMem Start1, LCLex, SLen1, IsValidRange
If IsValidRange = 0 Then MsgBox "Invalid Memory Data! (Entire or some parts of that Range)", vbExclamation, "Information": Exit Sub

Label2(2) = "Status:Processing References...."
DoEvents


GoCache Start1, End1, LCLex
MsgBox "References cached...", vbInformation, "Information"

Erase LCLex
NextB = 0
Form16.ReleaseShow 1
Label2(2) = "Status:Done"
Exit Sub

Dalje:
On Error GoTo 0
Label2(2) = "Status:Error"
MsgBox "Error during cache References!", vbCritical, "Error!"
Erase EREFSR
Set EINDEXESR = Nothing
Erase REFSR
Set INDEXESR = Nothing
Erase SREFSR
Set SINDEXESR = Nothing

End Sub

Private Sub Command4_Click()
If Len(Text3(0)) = 0 Then MsgBox "From Address isn't set!", vbInformation, "Information": Exit Sub
Dim LPP As Long
Dim IsValid As Byte
Dim TBuff(127) As Byte
LPP = CLng("&H" & Text3(0))
Text3(1) = ""
Label2(2) = "Status:Calculating..."
DoEvents
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
Label2(2) = "Status:Done"
End Sub

Private Sub Form_Load()

Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hwnd, &H194, ByVal 1500, ByVal 0&)
Dim Tabs() As Long
ReDim Tabs(2)
Tabs(0) = 20
Tabs(1) = 150
Tabs(2) = 190
Call SendMessage(List1.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

ReadModBySZ List1
End Sub



Private Sub List1_dblClick()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then Exit Sub
Dim XsDt() As String
XsDt = Split(List1.List(List1.ListIndex), vbTab)

Text3(0) = XsDt(1)
Text3(1) = Hex(AddBy8(CLng("&H" & XsDt(1)), CLng("&H" & XsDt(2))))

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub

Private Function GoCache(ByVal StartAdr As Long, ByVal LastAdr As Long, LCL() As Byte)
Dim IsOut As Long
Dim IsError As Byte
Dim ORGBYTE As Byte
Dim TFWR As Byte

Dim IsValidBP As Byte

Dim DMY As Long
Dim ConterX As Long

Dim INMod As Long 'first Address
Dim INModLast As Long 'Last address
ValidCRef = FindInModules(StartAdr, INMod, INModLast)
INMod = StartAdr
INModLast = LastAdr

CLcnt = 0
DMY = StartAdr
UseCache = 1
Do

DASM.BaseAddress = DMY
CMDS = DASM.DisAssemble(LCL, counterX, TFWR, 0, 0, IsError)


If NOTIFYVALG = 1 Then
IsStringOnAdr , StartAdr, 1
End If



CheckJC StartAdr, INMod, INModLast





StartAdr = StartAdr + TFWR
counterX = counterX + TFWR
Loop While StartAdr < LastAdr

UseCache = 0

End Function


Public Sub CheckJC(ByVal SAddress As Long, ByRef StartAdr As Long, ByRef LastAdr As Long)

Dim LxAddr As Long 'Where JMPS

If NOTIFYJMPCALL = 2 Or NOTIFYJMPCALL = 1 Then


LxAddr = VALUES1
Dim IsValidA As Long
Dim OredTemp() As Byte
GetDataFromMem VALUES1, OredTemp, 16, IsValidA
If IsValidA = 0 Then GoTo Obrada

Dim Ofwr As Byte
DASM.BaseAddress = VALUES1
NewVal = VALUES1
Call DASM.DisAssemble(OredTemp, 0, Ofwr, 0, 0)

If NOTIFYJMPCALL = 4 Then GoTo InNtf2


ElseIf NOTIFYJMPCALL = 3 Or NOTIFYJMPCALL = 4 Then
'CALL DWORD [ADR],JMP DWORD [ADR]
InNtf2:
Call ReadProcessMemory(ProcessHandle, ByVal VALUES1, LxAddr, 4, ByVal 0&)


ElseIf NOTIFYJMPCALL = 8 Then
LxAddr = VALUES1

ElseIf NOTIFYJMPCALL = 5 Then
Call ReadProcessMemory(ProcessHandle, ByVal VALUES1, LxAddr, 4, ByVal 0&)
If LxAddr = 0 Then Exit Sub

Dim TName As String

TName = FindInModules(LxAddr)
If Len(TName) = 0 Then Exit Sub

If Len(GetFromExportsSearch(TName, LxAddr)) = 0 Then Exit Sub


Else
Exit Sub

End If

Obrada:
If LxAddr = 0 Then Exit Sub

If LxAddr >= StartAdr And LxAddr <= LastAdr Then
'Add as Internals
AddInIndex INDEXESR, REFSR, SAddress, LxAddr
Else
'Add as Externals
AddInIndex EINDEXESR, EREFSR, SAddress, LxAddr
End If

End Sub


