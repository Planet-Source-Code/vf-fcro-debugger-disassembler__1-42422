VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Imports/Exports"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   8250
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
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
      Height          =   5310
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
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
      Height          =   7470
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   5775
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
      Height          =   1950
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exports Functions / Addresses"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imports By Module:"
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
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imports Modules"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModuleToShow As Long
Private IMPX As Byte
Private EXPX As Byte
Private Sub Command1_Click()
NextB = 0: Unload Me
End Sub



Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hwnd, &H194, ByVal 1000, ByVal 0&)
Call SendMessage(List3.hwnd, &H194, ByVal 1000, ByVal 0&)

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 180
Tabs(1) = 220

Call SendMessage(List1.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
Call SendMessage(List3.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))


ReadPE2 ModuleToShow, IMPX, EXPX
Dim u As Long

List1.Clear
List2.Clear
List3.Clear

If IMPX = 0 Then
List2.AddItem "No Imports!"
Else

For u = 0 To UBound(IMPS)
List2.AddItem IMPS(u).Module

Next u
End If


If EXPX = 0 Then
List3.AddItem "No Exports!"
Else

For u = 0 To UBound(ExPs.FuncNames)
If ExPs.FuncAddress(u) <> 0 Then
'List3.AddItem EXPS.FuncNames(u) & vbTab & "Ord:" & EXPS.Ord(u) & vbTab & Hex(EXPS.FuncAddress(u))
List3.AddItem ExPs.FuncNames(u) & vbTab & Hex(ExPs.FuncAddress(u))
End If
Next u
End If



End Sub

Private Sub List1_dblClick()
On Error GoTo Dalje
If List1.ListCount = 0 Or List1.ListIndex = -1 Then Exit Sub
Dim RLer() As String
RLer = Split(List1.List(List1.ListIndex), vbTab)
Dim GAddress As Long
GAddress = CLng("&H" & RLer(1))
ChoosedAdr = GAddress:  Unload Me
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Private Sub List2_dblClick()
Dim KLI As Long
Dim Xadr As String
KLI = List2.ListIndex
If (List2.ListCount = 1 And List2.List(0) = "No Imports!") Or KLI = -1 Then Exit Sub
Dim i As Long
List1.Clear
Label1 = "Imports By Module:" & IMPS(KLI).Module
For i = 0 To UBound(IMPS(KLI).FuncNames)
'List1.AddItem IMPS(KLI).FuncNames(i) & vbTab & "Ord:" & IMPS(KLI).Ord(i) & vbTab & Hex(IMPS(KLI).CallingAddresses(i))


If TestPTR(IMPS(KLI).CallingAddresses(i)) = 0 Then
Xadr = "Not Loaded Yet By ntdll.dll"
Else
Xadr = Hex(IMPS(KLI).CallingAddresses(i))
End If
List1.AddItem IMPS(KLI).FuncNames(i) & vbTab & Xadr
Next i

End Sub

Private Sub List3_dblClick()
On Error GoTo Dalje
If (List3.ListCount = 1 And List3.List(0) = "No Exports!") Or List3.ListIndex = -1 Then Exit Sub
Dim RLer() As String
RLer = Split(List3.List(List3.ListIndex), vbTab)
Dim GAddress As Long
GAddress = CLng("&H" & RLer(1))
ChoosedAdr = GAddress:  Unload Me
Exit Sub
Dalje:
On Error GoTo 0
End Sub
