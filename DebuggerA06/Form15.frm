VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imports Calling For"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   7470
      Left            =   7080
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   7800
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
      Height          =   7470
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calling From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imports Calling"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FRMTYPE As Byte

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub ReadStrs()

Dim MName As String
Dim WHr As Long
Dim IXX() As String
Dim u As Long
List1.Clear
LockWindowUpdate List1.hwnd
For u = 1 To SINDEXESR.count
IXX = SINDEXESR.Item(u)
List1.AddItem Hex(CLng(IXX(1))) & vbTab & IXX(2)
Next u
LockWindowUpdate 0

End Sub

Private Sub ReadExtn()
Dim ExPs As String 'Address To jmp
Dim MName As String
Dim WHr As Long
Dim IXX() As Long
Dim u As Long
List1.Clear
LockWindowUpdate List1.hwnd
For u = 1 To EINDEXESR.count
IXX = EINDEXESR.Item(u)
ExPs = GetFromExportsSearch(FindInModules(IXX(1)), IXX(1))
If Len(ExPs) <> 0 Then
MName = FindInModules(IXX(1))
List1.AddItem Hex(IXX(1)) & vbTab & MName & ":" & ExPs
End If
Next u
LockWindowUpdate 0

End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 35
Call SendMessage(List1.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

Call SendMessage(List1.hwnd, &H194, ByVal 2000, ByVal 0&)

If FRMTYPE = 1 Then
ReadStrs
Label3 = "Address/Strings"
Label1 = "Refers From:"
Else
Label3 = "Imports Calling"
Label1 = "Calling From:"
ReadExtn
End If
End Sub

Private Sub List1_dblClick()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then Exit Sub
Dim Vresp() As String
Dim SPT() As String
Dim SPT2() As String
Dim ExFm As String 'Address From
Dim u As Long
If FRMTYPE = 0 Then
Vresp = Split(List1.List(List1.ListIndex), vbTab)
ExFm = GetFromIndex(EINDEXESR, EREFSR, CLng("&H" & Vresp(0)))
SPT = Split(ExFm, "Jumps From:")
SPT2 = Split(SPT(1), ",")
Erase SPT
List2.Clear
For u = 0 To UBound(SPT2)
List2.AddItem SPT2(u)
Next u

Else
Vresp = Split(List1.List(List1.ListIndex), vbTab)
ExFm = GetFromStringIndex(SINDEXESR, SREFSR, CLng("&H" & Vresp(0)))
SPT = Split(ExFm, "Refs From:")
SPT2 = Split(SPT(1), ",")
Erase SPT
List2.Clear
For u = 0 To UBound(SPT2)
List2.AddItem SPT2(u)
Next u
End If

End Sub

Private Sub List2_dblClick()
ChoosedAdr = CLng("&H" & List2.List(List2.ListIndex))
Unload Me
End Sub
