VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Breakpoint Message"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
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
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TYPEINs As Byte

Private Sub Command2_Click()
InsertIsCancel = 1
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
If TYPEINs = 1 Then
List1.AddItem "WM_KEYDOWN=&H100"
List1.AddItem "WM_KEYUP=&H101"
List1.AddItem "WM_CHAR=&H102"
List1.AddItem "WM_LBUTTONDOWN=&H201"
List1.AddItem "WM_LBUTTONUP=&H202"
List1.AddItem "WM_LBUTTONDBLCLK=&H203"
List1.AddItem "WM_RBUTTONDOWN=&H204"
List1.AddItem "WM_RBUTTONUP=&H205"
List1.AddItem "WM_RBUTTONDBLCLK=&H206"
ElseIf TYPEINs = 2 Then
List1.AddItem "LBN_SELCHANGE =&H1"
List1.AddItem "LBN_DBLCLK=&H2"
List1.AddItem "LBN_SELCANCEL=&H3"
List1.AddItem "LBN_SETFOCUS=&H4"
List1.AddItem "LBN_KILLFOCUS=&H5"
List1.AddItem "CBN_SELCHANGE=&H1"
List1.AddItem "CBN_DBLCLK=&H2"
List1.AddItem "CBN_SETFOCUS=&H3"
List1.AddItem "CBN_KILLFOCUS=&H4"
List1.AddItem "CBN_EDITCHANGE=&H5"
List1.AddItem "CBN_EDITUPDATE=&H6"
List1.AddItem "CBN_DROPDOWN=&H7"
List1.AddItem "CBN_CLOSEUP=&H8"
List1.AddItem "CBN_SELENDOK=&H9"
List1.AddItem "CBN_SELENDCANCEL=&HA"
List1.AddItem "EN_SETFOCUS=&H100"
List1.AddItem "EN_KILLFOCUS=&H200"
List1.AddItem "EN_CHANGE=&H300"
List1.AddItem "EN_UPDATE=&H400"
List1.AddItem "EN_ERRSPACE=&H500"
List1.AddItem "EN_MAXTEXT=&H501"
List1.AddItem "EN_HSCROLL=&H601"
List1.AddItem "EN_VSCROLL=&H602"
List1.AddItem "BN_CLICKED=&H0"
List1.AddItem "BN_PAINT=&H1"
List1.AddItem "BN_HILITE=&H2"
List1.AddItem "BN_UNHILITE=&H3"
List1.AddItem "BN_DISABLE=&H4"
List1.AddItem "BN_DOUBLECLICKED=&H5"
List1.AddItem "BN_SETFOCUS=&H6"
List1.AddItem "BN_KILLFOCUS=&H7"


End If
End Sub

Private Sub List1_Click()
If List1.ListIndex = -1 Then Exit Sub
Dim WlX() As String
WlX = Split(List1.List(List1.ListIndex), "=&H")
InsertVL = WlX(1)
InsertIsCancel = 0
Unload Me
End Sub
