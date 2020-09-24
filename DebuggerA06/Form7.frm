VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Registers/Flags"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   2355
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   13
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   46
      Top             =   1770
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   12
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   44
      Top             =   1770
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   11
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   42
      Top             =   1410
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   10
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   40
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   9
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   38
      Top             =   690
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   8
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   36
      Top             =   1410
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   7
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   34
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   6
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   32
      Top             =   690
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   5
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   30
      Top             =   1410
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   28
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   26
      Top             =   690
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   24
      Top             =   1410
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   22
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   20
      Top             =   690
      Width           =   255
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
      Left            =   4440
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   8
      Left            =   7800
      MaxLength       =   8
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   7
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   6
      Left            =   5880
      MaxLength       =   8
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   5
      Left            =   4920
      MaxLength       =   8
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   4
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   3
      Left            =   3000
      MaxLength       =   8
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   2
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alignment Check"
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
      Height          =   255
      Index           =   22
      Left            =   6600
      TabIndex        =   47
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Virtual Mode"
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
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   45
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resume"
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
      Height          =   255
      Index           =   20
      Left            =   6600
      TabIndex        =   43
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nested"
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
      Height          =   255
      Index           =   19
      Left            =   6600
      TabIndex        =   41
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IOPL"
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
      Height          =   255
      Index           =   18
      Left            =   6600
      TabIndex        =   39
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Overflow"
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
      Height          =   255
      Index           =   17
      Left            =   4440
      TabIndex        =   37
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direction"
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
      Height          =   255
      Index           =   16
      Left            =   4440
      TabIndex        =   35
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interupt"
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
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   33
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trap"
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
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   31
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sign"
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
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   29
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zero"
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
      Height          =   255
      Index           =   12
      Left            =   2280
      TabIndex        =   27
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auxiliary"
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
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parity"
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
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carry"
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
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EBP"
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
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESP"
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
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDI"
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
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESI"
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
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDX"
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
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ECX"
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
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EBX"
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
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EAX"
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
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EIP"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SCTX As CONTEXT
Public ShowTH As Long
Private Sub Command1_Click()
On Error GoTo Dalje
SCTX.Eip = CLng("&H" & Text1(0))
SCTX.Eax = CLng("&H" & Text1(1))
SCTX.Ebx = CLng("&H" & Text1(2))
SCTX.Ecx = CLng("&H" & Text1(3))
SCTX.Edx = CLng("&H" & Text1(4))
SCTX.Esi = CLng("&H" & Text1(5))
SCTX.Edi = CLng("&H" & Text1(6))
SCTX.Esp = CLng("&H" & Text1(7))
SCTX.Ebp = CLng("&H" & Text1(8))

SCTX.EFlags = Text2(0)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(1)) * 4&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(2)) * 16&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(3)) * 64&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(4)) * 128&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(5)) * 256&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(6)) * 512&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(7)) * 1024&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(8)) * 2048&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(9)) * 4096&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(10)) * 8192&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(11)) * 32768)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(12)) * 65536)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(13)) * 131072)
SetContext ShowTH, SCTX


AddLastEip ShowTH, SCTX.Eip
ReadMem ActiveProcess, SCTX.Eip
ReadContext Form1.List1, SCTX
ActiveStackPosition = SCTX.Esp
ActiveBasePosition = SCTX.Ebp
Read9Stack Form1.List3, ActiveStackPosition, ActiveStackPosition, "ESP"
Read9Stack Form1.List4, ActiveBasePosition, ActiveBasePosition, "EBP"

Unload Me

Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Error in Values!", vbCritical, "Error!"
Form_Load
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

SCTX = GetContext(ShowTH)

Text1(0) = Hex(SCTX.Eip)
Text1(1) = Hex(SCTX.Eax)
Text1(2) = Hex(SCTX.Ebx)
Text1(3) = Hex(SCTX.Ecx)
Text1(4) = Hex(SCTX.Edx)
Text1(5) = Hex(SCTX.Esi)
Text1(6) = Hex(SCTX.Edi)
Text1(7) = Hex(SCTX.Esp)
Text1(8) = Hex(SCTX.Ebp)


Text2(0) = (SCTX.EFlags And 1&)
Text2(1) = ((SCTX.EFlags And 4&) / 4&)
Text2(2) = ((SCTX.EFlags And 16&) / 16&)
Text2(3) = ((SCTX.EFlags And 64&) / 64&)
Text2(4) = ((SCTX.EFlags And 128&) / 128&)
Text2(5) = ((SCTX.EFlags And 256&) / 256&)
Text2(6) = ((SCTX.EFlags And 512&) / 512&)
Text2(7) = ((SCTX.EFlags And 1024&) / 1024&)
Text2(8) = ((SCTX.EFlags And 2048&) / 2048&)
Text2(9) = ((SCTX.EFlags And 4096&) / 4096&)
Text2(10) = ((SCTX.EFlags And 8192&) / 8192&)
Text2(11) = ((SCTX.EFlags And 32768) / 32768)
Text2(12) = ((SCTX.EFlags And 65536) / 65536)
Text2(13) = ((SCTX.EFlags And 131072) / 131072)

End Sub





Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 49 Then KeyAscii = 0: Exit Sub
Text2(Index).SelStart = 0
Text2(Index).SelLength = 1
Text2(Index).SelText = Chr(KeyAscii)
End Sub
