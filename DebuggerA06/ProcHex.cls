VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Word Hex Processing
Public WithEvents Text1 As TextBox
Attribute Text1.VB_VarHelpID = -1
Public MAXLEN As Long
'Public ISLOCK As Boolean
Event UpdateAdr(ByVal Address As Long, ByVal DATA As Byte, CancelB As Boolean)


Private Sub Text1_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Or Text1.SelStart = MAXLEN Then KeyAscii = 0: Exit Sub


Dim CHARN As Long
Dim FirstCharN As Long
Dim CharPos As Long 'Character position in Line
CHARN = SendMessage(Text1.hWnd, EM_LINEFROMCHAR, ByVal Text1.SelStart, ByVal 0&) 'Active Line?!
FirstCharN = SendMessage(Text1.hWnd, EM_LINEINDEX, ByVal CHARN, ByVal 0) 'First Char in Line
CharPos = Text1.SelStart - FirstCharN 'Absolute Position in Line


If CharPos < 10 Then Text1.SelStart = FirstCharN + 10: KeyAscii = 0: Exit Sub
If CharPos > 75 Then KeyAscii = 0: Exit Sub
Text1.SelLength = 1

If CharPos >= 60 Then
'Ako je STRING OFFSET
ProcessStr2 CharPos, FirstCharN, KeyAscii
KeyAscii = 0


Else
'Ako je HEX OFFSET
If IsValidK(Chr(KeyAscii)) = 0 Or Text1.SelText = Chr(32) Then KeyAscii = 0: Text1.SelLength = 0: Exit Sub
Dim LNADR As Long 'Adresa iz reda
Dim PRESERVEPOS As Long
PRESERVEPOS = Text1.SelStart
Text1.SelStart = FirstCharN
Text1.SelLength = 8
LNADR = CLng("&H" & Text1.SelText)

Dim LC As String
Text1.SelStart = PRESERVEPOS
Text1.SelLength = 1
LC = Text1.SelText
Text1.SelText = UCase(Chr(KeyAscii))

If ProcessStr(CharPos, LNADR, FirstCharN) Then
Text1.SelStart = PRESERVEPOS
Text1.SelLength = 1
Text1.SelText = LC
End If

KeyAscii = 0

End If


End Sub
Private Sub ProcessStr2(ByVal CharPos As Long, ByVal FirstCharN As Long, ByVal CHRX As Byte)
'SREDILI

Dim LNADR As Long 'Adresa iz reda
Dim PRESERVEPOS As Long
Dim CB As Boolean
PRESERVEPOS = Text1.SelStart
Text1.SelStart = FirstCharN
Text1.SelLength = 8
LNADR = CLng("&H" & Text1.SelText)

'Dodali


If CHRX <= 13 Then Text1.SelStart = PRESERVEPOS + 1: Text1.SelLength = 0: Exit Sub
RaiseEvent UpdateAdr(LNADR + CharPos - 60, CHRX, CB)


If CB Then Text1.SelStart = PRESERVEPOS + 1: Text1.SelLength = 0: Exit Sub


Dim CBB As String
Text1.SelStart = PRESERVEPOS
Text1.SelLength = 1
Text1.SelText = Chr(CHRX)
CBB = Left(Hex(CHRX And &HF0), 1) & Hex(CHRX And &HF)
Text1.SelLength = 0
CharPos = CharPos - 60
Dim RELP As Long
RELP = 10 + (CharPos * 3)
Text1.SelStart = FirstCharN + RELP
Text1.SelLength = 2
Text1.SelText = CBB



Text1.SelStart = PRESERVEPOS + 1
Text1.SelLength = 0
End Sub
Private Function ProcessStr(ByVal CharPos As Long, ByVal ActiveAdr As Long, ByVal FirstCharN As Long) As Boolean
On Error Resume Next
Dim Is2 As Byte
Dim CB As Boolean
Dim PRESERVEPOS As Long
Dim VALUE As Byte
PRESERVEPOS = Text1.SelStart
Is2 = (CharPos - 10) Mod 3
CharPos = Int((CharPos - 10) / 3)
If Is2 = 0 Then
'Prvi Char
Text1.SelStart = Text1.SelStart - 1
Text1.SelLength = 2
VALUE = CByte("&H" & Text1.SelText)
If Err <> 0 Then On Error GoTo 0

RaiseEvent UpdateAdr(ActiveAdr + CharPos, VALUE, CB)

If CB Then Text1.SelStart = PRESERVEPOS: Text1.SelLength = 0: ProcessStr = True: Exit Function


Text1.SelStart = FirstCharN + 60 + CharPos
Text1.SelLength = 1
Text1.SelText = Chr(VALUE)
Text1.SelStart = PRESERVEPOS
Text1.SelLength = 0
Else
'Drugi Char
Text1.SelStart = Text1.SelStart - 2
Text1.SelLength = 2
VALUE = CByte("&H" & Text1.SelText)
RaiseEvent UpdateAdr(ActiveAdr + CharPos, VALUE, CB)

If CB Then Text1.SelStart = PRESERVEPOS: Text1.SelLength = 0: ProcessStr = True: Exit Function

Text1.SelStart = FirstCharN + 60 + CharPos
Text1.SelLength = 1
Text1.SelText = Chr(VALUE)
Text1.SelStart = PRESERVEPOS
Text1.SelLength = 0
End If

End Function







