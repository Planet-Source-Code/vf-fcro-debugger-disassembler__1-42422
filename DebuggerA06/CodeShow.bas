Attribute VB_Name = "CodeShow"
Public DebuggyOut As Long 'Query Window Proc Procedure
Public AccThreadX As Long 'Created Remote Thread
Public SusThreadX As Long 'Terminate Thread Procedure

Public MainPThread As Long 'Which one is main thread?

Public UseCache As Byte 'Signal to Disasm
Public ValidCRef As String 'Reference Valid For Module!

Public INDEXESR As New Collection 'Internals
Public REFSR() As New Collection

Public EINDEXESR As New Collection 'Externals
Public EREFSR() As New Collection

Public SINDEXESR As New Collection 'STRINGS
Public SREFSR() As New Collection

Public ChoosedAdr As Long

Public X1 As String
Public X2 As String
Public X3 As String


Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const HCBT_ACTIVATE = 5
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_CREATEWND = 3


Public Const WM_DESTROY = &H2
Public Const WM_CLOSE = &H10
Public Const WM_INITDIALOG = &H110
Public Const WM_SIZE = &H5
Public Const WM_SETREDRAW = &HB
Public Const WM_SIZING = &H214
Public Const WM_ACTIVATE = &H6
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_PAINT = &HF
Public Const WM_NCPAINT = &H85
Public Const WM_ERASEBKGND = &H14
Public Const WM_DRAWITEM = &H2B
Public Const WM_SETTEXT = &HC
Public Const WM_SETICON = &H80
Public Const WM_SETFONT = &H30
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CHAR = &H102
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_NOTIFY = &H4E
Public Const WM_COMMAND = &H111
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_INITMENU = &H116
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSKEYDOWN = &H104




Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public ActiveMemPos As Long
Public NotifyData1(29) As Long 'Real Notify flag
Public NotifyData2(29) As Long 'Values1
Public NotifyData3(29) As Long 'Values2
Public NotifyData4(29) As Long 'Values3
Public Function GetHexDump(ByVal Position As Long) As String

GetHexDump = Space(76)


Mid(GetHexDump, 1, 1) = Hex((Position) And &HF0000000)
Mid(GetHexDump, 2, 1) = Hex((Position) And &HF000000)
Mid(GetHexDump, 3, 1) = Hex((Position) And &HF00000)
Mid(GetHexDump, 4, 1) = Hex((Position) And &HF0000)
Mid(GetHexDump, 5, 1) = Hex((Position) And &HF000&)
Mid(GetHexDump, 6, 1) = Hex((Position) And &HF00&)
Mid(GetHexDump, 7, 1) = Hex((Position) And &HF0&)
Mid(GetHexDump, 8, 1) = Hex((Position) And &HF&)

Dim plc As Integer
Dim u As Long
Dim BytD As Byte
Dim ResZ As Byte
Dim Cnnt As Byte
Cnnt = 61
plc = 11
mmax = 16

For u = 0 To 15
ResZ = TestPTR(AddBy8(Position, u), BytD)

If ResZ = 0 Then

Mid(GetHexDump, plc, 2) = "??"
Mid(GetHexDump, Cnnt, 1) = "?"

Else

'Ako je ispravna mem lokacija
Mid(GetHexDump, plc, 1) = Hex(BytD And &HF0)
Mid(GetHexDump, plc + 1, 1) = Hex(BytD And &HF)

If BytD > 13 Then
Mid(GetHexDump, Cnnt, 1) = Chr$(BytD)
Else
Mid(GetHexDump, Cnnt, 1) = Chr(Asc("."))
End If

End If

Cnnt = Cnnt + 1
plc = plc + 3
Next u


End Function

Public Sub PrintDump(ByVal TXT As TextBox, ByVal Position As Long)
Dim u As Long
Dim DaX(19) As String
For u = 0 To 19
DaX(u) = GetHexDump(AddBy8(Position, (u * 16&))) & vbCrLf
Next u
TXT = Join(DaX, "")
End Sub

Public Function QueryMem(ByVal Address As Long, ByRef MemSData As String) As MEMORY_BASIC_INFORMATION
VirtualQueryEx ProcessHandle, ByVal Address, QueryMem, Len(QueryMem)
If QueryMem.AllocationBase = 0 Then
MemSData = "Invalid Memory Region"
Else
Dim MsDt(3) As String



MsDt(0) = "Test At Address:" & Hex(Address) & ",Base Address:" & Hex(QueryMem.BaseAddress) & ",Region Length:" & Hex(QueryMem.RegionSize)


MsDt(1) = "Protection:"
If (QueryMem.AllocationProtect And PAGE_GUARD) = PAGE_GUARD Then MsDt(1) = MsDt(1) & "GUARD "
If (QueryMem.AllocationProtect And PAGE_NOACCESS) = PAGE_NOACCESS Then MsDt(1) = MsDt(1) & "NOACCESS "
If (QueryMem.AllocationProtect And PAGE_NOCACHE) = PAGE_NOCACHE Then MsDt(1) = MsDt(1) & "NOCACHE "
If (QueryMem.AllocationProtect And PAGE_READONLY) = PAGE_READONLY Then MsDt(1) = MsDt(1) & "READONLY "
If (QueryMem.AllocationProtect And PAGE_READWRITE) = PAGE_READWRITE Then MsDt(1) = MsDt(1) & "READWRITE "
If (QueryMem.AllocationProtect And PAGE_WRITECOPY) = PAGE_WRITECOPY Then MsDt(1) = MsDt(1) & "WRITECOPY "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_READWRITE) = PAGE_EXECUTE_READWRITE Then MsDt(1) = MsDt(1) & "EXECUTEREADWRITE "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_READ) = PAGE_EXECUTE_READ Then MsDt(1) = MsDt(1) & "EXECUTEREAD "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_WRITECOPY) = PAGE_EXECUTE_WRITECOPY Then MsDt(1) = MsDt(1) & "EXECUTEWRITECOPY "

MsDt(2) = "State:"
If (QueryMem.State And MEM_COMMIT) = MEM_COMMIT Then MsDt(2) = MsDt(2) & "COMMIT "
If (QueryMem.State And MEM_RESERVE) = MEM_RESERVE Then MsDt(2) = MsDt(2) & "RESERVE "
If (QueryMem.State And MEM_RELEASE) = MEM_RELEASE Then MsDt(2) = MsDt(2) & "RELEASE "

MsDt(3) = "Type:"
If (QueryMem.lType And MEM_MAPPED) = MEM_MAPPED Then MsDt(3) = MsDt(3) & "MAPPED "
If (QueryMem.lType And MEM_IMAGE) = MEM_IMAGE Then MsDt(3) = MsDt(3) & "IMAGE "
If (QueryMem.lType And MEM_PRIVATE) = MEM_PRIVATE Then MsDt(3) = MsDt(3) & "PRIVATE "


MemSData = Join(MsDt, vbCrLf)

End If


End Function


Public Sub PlayS()
Beep 595, 3
Beep 2900, 1
Beep 11, 2
End Sub
Public Function IsStringOnAdr(Optional ByRef IsString As Long, Optional ByVal AddressFrom As Long, Optional ByVal ToCache As Byte) As String
Dim AaX As Long
If VALUES1 <> 0 Then
AaX = VALUES1
ElseIf VALUES3 <> 0 Then
AaX = VALUES3
ElseIf VALUES2 <> 0 Then
AaX = VALUES2
End If
IsString = 0
If AaX = 0 Then Exit Function

Dim IsWp As Long
Dim XYData() As Byte
Dim IsValidPRef As Long
GetDataFromMem SubBy8(AaX, 4), XYData, 260, IsWp


If IsWp = 0 Then Exit Function

Dim ret As Long

Dim TMaxLen As Long
TMaxLen = 256
Dim LenStrX As Long
Dim CCLen As Long
CopyMemory LenStrX, XYData(0), 4
CCLen = lstrlenW(XYData(4))


If LenStrX <= 0 Or LenStrX > 65535 Then GoTo InAnsi
If LenStrX > 256 Then CopyMemory XYData(0), TMaxLen, 4: LenStrX = 256


If (CCLen * 2) = LenStrX Then
'Pravi Unicode
IsStringOnAdr = Space(LenStrX / 2)
CopyMemory ByVal StrPtr(IsStringOnAdr), XYData(4), LenStrX
IsString = 1

If ToCache = 1 Then
AddInStringIndex SINDEXESR, SREFSR, AddressFrom, AaX, IsStringOnAdr
End If


Exit Function
Else
InAnsi:
XYData(259) = 0
Dim AClen As Long

AClen = lstrlen(XYData(4))
If AClen = 0 Then Exit Function
Dim u As Long
For u = 1 To AClen

If IsCharVD(XYData(3 + u)) = 0 Then Exit Function

Next u

'Ansi String!
IsStringOnAdr = Space(AClen)
CopyMemory ByVal IsStringOnAdr, XYData(4), AClen
IsString = 1

If ToCache = 1 Then
AddInStringIndex SINDEXESR, SREFSR, AddressFrom, AaX, IsStringOnAdr
End If

Exit Function
End If


IsString = 0
End Function
Public Function IsCharVD(ByRef BT As Byte) As Byte
Select Case BT
Case Is = 9, 13, 32 To 128
IsCharVD = 1
End Select
End Function


Public Sub RemoveX(ByVal hwnd As Long)
Const SC_CLOSE = &HF060
Dim hMenu As Long
hMenu = GetSystemMenu(hwnd, 0&)
If hMenu Then
Call DeleteMenu(hMenu, SC_CLOSE, 0)
DrawMenuBar (hwnd)
End If
CloseHandle hMenu
End Sub
Public Sub OnScreen(ByVal hwnd As Long)
ReleaseCapture
SendMessage hwnd, &HA1, 2, 0&
End Sub
Public Sub RemoveMx(ByVal hwnd As Long)
Dim osS As Long
osS = GetWindowLong(hwnd, -16)
If (osS And &H10000) = &H10000 Then osS = osS Xor &H10000
SetWindowLong hwnd, -16, osS
End Sub
