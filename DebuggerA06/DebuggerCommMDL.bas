Attribute VB_Name = "DebuggerStuff"
Public ConfigData(20) As Byte

Public InsertVL As String
Public InsertIsCancel As Byte


Public MEMINF As MEMORY_BASIC_INFORMATION
Public MEMStr As String


'SEARCHING
Public gSTARTADR As Long
Public gLASTADR As Long
Public gSTARTADR2 As Long
Public gLASTADR2 As Long
Public gBegAdr As Long
Public gLenAdr As Long


Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

Public BeepH As Long
Public BeepM As Long
Public Traffic As Long 'Communicator handle
Sub Main()

If App.PrevInstance = True Then MsgBox "Only One Instance of this Debugger is allowed!", vbCritical, "Information!": Exit Sub

RegClass "TRAFFIC", 0, AddressOf WndProc
Traffic = CreateWindowEx(0, "TRAFFIC", "DebuggerVF", 0, 0, 0, 0, 0, 0, 0, App.hinstance, 0)
InitDBGEvents

'Install Global Hook Sniffer
Dim DebuggyAdr As Long
DebuggyAdr = GetModuleHandle("debuggy.dll")
Call InstallHook(0, DebuggyAdr, 0)
CloseHandle DebuggyAdr


Form4.Show
End Sub
Public Sub ContinueDebug()
Dim X As Long
X = OpenEvent(EVENT_ALL_ACCESS, 0, "ContinueDBG")
SetEvent X 'Continue Debug
End Sub
Public Sub StopDebug()
Dim X As Long
X = OpenEvent(EVENT_ALL_ACCESS, 0, "StopDBG")
SetEvent X 'Stop Debug
End Sub



Public Sub ReadProcessesForDebugger(ByRef LB As ListBox)
Dim PCS() As Long
PCS = GetActiveProcessesId
Dim u As Long
Dim ProcH As Long
Dim MDLS() As Long
Dim N() As Long
Dim CNT As Long
Dim Name As String
Dim nLen As Long
LB.Clear
For u = 0 To UBound(PCS)
ProcH = OpenProcess(PROCESS_ALL_ACCESS, 0, PCS(u))
If ProcH <> 0 Then
ReDim N(1000)
ret = EnumProcessModules(ProcH, N(0), 1000, CNT)
Name = Space(260)
nLen = GetModuleFileNameExA(ProcH, N(0), Name, 260)
Name = Left(Name, nLen)
LB.AddItem PCS(u) & vbTab & Name
End If
Next u
Erase MDLS
Erase N
End Sub

Public Function GetTransferString(ByVal StringToTransfer As String) As String
On Error GoTo Eend:
Dim DTA() As String
DTA = Split(StringToTransfer, " ")
Dim OUT() As Byte
ReDim OUT(UBound(DTA) - 1)
Dim u As Long
For u = 0 To UBound(DTA) - 1
If Len(DTA(u)) > 2 Or Len(DTA(u)) < 1 Then GoTo Eend
OUT(u) = GetRealByte(DTA(u))
Next u
GetTransferString = Space(UBound(OUT) + 1)
CopyMemory ByVal GetTransferString, OUT(0), UBound(OUT) + 1
Erase OUT
Erase DTA
Exit Function
Eend:
On Error GoTo 0
MsgBox "Error in Hex String!", vbCritical, "Error"
GetTransferString = ""
End Function
Public Function GetRealByte(ByRef StrX As String) As Byte
GetRealByte = CByte("&H" & StrX)
End Function
Public Function GetHI(ByVal VALUE As Long) As Long
CopyMemory GetHI, ByVal (VarPtr(VALUE) + 2), 2
End Function
Public Function GetLO(ByVal VALUE As Long) As Long
CopyMemory GetLO, ByVal (VarPtr(VALUE)), 2
End Function
Public Function GetHIofLO(ByVal VALUE As Long) As Long
CopyMemory ByVal VarPtr(GetHIofLO), ByVal VarPtr(VALUE) + 2, 1
End Function
Public Function GetLOofLO(ByVal VALUE As Long) As Long
CopyMemory ByVal VarPtr(GetLOofLO), ByVal VarPtr(VALUE) + 1, 1
End Function
Public Sub GetCursor(ByRef X As Long, ByRef Y As Long)
Dim CSP(1) As Long
GetCursorPos CSP(0)
X = CSP(0)
Y = CSP(1)
End Sub
