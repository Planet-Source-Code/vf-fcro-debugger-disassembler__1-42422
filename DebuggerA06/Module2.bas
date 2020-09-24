Attribute VB_Name = "Module2"
Option Explicit

Public IsF11 As Boolean 'Is form11 visible


Public NameOfRunned As String

Public ISBPDisabled As Byte
Public IsLoadedProcess As Byte 'Provjera na Create Process
Public TerminateId As Long

Public GLOBALCOUNT As Long
Public GLOBALAFTERCOUNT As Long

Public NOTIFYJMPCALL As Byte 'JMP/CALL
Public NOTIFYVALG As Byte 'LEA,MOV,PUSH

Public DASM As New DisAsm

Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type


Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hinstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const CS_CLASSDC = &H40
Public Const CS_OWNDC = &H20
Public Const CS_GLOBALCLASS = &H4000
Public Const CS_HREDRAW = &H2
Public Const CS_PARENTDC = &H80
Public Const CS_VREDRAW = &H1

Public LastActiveIMP As Byte
Public LastActiveEXP As Byte

Public DATAPW() As Byte '??? Data for Search

Public ActiveH As Long '??
Public ActiveLength As Long  '??
'Public ActiveMName As String '??

Public ActiveProcess As Long 'trenutni process koji se debugira!
Public ActiveStackPosition As Long
Public ActiveBasePosition As Long
Public DISCOUNT As Long 'trenutna adresa gdje smo skrolali.
Public Forward As Byte
Public LAST As Long
Public NextF As Byte
Public NextB As Byte

Public ActiveThread As Long 'izabrani thread
Public ShowCaption As String

Public CTX As CONTEXT
Public ProcessHandle As Long
Public ACTIVEBREAKPOINTS As New Collection
Public PROCESSESTHREADS As New Collection
Public ACTMODULESBYPROCESS As New Collection
Public LASTTHREADEIP As New Collection
Public RTRIGGER As New Collection


Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hinstance As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal WMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hinstance As Long, lpParam As Any) As Long
 Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
 Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
 Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

 Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal WMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Public Sub AddLastEip(ByVal ThreadId As Long, ByVal Address As Long)
On Error Resume Next
LASTTHREADEIP.Remove "X" & ThreadId
If Err <> 0 Then On Error GoTo 0
LASTTHREADEIP.Add Address, "X" & ThreadId
End Sub
Public Function GetLastEip(ByVal ThreadId As Long) As Long
On Error GoTo Dalje
GetLastEip = LASTTHREADEIP.Item("X" & ThreadId)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub RemoveLastEip(ByVal ThreadId As Long)
On Error GoTo Dalje
LASTTHREADEIP.Remove "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub

'SREDIO
Public Sub RemoveBreakPoint(ByVal Address As Long)
On Error GoTo Dalje
Dim BPData() As Long
Dim ORGBYTE As Byte
BPData = ACTIVEBREAKPOINTS.Item("X" & Address)
ORGBYTE = CByte(BPData(1))
WriteProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
FlushInstructionCache ProcessHandle, ByVal Address, 1
ACTIVEBREAKPOINTS.Remove "X" & Address
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Sub AddBreakPoint(ByVal Address As Long, ByVal SkipWrite As Byte)
On Error Resume Next
Dim ORGBYTE As Byte
Dim HARDBP As Byte
Dim BPData(1) As Long
HARDBP = &HCC
Dim PHandle As Long
ACTIVEBREAKPOINTS.Remove "X" & Address
If Err <> 0 Then On Error GoTo 0
ReadProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
If SkipWrite = 0 Then
WriteProcessMemory ProcessHandle, ByVal Address, HARDBP, 1, ByVal 0&
FlushInstructionCache ProcessHandle, ByVal Address, 1
End If
BPData(0) = Address
BPData(1) = CLng(ORGBYTE)
ACTIVEBREAKPOINTS.Add BPData, "X" & Address
End Sub
'SREDIO

Public Function GetBreakPoint(ByVal Address As Long, ByRef IsFounded As Byte) As Byte
On Error GoTo Dalje
Dim BPData() As Long
BPData = ACTIVEBREAKPOINTS.Item("X" & Address)
GetBreakPoint = CByte(BPData(1))
IsFounded = 1
Exit Function
Dalje:
On Error GoTo 0
IsFounded = 0
End Function
Public Sub RestoreAllBreakPoints()
Dim u As Long
Dim BPData() As Long
Dim HBP As Byte
HBP = &HCC
For u = 1 To ACTIVEBREAKPOINTS.count
BPData = ACTIVEBREAKPOINTS.Item(u)
WriteProcessMemory ProcessHandle, ByVal BPData(0), HBP, 1, ByVal 0&
Next u
End Sub
Public Sub RestoreAllOriginalBytes()
Dim u As Long
Dim BPData() As Long
Dim ORGBYTE As Byte
For u = 1 To ACTIVEBREAKPOINTS.count
BPData = ACTIVEBREAKPOINTS.Item(u)
ORGBYTE = CByte(BPData(1))
WriteProcessMemory ProcessHandle, ByVal BPData(0), ORGBYTE, 1, ByVal 0&
Next u
End Sub
Public Sub RestoreOriginalBytes(ByVal Address As Long)
Dim IsValidBP As Byte
Dim ORGBYTE As Byte
ORGBYTE = GetBreakPoint(Address, IsValidBP)
If IsValidBP = 0 Then Exit Sub
WriteProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
End Sub
Public Sub RestoreBreakPoint(ByVal Address As Long)
Dim IsValidBP As Byte
Dim HBP As Byte
Call GetBreakPoint(Address, IsValidBP)
If IsValidBP = 0 Then Exit Sub
HBP = &HCC
WriteProcessMemory ProcessHandle, ByVal Address, HBP, 1, ByVal 0&
End Sub

'SREDIO
Public Function FindInModules(ByVal ActAddress As Long, Optional ByRef BaseAddress As Long, Optional ByRef Length As Long) As String
On Error GoTo Dalje
Dim u As Long
For u = 1 To ACTMODULESBYPROCESS.count
Dim Dat() As String
Dat = ACTMODULESBYPROCESS.Item(u)
If ActAddress >= CLng(Dat(1)) And ActAddress <= AddBy8(CLng(Dat(1)), CLng(Dat(2))) Then FindInModules = Dat(0): BaseAddress = CLng(Dat(1)): Length = CLng(Dat(2)): Exit Function
Next u
Exit Function
Dalje:
On Error GoTo 0
End Function
'DOBRO
Public Sub EnumActiveModules()
On Error GoTo Kraj
'Dim PHandle As Long
Dim N() As Long
Dim ret As Long
Dim PCSLength As Long
ReDim N(999)
ret = EnumProcessModules(ProcessHandle, N(0), 1000, PCSLength)
ReDim Preserve N(PCSLength / 4 - 1)
Dim u As Long
Set ACTMODULESBYPROCESS = Nothing
Dim S(3) As String
Dim INF1 As LPMODULEINFO
Dim MnM As String
Dim nLen As Long
For u = 0 To UBound(N)
MnM = Space(260)
nLen = GetModuleFileNameExA(ProcessHandle, N(u), MnM, 260)
MnM = Left(MnM, nLen)
Call GetModuleInformation(ProcessHandle, N(u), INF1, Len(INF1))
S(0) = MnM
S(1) = CStr(INF1.lpBaseOfDll)
S(2) = CStr(INF1.SizeOfImage)
S(3) = CStr(INF1.EntryPoint)
ACTMODULESBYPROCESS.Add S, "X" & CStr(INF1.lpBaseOfDll) 'Stavi kao base adresu

AddInExportsSearch S(0), INF1.lpBaseOfDll
Next u
Exit Sub
Kraj:
On Error GoTo 0
End Sub
Public Sub AddInActiveModules(ByVal Address As Long)
On Error GoTo Dalje
Dim S(3) As String
Dim EPPnt As Long
ReadPE2 Address, LastActiveIMP, LastActiveEXP
S(0) = ExPs.ModuleName
S(1) = Address
S(2) = NTHEADER.OptionalHeader.SizeOfImage
If NTHEADER.OptionalHeader.AddressOfEntryPoint <> 0 Then
EPPnt = Address + NTHEADER.OptionalHeader.AddressOfEntryPoint
End If
S(3) = EPPnt
ACTMODULESBYPROCESS.Add S, "X" & Address
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Function RemoveInActiveModules(ByVal Address As Long) As String
On Error GoTo Dalje
Dim S() As String
S = ACTMODULESBYPROCESS.Item("X" & CStr(Address))
RemoveInActiveModules = S(0)
ACTMODULESBYPROCESS.Remove "X" & CStr(Address)
Exit Function
Dalje:
On Error GoTo 0
End Function



Public Sub ReadModules(ByRef LB As ListBox)
On Error GoTo Dalje
LB.Clear
Dim Dat() As String
Dim u As Long
For u = 1 To ACTMODULESBYPROCESS.count
Dat = ACTMODULESBYPROCESS.Item(u)
LB.AddItem Dat(0) & vbTab & Hex(Dat(1)) & vbTab & Hex(Dat(2)) & vbTab & Hex(Dat(3))
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub

'SREDIO
'1-parametar ThreadId
'2-parametar Handle
Public Sub AddToThreadsList(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
On Error GoTo Dalje
Dim InTh(2) As Long
InTh(0) = ThreadId
InTh(1) = ThreadHandle
InTh(2) = 1
PROCESSESTHREADS.Add InTh, "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Sub RemoveFromThreadList(ByVal ThreadId As Long)
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
CloseHandle InTh(1)
PROCESSESTHREADS.Remove "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetHandleOfThread(ByVal ThreadId As Long, Optional ByRef IsRunning As Long) As Long
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
GetHandleOfThread = InTh(1)
IsRunning = InTh(2)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Function IsRunningThread(ByVal ThreadId As Long) As Long
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
IsRunningThread = InTh(2)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub ChangeStateThread(ByVal ThreadId As Long, ByVal StateTh As Long)
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
PROCESSESTHREADS.Remove "X" & ThreadId
InTh(2) = StateTh
PROCESSESTHREADS.Add InTh, "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub

'SREDIO
Public Sub ReadThreadsFromProcess(ByRef CB As ListBox)
CB.Clear
On Error GoTo Dalje
Dim i As String
Dim u As Long
Dim InTh() As Long
For u = 1 To PROCESSESTHREADS.count
InTh = PROCESSESTHREADS.Item(u)
If InTh(2) = 0 Then
i = "Suspend"
ElseIf InTh(2) = 1 Then
i = "Running"
ElseIf InTh(2) = 2 Then
i = "Waiting"
End If

If InTh(0) = MainPThread Then i = "(Main) " & i

CB.AddItem InTh(0) & "," & i
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub

' upravljanje hardware-skim breakpointima
Public Sub SetFlagInTrigger(ByVal ThreadId As Long, ByVal Address As Long, ByVal TriggerFlag As Long)
On Error GoTo Dalje
Dim TrDt() As Long
TrDt = RTRIGGER.Item("X" & ThreadId)
TrDt(1) = TriggerFlag
RTRIGGER.Remove "X" & ThreadId
RTRIGGER.Add TrDt, "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub SetInTrigger(ByVal ThreadId As Long, ByVal Address As Long)
On Error Resume Next
RTRIGGER.Remove "X" & ThreadId
If Err <> 0 Then On Error GoTo 0
Dim TrDt(1) As Long
TrDt(0) = Address
RTRIGGER.Add TrDt, "X" & ThreadId
End Sub
Public Sub RemoveInTrigger(ByVal ThreadId As Long)
On Error GoTo Dalje
RTRIGGER.Remove "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetFromTrigger(ByVal ThreadId As Long, ByRef IsValid As Byte, ByRef TriggerFlag As Long) As Long
On Error GoTo Dalje
Dim TrDt() As Long
TrDt = RTRIGGER.Item("X" & ThreadId)
GetFromTrigger = TrDt(0)
TriggerFlag = TrDt(1)
IsValid = 1
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub RegClass(ByVal Classname As String, ByVal hinstance As Long, ByVal AddressProc As Long)
Dim CLASSX As WNDCLASS
CLASSX.Style = CS_GLOBALCLASS
CLASSX.lpfnwndproc = AddressProc
CLASSX.hinstance = hinstance
CLASSX.lpszClassName = Classname
Call RegisterClass(CLASSX)
End Sub
Public Sub AddLine(ByRef StrX As String, ByRef RTB As RichTextBox)
RTB.SelStart = Len(RTB.Text)
RTB.SelLength = 0
RTB.SelText = StrX & vbCrLf
End Sub

'Communicator Messages Loop
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Connection Beetween Debugger Thread And our Project!
'The Best way for a harmless conversation with async thread out of our VB thread!


Select Case uMsg

Case 901
MsgBox "Cannot attach on Process:" & wParam & " ,I'm out...!", vbCritical, "Information"
StopAndClear: Unload Form1


Case 902
Form16.Show
'Form8.Caption = "Memory In Process:" & wParam
'Form8.Show
'Form12.Show
Form4.Enabled = False
Form4.Visible = False


Case 910
ProcessException wParam

Case 920 'Hooks Windows (Creating/Destroying Windows)

ProcessHooks wParam, lParam

Case 921 'Peek Message in Debugged process (This debugger process only WM_COMMAND)
'WH_CALLWNDPROC
ProcessMSG wParam, lParam, 0 'wparam=1 , lParam=PTR on struct

Case 922
'WH_CALLWNDPROCRET ***currently under progress
'ProcessMSG wParam, lParam, 1

Case 931
'Window proc -CHECKED FROM REMOTE THREAD CREATED BY THIS DEBUGGER!!!!!
Form14.Label8 = "Window Proc At Address:" & Hex(lParam)


Case 938


End Select


WndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
End Function
Public Sub ProcessMSG(ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Byte)
Dim CheckThreadId As Long
Dim CheckProcessId As Long

'Process Hooks for WH_CALLWNDPROC
'Message Chain designed by Vanja Fuckar @2002




Dim MSGS As CWPSTRUCT
Dim ret As Long
ret = ReadProcessMemory(ProcessHandle, ByVal Param2, MSGS, Len(MSGS), ByVal 0&)
If ret = 0 Then Exit Sub

Dim IsValidBRK As Byte
Dim WndSd() As Long
'0--HWND
'1--BRKMSG
'2--not used yet
If MSGS.message = WM_COMMAND Then
Dim NMSG As Long
NMSG = GetHI(MSGS.wParam)
WndSd = GetBreakWND(BRKWMCMD, MSGS.lParam, NMSG, 0, IsValidBRK)

If IsValidBRK = 1 Then
CheckThreadId = GetWindowThreadProcessId(MSGS.lParam, CheckProcessId)
PlayS
ActiveThread = CheckThreadId
Form16.InSuspend
AddLine "WM_COMMAND (Before Process) Breakpoint on value:" & NMSG & ",Class Name:" & ClassNameEx(MSGS.lParam) & ",Hwnd:" & Hex(MSGS.lParam) _
& ",On Parent Class Name:" & ClassNameEx(MSGS.hwnd) & ",Hwnd:" & Hex(MSGS.hwnd), Form12.rt1
If IsF11 Then Unload Form11
End If



Else
'Others WM_

WndSd = GetBreakWND(BRKW, MSGS.hwnd, MSGS.message, 0, IsValidBRK)

If IsValidBRK = 1 Then
CheckThreadId = GetWindowThreadProcessId(MSGS.hwnd, CheckProcessId)
PlayS
ActiveThread = CheckThreadId
If IsF11 Then Unload Form11
Form16.InSuspend
AddLine "WM_Value:" & Hex(MSGS.message) & " Breakpoint,Class Name:" & ClassNameEx(MSGS.hwnd) & ",Hwnd:" & Hex(MSGS.hwnd), Form12.rt1

End If
End If


'destroying windows
If MSGS.message = WM_DESTROY Then

CheckThreadId = GetWindowThreadProcessId(MSGS.hwnd, CheckProcessId)
AddLine "Destroy Window:" & Hex(MSGS.hwnd) & ",Class Name:" & ClassNameEx(MSGS.hwnd) & ",In Thread:" & CheckThreadId, Form12.rt1
RemoveProp MSGS.hwnd, "GOFORDEBUG" 'Remove property from window
RemoveWins MSGS.hwnd
RemoveEntireWND BRKW, MSGS.hwnd
RemoveEntireWND BRKWMCMD, MSGS.hwnd

If ConfigData(5) = 1 Then
ActiveThread = CheckThreadId
Form16.InSuspend
PlayS
End If


End If


End Sub

Public Function ClassNameEx(ByVal hwnd As Long) As String
Dim ClLen As Long
ClassNameEx = Space(260)
ClLen = GetClassName(hwnd, ClassNameEx, 260)
ClassNameEx = Left(ClassNameEx, ClLen)
End Function

Public Sub ProcessHooks(ByVal iMSG As Long, ByVal wParam As Long)
Dim CheckThreadId As Long
Dim CheckProcessId As Long

CheckThreadId = GetWindowThreadProcessId(wParam, CheckProcessId)
If CheckProcessId <> ActiveProcess Or ActiveProcess = 0 Then Exit Sub
Dim ClassNm As String

ClassNm = ClassNameEx(wParam)

'Window Hook message chain.
'Designed by Vanja Fuckar @2002
Select Case iMSG

Case HCBT_CREATEWND
AddLine "Create Window:" & Hex(wParam) & ",Class Name:" & ClassNm & ",In Thread:" & CheckThreadId, Form12.rt1
AddWins ClassNm, wParam, CheckThreadId
Call SetProp(wParam, "GOFORDEBUG", 1) 'Insert notify Flag for CallBack!
'If we don't do this,we will receive all messages from all of Desktop Shell Windows and Child Windows!
'surely that will do overload our project,yeeak!

If ConfigData(4) = 1 Then
ActiveThread = CheckThreadId
Form16.InSuspend
PlayS
End If


End Select

End Sub



Public Sub ProcessException(ByVal Param1 As Long)
'Param1=PTR on EVENT STRUCTURE!
Dim SHTC As Byte
Dim Addict As String
Dim TContext As CONTEXT
Dim MInf As LPMODULEINFO
Dim ModuleN As String
Dim TempThreadH As Long
Dim TempProcessH As Long
Dim IsValidBP As Byte
Dim IsValidTrigger As Byte
CopyMemory DBGEVENT, ByVal Param1, &H50&

Select Case DBGEVENT.dwDebugEventCode
Case EXCEPTION_DEBUG_EVENT
CopyMemory EXCEPTIONINFO, DBGEVENT.DATA(0), Len(EXCEPTIONINFO)

'--------
Select Case EXCEPTIONINFO.ExceptionRecord.ExceptionCode

    Case EXCEPTION_ACCESS_VIOLATION
    MsgBox "Access Violation At Address:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress) _
    & vbCrLf & "Try to Fix it by Changing the Registers,or Stop Debug this Process!", vbCritical, "Error"
    
    
    Case EXCEPTION_ILLEGAL_INSTRUCTION
    MsgBox "Illegal Instruction At Address:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress) _
    & vbCrLf & "Try to Fix it by Changing the Registers,or Stop Debug this Process!", vbCritical, "Error"
    

    Case EXCEPTION_INT_OVERFLOW
    MsgBox "Overflow At Address:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress), vbCritical, "Error"
    
    
    Case EXCEPTION_BREAKPOINT
    'Check if is our HARD BREAKPOINT!?
    GetBreakPoint EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, IsValidBP
    If IsValidBP = 1 Then
    CTX = GetContext(DBGEVENT.dwThreadId)
    CTX.Eip = CTX.Eip - 1
    SetContext DBGEVENT.dwThreadId, CTX
    SetInTrigger DBGEVENT.dwThreadId, EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    RestoreOriginalBytes CTX.Eip
    End If
    AddLine "Breakpoint encounted At:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress), Form12.rt1
    
    Case EXCEPTION_DATATYPE_MISALIGNMENT
 
    Case EXCEPTION_SINGLE_STEP
    Dim TrAdr As Long
    Dim TriggerFlag As Long
    TrAdr = GetFromTrigger(DBGEVENT.dwThreadId, IsValidTrigger, TriggerFlag)
    If IsValidTrigger = 1 Then
    RemoveInTrigger DBGEVENT.dwThreadId
    RestoreBreakPoint TrAdr
    If TriggerFlag = 1 Then
    'Ukoliko nije na single stepu nastavi
    ContinueDebug
    Exit Sub
    End If
   
    End If
           
    Case DBG_CONTROL_C

    Case Else

End Select

Case CREATE_THREAD_DEBUG_EVENT
CopyMemory CREATETHREADINFO, DBGEVENT.DATA(0), Len(CREATETHREADINFO)

If AccThreadX = DBGEVENT.dwThreadId Then ContinueDebug: Exit Sub

AddToThreadsList DBGEVENT.dwThreadId, CREATETHREADINFO.hThread
ReadThreadsFromProcess Form16.List1

AddLine "Create Thread:" & DBGEVENT.dwThreadId & Addict, Form12.rt1
SHTC = 2

If ConfigData(2) = 0 Then CTX = GetContext(DBGEVENT.dwThreadId): ShowThreadInfo SHTC: ContinueDebug: Exit Sub




Case CREATE_PROCESS_DEBUG_EVENT
CopyMemory CREATEPROCESSINFO, DBGEVENT.DATA(0), Len(CREATEPROCESSINFO)

If IsLoadedProcess = 1 Then
'Za sada single process debugger
MsgBox "Atempt to Create Another Process!" & vbCrLf & _
"This is a single Process Debugger.New Process will be terminated!", vbCritical, "Info"
TerminateProcess CREATEPROCESSINFO.hProcess, 0
TerminateId = DBGEVENT.dwProcessId
IsLoadedProcess = 2
AddLine "Create Process:" & DBGEVENT.dwProcessId & ",Thread:" & DBGEVENT.dwThreadId & " ***Terminated By This Debugger***", Form12.rt1
ContinueDebug
Exit Sub
End If


Call GetModuleInformation(CREATEPROCESSINFO.hProcess, CREATEPROCESSINFO.lpBaseOfImage, MInf, Len(MInf))
ProcessHandle = CREATEPROCESSINFO.hProcess
ActiveProcess = DBGEVENT.dwProcessId
MainPThread = DBGEVENT.dwThreadId
EnumActiveModules
AddToThreadsList DBGEVENT.dwThreadId, CREATEPROCESSINFO.hThread
ReadThreadsFromProcess Form16.List1
Form16.Caption = "Disassembling Process:" & ActiveProcess
AddLine "Create Process:" & DBGEVENT.dwProcessId & ",Thread:" & DBGEVENT.dwThreadId, Form12.rt1
IsLoadedProcess = 1
SHTC = 1

If Len(NameOfRunned) = 0 Then EnumWindows AddressOf EnumW, 0

'Dodan glavni modul...
'*********************

If Len(FindInModules(CREATEPROCESSINFO.lpBaseOfImage)) = 0 Then
'Ako nema tada
Dim S(3) As String
S(0) = NameOfRunned
AddInExportsSearch S(0), CREATEPROCESSINFO.lpBaseOfImage
S(1) = CREATEPROCESSINFO.lpBaseOfImage
S(2) = NTHEADER.OptionalHeader.SizeOfImage
S(3) = CREATEPROCESSINFO.lpBaseOfImage + NTHEADER.OptionalHeader.AddressOfEntryPoint
ACTMODULESBYPROCESS.Add S, "X" & CREATEPROCESSINFO.lpBaseOfImage

End If




Case EXIT_THREAD_DEBUG_EVENT
CopyMemory EXITTHREADINFO, DBGEVENT.DATA(0), Len(EXITTHREADINFO)

If AccThreadX = DBGEVENT.dwThreadId Then ContinueDebug: Exit Sub


RemoveLastEip DBGEVENT.dwThreadId
ActiveThread = 0

If ConfigData(3) = 0 Then
ContinueDebug
Else
ShowDatas DBGEVENT.dwProcessId, DBGEVENT.dwThreadId, CTX, 0
End If


RemoveFromThreadList DBGEVENT.dwThreadId
ReadThreadsFromProcess Form16.List1
AddLine "Exit Thread:" & DBGEVENT.dwThreadId & Addict, Form12.rt1


Exit Sub



Case EXIT_PROCESS_DEBUG_EVENT
If IsLoadedProcess = 2 Then
'Ako je process pokrenut od main process-a
AddLine "Exit Process:" & DBGEVENT.dwProcessId & " ***Terminated By This Debugger***", Form12.rt1
IsLoadedProcess = 1: ContinueDebug: Exit Sub
Else
StopAndClear
Unload Form16: Exit Sub
End If


Case LOAD_DLL_DEBUG_EVENT
CopyMemory LOADDLL, DBGEVENT.DATA(0), Len(LOADDLL)

'Preskoci ako je terminirajuci Process
If TerminateId = DBGEVENT.dwProcessId Then
AddLine "Loading ntdll.dll in process under the termination!", Form12.rt1
TerminateId = 0: ContinueDebug: Exit Sub
End If

If Len(NameOfRunned) = 0 Then
ModuleN = GetModuleNameFromHandle(LOADDLL.lpBaseOfDll)
Call GetModuleInformation(ProcessHandle, LOADDLL.lpBaseOfDll, MInf, Len(MInf))
If Len(ModuleN) = 0 Then GoTo AddEx
AddLine "Load Dll:" & ModuleN, Form12.rt1
Else
AddEx:
AddInActiveModules LOADDLL.lpBaseOfDll
AddInExportsSearch "", LOADDLL.lpBaseOfDll
AddLine "Load Dll:" & ExPs.ModuleName, Form12.rt1

End If
If ConfigData(0) = 0 Then ContinueDebug: Exit Sub

Case UNLOAD_DLL_DEBUG_EVENT
CopyMemory UNLOADDLL, DBGEVENT.DATA(0), Len(UNLOADDLL)
ModuleN = RemoveInActiveModules(UNLOADDLL.lpBaseOfDll)
DeleteInExportsSearch ModuleN
AddLine "Unload Dll:" & ModuleN, Form12.rt1


If ConfigData(1) = 0 Then ContinueDebug: Exit Sub

End Select


ShowDatas DBGEVENT.dwProcessId, DBGEVENT.dwThreadId, CTX, SHTC

End Sub
Public Sub ShowDatas(ByVal ProcessId As Long, ByVal ThreadId As Long, ByRef ContX As CONTEXT, ByVal ShowInfoTC As Byte)
ContX = GetContext(ThreadId)
AddLastEip ThreadId, ContX.Eip
ReadMem ProcessId, ContX.Eip
ActiveThread = ThreadId
ActiveStackPosition = ContX.Esp
ActiveBasePosition = ContX.Ebp
PlayS
ChangeStateThread ThreadId, 2
ReadThreadsFromProcess Form16.List1
If Form1.ShowingTHR = ThreadId Then Form1.ReadIT

ShowThreadInfo ShowInfoTC
End Sub

Public Sub ShowThreadInfo(ByVal ShowInfoTC As Byte)
If Len(NameOfRunned) <> 0 Then
If ShowInfoTC = 1 Then
AddLine "Process Start At Address:" & Hex(CTX.Eax), Form12.rt1
ElseIf ShowInfoTC = 2 Then
AddLine "Thread Start At Address:" & Hex(CTX.Eax), Form12.rt1
End If
End If
End Sub

Public Function GetModuleNameFromHandle(ByVal hModule As Long) As String
Dim lLen As Long
GetModuleNameFromHandle = Space(260)
lLen = GetModuleFileNameExA(ProcessHandle, hModule, GetModuleNameFromHandle, 260)
GetModuleNameFromHandle = Left(GetModuleNameFromHandle, lLen)
End Function

Public Sub ReadMem(ByVal ProcessId As Long, ByVal Address As Long)
Dim TName As String
TName = FindInModules(Address, ActiveH, ActiveLength)
If Len(TName) = 0 Then TName = "Unknown Address Or Not Valid"
Form16.Label7 = "Module:" & TName
DISCOUNT = Address
AddForward Form16.rt1, Form16.rt2, 25, ProcessId, Form16.List8
End Sub



Public Sub ReadContext(LBOX As ListBox, CTX As CONTEXT)
LBOX.Clear
LBOX.AddItem "EIP=" & Hex(CTX.Eip)
LBOX.AddItem "EAX=" & Hex(CTX.Eax)
LBOX.AddItem "EBX=" & Hex(CTX.Ebx)
LBOX.AddItem "ECX=" & Hex(CTX.Ecx)
LBOX.AddItem "EDX=" & Hex(CTX.Edx)
LBOX.AddItem "ESI=" & Hex(CTX.Esi)
LBOX.AddItem "EDI=" & Hex(CTX.Edi)
LBOX.AddItem "ESP=" & Hex(CTX.Esp)
LBOX.AddItem "EBP=" & Hex(CTX.Ebp)
LBOX.AddItem "EFLAGS=" & Hex(CTX.EFlags)
LBOX.AddItem "SEG CS=" & Hex(CTX.SegCs)
LBOX.AddItem "SEG DS=" & Hex(CTX.SegDs)
LBOX.AddItem "SEG ES=" & Hex(CTX.SegEs)
LBOX.AddItem "SEG FS=" & Hex(CTX.SegFs)
LBOX.AddItem "SEG GS=" & Hex(CTX.SegGs)
LBOX.AddItem "SEG SS=" & Hex(CTX.SegSs)
LBOX.AddItem "CARRY FLAG:" & (CTX.EFlags And 1&)
LBOX.AddItem "PARITY FLAG:" & ((CTX.EFlags And 4&) / 4&)
LBOX.AddItem "AUXILIARY FLAG:" & ((CTX.EFlags And 16&) / 16&)
LBOX.AddItem "ZERO FLAG:" & ((CTX.EFlags And 64&) / 64&)
LBOX.AddItem "SIGN FLAG:" & ((CTX.EFlags And 128&) / 128&)
LBOX.AddItem "TRAP FLAG:" & ((CTX.EFlags And 256&) / 256&)
LBOX.AddItem "INTERRUPT FLAG:" & ((CTX.EFlags And 512&) / 512&)
LBOX.AddItem "DIRECTION FLAG:" & ((CTX.EFlags And 1024&) / 1024&)
LBOX.AddItem "OVERFLOW FLAG:" & ((CTX.EFlags And 2048&) / 2048&)
LBOX.AddItem "IOPL FLAG:" & ((CTX.EFlags And 4096&) / 4096&)
LBOX.AddItem "NESTED FLAG:" & ((CTX.EFlags And 8192&) / 8192&)
LBOX.AddItem "RESUME FLAG:" & ((CTX.EFlags And 32768) / 32768)
LBOX.AddItem "VM FLAG:" & ((CTX.EFlags And 65536) / 65536)
LBOX.AddItem "AC FLAG:" & ((CTX.EFlags And 131072) / 131072)



End Sub
Public Function GetContext(ByVal ThreadId As Long) As CONTEXT
Dim ThHandle As Long
ThHandle = GetHandleOfThread(ThreadId)
GetContext.ContextFlags = CONTEXT_i486 Or CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT
GetThreadContext ThHandle, GetContext
End Function
Public Sub SetContext(ByVal ThreadId As Long, NewContext As CONTEXT)
Dim ThHandle As Long
ThHandle = GetHandleOfThread(ThreadId)
SetThreadContext ThHandle, NewContext
End Sub
Public Sub SetSingleStep(ByVal ThreadId As Long)
CTX = GetContext(ThreadId)
CTX.EFlags = CTX.EFlags Or &H100&
SetContext ThreadId, CTX
End Sub
Public Sub ClearSingleStep(ByVal ThreadId As Long)
CTX = GetContext(ThreadId)
If (CTX.EFlags And &H100&) = &H100& Then CTX.EFlags = CTX.EFlags Xor &H100&
SetContext ThreadId, CTX
End Sub

Public Sub GetDataFromMem(ByVal Address As Long, ByRef DataX() As Byte, ByVal NumOfBytes As Long, Optional ByRef IsValidMem As Long)
ReDim DataX(NumOfBytes - 1)
IsValidMem = ReadProcessMemory(ProcessHandle, ByVal Address, DataX(0), NumOfBytes, ByVal 0&)
End Sub
Public Sub AddForward25(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)
Dim u As Long
Dim ORGBYTE As Byte
Dim IsValidBP As Byte
Dim IsError As Byte
Dim DTX() As Byte
Dim CMDS As String
LAST = DISCOUNT

For u = 0 To Nums
GetDataFromMem LAST, DTX, 16
DASM.BaseAddress = LAST
ORGBYTE = GetBreakPoint(LAST, IsValidBP)
If IsValidBP <> 0 Then
DTX(0) = ORGBYTE
End If
CMDS = DASM.DisAssemble(DTX, 0, Forward, 0, 0, IsError)
LAST = LAST + Forward
Next u
DISCOUNT = LAST
NextB = 0
AddForward RTB, RTB2, Nums, ProcessId, LB2
End Sub
Public Sub AddBackward25(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)
Dim u As Long
For u = 0 To Nums - 1
AddBackward RTB, RTB2, Nums, ProcessId, LB2, 1
Next u
AddBackward RTB, RTB2, Nums, ProcessId, LB2
End Sub
Public Sub AddForward(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)
Dim DTX() As Byte
Dim CMDS As String
Dim AREF As String
Dim Allpr As String
Dim CHKC As String
Dim IsError As Byte
Dim CRef As String
Dim ExpSt As String
Dim IsValidBP As Byte
Dim ORGBYTE As Byte
Dim GFI As String
Dim IsString As Long
Dim i As Long
Dim u As Long
LAST = DISCOUNT

For u = 0 To Nums

GetDataFromMem LAST, DTX, 16
DASM.BaseAddress = LAST

ORGBYTE = GetBreakPoint(LAST, IsValidBP)
If IsValidBP = 0 Then
LB2.List(u) = ""
Else
LB2.List(u) = "*BP*"
DTX(0) = ORGBYTE
End If

CMDS = DASM.DisAssemble(DTX, 0, Forward, 0, 0, IsError)

NotifyData1(u) = NOTIFYJMPCALL
NotifyData2(u) = VALUES1
NotifyData3(u) = VALUES2
NotifyData4(u) = VALUES3



If NOTIFYVALG = 1 Then
AREF = IsStringOnAdr(IsString)
If IsString = 1 Then AREF = "(Possible) String: " & AREF
End If

ExpSt = GetFromExportsSearch(FindInModules(LAST), LAST)
If Len(ExpSt) <> 0 Then ExpSt = "Export:" & ExpSt


GFI = GetFromIndex(INDEXESR, REFSR, LAST)
If Len(GFI) = 0 Then
GFI = GetFromIndex(EINDEXESR, EREFSR, LAST)
End If


CHKC = CheckCALL(NotifyData2(u))
Allpr = ExpSt



If Len(Allpr) <> 0 And Len(CHKC) <> 0 Then
Allpr = Allpr & " ;" & CHKC
ElseIf Len(CHKC) <> 0 Then
Allpr = CHKC
End If

If Len(Allpr) <> 0 And Len(AREF) <> 0 Then
If Len(CHKC) = 0 Then
Allpr = Allpr & " ;" & AREF
End If
ElseIf Len(AREF) <> 0 Then
Allpr = AREF
End If


If Len(Allpr) <> 0 And Len(GFI) <> 0 Then
Allpr = Allpr & " ;" & GFI
ElseIf Len(GFI) <> 0 Then
Allpr = GFI
End If

RTB(u) = Hex(LAST) & vbTab & CMDS
RTB2(u) = Allpr
'RTB2(u) = ExpSt & AREF & CheckCALL(NotifyData2(u)) & GFI

If u = 0 Then
NextB = Forward
End If
LAST = LAST + Forward
AREF = ""
Next u
NextF = Forward

End Sub
Public Sub AddBackward(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox, Optional ByVal SkipPres As Byte)
Dim IsError As Byte
Dim DTX() As Byte
GetDataFromMem DISCOUNT - 49, DTX, 50
RestoreOrgBytes DTX, DISCOUNT - 49, ProcessId
Call DASM.DisassembleBack(DTX, 49, Forward, IsError)
DISCOUNT = DISCOUNT - Forward
NextB = 0
NextF = 0
If SkipPres = 0 Then
AddForward Form16.rt1, Form16.rt2, Nums, ProcessId, LB2
End If
End Sub
Public Sub RestoreOrgBytes(DATA() As Byte, ByVal StartAdr As Long, ByVal ProcessId As Long)
Dim u As Long
Dim IsValidBP As Byte
Dim ORGBYTE As Byte
For u = 0 To UBound(DATA)
ORGBYTE = GetBreakPoint(StartAdr + u, IsValidBP)
If IsValidBP = 1 Then
DATA(u) = ORGBYTE
End If
IsValidBP = 0
Next u
End Sub



'Public Function GetFromProcess(ByRef COL As Collection, ByVal ProcessId As Long, ByRef IsValid As Byte) As Long
'On Error GoTo Dalje
'GetFromProcess = COL.Item("X" & ProcessId)
'IsValid = 1
'Exit Function
'Dalje:
'On Error GoTo 0
'End Function


'Public Sub AddInProcess(ByRef COL As Collection, ByVal ProcessId As Long)
'On Error GoTo Dalje
'Dim ProcH As Long
'ProcH = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessId)
'Dim VHandle As Long
'VHandle = VirtualAllocEx(ProcH, ByVal 0&, ByVal 100&, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'COL.Add VHandle, "X" & ProcessId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
'Public Sub RemoveFromProcess(ByRef COL As Collection, ByVal ProcessId As Long)
'On Error GoTo Dalje
'Dim VHandle As Long
'VHandle = COL.Item("X" & ProcessId)
'Dim ProcH As Long
'ProcH = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessId)
'Call VirtualFreeEx(ProcH, ByVal VHandle, ByVal 100&, MEM_DECOMMIT)
'CloseHandle ProcH
'COL.Remove "X" & ProcessId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
Public Function IsValidK(ByRef S As String) As Byte
On Error GoTo Eend
Dim K As Byte
K = CByte("&H" & S)
IsValidK = 1
Eend:
On Error GoTo 0
End Function
Public Sub StopAndClear()
Set ModulesExports = Nothing
Set ACTIVEBREAKPOINTS = Nothing
Set PROCESSESTHREADS = Nothing
Set ACTMODULESBYPROCESS = Nothing
Set RTRIGGER = Nothing
Set LASTTHREADEIP = Nothing
Set WINS = Nothing
Set BRKW = Nothing
Set BRKWMCMD = Nothing
Erase REFSR
Set INDEXESR = Nothing
Erase EREFSR
Set EINDEXESR = Nothing
Erase SREFSR
Set SINDEXESR = Nothing
ValidCRef = ""
StopDebug
End Sub


Public Function TestPTR(ByVal Address As Long, Optional ByRef DataB As Byte) As Byte
TestPTR = ReadProcessMemory(ProcessHandle, ByVal Address, DataB, 1, ByVal 0&)
End Function

Public Sub ReadStackFrame(LB As ListBox, ByVal Eip As Long, ByVal Ebp As Long)
Dim IsErr As Byte
Dim TA As Long
Dim TA2 As Long
Dim TName As String
Dim u As Long
Dim RAdrs() As Long
LB.Clear
TName = FindInModules(Eip, TA, TA2)
LB.AddItem Hex(Eip) & vbTab & TName
RAdrs = CallStack(Ebp, IsErr)
If IsErr = 1 Then Exit Sub
For u = 0 To UBound(RAdrs)
TName = FindInModules(RAdrs(u), TA, TA2)
LB.AddItem Hex(RAdrs(u)) & vbTab & TName
Next u

End Sub


Public Function CallStack(ByVal Ebp As Long, ByRef IsError As Byte) As Long()
On Error GoTo Dalje
Dim StackFr() As Long
Dim count As Long
ReDim StackFr(10000)

Dim PtAddress(1) As Long
Do
Call ReadProcessMemory(ProcessHandle, ByVal Ebp, PtAddress(0), 8, ByVal 0&)
If Ebp = PtAddress(0) Or PtAddress(0) = 0 Then Exit Do
StackFr(count) = PtAddress(1)
count = count + 1
Ebp = PtAddress(0)
Loop
If count = 0 Then GoTo Dalje
ReDim Preserve StackFr(count - 1)
CallStack = StackFr
Exit Function
Dalje:
On Error GoTo 0
IsError = 1
End Function


Public Sub Read9Stack(LB As ListBox, ByVal StackPos As Long, ByVal EXXp As Long, ByVal EXXpS As String)
Dim FFAdr As Long
Dim u As Long
Dim Xret As Long
Dim Buffy As Long
Dim VD As String


FFAdr = SubBy8(StackPos, 12)
LB.Clear
For u = 0 To 6

If EXXp < FFAdr Then
VD = "+" & Hex(SubBy8(FFAdr, EXXp))
ElseIf EXXp > FFAdr Then
VD = "-" & Hex(SubBy8(EXXp, FFAdr))
Else
VD = "**"
End If



Xret = ReadProcessMemory(ProcessHandle, ByVal FFAdr, Buffy, 4, ByVal 0&)
If Xret = 0 Then
LB.AddItem "[" & EXXpS & VD & "]" & vbTab & "Not Valid"
Else
LB.AddItem "[" & EXXpS & VD & "]" & vbTab & Hex(Buffy)
End If
FFAdr = FFAdr + 4
Next u

End Sub



Public Function CheckCALL(ByRef NewVal As Long) As String
Dim Redr As String
Dim TName As String
Dim BaAdr As Long

If NOTIFYJMPCALL = 2 Or NOTIFYJMPCALL = 1 Then
'CALL ADR,JMP ADR
TName = FindInModules(VALUES1, BaAdr)
If Len(TName) = 0 Then Exit Function
CheckCALL = GetFromExportsSearch(TName, VALUES1)
If Len(CheckCALL) = 0 Then


Dim OIsValid As Byte
Dim OredTemp() As Byte
GetDataFromMem VALUES1, OredTemp, 16
Dim Ofwr As Byte
Dim Ored As String
DASM.BaseAddress = VALUES1
NewVal = VALUES1
Call DASM.DisAssemble(OredTemp, 0, Ofwr, 0, 0)
If NOTIFYJMPCALL = 4 Then
'Redir with JMP DWORD PTR[ ]
Redr = "Redirect to "
GoTo InNtf2

End If

Else
CheckCALL = "Import:" & TName & ":" & CheckCALL
End If

ElseIf NOTIFYJMPCALL = 3 Or NOTIFYJMPCALL = 4 Or NOTIFYJMPCALL = 5 Then
'CALL DWORD [ADR],JMP DWORD [ADR],MOV XXX,DWORD PTR[ADR]
InNtf2:
Dim LxAddr As Long
Call ReadProcessMemory(ProcessHandle, ByVal VALUES1, LxAddr, 4, ByVal 0&)


TName = FindInModules(LxAddr, BaAdr)
If Len(TName) = 0 Or LxAddr = 0 Then NewVal = 0: Exit Function
CheckCALL = GetFromExportsSearch(TName, LxAddr)
If Len(CheckCALL) <> 0 Then
CheckCALL = "Import:" & Redr & TName & ":" & CheckCALL
End If

If Len(Redr) = 0 Then NewVal = LxAddr



End If
End Function

Public Function GetFromIndex(INDX As Collection, REFX() As Collection, ByRef ToAdr As Long) As String
Dim STS() As String
Dim u As Long
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckExs(INDX, ToAdr, LYX)
If RZt = 0 Then Exit Function
ReDim STS(REFX(LYX).count - 1)
For u = 1 To REFX(LYX).count
STS(u - 1) = Hex(REFX(LYX).Item(u))
Next u
GetFromIndex = "Jumps From:" & Join(STS, ",")
End Function


Public Sub AddInIndex(INDX As Collection, REFX() As Collection, ByRef FromAdr As Long, ByRef ToAdr As Long)
On Error GoTo Dalje

Dim IXX(1) As Long
'0-index
'1-Address To JMP
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckExs(INDX, ToAdr, LYX)
If RZt = 0 Then
IXX(0) = INDX.count
IXX(1) = ToAdr
INDX.Add IXX, "X" & ToAdr
Set REFX(IXX(0)) = New Collection
REFX(IXX(0)).Add FromAdr
Else
REFX(LYX).Add FromAdr
End If
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Public Function CheckExs(INDX As Collection, ByRef ToAdr As Long, ByRef INX As Long) As Byte
On Error GoTo Dalje
Dim VD() As Long
VD = INDX("X" & ToAdr)
INX = VD(0) 'Uzmi index
CheckExs = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub AddInStringIndex(INDX As Collection, REFX() As Collection, ByRef FromAdr As Long, ByRef ToAdr As Long, ByRef StringX As String)
On Error GoTo Dalje

Dim IXX(2) As String
'0-index
'1-Address To JMP
'2-string
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckStringExs(INDX, ToAdr, LYX)
If RZt = 0 Then
IXX(0) = INDX.count
IXX(1) = ToAdr
IXX(2) = StringX
INDX.Add IXX, "X" & ToAdr
Set REFX(INDX.count - 1) = New Collection
REFX(INDX.count - 1).Add FromAdr
Else
REFX(LYX).Add FromAdr
End If
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function CheckStringExs(INDX As Collection, ByRef ToAdr As Long, ByRef INX As Long) As Byte
On Error GoTo Dalje
Dim VD() As String
VD = INDX("X" & ToAdr)
INX = CLng(VD(0)) 'Uzmi index
CheckStringExs = 1
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Function GetFromStringIndex(INDX As Collection, REFX() As Collection, ByRef ToAdr As Long) As String
Dim STS() As String
Dim u As Long
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckStringExs(INDX, ToAdr, LYX)
If RZt = 0 Then Exit Function
ReDim STS(REFX(LYX).count - 1)
For u = 1 To REFX(LYX).count
STS(u - 1) = Hex(CLng(REFX(LYX).Item(u)))
Next u
GetFromStringIndex = "Refs From:" & Join(STS, ",")
End Function
