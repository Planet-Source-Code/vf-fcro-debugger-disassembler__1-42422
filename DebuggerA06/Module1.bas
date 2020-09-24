Attribute VB_Name = "Module1"
Option Explicit



Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long

Public Type LPMODULEINFO
lpBaseOfDll As Long
SizeOfImage As Long
EntryPoint As Long
End Type

Public Const MEM_DECOMMIT = &H4000

'Mem state
Public Const MEM_RELEASE = &H8000
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000

Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const PAGE_EXECUTE_READ = &H20
Public Const PAGE_EXECUTE_WRITECOPY = &H80
Public Const PAGE_GUARD = &H100
Public Const PAGE_NOACCESS = &H1
Public Const PAGE_NOCACHE = &H200
Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOMBINE = &H400
Public Const PAGE_WRITECOPY = &H8

'Mem type
Public Const MEM_MAPPED = &H40000
Public Const MEM_IMAGE = &H1000000
Public Const MEM_PRIVATE = &H20000
Public Const MEM_PHYSICAL = &H400000



Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long

End Type

Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Declare Function VirtualLock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Any, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Declare Function VirtualUnlock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Any, ByVal dwFreeType As Long) As Long

Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Declare Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByRef lpmodinfo As LPMODULEINFO, ByVal CB As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal CB As Long, ByRef cbNeeded As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal CB As Long, ByRef cbNeeded As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const EVENT_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)

Public Const THREAD_GET_CONTEXT = (&H8)
Public Const THREAD_SET_CONTEXT = (&H10)
Public Const THREAD_SUSPEND_RESUME = (&H2)
Public Const THREAD_TERMINATE = (&H1)
Public Const THREAD_SET_THREAD_TOKEN = (&H80)
Public Const THREAD_SET_INFORMATION = (&H20)
Public Const THREAD_IDLE_TIMEOUT = 10
Public Const THREAD_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)


Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As Any) As Long
Public Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As Any) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Const WAIT_TIMEOUT = 258&
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'DEBUGGY.DLL Written in ASM By VANJA FUCKAR @2002!
'Designed for the Safety Communication with VB5/VB6 projects IN-DESIGN or RUNTIME...
Public Declare Function CreateDebuggerMainThread Lib "debuggy.dll" (ByVal ProcessToDebug As Long, ThreadId As Any, ByVal ConnectionHwnd As Long) As Long
Public Declare Function LoadP Lib "debuggy.dll" (lpApplicationName As Any, lpCommandLine As Any, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, lpCurrentDirectory As Any, lpStartupInfo As Any) As Long
Public Declare Sub UnLoadP Lib "debuggy.dll" ()
Public Declare Sub InitDBGEvents Lib "debuggy.dll" ()
Public Declare Sub UninitDBGEvents Lib "debuggy.dll" ()
Public Declare Function AddBy8 Lib "debuggy.dll" (ByVal Param1 As Long, ByVal Param2 As Long) As Long
Public Declare Function SubBy8 Lib "debuggy.dll" (ByVal Param1 As Long, ByVal Param2 As Long) As Long
Public Declare Function Search Lib "debuggy.dll" (ByVal SourceAddress As Long, ByVal SourceLength As Long, Pattern As Any, ByVal PatternLength As Long, ByVal MemHandle As Long) As Long
Public Declare Function Search2 Lib "debuggy.dll" (Source As Any, ByVal SourceLength As Any, Pattern As Any, ByVal PatternLength As Any) As Long
Public Declare Function InstallHook Lib "debuggy.dll" (ByVal Unused1 As Long, ByVal hinstance As Long, ByVal Unused2 As Long) As Long
Public Declare Sub UninstallHook Lib "debuggy.dll" ()
 Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public DBGEVENT As DEBUG_EVENT

Public EXCEPTIONINFO As EXCEPTION_DEBUG_INFO
Public LOADDLL As LOAD_DLL_DEBUG_INFO
Public UNLOADDLL As UNLOAD_DLL_DEBUG_INFO
Public CREATETHREADINFO As CREATE_THREAD_DEBUG_INFO
Public EXITTHREADINFO As EXIT_THREAD_DEBUG_INFO
Public CREATEPROCESSINFO As CREATE_PROCESS_DEBUG_INFO
Public EXITPROCESSINFO As EXIT_PROCESS_DEBUG_INFO
Public Function EnumW(ByVal hwnd As Long, ByVal Param1 As Long) As Long
Dim THR As Long
Dim PROC As Long
THR = GetWindowThreadProcessId(hwnd, PROC)
If PROC = ActiveProcess Then
Dim ClassNm As String
Dim ClLen As Long
ClassNm = Space(260)
ClLen = GetClassName(hwnd, ClassNm, 260)
ClassNm = Left(ClassNm, ClLen)
AddWins ClassNm, hwnd, THR
Call SetProp(hwnd, "GOFORDEBUG", 1)
EnumChildWindows hwnd, AddressOf EnumCW, THR
End If
EnumW = 1
End Function
Public Function EnumCW(ByVal hwnd As Long, ByVal Param1 As Long) As Long
Dim ClassNm As String
Dim ClLen As Long
ClassNm = Space(260)
ClLen = GetClassName(hwnd, ClassNm, 260)
ClassNm = Left(ClassNm, ClLen)
Call SetProp(hwnd, "GOFORDEBUG", 1)
AddWins ClassNm, hwnd, Param1
InE hwnd, Param1
EnumCW = 1
End Function
Public Sub InE(ByVal hwnd As Long, ByVal THR As Long)
EnumChildWindows hwnd, AddressOf EnumCW, THR
End Sub

Public Function GetActiveProcessesId() As Long()
Dim ret As Long
Dim ACTP() As Long
ReDim ACTP(100000)
Call EnumProcesses(ACTP(0), UBound(ACTP) + 1, ret)
ReDim Preserve ACTP(ret / 4 - 1)
GetActiveProcessesId = ACTP
End Function
Public Function PathFromName(ByVal PathEx As String) As String
Dim TmpRpc() As String
TmpRpc = Split(PathEx, "\")
ReDim Preserve TmpRpc(UBound(TmpRpc) - 1)
PathFromName = Join(TmpRpc, "\")
End Function
