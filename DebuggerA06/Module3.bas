Attribute VB_Name = "Module3"
Public Const SW_SHOWNORMAL = 1
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    CB As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (lpApplicationName As Any, lpCommandLine As Any, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, lpCurrentDirectory As Any, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Public Const INFINITE = &HFFFFFFFF
Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Declare Function WaitForDebugEvent Lib "kernel32.dll" (ByRef lpDebugEvent As DEBUG_EVENT, ByVal dwMilliseconds As Long) As Boolean
 

Public Const DBG_ATTACH = 14
Public Const DBG_BREAK = 6
Public Const DBG_CONTINUE = &H10002
Public Const DBG_CONTROL_BREAK = &H40010008
Public Const DBG_CONTROL_C = &H40010005
Public Const DBG_DIVOVERFLOW = 8
Public Const DBG_DLLSTART = 12
Public Const DBG_DLLSTOP = 13
Public Const DBG_DUMP_ADDRESS_AT_END = &H20000
Public Const DBG_DUMP_ADDRESS_OF_FIELD = &H10000
Public Const DBG_DUMP_CALL_FOR_EACH = &H8
Public Const DBG_DUMP_COMPACT_OUT = &H2000
Public Const DBG_DUMP_COPY_TYPE_DATA = &H40000
Public Const DBG_DUMP_FIELD_ARRAY = &H10
Public Const DBG_EXCEPTION_NOT_HANDLED = &H80010001
Public Const DBG_DUMP_NO_OFFSET = &H2
Public Const DBG_DUMP_NO_INDENT = &H1
Public Const DBG_DUMP_LIST = &H20

'Debugger can handle any exception of these
Public Const EXCEPTION_DEBUG_EVENT = 1
Public Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Public Const EXCEPTION_SINGLE_STEP = &H80000004
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Public Const EXCEPTION_BREAKPOINT = &H80000003
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Public Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Public Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Public Const EXCEPTION_INT_OVERFLOW = &HC0000095
Public Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Public Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096


Public Const CREATE_THREAD_DEBUG_EVENT = 2
Public Const CREATE_PROCESS_DEBUG_EVENT = 3
Public Const EXIT_THREAD_DEBUG_EVENT = 4
Public Const EXIT_PROCESS_DEBUG_EVENT = 5
Public Const LOAD_DLL_DEBUG_EVENT = 6
Public Const UNLOAD_DLL_DEBUG_EVENT = 7
Public Const OUTPUT_DEBUG_STRING_EVENT = 8

Public Const MAXIMUM_SUPPORTED_EXTENSION = 512
Public Const SIZE_OF_80387_REGISTERS = 80

Public Const CONTEXT_i486 = &H10000 '  // i486
Public Const CONTEXT_CONTROL = 1
Public Const CONTEXT_INTEGER = 2
Public Const CONTEXT_SEGMENTS = 4
Public Const CONTEXT_FLOATING_POINT = 8
Public Const CONTEXT_DEBUG_REGISTERS = 16
Public Const CONTEXT_EXTENDED_REGISTERS = 32
Public Const CONTEXT_EXTENDED_INTEGER = (CONTEXT_INTEGER Or &H10)
Public Const CONTEXT_FULL = (CONTEXT_CONTROL Or CONTEXT_FLOATING_POINT Or CONTEXT_INTEGER Or CONTEXT_EXTENDED_INTEGER)

Public Type CREATE_THREAD_DEBUG_INFO

    hThread As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
End Type

Public Type CREATE_PROCESS_DEBUG_INFO
   hFile As Long
    hProcess As Long
    hThread As Long
    lpBaseOfImage As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
    lpImageName As Long
    fUnicode As Integer
End Type

Public Type EXCEPTION_RECORD
   ExceptionCode As Long
   ExceptionFlags As Long
   pExceptionRecord As Long
   ExceptionAddress As Long
   NumberParameters As Long
   ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Public Type EXCEPTION_DEBUG_INFO
ExceptionRecord As EXCEPTION_RECORD
dwFirstChance As Long
End Type

Public Type EXIT_THREAD_DEBUG_INFO
   dwExitCode As Long
End Type

Public Type EXIT_PROCESS_DEBUG_INFO
  dwExitCode As Long
End Type

Public Type LOAD_DLL_DEBUG_INFO
  hFile As Long
    lpBaseOfDll As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpImageName As Long
    fUnicode As Integer
End Type

Public Type UNLOAD_DLL_DEBUG_INFO
 lpBaseOfDll As Long
End Type

Public Type OUTPUT_DEBUG_STRING_INFO
  lpDebugStringData As String
    fUnicode As Integer
    nDebugStringLength As Integer
End Type

Public Type RIP_INFO
 
 dwError As Long
    dwType As Long
End Type




Public Type DEBUG_EVENT
dwDebugEventCode As Long
dwProcessId As Long
dwThreadId As Long
DATA(20) As Long 'enough space
'UNION***NOT SUPPORTED BY VB
End Type

Public Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte
  
 'ST0 As Double
 'ST1 As Double
 'ST2 As Double
 'ST3 As Double
 'ST4 As Double
 'ST5 As Double
 'ST6 As Double
 'ST7 As Double
 'XXX As Double
 'YYY As Double
    Cr0NpxState As Long
End Type


Public Type CONTEXT

   ' //
   ' // The flags values within this flag control the contents of
   ' // a CONTEXT record.
   ' //
   ' // If the context record is used as an input parameter, then
   ' // for each portion of the context record controlled by a flag
   ' // whose value is set, it is assumed that that portion of the
   ' // context record contains valid context. If the context record
   ' // is being used to modify a threads context, then only that
   ' // portion of the threads context will be modified.
   ' //
   ' // If the context record is used as an IN OUT parameter to capture
   ' // the context of a thread, then only those portions of the thread's
   ' // context corresponding to set flags will be returned.
   ' //
   ' // The context record is never used as an OUT only parameter.
   ' //

    ContextFlags As Long

'    //
'    // This section is specified/returned if CONTEXT_DEBUG_REGISTERS is
'    // set in ContextFlags.  Note that CONTEXT_DEBUG_REGISTERS is NOT
'    // included in CONTEXT_FULL.
'    //

    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long
    
    
  '  //
  '  // This section is specified/returned if the
  '  // ContextFlags word contians the flag CONTEXT_FLOATING_POINT.
  '  //

   FloatSave As FLOATING_SAVE_AREA

  '  //
  '  // This section is specified/returned if the
  '  // ContextFlags word contians the flag CONTEXT_SEGMENTS.
  '  //

    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long

'  //
'    // This section is specified/returned if the
'    // ContextFlags word contians the flag CONTEXT_INTEGER.
'    //

    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long

'    //
'    // This section is specified/returned if the
'    // ContextFlags word contians the flag CONTEXT_CONTROL.
'    //

    Ebp As Long
    Eip As Long
    SegCs As Long       '       // MUST BE SANITIZED
    EFlags As Long      '       // MUST BE SANITIZED 'EFlags=&H100 For Single-Step Execution!!!!!!!!!
    Esp As Long
    SegSs As Long

'    //
'    // This section is specified/returned if the ContextFlags word
'    // contains the flag CONTEXT_EXTENDED_REGISTERS.
'    // The format and contexts are processor specific
'    //

   ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte

End Type


