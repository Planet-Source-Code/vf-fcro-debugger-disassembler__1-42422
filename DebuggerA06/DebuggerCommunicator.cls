VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebuggerCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private EventsD As New Collection
Private DebuggerThreads As New Collection
Private ProcessesD As New Collection
Dim EventsForThread() As Long 'Temporary storage,must be in DATA SECTION!
'if isn't,the new debugger thread cannot reach the Events!

Public Sub CreateCommunicator()
End Sub
Public Property Get DebuggersCount() As Long
DebuggersCount = DebuggerThreads.count
End Property
Public Property Get DebuggerThreadId(ByVal Index As Long) As Long
Dim TDA() As Long
TDA = DebuggerThreads.Item(Index)
DebuggerThreadId = TDA(1)
End Property
Public Property Get DebuggedProcess(ByVal Index As Long) As Long
DebuggedProcess = ProcessesD.Item(Index)
End Property

Public Function CreateDebuggerMain(ByVal ProcessId As Long, ByRef DebuggerMainThreadId As Long) As Long
On Error GoTo EendX
ReDim EventsForThread(4) As Long
Dim TDA(1) As Long
Dim u As Long
For u = 0 To 4
EventsForThread(u) = CreateEvent(ByVal 0&, 0, 0, "X" & CStr(ProcessId) & CStr(u + 1))
Next u
EventsD.Add EventsForThread, "X" & CStr(ProcessId)
CreateDebuggerMain = CreateDebuggerMainThread(ProcessId, DebuggerMainThreadId, Traffic, VarPtr(EventsForThread(0)))
TDA(0) = CreateDebuggerMain
TDA(1) = DebuggerMainThreadId
DebuggerThreads.Add TDA, "X" & CStr(ProcessId)
ProcessesD.Add ProcessId, "X" & CStr(ProcessId)
Exit Function
EendX:
On Error GoTo 0
For u = 0 To 4
CloseHandle EventsForThread(u)
Next u
End Function


Public Sub StartDebug(ByVal ProcessId As Long)

Dim X As Long
X = OpenEvent(EVENT_ALL_ACCESS, 0, "X" & CStr(ProcessId) & "1")
SetEvent X  'Start Debug

End Sub

Public Sub ContinueDebug(ByVal ProcessId As Long)
Dim X As Long
X = OpenEvent(EVENT_ALL_ACCESS, 0, "X" & CStr(ProcessId) & "2")
SetEvent X 'Continue Debug
End Sub

Public Function StopDebug(ByVal ProcessId As Long)
On Error GoTo Dalje
Dim X As Long
X = OpenEvent(EVENT_ALL_ACCESS, 0, "X" & CStr(ProcessId) & "3")
SetEvent X
DestroyHandles ProcessId
DebuggerThreads.Remove "X" & CStr(ProcessId)
ProcessesD.Remove "X" & CStr(ProcessId)
Exit Function
Dalje:
On Error GoTo 0
End Function

Private Sub DestroyHandles(ByVal ProcessId As Long)
EventsForThread = EventsD.Item("X" & CStr(ProcessId))
Dim u As Long
For u = 0 To 4
CloseHandle EventsForThread(u)
Next u
EventsD.Remove ("X" & CStr(ProcessId))
Dim DTA() As Long
DTA = DebuggerThreads.Item("X" & CStr(ProcessId))
TerminateThread DTA(0), 0
End Sub

Public Sub DestroyCommunicator()


End Sub
