Attribute VB_Name = "WndBreaks"


Public BRKW As New Collection
Public BRKWMCMD As New Collection 'On WM_COMMAND
Public Function AddBreakWND(ByRef COL As Collection, ByVal hwnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long) As Byte
On Error Resume Next
Dim C As Collection
Dim isEx As Byte
Set C = GetRemoveCol(COL, hwnd, isEx)
If isEx = 0 Then Set C = New Collection

Dim Rk(2) As Long
Rk(0) = hwnd
Rk(1) = BrkMsgValue
Rk(2) = AfterOrBefore
C.Add Rk, "X" & BrkMsgValue
If Err <> 0 Then
On Error GoTo 0
Else
AddBreakWND = 1
End If
COL.Add C, "X" & hwnd
End Function
Public Function GetRemoveCol(ByRef COL As Collection, ByVal hwnd As Long, ByRef IsExist As Byte) As Collection
On Error GoTo Dalje
Set GetRemoveCol = COL.Item("X" & hwnd)
COL.Remove "X" & hwnd
IsExist = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub RemoveBreakWND(ByRef COL As Collection, ByVal hwnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long)
On Error GoTo Dalje
Dim C As Collection
Set C = COL.Item("X" & hwnd)
C.Remove "X" & BrkMsgValue
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetBreakWND(ByRef COL As Collection, ByVal hwnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long, ByRef IsValidWNDBP As Byte) As Long()
On Error GoTo Dalje:
Dim C As Collection
Set C = COL.Item("X" & hwnd)
GetBreakWND = C.Item("X" & BrkMsgValue)
IsValidWNDBP = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub RemoveEntireWND(ByRef COL As Collection, ByVal hwnd As Long)
On Error GoTo Dalje
COL.Remove "X" & hwnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub

