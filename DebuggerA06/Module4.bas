Attribute VB_Name = "Module4"
Public ModulesExports As New Collection
Public WINS As New Collection



Public Sub AddWins(ByVal ClassNm As String, ByVal hwnd As Long, ByVal ThreadId As Long)
On Error GoTo Dalje
Dim WSd(2) As String
WSd(0) = ClassNm
WSd(1) = CStr(hwnd)
WSd(2) = CStr(ThreadId)

WINS.Add WSd, "X" & hwnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub RemoveWins(ByVal hwnd As Long)
On Error GoTo Dalje
WINS.Remove "X" & hwnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub


Public Sub AddInExportsSearch(ByVal MName As String, ByVal BaseAdr As Long)
On Error GoTo Dalje
Dim IMPFF As Byte
Dim EXPFF As Byte

ReadPE2 BaseAdr, IMPFF, EXPFF




Dim u As Long
Dim CNew As New Collection

If NTHEADER.OptionalHeader.AddressOfEntryPoint <> 0 Then
AddToCols CNew, AddBy8(BaseAdr, NTHEADER.OptionalHeader.AddressOfEntryPoint), "Entry Point"
End If

If EXPFF = 1 Then
For u = 0 To UBound(EXPS.FuncNames)
AddToCols CNew, EXPS.FuncAddress(u), EXPS.FuncNames(u)
If EXPS.FuncNames(u) = "VFDebuggerThread" Then DebuggyOut = EXPS.FuncAddress(u)
If EXPS.FuncNames(u) = "VFDebuggerTerminate" Then SusThreadX = EXPS.FuncAddress(u)
Next u
End If


If Len(MName) = 0 Then MName = EXPS.ModuleName
ModulesExports.Add CNew, MName
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Public Sub DeleteInExportsSearch(ByVal MName As String)
On Error GoTo Dalje
Dim COld As Collection
Set COld = ModulesExports.Item(MName)
Set COld = Nothing
ModulesExports.Remove (MName)
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Public Function GetFromExportsSearch(ByVal MName As String, ByVal Address As Long) As String
On Error GoTo Dalje
Dim COld As Collection
Set COld = ModulesExports.Item(MName)
GetFromExportsSearch = COld("X" & Address)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub AddToCols(COL As Collection, ByRef DataL As Long, ByRef DataS As String)
On Error GoTo Dalje
COL.Add DataS, "X" & DataL
Exit Sub
Dalje:
On Error GoTo 0
Dim S2 As String
S2 = COL.Item("X" & DataL)
If S2 = DataS Then Exit Sub
S2 = S2 & "," & DataS
COL.Remove "X" & DataL
COL.Add S2, "X" & DataL
End Sub


