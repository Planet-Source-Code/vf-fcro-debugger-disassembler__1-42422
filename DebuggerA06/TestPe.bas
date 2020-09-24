Attribute VB_Name = "TestPe"
Public Function ReadPE2(ByVal MAddress As Long, ByRef IMPORTSFLAG As Byte, ByRef EXPORTFLAG As Byte, Optional ByVal StopOnHeader As Byte) As Byte
On Error GoTo Kraj
Dim DATA() As Byte
Dim CNT As Long
Dim u As Long
Dim i As Long
Dim FND2 As Long
Dim FND As Long
Dim FREEF As Long

GetDataFromMem MAddress, DATA, 4096

Dim FNCMA As Long
Dim FNCNM As Long
Dim Counter As Long
IMPORTSFLAG = 0
EXPORTFLAG = 0
CopyMemory DOSHEADER, DATA(CNT), Len(DOSHEADER)
CopyMemory NTHEADER, DATA(DOSHEADER.e_lfanew), Len(NTHEADER)
CNT = CNT + DOSHEADER.e_lfanew + Len(NTHEADER)
If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then Exit Function
ReDim SECTIONSHEADER(NTHEADER.FileHeader.NumberOfSections - 1)
For u = 0 To UBound(SECTIONSHEADER)
CopyMemory SECTIONSHEADER(u), DATA(CNT), Len(SECTIONSHEADER(0))
CNT = CNT + Len(SECTIONSHEADER(0))
Next u

If StopOnHeader = 1 Then Exit Function

GetDataFromMem MAddress, DATA, NTHEADER.OptionalHeader.SizeOfImage

Dim SLen As Long
Dim LN As Long

Dim VVD As Long 'Imena iza
Dim VAD As Long 'Adresa
Dim VDIRC As Long

Dim TempOrd As Integer
'**********************IMPORTS
If NTHEADER.OptionalHeader.DataDirectory(1).VirtualAddress = 0 Then GoTo EXPT


CNT = NTHEADER.OptionalHeader.DataDirectory(1).VirtualAddress

LN = (NTHEADER.OptionalHeader.DataDirectory(1).size / 20) - 1
'duzina /20 -1,je broj import tabli
ReDim IMPORTS(LN - 1)
ReDim IMPS(LN - 1)

For u = 0 To LN - 1
CopyMemory IMPORTS(u), DATA(CNT), Len(IMPORTS(u))


If IMPORTS(u).dwRVAFunctionAddressList = 0 Then
If u = 0 Then Erase IMPS: GoTo EXPT
ReDim Preserve IMPORTS(u - 1): ReDim Preserve IMPS(u - 1): Exit For
End If


VAD = IMPORTS(u).dwRVAModuleName
SLen = lstrlen(DATA(VAD))
IMPS(u).Module = Space(SLen)
CopyMemory ByVal IMPS(u).Module, DATA(IMPORTS(u).dwRVAModuleName), SLen

ReDim IMPS(u).FuncNames(20000)
ReDim IMPS(u).Addresses(20000)
ReDim IMPS(u).Ord(20000)
ReDim IMPS(u).CallingAddresses(20000)

VAD = IMPORTS(u).dwRVAFunctionAddressList

Counter = 0
Do


CopyMemory FNCMA, DATA(VAD + Counter * 4), 4

If FNCMA = 0 Then Exit Do 'ako nema više u tabli izadji van!

If IMPORTS(u).dwRVAFunctionNameList = 0 Then
VVD = VAD
GoTo INVAD

Else

VVD = IMPORTS(u).dwRVAFunctionNameList
INVAD:
CopyMemory FNCNM, DATA(VVD + Counter * 4), 4



If FNCNM < 0 Then 'Ako nije ime,onda je ord.
IMPS(u).FuncNames(Counter) = "Hint:" & Hex(FNCNM And &HFFFF&) & "h"
IMPS(u).Ord(Counter) = FNCNM And &HFFFF&

Else

CopyMemory IMPS(u).Ord(Counter), DATA(FNCNM), 2
SLen = lstrlen(DATA(FNCNM + 2))
IMPS(u).FuncNames(Counter) = Space(SLen)
CopyMemory ByVal IMPS(u).FuncNames(Counter), DATA(FNCNM + 2), SLen

End If
End If


IMPS(u).Addresses(Counter) = IMPORTS(u).dwRVAFunctionAddressList + 4 * Counter
VDIRC = IMPS(u).Addresses(Counter)
CopyMemory IMPS(u).CallingAddresses(Counter), DATA(VDIRC), 4


Counter = Counter + 1
Loop

ReDim Preserve IMPS(u).Addresses(Counter - 1)
ReDim Preserve IMPS(u).FuncNames(Counter - 1)
ReDim Preserve IMPS(u).Ord(Counter - 1)


CNT = CNT + 20
Next u
OutI:
IMPORTSFLAG = 1

'SREDIO IMPORT
'****************************************************************


EXPT: 'Radi OK

Dim TEMPEXPS As EXPMOD

CNT = NTHEADER.OptionalHeader.DataDirectory(0).VirtualAddress
If CNT = 0 Then GoTo OutOf

CNT = NTHEADER.OptionalHeader.DataDirectory(0).VirtualAddress
CopyMemory EXPORTS, DATA(CNT), Len(EXPORTS)

SLen = lstrlen(DATA(EXPORTS.Name))
EXPS.ModuleName = Space(SLen)
CopyMemory ByVal EXPS.ModuleName, DATA(EXPORTS.Name), SLen




ReDim TEMPEXPS.FuncAddress(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.FuncNames(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.Ord(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.TempName(EXPORTS.NumberOfNames - 1)

ReDim EXPS.FuncAddress(UBound(TEMPEXPS.FuncAddress))
ReDim EXPS.FuncNames(UBound(TEMPEXPS.FuncNames))
ReDim EXPS.Ord(UBound(TEMPEXPS.Ord))

CNT = EXPORTS.AddressOfNames


For u = 0 To EXPORTS.NumberOfNames - 1
CopyMemory FNCMA, DATA(CNT + u * 4), 4
SLen = lstrlen(DATA(FNCMA))
TEMPEXPS.TempName(u) = Space(SLen)
CopyMemory ByVal TEMPEXPS.TempName(u), DATA(FNCMA), SLen
Next u

FNCMA = EXPORTS.AddressOfNameOrdinals

For u = 0 To EXPORTS.NumberOfNames - 1
CopyMemory TempOrd, DATA(FNCMA + u * 2), 2
TEMPEXPS.FuncNames(TempOrd + EXPORTS.Base) = TEMPEXPS.TempName(u)
Next u
Erase EXPS.TempName


CNT = EXPORTS.AddressOfFunctions

For u = 0 To EXPORTS.NumberOfFunctions - 1
CopyMemory FNCMA, DATA(CNT + u * 4), 4



If FNCMA <> 0 Then
TEMPEXPS.FuncAddress(u + EXPORTS.Base) = MAddress + FNCMA
TEMPEXPS.Ord(u + EXPORTS.Base) = u + EXPORTS.Base
If Len(TEMPEXPS.FuncNames(u + EXPORTS.Base)) = 0 Then
TEMPEXPS.FuncNames(u + EXPORTS.Base) = "Ord:" & Hex(u + EXPORTS.Base) & "h"
End If
End If

Next u


'Trimaj nevažece ulaze
Counter = 0
EXPS.BaseO = TEMPEXPS.BaseO
For u = 0 To UBound(TEMPEXPS.Ord)
If TEMPEXPS.FuncAddress(u) <> 0 Then
EXPS.FuncAddress(Counter) = TEMPEXPS.FuncAddress(u)
EXPS.FuncNames(Counter) = TEMPEXPS.FuncNames(u)
EXPS.Ord(Counter) = TEMPEXPS.Ord(u)
Counter = Counter + 1
End If
Next u
ReDim Preserve EXPS.FuncAddress(Counter - 1)
ReDim Preserve EXPS.FuncNames(Counter - 1)
ReDim Preserve EXPS.Ord(Counter - 1)
Erase TEMPEXPS.FuncAddress
Erase TEMPEXPS.FuncNames
Erase TEMPEXPS.Ord
EXPORTFLAG = 1
OutOf:
Exit Function
Kraj:
On Error GoTo 0

EXPORTFLAG = 0
IMPORTSFLAG = 0
End Function

