Attribute VB_Name = "PEread"
Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Public Const IMAGE_SIZEOF_SHORT_NAME = 8
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B


Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public Type IMAGE_IMPORT_DIRECTORY
    dwRVAFunctionNameList As Long
    TimeDateStamp As Long
    ForwarderChain As Long
    dwRVAModuleName As Long
    dwRVAFunctionAddressList As Long
End Type




Public Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type





Public Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
   
End Type




Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type


Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

'IMAGE DATA DIRECTORY:
'1-Export Table
'2-Import Table
'3-Resource Table
'4-Exception Table
'5-Certificate Table
'6-Relocation Table
'7-Debug Data
'8-Architecture Data
'9-Machine Value (MIPS GP)
'10-TLS Table
'11-Load Configuration Table
'12-Bound Import Table
'13-Import Address Table
'14-Delay Import Descriptor
'15-COM+ Runtime Header
'16-Reserved


Public Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
   OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type








Public Type IMPMOD
Module As String
FuncNames() As String
Addresses() As Long 'AS DWORD PTR[ADR]
CallingAddresses() As Long 'AS DIRECT CALL
Ord() As Integer
End Type

Public Type EXPMOD
BaseO As Long
ModuleName As String
FuncNames() As String
FuncAddress() As Long
Ord() As Integer
TempName() As String
End Type


Public DOSHEADER As IMAGEDOSHEADER
Public NTHEADER As IMAGE_NT_HEADERS
Public SECTIONSHEADER() As IMAGE_SECTION_HEADER




Public IMPORTSFLAG As Byte '0-no imp table!1-broken,2-
Public EXPORTFLAG As Byte '0-no Exp table!2-
Public IMPORTS() As IMAGE_IMPORT_DIRECTORY
Public EXPORTS As IMAGE_EXPORT_DIRECTORY
Public IMPS() As IMPMOD
Public EXPS As EXPMOD
Public EXPO As New Collection
Public IMPO() As New Collection


Public Function FindIn(SECS() As IMAGE_SECTION_HEADER, Addr As Long) As Byte
Dim u As Long
For u = 0 To UBound(SECS)
If Addr >= SECS(u).VirtualAddress And Addr <= SECS(u).VirtualAddress + SECS(u).SizeOfRawData Then FindIn = u: Exit Function
Next u
End Function

Public Function ReadPE(ByVal Filename As String, ByVal MAddress As Long, ByRef IMPORTSFLAG As Byte, ByRef EXPORTFLAG As Byte) As Byte
On Error GoTo Kraj
Dim DATA() As Byte
Dim CNT As Long
Dim u As Long
Dim i As Long
Dim FND2 As Long
Dim FND As Long
Dim FREEF As Long
FREEF = FreeFile

If Dir(Filename, vbHidden Or vbSystem Or vbReadOnly) = "" Then Dir "": Exit Function
Dir ""

Open Filename For Binary As #FREEF
ReDim DATA(4095) As Byte
Get #FREEF, , DATA


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

Dim SLen As Long
Dim LN As Long

Dim VVD As Long 'Imena iza
Dim VAD As Long 'Adresa
Dim VDIRC As Long

Dim TempOrd As Integer
'**********************IMPORTS
If NTHEADER.OptionalHeader.DataDirectory(1).VirtualAddress = 0 Then GoTo EXPT

FND = FindIn(SECTIONSHEADER, NTHEADER.OptionalHeader.DataDirectory(1).VirtualAddress)  'Nadji Sekciju!
CNT = SECTIONSHEADER(FND).PointerToRawData + (NTHEADER.OptionalHeader.DataDirectory(1).VirtualAddress _
- SECTIONSHEADER(FND).VirtualAddress)

LN = (NTHEADER.OptionalHeader.DataDirectory(1).size / 20) - 1
'duzina /20 -1,je broj import tabli
ReDim IMPORTS(LN - 1)
ReDim IMPS(LN - 1)

For u = 0 To LN - 1
Get #FREEF, CNT + 1, IMPORTS(u)

If IMPORTS(u).dwRVAFunctionAddressList = 0 Then
If u = 0 Then Erase IMPS: GoTo EXPT
ReDim Preserve IMPORTS(u - 1): ReDim Preserve IMPS(u - 1): Exit For
End If

'Nadji
FND = FindIn(SECTIONSHEADER, IMPORTS(u).dwRVAModuleName)
VAD = SECTIONSHEADER(FND).PointerToRawData + (IMPORTS(u).dwRVAModuleName _
- SECTIONSHEADER(FND).VirtualAddress)

'Ako je 0 tada su adrese iza


IMPS(u).Module = GetStr(VAD + 1, FREEF)

ReDim IMPS(u).FuncNames(20000)
ReDim IMPS(u).Addresses(20000)
ReDim IMPS(u).Ord(20000)
ReDim IMPS(u).CallingAddresses(20000)
'Adrese!********************************************************
FND = FindIn(SECTIONSHEADER, IMPORTS(u).dwRVAFunctionAddressList)
VAD = SECTIONSHEADER(FND).PointerToRawData + (IMPORTS(u).dwRVAFunctionAddressList _
- SECTIONSHEADER(FND).VirtualAddress)
'***********************************************************

Counter = 0
Do

Get #FREEF, VAD + Counter * 4 + 1, FNCMA

If FNCMA = 0 Then Exit Do 'ako nema više u tabli izadji van!

If IMPORTS(u).dwRVAFunctionNameList = 0 Then
VVD = VAD
GoTo INVAD

Else
FND = FindIn(SECTIONSHEADER, IMPORTS(u).dwRVAFunctionNameList)
VVD = SECTIONSHEADER(FND).PointerToRawData + (IMPORTS(u).dwRVAFunctionNameList _
- SECTIONSHEADER(FND).VirtualAddress)
INVAD:
Get #FREEF, VVD + Counter * 4 + 1, FNCNM

If FNCNM < 0 Then 'Ako nije ime,onda je ord.
IMPS(u).FuncNames(Counter) = "Hint:" & Hex(FNCNM And &HFFFF&) & "h"
IMPS(u).Ord(Counter) = FNCNM And &HFFFF&

Else
FNCNM = SECTIONSHEADER(FND).PointerToRawData + (FNCNM _
- SECTIONSHEADER(FND).VirtualAddress)
Get #FREEF, FNCNM + 1, IMPS(u).Ord(Counter)
IMPS(u).FuncNames(Counter) = GetStr(FNCNM + 3, FREEF)
End If
End If

IMPS(u).Addresses(Counter) = IMPORTS(u).dwRVAFunctionAddressList + 4 * Counter
FND2 = FindIn(SECTIONSHEADER, IMPS(u).Addresses(Counter))
VDIRC = SECTIONSHEADER(FND).PointerToRawData + (IMPS(u).Addresses(Counter) _
- SECTIONSHEADER(FND).VirtualAddress)
Get #FREEF, VDIRC + 1, IMPS(u).CallingAddresses(Counter)

Counter = Counter + 1
Loop

ReDim Preserve IMPS(u).Addresses(Counter - 1)
ReDim Preserve IMPS(u).FuncNames(Counter - 1)
ReDim Preserve IMPS(u).Ord(Counter - 1)


CNT = CNT + 20
Next u
OutI:
IMPORTSFLAG = 1


EXPT: 'Radi OK

Dim TEMPEXPS As EXPMOD

CNT = NTHEADER.OptionalHeader.DataDirectory(0).VirtualAddress
If CNT = 0 Then GoTo OutOf
FND = FindIn(SECTIONSHEADER, NTHEADER.OptionalHeader.DataDirectory(0).VirtualAddress)  'Nadji Sekciju!
CNT = SECTIONSHEADER(FND).PointerToRawData + (NTHEADER.OptionalHeader.DataDirectory(0).VirtualAddress _
- SECTIONSHEADER(FND).VirtualAddress)
Get #FREEF, CNT + 1, EXPORTS


ReDim TEMPEXPS.FuncAddress(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.FuncNames(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.Ord(EXPORTS.NumberOfFunctions + EXPORTS.Base - 1)
ReDim TEMPEXPS.TempName(EXPORTS.NumberOfNames - 1)

ReDim EXPS.FuncAddress(UBound(TEMPEXPS.FuncAddress))
ReDim EXPS.FuncNames(UBound(TEMPEXPS.FuncNames))
ReDim EXPS.Ord(UBound(TEMPEXPS.Ord))



'Uzmi samo imena imena!
FND = FindIn(SECTIONSHEADER, EXPORTS.AddressOfNames) 'Nadji Sekciju!
CNT = SECTIONSHEADER(FND).PointerToRawData + (EXPORTS.AddressOfNames _
- SECTIONSHEADER(FND).VirtualAddress)

For u = 0 To EXPORTS.NumberOfNames - 1
Get #FREEF, CNT + u * 4 + 1, FNCMA
FND = FindIn(SECTIONSHEADER, FNCMA) 'Nadji Sekciju!
FNCMA = SECTIONSHEADER(FND).PointerToRawData + (FNCMA - SECTIONSHEADER(FND).VirtualAddress)
TEMPEXPS.TempName(u) = GetStr(FNCMA + 1, FREEF)
Next u


FND = FindIn(SECTIONSHEADER, EXPORTS.AddressOfNameOrdinals) 'Nadji Sekciju!
FNCMA = SECTIONSHEADER(FND).PointerToRawData + (EXPORTS.AddressOfNameOrdinals - SECTIONSHEADER(FND).VirtualAddress)

'Uzmi imena i stavi na ordinal
For u = 0 To EXPORTS.NumberOfNames - 1
Get #FREEF, FNCMA + u * 2 + 1, TempOrd 'Uzmi sa AdrOfOrd
TEMPEXPS.FuncNames(TempOrd + EXPORTS.Base) = TEMPEXPS.TempName(u)
Next u
Erase EXPS.TempName


'Prodji kroz ordinal!
FND = FindIn(SECTIONSHEADER, EXPORTS.AddressOfFunctions) 'Nadji Sekciju!
CNT = SECTIONSHEADER(FND).PointerToRawData + (EXPORTS.AddressOfFunctions _
- SECTIONSHEADER(FND).VirtualAddress)
For u = 0 To EXPORTS.NumberOfFunctions - 1
Get #FREEF, CNT + u * 4 + 1, FNCMA
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
Close #1
Exit Function
Kraj:
On Error GoTo 0
Close #1
EXPORTFLAG = 0
IMPORTSFLAG = 0
End Function

Public Function GetStr(ByRef Position As Long, ByRef FREEF As Long) As String
Dim BSTRX() As Byte
Dim Counter2 As Long
ReDim BSTRX(255)
Counter2 = 0
Do
Get #FREEF, Position + Counter2, BSTRX(Counter2)
Counter2 = Counter2 + 1
Loop While BSTRX(Counter2 - 1) <> 0
GetStr = Space(Counter2 - 1)
CopyMemory ByVal GetStr, BSTRX(0), Counter2
End Function

