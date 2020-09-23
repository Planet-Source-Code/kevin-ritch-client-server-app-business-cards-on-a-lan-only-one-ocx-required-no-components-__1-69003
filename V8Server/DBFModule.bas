Attribute VB_Name = "dbaseModule"
Global Const BusinessCardTable = 1
Global tb As String
Global Record As Long
Global DataLine$
Global NumFlds(9) As Long
Global Records(9) As Long
Global HeaderSize(9) As Long
Global Pointer As Long
Global RecLen(9) As Long
Global FldLen(254, 9) As Integer
Global MidPos(254, 9) As Integer
Global FldType(254, 9) As Integer
Global FldNam$(254, 9)
Global FldData$(254)

Global FNameIndex As String
Global LNameIndex As String
Global CompanyIndex As String

Global DBF_File As String

Public Const FLD_MrMrsMs = 1
Public Const FLD_FName = 2
Public Const FLD_LName = 3
Public Const FLD_Company = 4
Public Const FLD_Phone = 5
Public Const FLD_Extension = 6
Public Const FLD_Fax = 7
Public Const FLD_Mobile = 8
Public Const FLD_Home = 9
Public Const FLD_EMail = 10
Public Const FLD_Website = 11
Public Const FLD_Addr1 = 12
Public Const FLD_Addr2 = 13
Public Const FLD_City = 14
Public Const FLD_State = 15
Public Const FLD_Zip = 16
Public Const FLD_IDStatus = 17
Public Const FLD_Owner = 18
Public Const FLD_DirectLine = 19
Public Const FLD_Department = 20
Public Const FLD_JobTitle = 21

Function AppendToDBF(DBF$, DL$, SubScript%) As Long
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 XL = HeaderSize(SubScript%) + (Records(SubScript%) * RecLen(SubScript%)) + 1
 NewRecs = Records(SubScript%) + 1 ' Increase Record Count
 Records(SubScript%) = NewRecs
 R1 = NewRecs Mod 256
 R2 = NewRecs \ 256
 RKS$ = Chr$(R1) + Chr$(R2) ' Number of Records
 Put #DF, XL, DL$
 Put #DF, 5, RKS$
 Close #DF
'==============
'UPDATE INDEXES
'==============
 CompanyIndex = CompanyIndex & UCase$(Mid$(DL$, MidPos(FLD_Company, SubScript%), 1))
 FNameIndex = FNameIndex & UCase$(Mid$(DL$, MidPos(FLD_FName, SubScript%), 1))
 LNameIndex = LNameIndex & UCase$(Mid$(DL$, MidPos(FLD_LName, SubScript%), 1))
'========================
'Return NEW Record Number
'========================
 AppendToDBF = NewRecs
End Function
Sub DeleteRecord(DBF$, RecNum, SubScript%)
 Record = RecNum
 If Record < 1 Or Record > Records(SubScript%) Then
  Exit Sub
 End If
 DF = FreeFile
 DelChar$ = "*"
 Open DBF$ For Binary Shared As #DF
 Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + 1
 Put #DF, Pointer, DelChar$
 Close #DF
End Sub
Function GetRecord(DBF$, RecNum, SubScript%) As String
 Record = RecNum
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + 1
 Seek #DF, Pointer
 DataLine$ = Input$(RecLen(SubScript%), #DF)
 GetRecord = DataLine$
 Close #DF
' X = 2
' For i = 1 To NumFlds(SubScript%)
'  FldData$(i) = Trim$(Mid$(DataLine$, X, FldLen(i, SubScript%)))
'  X = X + FldLen(i, SubScript%)
' Next i
End Function
Sub GetDBFRecordCount(DBF$, SubScript%)
 Dim FB As Long
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 FileSize = LOF(DF)
 Seek #DF, 5
 X1$ = Input$(1, #DF)
 X2$ = Input$(1, #DF)
 X3$ = Input$(1, #DF)
 Records(SubScript%) = Asc(X3$) * (256# * 256#)
 Records(SubScript%) = Records(SubScript%) + Asc(X1$) + (Asc(X2$) * 256#)
 Close #DF
End Sub
Sub GetDBStructure(DBF$, SubScript%)
 Dim FB As Long
 DF = FreeFile
 NumFlds(SubScript%) = 0
 Open DBF$ For Binary Shared As #DF
 FileSize = LOF(DF)
 Seek #DF, 5
 X1$ = Input$(1, #DF)
 X2$ = Input$(1, #DF)
 X3$ = Input$(1, #DF)
 Records(SubScript%) = Asc(X3$) * (256# * 256#)
 Records(SubScript%) = Records(SubScript%) + Asc(X1$) + (Asc(X2$) * 256#)
 Seek #DF, 9
 X1$ = Input$(1, #DF)
 X2$ = Input$(1, #DF)
 HeaderSize(SubScript%) = Asc(X1$) + (Asc(X2$) * 256#)
 Seek #DF, 11
 X1$ = Input$(1, #DF)
 X2$ = Input$(1, #DF)
 RecLen(SubScript%) = Asc(X1$) + (Asc(X2$) * 256#)
 NumFlds(SubScript%) = ((HeaderSize(SubScript%) + 20) \ 32) - 1
 Seek #DF, 33
 MMid = 2
 If NumFlds(SubScript%) > 999 Then
  NumFlds(SubScript%) = 999
 End If
 For i = 1 To NumFlds(SubScript%)
  FldName$ = Input$(10, #DF) + Chr$(0)
  Ignore$ = Input$(1, #DF)
  FldTyp$ = UCase$(Input$(1, #DF))
  Select Case FldTyp$
   Case "N" ' NUMERIC
    FT = 2
   Case "D" ' DATE FIELD
    FT = 3
   Case "L" ' Logical Boolean (T/F)
    FT = 4
   Case Else
    FT = 1 ' CHARACTER
  End Select
  FldType(i, SubScript) = FT
  Ignore$ = Input$(3, #DF)
  Ignore$ = Input$(1, #DF)
  FldWid$ = Input$(1, #DF)
  FldDec$ = Input$(1, #DF)
  Ignore$ = Input$(14, #DF)
  FldName$ = Left$(FldName$, InStr(FldName$, Chr$(0)) - 1) + Space$(10)
  FldNam$(i, SubScript) = Left$(UCase$(RTrim$(FldName$)), 10)
  FldLen(i, SubScript) = Asc(FldWid$)
  MidPos(i, SubScript) = MMid
  MMid = MMid + FldLen(i, SubScript)
 Next i
 Close #DF
End Sub
Sub UnDelRecord(DBF$, RecNum, SubScript%)
 Record = RecNum
 If Record < 1 Or Record > Records(SubScript%) Then
  Exit Sub
 End If
 DF = FreeFile
 DelChar$ = " "
 Open DBF$ For Binary Shared As #DF
 Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + 1
 Put #DF, Pointer, DelChar$
 Close #DF
End Sub
Sub UpdateDBF(DBF$, RecNum, DL$, SubScript%)
 Record = RecNum
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + 1
 Put #DF, Pointer, DL$
 Close #DF
'==============
'UPDATE INDEXES
'==============
 Mid$(CompanyIndex, Record, 1) = UCase$(Mid$(DL$, MidPos(FLD_Company, SubScript%), 1))
 Mid$(FNameIndex, Record, 1) = UCase$(Mid$(DL$, MidPos(FLD_FName, SubScript%), 1))
 Mid$(LNameIndex, Record, 1) = UCase$(Mid$(DL$, MidPos(FLD_LName, SubScript%), 1))
End Sub
Public Function AsciiToHex(TypedData$) As String
 On Error Resume Next
 tmp$ = ""
 For i = 1 To Len(TypedData$)
  n = Asc(Mid$(TypedData$, i, 1))
  HV$ = "00" & Hex$(n)
  HV$ = Right$(HV$, 2)
  tmp$ = tmp$ & HV$
 Next i
 AsciiToHex$ = tmp$
End Function
Public Function HexToASCII(TheString) As String
 On Error Resume Next
 BYTES = Len(TheString)
 For i = 1 To BYTES
  HV = "&H" & Mid(TheString, i, 2)
  HB = CInt(HV)
  TmpStr = TmpStr & Chr(HB)
   i = i + 1
 Next
 HexToASCII = TmpStr
End Function
Public Sub LoadIndexes(DBF$, SubScript%)
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 CompanyIndex = Space$(Records(SubScript%))
 FNameIndex = Space$(Records(SubScript%))
 LNameIndex = Space$(Records(SubScript%))
 For Record = 1 To Records(SubScript%)
  Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + MidPos(FLD_Company, SubScript%)
  Seek #DF, Pointer
  FirstCharacter$ = Input$(1, #DF)
  Mid$(CompanyIndex, Record, 1) = UCase$(FirstCharacter$)
  Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + MidPos(FLD_FName, SubScript%)
  Seek #DF, Pointer
  FirstCharacter$ = Input$(1, #DF)
  Mid$(FNameIndex, Record, 1) = UCase$(FirstCharacter$)
  Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + MidPos(FLD_LName, SubScript%)
  Seek #DF, Pointer
  FirstCharacter$ = Input$(1, #DF)
  Mid$(LNameIndex, Record, 1) = UCase$(FirstCharacter$)
 Next Record
 Close #DF
End Sub
Function SendableRecord(DBF$, RecNum, SubScript%) As String
 Record = RecNum
 DF = FreeFile
 Open DBF$ For Binary Shared As #DF
 Pointer = HeaderSize(SubScript%) + ((Record - 1) * RecLen(SubScript%)) + 1
 Seek #DF, Pointer
 DataLine$ = Input$(RecLen(SubScript%), #DF)
 Close #DF
 a$ = ""
 X = 2
 For i = 1 To NumFlds(SubScript%)
  a$ = a$ & Trim$(Mid$(DataLine$, X, FldLen(i, SubScript%))) & Chr$(9)
  X = X + FldLen(i, SubScript%)
 Next i
 SendableRecord = a$
End Function
