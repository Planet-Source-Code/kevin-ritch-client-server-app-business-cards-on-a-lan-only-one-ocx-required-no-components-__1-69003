VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   " V8Server - Managing your    ""NETWORK BUSINESS CARDS"""
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmMain.frx":08CA
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DisplayAddress 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "{Stopped - Please Start Server}"
      Top             =   2400
      Width           =   6975
   End
   Begin VB.TextBox Incoming 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C80B1&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3960
      Width           =   7095
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "Start Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Sport 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Text            =   "6954"
      Top             =   720
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   120
   End
   Begin MSWinsockLib.Winsock WinSock1 
      Index           =   0
      Left            =   6360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No OCXs or DLLs are required for your Client Applications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   645
      TabIndex        =   7
      Top             =   4920
      Width           =   5970
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You may copy this string into your Client Application ""ServerAddress"" Field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   6360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PORT #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   705
   End
   Begin VB.Label StatusLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS : STOPPED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   3555
   End
   Begin VB.Image Running 
      Height          =   480
      Left            =   720
      Picture         =   "FrmMain.frx":10ADC
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmMain.frx":113A6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   7065
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ServerResetTimer As Integer
Dim RecordNumber As Long
Dim ContactRecord As String
Dim TheASPQueryString(999, 2) ' Str NAME and VALUE - up to 99 Variables!
Dim TheASPQueryStringSubscriptCounter As Integer
Dim UserList As String
Dim UserRequest$(999)
Dim ServerPort As Long

Private Sub StartButton_Click()
 StartButton.Enabled = False
'========================================
'Define the DBF File used for the Rolodex
'========================================
 DBF_File = "C:\Program Files\V8Server\Data\Rolodex.dbf"
 Call GetDBStructure(DBF_File, BusinessCardTable)
 Call LoadIndexes(DBF_File, BusinessCardTable)
'=======================================================
 Sport.Enabled = False ' Once run - no can change Server Port!
 ServerPort = Val(Sport)
 DisplayAddress = "http://" & Trim$(WinSock1(0).LocalIP) & ":" & Trim$(ServerPort) & "/"
 WinSock1(0).Bind ServerPort
 WinSock1(0).Listen
 Running.Visible = True
 StatusLabel.Caption = "STATUS : ON LINE"
 StatusLabel.ForeColor = &HFFFF00
 Label12.Visible = True
End Sub
Private Sub Form_Load()
 If App.PrevInstance Then End
 tb = Chr$(9)
 Call SetupSystem
End Sub
Private Sub Form_Unload(Cancel As Integer)
 End
End Sub
Private Sub WinSock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
 WinSock1(Index).Close
 WinSock1(Index).Accept requestID
 Load WinSock1(Index + 1)
 WinSock1(Index + 1).Bind ServerPort
 WinSock1(Index + 1).Listen
End Sub
Private Sub WinSock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 '=========================================================================
 'BECAUSE THIS GENERIC SOFTWARE HANDLES URLS OF *ANY LENGTH* STRINGS MAY BE
 'SENT IN SEGMENTS & THEN SEQUENTIALLY CONCATENATED.
 '=========================================================================
 '
 '                          HEADER DESCRIPTION:
 '
 ' EACH REQUEST IS PRECEDED BY A 10 BYTE HEADER - DESCRIBED AS FOLLOWS:
 '
 ' BYTE 1 - DEFINES PORTION OF URL AS COMPLETE, NEW, APPEND OR PROCESS!
 '
 '         "C" = THIS IS A COMPLETE COMMAND STRING. PROCESS IMMEDIATELY
 '         "N" = THIS IS A NEW COMMAND STRING FROM USER, MORE TO FOLLOW
 '         "A" = APPEND TO EXISTING COMMAND STRING. MORE DATA TO FOLLOW
 '         "P" = APPEND STRING AND FINALLY, GO AND PROCESS THIS COMMAND
 '
 ' BYTES 2 to 10 (9 BYTES) COMPRISES THE HARD DISK SERIAL NUMBER OF USER
 '
 Dim str As String
 WinSock1(Index).GetData str
 Dim temp As String
 str = Replace(str, "+", " ")
 LTest$ = "GET /"
 RTest$ = " HTTP/1.1"
 If Left$(str, 5) = LTest$ And InStr(str, RTest$) Then
  s = InStr(str, RTest$)
  str = Left$(str, s - 1)
  temp = Right$(str, Len(str) - 4)
 Else
  GoTo SendTheReply:
 End If
'====================================
'CHECK IF THE COMMAND REQUEST IS REAL
'====================================
 R$ = temp
 If R$ <> "" Then
  If Left$(R$, 1) = "/" Then
   R$ = Right$(R$, Len(R$) - 1)
  End If
 End If
 Reply$ = "{NONE}"
 RealRequest = Len(R$) >= 10
 If RealRequest Then
 '===================
 'PROCESS THE REQUEST
 '===================
  CMD_Type$ = Mid$(R$, 1, 1)
  CMD_User$ = Mid$(R$, 2, 9)
  CMD_String$ = Right$(R$, Len(R$) - 10)
 '=============================================================
 'ESTABLISH THE USER'S SUBSCRIPT USING PRETTY SIMPLE ARITHMETIC
 '=============================================================
 '
 ' EACH USER = THEIR HARD-DISK SERIAL NUMBER (9 BYTES) & CHR$(9)
 '
 ' EACH *NEW* INCOMING USER IS ADDED TO THE STRING "USERLIST".
 '
 ' IT STANDS TO REASON THAT WITH THE INSTR COMMAND, WE CAN LOCATE
 ' THE POSITION AND CONSEQUENTLY, THE SEQUENTIAL "CURRENT REQUEST"
 ' USER SUBSCRIPT FOR CONCATENTION OF THEIR REQUEST STRINGS.
 '
 ' WINSOCK HAS A STRING LIMITATION OF, I BELIEVE, 8096 BYTES.
 '
 ' SO WHAT? ALL THIS MEANS IS THAT THE CLIENT SOFTWARE MAY NEED
 ' TO SUBMIT MULTIPLE CHUNKS OF DATA FOR INCREDIBLY LONG STRINGS.
 '
 ' THIS V8SERVER SOFTWARE HAS NO URL "STRING LENGTH LIMITATION".
 '
 ' SO USER COMMANDS MAY BE SENT IN PIECES UNTIL THE WHOLE STRING
 ' HAS BEEN SUBMITTED AND IS READY FOR PROCESSING.
 '
  UserSubscript = 0
  UserCount = Len(UserList) + 9
  UserCount = Fix(UserCount / 10)
 '==============================================================
 'OK - EITHER FIND OR ADD THIS USER AND ESTABLISH USER SUBSCRIPT
 '==============================================================
  FindUser$ = CMD_User$ & tb ' 10 BYTES EXACTLY
  s = InStr(UserList, FindUser$)
  If s Then ' USER IS IN THE LIST
   UserSubscript = s + 9
   UserSubscript = Fix(UserSubscript / 10)
  Else ' ADD THIS USER TO LIST OF USERS
   UserList = UserList + FindUser$
   UserSubscript = UserSubscript + 1
  End If
 '===================================
 'DEAL WITH THE REQUEST FROM THE USER
 '===================================
 '
 ' EXAMPLE OF TWO LEGITIMATE INCOMING URLS (WE WOULD NOT ACTUALLY TRUNCATE IF THIS SMALL IN REALITY)
 '
 '  IPADDRESS + PORT + Type of command + Hard Disk Serial # + SEGMENT OF(ASP PAGE NAME + VARIABLES)
 '
 ' "http://192.168.0.6:6954/N1234-5432EditContactField.asp?RecNum=5&FldName=Mobile&FldData=(631)"
 ' "http://192.168.0.6:6954/P1234-5432 961-0594"
 '
 ' 1ST STRING IS CMD_TYPE "N" - New, more to follow (In example, missing rest of the Mobile number)
 ' 2ND STRING IS CMD_TYPE "P" - Process the whole string now! (Example includes balance Mobile #)
 '
 ' ========================================
 ' THE FINAL DATA READY TO PROCESS BECOMES:
 ' ========================================
 ' CMD_USER = "1234-5432"
 ' CMD_String = "EditContactField.asp?RecNum=5&FldName=Phone&FldData=(631) 961-0594"
 ' =================================================================================
 '
  Incoming = R$ ' (Displays the incoming data string on screen)
  ProcessNow = False
  Select Case CMD_Type$
   Case "C" ' THIS IS A COMPLETE STRING - PROCESS IMMEDIATELY
    ProcessNow = True
   Case "N" ' THIS IS A NEW COMMAND STRING FOR USER (MORE TO FOLLOW)
    UserRequest(UserSubscript) = CMD_String$
    Reply$ = "{OK}"
   Case "A" ' APPEND TO COMMAND STRING - MORE DATA TO FOLLOW
    UserRequest(UserSubscript) = UserRequest(UserSubscript) & CMD_String$
    Reply$ = "{OK}"
   Case "P" ' APPEND STRING AND PROCESS RIGHT AWAY
    CMD_String$ = UserRequest(UserSubscript) & CMD_String$
    ProcessNow = True
   Case Else
    Reply$ = "{CANCELLED}"
  End Select
 '===========================================================================
 'IF THE STRING IS READY TO USE - PROCESS THAT STRING AND FORMAT THE RESPONSE
 '===========================================================================
  If ProcessNow Then
   Reply$ = ProcessTheRequest(CMD_String$)
  End If
 End If ' REQUEST STRING LEN => 10
'====================================================================
'SEND THE REPLY TEXT TO THE USER OF THE CURRENT INSTANCE OF A WINSOCK
'====================================================================
SendTheReply:
 HTMLString$ = "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & "Content-Length: " & Len(Reply$) & vbCrLf & vbCrLf & Reply$ '& vbCrLf & vbCrLf & vbCrLf
 WinSock1(Index).SendData HTMLString$
'====================================================================
End Sub
Private Sub Timer1_Timer()
 ServerResetTimer = ServerResetTimer + 1
 If ServerResetTimer > 255 Then
  Call ResetTheServer
  ServerResetTimer = 0
 End If
End Sub
Sub ResetTheServer()
 If WinSock1.ubound <> 0 Then
  For i = 1 To WinSock1.ubound
   WinSock1(i).Close
   Unload WinSock1(i)
  Next i
  WinSock1(0).Close
  WinSock1(0).Listen
 End If
End Sub
Sub SetupSystem()
 On Error Resume Next
 MkDir "c:\Program Files\"
 MkDir "c:\Program Files\V8Server"
 MkDir "c:\Program Files\V8Server\System"
 MkDir "c:\Program Files\V8Server\Data"
 If Dir("c:\Program Files\V8Server\Data\Rolodex.dbf") = "" Then
  Call InstallBlankRolodexDatabase
 End If
End Sub
Function ProcessTheRequest(TheCommandString As String) As String
'=================================================================
'FOR ASP PROGRAMMERS, THIS IS GOING TO BE A BLOOIN' PIECE OF CAKE!
'
' NB: A COMPLETE AND GROOVY URL LOOKS SOMETHING LIKE THIS:
'
' http://192.168.0.6:6954/P1234-5678DeleteContact.asp?RecNum=55
'
' EVERYTHING UP TO THE LAST "/" WAS ALREADY STRIPPED OFF WHEN THE REQUEST ARRIVED
'
' =========================================================================================
' PLEASE DO GIVE THE DESCRIPTION OF THE HEADER A 'BUTCHER'S HOOK' in Sub WinSock1_DataArrival
' =========================================================================================
'
' =========================================
' EXAMPLES OF ACCEPTABLE ASP STYLE REQUESTS
' =========================================
'
' GetContact.asp?RecNum=55
' AddContact.asp?FName=Kevin&LName=Ritch&Addr1=123 Test Street&Addr2=&City=Luton&State=Beds&Zip=MK45 1JG&Phone=0171-784-138
' ========================================================================================================
' Editing a record is ALWAYS done on a field-by-field basis. (In Client App - On TextBox(Index) LOSTFOCUS)
' ========================================================================================================
' EditContactField.asp?RecNum=55&FldNam=Addr1&FldData=127 Test Street
' EditContactFieldByNum.asp?RecNum=55&FldNum=12&FldData=127 Test Street
' ========================================================================================================
' DeleteContact.asp?RecNum=55
' LookupCompany.asp?SearchString=Software&LookupType=Contains
' LookupFName.asp?SearchString=Alex&LookupType=StartsWith
' LookupLName.asp?SearchString=Ritch&LookupType=ExactMatch
' LookupIDStatus.asp?SearchString=Director&LookupType=Contains
'
' LOOKUPS ARE *NOT* CASE SENSITIVE. (ALL TREATED AS UPPER CASE)
'
'=======================================================================
'ABOUT DBF FILE DATA PROCESSING
'==============================
'EACH TIME A RECORD IS WORKED ON, THE FULL RECORD IS LOADED FROM THE DBF
'DBASE STORES DATA IN FIXED RECORD LENGTHS. EACH RECORD STARTS WITH " ".
'IN DBASE, WHEN A RECORD IS DELETED, THE FIRST BYTE IS REPLACED WITH "*"
'=======================================================================
'TO GET THE EXACT SHAPE OF THE DATABASE AND NUMBER OF RECORDS
'============================================================
' Call GetDBStructure(DBF_File, BusinessCardTable) ' CALLED ON STARTUP
'================================================
'EXTRACT THE REQUIREMENTS FROM THE COMMAND STRING
'================================================
 ProcessTheRequest = "{ERROR IN URL}" ' DEFAULT IS FAILURE (UNLESS CORRECT:-)
 On Error GoTo BadlyFormedURL:
'===============================
'STEP 1 - GET THE COMMAND STRING
'===============================
 a$ = Trim$(TheCommandString)
'==========================
'DEFINE THE "ASP PAGE" NAME
'==========================
 a$ = Replace$(a$, "?", "&")
 Q = InStr(a$, "&")
 ASPPage$ = UCase$(Left$(a$, Q - 1)) ' e.g. GetContact.asp
'=========================== ======================================================
'Extract submitted variables (Prepare program for our Request.QueryString(Whatever)
'=========================== ======================================================
 a$ = "&" & Right$(a$, Len(a$) - Q)
 TheASPQueryStringSubscriptCounter = 0
 a$ = a$ & "&"
 While InStr(a$, "&") > 0 And InStr(a$, "=") > 0
  S1 = InStr(a$, "&")
  S2 = InStr(S1, a$, "=")
  S3 = InStr(S2, a$, "&")
  TheASPQueryStringSubscriptCounter = TheASPQueryStringSubscriptCounter + 1
  TheASPQueryString(TheASPQueryStringSubscriptCounter, 1) = UCase$(Mid$(a$, S1 + 1, (S2 - (S1 + 1))))
  TheASPQueryString(TheASPQueryStringSubscriptCounter, 2) = Mid$(a$, S2 + 1, (S3 - (S2 + 1)))
 '===================================================================================
 'ALL SUBMITTED AMPERSANDS (IN TYPE DATA) ARE SENT AS UNDERSCORES, SO SWAP THEM OVER!
 '===================================================================================
  TheASPQueryString(TheASPQueryStringSubscriptCounter, 2) = Replace$(TheASPQueryString(TheASPQueryStringSubscriptCounter, 2), "_", "&")
 '===================================================================================
 'ALL SUBMITTED QUESTION MARKS (IN TYPED DATA) ARE SENT AS CARETS, SO SWAP THEM OVER!
 '===================================================================================
  TheASPQueryString(TheASPQueryStringSubscriptCounter, 2) = Replace$(TheASPQueryString(TheASPQueryStringSubscriptCounter, 2), "^", "?")
  a$ = Right$(a$, Len(a$) - (S3 - 1))
 Wend
'========================================================
'SET THE CURRENT CONTACT SIMPLY AS A BLANK STRING FOR NOW
'========================================================
 ContactRecord = Space$(RecLen(1))
 RecordNumber = 0 ' We don't know what the physical dbase record number is yet!
'=================================
 ResponseDotWrite$ = ""
 Dim TheFieldNumber As Integer
 Select Case UCase$(ASPPage$)
 
  Case "GETCONTACT.ASP" ' GetContact.asp?RecNum=55
   Record = Val(RequestDotQueryString("RecNum"))
   ResponseDotWrite$ = SendableRecord(DBF_File, Record, BusinessCardTable)
 
  Case "ADDCONTACT.ASP"
  '===========================================================
  'Extract Data from submission and insert into a blank record
  '===========================================================
   Call InsertField(FLD_MrMrsMs, RequestDotQueryString("MrMrsMs"), BusinessCardTable)
   Call InsertField(FLD_FName, RequestDotQueryString("FName"), BusinessCardTable)
   Call InsertField(FLD_LName, RequestDotQueryString("LName"), BusinessCardTable)
   Call InsertField(FLD_Company, RequestDotQueryString("Company"), BusinessCardTable)
   Call InsertField(FLD_Phone, RequestDotQueryString("Phone"), BusinessCardTable)
   Call InsertField(FLD_Extension, RequestDotQueryString("Extension"), BusinessCardTable)
   Call InsertField(FLD_Fax, RequestDotQueryString("Fax"), BusinessCardTable)
   Call InsertField(FLD_Mobile, RequestDotQueryString("Mobile"), BusinessCardTable)
   Call InsertField(FLD_Home, RequestDotQueryString("Home"), BusinessCardTable)
   Call InsertField(FLD_EMail, RequestDotQueryString("EMail"), BusinessCardTable)
   Call InsertField(FLD_Website, RequestDotQueryString("Website"), BusinessCardTable)
   Call InsertField(FLD_Addr1, RequestDotQueryString("Addr1"), BusinessCardTable)
   Call InsertField(FLD_Addr2, RequestDotQueryString("Addr2"), BusinessCardTable)
   Call InsertField(FLD_City, RequestDotQueryString("City"), BusinessCardTable)
   Call InsertField(FLD_State, RequestDotQueryString("State"), BusinessCardTable)
   Call InsertField(FLD_Zip, RequestDotQueryString("Zip"), BusinessCardTable)
   Call InsertField(FLD_IDStatus, RequestDotQueryString("IDStatus"), BusinessCardTable)
  '======================
  'ADD NEW CONTACT RECORD
  '======================
   RecordNumber = AppendToDBF(DBF_File, ContactRecord, BusinessCardTable)
   ResponseDotWrite$ = "{RECORD NUMBER " & Trim(RecordNumber) & " ADDED}"
   
  Case "EDITCONTACTFIELD.ASP" ' EditContactField.asp?RecNum=55&FldName=Addr1&FldData=127 Test Street
   Record = Val(RequestDotQueryString("RecNum"))
   TheFieldName$ = RequestDotQueryString("FldName")
   TheFieldNumber = EvaluateFieldNumber(TheFieldName$, BusinessCardTable)
   ContactRecord = GetRecord(DBF_File, Record, BusinessCardTable)
   Call InsertField(TheFieldNumber, RequestDotQueryString("FldData"), BusinessCardTable)
  '=====================
  'UPDATE CONTACT RECORD
  '=====================
   Call UpdateDBF(DBF_File, Record, ContactRecord, BusinessCardTable)
   ResponseDotWrite$ = "{UPDATED}"
   
  Case "EDITCONTACTFIELDBYNUM.ASP" ' EditContactFieldByNum.asp?RecNum=55&FldNum=12&FldData=127 Test Street
   Record = Val(RequestDotQueryString("RecNum"))
   TheFieldNumber = RequestDotQueryString("FldNum")
   ContactRecord = GetRecord(DBF_File, Record, BusinessCardTable)
   Call InsertField(TheFieldNumber, RequestDotQueryString("FldData"), BusinessCardTable)
  '=====================
  'UPDATE CONTACT RECORD
  '=====================
   Call UpdateDBF(DBF_File, Record, ContactRecord, BusinessCardTable)
   ResponseDotWrite$ = "{UPDATED}"
   
  Case "DELETECONTACT.ASP" ' DeleteContact.asp?RecNum=55
   Record = Val(RequestDotQueryString("RecNum"))
   Call DeleteRecord(DBF_File, Record, BusinessCardTable)
   ResponseDotWrite$ = "{DELETED}"
   
  Case "LOOKUPFNAME.ASP"
   ResponseDotWrite$ = LookupList("FNAME", RequestDotQueryString("SearchString"), RequestDotQueryString("LookupType"))
  
  Case "LOOKUPLNAME.ASP"
   ResponseDotWrite$ = LookupList("LNAME", RequestDotQueryString("SearchString"), RequestDotQueryString("LookupType"))
  
  Case "LOOKUPCOMPANY.ASP"
   ResponseDotWrite$ = LookupList("COMPANY", RequestDotQueryString("SearchString"), RequestDotQueryString("LookupType"))
  
  Case "LOOKUPIDSTATUS.ASP"
   ResponseDotWrite$ = LookupList("IDSTATUS", RequestDotQueryString("SearchString"), RequestDotQueryString("LookupType"))
   
  Case Else
   ResponseDotWrite$ = "{ERROR 404}"
 End Select
 ProcessTheRequest = ResponseDotWrite$
 Exit Function
BadlyFormedURL:
 ProcessTheRequest = "{ERROR IN URL}"
End Function
Function RequestDotQueryString(FindString As String) As String
 FindString = UCase$(FindString)
 For i = 1 To TheASPQueryStringSubscriptCounter
  If TheASPQueryString(i, 1) = FindString Then
   RequestDotQueryString = TheASPQueryString(i, 2)
   Exit For
  End If
 Next i
End Function
Sub InsertField(FieldNumber As Integer, FieldValue As String, TableNumber As Integer)
 FieldValue = FieldValue & Space$(FldLen(FieldNumber, TableNumber))
 Mid$(ContactRecord, MidPos(FieldNumber, TableNumber), FldLen(FieldNumber, TableNumber)) = Left$(FieldValue, FldLen(FieldNumber, TableNumber))
End Sub
Function EvaluateFieldNumber(TheFieldName As String, SubScript As Integer) As Integer
 TheFieldName = Trim$(UCase$(TheFieldName))
 For i = 1 To NumFlds(SubScript)
  If TheFieldName = FldNam$(i, SubScript) Then
   EvaluateFieldNumber = i
   Exit For
  End If
 Next i
End Function
Function LookupList(SearchField As String, SearchString As String, LookupType As String) As String
 SearchString = Trim$(UCase$(SearchString))
 LS = Len(SearchString)
 If LS = 0 Then
  Exit Function
 End If
 FirstLetter$ = Left$(SearchString, 1)
 Call GetDBFRecordCount(DBF_File, BusinessCardTable)
 If Records(BusinessCardTable) = 0 Then
  Exit Function
 End If
 LookupType = UCase$(LookupType)
 SearchField = UCase$(SearchField)
 Select Case SearchField
  Case "FNAME"
   IndexString = FNameIndex
   FldToCheck = FLD_FName
  Case "LNAME"
   IndexString = LNameIndex
   FldToCheck = FLD_LName
  Case "COMPANY"
   IndexString = CompanyIndex
   FldToCheck = FLD_Company
  Case "IDSTATUS"
   IndexString = Chr$(255)
   FldToCheck = FLD_IDStatus
  Case Else
   Exit Function
 End Select
'===================================================================================
'IF AN INDEX WAS EXPECTED BUT A "CONTAINS" SEARCH HAS BEEN REQUESTED - THEN NO INDEX
'===================================================================================
 IndexString = IIf(LookupType$ = "CONTAINS", Chr$(255), IndexString)
 IsIndexed = IndexString <> Chr$(255)
'===================================================================================
 Result$ = ""
 DF = FreeFile
 Open DBF_File For Binary Shared As #DF
 For Record = 1 To Records(BusinessCardTable)
  IsMatch = False
 '========================
 'IS THIS AN INDEX SEARCH?
 '========================
  If IsIndexed Then
   If Mid$(IndexString, Record, 1) <> FirstLetter$ Then
    GoTo NextRecord:
   End If
  End If
 '==================
 'IS RECORD DELETED?
 '==================
  Pointer = HeaderSize(BusinessCardTable) + ((Record - 1) * RecLen(BusinessCardTable)) + 1
  Seek #DF, Pointer
  Test$ = Input$(1, #DF)
  If Test$ = "*" Then ' Deleted!
   GoTo NextRecord:
  End If
 '=========================
 'EXTRACT COMPARISON STRING
 '=========================
  Pointer = HeaderSize(BusinessCardTable) + ((Record - 1) * RecLen(BusinessCardTable)) + MidPos(FldToCheck, BusinessCardTable)
  Seek #DF, Pointer
  Test$ = Trim$(UCase$(Input$(FldLen(FldToCheck, BusinessCardTable), #DF)))
 '==================================
 'APPLY THE TYPE OF LOOKUP REQUESTED
 '==================================
  Select Case LookupType$
   Case "STARTSWITH"
    IsMatch = Left$(Test$, LS) = SearchString
   Case "EXACTMATCH"
    IsMatch = Test$ = SearchString
   Case "CONTAINS"
    IsMatch = InStr(Test$, SearchString) > 0
   Case Else
    Close #DF
    Exit Function
  End Select
 '========================================
 'IF WE HAVE A MATCH - ADD TO CONTACT LIST
 '========================================
  If IsMatch Then
   Pointer = HeaderSize(BusinessCardTable) + ((Record - 1) * RecLen(BusinessCardTable)) + MidPos(FLD_FName, BusinessCardTable)
   Seek #DF, Pointer
   FirstName$ = Trim$(Input$(FldLen(FLD_FName, BusinessCardTable), #DF))
   Pointer = HeaderSize(BusinessCardTable) + ((Record - 1) * RecLen(BusinessCardTable)) + MidPos(FLD_LName, BusinessCardTable)
   Seek #DF, Pointer
   LastName$ = Trim$(Input$(FldLen(FLD_LName, BusinessCardTable), #DF))
   Pointer = HeaderSize(BusinessCardTable) + ((Record - 1) * RecLen(BusinessCardTable)) + MidPos(FLD_Company, BusinessCardTable)
   Seek #DF, Pointer
   Company$ = Input$(FldLen(FLD_Company, BusinessCardTable), #DF)
   Disp$ = Left$((FirstName$ & " " & LastName$ & Space$(30)), 30) & " " & Company$ & String$(5, 9) & Trim$(Record) & vbCrLf
   Result$ = Result$ & Disp$
  End If
NextRecord:
 Next Record
 Close #DF
 LookupList = Result$
End Function
Private Sub InstallBlankRolodexDatabase()
 Dim bytResourceData()   As Byte
 bytResourceData = LoadResData(101, "CUSTOM")
 Open "c:\Program Files\V8Server\Data\Rolodex.dbf" For Binary Shared As #1
 Put #1, 1, bytResourceData
 Close
End Sub

