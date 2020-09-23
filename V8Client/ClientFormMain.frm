VERSION 5.00
Begin VB.Form ClientFormMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   " Network Business Cards "
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "ClientFormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "ClientFormMain.frx":030A
   ScaleHeight     =   7980
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4440
      Top             =   4080
   End
   Begin VB.CommandButton CreateNewCardButton 
      Height          =   2535
      Left            =   6120
      Picture         =   "ClientFormMain.frx":55B4C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click this button to CREATE A NEW Business Card"
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   8175
   End
   Begin VB.CommandButton LookupButton 
      Height          =   735
      Left            =   5040
      Picture         =   "ClientFormMain.frx":6410E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "LOOKUP BUTTON"
      Top             =   4560
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Matches"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox SearchText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   4200
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contains"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   3240
      Width           =   930
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Starts with"
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   2880
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.TextBox ServerAddress 
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
      Left            =   360
      TabIndex        =   8
      Text            =   "http://192.168.0.3:6954/"
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   25
      Left            =   7800
      TabIndex        =   38
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   24
      Left            =   7560
      TabIndex        =   37
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   23
      Left            =   7320
      TabIndex        =   36
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   22
      Left            =   7080
      TabIndex        =   35
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   21
      Left            =   6840
      TabIndex        =   34
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   20
      Left            =   6600
      TabIndex        =   33
      Top             =   5040
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   19
      Left            =   6360
      TabIndex        =   32
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   18
      Left            =   6120
      TabIndex        =   31
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   17
      Left            =   8040
      TabIndex        =   30
      Top             =   4800
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   16
      Left            =   7800
      TabIndex        =   29
      Top             =   4800
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   15
      Left            =   7560
      TabIndex        =   28
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   14
      Left            =   7320
      TabIndex        =   27
      Top             =   4800
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   13
      Left            =   7080
      TabIndex        =   26
      Top             =   4800
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   12
      Left            =   6840
      TabIndex        =   25
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   11
      Left            =   6600
      TabIndex        =   24
      Top             =   4800
      Width           =   120
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   10
      Left            =   6360
      TabIndex        =   23
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   6120
      TabIndex        =   22
      Top             =   4800
      Width           =   105
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   8
      Left            =   8040
      TabIndex        =   21
      Top             =   4560
      Width           =   75
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   7800
      TabIndex        =   20
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   7560
      TabIndex        =   19
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   7320
      TabIndex        =   18
      Top             =   4560
      Width           =   120
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   7080
      TabIndex        =   17
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   6840
      TabIndex        =   16
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   6600
      TabIndex        =   15
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   14
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label LETTER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   6120
      TabIndex        =   13
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Double click your mouse on any person's business card to open it."
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7680
      Width           =   8175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lookup text"
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please set the correct URL into this ""ServerAddress"" Field"
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
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lookup using"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "ClientFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LookupResultText





Private Sub CreateNewCardButton_Click()
 EditingBusinessCard = False
 CreatedNewBusinessCard = False
 LeResult = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
 EditCardForm.Caption = "CREATE A NEW BUSINESS CARD"
 EditCardForm.Show vbModal, Me
 LeResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 EditingBusinessCard = False
 If CreatedNewBusinessCard Then
  List2.Clear
  List2.AddItem AutoListing$
  List2.ListIndex = 0
  Call List2_DblClick
 End If
End Sub

Private Sub Form_Load()
 If App.PrevInstance Then
  End
 End If
 tb = Chr$(9)
'=======================================
'COLLECT SNSTR (HARD DISK SERIAL NUMBER)
'=======================================
 Call HardDiskSerial
'=======================================
'POPULATE LIST WITH LOOKUP FIELD OPTIONS
'=======================================
 List1.AddItem "First Name" & String$(5, 9) & "LookupFName.asp?SearchString="
 List1.AddItem "Last Name" & String$(5, 9) & "LookupLName.asp?SearchString="
 List1.AddItem "Company" & String$(5, 9) & "LookupCompany.asp?SearchString="
 List1.AddItem "ID/Status" & String$(5, 9) & "LookupIDStatus.asp?SearchString="
 List1.ListIndex = 1
'=================================
'THE RESULT LIST (LIST2) IS SORTED
'=================================
End Sub

Private Sub LETTER_Click(Index As Integer)
 Option1.Value = True
 SearchText = Chr$(Index + 65)
 Call LookupButton_Click
End Sub

Private Sub List2_DblClick()
 a$ = List2
 s = InStr(a$, tb) + 4
 CurrentContactRecNum = Right$(a$, Len(a$) - s)
 WebURL$ = Trim$(UCase$(ServerAddress))
 WebURL$ = WebURL$ & "P" & SnStr & "GetContact.asp?RecNum=" & CurrentContactRecNum
'=========================================================================================
'PAD the result with a preceding tab and 20 more to move subscript up one to match globals
'=========================================================================================
 Result$ = tb & GetUrlSource$(WebURL$) & String$(20, 9)
 ContactData = Split(Result$, tb)
 BusinessCard.Show vbModal, Me
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If List2.ListCount > 0 Then
   Call List2_DblClick
  End If
 End If
End Sub

Private Sub LookupButton_Click()
 SearchText = UCase$(Trim$(SearchText))
 If SearchText = "" Then
  MsgBox "Sorry, you do need to enter something for me to search for!", vbApplicationModal + vbExclamation, "Whoops!"
  On Error Resume Next
  SearchText.SetFocus
  Exit Sub
 End If
 LookupType = "StartsWith"
 If Option2 Then
  LookupType = "Contains"
 End If
 If Option3 Then
  LookupType = "ExactMatch"
 End If
 WebURL$ = Trim$(UCase$(ServerAddress))
 If InStr(WebURL$, ":") = 0 Then
  BadURL = True
 End If
 If Left$(WebURL$, 7) <> "HTTP://" Then
  BadURL = True
 End If
 If Right$(WebURL$, 1) <> "/" Then
  BadURL = True
 End If
 If BadURL Then
  MsgBox "Sorry, the ServerAddress Field is incorrectly formatted." & String$(2, 10) & "Please check with the administrator or visit V8Software.com for assistance!", vbApplicationModal + vbExclamation, "Whoops!"
  On Error Resume Next
  ServerAddress.SelStart = 0
  ServerAddress.SelLength = Len(ServerAddress)
  ServerAddress.SetFocus
  Exit Sub
 End If
 WebURL$ = WebURL$ & "P" & SnStr
 Y = List1.ListIndex
 FT$ = List1.List(Y)
 s = InStr(FT$, Chr$(9)) + 4
 Mid$(FT$, 1, s) = Space$(s)
 WebURL$ = WebURL$ & Trim$(FT$) & SearchText & "&LookupType=" & LookupType
'=========================================
'EXAMPLES OF ACCEPTABLE ASP STYLE REQUESTS
'=========================================
'GetContact.asp?RecNum=55
'AddContact.asp?FName=Kevin&LName=Ritch&Addr1=123 Test Street&Addr2=&City=Luton&State=Beds&Zip=MK45 1JG&Phone=0171-784-138
'========================================================================================================
'Editing a record is ALWAYS done on a field-by-field basis. (In Client App - On TextBox(Index) LOSTFOCUS)
'========================================================================================================
'EditContactField.asp?RecNum=55&FldNam=Addr1&FldData=127 Test Street
'========================================================================================================
'DeleteContact.asp?RecNum=55
'LookupCompany.asp?SearchString=Software&LookupType=Contains
'LookupFName.asp?SearchString=Alex&LookupType=StartsWith
'LookupLName.asp?SearchString=Ritch&LookupType=ExactMatch
'LookupIDStatus.asp?SearchString=Director&LookupType=Contains
 Result$ = GetUrlSource$(WebURL$)
 LookupResultText = Split(Result$, vbCrLf)
 List2.Clear
 For i = 1 To UBound(LookupResultText)
  List2.AddItem LookupResultText(i - 1)
 Next i
 If List2.ListCount Then
  List2.ListIndex = 0
  If List2.ListCount = 1 Then
   Call List2_DblClick
   Exit Sub
  End If
  On Error Resume Next
  List2.SetFocus
 Else
  MsgBox "Sorry, no Business Cards have been found that match your LOOKUP criteria!" & String$(2, 10) & "Please try again.", vbApplicationModal + vbExclamation, "NO RECORDS MATCH"
 ' On Error Resume Next
 ' SearchText.SelStart = 0
 ' SearchText.SelLength = Len(SearchText)
 ' SearchText.SetFocus
 ' Exit Sub
 End If
 
End Sub


Private Sub SearchText_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Call LookupButton_Click
 End If
 KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
 On Error Resume Next
 Me.SearchText.SetFocus
End Sub
