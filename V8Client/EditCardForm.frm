VERSION 5.00
Begin VB.Form EditCardForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CAPTION IS UPDATED BY CALLING PROGRAM"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "EditCardForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "EditCardForm.frx":030A
   ScaleHeight     =   4470
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7200
      Top             =   3960
   End
   Begin VB.CommandButton SaveChangesAndExitButton 
      Height          =   1935
      Left            =   8040
      Picture         =   "EditCardForm.frx":13300C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Click this button to SAVE this Business Card"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   6240
      TabIndex        =   18
      Tag             =   "17"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Tag             =   "16"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   6240
      TabIndex        =   16
      Tag             =   "15"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   6240
      TabIndex        =   15
      Tag             =   "14"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   6240
      TabIndex        =   14
      Tag             =   "13"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   6240
      TabIndex        =   13
      Tag             =   "12"
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1440
      TabIndex        =   12
      Tag             =   "10"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   1440
      TabIndex        =   11
      Tag             =   "11"
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1440
      TabIndex        =   10
      Tag             =   "7"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Tag             =   "8"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Tag             =   "19"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   7
      Tag             =   "6"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Tag             =   "5"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Tag             =   "20"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Tag             =   "21"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Tag             =   "2"
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox FldValue 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Tag             =   "4"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID/STATUS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   5040
      TabIndex        =   38
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   5040
      TabIndex        =   37
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   5040
      TabIndex        =   36
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CITY"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   5040
      TabIndex        =   35
      Top             =   840
      Width           =   480
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "         "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   5040
      TabIndex        =   34
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   5040
      TabIndex        =   33
      Top             =   120
      Width           =   840
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   600
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WEBSITE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Width           =   840
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   30
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   3480
      TabIndex        =   27
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOB TITLE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MR/MRS/MS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label FldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "EditCardForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChangesMade As Boolean
Dim Loading As Boolean

Private Sub FldValue_Change(Index As Integer)
 If Loading Then
  Exit Sub
 End If
 ChangesMade = True
End Sub
Private Sub FldValue_LostFocus(Index As Integer)
 If ChangesMade Then
 '===================================================================
 'EditContactFieldByNum.asp?RecNum=55&FldNum=12&FldData=127 Test Street
 '===================================================================
  If EditingBusinessCard Then
   ChangesMade = False
   WebURL$ = ClientFormMain.ServerAddress & "P" & SnStr
   WebURL$ = WebURL$ & "EditContactFieldByNum.asp"
   WebURL$ = WebURL$ & "?RecNum=" & CurrentContactRecNum
   WebURL$ = WebURL$ & "&FldNum=" & Trim$(FldValue(Index).Tag)
   NewData$ = Trim$(FldValue(Index).Text)
   NewData$ = Replace$(NewData$, " ", "+")
   NewData$ = Replace$(NewData$, "&", "_")
   NewData$ = Replace$(NewData$, "?", "^")
   WebURL$ = WebURL$ & "&FldData=" & NewData$
   Result$ = GetUrlSource$(WebURL$)
  End If
 End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
  Exit Sub
 End If
End Sub

Private Sub SaveChangesAndExitButton_Click()
 If EditingBusinessCard Then
  Unload Me ' (Already saved)
  Exit Sub
 Else
 '======================
 'SAVE NEW BUSINESS CARD
 '======================
  If ChangesMade = False Then
   Unload Me
   Exit Sub
  End If
 '=======================================
 'CHECK SIGNIFICANT DETAILS BEFORE SAVING
 '=======================================
  CompanyStr$ = Trim$(FldValue(0).Text)
  MrMrsMsStr$ = Trim$(FldValue(1).Text)
  FNameStr$ = Trim$(FldValue(2).Text)
  LNameStr$ = Trim$(FldValue(3).Text)
  If CompanyStr$ = "" And FNameStr$ = "" And LNameStr$ = "" Then
   MsgBox "Sorry, you do need to have SOMETHING in this business card that can be indexed!" & String$(2, 10) & "Please enter a NAME or COMPANY!", vbExclamation, "Whoops!"
   Exit Sub
  End If
 '=========================================================================================================================
 'AddContact.asp?FName=Kevin&LName=Ritch&Addr1=123 Test Street&Addr2=&City=Luton&State=Beds&Zip=MK45 1JG&Phone=0171-784-138
 '=========================================================================================================================
  WebURL$ = ClientFormMain.ServerAddress & "P" & SnStr
  WebURL$ = WebURL$ & "AddContact.asp"
 '============
 'MR,MRS or MS
 '============
  NewData$ = Replace$(MrMrsMsStr$, " ", "+"):  NewData$ = Replace$(NewData$, "&", "_"):  NewData$ = Replace$(NewData$, "?", "^")
  WebURL$ = WebURL$ & "?MrMrsMs=" & NewData$
 '==========
 'FIRST NAME
 '==========
  NewData$ = Replace$(FNameStr$, " ", "+"):  NewData$ = Replace$(NewData$, "&", "_"):  NewData$ = Replace$(NewData$, "?", "^")
  WebURL$ = WebURL$ & "&FName=" & NewData$
 '========
 'LAST NAME
 '========
  NewData$ = Replace$(LNameStr$, " ", "+"):  NewData$ = Replace$(NewData$, "&", "_"):  NewData$ = Replace$(NewData$, "?", "^")
  WebURL$ = WebURL$ & "&LName=" & NewData$
 '=======
 'COMPANY
 '=======
  NewData$ = Replace$(CompanyStr$, " ", "+"):  NewData$ = Replace$(NewData$, "&", "_"):  NewData$ = Replace$(NewData$, "?", "^")
  WebURL$ = WebURL$ & "&Company=" & NewData$
  Result$ = GetUrlSource$(WebURL$)
  s = InStr(Result$, "{RECORD NUMBER")
  If s Then
   CurrentContactRecNum = Trim$(Val(Mid$(Result$, s + 14)))
   For UpDateField = 4 To 18
    WebURL$ = ClientFormMain.ServerAddress & "P" & SnStr
    WebURL$ = WebURL$ & "EditContactFieldByNum.asp"
    WebURL$ = WebURL$ & "?RecNum=" & CurrentContactRecNum
    WebURL$ = WebURL$ & "&FldNum=" & Trim$(FldValue(UpDateField).Tag)
    NewData$ = Trim$(FldValue(UpDateField).Text)
    NewData$ = Replace$(NewData$, " ", "+")
    NewData$ = Replace$(NewData$, "&", "_")
    NewData$ = Replace$(NewData$, "?", "^")
    WebURL$ = WebURL$ & "&FldData=" & NewData$
    Result$ = GetUrlSource$(WebURL$)
   Next UpDateField
  '=======================
  'LOAD THE CARD TO MEMORY
  '=======================
   AutoListing$ = Left$((Trim$(FldValue(2).Text) & " " & Trim$(FldValue(3).Text) & Space$(30)), 30) & " " & Trim$(FldValue(0).Text) & String$(5, 9) & CurrentContactRecNum
   CreatedNewBusinessCard = True
   Unload Me
  End If
 End If
End Sub

Private Sub Form_Load()
 ChangesMade = False
 LeResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'======================================================================================
'The FldValue Textboxes (Indexed array) use the TAG property for Physical Field Numbers
'======================================================================================
 On Error Resume Next
 If EditingBusinessCard Then
  Loading = True
  For i = 0 To 99 ' Allow for 99 fields
   SS = Val(FldValue(i).Tag)
   FldValue(i).Text = ContactData(SS)
  Next i
 End If
 Loading = False
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
 On Error Resume Next
 FldValue(0).SetFocus
End Sub
