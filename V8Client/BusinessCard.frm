VERSION 5.00
Begin VB.Form BusinessCard 
   Caption         =   " Network Business Cards"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   Icon            =   "BusinessCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BusinessCard.frx":030A
   ScaleHeight     =   8820
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton EditCardButton 
      Height          =   1935
      Left            =   5400
      Picture         =   "BusinessCard.frx":D96C4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click this button to EDIT the Business Card"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label DirectLine 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   15
      Top             =   5520
      Width           =   2490
   End
   Begin VB.Label IDStatus 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ID/Status of this contact"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   1440
      TabIndex        =   14
      Top             =   6480
      Width           =   2835
   End
   Begin VB.Label email 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email@Domain.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4440
      TabIndex        =   13
      Top             =   6240
      Width           =   2490
   End
   Begin VB.Label Fax 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   12
      Top             =   6000
      Width           =   2490
   End
   Begin VB.Label Mobile 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   11
      Top             =   5760
      Width           =   2490
   End
   Begin VB.Label Extension 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   6360
      TabIndex        =   10
      Top             =   5280
      Width           =   570
   End
   Begin VB.Label Phone 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   9
      Top             =   5280
      Width           =   1770
   End
   Begin VB.Label CityStateZip 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "City, State Zip"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   8
      Top             =   5040
      Width           =   2490
   End
   Begin VB.Label AddressLine2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line One Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   2490
   End
   Begin VB.Label AddressLine1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line One Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Width           =   2490
   End
   Begin VB.Label Website 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "www.website.here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   4080
      Width           =   5490
   End
   Begin VB.Label Department 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   5040
      Width           =   2565
   End
   Begin VB.Label JobTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   4800
      Width           =   2610
   End
   Begin VB.Label Contact 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact name here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   4560
      Width           =   2610
   End
   Begin VB.Label Company 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name Shows Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   1425
      TabIndex        =   1
      Top             =   3765
      Width           =   5430
   End
End
Attribute VB_Name = "BusinessCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EditCardButton_Click()
 EditingBusinessCard = True
 LeResult = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
 EditCardForm.Caption = "EDIT BUSINESS CARD"
 EditCardForm.Show vbModal, Me
 LeResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 EditingBusinessCard = False
 WebURL$ = Trim$(UCase$(ClientFormMain.ServerAddress))
 WebURL$ = WebURL$ & "P" & SnStr & "GetContact.asp?RecNum=" & CurrentContactRecNum
 Result$ = tb & GetUrlSource(WebURL$) & String$(99, 9)
 ContactData = Split(Result$, tb)
 Call LoadCardToScreen
 Me.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
 End If
 If KeyAscii = 13 Then
  Call EditCardButton_Click
 End If
End Sub

Private Sub Form_Load()
 LeResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 Call LoadCardToScreen
End Sub
Sub LoadCardToScreen()
 Company = Replace$(ContactData(FLD_Company), "&", "&&")
 Contact = (ContactData(FLD_MrMrsMs) & " " & ContactData(FLD_FName) & " " & ContactData(FLD_LName))
 Website = ContactData(FLD_Website)
 JobTitle = Replace$(ContactData(FLD_JobTitle), "&", "&&")
 Department = Replace$(ContactData(FLD_Department), "&", "&&")
 AddressLine1 = Replace$(ContactData(FLD_Addr1), "&", "&&")
 AddressLine2 = Replace$(Trim$(ContactData(FLD_Addr2)), "&", "&&")
 CityStateZip = ContactData(FLD_City) & ", " & ContactData(FLD_State) & " " & ContactData(FLD_Zip)
 CityStateZip = IIf(Trim(CityStateZip) = ",", "", CityStateZip)
 If AddressLine2 = "" Then
  AddressLine2 = CityStateZip
  CityStateZip = ""
 End If
 Phone = "Phone " & ContactData(FLD_Phone)
 Extension = "Ext. " & ContactData(FLD_Extension)
 DirectLine = "Direct " & ContactData(FLD_DirectLine)
 Mobile = "Mobile " & ContactData(FLD_Mobile)
 Fax = "Fax " & ContactData(FLD_Fax)
 email = ContactData(FLD_EMail)
 IDStatus = ContactData(FLD_IDStatus)
End Sub
