VERSION 5.00
Begin VB.Form SplashForm 
   Caption         =   " Welcome to the V8Server for ""Network Business Cards"""
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   ClipControls    =   0   'False
   Icon            =   "SplashForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CloseMeButton 
      Cancel          =   -1  'True
      Caption         =   "Close This Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source (C)opyright V8Software.com"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"SplashForm.frx":058A
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   120
      Picture         =   "SplashForm.frx":06BD
      Top             =   120
      Width           =   8490
   End
End
Attribute VB_Name = "SplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseMeButton_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Label2 = "Open Source Code " & Chr$(169) & " Copyright 2007 - V8Software.com"
'
' OK FOLKS - THIS IS WHERE YOU CHECK TO SEE IF THE WINSOCK OCX IS ALREADY INSTALLED
'
' IF NOT - HERE'S WHERE YOU OUGHT TO INSTALL AND REGISTER IT
'
' F.W.I.W. THAT OCX IS PRETTY SMALL - SO SEE HOW I INSTALLED THE DBF AND TAKE IT FROM THERE
'
' I HAVE A TUTORIAL ON INSTALLING AN OCX FROM YOU APP AT :
'
' http://v8software.com/VBResourceFeature.doc
'
' (DON'T FORGET TO SHELL REGSVR32 TO REGISTER IT OF COURSE)
'
' Kevin Ritch, V8Software.com - check for phone numbers on our website - CALL!!!! :-)
'
' REMEMBER - This is 100% completely "open source" code!
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
 FrmMain.Show

End Sub
