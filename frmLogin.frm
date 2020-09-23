VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmLogin 
   Caption         =   "Breathe SMS Sender"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   1680
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   3975
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login To Breathe"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtPassWord 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLogin_Click()
    ' Check the user has entered name and password
    If ((txtUsername.Text <> "") And (txtPassWord.Text <> "")) Then
        cmdLogin.Enabled = False
        ' Login to breathe
        Web1.Navigate2 "http://www.breathe.com/cgi-bin/login.cgi?&extension-attribute-11=" & txtUsername.Text & "&extension-attribute-12=" & txtPassWord.Text & "&SUBMIT"
        ' Make the login invisable
        Me.Visible = False
        ' let the page finish opening
        Call Sleep(4000)
        ' show main form
        Call Load(frmSend)
        frmSend.Show
        ' enable the timer to unload the login form after enough time has
        ' passed to let the user log in
        Timer1.Enabled = True
    ' Now tell the program what to do if the user does not enter any Info
    Else:
        If (txtUsername.Text = "") Then
            Call Error
            txtUsername.SetFocus
            Exit Sub
        End If
        If (txtPassWord.Text = "") Then
            Call Error
            txtPassWord.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    ' unload the form after 60 seconds
    Unload Me
End Sub
