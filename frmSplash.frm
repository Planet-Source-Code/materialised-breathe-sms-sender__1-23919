VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer2 
         Interval        =   18000
         Left            =   1200
         Top             =   3480
      End
      Begin VB.Timer Timer1 
         Interval        =   22000
         Left            =   360
         Top             =   3480
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash1 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
         _cx             =   22818619
         _cy             =   22811423
         Movie           =   "http://jason-n3xt.org/Five-Times/images/5tp.swf"
         Src             =   "http://jason-n3xt.org/Five-Times/images/5tp.swf"
         WMode           =   "Window"
         Play            =   0   'False
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   -1  'True
         BGColor         =   ""
         SWRemote        =   ""
         Stacking        =   "below"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Presents"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "ur website here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   3600
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public r As Integer
Private Sub Form_Load()
    ' Play Movie and configure unload timer
    flash1.Play
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Change the colour of the text on the label when mouse is taken off
    Label1.ForeColor = vbWhite
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Change the colour of the text on the label when mouse is taken off
    Label1.ForeColor = vbWhite
End Sub

Private Sub Label1_Click()
    ' Use Windows API to open the webbrowser on page specified below
    ShellExecute Me.hwnd, "open", "www.whatever.com", "", "", 10
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Change the colour of the text on the label when mouse is taken off
    Label1.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
    ' When timer reaches end, unload the form
    Unload Me
    ' Ask the user if he wants to view the inroduction
    r = MsgBox("Do you want to Play the introduction?", vbYesNo + vbQuestion, "User Input")
    ' Determine the users answer
    If r = vbYes Then
        Call Load(frmSplash1)
        Call frmSplash1.Show
    End If
    If r = vbNo Then
        Call Load(frmLogin)
        Call frmLogin.Show
    End If
    
    
End Sub

Private Sub Timer2_Timer()
    ' Change the visible label on the form
    Label2.Visible = True
    Label1.Visible = False
End Sub
