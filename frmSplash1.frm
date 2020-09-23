VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmSplash1 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   36000
      Left            =   2400
      Top             =   4440
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash2 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   -480
      Width           =   7095
      _cx             =   22884579
      _cy             =   22878229
      Movie           =   "http://www.breathe.com/breathepromo/flash/breathe3.swf"
      Src             =   "http://www.breathe.com/breathepromo/flash/breathe3.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ExactFit"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    ' Play the movie when the form loads
    flash2.Play
End Sub

Private Sub Timer3_Timer()
    ' Stop the movie when the timer expires
    flash2.Stop
    ' Suspend execution for 4 seconds
    Call Sleep(4000)
    ' Unload the form and show the login form
    Unload Me
    Call Load(frmLogin)
    Call frmLogin.Show
End Sub
