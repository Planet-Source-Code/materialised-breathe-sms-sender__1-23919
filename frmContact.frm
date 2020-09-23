VERSION 5.00
Begin VB.Form frmContact 
   Caption         =   "Contacting Me"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmContact.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      Caption         =   $"frmContact.frx":0442
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Five-Times"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4935
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub
