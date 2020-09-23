VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmSend 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Breathe SMS Sender by Five-Times"
   ClientHeight    =   4455
   ClientLeft      =   3105
   ClientTop       =   3690
   ClientWidth     =   5160
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4815
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   8493
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
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      Picture         =   "frmSend.frx":0442
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdManage 
      Caption         =   "Number Manager"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Picture         =   "frmSend.frx":2864
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAddress 
      Caption         =   "Load Number Book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Picture         =   "frmSend.frx":2CA6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox lstAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Picture         =   "frmSend.frx":30E8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&Contact"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      Picture         =   "frmSend.frx":352A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSend.frx":396C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   2  'Center
      Height          =   1365
      Left            =   120
      MaxLength       =   145
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "Enter Body Text Here"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtPhoneNumber 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   960
      Y2              =   4800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Address Book"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Message Body"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter the Destination Number Here"
      ForeColor       =   &H80000006&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Menu mnuOps 
      Caption         =   "Options"
      Begin VB.Menu mnuSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout2 
      Caption         =   "About"
      Begin VB.Menu mnuURL 
         Caption         =   "Website"
      End
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
    Call Load(frmContact)
    Call frmContact.Show
End Sub

Private Sub cmdAddress_Click()
    Call LoadAddressBook
End Sub

Private Sub cmdHelp_Click()
    ' Load the help form
    Call Load(frmHelp)
    Call frmHelp.Show
End Sub

Private Sub cmdManage_Click()
    ' Load the address manager
    Call Load(frmAddress)
    Call frmAddress.Show
End Sub

Private Sub cmdSend_Click()
    ' Disable the command Button
    cmdSend.Enabled = False
    
    Dim Number As String
    Dim Message As String * 146
    ' Check the user has entered values in the text boxes
    If txtMessage.Text = "" Then
        Call Error
        cmdSend.Enabled = True
        Exit Sub
    End If
    If txtPhoneNumber.Text = "" Then
        Call Error
        cmdSend.Enabled = True
        Exit Sub
    End If
    ' Transfer the inputed number to a varable
    Number = txtPhoneNumber.Text
    ' Because when sending SMS's through breathes server you have to replace all spaces with +'s
    ' We need to use the code below
    ' The syntax is Varable =Replace(String, Search Key, Replace Character)
    Message = Replace(txtMessage.Text, " ", "+")
    ' Open the web browser control to send the message
    Web.Navigate "http://www.breathe.com/services/textmessaging.html?number=" & Number & "&message=" & Message & "&charleft=139%2F146&submit.x=16&submit.y=9"
    ' Call API function sleep to suspend execution so the program can send the data
    Call Sleep(5000)
    ' Enable the command button to allow the user to send another message
    cmdSend.Enabled = True
    
End Sub

Private Sub Form_Load()
     ' If the file exists, do nothing
     If Dir("c:\addressbook.txt") <> "" Then
        DoEvents
    Else
        ' if it doesnt exist, create the address book file
        Dim filepath3 As String
        filepath3 = "c:\addressbook.txt"
        Open filepath3 For Output As #3
        ' close the file
        Close #3
    End If
End Sub

Private Sub lstAddress_Click()
    Dim blank As Integer
    Dim transfer As String
    Dim final As String
    ' Assign the value clicked in the list box to a varable so we can minipulate it
    transfer = (lstAddress.List(lstAddress.ListIndex))
    ' Use InStr to find the first space in the string, and assign its location to blank
    blank = InStr(1, transfer, " ")
    ' Assign the string from transfer to the varable final, triming it by the value produced in blank and -1 to delete the space
    final = Left$(transfer, blank - 1)
    ' Transfer the formated phone number to the text box
    txtPhoneNumber.Text = final
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSend_Click()
    Call cmdSend_Click
End Sub

Private Sub mnuURL_Click()
    ' Use Windows API to open the webbrowser on page specified below
    ShellExecute Me.hwnd, "open", "www.whatever.com", "", "", 10
End Sub
