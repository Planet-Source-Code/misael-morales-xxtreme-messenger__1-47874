VERSION 5.00
Begin VB.Form xtrememsn 
   BackColor       =   &H8000000D&
   Caption         =   "Extreme Messenger"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5775
   Icon            =   "Xtreme Messenger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   4920
      Top             =   120
   End
   Begin VB.ListBox offline 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.ListBox online 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "Xtreme Messenger.frx":0442
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Offline Buddys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connected Buddys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      Picture         =   "Xtreme Messenger.frx":183C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5835
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu mnusignin 
         Caption         =   "Sing in"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu others 
         Caption         =   "Others"
         Begin VB.Menu mnumultiplemsg 
            Caption         =   "Multiple Mesage"
         End
         Begin VB.Menu mnumassivemessage 
            Caption         =   "Massive Message"
         End
      End
      Begin VB.Menu mnuemails 
         Caption         =   "Msn Emails"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "xtrememsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MSNAPI  As New MessengerAPI.Messenger
Dim MSN As New MsgrObject
Dim widthmsn As String
Dim heigthmsn  As String

Private Sub Form_Load()
heigthmsn = xtrememsn.Height
widthmsn = xtrememsn.Width
End Sub

Private Sub Form_Resize()
On Error Resume Next

xtrememsn.Height = heigthmsn
xtrememsn.Width = widthmsn
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuabout_Click()
Load credits
credits.Show 1
End Sub


Private Sub mnuclose_Click()
End
End Sub

Private Sub mnuemails_Click()
Load msnemails
msnemails.Show
End Sub

Private Sub mnulogoff_Click()
On Error Resume Next
MSN.Logoff
End Sub

Private Sub mnumassivemessage_Click()
Load massive
massive.Show
End Sub

Private Sub mnumultiplemsg_Click()
Load multiplemessages
multiplemessages.Show
End Sub

Private Sub mnuoptions_Click()
Load options
options.Show 1
End Sub

Private Sub mnusignin_Click()
Load signin
signin.Show 1
End Sub
Private Sub offline_DblClick()
MsgBox offline.Text
End Sub

Private Sub online_Click()
contact.UserName = online.Text
End Sub

Private Sub online_DblClick()
Load message
message.Show
End Sub


Private Sub online_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton And online.ListIndex >= 0 Then PopupMenu others
End Sub

Private Sub Timer1_Timer()
Call RefreshList1

End Sub
Private Sub MSN_OnUserStateChanged(ByVal pUser As Messenger.IMsgrUser, ByVal mPrevState As _
            Messenger.MSTATE, pfEnableDefault As Boolean)
                        End Sub


