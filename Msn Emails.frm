VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form msnemails 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Msn Emails"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "Msn Emails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog saveit 
      Left            =   2520
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Msn Emails.frx":0442
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   0
      Picture         =   "Msn Emails.frx":0448
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5040
   End
End
Attribute VB_Name = "msnemails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo handleit
saveit.CancelError = False
saveit.Filter = "Text|*.txt"
saveit.DialogTitle = "Save Msn Emails LIst"
saveit.ShowSave

Open saveit.FileName For Output As #1
Print #1, Text1.Text
Close savelist


handleit:
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1 = "Lista De Emails" & vbCrLf

For online = 0 To xtrememsn.online.ListCount - 1
xtrememsn.online.ListIndex = online
Text1.Text = Text1.Text & vbCrLf & xtrememsn.online.Text
Next online

For offline = 0 To xtrememsn.offline.ListCount - 1
xtrememsn.offline.ListIndex = offline
Text1.Text = Text1.Text & vbCrLf & xtrememsn.offline.Text
Next offline
End Sub
