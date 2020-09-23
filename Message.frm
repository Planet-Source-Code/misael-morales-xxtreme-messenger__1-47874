VERSION 5.00
Begin VB.Form message 
   BackColor       =   &H8000000D&
   Caption         =   "Message"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000D&
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Message.frx":0442
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   4740
      Left            =   0
      Picture         =   "Message.frx":0448
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject
Dim user As IMsgrUser
            
            Private Sub cmdSendIM_Click()
            
            End Sub

Private Sub Command1_Click()

            Dim bstrMsgHeader As String
            Dim bstrMsgText As String
            
            If MSN.LocalState = MSTATE_OFFLINE Then
            MsgBox "You are not Signed In"
            Else
            If MSN.LocalState = MSTATE_INVISIBLE Then
            MsgBox "Change you status first !"
            Else
            
            Set user = MSN.CreateUser(xtrememsn.online.Text, MSN.Services.PrimaryService)
            bstrMsgText = Text2.Text
            Text1.Text = Text1.Text & vbCrLf & Text2.Text
            Text2.Text = ""
            user.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
                       
            End If
            End If
End Sub
Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = ""
Me.Caption = xtrememsn.online.Text
End Sub



