VERSION 5.00
Begin VB.Form massive 
   BackColor       =   &H8000000D&
   Caption         =   "Massive to -"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   Icon            =   "Massive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "10"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Times Sended:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Times To Send:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "massive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject
Dim user As IMsgrUser
Dim counter As Integer
      
Private Sub Command1_Click()
On Error Resume Next
            Dim bstrMsgHeader As String
            Dim bstrMsgText As String
            
            If MSN.LocalState = MSTATE_OFFLINE Then
            MsgBox "You are not Signed In"
            Else
            If MSN.LocalState = MSTATE_INVISIBLE Then
            MsgBox "Change you status first !"
            Else
            
            Set user = MSN.CreateUser(xtrememsn.online.Text, MSN.Services.PrimaryService)
            bstrMsgText = Text1.Text
                                   
            Do Until counter = Val(Text2.Text)
            DoEvents
            counter = counter + 1
            DoEvents
            user.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
            Label3.Caption = counter
            DoEvents
            Loop
            counter = 0
            End If
            End If

End Sub

Private Sub Form_Load()
Text1.Text = ""
Me.Caption = "Massive To - " & xtrememsn.online.Text
End Sub
