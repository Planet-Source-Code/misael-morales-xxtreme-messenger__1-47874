VERSION 5.00
Begin VB.Form multiplemessages 
   BackColor       =   &H8000000D&
   Caption         =   "Multiple Messages"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   Icon            =   "Multiple Messages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Delete"
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Send"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Multiple Messages.frx":0442
      Top             =   2640
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesage:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "multiplemessages"
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
            
                                              
            For send = 0 To List1.ListCount - 1
            DoEvents
            
            List1.ListIndex = send
            Set user = MSN.CreateUser(List1.Text, MSN.Services.PrimaryService)
            
            bstrMsgText = Text1.Text
            DoEvents
            
            user.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
            Label3.Caption = counter
            DoEvents
            Next send
            
            End If
            End If
End Sub

Private Sub Command2_Click()
List1.RemoveItem List1.ListIndex
List1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""

For Add = 0 To xtrememsn.online.ListCount - 1
xtrememsn.online.ListIndex = Add
List1.AddItem xtrememsn.online.Text
Next Add

End Sub
