VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form options 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdchgstatus 
      BackColor       =   &H8000000D&
      Caption         =   "Change"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdchangenick 
      BackColor       =   &H8000000D&
      Caption         =   "Change"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtnewnick 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   5953
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Nickname:"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject
Dim state As Integer

Private Sub cmdChangeNick_Click()
If MSN.LocalState = MSTATE_OFFLINE Then
MsgBox "You are not Signed In"
Else
MSN.Services.PrimaryService.FriendlyName = txtnewnick.Text
txtnewnick.Text = MSN.LocalFriendlyName
End If
End Sub

Private Sub cmdchgstatus_Click()

Call stateschoose(state + 1)
End Sub



Private Sub Form_Load()
On Error Resume Next

txtnewnick.Text = MSN.LocalFriendlyName



With List1
.AddItem "ONLINE"
.AddItem "BUSY"
.AddItem "BE RIGHT BACK"
.AddItem "AWAY"
.AddItem "ON THE PHONE"
.AddItem "OUT TO LUNCH"
.AddItem "INVISIBLE"


End With
With Tree.Nodes
.Add , , "Options", "Options "
.Add "Options", tvwChild, "NickName", "Nickname"
.Add "Options", tvwChild, "Status", "Status"


End With
' Put in invisible mode all items
cmdchgstatus.Visible = False
List1.Visible = False
txtnewnick.Visible = False
cmdchangenick.Visible = False
Label1.Visible = False
End Sub
Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Key = "NickName" Then
cmdchgstatus.Visible = False
List1.Visible = False
txtnewnick.Visible = True
cmdchangenick.Visible = True
Label1.Visible = True
End If

If Node.Key = "Status" Then
cmdchgstatus.Visible = True
List1.Visible = True
txtnewnick.Visible = False
cmdchangenick.Visible = False
Label1.Visible = False
End If

If Node.Key = "Options" Then
cmdchgstatus.Visible = False
List1.Visible = False
txtnewnick.Visible = False
cmdchangenick.Visible = False
Label1.Visible = False
End If

End Sub
Private Sub List1_Click()
On Error Resume Next
state = List1.ListIndex
End Sub


Private Sub stateschoose(statenumber As Integer)

Select Case statenumber

Case 1
MSN.LocalState = MSTATE_ONLINE
Case 2
MSN.LocalState = MSTATE_BUSY
Case 3
MSN.LocalState = MSTATE_BE_RIGHT_BACK
Case 4
MSN.LocalState = MSTATE_AWAY
Case 5
MSN.LocalState = MSTATE_ON_THE_PHONE
Case 6
MSN.LocalState = MSTATE_OUT_TO_LUNCH
Case 7
MSN.LocalState = MSTATE_INVISIBLE

End Select


End Sub

