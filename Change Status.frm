VERSION 5.00
Begin VB.Form chgstatus 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1815
   Icon            =   "Change Status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   1815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H8000000D&
      Caption         =   "Close"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdchgstatus 
      BackColor       =   &H8000000D&
      Caption         =   "Change"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "chgstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim state As Integer
Private MSN As New MsgrObject

Private Sub cmdchgstatus_Click()
Call stateschoose(state + 1)
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
With List1
.AddItem "ONLINE"
.AddItem "BUSY"
.AddItem "BE RIGHT BACK"
.AddItem "AWAY"
.AddItem "ON THE PHONE"
.AddItem "OUT TO LUNCH"
.AddItem "INVISIBLE"


End With

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
