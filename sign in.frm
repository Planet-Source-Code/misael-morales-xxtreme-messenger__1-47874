VERSION 5.00
Begin VB.Form signin 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign In......"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "sign in.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Sign In"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "thedragon"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Email 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "johnsurfer21@hotmail.com"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   0
      Picture         =   "sign in.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   15
   End
End
Attribute VB_Name = "signin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub Command1_Click()
On Error Resume Next
MSN.Logon Email.Text, Password.Text, MSN.Services.PrimaryService
Call RefreshList1
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
