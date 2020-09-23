Attribute VB_Name = "Module1"
Type user
UserName As String * 20
UserEmail As String * 50
UserState As String

End Type

Global contact As user
Private MSN As New MsgrObject
Public Sub RefreshList1()
Dim user As IMsgrUser

xtrememsn.online.Clear
xtrememsn.offline.Clear
For Each user In MSN.List(MLIST_CONTACT)
If user.state = MSTATE_OFFLINE Then
xtrememsn.offline.AddItem (user.EmailAddress)
Else
xtrememsn.online.AddItem (user.EmailAddress)
End If
Next

End Sub
