VERSION 5.00
Begin VB.MDIForm MDIadmin 
   BackColor       =   &H8000000C&
   Caption         =   "Admin"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   4365
   ClientWidth     =   19065
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIadmin.frx":0000
   Begin VB.Menu menuvoter 
      Caption         =   "&Voter_list"
   End
   Begin VB.Menu menunominee 
      Caption         =   "&Nominee"
   End
   Begin VB.Menu mnuelection 
      Caption         =   "&Election"
   End
   Begin VB.Menu mnuresult 
      Caption         =   "&Result"
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
   End
   Begin VB.Menu mnulogout 
      Caption         =   "&Logout"
   End
End
Attribute VB_Name = "MDIadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
Width = 21000
Top = 0
Left = 0
Height = 11000
End Sub

Private Sub menunominee_Click()
Module1.nomineelist = "admin"
frmnominee.Show

Unload Me
End Sub

Private Sub menuvoter_Click()
Module1.voterlist = "admin"
frmvoterlist.Show
Unload Me
End Sub

Private Sub mnuelection_Click()

frmelection.Show
Unload Me
End Sub

Private Sub mnulogout_Click()
'Confirmation
Dim wish As Integer
wish = MsgBox("Do you really want to Logout ?", vbQuestion + vbYesNo)
If wish <> vbYes Then
  Exit Sub
End If

frmhome.Show
Unload Me
End Sub

Private Sub mnureport_Click()
frmreport.Show
Unload Me
End Sub

Private Sub mnuresult_Click()
Module1.result = "admin"
Unload Me
frmresult.Show
End Sub
