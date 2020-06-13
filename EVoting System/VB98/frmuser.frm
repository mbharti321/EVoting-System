VERSION 5.00
Begin VB.Form frmuser 
   BackColor       =   &H00FFFF80&
   Caption         =   "Voter/User"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   Picture         =   "frmuser.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   15630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6855
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   12855
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2895
         Begin VB.CommandButton cmddetails 
            Caption         =   "Your_Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   2295
         End
         Begin VB.CommandButton cmdnominee 
            Caption         =   "Nominee_Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmdresult 
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   8
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton cmdlogout 
            Caption         =   "Logout"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   2280
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdvote 
         BackColor       =   &H00C0C0C0&
         Caption         =   "VOTE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8880
         Picture         =   "frmuser.frx":C035F
         TabIndex        =   5
         Top             =   5280
         Width           =   3495
      End
      Begin VB.Image Image6 
         Height          =   2655
         Left            =   480
         Picture         =   "frmuser.frx":C24FB
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Image Image4 
         Height          =   1815
         Left            =   3600
         Picture         =   "frmuser.frx":C7514
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   4935
      End
      Begin VB.Image Image3 
         Height          =   1695
         Left            =   3600
         Picture         =   "frmuser.frx":CD99E
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Image Image2 
         Height          =   2175
         Left            =   3600
         Picture         =   "frmuser.frx":D3D80
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3615
      End
      Begin VB.Image Image5 
         Height          =   2175
         Left            =   7560
         Picture         =   "frmuser.frx":D4EE3
         Stretch         =   -1  'True
         Top             =   600
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   8880
         Picture         =   "frmuser.frx":D9BE2
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.Label lblusername 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "user_name"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome "
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim img1 As String
'Back button
Private Sub cmdback_Click()
'Confirmation
Dim wish As Integer
wish = MsgBox("Do you really want to Exit ?", vbQuestion + vbYesNo)
If wish <> vbYes Then
  Exit Sub
End If

frmhome.Show
Unload Me
End Sub
'Your_details button
Private Sub cmddetails_Click()
 'if the user is nominee
 If usertype = "nominee" Then
    'To change the value of nomineelist variable's value presented in Module1
    Module1.nomineelist = "user"
    
    frmnominee.Show
    frmnominee.Label1.Caption = "User_Details"
    
    'To enable password changing frame
     frmnominee.Frame4.Visible = True
     
    'for hiding users details other than loggein user's details
    frmnominee.Frame2.Visible = False
    frmnominee.Frame3.Visible = False
    frmnominee.DataGrid1.Visible = False
    frmnominee.cmdimage.Visible = False
    'for loading user's data in different fields
    rs.Open "select * from nomineelist where nomineeid='" & logged_userid & "'", conn, adOpenDynamic, adLockOptimistic
    frmnominee.txtnomineeid.Text = rs(0)
    frmnominee.txtpassword.Text = rs(1)
    frmnominee.txtname.Text = rs(2)
    frmnominee.txtfather.Text = rs(3)
    frmnominee.txtqualification.Text = rs(4)
    frmnominee.cmbgender.Text = rs(5)
    frmnominee.DTPicker1.Value = rs(6)
    frmnominee.imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    frmnominee.txtaddress.Text = rs(8)
    MsgBox "You are a Nominee!!! You can not alter details other than 'PASSWORD' !!!!", vbInformation + vbOKOnly, "User_Details"
    Unload Me
 'If the user is not nominee, means he/she is a voter
 ElseIf usertype = "voter" Then
    'To change the value of voterlist variable's value presented in Module1
    Module1.voterlist = "user"
    frmvoterlist.Show
    frmvoterlist.Label1.Caption = "User_Details"
    
    'To enable password changing frame
     frmvoterlist.Frame4.Visible = True
    
    'for hiding users details other than loggedin user's details
    frmvoterlist.DataGrid1.Visible = False
    frmvoterlist.Frame2.Visible = False
    frmvoterlist.Frame3.Visible = False
    frmvoterlist.cmdimage.Visible = False
    'for loading user's data in different fields
    rs.Open "select * from voterlist where voterid='" & logged_userid & "'", conn, adOpenDynamic, adLockOptimistic
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    img1 = rs(6)
    frmvoterlist.txtaddress.Text = rs(7)
    frmvoterlist.imgvoter.Picture = LoadPicture(img1)
    MsgBox "!!!You can not edit your details other than 'PASSWORD' !!!!", vbInformation + vbOKOnly, "User_Details"
    Unload Me
 End If
End Sub

Private Sub cmdlogout_Click()
'Confirmation
Dim wish As Integer
wish = MsgBox("Do you really want to Exit ?", vbQuestion + vbYesNo)
If wish <> vbYes Then
  Exit Sub
End If

frmhome.Show
Unload Me
End Sub

Private Sub cmdnominee_Click()
'to change the value of nomineelist variable's value presented in Module1
Module1.nomineelist = "user"
frmnominee.Show
frmnominee.Label1.Caption = "Nominee_Details"
'for hiding users details other than loggein user's details
frmnominee.Frame2.Visible = False
frmnominee.Frame3.Left = 3000
frmnominee.DataGrid1.Visible = False
frmnominee.lbluserid.Visible = False
frmnominee.txtnomineeid.Visible = False
frmnominee.lblpassword.Visible = False
frmnominee.txtpassword.Visible = False
frmnominee.cmdimage.Visible = False

MsgBox "!!!You can not alter details!!!!", vbInformation + vbOKOnly, "User_Details"
Unload Me
End Sub

Private Sub cmdresult_Click()
rs1.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
If rs1.EOF Then
   MsgBox "Sorry,But the Result_Page is not visible now. Reason may be either 'vote hasn't started yet' or 'Voting is still going on' or 'Result hasn't decleared yet'", vbInformation, "Result"
   rs1.Close
   Exit Sub
End If
If rs1(5) = 1 Then
    Module1.result = "user"
    frmresult.Show
    frmresult.Frame4.Visible = False
    Unload Me
Else
    MsgBox "Sorry,But the Result_Page is not visible now. Reason may be either 'vote hasn't started yet' or 'Voting is still going on' or 'Result hasn't decleared yet'", vbInformation, "Result"
End If
rs1.Close
End Sub

Private Sub cmdvote_Click()
frmvoting.Show

'if no nominee is there
rs2.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
If rs2.EOF Then
    MsgBox "Sorry,but you can't vot now.. Reason:-No any nominee is listed yet.."
    frmvoting.cmdconfirm.Enabled = False
    Exit Sub
    rs2.Close
Else
    frmvoting.cmdconfirm.Enabled = True
    rs2.Close
End If

rs2.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
If rs2.EOF Then
   MsgBox "Sorry,But you can't cast your vote now. please try after some time.. Reason may be either ''vote hasn't started yet'' or '' Voting time has been collapsed''", vbInformation, "Result"
   rs2.Close
   frmvoting.cmdconfirm.Enabled = False
   Exit Sub
End If

If rs2(4) <> 1 Then
    MsgBox "Sorry,But you can't cast your vote now. please try after some time.. Reason may be either ''vote hasn't started yet'' or '' Voting time has been collapsed''", vbInformation, "Result"
    frmvoting.cmdconfirm.Enabled = False
Else

    rs1.Open "select * from vote_user where userid='" & Module1.logged_userid & "'", conn, adOpenDynamic, adLockOptimistic
    If rs1.EOF <> True Then 'Means if user has already voted
        Dim nominee As String
        nominee = rs1(1) 'Nominee_id
        rs1.Close
     
       'To load data of nominee whom user has voted
        rs1.Open "select * from nomineelist where nomineeid='" & nominee & "'", conn, adOpenDynamic, adLockOptimistic
        frmvoting.lblnomineename.Caption = rs1(2)
        frmvoting.lblqualification.Caption = rs1(4)
        frmvoting.imgnominee.Picture = LoadPicture(rs1(7))
       'combobox with nominee details
        frmvoting.Combo1.Clear
        frmvoting.Combo1.AddItem rs1(0)
        frmvoting.Combo1.Text = rs1(0)
        frmvoting.Combo1.Enabled = False
     
        frmvoting.cmdconfirm.Enabled = False
        'To change the caption of lbldisplay
         frmvoting.lbldisplay.Caption = "You have voted the below candidate:"
     
         MsgBox "you have already voted.Please wait for 'RESULT'", vbInformation + vbOKOnly, "Voting"
         rs1.Close
    Else
     MsgBox "!!!!!!!!_ _ _It's voting time, Please vote wisely!_ _ _!!!!!!!!", vbInformation + vbOKOnly, "Voting"
     rs1.Close
    End If
rs2.Close
End If

Unload Me
End Sub

Private Sub Form_Load()
'To load username in username_display
 lblusername = Module1.logged_username
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"

Width = 21000
Top = 0
Left = 0
Height = 11000

End Sub

