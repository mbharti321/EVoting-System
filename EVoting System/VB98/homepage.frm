VERSION 5.00
Begin VB.Form frmhome 
   BackColor       =   &H000080FF&
   Caption         =   "HOMEPAGE"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "homepage.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   16320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   1215
      Left            =   16200
      TabIndex        =   16
      Top             =   7560
      Width           =   2055
      Begin VB.CommandButton cmdexit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         MaskColor       =   &H80000007&
         Picture         =   "homepage.frx":C035F
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7095
      Left            =   6840
      TabIndex        =   7
      Top             =   2280
      Width           =   6135
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   360
         TabIndex        =   9
         Top             =   4080
         Width           =   5295
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   14
            Top             =   1920
            Width           =   3615
            Begin VB.CommandButton cmdclear 
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   6
               Top             =   120
               Width           =   1095
            End
            Begin VB.CommandButton cmdlogin 
               Caption         =   "Login"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   360
               TabIndex        =   5
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   4935
            Begin VB.TextBox txtpassword 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   2040
               PasswordChar    =   "*"
               TabIndex        =   4
               ToolTipText     =   "Enter Password"
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label Label5 
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txtuserid 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               TabIndex        =   3
               ToolTipText     =   "Enter User_Id"
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label Label4 
               Caption         =   "User_Id:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1575
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "User_Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   5535
         Begin VB.OptionButton optnominee 
            BackColor       =   &H008080FF&
            Caption         =   "Nominee_Login"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optvoter 
            BackColor       =   &H008080FF&
            Caption         =   "Voter_Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optadmin 
            BackColor       =   &H008080FF&
            Caption         =   "Admin_Login"
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
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Image Image5 
         Height          =   2775
         Left            =   3120
         Picture         =   "homepage.frx":E58DF
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image4 
         Height          =   2775
         Left            =   240
         Picture         =   "homepage.frx":E9FE8
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Image Image7 
      Height          =   1815
      Left            =   2760
      Picture         =   "homepage.frx":113272
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Image Image6 
      Height          =   2535
      Left            =   2760
      Picture         =   "homepage.frx":12225D
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   2115
      Left            =   13320
      Picture         =   "homepage.frx":13E38A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   5325
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   2760
      Picture         =   "homepage.frx":14432A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   4665
      Left            =   13320
      Picture         =   "homepage.frx":146F88
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to E-Voting System "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   10335
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

'clear Buttton
Private Sub cmdclear_Click()
txtuserid.Text = ""
txtpassword.Text = ""
txtuserid.SetFocus
End Sub
'Exit Button
Private Sub cmdexit_Click()
'Confirmation
Dim wish As Integer
wish = MsgBox("Do you really want to Exit ?", vbQuestion + vbYesNo)
If wish <> vbYes Then
  Exit Sub
End If

End
End Sub
'Login button
Private Sub cmdlogin_Click()
Dim userid As String
'Checking userid is blank?
If txtuserid.Text = "" Then
   MsgBox "Enter userid", vbInformation + vbOKOnly, "LOgin"
   txtuserid.SetFocus
   Exit Sub
End If
'chenking password is blank?
If txtpassword.Text = "" Then
   MsgBox "Enter Password", vbInformation + vbOKOnly, "LOgin"
   txtpassword.SetFocus
   Exit Sub
End If

'if both the field is not blank then
If txtuserid.Text <> "" And txtpassword.Text <> "" Then
 'For admin login
  If optadmin.Value = True Then
    userid = txtuserid.Text
    rs.Open "Select * from admin where adminid='" & userid & "' and password='" & txtpassword.Text & "'", conn, adOpenDynamic, adLockOptimistic, adCmdText
    If rs.EOF = True Then
     MsgBox "Invalid Admin_id or password!!!!!!! Please enter a valid Admin_id and password", vbCritical + vbOKOnly, "login"
     txtuserid.Text = ""
     txtpassword.Text = ""
     txtuserid.SetFocus
    Else
     MsgBox "Adminid and password are correct", vbInformation, "Login"
     MDIadmin.Show
     Unload Me
    End If
   rs.Close
  End If
 
   
  ' for voter login
  If optvoter.Value = True Then
   userid = UCase(txtuserid.Text)
   rs.Open "Select * from voterlist where voterid='" & userid & "' and password='" & txtpassword.Text & "' ", conn, adOpenDynamic, adLockOptimistic, adCmdText
   If rs.EOF = True Then
     MsgBox "Invalid voter_id or password!!!!!!! Please enter a valid User_id and password", vbCritical + vbOKOnly, "login"
     txtuserid.Text = ""
     txtpassword.Text = ""
     txtuserid.SetFocus
   Else
    MsgBox "Userid and password are correct", vbInformation, "Login"
   'To save the details of user going to login
     Module1.usertype = "voter"
     Module1.logged_userid = userid
     Module1.logged_username = rs(2)
     frmuser.Show
     Unload Me
   End If
   rs.Close
  End If
  
  'Nominee login
  If optnominee = True Then
      userid = UCase(txtuserid.Text)
       rs.Open "Select * from nomineelist where nomineeid='" & userid & "' and password='" & txtpassword.Text & "' ", conn, adOpenDynamic, adLockOptimistic, adCmdText
       If rs.EOF Then
          MsgBox "Invalid User_id or password!!!!!!! Please enter a valid User_id and password", vbCritical + vbOKOnly, "login"
          txtuserid.Text = ""
          txtpassword.Text = ""
          txtuserid.SetFocus
       Else
          MsgBox "Userid and password are correct", vbInformation, "Login"
        'To save the details of user going to login
         Module1.usertype = "nominee"
         Module1.logged_userid = userid
         Module1.logged_username = rs(2)
         frmuser.Show
         Unload Me
       End If
         rs.Close
 End If
   
End If
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"

Width = 21000
Top = 0
Left = 0
Height = 11000
End Sub

