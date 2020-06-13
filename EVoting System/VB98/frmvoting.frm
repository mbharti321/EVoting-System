VERSION 5.00
Begin VB.Form frmvoting 
   BackColor       =   &H8000000B&
   Caption         =   "Voting"
   ClientHeight    =   9210
   ClientLeft      =   4800
   ClientTop       =   3165
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   Picture         =   "frmvoting.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   7215
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   13455
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Nominee_Details:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   4320
         TabIndex        =   15
         Top             =   3000
         Width           =   5295
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            ItemData        =   "frmvoting.frx":C035F
            Left            =   3360
            List            =   "frmvoting.frx":C036C
            Style           =   1  'Simple Combo
            TabIndex        =   16
            Text            =   "Nominees"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lbldisplay 
            Alignment       =   2  'Center
            Caption         =   "Please Select your Nominee whom you want to Vote:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Nominee_Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblnomineename 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Image:"
            Height          =   255
            Left            =   1200
            TabIndex        =   19
            Top             =   1560
            Width           =   975
         End
         Begin VB.Image imgnominee 
            Height          =   2055
            Left            =   240
            Picture         =   "frmvoting.frx":C0388
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Qualification:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblqualification 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000080FF&
         Caption         =   "Election"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   8535
         Begin VB.TextBox txtdescription 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label Label5 
            Caption         =   "Ending_Time:"
            Height          =   255
            Left            =   6360
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Descripition"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label label2 
            Caption         =   "Election_name"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblelectionname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblendingtime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5760
            TabIndex        =   10
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label11 
            Caption         =   "Starting_Time:"
            Height          =   255
            Left            =   6240
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblstartingtime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5760
            TabIndex        =   8
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000E&
         Height          =   1455
         Left            =   10200
         TabIndex        =   3
         Top             =   5400
         Width           =   2775
         Begin VB.CommandButton cmdconfirm 
            Caption         =   "Confirm_Vote"
            BeginProperty Font 
               Name            =   "Lucida Calligraphy"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            MaskColor       =   &H0000FF00&
            TabIndex        =   5
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmdexit 
            BackColor       =   &H8000000B&
            Caption         =   "EXIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   10200
         Picture         =   "frmvoting.frx":C184A
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   2160
         Left            =   10200
         Picture         =   "frmvoting.frx":CE58C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2700
      End
      Begin VB.Image Image4 
         Height          =   3975
         Left            =   1080
         Picture         =   "frmvoting.frx":D084B
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "VOTING  TIME"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "frmvoting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim votecount As String
Private Sub cmdback_Click()
Unload Me
frmuser.Show
End Sub

Private Sub cmdconfirm_Click()

'To check nominee is selected or not
If Combo1.Text = "Nominees!!!" Then
 MsgBox "Please select any nominee for voting confirmation", vbCritical, "Voting"
 Exit Sub
End If

'Confirmation
Dim wish As Integer
wish = MsgBox("Do you want to confirm your vote? After confirmation you won't be able to change your vote..", vbQuestion + vbYesNo, "Voting_Confirmation")
If wish <> vbYes Then
  Exit Sub
End If


'To check whether voter has already voted
rs1.Open "select * from vote_user where userid='" & Module1.logged_userid & "'", conn, adOpenDynamic, adLockOptimistic
   If rs1.EOF Then
    rs1.Close
    GoTo updatedata
   Else
     MsgBox "you have already voted"
     rs1.Close
     Exit Sub
   End If
    
    
updatedata:

'To increase the count of votes in table of selected nominee
rs1.Open "select * from nomineelist where nomineeid='" & Combo1.Text & "'", conn, adOpenDynamic, adLockOptimistic
votecount = rs1(9)
votecount = votecount + 1
rs1(9) = votecount
rs1.Update
rs1.Close
'To add data in vote_user table
rs1.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1(0) = Module1.logged_userid 'userid
rs1(1) = Combo1.Text           'nomineeid
'To store nominee id for displaying after vote conformation
    Dim nominee_id As String
    nominee_id = Combo1.Text
rs1.Update
rs1.Requery
rs1.Close
MsgBox "!!!!!_ _Thanks for voting_ _!!!!!!", vbInformation + vbOKOnly, "Voting"

'To load data of nominee whom user has voted
     rs1.Open "select * from nomineelist where nomineeid='" & nominee_id & "'", conn, adOpenDynamic, adLockOptimistic
     frmvoting.lblnomineename.Caption = rs1(2)
     frmvoting.lblqualification.Caption = rs1(4)
     frmvoting.imgnominee.Picture = LoadPicture(rs1(7))
'combobox with nominee details
     frmvoting.Combo1.Clear
     frmvoting.Combo1.AddItem rs1(0)
     frmvoting.Combo1.Text = rs1(0)
 'To disable controls
     frmvoting.Combo1.Enabled = False
     frmvoting.cmdconfirm.Enabled = False
     frmvoting.Frame1.Enabled = False
 'To change the caption of lbldisplay
     frmvoting.lbldisplay.Caption = "You have voted the below candidate:"
  rs1.Close
End Sub

Private Sub cmdexit_Click()
Unload Me
frmuser.Show
End Sub
'To load nominees' data when  comboitemes are dblclicked
Private Sub Combo1_DblClick()
rs1.Open "select * from nomineelist where nomineeid='" & Combo1.Text & "'", conn, adOpenDynamic, adLockOptimistic
lblnomineename.Caption = rs1(2)
lblqualification.Caption = rs1(4)
imgnominee.Picture = LoadPicture(rs1(7))
rs1.Close

End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"
'to Load data in fields
 rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
 If rs.EOF Then
   frmvoting.Show
   MsgBox "Election details are not available!!!"
   cmdconfirm.Enabled = False
   'combo1.Enabled=False
 Else
    lblelectionname.Caption = rs(0)
    txtdescription.Text = rs(1)
    lblstartingtime.Caption = rs(2)
    lblendingtime.Caption = rs(3)
    rs.Close
End If

'To load nominees' nmae in combobox
    rs1.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
    Combo1.Clear
    Combo1.Text = "Nominees!!!"
    Do While (rs1.EOF <> True)
     Combo1.AddItem UCase(rs1(0))
     rs1.MoveNext
    Loop
   rs1.Close

Width = 21000
Top = 0
Left = 0
Height = 11000
End Sub

