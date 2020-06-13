VERSION 5.00
Begin VB.Form frmelection 
   BackColor       =   &H80000012&
   Caption         =   "Election"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   Picture         =   "frmelection.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Election_Details:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   4800
      TabIndex        =   4
      Top             =   2760
      Width           =   9375
      Begin VB.TextBox txtendingtime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtstartingtime 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtdescription 
         Height          =   1095
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   6975
      End
      Begin VB.TextBox txtelectionname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   6
         Top             =   360
         Width           =   6975
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   120
         Picture         =   "frmelection.frx":35DB5
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   9135
      End
      Begin VB.Label Label5 
         Caption         =   "Ending_Time:"
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
         Left            =   4920
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Starting_Time:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Descripition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label label2 
         Caption         =   "Election_name"
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   14520
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
      Begin VB.Frame Frame3 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1935
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
         Begin VB.CommandButton cmdenable 
            Caption         =   "Enable_voting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1815
         End
         Begin VB.CommandButton cmddisable 
            Caption         =   "Disable_Voting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit_Details"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "EXIT"
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
         TabIndex        =   3
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmdsubmit 
         BackColor       =   &H0000C0C0&
         Caption         =   "SUBMIT"
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
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "ELECTION"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
   End
End
Attribute VB_Name = "frmelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim conn As New ADODB.Connection
Private Sub cmdback_Click()
MDIadmin.Show
Unload Me
End Sub

Private Sub cmddisable_Click()
wish = MsgBox("Are you sure to Disable voting(Y/N)? Before disabling voting, please check voting time.As after this, Voters can't be able to Vote..", vbQuestion + vbYesNo, "Delete")
If wish = vbYes Then
    rs1.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
    If rs1.EOF Then
        MsgBox "No data present in Election table..", vbInformation, "Result_Visibility"
       rs1.Close
    ElseIf rs1(4) = 0 Then
       MsgBox "Voting is already disabled for users...", vbInformation, "Result_Visibility"
       rs1.Close
    Else
        rs1(4) = 0 'disabling voting
       rs1(5) = 0 'disabling result_page visibility
       rs1.Update
      rs1.Close
      MsgBox "Voting has been Disabled for users. Now, Users can't able to caste their vote...", vbInformation, "Result_Visibility"
    End If
End If
End Sub

Private Sub cmdedit_Click()
'To enable data fields and buttons
 txtelectionname.Enabled = True
 txtdescription.Enabled = True
 txtstartingtime.Enabled = True
 txtendingtime.Enabled = True
 cmdsubmit.Enabled = True
  
 Frame3.Enabled = False
 MsgBox "After editing click submit button for updation.", vbInformation, "Election"
End Sub

Private Sub cmdenable_Click()
wish = MsgBox("Are you sure to Enable voting? Before enabling voting please check voting details...(Y/N)?", vbQuestion + vbYesNo, "Delete")
If wish = vbYes Then
    rs1.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
    If rs1.EOF Then
       MsgBox "No data present in Election table..", vbInformation, "Election"
       rs1.Close
    ElseIf rs1(4) = 1 Then
       MsgBox "Voting is already Enabled for users...", vbInformation, "Election"
       rs1.Close
    Else
       rs1(4) = 1
       rs1.Update
       rs1.Close
       MsgBox "Voting has been Enabled for users. Now,Users can be able to cast their vote...", vbInformation, "Result_Visibility"
    End If
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
MDIadmin.Show
End Sub


Private Sub cmdsubmit_Click()
If txtelectionname.Text <> "" And txtdescription.Text <> "" And txtstartingtime.Text <> "" And txtendingtime.Text <> "" Then
 rs(0) = UCase(txtelectionname.Text)
 rs(1) = txtdescription.Text
 rs(2) = UCase(txtstartingtime.Text)
 rs(3) = UCase(txtendingtime.Text)
 rs.Update
 rs.Requery
 MsgBox "Record saved", vbInformation, "ELECTION"
 Frame3.Enabled = True
 cmdedit.Enabled = True
 cmdsubmit.Enabled = False
Else
 MsgBox "Field sholud not be left blank"
End If
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"
rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
If rs.EOF <> True Then
 'to Load data in fields
   txtelectionname.Text = rs(0)
   txtdescription.Text = rs(1)
   txtstartingtime.Text = rs(2)
   txtendingtime.Text = rs(3)
 'To disable data fields and buttons
   txtelectionname.Enabled = False
   txtdescription.Enabled = False
   txtstartingtime.Enabled = False
   txtendingtime.Enabled = False
   cmdsubmit.Enabled = False
Else
   frmelection.Show
   cmdedit.Enabled = False
   Frame3.Enabled = False
   MsgBox "No data is available to show!!  So, Please Enter details of election....", vbInformation, "Election_Details"
   rs.AddNew
End If

Width = 21000
Top = 0
Left = 0
Height = 11000
End Sub

