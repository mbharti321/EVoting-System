VERSION 5.00
Begin VB.Form frmresult 
   BackColor       =   &H80000012&
   Caption         =   "Result"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   Picture         =   "frmresult.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   16170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   5040
      TabIndex        =   0
      Top             =   1320
      Width           =   10695
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1080
         TabIndex        =   15
         Top             =   1200
         Width           =   8295
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "has won the Election with"
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
            Left            =   5280
            TabIndex        =   21
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblwinnner 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "WINNER NOMINEE NAME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblnomineeid 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "nominee_id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "with nominee_id"
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
            Left            =   2280
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblcount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COUNT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5040
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "votes....."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Visibility Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   8760
         TabIndex        =   12
         Top             =   3360
         Width           =   1575
         Begin VB.CommandButton cmdhide 
            Caption         =   "Hide"
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
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdvisible 
            Caption         =   "Visible"
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
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080FF80&
         Caption         =   "Vote_Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   6975
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            ItemData        =   "frmresult.frx":5D2F1
            Left            =   5160
            List            =   "frmresult.frx":5D2F3
            Style           =   1  'Simple Combo
            TabIndex        =   22
            Text            =   "Nominee"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblvotecount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Vote_count:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblnomineename 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "nominee"
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
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1935
         End
         Begin VB.Image imgnominee 
            Height          =   1935
            Left            =   2280
            Picture         =   "frmresult.frx":5D2F5
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Nominee_Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Select nominee name here for vote details:"
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
            TabIndex        =   7
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   5
         Top             =   5640
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1440
         TabIndex        =   2
         Top             =   5520
         Width           =   6975
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total_Vote Count:"
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
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lbltotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            TabIndex        =   3
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Voting   Result"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim conn As New ADODB.Connection
Private Sub cmdback_Click()
If Module1.result = "user" Then
frmuser.Show
End If
If Module1.result = "admin" Then
MDIadmin.Show
End If
Unload Me
End Sub
Private Sub cmdexit_Click()
If Module1.result = "user" Then
frmuser.Show
End If
If Module1.result = "admin" Then
MDIadmin.Show
End If
Unload Me
End Sub

Private Sub cmdhide_Click()

Dim wish As Integer
wish = MsgBox("This will disable 'Result' page for ALL users. Do you really want to Disable?", vbQuestion + vbYesNo, "Result_Disable")
If wish <> vbYes Then
 Exit Sub
End If


rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
   MsgBox "No data present in Election table..", vbInformation, "Result_Visibility"
   rs.Close
ElseIf rs(5) = 0 Then
   MsgBox "Result_page is already disabled for users...", vbInformation, "Result_Visibility"
   rs.Close
Else
   rs(5) = 0
   rs.Update
   rs.Close
   MsgBox "Result_page has been Disabled for users. Now, This page is not visible to any user...", vbInformation, "Result_Visibility"
End If
End Sub

Private Sub cmdvisible_Click()
Dim wish As Integer
wish = MsgBox("This will enable 'Result' page for ALL users. Do you really want to Enable?", vbQuestion + vbYesNo, "Result_Enable")
If wish <> vbYes Then
 Exit Sub
End If

rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
   MsgBox "No data present in Election table..", vbInformation, "Result_Visibility"
   rs.Close
ElseIf rs(5) = 1 Then
   MsgBox "Result_page is already Enabled for users...", vbInformation, "Result_Visibility"
   rs.Close
Else
   rs(5) = 1
   rs.Update
   rs.Close
   MsgBox "Result_page has been Enabled for users. Now, This page is visible to all users...", vbInformation, "Result_Visibility"
End If
End Sub

Private Sub Combo1_DblClick()
rs1.Open "select * from nomineelist where nomineeid='" & Combo1.Text & "'", conn, adOpenDynamic, adLockOptimistic
lblnomineename.Caption = rs1(2)
lblvotecount.Caption = rs1(9)
imgnominee.Picture = LoadPicture(rs1(7))
rs1.Close
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"

rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  
'if No data present in nominee list.
  If rs.EOF Then
     MsgBox "No data available to show !!!!!!!!!"
     Frame4.Enabled = False
     Combo1.Clear
     Combo1.Enabled = False
     rs.Close
     Exit Sub
  End If
'To load nominees' nmae in combobox
  Combo1.Clear
  Combo1.Text = "Nominees!!!"
  Do While (rs.EOF <> True)
    Combo1.AddItem rs(0)
    rs.MoveNext
  Loop
  rs.Close
  
'To load total vote count
  rs1.Open "select count(*) from vote_user", conn, adOpenDynamic, adLockOptimistic
  lbltotal.Caption = rs1(0)
  rs1.Close
  
'To load the winner name and nomineeid
    Dim winner_id As String
    Dim flage As Boolean
     flage = False
    Dim maxcount As Integer
     maxcount = 0
    
    rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
    
    Do While rs.EOF <> True
      If rs(9) >= maxcount Then
         If rs(9) = maxcount Then
            flage = True
         Else
            flage = False
            maxcount = rs(9)
            winner_id = rs(0) 'To save the Nominiee_Id of the nominee having max vote count
         End If
      End If
      rs.MoveNext
    Loop
    
    If flage = True Then 'if two nominee have same max_vote_count
       frmresult.Show
       MsgBox "Can't determined the winner !!!! Reason may be either two or more nominee have same maximum_vote_count or elction hasn't started yet...", vbInformation, "Result"
       Frame5.Visible = False
       Frame3.Top = 1500    'to relocate frame3
       Frame2.Top = 4500    'to relocate frame2
       rs.Close
       Exit Sub
    Else
     rs.Requery
     rs.Find "nomineeid=" & winner_id 'finding rcordset of winner
     
     lblwinnner.Caption = rs(2)
     lblnomineeid.Caption = rs(0)
     lblcount.Caption = rs(9)
     
     lblnomineename.Caption = rs(2)
     lblvotecount.Caption = rs(9)
     imgnominee.Picture = LoadPicture(rs(7))
     Combo1.Text = rs(0)
     rs.Close
    End If

Width = 21000
Top = 0
Left = 0
Height = 11000

End Sub
