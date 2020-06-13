VERSION 5.00
Begin VB.Form Table_Creation 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3840
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
      Begin VB.CommandButton cmdnomineelist 
         Caption         =   "nomineelist"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdvoterlist 
         Caption         =   "Voterlist"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdvoteuser 
         Caption         =   "Vote_User"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdelection 
         Caption         =   "Election"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdadmin 
         Caption         =   "Admin"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Table creation"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Table_Creation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub cmdadmin_Click()
On Error GoTo create_table
rs.Open "Select * from admin", conn
    MsgBox "Admin table already exist"
rs.Close

Exit Sub

create_table:
    conn.Execute " CREATE TABLE admin(" & _
            "adminid VARCHAR2(20) Primary key," & _
            "password VARCHAR2(20) not null," & _
            "name VARCHAR2(20) not null)"
    MsgBox "Admin table createed", vbInformation
    conn.Execute "Insert into admin values('m','p','Manish')"
    conn.Execute "Insert into admin values('s','p','Sunny')"
    conn.Execute "Insert into admin values('d','p','Dharam')"
    MsgBox "Admin table createed and data inserted....", vbInformation
End Sub

Private Sub cmdelection_Click()
On Error GoTo create_table
rs.Open "Select * from election", conn
    MsgBox "Election table already exist"
rs.Close

Exit Sub


create_table:
    conn.Execute " CREATE TABLE election(" & _
            "electionname VARCHAR2(50) not null," & _
            "description VARCHAR2(250) not null," & _
            "starttime VARCHAR2(30) not null," & _
            "endtime VARCHAR2(30) not null," & _
            "vote_enable number(1) default 0," & _
            "rslt_enable number(1) default 0)"
       MsgBox "Election table createed ....", vbInformation
End Sub

Private Sub cmdnomineelist_Click()
On Error GoTo create_table
rs.Open "Select * from nomineelist", conn
    MsgBox "Nomineelist table already exist"
rs.Close

Exit Sub


create_table:
    conn.Execute " CREATE TABLE nomineelist(" & _
           "nomineeid VARCHAR2(20) Primary key," & _
            "password VARCHAR2(20) not null," & _
            "name VARCHAR2(20) not null," & _
            "father VARCHAR2(20) not null," & _
            "academic VARCHAR2(20) not null," & _
            "gender VARCHAR2(10) not null," & _
            "dob date not null," & _
            "impath VARCHAR2(100) not null," & _
            "address VARCHAR2(100) not null," & _
            "vote_count number(5) default 0)"
   MsgBox "Nomineelist table createed ....", vbInformation
End Sub

Private Sub cmdvoterlist_Click()
On Error GoTo create_table
rs.Open "Select * from voterlist", conn
    MsgBox "voterlist table already exist"
rs.Close

Exit Sub


create_table:
    conn.Execute " CREATE TABLE voterlist(" & _
            "voterid VARCHAR2(20) Primary key," & _
            "password VARCHAR2(20) not null," & _
            "name VARCHAR2(20) not null," & _
            "father VARCHAR2(20) not null," & _
            "gender VARCHAR2(10) not null," & _
            "dob date not null," & _
            "impath VARCHAR2(100) not null," & _
            "address VARCHAR2(100) not null)"
    MsgBox "voterlist table createed ....", vbInformation
End Sub



Private Sub cmdvoteuser_Click()
On Error GoTo create_table
rs.Open "Select * from vote_user", conn
    MsgBox "Vote_user table already exist"
rs.Close

Exit Sub


create_table:
    conn.Execute " CREATE TABLE VOTE_USER(" & _
            "userid varchar2(20)," & _
            "nomineeid varchar2(20))"
            
    MsgBox "vote_User table createed ....", vbInformation
             
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"
End Sub
