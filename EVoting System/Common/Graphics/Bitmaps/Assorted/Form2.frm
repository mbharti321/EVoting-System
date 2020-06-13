VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmvoterlist 
   BackColor       =   &H00000009&
   Caption         =   "Voterlist_Management"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16530
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   16530
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15120
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=password;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=password;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "password"
      RecordSource    =   "voterlist"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":C035F
      Height          =   2895
      Left            =   2640
      TabIndex        =   20
      Top             =   7080
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      DefColWidth     =   107
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Voter_List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14280
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Voter_Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4335
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   13095
      Begin VB.Frame Frame2 
         BackColor       =   &H008080FF&
         Caption         =   "Manipulation Botton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   5535
         Begin VB.CommandButton cmdadd 
            Caption         =   "ADD"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "SAVE"
            Height          =   375
            Left            =   1080
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "Update"
            Height          =   375
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "DELETE"
            Height          =   375
            Left            =   3240
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdsearch 
            Caption         =   "SEARCH"
            Height          =   375
            Left            =   4440
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H008080FF&
         Caption         =   "Navigation Botton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   25
         Top             =   3480
         Width           =   5415
         Begin VB.CommandButton cmdfirst 
            Caption         =   "FIRST"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdlast 
            Caption         =   "LAST"
            Height          =   255
            Left            =   4080
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "NEXT"
            Height          =   255
            Left            =   1560
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdprevious 
            Caption         =   "PREVIOUS"
            Height          =   255
            Left            =   2760
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Left            =   3840
         TabIndex        =   22
         Top             =   3360
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton cmdpswdupdate 
            Caption         =   "Update"
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
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdchangepswd 
            Caption         =   "Change_Password"
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
            TabIndex        =   23
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdcancleadding 
         Caption         =   "Cancle_Adding"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   21
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11760
         TabIndex        =   11
         Top             =   3480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   125501441
         CurrentDate     =   36526
      End
      Begin VB.CommandButton cmdimage 
         Caption         =   "Select image here"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbgender 
         Height          =   315
         ItemData        =   "Form2.frx":C0374
         Left            =   5280
         List            =   "Form2.frx":C0381
         TabIndex        =   6
         Text            =   "Gender"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         MaxLength       =   20
         TabIndex        =   4
         Top             =   405
         Width           =   1455
      End
      Begin VB.TextBox txtaddress 
         Height          =   735
         Left            =   1320
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox txtfather 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtvoterid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   3375
         Left            =   9480
         Picture         =   "Form2.frx":C039A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3600
      End
      Begin VB.Label Label9 
         Caption         =   "Voter_image"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgvoter 
         Height          =   2535
         Left            =   6960
         Picture         =   "Form2.frx":C1EED
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "D.O.B"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Gender"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Addrress"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Father"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblpassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbluserid 
         BackStyle       =   0  'Transparent
         Caption         =   "voter_id"
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
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VOterlist Management"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   735
      Left            =   6360
      TabIndex        =   0
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "frmvoterlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_pswd As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim img1 As String
'Add button
Private Sub cmdadd_Click()
txtvoterid.Text = ""
txtpassword.Text = ""
txtname.Text = ""
txtfather.Text = ""
cmbgender.Text = "Gender"
DTPicker1.Value = "01-01-2000"
img1 = ""
imgvoter.Picture = LoadPicture("")
txtaddress.Text = ""
If rs.EOF <> True Then
  rs.MoveLast
End If
rs.AddNew
'for enabling data entry fields
txtvoterid.Enabled = True
txtpassword.Enabled = True
txtname.Enabled = True
txtfather.Enabled = True
cmbgender.Enabled = True
DTPicker1.Enabled = True
txtaddress.Enabled = True
cmdimage.Enabled = True
txtvoterid.SetFocus
'To enable Save button
 cmdsave.Enabled = True
'for disabling buttons
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdsearch.Enabled = False
Frame3.Enabled = False
'for enabling cancle_addig button
cmdcancleadding.Visible = True
cmdcancleadding.Enabled = True
'To disable add button
 cmdadd.Enabled = False
End Sub

Private Sub cmdback_Click()
If Module1.voterlist = "user" Then
frmuser.Show
End If
If Module1.voterlist = "admin" Then
MDIadmin.Show
End If

Unload Me
End Sub

Private Sub cmdcancleadding_Click()
'for enableabling buttons which are disabled by add_button
 rs.CancelUpdate
 rs.Requery
 cmdadd.Enabled = True
 cmdupdate.Enabled = True
 cmddelete.Enabled = True
 cmdsearch.Enabled = True
 Frame3.Enabled = True
 frmvoterlist.Refresh
 'for loading data in datafields
 If rs.EOF <> True Then
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    frmvoterlist.txtaddress.Text = rs(7)
  Else
    'To clear data field
     frmvoterlist.txtvoterid.Text = ""
     frmvoterlist.txtpassword.Text = ""
     frmvoterlist.txtname.Text = ""
     frmvoterlist.txtfather.Text = ""
     frmvoterlist.cmbgender.Text = ""
     frmvoterlist.DTPicker1.Value = "01-01-2000"
     'to store the path of image
      img1 = ""
     frmvoterlist.imgvoter.Picture = LoadPicture("")
     frmvoterlist.txtaddress.Text = ""
     
     cmdupdate.Enabled = False
  End If
    
'for disabling data entry fields
 txtvoterid.Enabled = False
 txtpassword.Enabled = False
 txtname.Enabled = False
 txtfather.Enabled = False
 cmbgender.Enabled = False
 DTPicker1.Enabled = False
 txtaddress.Enabled = False
 cmdimage.Enabled = False
'to disable Save button
 cmdsave.Enabled = False
    
'To hide cancle_adding button
 cmdcancleadding.Visible = False
 cmdcancleadding.Enabled = False
 
End Sub

Private Sub cmdchangepswd_Click()
txtpassword.Enabled = True
MsgBox "Press update after changing password to change your password", vbInformation, "User_Details"
txtpassword.SetFocus
End Sub

Private Sub cmddelete_Click()
Dim wish As Integer
On Error GoTo errmsg
wish = MsgBox("Are you sure to delete(Y/N)?", vbQuestion + vbYesNo, "Delete")
If wish = vbYes Then
  If rs.EOF Then
        MsgBox "Record not found for deletion"
  Else
    rs.Delete
    rs.Requery
    'rs.MoveNext
    If rs.EOF Then
     'To clear data field
     frmvoterlist.txtvoterid.Text = ""
     frmvoterlist.txtpassword.Text = ""
     frmvoterlist.txtname.Text = ""
     frmvoterlist.txtfather.Text = ""
     frmvoterlist.cmbgender.Text = ""
     frmvoterlist.DTPicker1.Value = "01-01-2000"
     'to store the path of image
      img1 = ""
     frmvoterlist.imgvoter.Picture = LoadPicture("")
     frmvoterlist.txtaddress.Text = ""
     
     cmdupdate.Enabled = False
     Adodc1.Refresh
     Exit Sub
    End If
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    frmvoterlist.txtaddress.Text = rs(7)
  End If
End If
Adodc1.Refresh

Exit Sub
errmsg:
        MsgBox "Record not found for deletion"
End Sub

Private Sub cmdexit_Click()
If Module1.voterlist = "user" Then
frmuser.Show
End If
If Module1.voterlist = "admin" Then
MDIadmin.Show
End If

Unload Me
End Sub



Private Sub cmdfirst_Click()
On Error GoTo errmsg
rs.MoveFirst
frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    frmvoterlist.txtaddress.Text = rs(7)
    Exit Sub
errmsg:
        MsgBox "Record not found"
    
End Sub

Private Sub cmdimage_Click()
On Error GoTo errmsg
CommonDialog1.ShowOpen
img1 = CommonDialog1.FileName
imgvoter.Picture = LoadPicture(img1)
Exit Sub

errmsg:
   img1 = ""
   MsgBox "please select a valid image with jpg/bmp extention.Don't select .png image.", vbCritical, "Image_Selection"
End Sub

Private Sub cmdlast_Click()
On Error GoTo errmsg
rs.MoveLast
frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    frmvoterlist.txtaddress.Text = rs(7)
    Exit Sub
errmsg:
        MsgBox "Record not found"
    
End Sub

Private Sub cmdnext_Click()
On Error GoTo errmsg
rs.MoveNext
If rs.EOF Then
    rs.MoveLast
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.txtaddress.Text = rs(7)
    MsgBox "this the last record in list."
Else
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.txtaddress.Text = rs(7)
End If
Exit Sub
errmsg:
        MsgBox "Record not found"
End Sub

Private Sub cmdprevious_Click()
On Error GoTo errmsg
rs.MovePrevious
If rs.BOF Then
    rs.MoveFirst
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.txtaddress.Text = rs(7)
    MsgBox "this the first record in list."
Else
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    'to store the path of image
     img1 = rs(6)
     frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    frmvoterlist.txtaddress.Text = rs(7)
End If
Exit Sub
errmsg:
        MsgBox "Record not found"
End Sub

Private Sub cmdpswdupdate_Click()
Dim pswd As String
pswd = InputBox("Please enter your password for confirmation!!", "Password_Confirmation")
If pswd <> txtpassword.Text Then
    MsgBox "Password didn't match....So, Please recheck entered password, and try again..", vbCritical, "User_Details"
Else
    rs_pswd.Open "Select * from voterlist where voterid='" & txtvoterid.Text & "'", conn, adOpenDynamic, adLockOptimistic, adCmdText
    rs_pswd(1) = txtpassword.Text
    rs_pswd.Update
    rs_pswd.Close
    MsgBox "Your password has been changed", vbInformation, "User_Details"
End If
End Sub

'Save botton
Private Sub cmdsave_Click()
'To check whether any field is empty
If txtvoterid.Text <> "" And txtpassword.Text <> "" And txtname.Text <> "" And txtfather.Text <> "" And cmbgender.Text <> "" And DTPicker1.Value <> "" And img1 <> "" And txtaddress.Text <> "" Then
 rs(0) = UCase(txtvoterid.Text)
 rs(1) = txtpassword.Text
 rs(2) = UCase(txtname.Text)
 rs(3) = UCase(txtfather.Text)
 rs(4) = cmbgender.Text
 rs(5) = DTPicker1.Value
 rs(6) = img1
 rs(7) = UCase(txtaddress.Text)
 rs.Update
 MsgBox "Record saved !!!!!_ _Going to load First voter's details_ _", vbInformation, "Voter_List"
 rs.Requery
 'To load data of first voter's details
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
     frmvoterlist.txtaddress.Text = rs(7)
 'To refresh database and data grid
   Adodc1.Refresh
 'To disable data fields
 txtvoterid.Enabled = False
 txtpassword.Enabled = False
 txtname.Enabled = False
 txtfather.Enabled = False
 cmbgender.Enabled = False
 DTPicker1.Enabled = False
 txtaddress.Enabled = False
 cmdimage.Enabled = False
 'to disable Save button
 cmdsave.Enabled = False
 'for disabling cancle_add button
 cmdcancleadding.Visible = False
 
'for enableabling buttons which are disabled by add_button
 cmdadd.Enabled = True
 cmdupdate.Enabled = True
 cmddelete.Enabled = True
 cmdsearch.Enabled = True
 Frame3.Enabled = True
 'rs.Close
Else
 MsgBox "Field sholud not be left blank"
End If
End Sub
'Search button
Private Sub cmdsearch_Click()
Dim userid As String
userid = Trim(InputBox("Enter voterid which you want to search"))
On Error GoTo errmsg
'rs.Open "select * from voterlist", conn, adOpenDynamic, adLockOptimistic
rs.Find "voterid=" & UCase(userid)
If rs.EOF Then
 MsgBox "Voterid not found"
 rs.Requery
 
Else
 'for loading data in the form with first voterlist
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
    frmvoterlist.txtaddress.Text = rs(7)
 'for disabling data entry fields
    txtvoterid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    txtaddress.Enabled = False
    cmdimage.Enabled = False
End If
Exit Sub

errmsg:
   MsgBox "Plesae enter a valid voterid"
End Sub
'Update button
Private Sub cmdupdate_Click()
MsgBox "Modify and click save to update changes"
'To disable voter_id text_field
txtvoterid.Enabled = False
'To disable buttons
cmdadd.Enabled = False
cmdupdate.Enabled = False
cmdsearch.Enabled = False

'to enable datafied
txtpassword.Enabled = True
txtname.Enabled = True
txtfather.Enabled = True
cmbgender.Enabled = True
DTPicker1.Enabled = True
txtaddress.Enabled = True
cmdimage.Enabled = True
txtpassword.SetFocus
'To enable Save button
 cmdsave.Enabled = True
End Sub





'To check whether the user is 18 years old or not
Private Sub DTPicker1_LostFocus()
Dim age As Integer
age = (DateValue(Date) - DateValue(DTPicker1.Value)) / 365
If age < 18 Then
MsgBox "Voter must be 18 or more years old.So, Please select a proper date of Birth.", vbInformation + vbOKOnly, "Date_Of_Birth"
DTPicker1.Value = "01-01-2000"
End If
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs_pswd = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=password;Persist Security Info=True;User ID=system"

rs.Open "select * from voterlist", conn, adOpenDynamic, adLockOptimistic
  'for disabling data entry fields
    txtvoterid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    txtaddress.Enabled = False
    cmdimage.Enabled = False
    'to disable Save button
     cmdsave.Enabled = False
    'To hide cancle_adding button
     cmdcancleadding.Visible = False
   
   'for loading data in the form with first voterlist
   If rs.EOF <> True Then
    frmvoterlist.txtvoterid.Text = rs(0)
    frmvoterlist.txtpassword.Text = rs(1)
    frmvoterlist.txtname.Text = rs(2)
    frmvoterlist.txtfather.Text = rs(3)
    frmvoterlist.cmbgender.Text = rs(4)
    frmvoterlist.DTPicker1.Value = rs(5)
    frmvoterlist.imgvoter.Picture = LoadPicture(rs(6))
    'to store the path of image
     img1 = rs(6)
     frmvoterlist.txtaddress.Text = rs(7)
   Else
    cmdupdate.Enabled = False 'if there is no data then it doesnot make sence of updating records
   End If


Width = 21000
Top = 0
Left = 0
Height = 11000

End Sub




Private Sub txtpassword_GotFocus()
If txtpassword.Text = "" Then
    Dim p1 As String
    Dim p2 As String
    Dim p3 As String
    Dim p4 As String
    Dim p5 As String
    Dim p6 As String
    p1 = Int(Rnd() * 26) + 97
    p2 = Int(Rnd() * 26) + 97
    p3 = Int(Rnd() * 10) + 48
    p4 = Int(Rnd() * 10) + 48
    p5 = Int(Rnd() * 26) + 97
    p6 = Int(Rnd() * 26) + 97
    Dim pswd As String
    pswd = Chr(p1) + Chr(p2) + Chr(p3) + Chr(p4) + Chr(p5) + Chr(p6)
    txtpassword.Text = pswd
End If
End Sub

Private Sub txtvoterid_LostFocus()
Dim userid As String
If txtvoterid.Enabled = True Then
userid = UCase(txtvoterid.Text)
 'TO CHECK IN VOTERLIST
  rs1.Open "select * from voterlist where voterid='" & userid & "'", conn, adOpenDynamic, adLockOptimistic
  If rs1.EOF = True Then
   'TO CHEK IN NOMINEELIST
    rs2.Open "select * from nomineelist where nomineeid='" & userid & "'", conn, adOpenDynamic, adLockOptimistic
     If rs2.EOF = True Then
      rs1.Close
      rs2.Close
     Else
       MsgBox "Userid is already exist in Nomineelist.Please choose an unique useridid which does not exist."
       rs1.Close
       rs2.Close
       txtvoterid.Text = ""
       txtvoterid.SetFocus
       Exit Sub
     End If
  Else
   MsgBox "Userid is already exist in Voterlist.Please choose an unique Userid which does not exist."
   rs1.Close
   txtvoterid.Text = ""
   txtvoterid.SetFocus
End If
End If
End Sub
