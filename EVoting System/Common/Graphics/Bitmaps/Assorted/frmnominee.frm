VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmnominee 
   Caption         =   "Nominee_Management"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16380
   LinkTopic       =   "Form2"
   Picture         =   "frmnominee.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14160
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "nomineelist"
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
      Bindings        =   "frmnominee.frx":C035F
      Height          =   2295
      Left            =   2400
      TabIndex        =   22
      Top             =   6960
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   93
      Enabled         =   -1  'True
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nominee_Details"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Nominee_Details"
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
      Height          =   4575
      Left            =   2760
      TabIndex        =   12
      Top             =   2040
      Width           =   12855
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
         Left            =   6000
         TabIndex        =   33
         Top             =   3720
         Width           =   5295
         Begin VB.CommandButton cmdfirst 
            Caption         =   "FIRST"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdlast 
            Caption         =   "LAST"
            Height          =   255
            Left            =   3960
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "NEXT"
            Height          =   255
            Left            =   1440
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdprevious 
            Caption         =   "PREVIOUS"
            Height          =   255
            Left            =   2640
            TabIndex        =   34
            Top             =   360
            Width           =   1095
         End
      End
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
         TabIndex        =   27
         Top             =   3720
         Width           =   5655
         Begin VB.CommandButton cmdadd 
            Caption         =   "ADD"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "SAVE"
            Height          =   375
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdupdate 
            Caption         =   "Update"
            Height          =   375
            Left            =   2160
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "DELETE"
            Height          =   375
            Left            =   3240
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdsearch 
            Caption         =   "SEARCH"
            Height          =   375
            Left            =   4560
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Left            =   4080
         TabIndex        =   24
         Top             =   3600
         Visible         =   0   'False
         Width           =   3975
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
            TabIndex        =   26
            Top             =   240
            Width           =   2295
         End
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
            TabIndex        =   25
            Top             =   240
            Width           =   1095
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
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000B&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11640
         TabIndex        =   10
         Top             =   3840
         Width           =   1095
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
         Height          =   405
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtaddress 
         Height          =   525
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox txtfather 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtnomineeid 
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
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdimage 
         Caption         =   "Select image here"
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cmbgender 
         Height          =   315
         ItemData        =   "frmnominee.frx":C0374
         Left            =   5160
         List            =   "frmnominee.frx":C0381
         TabIndex        =   4
         Text            =   "Gender"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtqualification 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   135856129
         CurrentDate     =   43351
      End
      Begin VB.Image imgnominee 
         Height          =   2535
         Left            =   6720
         Picture         =   "frmnominee.frx":C039A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "D.O.B"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Gender"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Addrress"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Father"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblpassword 
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
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbluserid 
         Caption         =   "User_Id"
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
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Nominee_img"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   3375
         Left            =   9600
         Picture         =   "frmnominee.frx":DCAFE
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   " Academic Qualification:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14520
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nominee Management"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   6240
      TabIndex        =   0
      Top             =   840
      Width           =   7455
   End
End
Attribute VB_Name = "frmnominee"
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
txtnomineeid.Text = ""
txtpassword.Text = ""
txtname.Text = ""
txtfather.Text = ""
cmbgender.Text = "Gender"
DTPicker1.Value = "01-01-2000"
txtqualification = ""
img1 = ""
imgnominee.Picture = LoadPicture("")
txtaddress.Text = ""
If rs.EOF <> True Then
  rs.MoveLast
End If
rs.AddNew
'for enabling data entry fields
    txtnomineeid.Enabled = True
    txtpassword.Enabled = True
    txtname.Enabled = True
    txtfather.Enabled = True
    txtqualification.Enabled = True
    cmbgender.Enabled = True
    DTPicker1.Enabled = True
    cmdimage.Enabled = True
    txtaddress.Enabled = True
    txtnomineeid.SetFocus
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

'Back button
Private Sub cmdback_Click()
If Module1.nomineelist = "user" Then
frmuser.Show
End If
If Module1.nomineelist = "admin" Then
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
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    txtaddress.Text = rs(8)
  Else
    'To clear data field
      txtnomineeid.Text = ""
     txtpassword.Text = ""
     txtname.Text = ""
     txtfather.Text = ""
     txtqualification.Text = ""
     cmbgender.Text = ""
     DTPicker1.Value = "01-01-2000"
     imgnominee.Picture = LoadPicture("")
     txtaddress.Text = ""
    
    cmdupdate.Enabled = False
  End If
    
'for disabling datafields
    txtnomineeid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    txtqualification.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    cmdimage.Enabled = False
    txtaddress.Enabled = False
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
     txtnomineeid.Text = ""
     txtpassword.Text = ""
     txtname.Text = ""
     txtfather.Text = ""
     txtqualification.Text = ""
     cmbgender.Text = ""
     DTPicker1.Value = "01-01-2000"
     imgnominee.Picture = LoadPicture("")
     txtaddress.Text = rs(8)
    
     cmdupdate.Enabled = False
     Adodc1.Refresh
     Exit Sub
    End If
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
    rs.Requery
    Adodc1.Refresh
  End If
End If
Adodc1.Refresh
Exit Sub

errmsg:
        MsgBox "Record not found for deletion"
End Sub

'Exit button
Private Sub cmdexit_Click()
If Module1.nomineelist = "user" Then
frmuser.Show
End If
If Module1.nomineelist = "admin" Then
MDIadmin.Show
End If
Unload Me
End Sub

Private Sub cmdfirst_Click()
On Error GoTo errmsg
rs.MoveFirst
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
    Exit Sub
errmsg:
        MsgBox "Record not found"
    
End Sub

'Image selection button
Private Sub cmdimage_Click()
On Error GoTo errmsg
CommonDialog1.ShowOpen
img1 = CommonDialog1.FileName
imgnominee.Picture = LoadPicture(img1)
Exit Sub

errmsg:
   img1 = ""
   MsgBox "please select a valid image with jpg/bmp extention.Don't select .png image.", vbCritical, "Image_Selection"

End Sub

Private Sub cmdlast_Click()
On Error GoTo errmsg
rs.MoveLast
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
    Exit Sub
errmsg:
        MsgBox "Record not found"
    
End Sub

Private Sub cmdnext_Click()
On Error GoTo errmsg
rs.MoveNext
If rs.EOF Then
    rs.MoveLast
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
    MsgBox "This is the last record in list."
Else
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
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
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
    MsgBox "This is the first record in list."
Else
   txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
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
    rs_pswd.Open "Select * from nomineelist where nomineeid='" & txtnomineeid.Text & "'", conn, adOpenDynamic, adLockOptimistic, adCmdText
    rs_pswd(1) = txtpassword.Text
    rs_pswd.Update
    rs_pswd.Close
    MsgBox "Your password has been changed", vbInformation, "User_Details"
    txtpassword.Enabled = False
End If
End Sub

Private Sub cmdsave_Click()
'To check whether any field is empty
If txtnomineeid.Text <> "" And txtpassword.Text <> "" And txtname.Text <> "" And txtfather.Text <> "" And cmbgender.Text <> "" And DTPicker1.Value <> "" And txtqualification <> "" And img1 <> "" And txtaddress.Text <> "" Then
 rs(0) = UCase(txtnomineeid.Text)
 rs(1) = txtpassword.Text
 rs(2) = UCase(txtname.Text)
 rs(3) = UCase(txtfather.Text)
 rs(4) = UCase(txtqualification.Text)
 rs(5) = cmbgender.Text
 rs(6) = DTPicker1.Value
 rs(7) = img1
 rs(8) = UCase(txtaddress.Text)
 rs.Update
 MsgBox "Record saved !!!!!_ _Going to load First nominee's details_ _ ", vbInformation, "Nominee_List"
 rs.Requery 'refresh
 'To load first nominee's data
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
 'to refreash database and datagrid
    Adodc1.Refresh
 'for disabling data entry fields
    txtnomineeid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    txtqualification.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    cmdimage.Enabled = False
    txtaddress.Enabled = False
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
'search button
Private Sub cmdsearch_Click()
Dim userid As String
userid = Trim(InputBox("Enter voterid which you want to search"))
On Error GoTo errmsg

rs.Find "nomineeid=" & UCase(userid)
If rs.EOF Then
 MsgBox "Nomineeid not found"
 rs.Requery
 
Else
 'for loading data in the form with first voterlist
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
'for disabling data entry fields
    txtnomineeid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    txtqualification.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    cmdimage.Enabled = False
    txtaddress.Enabled = False
End If
Exit Sub

errmsg:
   MsgBox "Plesae enter a valid nomineeid"
End Sub
'update button
Private Sub cmdupdate_Click()
MsgBox "Modify and click save to update changes"
'To disable nominee_id text_field
txtnomineeid.Enabled = False
'To disable buttons
cmdadd.Enabled = False
cmdupdate.Enabled = False
cmdsearch.Enabled = False

'for enabling data entry fields
    txtpassword.Enabled = True
    txtname.Enabled = True
    txtfather.Enabled = True
    txtqualification.Enabled = True
    cmbgender.Enabled = True
    DTPicker1.Enabled = True
    cmdimage.Enabled = True
    txtaddress.Enabled = True
    txtpassword.SetFocus
'To enable Save button
 cmdsave.Enabled = True
End Sub



'To check whether the user is 18 years old or not
Private Sub DTPicker1_LostFocus()
Dim age As Integer
age = (DateValue(Date) - DateValue(DTPicker1.Value)) / 365
If age < 18 Then
MsgBox "Nominee must be 18 or more years old.So, Please select a proper date of Birth.", vbInformation + vbOKOnly, "Date_Of_Birth"
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

rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
   'for disabling datafields
    txtnomineeid.Enabled = False
    txtpassword.Enabled = False
    txtname.Enabled = False
    txtfather.Enabled = False
    txtqualification.Enabled = False
    cmbgender.Enabled = False
    DTPicker1.Enabled = False
    cmdimage.Enabled = False
    txtaddress.Enabled = False
    'to disable Save button
     cmdsave.Enabled = False
    'To hide cancle_adding button
     cmdcancleadding.Visible = False
   
   'for loading data in the form with first voterlist
   If rs.EOF <> True Then 'Before loding data it will check whether data is there in table nomineelit or not
    txtnomineeid.Text = rs(0)
    txtpassword.Text = rs(1)
    txtname.Text = rs(2)
    txtfather.Text = rs(3)
    txtqualification.Text = rs(4)
    cmbgender.Text = rs(5)
    DTPicker1.Value = rs(6)
    imgnominee.Picture = LoadPicture(rs(7))
    'to save the path of image
    img1 = rs(7)
    txtaddress.Text = rs(8)
   Else
    cmdupdate.Enabled = False 'if there is no data then it doesnot make sence of updating records
   End If



Width = 21000
Top = 0
Left = 0
Height = 11000


End Sub





Private Sub txtnomineeid_LostFocus()
Dim userid As String
If txtnomineeid.Enabled = True Then
 userid = UCase(txtnomineeid.Text)
 'TO CHECK IN VOTERLIST
  rs1.Open "select * from voterlist where voterid='" & userid & "'", conn, adOpenDynamic, adLockOptimistic
  If rs1.EOF = True Then
   'TO CHEK IN NOMINEELIST
    rs2.Open "select * from nomineelist where nomineeid='" & userid & "'", conn, adOpenDynamic, adLockOptimistic
    If rs2.EOF = True Then
      rs1.Close
      rs2.Close
    Else
      MsgBox "Userid is already exist in nomineelist.Please choose an unique userid which does not exist."
      rs1.Close
      rs2.Close
      txtnomineeid.Text = ""
      txtnomineeid.SetFocus
      Exit Sub
    End If
  Else
   MsgBox "Userid is already exist in voterlist.Please choose an unique userid which does not exist."
   rs1.Close
   txtnomineeid.Text = ""
   txtnomineeid.SetFocus
 End If
End If
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

