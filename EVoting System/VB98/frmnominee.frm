VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnominee 
   Caption         =   "Nominee_Management"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form2"
   Picture         =   "frmnominee.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
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
      Height          =   4935
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   6255
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160628737
         CurrentDate     =   43348
      End
      Begin VB.TextBox txtqualification 
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmnominee.frx":1DA7F
         Left            =   1320
         List            =   "frmnominee.frx":1DA8C
         TabIndex        =   29
         Text            =   "Gender"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmdimage 
         Caption         =   "Select your image here"
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtuserid 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtfather 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtaddress 
         Height          =   405
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   4440
         Width           =   4455
      End
      Begin VB.TextBox txtpassword 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   " Academic Qualification:"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   2535
         Left            =   2880
         Picture         =   "frmnominee.frx":1DAA5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Nominee_img"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lbluserid 
         Caption         =   "User_Id"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblpassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Father"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Addrress"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "D.O.B"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   3840
         Picture         =   "frmnominee.frx":1F938
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   6
      Top             =   6120
      Width           =   5895
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   7560
      TabIndex        =   1
      Top             =   6240
      Width           =   5895
      Begin VB.CommandButton cmdfirst 
         Caption         =   "FIRST"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdlast 
         Caption         =   "LAST"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "NEXT"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdprevious 
         Caption         =   "PREVIOUS"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   7680
      Width           =   2175
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   4815
      Left            =   6840
      ScaleHeight     =   4755
      ScaleWidth      =   6075
      TabIndex        =   25
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Nominee Management"
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
      Left            =   3000
      TabIndex        =   26
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmnominee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
'MDIadmin.Show
'Unload Me
If nomineelist = "user" Then
frmuser.Show
End If
If nomineelist = "admin" Then
MDIadmin.Show
End If
Unload Me
End Sub

Private Sub cmdexit_Click()
'Unload Me
'MDIadmin.Show
If nomineelist = "user" Then
frmuser.Show
End If
If nomineelist = "admin" Then
MDIadmin.Show
End If
Unload Me
End Sub

Private Sub cmdhide_Click()
txtpassword.PasswordChar = "*"
End Sub

Private Sub cmdimage_Click()
Dim img1 As String
CommonDialog1.ShowOpen
img1 = CommonDialog1.FileName
Image1.Picture = LoadPicture(img1)
End Sub

