VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmreport 
   BackColor       =   &H00FFFF80&
   Caption         =   "frmReport"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16380
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3735
      Left            =   15000
      TabIndex        =   12
      Top             =   5520
      Width           =   3495
      Begin VB.CommandButton cmdclear_all 
         Caption         =   "Clear_All_Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "This will clear all data from ""ALL"" tables"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CommandButton cmdclear_nomineelist 
         Caption         =   "Clear_Nomineeelist"
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
         TabIndex        =   17
         ToolTipText     =   "This will clear all data from NOMINEELIST tabl"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton cmdclear_voterlist 
         Caption         =   "Clear_Voterlist"
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
         Left            =   600
         TabIndex        =   16
         ToolTipText     =   "This will clear all data from VOTERLIST table"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdclear_election 
         Caption         =   "Clear_Election_Data"
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
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "This will clear all data from ELECTION tabl"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CommandButton cmdclear_voting 
         Caption         =   "Clear_Voting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   13
         ToolTipText     =   "This will delete all data from VOTE_USER  table"
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Vote_User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   15000
      TabIndex        =   9
      Top             =   1200
      Width           =   3495
      Begin MSAdodcLib.Adodc adodc_vote_user 
         Height          =   330
         Left            =   1320
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         RecordSource    =   "vote_user"
         Caption         =   "Adodc5"
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmreport.frx":0000
         Height          =   3135
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
         DefColWidth     =   90
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
         Caption         =   "Vote_User"
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
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Election_Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7080
      TabIndex        =   7
      Top             =   1200
      Width           =   7695
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmreport.frx":001E
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         DefColWidth     =   80
         ForeColor       =   192
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Election"
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
      Begin MSAdodcLib.Adodc adodc_election 
         Height          =   330
         Left            =   2520
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         RecordSource    =   "election"
         Caption         =   "Adodc4"
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
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF80FF&
      Caption         =   "Nominee_List"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   6600
      Width           =   12495
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmreport.frx":003B
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         DefColWidth     =   87
         ForeColor       =   192
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Nominee_List"
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
      Begin MSAdodcLib.Adodc adodc_nominee 
         Height          =   330
         Left            =   3840
         Top             =   120
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
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
         Caption         =   "           ""Nominee_List"""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Voter_List"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   12495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmreport.frx":0057
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483634
         DefColWidth     =   97
         ForeColor       =   192
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSAdodcLib.Adodc adodc_voter 
         Height          =   330
         Left            =   3720
         Top             =   120
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   582
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
         Caption         =   "              ""Voter_List"" "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Admin_table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmreport.frx":0071
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         DefColWidth     =   87
         ForeColor       =   192
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Admin"
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
      Begin MSAdodcLib.Adodc adodc_admin 
         Height          =   330
         Left            =   1800
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         BackColor       =   16777215
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
         RecordSource    =   "admin"
         Caption         =   "                "" ADMIN  DATA"""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "<<--Back"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report And Database Table"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub cmdback_Click()
MDIadmin.Show
Unload Me
End Sub


Private Sub cmdclear_all_Click()
Dim wish As Integer
wish = MsgBox("You are going to clear all data  from ALL table other than Admin table. Do you really want to clear data(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
If wish = vbYes Then
 'Reconformation
 wish = MsgBox("_ _ _ _Reconfirmation_ _ _ _You are going to clear all data  from 'all tables'. Do you really want to clear data(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
 If wish = vbYes Then
  'To alter the value of vote_count as 0 in nomineelist table
  rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
   Do While (rs.EOF <> True)
     rs.Delete
      rs.MoveNext
   Loop
   rs.Close
   adodc_nominee.Refresh
  Else
   rs.Close
   MsgBox "There is no data present in nomineelist"
  End If
  
  'To clear data from vote_user table
  rs.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_vote_user.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in vote_user table"
  End If
  
 'To dalete data from Election Table
  rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_election.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in election table"
  End If
  
  'To delete data from voterlist table
  rs.Open "select * from voterlist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_voter.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in voterlist table"
  End If
 End If
End If
End Sub

Private Sub cmdclear_election_Click()
Dim wish As Integer
wish = MsgBox("You are going to clear all data  from election table and clear all data from Vote_user table. Do you really want to clear voting(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
If wish = vbYes Then
   
  'To alter the value of vote_count as 0 in nomineelist table
  rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
   Do While (rs.EOF <> True)
      rs(9) = 0
      rs.Update
      rs.MoveNext
   Loop
   rs.Requery
   rs.Close
   adodc_nominee.Refresh
  Else
   rs.Close
   MsgBox "There is no data present in nomineelist"
  End If
  
  'To clear data from vote_user table
  rs.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_vote_user.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in vote_user table"
  End If
  
 'To dalete data from Election Table
  rs.Open "select * from election", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_election.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in election table"
  End If
End If
End Sub

Private Sub cmdclear_voting_Click()
Dim wish As Integer
wish = MsgBox("Clear_voting will clear vote_count from all nominee table and alter them as 0 and clear all data from Vote_user table. Do you still want to clear voting(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
If wish = vbYes Then
  'To alter the value of vote_count as 0 in nomineelist table
  rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
   Do While (rs.EOF <> True)
      rs(9) = 0
      rs.Update
      rs.MoveNext
   Loop
   rs.Requery
   rs.Close
   adodc_nominee.Refresh
  Else
   rs.Close
   MsgBox "There is no data present in nomineelist table"
  End If
  
  'To clear data from vote_user table
  rs.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_vote_user.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in vote_user table"
  End If
End If
End Sub

Private Sub cmdclear_nomineelist_Click()
Dim wish As Integer
wish = MsgBox("You are going to clear all from nomineelist table and clear all data from Vote_user table. Do you really want to clear Nomineelist(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
If wish = vbYes Then
  'To alter the value of vote_count as 0 in nomineelist table
  rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
   Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
   Loop
   adodc_nominee.Refresh
   rs.Close
  Else
   rs.Close
   MsgBox "There is no data present in nomineelist"
  End If
  
  'To clear data from vote_user table
  rs.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_vote_user.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in vote_user table"
  End If
End If
End Sub

Private Sub cmdclear_voterlist_Click()
Dim wish As Integer
wish = MsgBox("You are going to clear all data from voterlist table and clear all data from Vote_user table. Do you really want to clear voterlist(Y/N)?", vbQuestion + vbYesNo + vbCritical, "Delete")
If wish = vbYes Then
  'To alter the value of vote_count as 0 in nomineelist table
  rs.Open "select * from nomineelist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
   Do While (rs.EOF <> True)
      rs(9) = 0
      rs.Update
      rs.MoveNext
   Loop
   rs.Requery
   rs.Close
   adodc_nominee.Refresh
  Else
   rs.Close
   MsgBox "There is no data present in nomineelist"
  End If
  
  'To clear data from vote_user table
  rs.Open "select * from vote_user", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_vote_user.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in vote_user table"
  End If
  'To delete data from voterlist table
  rs.Open "select * from voterlist", conn, adOpenDynamic, adLockOptimistic
  If rs.EOF <> True Then
    Do While (rs.EOF <> True)
      rs.Delete
      rs.MoveNext
    Loop
    rs.Close
    adodc_voter.Refresh
  Else
    rs.Close
    MsgBox "There is no data present in voterlist table"
  End If
End If
End Sub


Private Sub cmdexit_Click()
MDIadmin.Show
Unload Me
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

Private Sub Form_LostFocus()
Unload Me
End Sub

