VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form add_kl 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«÷«›Â ò—œ‰ ò«·«"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   18180
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_kl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   18180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      DisabledPicture =   "add_kl.frx":030A
      DownPicture     =   "add_kl.frx":24F84
      DragIcon        =   "add_kl.frx":49BFE
      Height          =   330
      Left            =   17640
      Picture         =   "add_kl.frx":6E878
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   4800
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4440
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   9180
      Width           =   18180
      _ExtentX        =   32068
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      Begin MSAdodcLib.Adodc F_kl 
         Height          =   375
         Left            =   240
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_kl"
         Caption         =   "F_kl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_sttg 
         Height          =   375
         Left            =   240
         Top             =   2520
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_sttg"
         Caption         =   "F_sttg"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_tbq 
         Height          =   375
         Left            =   240
         Top             =   2160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_tbq"
         Caption         =   "F_tbq"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_fctr_msht 
         Height          =   375
         Left            =   240
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_fctr_msht"
         Caption         =   "F_fctr_msht"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_fctr 
         Height          =   375
         Left            =   240
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_fctr"
         Caption         =   "F_fctr"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_msht 
         Height          =   375
         Left            =   240
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_msht"
         Caption         =   "F_msht"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_usr 
         Height          =   375
         Left            =   240
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_usr"
         Caption         =   "F_usr"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_kl_gh 
         Height          =   375
         Left            =   240
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_kl_gh"
         Caption         =   "F_kl_gh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc F_gh 
         Height          =   375
         Left            =   240
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM F_gh"
         Caption         =   "F_gh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   " ’ÊÌ— ò«·«"
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   120
      Width           =   4695
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "0.0.0."
         DataField       =   "ID_kl"
         DataSource      =   "F_kl"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1080
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "À»  ò«·«"
      ForeColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   12975
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Left            =   4680
         TabIndex        =   37
         Top             =   3960
         Width           =   6975
      End
      Begin VB.ComboBox Combo3 
         Height          =   465
         Left            =   9720
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   2280
         Width           =   6495
      End
      Begin VB.CheckBox ch_Jostan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404000&
         Caption         =   "Ã” ÃÊ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   11760
         TabIndex        =   28
         Top             =   3960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000C0C0&
         Caption         =   "«’·«Õ „‘Œ’« "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   465
         Left            =   8400
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   10440
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   10440
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› ò«·«"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
         Caption         =   "À»  ò«·«"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   2160
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   8400
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   6480
         TabIndex        =   7
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   855
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2880
         Width           =   9495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "ò«·«Ì ÃœÌœ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "ê“«—‘"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   8880
         TabIndex        =   32
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "Ê÷⁄Ì "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   11880
         TabIndex        =   31
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label tedad_yaft 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   3840
         TabIndex        =   30
         Top             =   3960
         Width           =   105
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "„Ê—œ Ì«›  ‘œ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   2160
         TabIndex        =   29
         Top             =   3960
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "—Ì«·"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   2280
         TabIndex        =   27
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "„ﬁœ«—/ ⁄œ«œ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   11880
         TabIndex        =   26
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "ﬁÌ„ "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   5880
         TabIndex        =   23
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "ê—ÊÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   7560
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "òœ ò«·«"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   11880
         TabIndex        =   21
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "‰Ê⁄ ò«·«"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   7560
         TabIndex        =   20
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "„‘Œ’Â"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   11880
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   11760
         TabIndex        =   18
         Top             =   3120
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "add_kl.frx":934F2
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   30
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "·Ì”  ò«·«Â«Ì À»  ‘œÂ"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ID_kl"
         Caption         =   "òœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "code"
         Caption         =   "òœ ò«·«"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "xnoe"
         Caption         =   "‰Ê⁄ ò«·«"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "tmyz"
         Caption         =   "„‘Œ’Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "tozih"
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "qeymat"
         Caption         =   "ﬁÌ„ "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "id_usr"
         Caption         =   "ò«—»—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "xdate"
         Caption         =   " «—ÌŒ À» "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "end_event"
         Caption         =   "¬Œ—Ì‰ Ê÷⁄Ì "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "tedad"
         Caption         =   " ⁄œ«œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "vahed"
         Caption         =   "Ê«Õœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Label no_ 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label yes_ 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   960
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label question_ 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „ÿ„∆‰ Â” Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1680
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Natije_color_time 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Menu mnufindInhavalie 
      Caption         =   "Ã” ÃÊÌ «Ì‰ ò«·« œ— ÕÊ«·Â Â«"
   End
End
Attribute VB_Name = "add_kl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer_for_natije_color_ As Integer

Function TEXT_CH_JOS_IN_F_KL(filde_, str_)
On Error Resume Next

If ch_Jostan.Value = 1 Then
F_kl.Refresh
F_kl.RecordSource = "select * from F_kl where " & filde_ & " like ('%" & str_ & "%')"
F_kl.Refresh
tedad_yaft.Caption = F_kl.Recordset.RecordCount
End If
End Function
Function Clear_text()
On Error Resume Next

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

End Function
Function Add_new_record()
On Error Resume Next

If ch_Jostan.Value = 1 Then
ch_Jostan.Value = 0
Clear_text
ch_Jostan.Value = 1
Else
Clear_text
End If
End Function
Function KOD_INS_COMBO_F_STTG(KOD_, Combo_, id_)
On Error Resume Next

F_sttg.Refresh
F_sttg.RecordSource = "select * from f_sttg where kod like ('" & KOD_ & "')"
F_sttg.Refresh
For I = 1 To F_sttg.Recordset.RecordCount
If id_ = 1 Then
Combo_.AddItem (F_sttg.Recordset.Fields("id_sttg") & " _ " & F_sttg.Recordset.Fields("xvalue"))
Else
Combo_.AddItem (F_sttg.Recordset.Fields("xvalue"))
End If
F_sttg.Recordset.MoveNext
Next I
End Function
Function Find_id_f_sttg(Combo_)
On Error Resume Next

a = Split(Combo_.Text, " _ ")
F_sttg.Refresh
F_sttg.RecordSource = "select * from f_sttg where id_sttg like ('" & a(0) & "')"
F_sttg.Refresh
End Function
Function Time_color_natije(labe_, xtime_, color_, str_)
On Error Resume Next

labe_.Visible = True
Timer_for_natije_color_ = xtime_
Timer2.Enabled = True
labe_.BackColor = color_
labe_.Caption = str_

End Function


Private Sub Command1_Click()
On Error Resume Next

Add_new_record
a = Time_color_natije(Natije_color_time, 40, &H8080&, "„‘Œ’«  ò«·«Ì ÃœÌœ —« Ê«—œ ò‰Ìœ")
Text2.SetFocus

End Sub

Private Sub Command11_Click()
Dim oExcel As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub
2:
Me.F_kl.Recordset.MoveFirst

Dim BePAR As Integer

BePAR = 0
Set oExcel = GetObject(App.Path & "\formXLS\Mojod_giri.xls")


oExcel.ActiveSheet.Range("H3").Value = Taqvim.KKK.Caption

For I = 5 To F_kl.Recordset.RecordCount + 4
oExcel.ActiveSheet.Range("B" & I).Value = F_kl.Recordset.Fields("id_kl")
oExcel.ActiveSheet.Range("C" & I).Value = F_kl.Recordset.Fields("xnoe")
oExcel.ActiveSheet.Range("D" & I).Value = F_kl.Recordset.Fields("tmyz")
oExcel.ActiveSheet.Range("E" & I).Value = F_kl.Recordset.Fields("code")
oExcel.ActiveSheet.Range("F" & I).Value = F_kl.Recordset.Fields("tedad") & " " & F_kl.Recordset.Fields("vahed")
oExcel.ActiveSheet.Range("G" & I).Value = F_kl.Recordset.Fields("end_event")
oExcel.ActiveSheet.Range("H" & I).Value = F_kl.Recordset.Fields("tozih")

F_kl.Recordset.MoveNext
Next I


oExcel.Application.Visible = True
On Error GoTo 722


oExcel.Parent.Windows(2).Visible = True
GoTo 910
722:

oExcel.Parent.Windows(1).Visible = True
910:
''''''
Dim X As String
X = App.Path & "\temp\" & Taqvim.KKK.Caption & "-" & "mojodi"
On Error GoTo 9

oExcel.SaveAs X
GoTo 8
9:

oExcel.ActiveSheet.Range("H3").Value = ""

For I = 5 To F_kl.Recordset.RecordCount + 4
oExcel.ActiveSheet.Range("B" & I).Value = ""
oExcel.ActiveSheet.Range("C" & I).Value = ""
oExcel.ActiveSheet.Range("D" & I).Value = ""
oExcel.ActiveSheet.Range("E" & I).Value = ""
oExcel.ActiveSheet.Range("F" & I).Value = ""
Next I
oExcel.Close
MsgBox "Œÿ« œ— ›«Ì· Œ—ÊÃÌ: Ê÷⁄Ì  ›«Ì· Â«Ì »«“ —« »——”Ì ò‰Ìœ ::   Ìò ›«Ì· »«“ »« Â„Ì‰ „‘Œ’«  ÊÃÊœ œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
8:

End Sub

Private Sub Command2_Click()
On Error Resume Next

If Text3.Text = "" Or Text6.Text = "" Then
If Text3.Text = "" Then Text3.SetFocus
If Text6.Text = "" Then Text6.SetFocus

Beep
Exit Sub
End If

cod_1_ = Text1.Text & " :: " & Text2.Text
If cod_1_ = " :: " Then GoTo 2

F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where code like ('" & cod_1_ & "')"
F_kl.Refresh
If F_kl.Recordset.BOF = False Or F_kl.Recordset.EOF = False Then
a = Time_color_natije(Natije_color_time, 20, &H80FF&, "«Ì‰ „Ê—œ ﬁ»·« «÷«›Â ‘œÂ «” ")
Exit Sub
End If
2:

F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where xnoe like ('" & Text3.Text & "')"
F_kl.Refresh
If F_kl.Recordset.BOF = False Or F_kl.Recordset.EOF = False Then
If MsgBox("«Ì‰ ‰Ê⁄ ò«·« œ— ·Ì”  ÊÃÊœ œ«—œ ¬Ì« „Ì ŒÊ«ÂÌœ œ— Â— ’Ê—  ò«·«Ì ÃœÌœ —« «÷«›Â ò‰Ìœ", vbExclamation + vbYesNo, "À»  ò«·«") = vbYes Then
GoTo 1
Else
Exit Sub
End If
End If

1:
c = Split(Combo1.Text, " _ ")

F_kl.Refresh
F_kl.Recordset.AddNew
F_kl.Recordset.Fields("code") = cod_1_
F_kl.Recordset.Fields("xnoe") = Text3.Text
F_kl.Recordset.Fields("tmyz") = Text4.Text
F_kl.Recordset.Fields("tozih") = Text5.Text
F_kl.Recordset.Fields("qeymat") = Text7.Text
F_kl.Recordset.Fields("id_gh") = c(0)
F_kl.Recordset.Fields("id_usr") = menu.StatusBar1.Panels(1).Text
F_kl.Recordset.Fields("xdate") = Taqvim.KKK.Caption
F_kl.Recordset.Fields("end_event") = Combo3.Text
F_kl.Recordset.Fields("gozaresh") = Text8.Text
F_kl.Recordset.Fields("tedad") = Text6.Text
F_kl.Recordset.Fields("vahed") = Combo2.Text
F_kl.Recordset.Update
F_kl.Refresh
a = Time_color_natije(Natije_color_time, 20, &HC000&, "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ")
Command1.SetFocus

End Sub

Private Sub Command3_Click()
On Error Resume Next

question_.Visible = True
yes_.Visible = True
no_.Visible = True
End Sub

Private Sub Command6_Click()
On Error Resume Next

edite_kl.Text9.Text = Me.F_kl.Recordset.Fields("id_kl")
edite_kl.Check1.Value = 1

edite_kl.Show

End Sub

Private Sub Form_Load()
On Error Resume Next

SB1.Panels(2).Text = Taqvim.KKK.Caption
Addsaa = &H80FF&

a = KOD_INS_COMBO_F_STTG("te_me", Combo2, 0)
a = KOD_INS_COMBO_F_STTG("s_m", Combo3, 0)
Combo3.Text = Combo3.List(0)
Combo2.Text = Combo2.List(0)
Set T = New Fx

a = T.INSERT_F_GH(Combo1, 1)
End Sub



Private Sub Label18_Click()

End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label6_Change()
On Error GoTo 1
GoTo 2
1:
Image1.Picture = LoadPicture(App.Path & "\pic\no_img.jpg")
Exit Sub
2:
Image1.Picture = LoadPicture(App.Path & "\pic\f_kl\" & Label6.Caption & ".jpg")
End Sub

Private Sub mnufindInhavalie_Click()
On Error Resume Next

Find_havale.id_kl_text.Text = Me.F_kl.Recordset.Fields("id_kl")
Find_havale.Show

End Sub

Private Sub no__Click()
On Error Resume Next

question_.Visible = False
yes_.Visible = False
no_.Visible = False

End Sub

Private Sub Text1_Change()
a = TEXT_CH_JOS_IN_F_KL("code", Text1.Text)

End Sub

Private Sub Text2_Change()
a = TEXT_CH_JOS_IN_F_KL("code", Text2.Text)

End Sub

Private Sub Text3_Change()
a = TEXT_CH_JOS_IN_F_KL("xnoe", Text3.Text)

End Sub

Private Sub Text4_Change()
a = TEXT_CH_JOS_IN_F_KL("tmyz", Text4.Text)

End Sub

Private Sub Text5_Change()
a = TEXT_CH_JOS_IN_F_KL("tozih", Text5.Text)

End Sub

Private Sub Text6_Change()
a = TEXT_CH_JOS_IN_F_KL("tedad", Text6.Text)

End Sub

Private Sub Text7_Change()
a = TEXT_CH_JOS_IN_F_KL("qeymat", Text7.Text)

End Sub

Private Sub Text9_Change()
On Error Resume Next

Dim T As String
T = Text9.Text

F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where id_kl like ('" & T & "') or code like ('%" & T & "%') or xnoe like ('%" & T & "%') or tmyz like ('%" & T & "%') or tozih like ('%" & T & "%') or qeymat like ('%" & T & "%') or xdate like ('%" & T & "%') or end_event like ('%" & T & "%') or tedad like ('%" & T & "%') or vahed like ('%" & T & "%')"
F_kl.Refresh
tedad_yaft.Caption = F_kl.Recordset.RecordCount

End Sub

Private Sub Timer1_Timer()
On Error Resume Next

SB1.Panels(1).Text = Format(Now, "short time")

End Sub

Private Sub Timer2_Timer()
On Error Resume Next

Timer_for_natije_color_ = Timer_for_natije_color_ - 1
If Timer_for_natije_color_ = 0 Then
Natije_color_time.Visible = False
Timer2.Enabled = False
End If



End Sub

Private Sub yes__Click()
On Error Resume Next

If F_kl.Recordset.RecordCount = 0 Then
a = Time_color_natije(Natije_color_time, 20, &H80FF&, "„Ê—œÌ »—«Ì Õ–› ÊÃÊœ ‰œ«—œ")
GoTo 1

End If

F_kl.Recordset.Delete
a = Time_color_natije(Natije_color_time, 20, &HC0&, "«ÿ·«⁄«  Õ–› ‘œ")
1:

question_.Visible = False
yes_.Visible = False
no_.Visible = False
End Sub
