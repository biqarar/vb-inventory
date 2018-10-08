VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form edite_kl 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’·«Õ „‘Œ’«  ò«·«"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "edite_kl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404000&
      Caption         =   "«⁄„«·  €ÌÌ—« "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   26
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "«’·«Õ „‘Œ’«  ò«·«"
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   11175
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   9840
         TabIndex        =   24
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2880
         Width           =   9495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   6480
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   960
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
         Caption         =   " »   €ÌÌ—« "
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   8520
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   8520
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   465
         Left            =   6480
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   6495
      End
      Begin VB.ComboBox Combo3 
         Height          =   465
         Left            =   7800
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Natije_color_time 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   2760
         TabIndex        =   25
         Top             =   3840
         Visible         =   0   'False
         Width           =   3015
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
         Left            =   9840
         TabIndex        =   23
         Top             =   3120
         Width           =   720
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
         Left            =   9960
         TabIndex        =   22
         Top             =   1800
         Width           =   615
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
         Left            =   5640
         TabIndex        =   21
         Top             =   600
         Width           =   645
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
         Left            =   9960
         TabIndex        =   20
         Top             =   480
         Width           =   585
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
         Left            =   5640
         TabIndex        =   19
         Top             =   1080
         Width           =   375
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
         Left            =   3960
         TabIndex        =   18
         Top             =   1680
         Width           =   420
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
         Left            =   9960
         TabIndex        =   17
         Top             =   1200
         Width           =   1005
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
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   375
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
         Left            =   9960
         TabIndex        =   15
         Top             =   2400
         Width           =   555
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
         Left            =   6960
         TabIndex        =   14
         Top             =   2280
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   0
      TabIndex        =   12
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
End
Attribute VB_Name = "edite_kl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer_for_natije_color_  As Integer

Private Sub Check1_Click()
On Error Resume Next

If Check1.Value = 1 Then
Me.Frame3.Enabled = True
Else
Me.Frame3.Enabled = False
End If

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


F_kl.Recordset.Fields("code") = cod_1_
F_kl.Recordset.Fields("xnoe") = Text3.Text
F_kl.Recordset.Fields("tmyz") = Text4.Text
F_kl.Recordset.Fields("tozih") = Text5.Text
F_kl.Recordset.Fields("qeymat") = Text7.Text
c = Split(Combo1.Text, " _ ")
F_kl.Recordset.Fields("id_gh") = c(0)
'F_kl.Recordset.Fields("id_usr") = 1
'F_kl.Recordset.Fields("xdate") = Taqvim.KKK.Caption
F_kl.Recordset.Fields("end_event") = Combo3.Text
F_kl.Recordset.Fields("gozaresh") = Text8.Text
F_kl.Recordset.Fields("tedad") = Text6.Text
F_kl.Recordset.Fields("vahed") = Combo2.Text
F_kl.Recordset.Update
F_kl.Refresh
a = Time_color_natije(Natije_color_time, 20, &HC000&, " €ÌÌ—«  »« „Ê›ﬁÌ  À»  ‘œ")

Beep
add_kl.F_kl.Refresh
select_kl.F_kl.Refresh
select_kl.DataGrid1.Refresh

add_kl.DataGrid1.Refresh

Unload Me

End Sub

Private Sub Form_Load()
Set KOD_ = New Fx
'KOD_.KOD_INS_COMBO_F_STTG ("SD")
a = KOD_.KOD_INS_COMBO_F_STTG("te_me", Combo2, 0)
a = KOD_.KOD_INS_COMBO_F_STTG("s_m", Combo3, 0)


ay = KOD_.INSERT_F_GH(Combo1, 1)
End Sub

Function Time_color_natije(labe_, xtime_, color_, str_)
On Error Resume Next

labe_.Visible = True
Timer_for_natije_color_ = xtime_
Timer2.Enabled = True
labe_.BackColor = color_
labe_.Caption = str_

End Function

Private Sub Text9_Change()
On Error Resume Next

F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where id_kl like ('" & Text9.Text & "')"
F_kl.Refresh

a = Split(F_kl.Recordset.Fields("code"), " :: ")
Text1.Text = a(0)
Text2.Text = a(1)

MoToR.F_gh.Refresh
MoToR.F_gh.RecordSource = "select * from F_gh where id_gh like ('" & Me.F_kl.Recordset.Fields("id_gh") & "')  " ' kod like ('" & KOD_ & "')"
MoToR.F_gh.Refresh
Combo1.Text = (MoToR.F_gh.Recordset.Fields("id_gh") & " _ " & MoToR.F_gh.Recordset.Fields("xname") & "  _  " & MoToR.F_gh.Recordset.Fields("tozih"))
Text3.Text = F_kl.Recordset.Fields("xnoe")
Text4.Text = F_kl.Recordset.Fields("tmyz")
Text5.Text = F_kl.Recordset.Fields("tozih")
Text7.Text = F_kl.Recordset.Fields("qeymat")

'F_kl.Recordset.Fields("id_usr") = 1
'F_kl.Recordset.Fields("xdate") = Taqvim.KKK.Caption
'a = Split(F_kl.Recordset.Fields("end_event"), " :: ")
'Combo3.Text = a(0)
'Text8.Text = a(1)
Combo3.Text = F_kl.Recordset.Fields("end_event")
Text8.Text = F_kl.Recordset.Fields("gozaresh")
Text6.Text = F_kl.Recordset.Fields("tedad")
Combo2.Text = F_kl.Recordset.Fields("vahed")

End Sub

Private Sub Timer2_Timer()
On Error Resume Next

Timer_for_natije_color_ = Timer_for_natije_color_ - 1
If Timer_for_natije_color_ = 0 Then
Natije_color_time.Visible = False
Timer2.Enabled = False
End If


End Sub
