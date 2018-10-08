VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form select_kl 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇäÊÎÇÈ ÇÞáÇã ÍæÇáå"
   ClientHeight    =   8625
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15945
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "select_kl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   15945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÇÖÇÝå ˜ÑÏä Èå áíÓÊ"
      Default         =   -1  'True
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "ãÌÏÏ"
      Height          =   465
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ËÈÊ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   4335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808000&
      CausesValidation=   0   'False
      ForeColor       =   &H00FFFFC0&
      Height          =   2475
      ItemData        =   "select_kl.frx":030A
      Left            =   4680
      List            =   "select_kl.frx":030C
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5760
      Width           =   11175
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "ÊÕæíÑ ˜ÇáÇ"
      ForeColor       =   &H0000FFFF&
      Height          =   3735
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.Image Image1 
         Height          =   3255
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   4095
      End
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
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "select_kl.frx":030E
      Height          =   4815
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
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
      Caption         =   "áíÓÊ ˜ÇáÇåÇí ËÈÊ ÔÏå"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ID_kl"
         Caption         =   "˜Ï"
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
         Caption         =   "˜Ï ˜ÇáÇ"
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
         Caption         =   "äæÚ ˜ÇáÇ"
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
         Caption         =   "ãÔÎÕå"
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
         Caption         =   "ÊæÖíÍÇÊ"
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
         Caption         =   "ÞíãÊ"
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
         Caption         =   "˜ÇÑÈÑ"
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
         Caption         =   "ÊÇÑíÎ ËÈÊ"
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
         Caption         =   "ÂÎÑíä æÖÚíÊ"
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
         Caption         =   "ÊÚÏÇÏ"
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
         Caption         =   "æÇÍÏ"
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
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2775.118
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "ÍÐÝ"
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ÊÚÏÇÏ ÏÑÎæÇÓÊí"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ÚÏÏ"
      DataField       =   "vahed"
      DataSource      =   "F_kl"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Ç˜ÓÇÒí áíÓÊ"
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   14400
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      DataField       =   "gozaresh"
      DataSource      =   "F_kl"
      ForeColor       =   &H00404000&
      Height          =   1335
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      DataField       =   "end_event"
      DataSource      =   "F_kl"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÒÇÑÔ"
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
      Left            =   3840
      TabIndex        =   18
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "æÖÚíÊ ˜ÇáÇ"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   5880
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ÚÏÏ"
      DataField       =   "vahed"
      DataSource      =   "F_kl"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label tahvil_girande_label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ãæÌæÏí ÝÚáí"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label tahvil_dahande_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ãæÌæÏí ˜áí"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "ÊÛííÑ ÏÑ ãæÌæÏí"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÌÓÊÌæ"
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
      Left            =   14760
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Menu mnuadd_kl 
      Caption         =   "ÇÖÇÝå ˜ÑÏä ˜ÇáÇí ÌÏíÏ"
   End
   Begin VB.Menu mnufindP 
      Caption         =   "ÌÓÊÌæí Çíä ˜ÇáÇ ÏÑ ÍæÇáå åÇ"
   End
End
Attribute VB_Name = "select_kl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
On Error Resume Next

b = Find_id_f_sttg(Combo1)
Text5.Enabled = F_sttg.Recordset.Fields("x1")
Text5.BackColor = (F_sttg.Recordset.Fields("x2"))
If F_sttg.Recordset.Fields("xtext") = "No" Then Text5.Text = ""

End Sub
Function Find_id_f_sttg(Combo_)
On Error Resume Next

a = Split(Combo_.Text, " _ ")
F_sttg.Refresh
F_sttg.RecordSource = "select * from f_sttg where id_sttg like ('" & a(0) & "')"
F_sttg.Refresh
End Function

Private Sub Command1_Click()
Call Text9_Change

End Sub

Private Sub Command2_Click()
On Error Resume Next

For I = 1 To Me.List1.ListCount
Factor_f.List1.AddItem (Me.List1.List(I - 1))
Next I
Beep

Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Val(Text1.Text) > Val(Label5.Caption) Then Exit Sub
If Val(Text1.Text) = 0 Then Exit Sub
If F_kl.Recordset.RecordCount <> 0 Then
For I = 1 To List1.ListCount
a = Split(List1.List(I - 1), " _ ")
If Val(a(0)) = F_kl.Recordset.Fields("id_kl") Then
MsgBox "Çíä ˜ÇáÇ ÏÑ áíÓÊ æÌæÏ ÏÇÑÏ", vbInformation, ".:.:."
Exit Sub

Else
End If
Next I

List1.AddItem (F_kl.Recordset.Fields("id_kl") & " _ " & F_kl.Recordset.Fields("xnoe") & "   |   " & F_kl.Recordset.Fields("tmyz") & "   ||   " & F_kl.Recordset.Fields("tozih") & "   ÊÚÏÇÏ   " & Text1.Text)
Text9.SetFocus
Text1.Text = 1

End If

End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next

edite_kl.Text9.Text = Me.F_kl.Recordset.Fields("id_kl")

edite_kl.Show
End Sub

Private Sub Form_Load()
'Set KOD_ = New Fx

'a = KOD_.KOD_INS_COMBO_F_STTG("s_m", Combo1, 1)
End Sub

Private Sub Label11_Click()
List1.Clear

End Sub

Private Sub Label13_Click()
On Error Resume Next


List1.RemoveItem (List1.ListIndex)


End Sub

Private Sub Label2_Click()
change_mojodi.Text9.Text = Me.F_kl.Recordset.Fields("id_kl")
change_mojodi.Show

End Sub

Private Sub List1_Click()
On Error Resume Next

a = Split(List1.Text, " _ ")
F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where id_kl like ('" & a(0) & "')"
F_kl.Refresh

End Sub

Private Sub mnuadd_kl_Click()
add_kl.Show

End Sub

Private Sub mnufindP_Click()
On Error Resume Next

Find_havale.id_kl_text.Text = Me.F_kl.Recordset.Fields("id_kl")
Find_havale.Show
End Sub

Private Sub Text1_Change()
On Error Resume Next

If Val(Text1.Text) > Val(Label5.Caption) Then
Beep
Text1.Text = Label5.Caption
End If


End Sub

Private Sub Text9_Change()
On Error Resume Next

Dim T As String
T = Text9.Text

F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where id_kl like ('" & T & "') or code like ('%" & T & "%') or xnoe like ('%" & T & "%') or tmyz like ('%" & T & "%') or tozih like ('%" & T & "%') or qeymat like ('%" & T & "%') or xdate like ('%" & T & "%') or end_event like ('%" & T & "%') or tedad like ('%" & T & "%') or vahed like ('%" & T & "%')"
F_kl.Refresh
'tedad_yaft.Caption = F_kl.Recordset.RecordCount

End Sub

Private Sub Label6_Change()
On Error Resume Next

Label3.Caption = Me.F_kl.Recordset.Fields("tedad")
F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "select * from F_fctr_msht where id_kl like ('" & Label6.Caption & "') and vazeyat like ('out')"
F_fctr_msht.Refresh
ted = 0
For I = 1 To F_fctr_msht.Recordset.RecordCount
ted = ted + Val(F_fctr_msht.Recordset.Fields("tedad"))
Next I
Label5.Caption = Val(Label3.Caption) - Val(ted)
On Error GoTo 1
GoTo 2
1:
Image1.Picture = LoadPicture(App.Path & "\pic\no_img.jpg")
Exit Sub
2:
Image1.Picture = LoadPicture(App.Path & "\pic\f_kl\" & Label6.Caption & ".jpg")
End Sub
