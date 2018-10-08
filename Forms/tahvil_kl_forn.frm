VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tahvil_kl_forn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ· ÕÊ«·Â "
   ClientHeight    =   8655
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tahvil_kl_forn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   13170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   12960
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   4410
      ItemData        =   "tahvil_kl_forn.frx":030A
      Left            =   120
      List            =   "tahvil_kl_forn.frx":030C
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   12855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   7080
      Width           =   4815
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   8640
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "‰„«Ì‘"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«»ÿ«· ÕÊ«·Â"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   2775
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   4245
      ItemData        =   "tahvil_kl_forn.frx":030E
      Left            =   120
      List            =   "tahvil_kl_forn.frx":0310
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   2520
      Width           =   12855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   120
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ÂÌç òœ«„"
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
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label14 
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
      Left            =   7920
      TabIndex        =   31
      Top             =   7320
      Width           =   585
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ÂÌç òœ«„"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "„⁄òÊ”"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "«‰ Œ«» Â„Â"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "‘„«—Â ÕÊ«·Â"
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
      TabIndex        =   27
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label f_ctr_id_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "Ê÷⁄Ì  ò«·«Â«Ì «‰ Œ«» ‘œÂ"
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
      Left            =   10800
      TabIndex        =   25
      Top             =   7320
      Width           =   2220
   End
   Begin VB.Label chap_la 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "Ê÷⁄Ì  ç«Å"
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
      Left            =   12000
      TabIndex        =   23
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label bazgasht_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label tozih_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Width           =   8175
   End
   Begin VB.Label noe_ha 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "‘„«—Â ÕÊ«·Â"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   " ÕÊÌ· êÌ—‰œÂ"
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
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   " ÕÊÌ· œÂ‰œÂ"
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
      TabIndex        =   17
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   " «—ÌŒ  ÕÊÌ·"
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
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label date_ 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label tahvil_girande_label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label tahvil_dahande_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "‰Ê⁄ ÕÊ«·Â"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   " «—ÌŒ »«“ê‘ "
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
      TabIndex        =   11
      Top             =   1920
      Width           =   1110
   End
   Begin VB.Label Label10 
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
      Left            =   8400
      TabIndex        =   10
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label time_ 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu Mnuparvande 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu mnushowAll 
         Caption         =   "‰„«Ì‘ ò«„· ÕÊ«·Â"
      End
      Begin VB.Menu mnufarkh 
         Caption         =   "œ—ŒÊ«”  ÕÊ«·Â"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnufing 
      Caption         =   "Ã” ÃÊ œ— ÕÊ«·Â »« ‰«„ «‘Œ«’"
   End
   Begin VB.Menu mnufkl 
      Caption         =   "Ã” ÃÊ œ— ÕÊ«·Â »« ‰Ê⁄ ò«·«"
   End
End
Attribute VB_Name = "tahvil_kl_forn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

If Text9.Text = "" Then Exit Sub
List2.Visible = False
List1.Visible = True

F_fctr.Refresh
F_fctr.RecordSource = "select * from f_fctr where id_fctr like ('" & Text9.Text & "') " 'and vazeyat like ('" & "out" & "')"
F_fctr.Refresh
If F_fctr.Recordset.RecordCount = 0 Then
f_ctr_id_l.Caption = "‘„«—Â ÕÊ«·Â Ì«›  ‰‘œ"
Exit Sub
End If

f_ctr_id_l.Caption = F_fctr.Recordset.Fields("id_fctr")
Label15.Caption = F_fctr.Recordset.Fields("vazeyat")
date_.Caption = F_fctr.Recordset.Fields("xdate")
time_.Caption = F_fctr.Recordset.Fields("xtime")
F_msht.Refresh
F_msht.RecordSource = "select * from f_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_msht") & "')"
F_msht.Refresh
tahvil_girande_label.Caption = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
F_usr.Refresh
F_usr.RecordSource = "select * from f_usr where id_usr like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_usr.Refresh
F_msht.Refresh
F_msht.RecordSource = "select * from f_msht where id_msht like ('" & F_usr.Recordset.Fields("id_msht") & "')"
F_msht.Refresh
tahvil_dahande_l.Caption = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")


 bazgasht_l.Caption = F_fctr.Recordset.Fields("xend_date")
noe_ha.Caption = F_fctr.Recordset.Fields("noe")

chap_la.Caption = F_fctr.Recordset.Fields("chap")
tozih_l.Caption = F_fctr.Recordset.Fields("tozih")
F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "select * from f_fctr_msht where id_fctr like ('" & F_fctr.Recordset.Fields("id_fctr") & "') and vazeyat like ('" & "out" & "')"
F_fctr_msht.Refresh
List1.Clear

For I = 1 To F_fctr_msht.Recordset.RecordCount
F_kl.Refresh
F_kl.RecordSource = "select * from F_kl where id_kl like ('" & F_fctr_msht.Recordset.Fields("id_kl") & "')"
F_kl.Refresh
List1.AddItem ("  òœ  " & F_kl.Recordset.Fields("id_kl") & " _ " & F_kl.Recordset.Fields("xnoe") & "   |   " & F_kl.Recordset.Fields("tmyz") & "   ||   " & F_kl.Recordset.Fields("tozih") & "    ⁄œ«œ   " & F_fctr_msht.Recordset.Fields("tedad") & " :: ")

F_fctr_msht.Recordset.MoveNext
Next I

Beep

End Sub



Private Sub Command2_Click()
On Error Resume Next

Dim u(0 To 1) As String
Dim i_CoUnt As Integer
i_CoUnt = 0
For I = 1 To List1.ListCount
If List1.Selected(I - 1) = False Then GoTo 3
i_CoUnt = i_CoUnt + 1
a = Split(List1.List(I - 1), " _ ")
T = Split(a(0), " òœ ")
F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "select * from f_fctr_msht where id_fctr like ('" & f_ctr_id_l.Caption & "') and id_kl like ('" & Val(T(1)) & "')"
F_fctr_msht.Refresh
g = Split(List1.List(I - 1), " :: ")
If g(1) = "" Then GoTo 2

g = Split(List1.List(I - 1), "   Ê÷⁄Ì :   ")
gozaresh_ = Split(g(1), "   :ê“«—‘   ")
F_fctr_msht.Recordset.Fields("v_daryaft") = gozaresh_(0)

g = Split(List1.List(I - 1), "   :ê“«—‘   ")
gozaresh_ = Split(g(1), "    ⁄œ«œ »«“ê‘ Ì:   ")

F_fctr_msht.Recordset.Fields("gozaresh") = gozaresh_(0)

g = Split(List1.List(I - 1), "    ⁄œ«œ »«“ê‘ Ì:   ")
F_fctr_msht.Recordset.Fields("tedad_bazgasht") = g(1)


'tahvil_kl_forn.List1.List(a(0)) = tahvil_kl_forn.List1.List(a(0)) & " :: " & "   Ê÷⁄Ì :   " & Me.Combo2.Text & "   :ê“«—‘   " & Text9.Text & "    ⁄œ«œ »«“ê‘ Ì:   " & Text1.Text


'gozaresh nadare
GoTo 1
2:
F_fctr_msht.Recordset.Fields("v_daryaft") = Combo2.Text
F_fctr_msht.Recordset.Fields("gozaresh") = Text1.Text
F_fctr_msht.Recordset.Fields("tedad_bazgasht") = F_fctr_msht.Recordset.Fields("tedad")



1:
F_fctr_msht.Recordset.Fields("vazeyat") = "in"
F_fctr_msht.Recordset.Fields("date_bazgasht") = Taqvim.KKK.Caption

F_fctr_msht.Recordset.Update




3: Next I
If i_CoUnt = List1.ListCount Then
F_fctr.Refresh
F_fctr.RecordSource = "select * from F_fctr where id_fctr like ('" & f_ctr_id_l.Caption & "')" ' and id_kl like ('" & Val(T(1)) & "')"
F_fctr.Refresh
F_fctr.Recordset.Fields("vazeyat") = "in"
F_fctr.Recordset.Update
MsgBox "ÕÊ«·Â »Â ’Ê—  ò«„· œ—Ì«›  Ê »«ÿ· ‘œ", vbInformation, "«»ÿ«· ÕÊ«·Â"

End If

Call Command1_Click

End Sub

Private Sub Form_Load()

Set T = New Fx
a = T.KOD_INS_COMBO_F_STTG("s_m", Combo2, 0)
Combo2.Text = Combo2.List(0)
date_.Caption = Taqvim.KKK.Caption

End Sub

Private Sub Label11_Click()
On Error Resume Next

For I = 1 To List1.ListCount
List1.Selected(I - 1) = True
Next I

End Sub

Private Sub Label12_Click()
On Error Resume Next

For I = 1 To List1.ListCount
List1.Selected(I - 1) = False
Next I

End Sub

Private Sub Label15_Change()
On Error Resume Next

If Label15.Caption = "out" Then
Label15.Caption = "ÕÊ«·Â ›⁄«·"
ElseIf Label15.Caption = "in" Then
Label15.Caption = "ÕÊ«·Â »«ÿ· ‘œÂ"
End If
End Sub

Private Sub Label6_Click()
On Error Resume Next

For I = 1 To List1.ListCount
 If List1.Selected(I - 1) = True Then
 List1.Selected(I - 1) = False
 Else
 List1.Selected(I - 1) = True
 End If
 
Next I

End Sub

Private Sub List1_DblClick()
On Error Resume Next

gozaresh_tahvil_form.list_index_l.Caption = Me.List1.ListIndex & " _ " & Me.List1.Text
a = Split(Me.List1.Text, "  ⁄œ«œ ")
gozaresh_tahvil_form.tedad_l.Caption = Val(a(1))
g = Split(List1.Text, " :: ")
If g(1) = "" Then GoTo 2

g = Split(List1.Text, "   Ê÷⁄Ì :   ")
gozaresh_ = Split(g(1), "   :ê“«—‘   ")
gozaresh_tahvil_form.Combo2.Text = gozaresh_(0)

g = Split(List1.Text, "   :ê“«—‘   ")
gozaresh_ = Split(g(1), "    ⁄œ«œ »«“ê‘ Ì:   ")

gozaresh_tahvil_form.Text9.Text = gozaresh_(0)

g = Split(List1.Text, "    ⁄œ«œ »«“ê‘ Ì:   ")
gozaresh_tahvil_form.Text1.Text = g(1)
2

gozaresh_tahvil_form.Show


End Sub

Private Sub List2_DblClick()
On Error Resume Next

show_more_fctr_kl.Show
show_more_fctr_kl.Text1.Text = Me.List2.Text
End Sub

Private Sub mnufarkh_Click()
Factor_f.Show

End Sub

Private Sub mnufing_Click()
select_msht.Show

End Sub

Private Sub mnufkl_Click()
select_kl.Show

End Sub

Private Sub mnushowall_Click()
On Error Resume Next

If Text9.Text = "" Then Exit Sub
List1.Clear

List2.Visible = True
List1.Visible = False

F_fctr.Refresh
F_fctr.RecordSource = "select * from f_fctr where id_fctr like ('" & Text9.Text & "')" ' and vazeyat like ('" & "out" & "')"
F_fctr.Refresh
If F_fctr.Recordset.RecordCount = 0 Then
f_ctr_id_l.Caption = "‘„«—Â ÕÊ«·Â Ì«›  ‰‘œ"
Exit Sub
End If

f_ctr_id_l.Caption = F_fctr.Recordset.Fields("id_fctr")
Label15.Caption = F_fctr.Recordset.Fields("vazeyat")
date_.Caption = F_fctr.Recordset.Fields("xdate")
time_.Caption = F_fctr.Recordset.Fields("xtime")
F_msht.Refresh
F_msht.RecordSource = "select * from f_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_msht") & "')"
F_msht.Refresh
tahvil_girande_label.Caption = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
F_usr.Refresh
F_usr.RecordSource = "select * from f_usr where id_usr like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_usr.Refresh
F_msht.Refresh
F_msht.RecordSource = "select * from f_msht where id_msht like ('" & F_usr.Recordset.Fields("id_msht") & "')"
F_msht.Refresh
tahvil_dahande_l.Caption = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")


 bazgasht_l.Caption = F_fctr.Recordset.Fields("xend_date")
noe_ha.Caption = F_fctr.Recordset.Fields("noe")

chap_la.Caption = F_fctr.Recordset.Fields("chap")
tozih_l.Caption = F_fctr.Recordset.Fields("tozih")
F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "select * from f_fctr_msht where id_fctr like ('" & F_fctr.Recordset.Fields("id_fctr") & "')" ' and vazeyat like ('" & "out" & "')"
F_fctr_msht.Refresh
List2.Clear
Dim Te_ As String
Te_ = ""

For I = 1 To F_fctr_msht.Recordset.RecordCount
Te_ = ""
F_kl.Refresh
F_kl.RecordSource = "select * from F_kl where id_kl like ('" & F_fctr_msht.Recordset.Fields("id_kl") & "')"
F_kl.Refresh
Te_ = Te_ & " Ê÷⁄Ì   ÕÊÌ· " & F_fctr_msht.Recordset.Fields("v_tahvil")
Te_ = Te_ & " Ê÷⁄Ì  œ—Ì«›  " & F_fctr_msht.Recordset.Fields("v_daryaft")
Te_ = Te_ & " ê“«—‘ " & F_fctr_msht.Recordset.Fields("gozaresh")
Te_ = Te_ & "  ⁄œ«œ »«“ê‘ Ì " & F_fctr_msht.Recordset.Fields("tedad_bazgasht")
Te_ = Te_ & "  «—ÌŒ »«“ê‘  " & F_fctr_msht.Recordset.Fields("date_bazgasht")

List2.AddItem ("  òœ  " & F_kl.Recordset.Fields("id_kl") & " _ " & F_kl.Recordset.Fields("xnoe") & "   |   " & F_kl.Recordset.Fields("tmyz") & "   ||   " & F_kl.Recordset.Fields("tozih") & "    ⁄œ«œ  ÕÊÌ· œ«œÂ ‘œÂ   " & F_fctr_msht.Recordset.Fields("tedad") & " :: " & Te_)

F_fctr_msht.Recordset.MoveNext
Next I

Beep
End Sub

Private Sub Text2_Change()
Call Command1_Click

End Sub

Private Sub Timer1_Timer()
Dim HH, MM, SS As String
HH = Hour(Now)
MM = Minute(Now)
SS = Second(Now)
If Val(HH) < 10 Then HH = "0" & HH
If Val(MM) < 10 Then MM = "0" & MM
If Val(SS) < 10 Then SS = "0" & SS


time_.Caption = HH & ":" & MM & ":" & SS
End Sub
