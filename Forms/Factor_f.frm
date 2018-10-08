VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Factor_f 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "œ—ŒÊ«”  ÕÊ«·Â «“ «‰»«—"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Factor_f.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "..."
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2520
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ç«Å ÕÊ«·Â "
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
      TabIndex        =   8
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox tozih_t 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   8175
   End
   Begin VB.TextBox bargasht_ 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "«‰ Œ«» «ﬁ·«„"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
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
      Height          =   4410
      ItemData        =   "Factor_f.frx":030A
      Left            =   240
      List            =   "Factor_f.frx":030C
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   12855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "À»  "
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÃœÌœ"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
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
   Begin VB.Label time_ 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   " «—ÌŒ Ê ”«⁄ "
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
      TabIndex        =   20
      Top             =   240
      Width           =   1575
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
      Left            =   8520
      TabIndex        =   19
      Top             =   1440
      Width           =   720
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
      Left            =   11880
      TabIndex        =   18
      Top             =   1440
      Width           =   1110
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
      Left            =   4080
      TabIndex        =   17
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin VB.Label tahvil_dahande_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ò«—»— Ã«—Ì"
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
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label tahvil_girande_label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ÂÌç ›—œÌ «‰ Œ«» ‰‘œÂ «” "
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
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label date_ 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   " «—ÌŒ Ê ”«⁄ "
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   " «—ÌŒ Ê ”«⁄ "
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
      Left            =   4080
      TabIndex        =   12
      Top             =   240
      Width           =   1080
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
      Left            =   13200
      TabIndex        =   11
      Top             =   240
      Width           =   1035
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
      Left            =   13200
      TabIndex        =   10
      Top             =   840
      Width           =   1065
   End
   Begin VB.Menu mnufp 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu mnuftath 
         Caption         =   " ÕÊÌ· ÕÊ«·Â"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "Factor_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim last_id_inserter_f_fctR
Private Sub Command1_Click()
On Error Resume Next

List1.Clear
bargasht_.Text = ""
tozih_t.Text = ""
tahvil_girande_label.Caption = "ÂÌç ›—œÌ «‰ Œ«» ‰‘œÂ «” "

End Sub

Private Sub Command2_Click()
On Error Resume Next

If tahvil_girande_label.Caption = "ÂÌç ›—œÌ «‰ Œ«» ‰‘œÂ «” " Or tahvil_girande_label.Caption = "" Then Exit Sub
If List1.ListCount = 0 Then Exit Sub

tahvil_dahande_ = Split(tahvil_dahande_l.Caption, " _ ")
tahvil_girande_ = Split(tahvil_girande_label.Caption, " _ ")

F_fctr.Refresh
F_fctr.RecordSource = "select * from f_fctr where xdate like ('" & date_.Caption & "') and id_msht like ('" & tahvil_girande_(0) & "') and id_usr like ('" & tahvil_dahande_(0) & "')"
F_fctr.Refresh
If F_fctr.Recordset.BOF = False Or F_fctr.Recordset.EOF = False Then
If MsgBox("ç‰œ „Ê—œ »« „‘Œ’«  ‰“œÌò »Â «Ì‰ ÕÊ«·Â Ì«›  ‘œ ¬Ì« „ÿ„∆‰ Â” Ìœ òÂ «Ì‰ ÕÊ«·Â  ò—«—Ì ‰Ì” ø", vbQuestion + vbYesNo, "Â‘œ«—") = vbYes Then

Else
Exit Sub
End If
End If


ProgressBar1.Visible = True
ProgressBar1.Max = 2 + Val(Me.List1.ListCount)
ProgressBar1.Value = 1


F_fctr.Refresh
F_fctr.Recordset.AddNew
F_fctr.Recordset.Fields("xdate") = date_.Caption
F_fctr.Recordset.Fields("xtime") = time_.Caption
F_fctr.Recordset.Fields("id_msht") = tahvil_girande_(0)
F_fctr.Recordset.Fields("id_usr") = tahvil_dahande_(0)
F_fctr.Recordset.Fields("xend_date") = bargasht_.Text
F_fctr.Recordset.Fields("noe") = Combo1.Text
F_fctr.Recordset.Fields("chap") = "ç«Å ‰‘œÂ"
F_fctr.Recordset.Fields("vazeyat") = "out"

F_fctr.Recordset.Fields("tozih") = tozih_t.Text
F_fctr.Recordset.Update
last_id_inserter_f_fctR = F_fctr.Recordset.Fields("id_fctr")
F_fctr.Refresh
ProgressBar1.Value = 2

For I = 1 To Me.List1.ListCount

F_fctr_msht.Refresh
F_fctr_msht.Recordset.AddNew
F_fctr_msht.Recordset.Fields("id_fctr") = last_id_inserter_f_fctR
a = Split(Me.List1.List(I - 1), " _ ")
F_fctr_msht.Recordset.Fields("id_kl") = a(0)
F_kl.Refresh
F_kl.RecordSource = "select * from f_kl where id_kl like ('" & a(0) & "')"
F_kl.Refresh
F_fctr_msht.Recordset.Fields("v_tahvil") = F_kl.Recordset.Fields("end_event")
F_fctr_msht.Recordset.Fields("v_daryaft") = " ÕÊÌ· œ«œÂ ‰‘œÂ"
F_fctr_msht.Recordset.Fields("gozaresh") = "«ÿ·«⁄« Ì œ— œ” —” ‰Ì” "
F_fctr_msht.Recordset.Fields("vazeyat") = "out"
a = Split(Me.List1.List(I - 1), "    ⁄œ«œ   ")
F_fctr_msht.Recordset.Fields("tedad") = Val(a(1))

F_fctr_msht.Recordset.Update
F_fctr_msht.Refresh
ProgressBar1.Value = ProgressBar1.Value + 1
Next I
ProgressBar1.Value = 0
ProgressBar1.Visible = False
Call Command1_Click

'MsgBox "ÕÊ«·Â »« „Ê›ﬁÌ  À»  ‘œ", vbInformation, "À»  ÕÊ«·Â"
If MsgBox("ÕÊ«·Â »« „Ê›ﬁÌ  À»  ‘œ. ¬Ì« ‰„Ê‰Â ﬁ«»· ç«Å «“ «Ì‰ ÕÊ«·Â —« „Ì ŒÊ«ÂÌœø", vbQuestion + vbYesNo, "ç«Å ÕÊ«·Â") = vbYes Then
a = Print_factore(last_id_inserter_f_fctR)
End If
End Sub
Function Print_factore(id_factore)
Set T = New Fx
a = T.XLSX_Havale(id_factore)
End Function
Private Sub Command4_Click()
select_kl.Show

End Sub

Private Sub Command5_Click()
Call Label6_Click

End Sub

Private Sub Command7_Click()
If last_id_inserter_f_fctR = 0 Then Exit Sub
a = Print_factore(last_id_inserter_f_fctR)
End Sub

Private Sub Form_Load()
Set KOD_ = New Fx
'KOD_.KOD_INS_COMBO_F_STTG ("SD")
a = KOD_.KOD_INS_COMBO_F_STTG("noe_havale", Combo1, 0)
tahvil_dahande_l.Caption = menu.StatusBar1.Panels(2).Text & " _ " & menu.StatusBar1.Panels(1).Text
date_.Caption = Taqvim.KKK.Caption
Combo1.Text = Combo1.List(0)

last_id_inserter_f_fctR = 0

End Sub

Private Sub Label6_Click()
select_msht.Show

End Sub

Private Sub mnuftath_Click()
tahvil_kl_forn.Show

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
