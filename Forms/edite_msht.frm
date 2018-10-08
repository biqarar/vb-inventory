VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form edite_msht 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’·«Õ „‘Œ’«  „‘ —Ì"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "edite_msht.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   " »   €ÌÌ—« "
      ForeColor       =   &H0000FFFF&
      Height          =   4215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame1 
         BackColor       =   &H00008080&
         Caption         =   "ADODC"
         Height          =   255
         Left            =   0
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
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   12600
         TabIndex        =   24
         Top             =   4800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox tel1_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7560
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox pedar_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox name_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   7560
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
         Caption         =   "À»  "
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
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   360
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox kod_meli_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox famil_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox tavalod_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   7560
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox adress_t 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   9495
      End
      Begin VB.TextBox mob_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox tel2_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox tozih_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2880
         Width           =   9495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " ·›‰ À«» "
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
         Left            =   10080
         TabIndex        =   23
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "‰«„ Åœ—"
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
         Left            =   2160
         TabIndex        =   22
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "òœ „·Ì"
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
         Left            =   6600
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "”„ "
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
         TabIndex        =   20
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "‰«„"
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
         Left            =   10200
         TabIndex        =   19
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
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
         Left            =   6360
         TabIndex        =   18
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " «—ÌŒ  Ê·œ"
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
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   "¬œ—”"
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
         Left            =   10200
         TabIndex        =   16
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " ·›‰ Â„—«Â"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         Caption         =   " ·›‰ «÷ÿ—«—Ì"
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
         Left            =   2760
         TabIndex        =   14
         Top             =   1800
         Width           =   1155
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
         Left            =   10080
         TabIndex        =   13
         Top             =   3000
         Width           =   720
      End
   End
End
Attribute VB_Name = "edite_msht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error Resume Next

If name_t.Text = "" Or famil_t.Text = "" Then Exit Sub
F_msht.Recordset.Fields("xname") = name_t.Text
F_msht.Recordset.Fields("famil") = famil_t.Text
F_msht.Recordset.Fields("pedar") = pedar_t.Text
F_msht.Recordset.Fields("tavalod") = tavalod_t.Text
F_msht.Recordset.Fields("kodmeli") = kod_meli_t.Text
F_msht.Recordset.Fields("tel1") = tel1_t.Text
F_msht.Recordset.Fields("tel2") = tel2_t.Text
F_msht.Recordset.Fields("mob") = mob_t.Text
F_msht.Recordset.Fields("adress") = adress_t.Text
F_msht.Recordset.Fields("semat") = Combo1.Text
F_msht.Recordset.Fields("tozih") = tozih_t.Text

F_msht.Recordset.Update
F_msht.Refresh
MsgBox " €ÌÌ—«  «⁄„«· ‘œ", vbInformation, "«’·«Õ «ÿ·«⁄« "
add_msht.Show
add_msht.F_msht.Refresh
select_msht.F_msht.Refresh
select_msht.DataGrid1.Refresh

add_msht.DataGrid1.Refresh
Unload Me

End Sub

Private Sub Form_Load()
Set T = New Fx
a = T.KOD_INS_COMBO_F_STTG("semat", Combo1, 0)
End Sub

Private Sub Text1_Change()
On Error Resume Next

F_msht.Refresh
F_msht.RecordSource = "select * from f_msht where id_msht like ('" & Text1.Text & "')"
F_msht.Refresh
name_t.Text = F_msht.Recordset.Fields("xname")
famil_t.Text = F_msht.Recordset.Fields("famil")
pedar_t.Text = F_msht.Recordset.Fields("pedar")
tavalod_t.Text = F_msht.Recordset.Fields("tavalod")
kod_meli_t.Text = F_msht.Recordset.Fields("kodmeli")
tel1_t.Text = F_msht.Recordset.Fields("tel1")
tel2_t.Text = F_msht.Recordset.Fields("tel2")
mob_t.Text = F_msht.Recordset.Fields("mob")
adress_t.Text = F_msht.Recordset.Fields("adress")
Combo1.Text = F_msht.Recordset.Fields("semat")
tozih_t.Text = F_msht.Recordset.Fields("tozih")

End Sub
