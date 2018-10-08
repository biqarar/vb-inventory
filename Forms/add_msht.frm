VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form add_msht 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "À»  „‘ —Ì"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16050
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_msht.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   16050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      DisabledPicture =   "add_msht.frx":030A
      DownPicture     =   "add_msht.frx":24F84
      DragIcon        =   "add_msht.frx":49BFE
      Height          =   330
      Left            =   15600
      Picture         =   "add_msht.frx":6E878
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "«‰ ﬁ«· ÃœÊ· »Â »—‰«„Â «ò”·"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Caption         =   "ADODC"
      Height          =   255
      Left            =   0
      TabIndex        =   32
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "À»  „‘ —Ì"
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   120
      Width           =   15735
      Begin VB.TextBox tozih_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   14295
      End
      Begin VB.TextBox tel2_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5040
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox mob_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8640
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox adress_t 
         Alignment       =   1  'Right Justify
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox tavalod_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   12240
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox famil_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   7560
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox kod_meli_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   7560
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   5040
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› "
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox name_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   12240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox pedar_t 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   1575
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
         Height          =   615
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CheckBox ch_Jostan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404000&
         Caption         =   "Ã” ÃÊ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   14760
         TabIndex        =   18
         Top             =   2880
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox tel1_t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   12240
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Left            =   9120
         TabIndex        =   15
         Top             =   2880
         Width           =   5415
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
         Left            =   14760
         TabIndex        =   31
         Top             =   2400
         Width           =   720
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
         Left            =   7440
         TabIndex        =   30
         Top             =   1800
         Width           =   1155
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
         Left            =   11160
         TabIndex        =   29
         Top             =   1800
         Width           =   840
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
         Left            =   4200
         TabIndex        =   28
         Top             =   360
         Width           =   540
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
         Left            =   14640
         TabIndex        =   27
         Top             =   1200
         Width           =   825
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
         Left            =   11040
         TabIndex        =   26
         Top             =   600
         Width           =   1080
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
         Left            =   14760
         TabIndex        =   25
         Top             =   480
         Width           =   240
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
         Left            =   6960
         TabIndex        =   24
         Top             =   1200
         Width           =   405
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
         Left            =   11280
         TabIndex        =   23
         Top             =   1080
         Width           =   615
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
         Left            =   6840
         TabIndex        =   22
         Top             =   480
         Width           =   570
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
         Left            =   6840
         TabIndex        =   21
         Top             =   3000
         Width           =   1140
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
         Left            =   8520
         TabIndex        =   20
         Top             =   3000
         Width           =   105
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
         Left            =   14760
         TabIndex        =   19
         Top             =   1800
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "add_msht.frx":934F2
      Height          =   4935
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   8705
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
      Caption         =   "·Ì”  „‘ —Ì«‰"
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "ID_msht"
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
         DataField       =   "xname"
         Caption         =   "‰«„"
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
         DataField       =   "famil"
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
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
         DataField       =   "pedar"
         Caption         =   "‰«„ Åœ—"
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
         DataField       =   "tavalod"
         Caption         =   " «—ÌŒ  Ê·œ"
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
         DataField       =   "kodmeli"
         Caption         =   "òœ „·Ì"
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
         DataField       =   "tel1"
         Caption         =   " ·›‰ À«» "
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
         DataField       =   "mob"
         Caption         =   "Â„—«Â"
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
         DataField       =   "tel2"
         Caption         =   " ·›‰ «÷ÿ—«—Ì"
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
         DataField       =   "adress"
         Caption         =   "¬œ—”"
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
         DataField       =   "semat"
         Caption         =   "”„ "
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
      BeginProperty Column11 
         DataField       =   "emza"
         Caption         =   "emza"
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
      BeginProperty Column12 
         DataField       =   "ax"
         Caption         =   "ax"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
         DataField       =   "id_usr"
         Caption         =   "id_usr"
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
      BeginProperty Column15 
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1920.189
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
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnupic 
      Caption         =   "‰„«Ì‘ ò«„· «ÿ·«⁄« "
   End
   Begin VB.Menu mnufind_havale 
      Caption         =   "‰„«Ì‘ ÕÊ«·Â Â«Ì ›⁄«· «Ì‰ „‘ —Ì"
   End
End
Attribute VB_Name = "add_msht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adress_t_Change()
On Error Resume Next

a = TEXT_CH_JOS_IN_F_KL("adress", adress_t.Text)

End Sub

Private Sub Combo1_Click()
On Error Resume Next

a = TEXT_CH_JOS_IN_F_KL("semat", Combo1.Text)

End Sub

Private Sub Command1_Click()
On Error Resume Next
Add_new_record



End Sub
Function Clear_text()
On Error Resume Next
name_t.Text = ""
famil_t.Text = ""
pedar_t.Text = ""
tavalod_t.Text = ""
kod_meli_t.Text = ""
tel1_t.Text = ""
tel2_t.Text = ""
mob_t.Text = ""
adress_t.Text = ""

tozih_t.Text = ""
name_t.SetFocus

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

Private Sub Command11_Click()
Dim oExcel As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Sub
2:
Me.F_msht.Recordset.MoveFirst

Dim BePAR As Integer

BePAR = 0
Set oExcel = GetObject(App.Path & "\formXLS\Emza_Pic.xls")


oExcel.ActiveSheet.Range("G3").Value = Taqvim.KKK.Caption

For I = 5 To F_msht.Recordset.RecordCount + 4
oExcel.ActiveSheet.Range("B" & I).Value = Me.F_msht.Recordset.Fields("id_msht")

text_ = ""
text_ = text_ & "  " & Me.F_msht.Recordset.Fields("xname")
text_ = text_ & "  " & Me.F_msht.Recordset.Fields("famil")
text_ = text_ & " ‰«„ Åœ— " & Me.F_msht.Recordset.Fields("pedar")
text_ = text_ & "  «—ÌŒ  Ê·œ " & Me.F_msht.Recordset.Fields("tavalod")
text_ = text_ & " òœ „·Ì " & Me.F_msht.Recordset.Fields("kodmeli")
text_ = text_ & "  ·›‰ À«»  " & Me.F_msht.Recordset.Fields("tel1")
text_ = text_ & "  ·›‰ «÷ÿ—«—Ì " & Me.F_msht.Recordset.Fields("tel2")
text_ = text_ & "  ·›‰ Â„—«Â " & Me.F_msht.Recordset.Fields("mob")
text_ = text_ & " ¬œ—” " & Me.F_msht.Recordset.Fields("adress")
text_ = text_ & " ”„  " & Me.F_msht.Recordset.Fields("semat")
text_ = text_ & "  Ê÷ÌÕ«  " & Me.F_msht.Recordset.Fields("tozih")

oExcel.ActiveSheet.Range("C" & I).Value = text_

F_msht.Recordset.MoveNext
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
X = App.Path & "\temp\" & Taqvim.KKK.Caption & "-" & "emza"
On Error GoTo 9

oExcel.SaveAs X
GoTo 8
9:

oExcel.ActiveSheet.Range("G3").Value = ""

For I = 5 To F_msht.Recordset.RecordCount + 4
oExcel.ActiveSheet.Range("B" & I).Value = ""
oExcel.ActiveSheet.Range("C" & I).Value = ""

Next I
oExcel.Close
MsgBox "Œÿ« œ— ›«Ì· Œ—ÊÃÌ: Ê÷⁄Ì  ›«Ì· Â«Ì »«“ —« »——”Ì ò‰Ìœ ::  Ìò ›«Ì· »«“ »« Â„Ì‰ „‘Œ’«  ÊÃÊœ œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Sub
8:

End Sub

Private Sub Command2_Click()
On Error Resume Next

If name_t.Text = "" Or famil_t.Text = "" Then Exit Sub
F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where xname like ('" & name_t.Text & "') and famil like ('" & famil_t.Text & "')"
F_msht.Refresh
If F_msht.Recordset.RecordCount >= 1 Then
If MsgBox("  ⁄œ«œ " & F_msht.Recordset.RecordCount & " „Ê—œ »« Â„Ì‰ ‰«„ Ê ‰«„ Œ«‰Ê«œêÌ Ì«›  ‘œ ¬Ì« „Ì ŒÊ«ÂÌœ œ— Â— ’Ê—  „‘ —Ì ÃœÌœ —« «÷«›Â ò‰Ìœø", vbQuestion + vbYesNo, "À»  „‘ —Ì") = vbYes Then
GoTo 1
Else
Exit Sub
End If
End If
'Exit Sub

1

F_msht.Refresh
F_msht.Recordset.AddNew
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
F_msht.Recordset.Fields("id_usr") = menu.StatusBar1.Panels(2).Text
F_msht.Recordset.Fields("xdate") = Taqvim.KKK.Caption

F_msht.Recordset.Update
F_msht.Refresh
MsgBox "„‘ —Ì ÃœÌœ À»  ‘œ", vbInformation, "À»  „‘ —Ì"
End Sub

Private Sub Command3_Click()
On Error Resume Next

If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ „‘ —Ì —« Õ–› ò‰Ìœø", vbQuestion + vbYesNo, "Õ–› „‘ —Ì") = vbYes Then

F_usr.Refresh
F_usr.RecordSource = "select * from f_usr where id_msht like ('" & F_msht.Recordset.Fields("id_msht") & "')"
F_usr.Refresh
If F_usr.Recordset.EOF = True Or F_usr.Recordset.BOF = True Then
F_msht.Recordset.Delete
MsgBox "„‘ —Ì Õ–› ‘œ", vbInformation, "Õ–› „‘ —Ì"
Else
MsgBox "«„ò«‰ Õ–› «Ì‰ „‘ —Ì ÊÃÊœ ‰œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
End If
End If

End Sub

Private Sub Command6_Click()
On Error Resume Next

edite_msht.Show
edite_msht.Text1.Text = Me.F_msht.Recordset.Fields("id_msht")

End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next

Call Command6_Click

End Sub

Private Sub famil_t_Change()
On Error Resume Next

a = TEXT_CH_JOS_IN_F_KL("famil", famil_t.Text)


End Sub

Private Sub Form_Load()
On Error Resume Next

Set T = New Fx
a = T.KOD_INS_COMBO_F_STTG("semat", Combo1, 0)
End Sub
Function TEXT_CH_JOS_IN_F_KL(filde_, str_)
If ch_Jostan.Value = 1 Then
F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where " & filde_ & " like ('%" & str_ & "%')"
F_msht.Refresh
tedad_yaft.Caption = F_msht.Recordset.RecordCount
End If
End Function

Private Sub kod_meli_t_Change()
On Error Resume Next

a = TEXT_CH_JOS_IN_F_KL("kodmeli", kod_meli_t.Text)

End Sub

Private Sub mnufind_havale_Click()
On Error Resume Next

Find_havale.id_msht_text.Text = Me.F_msht.Recordset.Fields("id_msht")
Find_havale.Show

End Sub

Private Sub mnupic_Click()
On Error Resume Next

sho_pic_emza_f.Show
sho_pic_emza_f.Text1 = Me.F_msht.Recordset.Fields("id_msht")

End Sub

Private Sub mob_t_Change()
a = TEXT_CH_JOS_IN_F_KL("mob", mob_t.Text)

End Sub

Private Sub name_t_Change()
a = TEXT_CH_JOS_IN_F_KL("xname", name_t.Text)

End Sub

Private Sub pedar_t_Change()
a = TEXT_CH_JOS_IN_F_KL("pedar", pedar_t.Text)

End Sub

Private Sub tavalod_t_Change()
a = TEXT_CH_JOS_IN_F_KL("tavalod", tavalod_t.Text)

End Sub

Private Sub tel1_t_Change()
a = TEXT_CH_JOS_IN_F_KL("tel1", tel1_t.Text)

End Sub

Private Sub tel2_t_Change()
a = TEXT_CH_JOS_IN_F_KL("tel2", tel2_t.Text)

End Sub

Private Sub Text9_Change()
On Error Resume Next

F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where  xname + ' ' + famil + ' ' + pedar   like ('%" & Text9.Text & "%')"
F_msht.Refresh
tedad_yaft.Caption = F_msht.Recordset.RecordCount

End Sub

Private Sub tozih_t_Change()
On Error Resume Next

a = TEXT_CH_JOS_IN_F_KL("tozih", tozih_t.Text)

End Sub
