VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Find_havale 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã” ÃÊ œ— ÕÊ«·Â Â«"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14655
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Find_havale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   14655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "‰„«Ì‘ ÕÊ«·Â Â«Ì »«ÿ· ‘œÂ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "ç«Å ÕÊ«·Â"
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
      Height          =   435
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "‰„«Ì‘  „«„Ì ÕÊ«·Â Â«Ì ›⁄«·"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Caption         =   "ADODC"
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
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
   Begin VB.TextBox id_msht_text 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   8760
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "›—„  ÕÊÌ· ÕÊ«·Â"
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
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.TextBox id_kl_text 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   7200
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   2670
      ItemData        =   "Find_havale.frx":030A
      Left            =   240
      List            =   "Find_havale.frx":030C
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   14295
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
      Height          =   2670
      ItemData        =   "Find_havale.frx":030E
      Left            =   240
      List            =   "Find_havale.frx":0310
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   14295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÕÊ«·Â Â«Ì ›⁄«·"
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   6840
      TabIndex        =   10
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÕÊ«·Â Â«Ì ›⁄«·"
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   13200
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "«ﬁ·«„ „ÊÃÊœ œ— «Ì‰ ÕÊ«·Â"
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   12480
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
End
Attribute VB_Name = "Find_havale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

a = Split(Me.List2.Text, " _ ")
tahvil_kl_forn.Text2.Text = ""
tahvil_kl_forn.Text2.Text = 1

tahvil_kl_forn.Text9.Text = a(0)
tahvil_kl_forn.Show
Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next

F_fctr.Refresh
F_fctr.RecordSource = "select * from F_fctr where vazeyat like ('out')"
F_fctr.Refresh
If F_fctr.Recordset.RecordCount = 0 Then Exit Sub


List2.Clear

For I = 1 To F_fctr.Recordset.RecordCount
F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_msht.Refresh
nae_famil = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
List2.AddItem (F_fctr.Recordset.Fields("id_fctr") & " _ " & "  «—ÌŒ " & F_fctr.Recordset.Fields("xdate") & " ‰Ê⁄ ÕÊ«·Â " & F_fctr.Recordset.Fields("noe") & "  ÕÊÌ· œÂ‰œÂ " & nae_famil & "  Ê÷ÌÕ« : " & F_fctr.Recordset.Fields("tozih"))
F_fctr.Recordset.MoveNext
Next I
End Sub

Private Sub Command3_Click()
On Error Resume Next

a = Split(Me.List2.Text, " _ ")
T = Print_factore(a(0))
End Sub
Function Print_factore(id_factore)
'some code....
Set T = New Fx
a = T.XLSX_Havale(id_factore)

End Function

Private Sub Command4_Click()
On Error Resume Next

F_fctr.Refresh
F_fctr.RecordSource = "select * from F_fctr where vazeyat like ('in')"
F_fctr.Refresh
If F_fctr.Recordset.RecordCount = 0 Then Exit Sub


List2.Clear

For I = 1 To F_fctr.Recordset.RecordCount
F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_msht.Refresh
nae_famil = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
List2.AddItem (F_fctr.Recordset.Fields("id_fctr") & " _ " & "  «—ÌŒ " & F_fctr.Recordset.Fields("xdate") & " ‰Ê⁄ ÕÊ«·Â " & F_fctr.Recordset.Fields("noe") & "  ÕÊÌ· œÂ‰œÂ " & nae_famil & "  Ê÷ÌÕ« : " & F_fctr.Recordset.Fields("tozih"))
F_fctr.Recordset.MoveNext
Next I
End Sub

Private Sub id_kl_text_Change()
On Error Resume Next

F_kl.Refresh
F_kl.RecordSource = "select * from F_kl where id_kl like ('" & id_kl_text.Text & "')"
F_kl.Refresh
If F_kl.Recordset.RecordCount = 0 Then Exit Sub


Label3.Caption = F_kl.Recordset.Fields("xnoe")
F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "SELECT DISTINCT id_fctr from f_fctr_msht where vazeyat like ('out') and id_kl like ('" & id_kl_text.Text & "')"
F_fctr_msht.Refresh
List2.Clear

For I = 1 To F_fctr_msht.Recordset.RecordCount
F_fctr.Refresh
F_fctr.RecordSource = "select * from F_fctr where id_fctr like ('" & F_fctr_msht.Recordset.Fields("id_fctr") & "')"
F_fctr.Refresh

F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_msht.Refresh
nae_famil = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")

F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_msht") & "')"
F_msht.Refresh
moshtari = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")

List2.AddItem (F_fctr.Recordset.Fields("id_fctr") & " _ " & "  «—ÌŒ " & F_fctr.Recordset.Fields("xdate") & " ‰Ê⁄ ÕÊ«·Â: " & F_fctr.Recordset.Fields("noe") & "  ÕÊÌ· œÂ‰œÂ: " & nae_famil & "  ÕÊÌ· êÌ—‰œÂ: " & moshtari & "  Ê÷ÌÕ« : " & F_fctr.Recordset.Fields("tozih"))
F_fctr.Recordset.MoveNext
Next I
End Sub

Private Sub Text1_Change()



End Sub

Private Sub id_msht_text_Change()
On Error Resume Next

F_fctr.Refresh
F_fctr.RecordSource = "select * from F_fctr where id_msht like ('" & Me.id_msht_text.Text & "') and vazeyat like ('out')"
F_fctr.Refresh
If F_fctr.Recordset.RecordCount = 0 Then Exit Sub

F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & Me.id_msht_text.Text & "')"
F_msht.Refresh
Label3.Caption = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
List2.Clear

For I = 1 To F_fctr.Recordset.RecordCount
F_msht.Refresh
F_msht.RecordSource = "select * from F_msht where id_msht like ('" & F_fctr.Recordset.Fields("id_usr") & "')"
F_msht.Refresh
nae_famil = F_msht.Recordset.Fields("xname") & " " & F_msht.Recordset.Fields("famil")
List2.AddItem (F_fctr.Recordset.Fields("id_fctr") & " _ " & "  «—ÌŒ " & F_fctr.Recordset.Fields("xdate") & " ‰Ê⁄ ÕÊ«·Â " & F_fctr.Recordset.Fields("noe") & "  ÕÊÌ· œÂ‰œÂ " & nae_famil & "  Ê÷ÌÕ« : " & F_fctr.Recordset.Fields("tozih"))
F_fctr.Recordset.MoveNext
Next I

End Sub

Private Sub List2_Click()
On Error Resume Next

Command1.Visible = True
Command3.Visible = True

a = Split(Me.List2.Text, " _ ")

F_fctr_msht.Refresh
F_fctr_msht.RecordSource = "SELECT * from f_fctr_msht where id_fctr like ('" & a(0) & "')"
F_fctr_msht.Refresh
List1.Clear

For I = 1 To F_fctr_msht.Recordset.RecordCount

F_kl.Refresh
F_kl.RecordSource = "select * from F_kl where id_kl like ('" & F_fctr_msht.Recordset.Fields("id_kl") & "')"
F_kl.Refresh
List1.AddItem (F_kl.Recordset.Fields("id_kl") & " _ " & F_kl.Recordset.Fields("xnoe") & "   Ê÷€Ì   ÕÊÌ·:     " & F_fctr_msht.Recordset.Fields("v_tahvil") & "    Ê÷⁄Ì  œ—Ì«› :     " & F_fctr_msht.Recordset.Fields("v_tahvil") & "   ê“«—‘   " & F_fctr_msht.Recordset.Fields("gozaresh") & " :: ")

F_fctr_msht.Recordset.MoveNext
Next I


End Sub
