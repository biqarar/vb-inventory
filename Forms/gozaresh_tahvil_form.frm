VERSION 5.00
Begin VB.Form gozaresh_tahvil_form 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ËÈÊ ÒÇÑÔ ÊÍæíá ÍæÇáå"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "gozaresh_tahvil_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "ËÈÊ"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   6615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "ÊÚÏÇÏ ÊÍæíá ÏÇÏå ÔÏå"
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
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label tedad_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÊÚÏÇÏ ÈÇÒÔÊí"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label list_index_l 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label1 
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
      Left            =   6840
      TabIndex        =   5
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "æÖÚíÊ"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   555
   End
End
Attribute VB_Name = "gozaresh_tahvil_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

If Combo2.Text = "" Or Text1.Text = "" Then Exit Sub
a = Split(list_index_l.Caption, " _ ")
b = Split(tahvil_kl_forn.List1.List(a(0)), " :: ")

tahvil_kl_forn.List1.List(a(0)) = b(0) & " :: " & "   æÖÚíÊ:   " & Me.Combo2.Text & "   :ÒÇÑÔ   " & Text9.Text & "   ÊÚÏÇÏ ÈÇÒÔÊí:   " & Text1.Text
tahvil_kl_forn.Show
Unload Me

End Sub

Private Sub Form_Load()
Set T = New Fx
a = T.KOD_INS_COMBO_F_STTG("s_m", Combo2, 0)

End Sub

Private Sub Text1_Change()
On Error Resume Next

If Val(Text1.Text) > Val(tedad_l.Caption) Then Text1.Text = tedad_l.Caption

End Sub
