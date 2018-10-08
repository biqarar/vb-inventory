VERSION 5.00
Begin VB.Form Add_goroh_f 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ê—ÊÂ Â«"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Add_goroh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   1200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "Add_goroh.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Õ–› ê—ÊÂ"
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "Add_goroh.frx":48CF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "À»  ê—ÊÂ"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox NAME_GOROH_T 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox TOZIH_GOROH_T 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2280
      Width           =   6615
      Begin VB.ListBox goroh_list 
         BackColor       =   &H00C0E0FF&
         Height          =   3165
         ItemData        =   "Add_goroh.frx":832C
         Left            =   120
         List            =   "Add_goroh.frx":832E
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ê—ÊÂ Â«Ì „ÊÃÊœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "ÿ»ﬁÂ"
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
      Left            =   6120
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ê—ÊÂ —« Õ–› ò‰Ìœø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label ins_l 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ê—ÊÂ «÷«›Â ‘œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label eror_l 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "«Ì‰ ê—ÊÂ œ— ·Ì”  ÊÃÊœ œ«—œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Ê÷ÌÕ« "
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   6000
      TabIndex        =   11
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ê—ÊÂ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   6000
      TabIndex        =   10
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "Add_goroh_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Secent_for_show_and_hide_label As Integer

Private Sub Text3_Change()

End Sub

Private Sub Command10_Click()
On Error Resume Next

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Command10.Visible = False
Command7.Visible = False



End Sub

Private Sub Command7_Click()
On Error Resume Next

If NAME_GOROH_T.Text = "" And TOZIH_GOROH_T.Text = "" Then Exit Sub
If NAME_GOROH_T.Text = "" Then Exit Sub
MoToR.F_gh.Refresh
MoToR.F_gh.Recordset.AddNew

MoToR.F_gh.Recordset.Fields("xname") = NAME_GOROH_T.Text
MoToR.F_gh.Recordset.Fields("tozih") = TOZIH_GOROH_T.Text
a = Split(Combo1.Text, " _ ")
MoToR.F_gh.Recordset.Fields("id_tbq") = a(0)
MoToR.F_gh.Recordset.Fields("id_usr") = menu.StatusBar1.Panels(2).Text
MoToR.F_gh.Recordset.Update
MoToR.F_gh.Refresh
a = Show_error_or_insert("ê—ÊÂ ÃœÌœ «÷«›Â ‘œ")

Refresh_goroh_list


End Sub
Function Refresh_goroh_list()
On Error Resume Next

MoToR.F_gh.Refresh
MoToR.F_gh.RecordSource = "select * from F_gh" ' where xname like ('%" & "" & "%')"
MoToR.F_gh.Refresh
goroh_list.Clear

For I = 1 To MoToR.F_gh.Recordset.RecordCount
goroh_list.AddItem (MoToR.F_gh.Recordset.Fields("id_gh") & " _ " & MoToR.F_gh.Recordset.Fields("id_tbq") & " _ " & MoToR.F_gh.Recordset.Fields("xname") & "  ::  " & MoToR.F_gh.Recordset.Fields("tozih"))
MoToR.F_gh.Recordset.MoveNext
Next I
goroh_list.Text = goroh_list.List(0)

End Function

Private Sub Form_Load()
On Error Resume Next

Refresh_goroh_list
Set T = New Fx
a = T.INSERT_F_TBQ(Combo1, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Me

End Sub

Private Sub goroh_list_Click()
On Error Resume Next

Command10.Visible = True

End Sub

Private Sub goroh_list_DblClick()
On Error Resume Next

show_more_fctr_kl.Show
show_more_fctr_kl.Text1.Text = Me.goroh_list.Text

End Sub

Private Sub Label2_Click()
On Error Resume Next

'On Error Resume Next
a = Split(goroh_list.Text, " _ ")
MoToR.F_kl_gh.Refresh
MoToR.F_kl_gh.RecordSource = "select * from f_kl_gh where id_gh like ('" & a(0) & "')"
MoToR.F_kl_gh.Refresh

If MoToR.F_kl_gh.Recordset.EOF = True Or MoToR.F_kl_gh.Recordset.BOF = True Then
MoToR.F_gh.Refresh
MoToR.F_gh.RecordSource = "select * from F_gh where id_gh like ('" & a(0) & "')"
MoToR.F_gh.Refresh
MoToR.F_gh.Recordset.Delete
a = Show_error_or_insert("ê—ÊÂ Õ–› ‘œ")
Refresh_goroh_list

Else

a = Show_error_or_insert("«Ì‰ ê—ÊÂ œ— Õ«· «” ›«œÂ „Ì »«‘œ")
End If





Command10.Visible = True

Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub
Function Show_error_or_insert(str_)
On Error Resume Next

Command10.Visible = True
'Refresh_goroh_list
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
eror_l.Visible = True
eror_l.Caption = str_
eror_l.BackColor = &HFF&
Secent_for_show_and_hide_label = 40

Timer1.Enabled = True


End Function
Private Sub Label3_Click()
On Error Resume Next

Command10.Visible = True
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Secent_for_show_and_hide_label = Secent_for_show_and_hide_label - 1
If Secent_for_show_and_hide_label = 0 Then
eror_l.Visible = False
ins_l.Visible = False

Timer1.Enabled = False
End If


End Sub
