VERSION 5.00
Begin VB.Form tabaqe 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»ﬁÂ »‰œÌ «‰»«—"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tabaqe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox salon_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox radif_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox tabaqe_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      TabIndex        =   4
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox makan_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox shobe_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      TabIndex        =   0
      Top             =   360
      Width           =   3735
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
      Height          =   4455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   7575
      Begin VB.ListBox goroh_list 
         BackColor       =   &H00C0E0FF&
         Height          =   3510
         ItemData        =   "tabaqe.frx":030A
         Left            =   120
         List            =   "tabaqe.frx":030C
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   7335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ»ﬁÂ »‰œÌ Â«Ì „ÊÃÊœ"
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
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.TextBox tozih_t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   465
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   3735
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
      Left            =   7920
      Picture         =   "tabaqe.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "À»  ê—ÊÂ"
      Top             =   360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   1200
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
      Left            =   7920
      Picture         =   "tabaqe.frx":3D6B
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Õ–› ê—ÊÂ"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”«·‰"
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
      Left            =   12840
      TabIndex        =   21
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
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
      Left            =   12840
      TabIndex        =   20
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   12840
      TabIndex        =   19
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ò«‰"
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
      Left            =   12840
      TabIndex        =   18
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘⁄»Â"
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
      Left            =   12840
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   12840
      TabIndex        =   16
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label eror_l 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "«Ì‰ ê—ÊÂ œ— ·Ì”  ÊÃÊœ œ«—œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label ins_l 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ê—ÊÂ «÷«›Â ‘œ"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «ÿ·«⁄«  —« Õ–› òÌ‰œø"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "»·Ì"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9000
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ŒÌ—"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "tabaqe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Secent_for_show_and_hide_label As Integer
Private Sub Command10_Click()
On Error Resume Next

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Command10.Visible = False
Command7.Visible = False

red = &HFF&
green = &HC000&

End Sub

Private Sub Command7_Click()
On Error Resume Next

If salon_t.Text = "" Or makan_t.Text = "" Or shobe_t.Text = "" Or radif_t.Text = "" Or tabaqe_t.Text = "" Then Exit Sub
MoToR.F_tbq.Refresh
MoToR.F_tbq.RecordSource = "select * from f_tbq where shobe like ('" & shobe_t.Text & "') and makan like ('" & makan_t.Text & "') and salon like ('" & salon_t.Text & "') and radif like ('" & radif_t.Text & "') and tabaqe like ('" & tabaqe_t.Text & "') and tozih like ('" & tozih_t.Text & "')"
MoToR.F_tbq.Refresh
If MoToR.F_tbq.Recordset.EOF = False Or MoToR.F_tbq.Recordset.BOF = False Then
a = Show_error_or_insert("«Ì‰ „Ê—œ ﬁ»·« «÷«›Â ‘œÂ «” ", &HFF&)
Exit Sub
End If

MoToR.F_tbq.Refresh
MoToR.F_tbq.Recordset.AddNew
MoToR.F_tbq.Recordset.Fields("shobe") = shobe_t.Text
MoToR.F_tbq.Recordset.Fields("makan") = makan_t.Text
MoToR.F_tbq.Recordset.Fields("salon") = salon_t.Text
MoToR.F_tbq.Recordset.Fields("radif") = radif_t.Text
MoToR.F_tbq.Recordset.Fields("tabaqe") = tabaqe_t.Text
MoToR.F_tbq.Recordset.Fields("tozih") = tozih_t.Text
MoToR.F_tbq.Recordset.Fields("id_usr") = menu.StatusBar1.Panels(2).Text

MoToR.F_tbq.Recordset.Update
MoToR.F_tbq.Refresh
a = Show_error_or_insert("«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", &HC000&)
refresh_tabaqe_list

End Sub

Private Sub Form_Load()
refresh_tabaqe_list

End Sub
Function Show_error_or_insert(str_, color_)
On Error Resume Next

Command10.Visible = True
'Refresh_goroh_list
Command7.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
eror_l.Visible = True
eror_l.Caption = str_
eror_l.BackColor = color_
Secent_for_show_and_hide_label = 20

Timer1.Enabled = True


End Function
Function refresh_tabaqe_list()
On Error Resume Next

MoToR.F_tbq.Refresh
MoToR.F_tbq.RecordSource = "select * from F_tbq  " ' kod like ('" & KOD_ & "')"
MoToR.F_tbq.Refresh
MoToR.F_tbq.Recordset.Sort = "id_tbq"
goroh_list.Clear

For I = 1 To MoToR.F_tbq.Recordset.RecordCount

goroh_list.AddItem (MoToR.F_tbq.Recordset.Fields("id_tbq") & " _ " & "  ‘⁄»Â  " & MoToR.F_tbq.Recordset.Fields("shobe") & "  „Êﬁ⁄Ì   " & MoToR.F_tbq.Recordset.Fields("makan") & " ”«·‰ " & MoToR.F_tbq.Recordset.Fields("salon") & "  —œÌ›  " & MoToR.F_tbq.Recordset.Fields("radif") & "  ÿ»ﬁÂ  " & MoToR.F_tbq.Recordset.Fields("tabaqe") & "  ::  " & MoToR.F_tbq.Recordset.Fields("tozih"))


MoToR.F_tbq.Recordset.MoveNext
Next I
End Function

Private Sub goroh_list_Click()

End Sub

Private Sub goroh_list_DblClick()
On Error Resume Next

show_more_fctr_kl.Show
show_more_fctr_kl.Text1.Text = Me.goroh_list.Text
End Sub

Private Sub Label2_Click()
On Error GoTo 1


a = Split(goroh_list.Text, " _ ")
MoToR.F_gh.Refresh
MoToR.F_gh.RecordSource = "select * from f_gh where id_tbq like ('" & a(0) & "')"
MoToR.F_gh.Refresh
If MoToR.F_gh.Recordset.EOF = True Or MoToR.F_gh.Recordset.BOF = True Then
MoToR.F_tbq.Refresh
MoToR.F_tbq.RecordSource = "select * from F_tbq where id_tbq like ('" & a(0) & "')"
MoToR.F_tbq.Refresh
MoToR.F_tbq.Recordset.Delete
a = Show_error_or_insert("ê—ÊÂ Õ–› ‘œ", &HC000&)
refresh_tabaqe_list


Else

a = Show_error_or_insert("«Ì‰ ê—ÊÂ œ— Õ«· «” ›«œÂ „Ì »«‘œ", &HFF&)
End If

1:
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Command10.Visible = True
Command7.Visible = True

End Sub

Private Sub Label3_Click()
On Error Resume Next

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Command10.Visible = True
Command7.Visible = True


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
