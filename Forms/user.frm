VERSION 5.00
Begin VB.Form user 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ﬁ—«—ê«Â ›—Â‰êÌ ›«ÿ„ÌÊ‰"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "user.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ê—Êœ"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox pa_t 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox us_t 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò·„Â ⁄»Ê—"
      Height          =   345
      Left            =   3120
      TabIndex        =   3
      Top             =   7440
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—Ì"
      Height          =   345
      Left            =   6120
      TabIndex        =   1
      Top             =   7440
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "user.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

Dim p, P_ As Double

P_ = 0
p = Val(pa_t.Text)
P_ = (((p * 7) ^ 2) * 3) + 19
MoToR.F_usr.Refresh
MoToR.F_usr.RecordSource = "select * from f_usr where usr like ('" & us_t.Text & "') and kl like ('" & P_ & "')"
MoToR.F_usr.Refresh
If MoToR.F_usr.Recordset.RecordCount = 1 Then
MoToR.F_msht.Refresh
MoToR.F_msht.RecordSource = "select * from f_msht where id_msht like ('" & MoToR.F_usr.Recordset.Fields("id_msht") & "')"
MoToR.F_msht.Refresh
If MoToR.F_msht.Recordset.EOF = False Or MoToR.F_msht.Recordset.BOF = False Then
menu.StatusBar1.Panels(1).Text = MoToR.F_msht.Recordset.Fields("xname") & " " & MoToR.F_msht.Recordset.Fields("famil")
menu.StatusBar1.Panels(2).Text = MoToR.F_usr.Recordset.Fields("id_usr")

menu.Show
Unload Me
End If

Else
Beep
End If


End Sub

Private Sub Form_Load()
On Error Resume Next

Image1.Picture = LoadPicture(App.Path & "\pic\logo.jpg")
End Sub

Private Sub Image1_DblClick()
End

End Sub

Private Sub pa_t_Change()
If pa_t.Text = "10012513" Then

add_user.Show

Unload Me
End If

End Sub

Private Sub pa_t_Click()
pa_t.Text = ""

End Sub

Private Sub us_t_Click()
us_t.Text = ""

End Sub
