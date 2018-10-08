VERSION 5.00
Begin VB.Form Change_pass 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " €ÌÌ— ò·„Â ⁄»Ê—"
   ClientHeight    =   3540
   ClientLeft      =   8100
   ClientTop       =   3435
   ClientWidth     =   4890
   Icon            =   "newuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6120
      Picture         =   "newuser.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2500
      End
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "newuser.frx":084C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "newuser.frx":0D8E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   2500
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FF80&
         Caption         =   "À» "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2500
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   2500
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   7
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò·„Â ⁄»Ê— ›⁄·Ì"
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
         Left            =   3165
         TabIndex        =   14
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ ò«—»—Ì"
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
         Index           =   3
         Left            =   3165
         TabIndex        =   13
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò·„Â ⁄»Ê— ÃœÌœ"
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
         Left            =   3165
         TabIndex        =   12
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ﬂ—«— ﬂ·„Â ⁄»Ê— ÃœÌœ"
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
         Left            =   3120
         TabIndex        =   11
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò«—»—"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   2400
         Width           =   285
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2500
   End
End
Attribute VB_Name = "Change_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim D_ As Double

Function f(d)
On Error Resume Next

Dim p, P_ As Double
P_ = 0
p = Val(d)
D_ = (((p * 7) ^ 2) * 3) + 19




End Function
Private Sub Command11_Click()
On Error Resume Next


If Picture7.Visible = True And Picture8.Visible = True Then
MoToR.F_usr.Refresh
MoToR.F_usr.RecordSource = "select * from f_usr where id_usr like ('" & Text1.Text & "')"
MoToR.F_usr.Refresh
a = f(Val(Text7.Text))
If Val(MoToR.F_usr.Recordset.Fields("kl")) = D_ Then

d = f(Val(Text8.Text))
MoToR.F_usr.Recordset.Fields("kl") = D_
MoToR.F_usr.Recordset.Fields("usr") = Text11.Text
MoToR.F_usr.Recordset.Update
MsgBox " €ÌÌ— ò·„Â ⁄»Ê— «‰Ã«„ ‘œ", vbInformation, " €ÌÌ— ò·„Â ⁄»Ê—"
Unload Me

Else
MsgBox "ò·„Â ⁄»Ê— ›⁄·Ì «‘ »«Â «” ", vbCritical + vbOKOnly, " €ÌÌ— ò·„Â ⁄»Ê—"
Exit Sub

End If
End If

End Sub

Private Sub Text1_Change()
On Error Resume Next

MoToR.F_usr.Refresh
MoToR.F_usr.RecordSource = "select * from f_usr where id_usr like ('" & Text1.Text & "')"
MoToR.F_usr.Refresh
Text11.Text = MoToR.F_usr.Recordset.Fields("usr")
MoToR.F_msht.Refresh
MoToR.F_msht.RecordSource = "select * from f_msht where id_msht like ('" & MoToR.F_usr.Recordset.Fields("id_msht") & "')"
MoToR.F_msht.Refresh

If MoToR.F_msht.Recordset.EOF = False Or MoToR.F_msht.Recordset.BOF = False Then
Text10.Text = MoToR.F_msht.Recordset.Fields("xname") & " " & MoToR.F_msht.Recordset.Fields("famil")
'menu.StatusBar1.Panels(2).Text = MoToR.F_usr.Recordset.Fields("id_usr")
End If




End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text9_Change()
If Text9.Text = Text8.Text Then
Picture8.Visible = True
Picture7.Visible = True
Else
Picture8.Visible = False
Picture7.Visible = False
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
