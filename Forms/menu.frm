VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form menu 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁ—«—ê«Â ›—Â‰êÌ ›«ÿ„ÌÊ‰"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C000&
      Caption         =   "Ã” ÃÊ œ— ÕÊ«·Â Â«"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   " ÕÊÌ· ÕÊ«·Â"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "ÿ»ﬁÂ »‰œÌ «‰»«—"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "ê—ÊÂ »‰œÌ ÿ»ﬁ« "
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "œ—ŒÊ«”  ÕÊ«·Â «“ «‰»«—"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "À»  „‘ —Ì"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3495
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5256
            MinWidth        =   5256
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "œ—»«—Â »—‰«„Â"
            TextSave        =   "œ—»«—Â »—‰«„Â"
            Object.ToolTipText     =   "»—«Ì ‰„«Ì‘ ò·Ìò ò‰Ìœ"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "À»  ò«·«"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "—ÊÌœ«œ Â«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3255
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "„‘ —Ì Â«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   7800
      TabIndex        =   10
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "ê—ÊÂ »‰œÌ Ê ÿ»ﬁÂ »‰œÌ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "ò«·« Â«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   7800
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu nmufile 
      Caption         =   "Å—Ê‰œÂ"
      Begin VB.Menu mnu_user 
         Caption         =   "„œÌ—Ì  ò«—»—«‰"
         Begin VB.Menu mnu_change_password 
            Caption         =   " €ÌÌ— ò·„Â ⁄»Ê—"
         End
      End
      Begin VB.Menu mnufile_open 
         Caption         =   "»«“ ò—œ‰ ÅÊ‘Â"
      End
      Begin VB.Menu d 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Mnuend 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
add_msht.Show

End Sub

Private Sub Command2_Click()
add_kl.Show

End Sub

Private Sub Command3_Click()
Factor_f.Show

End Sub

Private Sub Command4_Click()
Add_goroh_f.Show

End Sub

Private Sub Command5_Click()
tabaqe.Show

End Sub

Private Sub Command6_Click()
tahvil_kl_forn.Show

End Sub

Private Sub Command7_Click()
Find_havale.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub mnu_change_password_Click()
Change_pass.Text1.Text = Me.StatusBar1.Panels(2).Text
Change_pass.Show

End Sub

Private Sub Mnuend_Click()
End

End Sub

Private Sub mnufile_open_Click()
Shell "explorer.exe " & App.Path, vbNormalFocus


End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
WE.Show

End Sub
