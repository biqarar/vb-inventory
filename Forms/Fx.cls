VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function KOD_INS_COMBO_F_STTG(KOD_, Combo_, id_)
On Error Resume Next

MoToR.F_sttg.Refresh
MoToR.F_sttg.RecordSource = "select * from f_sttg where kod like ('" & KOD_ & "')"
MoToR.F_sttg.Refresh
For I = 1 To MoToR.F_sttg.Recordset.RecordCount
If id_ = 1 Then
Combo_.AddItem (MoToR.F_sttg.Recordset.Fields("id_sttg") & " _ " & MoToR.F_sttg.Recordset.Fields("xvalue"))
Else
Combo_.AddItem (MoToR.F_sttg.Recordset.Fields("xvalue"))
End If
MoToR.F_sttg.Recordset.MoveNext
Next I
End Function
Function INSERT_F_TBQ(Combo_, id_)
On Error Resume Next

MoToR.F_tbq.Refresh
MoToR.F_tbq.RecordSource = "select * from F_tbq  " ' kod like ('" & KOD_ & "')"
MoToR.F_tbq.Refresh
MoToR.F_tbq.Recordset.Sort = "id_tbq"

For I = 1 To MoToR.F_tbq.Recordset.RecordCount
If id_ = 1 Then
Combo_.AddItem (MoToR.F_tbq.Recordset.Fields("id_tbq") & " _ " & " ”«·‰ " & MoToR.F_tbq.Recordset.Fields("salon") & "  —œÌ›  " & MoToR.F_tbq.Recordset.Fields("radif") & "  ÿ»ﬁÂ  " & MoToR.F_tbq.Recordset.Fields("tabaqe") & "  ::  " & MoToR.F_tbq.Recordset.Fields("tozih"))
Else
Combo_.AddItem (" ”«·‰ " & MoToR.F_tbq.Recordset.Fields("salon") & "  —œÌ›  " & MoToR.F_tbq.Recordset.Fields("radif") & "  ÿ»ﬁÂ  " & MoToR.F_tbq.Recordset.Fields("tabaqe") & "  ::  " & MoToR.F_tbq.Recordset.Fields("tozih"))
End If
MoToR.F_tbq.Recordset.MoveNext
Next I
End Function
Function INSERT_F_GH(Combo_, id_)
On Error Resume Next

MoToR.F_gh.Refresh
MoToR.F_gh.RecordSource = "select * from F_gh  " ' kod like ('" & KOD_ & "')"
MoToR.F_gh.Refresh
MoToR.F_gh.Recordset.Sort = "id_gh"

For I = 1 To MoToR.F_gh.Recordset.RecordCount
If id_ = 1 Then
Combo_.AddItem (MoToR.F_gh.Recordset.Fields("id_gh") & " _ " & MoToR.F_gh.Recordset.Fields("xname") & "  _  " & MoToR.F_gh.Recordset.Fields("tozih"))
Else
Combo_.AddItem (MoToR.F_gh.Recordset.Fields("xname") & "  _  " & MoToR.F_gh.Recordset.Fields("tozih"))
End If
MoToR.F_gh.Recordset.MoveNext
Next I
End Function
Function XLSX_Havale(id_fcactor_)
Dim oExcel As Object
Dim AD As String
'On Error GoTo 1
GoTo 2
1: MsgBox "›«Ì· ÅÌœ« ‰‘œ∫ »—«Ì «Ã—«Ì «Ì‰ ›—„«‰ »«Ìœ «“ ”Ì” „ ”—Ê— «” ›«œÂ ò‰Ìœ", vbCritical, "Œÿ«"
Exit Function
2:
MoToR.F_fctr_msht.Refresh
MoToR.F_fctr_msht.RecordSource = "select * from f_fctr_msht where id_fctr like ('" & id_fcactor_ & "')"
MoToR.F_fctr_msht.Refresh
Dim BePAR As Integer

BePAR = 0
If Val(MoToR.F_fctr_msht.Recordset.RecordCount) <= 10 Then
Set oExcel = GetObject(App.Path & "\formXLS\havaleA5.xls")
ElseIf Val(MoToR.F_fctr_msht.Recordset.RecordCount) > 10 And Val(MoToR.F_fctr_msht.Recordset.RecordCount) <= 30 Then
Set oExcel = GetObject(App.Path & "\formXLS\havaleA4.xls")
ElseIf Val(MoToR.F_fctr_msht.Recordset.RecordCount) > 30 Then
Set oExcel = GetObject(App.Path & "\formXLS\havalefull.xls")
BePAR = 1
End If

oExcel.ActiveSheet.Range("G2").Value = id_fcactor_

MoToR.F_fctr.Refresh
MoToR.F_fctr.RecordSource = "select * from F_fctr where id_fctr like ('" & id_fcactor_ & "')"
MoToR.F_fctr.Refresh
oExcel.ActiveSheet.Range("G3").Value = MoToR.F_fctr.Recordset.Fields("xdate")
oExcel.ActiveSheet.Range("A16").Value = "   «—ÌŒ  ÕÊÌ·:  " & MoToR.F_fctr.Recordset.Fields("xend_date") & "   Ê÷ÌÕ« :  " & MoToR.F_fctr.Recordset.Fields("tozih")
oExcel.ActiveSheet.Range("D3").Value = MoToR.F_fctr.Recordset.Fields("noe")

MoToR.F_msht.Refresh
MoToR.F_msht.RecordSource = "select * from F_msht where id_msht like ('" & MoToR.F_fctr.Recordset.Fields("id_msht") & "')"
MoToR.F_msht.Refresh

oExcel.ActiveSheet.Range("C4").Value = MoToR.F_msht.Recordset.Fields("xname") & "  " & MoToR.F_msht.Recordset.Fields("famil")
oExcel.ActiveSheet.Range("F4").Value = MoToR.F_msht.Recordset.Fields("tel1") & "  -  " & MoToR.F_msht.Recordset.Fields("mob")

If BePAR = 1 Then GoTo 11
oExcel.ActiveSheet.Range("A18").Value = MoToR.F_msht.Recordset.Fields("xname") & "  " & MoToR.F_msht.Recordset.Fields("famil")

MoToR.F_msht.Refresh
MoToR.F_msht.RecordSource = "select * from F_msht where id_msht like ('" & MoToR.F_fctr.Recordset.Fields("id_usr") & "')"
MoToR.F_msht.Refresh
oExcel.ActiveSheet.Range("C18").Value = "   " & MoToR.F_msht.Recordset.Fields("xname") & "  " & MoToR.F_msht.Recordset.Fields("famil")
11:

MoToR.F_fctr_msht.Refresh
MoToR.F_fctr_msht.RecordSource = "select * from f_fctr_msht where id_fctr like ('" & id_fcactor_ & "')"
MoToR.F_fctr_msht.Refresh

For I = 6 To MoToR.F_fctr_msht.Recordset.RecordCount + 5
oExcel.ActiveSheet.Range("B" & I).Value = MoToR.F_fctr_msht.Recordset.Fields("id_kl")

MoToR.F_kl.Refresh
MoToR.F_kl.RecordSource = "select * from F_kl where id_kl like ('" & MoToR.F_fctr_msht.Recordset.Fields("id_kl") & "')"
MoToR.F_kl.Refresh

oExcel.ActiveSheet.Range("C" & I).Value = MoToR.F_kl.Recordset.Fields("xnoe")
oExcel.ActiveSheet.Range("D" & I).Value = MoToR.F_fctr_msht.Recordset.Fields("tedad") & " " & MoToR.F_kl.Recordset.Fields("vahed")
oExcel.ActiveSheet.Range("E" & I).Value = MoToR.F_kl.Recordset.Fields("code")
oExcel.ActiveSheet.Range("F" & I).Value = MoToR.F_fctr_msht.Recordset.Fields("v_tahvil")

MoToR.F_fctr_msht.Recordset.MoveNext
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
X = App.Path & "\temp\" & Taqvim.KKK.Caption & "-" & id_fcactor_
On Error GoTo 9

oExcel.SaveAs X
MoToR.F_fctr.Refresh
MoToR.F_fctr.RecordSource = "select * from F_fctr where id_fctr like ('" & id_fcactor_ & "')"
MoToR.F_fctr.Refresh
MoToR.F_fctr.Recordset.Fields("chap") = " ç«Å ‘œÂ œ—  «—ÌŒ " & Taqvim.KKK.Caption
MoToR.F_fctr.Recordset.Update


GoTo 8
9:
oExcel.ActiveSheet.Range("G2").Value = ""
oExcel.ActiveSheet.Range("G3").Value = ""
oExcel.ActiveSheet.Range("A16").Value = ""
oExcel.ActiveSheet.Range("D3").Value = ""

oExcel.ActiveSheet.Range("C4").Value = ""
oExcel.ActiveSheet.Range("F4").Value = ""

If BePAR = 1 Then GoTo 118
oExcel.ActiveSheet.Range("A18").Value = ""
oExcel.ActiveSheet.Range("C18").Value = ""
118:
For I = 6 To MoToR.F_fctr_msht.Recordset.RecordCount + 5
oExcel.ActiveSheet.Range("B" & I).Value = ""
oExcel.ActiveSheet.Range("C" & I).Value = ""
oExcel.ActiveSheet.Range("D" & I).Value = ""
oExcel.ActiveSheet.Range("E" & I).Value = ""
oExcel.ActiveSheet.Range("F" & I).Value = ""
Next I
oExcel.Close
MsgBox "Œÿ« œ— ›«Ì· Œ—ÊÃÌ: Ê÷⁄Ì  ›«Ì· Â«Ì »«“ —« »——”Ì ò‰Ìœ ::   Ìò ›«Ì· »«“ »« Â„Ì‰ „‘Œ’«  ÊÃÊœ œ«—œ", vbCritical + vbOKOnly, "Œÿ«"
Exit Function
8:
End Function
