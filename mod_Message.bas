Attribute VB_Name = "mod_Message"
Option Explicit

Private Function ComponentFileName(Cmp As VBComponent) As String

  ComponentFileName = Cmp.Name & "." & ModuleExt(Cmp.CodeModule)

End Function

Public Sub MemberMessage(ByVal StrCap As String, _
                         ByVal lngCurValue As Double, _
                         ByVal lngMaxValue As Long)

  If lngMaxValue Then
    mObjDoc.DrawPercent2 4, lngCurValue / lngMaxValue * 100, StrCap
  End If

End Sub

Private Function ModuleExt(codeMod As CodeModule) As String

  'Based on code found at
  'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42065&lngWId=1
  'Submitted on: 1/1/2003 4:26:20 PM
  'By: Mark Nemtsas

  Select Case codeMod.Parent.Type
   Case vbext_ct_StdModule    '1 standard module.
    ModuleExt = "bas"
   Case vbext_ct_ClassModule '2 class module.
    ModuleExt = "cls"
   Case vbext_ct_MSForm       '3 form.
    ModuleExt = "frm"
   Case vbext_ct_ResFile      '4 standard resource file.
    ModuleExt = "res"
   Case vbext_ct_VBForm       '5 Visual Basic form.
    ModuleExt = "frm"
   Case vbext_ct_VBMDIForm    '6 The component is an MDI form.
    ModuleExt = "frm"
   Case vbext_ct_PropPage     '7 property page.
    ModuleExt = "pag"
   Case vbext_ct_UserControl  '8 user control.
    ModuleExt = "ctl"
   Case vbext_ct_DocObject     '9 RelatedDocument.
    ModuleExt = "dob"
   Case vbext_ct_ActiveXDesigner '11 ActiveX designer.
    ModuleExt = "dsr"
  End Select

End Function

Public Sub ModuleMessage(Cmp As VBComponent, _
                         ByVal Counter As Double)

  mObjDoc.DrawPercent2 2, Counter / ComponentCount * 100, ComponentFileName(Cmp)

End Sub

Public Sub SectionMessage(ByVal StrCap As String, _
                          ByVal Counter As Double)

  mObjDoc.DrawPercent2 1, Counter * 100, StrCap

End Sub

Public Sub WorkingMessage(ByVal StrCap As String, _
                          ByVal lngCurValue As Long, _
                          ByVal lngMaxValue As Long)

  If lngMaxValue > 0 Then
    mObjDoc.DrawPercent2 3, lngCurValue / lngMaxValue * 100, StrCap
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:16:33 AM) 1 + 69 = 70 Lines Thanks Ulli for inspiration and lots of code.

