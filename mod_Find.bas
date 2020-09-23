Attribute VB_Name = "mod_Find"
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
''Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Option Explicit

Public Function GetCurrentLSVLine(Lsv As ListView) As Long

  If Lsv.ListItems.Count Then
    GetCurrentLSVLine = Lsv.SelectedItem.Index
   Else
    GetCurrentLSVLine = 0
  End If

End Function

Public Function InProcedure(ByVal strTest As String, _
                            Cmp As VBComponent, _
                            ByVal Sline As Long) As Boolean

  On Error Resume Next
  If strTest = "(Declarations)" And Sline <= Cmp.CodeModule.CountOfDeclarationLines Then
    'dummy for detecing that item is in Declaration section
    InProcedure = True
   ElseIf strTest = Cmp.CodeModule.ProcOfLine(Sline, vbext_pk_Proc) Then
    InProcedure = True
   ElseIf strTest = Cmp.CodeModule.ProcOfLine(Sline, vbext_pk_Let) Then
    InProcedure = True
   ElseIf strTest = Cmp.CodeModule.ProcOfLine(Sline, vbext_pk_Get) Then
    InProcedure = True
   ElseIf strTest = Cmp.CodeModule.ProcOfLine(Sline, vbext_pk_Set) Then
    InProcedure = True
   Else
    InProcedure = False
  End If
  On Error GoTo 0

End Function

Public Function IsAlphaIntl(ByVal sChar As String) As Boolean

  IsAlphaIntl = Not (UCase$(sChar) = LCase$(sChar))

End Function

Public Function IsNumeral(ByVal strTest As String) As Boolean

  IsNumeral = InStr("1234567890", strTest) > 0

End Function

Public Function IsPunct(ByVal strTest As String) As Boolean

  'Detect punctuation

  If IsNumeral(strTest) Then
    IsPunct = False
   Else
    IsPunct = Not IsAlphaIntl(strTest)
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:14:31 AM) 4 + 58 = 62 Lines Thanks Ulli for inspiration and lots of code.

