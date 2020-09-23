Attribute VB_Name = "mod_Useage2"
Option Explicit
Public Enum SearchRange
  WholeMod
  DecOnly
  CodeOnly
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private WholeMod, DecOnly, CodeOnly
#End If
Public Enum CallAnalysis
  CallsZero
  CallsControlOnly
  CallsInternal
  CallsClassOnly
  CallsClassForm
  CallsClassMod
  CallsClassModForm
  CallsModonly
  CallsModMod
  CallsModModForm
  CallsModForm
  CallsModModFormForm
  CallsFormOnly
  CallsFormMod
  CallsFormForm
  CallsFormFormMod
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private CallsZero, CallsInternal, CallsClassOnly, CallsModForm, CallsModMod, CallsModonly, CallsFormForm, CallsFormOnly, CallsClassMod, CallsClassForm, CallsClassModForm
#End If


Public Function EnumUseageArrayTest(strFind As String, _
                                    ByVal strOrig As String) As Variant

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim StartLine  As Long
  Dim GuardLine  As Long
  Dim L_CodeLine As String
  Dim strArr     As String
  Dim UseCount   As Long
  Dim TPos       As Long

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = Comp.CodeModule.CountOfDeclarationLines
        UseCount = 0
        'Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, L_CodeLine, StartLine)
        Do While Comp.CodeModule.Find(strFind, StartLine, 1, -1, -1, True, True, False)
          L_CodeLine = Comp.CodeModule.Lines(StartLine, 1)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          If Not SmartLeft(strOrig, L_CodeLine) Then
            'because the strORig may have been rebuilt by detector to no linecont code line
            '          UseCount = UseCount + CountCodeSubString(L_CodeLine, srtfind)
            TPos = InStrWholeWordRX(L_CodeLine, strFind)
            Do While TPos
              If InCode(L_CodeLine, TPos) Then
                UseCount = UseCount + 1
              End If
              TPos = InStrWholeWordRX(L_CodeLine, strFind, TPos + 1)
            Loop
          End If
          StartLine = StartLine + 1
          GuardLine = StartLine
        Loop
        If UseCount Then
          strArr = AccumulatorString(strArr, strInBrackets(UseCount) & Comp.Name & "[" & Comp.Type & "]")
        End If
      End If
    Next Comp
  Next Proj
  EnumUseageArrayTest = Split(strArr, ",")

End Function

'Public Function GetWholeCaseMatchCodeLine(VarProjName As Variant,
''                                          VarModName As Variant,
''                                          varFind As Variant,
''                                          strCode As String,
''                                          Optional lngStartLine As Long,
''                                          Optional SearchIn As SearchRange = WholeMod,
''                                          Optional colFind As Long = 0) As Boolean
'  Dim Comp    As VBComponent
'  Dim CompMod As CodeModule
'  Dim EndLine As Long
'  Dim SCol    As Long
''Check that a Found line is still in the code
'  If Len(varFind) Then
'  If Not Comp Is Nothing Then
'    Set Comp = VBInstance.VBProjects(VarProjName).VBComponents(VarModName)
'    Set CompMod = Comp.CodeModule
'    Select Case SearchIn
'     Case WholeMod
'      EndLine = -1
'     Case DecOnly
'      EndLine = CompMod.CountOfDeclarationLines + 1
'      If EndLine < lngStartLine Then
'        GoTo NoActionExit
'      End If
'     Case CodeOnly
'      EndLine = -1
'    End Select
'      GetWholeCaseMatchCodeLine = CompMod.Find(varFind, lngStartLine, SCol, EndLine, -1, True, True, False)
'      If GetWholeCaseMatchCodeLine Then
'        strCode = CompMod.Lines(lngStartLine, 1)
'        colFind = SCol
'       Else
'        strCode = vbNullString
'      End If
'    End If
'  End If
'NoActionExit:
'  Set Comp = Nothing
'  Set CompMod = Nothing
'End Function
Public Function ProcedureUseageArray(strFind As String) As Variant

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim StartLine  As Long
  Dim GuardLine  As Long
  Dim L_CodeLine As String
  Dim strArr     As String
  Dim UseCount   As Long
  Dim TPos       As Long

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        UseCount = 0
        Do While Comp.CodeModule.Find(strFind, StartLine, 1, -1, -1, True, True, False)
          L_CodeLine = Comp.CodeModule.Lines(StartLine, 1)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          TPos = InStrWholeWordRX(L_CodeLine, strFind)
          Do While TPos
            If InCode(L_CodeLine, TPos) Then
              If Not isProcHead(L_CodeLine) Then
                UseCount = UseCount + 1
              End If
            End If
            TPos = InStrWholeWordRX(L_CodeLine, strFind, TPos + 1)
          Loop
          StartLine = StartLine + 1
          GuardLine = StartLine
        Loop
        If UseCount Then
          strArr = AccumulatorString(strArr, strInBrackets(UseCount) & Comp.Name & "[" & Comp.Type & "]")
        End If
      End If
    Next Comp
  Next Proj
  ProcedureUseageArray = Split(strArr, ",")

End Function

Public Function VariableUseageArray(strFind As String, _
                                    ByVal strOrig As String) As Variant

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim StartLine  As Long
  Dim GuardLine  As Long
  Dim L_CodeLine As String
  Dim strArr     As String
  Dim UseCount   As Long
  Dim TPos       As Long

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        UseCount = 0
        'Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, L_CodeLine, StartLine)
        Do While Comp.CodeModule.Find(strFind, StartLine, 1, -1, -1, True, True, False)
          L_CodeLine = Comp.CodeModule.Lines(StartLine, 1)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          If Not SmartLeft(strOrig, L_CodeLine) Then
            'because the strORig may have been rebuilt by detector to no linecont code line
            '          UseCount = UseCount + CountCodeSubString(L_CodeLine, srtfind)
            TPos = InStrWholeWordRX(L_CodeLine, strFind)
            Do While TPos
              If InCode(L_CodeLine, TPos) Then
                UseCount = UseCount + 1
              End If
              TPos = InStrWholeWordRX(L_CodeLine, strFind, TPos + 1)
            Loop
          End If
          StartLine = StartLine + 1
          GuardLine = StartLine
        Loop
        If UseCount Then
          strArr = AccumulatorString(strArr, strInBrackets(UseCount) & Comp.Name & "[" & Comp.Type & "]")
        End If
      End If
    Next Comp
  Next Proj
  VariableUseageArray = Split(strArr, ",")

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:24:41 AM) 30 + 185 = 215 Lines Thanks Ulli for inspiration and lots of code.

