Attribute VB_Name = "mod_ScopeTest"
Option Explicit
Private Enum ScopeLevel
  ScopeUnused
  ScopePrivate
  ScopePublic
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ScopeUnused, ScopePrivate, ScopePublic
#End If


Public Function isProcHead(ByVal strCode As String) As Boolean

  ' protects from detecting comments

  If ExtractCode(strCode) Then
    isProcHead = InstrAtPositionSetArray(strCode, ipLeftOr2ndOr3rd, True, ArrFuncPropSub)
    If isProcHead Then
      isProcHead = LeftWord(strCode) <> "End"
    End If
  End If

End Function

Public Function ScopeGenerator(varTest As Variant, _
                               Optional ByVal bEnumTypeTest As Boolean = False) As String

  Dim TestProcHead As Boolean

  Select Case ScopeSearch(varTest, TestProcHead)
   Case ScopePublic
    ScopeGenerator = "Public "
   Case Else
    If bEnumTypeTest Then
      ScopeGenerator = IIf(TestProcHead, "Public ", "Private ")
     Else
      ScopeGenerator = "Private "
    End If
  End Select

End Function

Private Function ScopeSearch(ByVal strFind As String, _
                             Optional InProcLine As Boolean = False) As ScopeLevel

  Dim strCode        As String
  Dim CompMod        As CodeModule
  Dim Comp           As VBComponent
  Dim Proj           As VBProject
  Dim StartLine      As Long
  Dim CurCompCount   As Long
  Dim HitsArray()    As Variant
  Dim HitSum         As Long
  Dim HitArrayMember As Long
  Dim I              As Long

  If LenB(strFind) Then
    On Error Resume Next
    ScopeSearch = 0 'default value
    If LenB(strFind) > 0 Then
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If SafeCompToProcess(Comp, CurCompCount) Then
            I = I + 1
          End If
        Next Comp
      Next Proj
      CurCompCount = 0
      ReDim HitsArray(I) As Variant
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If SafeCompToProcess(Comp, CurCompCount) Then
            Set CompMod = Comp.CodeModule
            StartLine = 1 'initialize search range
            'If GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, strCode, StartLine) Then
            'Do
            Do While CompMod.Find(strFind, StartLine, 1, -1, -1, True, True)
              strCode = CompMod.Lines(StartLine, 1)
              If InCode(strCode, InStr(strCode, strFind)) Then
                'Got one
                HitsArray(HitArrayMember) = HitsArray(HitArrayMember) + 1
                If isProcHead(strCode) Then
                  'Get whole prochead
                  strCode = GetWholeLineCodeModule(CompMod, StartLine, StartLine)
                  InProcLine = InStrCode(strFind, strCode)
                End If
              End If
              StartLine = StartLine + 1
              If StartLine >= CompMod.CountOfLines Then
                Exit Do
              End If
            Loop
            ' While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, strCode, StartLine)
            'End If
          End If
          HitArrayMember = HitArrayMember + 1
        Next Comp
      Next Proj
      Set Comp = Nothing
      Set CompMod = Nothing
      Set Proj = Nothing
    End If
    'assume is public
    ScopeSearch = ScopePublic
    'test used at all
    For I = LBound(HitsArray) To UBound(HitsArray)
      HitSum = HitSum + HitsArray(I)
    Next I
    If HitSum <= 1 Then
      ScopeSearch = ScopeUnused
     Else
      'test if only a single module equals hitsum
      For I = LBound(HitsArray) To UBound(HitsArray)
        If HitSum = HitsArray(I) Then
          ScopeSearch = ScopePrivate
          Exit For
        End If
      Next I
    End If
    On Error GoTo 0
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:20:48 AM) 9 + 115 = 124 Lines Thanks Ulli for inspiration and lots of code.

