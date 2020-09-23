Attribute VB_Name = "mod_LocalFix2"
Option Explicit

Public Function ArrayMember(ByVal Tval As Variant, _
                            ParamArray pMembers() As Variant) As Boolean

  Dim wal As Variant

  'returns true if any member of pMembers equals Tval
  For Each wal In pMembers
    If Tval = wal Then
      ArrayMember = True
      Exit For 'unction
    End If
  Next wal

End Function

Public Function BuildExistingDimArray(ByVal strR As String) As Variant

  Dim arrLine         As Variant
  Dim arrDims         As Variant
  Dim K               As Long
  Dim J               As Long
  Dim L_CodeLine      As String
  Dim StrDetectedDims As String
  Dim I               As Long
  Dim lngdummy        As Long
  Dim strDummy        As String

  arrLine = Split(strR, vbNewLine)
  L_CodeLine = GetRoutineDeclaration(arrLine, lngdummy, lngdummy, strDummy)
  ExtractParameters L_CodeLine, StrDetectedDims
  For I = LBound(arrLine) To UBound(arrLine) - 1
    L_CodeLine = arrLine(I)
    'if it is a function then it will appear as a potential Dim of form 'FuncName = WhatEver
    If ExtractCode(L_CodeLine) Then
      If IsDimLine(L_CodeLine) Then
        L_CodeLine = ExpandForDetection(Mid$(L_CodeLine, GetSpacePos(L_CodeLine) + 1))
        arrDims = Split(L_CodeLine, ",")
        For J = LBound(arrDims) To UBound(arrDims)
          'Hi Fredric' the bug was that next line started ArrLine(J)=
          'fortunately ArrLine is local so the bug didn't feed back to rest of code
          arrDims(J) = Trim$(arrDims(J))
          If GetSpacePos(arrDims(J)) Then
            arrDims(J) = Left$(arrDims(J), GetSpacePos(arrDims(J)))
          End If
          StrDetectedDims = AccumulatorString(StrDetectedDims, arrDims(J), ",")
        Next J
      End If
    End If
  Next I
  If LenB(StrDetectedDims) Then
    StrDetectedDims = ReplaceArray(StrDetectedDims, ",Dim ", ",", ",Static ", ",", ",Const ", ",", " As ", SngSpace, DblSpace, SngSpace)
    arrDims = Split(StrDetectedDims, ",")
    For K = LBound(arrDims) To UBound(arrDims)
      arrDims(K) = Trim$(arrDims(K))
      If GetSpacePos(arrDims(K)) Then
        If Right$(arrDims(K), 6) = "String" Then
          arrDims(K) = Trim$(Left$(arrDims(K), GetSpacePos(arrDims(K)) - 1))
          arrDims(K) = Trim$(arrLine(K) & SngSpace & arrDims(K) & "$")
         Else
          arrDims(K) = Trim$(Left$(arrDims(K), GetSpacePos(arrDims(K))))
        End If
      End If
    Next K
  End If
  If Not IsEmpty(arrDims) Then
    For K = LBound(arrDims) To UBound(arrDims)
      'parameters with type suffixes may be present without the suffix so generate them
      If MultiRight(arrDims(K), True, "!", "@", "#", "$", "%", "&") Then
        arrDims(K) = arrDims(K) & SngSpace & Left$(arrDims(K), Len(arrDims(K)) - 1)
      End If
    Next K
    BuildExistingDimArray = QuickSortArray(Split(Join(arrDims)))
   Else
    BuildExistingDimArray = Split("")
  End If

End Function

Public Function CharReplace(varText As Variant, _
                            VarRep As Variant, _
                            ByVal StartPos As Long, _
                            Optional Range As Long = 1) As String

  CharReplace = Left$(varText, StartPos - 1) & VarRep & Mid$(varText, StartPos + Range)

End Function

Public Function CntrlDescMember(varName As Variant, _
                                Optional ByVal StartPoint As Long = 0) As Long

  Dim I As Long

  CntrlDescMember = -1 ' report control does not exist with this return
  If bCtrlDescExists Then
    For I = StartPoint To UBound(CntrlDesc)
      If CntrlDesc(I).CDName = varName Then
        CntrlDescMember = I
        Exit For
      End If
    Next I
  End If

End Function

Private Function ContainsDecimal(strTest As String) As Boolean

  ContainsDecimal = InStr(strTest, "0.") > 0

End Function

Private Sub ExtractParameters(ByVal Cline As String, _
                              strDims As String)

  Dim DelArray As Variant
  Dim strTmp   As String
  Dim FuncLoc  As Long
  Dim arrLine  As Variant
  Dim I        As Long

  If InstrAtPositionSetArray(Cline, ipLeftOr2ndOr3rd, True, ArrFuncPropSub) Then
    If ExtractCode(Cline) Then
      Cline = ConcealParameterCommas(Cline)
      If InStr(Cline, DQuote) Then
        For I = 1 To Len(Cline)
          If InLiteral(Cline, I) Then
            If Mid$(Cline, I, 1) <> DQuote Then
              Mid$(Cline, I, 1) = "~"
            End If
          End If
        Next I
        Do While InStrCode(Cline, "~~")
          Cline = CharReplace(Cline, "~", InStrCode(Cline, "~~"), 2)
        Loop
        Do While InStr(Cline, DQuote & "~" & DQuote)
          Cline = Replace$(Cline, DQuote & "~" & DQuote, EmptyString)
        Loop
      End If
      Cline = ReplaceArray(Cline, LBracket, " (", RBracket, ") ", "ByRef ", vbNullString, "ByVal ", vbNullString, "Optional ", vbNullString, EqualInCode, " =", ",", SpacePad(","), " As ", "*As*")
      arrLine = Split(Cline)
      For I = LBound(arrLine) To UBound(arrLine)
        If IsInArray(arrLine(I), ArrFuncPropSub) Then
          Select Case arrLine(I)
           Case "Function"
            strDims = AccumulatorString(strDims, arrLine(I + 1))
            arrLine(I + 1) = vbNullString
           Case "Property"
            strDims = AccumulatorString(strDims, arrLine(I + 2))
            arrLine(I + 1) = vbNullString
            arrLine(I + 2) = vbNullString
           Case "Sub"
            arrLine(I + 1) = vbNullString
          End Select
        End If
      Next I
      DelArray = Array("Friend", "Function", "Private", "Property", "Public", "Sub")
      For I = LBound(arrLine) To UBound(arrLine)
        If InQSortArray(DelArray, arrLine(I)) Then
          arrLine(I) = vbNullString
        End If
      Next I
      strTmp = Trim$(Join(arrLine))
      arrLine = Split(strTmp, ",")
      For I = LBound(arrLine) To UBound(arrLine)
        FuncLoc = InStr(arrLine(I), "=")
        If FuncLoc Then
          arrLine(I) = Left$(arrLine(I), FuncLoc - 1)
        End If
      Next I
      strTmp = Trim$(Join(arrLine))
      arrLine = Split(strTmp)
      For I = LBound(arrLine) To UBound(arrLine)
        If Left$(arrLine(I), 1) = LBracket Then
          arrLine(I) = Mid$(arrLine(I), 2)
        End If
        If Right$(arrLine(I), 1) = RBracket Then
          arrLine(I) = Left$(arrLine(I), Len(arrLine(I)) - 1)
        End If
      Next I
      strTmp = Trim$(Join(arrLine))
      arrLine = Split(strTmp)
      For I = LBound(arrLine) To UBound(arrLine)
        FuncLoc = InStr(arrLine(I), "*As*")
        If FuncLoc Then
          arrLine(I) = Left$(arrLine(I), FuncLoc - 1)
        End If
        strDims = AccumulatorString(strDims, arrLine(I))
      Next I
    End If
  End If

End Sub

Public Function FromDefType(varTest As Variant, _
                            ByVal ModuleNumber As Long) As String

  Dim str1stlet As String

  str1stlet = UCase$(Left$(varTest, 1))
  If InQSortArray(ArrQDefInt(ModuleNumber), str1stlet) Then
    FromDefType = " As Integer"
   ElseIf InQSortArray(ArrQDefLng(ModuleNumber), str1stlet) Then
    FromDefType = " As Long"
   ElseIf InQSortArray(ArrQDefDbl(ModuleNumber), str1stlet) Then
    FromDefType = " As Double"
   ElseIf InQSortArray(ArrQDefSng(ModuleNumber), str1stlet) Then
    FromDefType = " As Single"
   ElseIf InQSortArray(ArrQDefBool(ModuleNumber), str1stlet) Then
    FromDefType = " As Boolean"
   ElseIf InQSortArray(ArrQDefByte(ModuleNumber), str1stlet) Then
    FromDefType = " As Byte"
   ElseIf InQSortArray(ArrQDefCur(ModuleNumber), str1stlet) Then
    FromDefType = " As Currency"
   ElseIf InQSortArray(ArrQDefDate(ModuleNumber), str1stlet) Then
    FromDefType = " As Date"
   ElseIf InQSortArray(ArrQDefStr(ModuleNumber), str1stlet) Then
    FromDefType = " As String"
   ElseIf InQSortArray(ArrQDefObj(ModuleNumber), str1stlet) Then
    FromDefType = " As Object"
   ElseIf InQSortArray(ArrQDefVar(ModuleNumber), str1stlet) Then
    FromDefType = " As Variant"
  End If

End Function

Public Function GetDimTypeFromInternalEvidence(ByVal ModuleNumber As Long, _
                                               ArrProc As Variant, _
                                               varTest As Variant, _
                                               ByVal CompName As String) As String

  
  Dim strRecursive     As String
  Dim I                As Long
  Dim strTmp           As String
  Dim ExistingDimArray As Variant
  Dim ArrLine2         As Variant
  Dim K                As Long
  Dim X                As Long
  Dim arrLine          As Variant
  Dim EqPos            As Long

  'Additions Detect if routine parameters are feed to variable, test if prog functions are called and have type
  If varTest <> TypeUpdate(varTest) Then
    GetDimTypeFromInternalEvidence = "<-1" & Mid$(TypeUpdate(varTest), Len(varTest))
   Else
    If UsingDefTypes Then
      GetDimTypeFromInternalEvidence = FromDefType(varTest, ModuleNumber)
     Else
      'test for paramater hints
      For I = LBound(ArrProc) + 1 To UBound(ArrProc) - 1
        If Not JustACommentOrBlankOrDimLine(ArrProc(I)) Then
          If InstrAtPosition(ExpandForDetection(ArrProc(I)), varTest, ipAny) Then
            'test for 'For Each X in Y' the For Each variable MUST be variant
            If InstrAtPosition(ExpandForDetection(ArrProc(I)), varTest, ip3rd) Then
              If InstrAtPosition(ExpandForDetection(ArrProc(I)), "Each", ip2nd) Then
                GetDimTypeFromInternalEvidence = " As Variant"
                Exit Function
              End If
            End If
            'Tests for varTest being 1st or 2nd(allows 'Set varTest' and 'For  varTest') position
            'v2.6.4 changed to ipLeftOr2nd
            If InstrAtPosition(ExpandForDetection(ArrProc(I)), varTest, ipLeftOr2nd) Then
              ', ip2nd) Then
              arrLine = Split(ExpandForDetection(ArrProc(I)))
collapsedBrackets:
              If UBound(arrLine) > 1 Then
                Select Case arrLine(1)
                 Case "="
                  If LenB(InternalEvidence(arrLine(2))) Then
                    If arrLine(2) <> varTest Then
                      'don't test if it is format x = x + y
                      'if X is a function name as a variable inside itself
                      ' this would generate an eternal loop
                      GetDimTypeFromInternalEvidence = InternalEvidence(arrLine(2))
                    End If
                    Exit For
                   ElseIf IsDeclaration(CStr(arrLine(2)), "Private", , CompName) Then
                    strTmp = GetDeclarationType(CStr(arrLine(2)), "Private", , CompName)
                    If LenB(strTmp) Then
                      GetDimTypeFromInternalEvidence = " As " & strTmp
                      Exit For
                    End If
                   ElseIf UBound(arrLine) > 2 Then
                    If LenB(InternalEvidence(arrLine(3))) Then
                      GetDimTypeFromInternalEvidence = InternalEvidence(arrLine(3))
                      Exit For
                    End If
                    If IsDeclaration(CStr(arrLine(2)), "Public") Then
                      strTmp = GetDeclarationType(CStr(arrLine(2)), "Public")
                      If LenB(strTmp) Then
                        GetDimTypeFromInternalEvidence = " As " & strTmp
                        Exit For
                      End If
                      If LenB(InternalEvidence(arrLine(3))) Then
                        GetDimTypeFromInternalEvidence = InternalEvidence(arrLine(3))
                        Exit For
                      End If
                     ElseIf ContainsDecimal(JoinPartial(arrLine, 3, UBound(arrLine))) Then
                      GetDimTypeFromInternalEvidence = " As Double"
                     ElseIf InstrArrayHard(JoinPartial(arrLine, 3, UBound(arrLine)), arrPatternDetector2) Then
                      GetDimTypeFromInternalEvidence = " As Single"
                      Exit For
                     ElseIf InStrWholeWord(JoinPartial(arrLine, 3, UBound(arrLine)), "Mod") Then
                      GetDimTypeFromInternalEvidence = " As Long"
                      Exit For
                    End If
                   Else
                    'ver 1.0.99
                    'this gets a Type from internal Dims(and Consts)
                    'if there is a code line of format 'VarTest = Dim(or Const) '
                    ' as long as the Dim(Const) has a Type
                    ExistingDimArray = BuildExistingDimArray(Join(ArrProc, vbNewLine))
                    If Not InQSortArray(ExistingDimArray, arrLine(2)) Then
                      For K = 0 To UBound(ExistingDimArray)
                        If InStr(ArrProc(K), arrLine(2)) And MultiLeft(ArrProc(K), True, "Dim ", "Const ") Then
                          If Get_As_Pos(ArrProc(K)) > 0 Then
                            GetDimTypeFromInternalEvidence = GetType(ArrProc(K))
                            'ver1.1.41   becuase the dim lines aren't yet expanded this stop multimembers of dim being assignned
                            'rare only happens if the only way to access a Type is to read from dim lines
                            If InStr(GetDimTypeFromInternalEvidence, CommaSpace) Then
                              GetDimTypeFromInternalEvidence = Left$(GetDimTypeFromInternalEvidence, InStr(GetDimTypeFromInternalEvidence, CommaSpace) - 1)
                            End If
                            EqPos = InStr(GetDimTypeFromInternalEvidence, EqualInCode)
                            If EqPos Then
                              'remove the Value assigned to Constant
                              GetDimTypeFromInternalEvidence = Left$(GetDimTypeFromInternalEvidence, EqPos - 1)
                              Exit For
                            End If
                          End If
                        End If
                      Next K
                    End If
                  End If
                 Case varTest
                  If InstrArrayHard(JoinPartial(arrLine, 3, UBound(arrLine)), arrPatternDetector2) Then
                    GetDimTypeFromInternalEvidence = " As Single"
                    Exit For
                   ElseIf InStrWholeWord(JoinPartial(arrLine, 3, UBound(arrLine)), "Mod") Then
                    GetDimTypeFromInternalEvidence = " As Long"
                    Exit For
                   ElseIf UBound(arrLine) > 2 Then
                    If arrLine(2) <> LBracket Then
                      If LenB(InternalEvidence(arrLine(3))) Then
                        GetDimTypeFromInternalEvidence = InternalEvidence(arrLine(3))
                        Exit For
                      End If
                    End If
                   ElseIf InstrAtPosition(ArrProc(I), varTest, IpRight) Then
                    arrLine = Split(ArrProc(I))
                    If IsDeclaration(CStr(arrLine(0)), "Private", , CompName) Then
                      GetDimTypeFromInternalEvidence = " As " & GetDeclarationType(CStr(arrLine(0)), "Private", , CompName)
                      Exit For
                    End If
                   ElseIf IsDeclaration(CStr(arrLine(0)), "Public") Then
                    GetDimTypeFromInternalEvidence = " As " & GetDeclarationType(CStr(arrLine(0)), "Public")
                    Exit For
                   ElseIf InstrArray(JoinPartial(arrLine, 0, 1), DQuote, "$", " & ") Then
                    GetDimTypeFromInternalEvidence = " As String"
                   ElseIf arrLine(0) = "Set" Then
                    GetDimTypeFromInternalEvidence = " As Variant' Could be As Object"
                   ElseIf arrLine(0) = "For" Then
                    strRecursive = GetDimTypeFromInternalEvidence(ModuleNumber, ArrProc, arrLine(3), CompName)
                    If Len(strRecursive) Then
                      GetDimTypeFromInternalEvidence = strRecursive  'Variant' Could be As Object"
                      Exit For
                    End If
                   ElseIf InstrArray(JoinPartial(arrLine, 3, UBound(arrLine)), DQuote, "$", " & ") Then
                    GetDimTypeFromInternalEvidence = " As String"
                    Exit For
                  End If
                 Case LBracket
                  ArrLine2 = arrLine
                  For X = LBound(ArrLine2) To UBound(ArrLine2)
                    arrLine(X) = vbNullString
                    ArrLine2(0) = ArrLine2(0) & ArrLine2(X)
                    If ArrLine2(X) = RBracket Then
                      ArrLine2(X) = vbNullString
                      Exit For
                    End If
                    If X > 0 Then
                      ArrLine2(X) = vbNullString
                    End If
                  Next X
                  arrLine = CleanArray(ArrLine2)
                  GoTo collapsedBrackets
                End Select
              End If
            End If
            If InstrAtPosition(ExpandForDetection(ArrProc(I)), varTest, IpRight) Then
              If Left$(ArrProc(I), 12) = "Line Input #" Then
                GetDimTypeFromInternalEvidence = " As String' Possibly: As Variant"
                Exit For
              End If
            End If
            If InstrAtPosition(ExpandForDetection(ArrProc(I)), varTest, ipAny) Then
              If Left$(ArrProc(I), 7) = "Input #" Then
                GetDimTypeFromInternalEvidence = " As String' Possibly: As Variant: As Long"
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  If LenB(GetDimTypeFromInternalEvidence) Then
    If Right$(GetDimTypeFromInternalEvidence, 2) = "()" Then
      'v3.0.2 improved test
      GetDimTypeFromInternalEvidence = Left$(GetDimTypeFromInternalEvidence, Len(GetDimTypeFromInternalEvidence) - 2)
    End If
  End If

End Function

Public Function GetProcClassStr(ByVal strCode As String) As String

  Dim arrTest As Variant

  arrTest = Array("Public", "Private", "Friend", "Static", "Set", "Let", "Get")
  If isProcHead(strCode) Then
    strCode = ExpandForDetection(strCode)
    Do While IsInArray(LeftWord(strCode), arrTest)
      strCode = Trim$(Mid$(strCode, Len(LeftWord(strCode)) + 1))
    Loop
    GetProcClassStr = LeftWord(Trim$(strCode))
  End If

End Function

Public Function GetProcNameStr(ByVal strCode As String) As String

  Dim arrTmp  As Variant
  Dim strTmp  As String
  Dim arrTest As Variant

  'v3.0.4 switched to faster IsInArray from ArrayMember
  arrTest = Array("Public", "Private", "Friend", "Static", "Sub", "Function", "Property", "Set", "Let", "Get")
  arrTmp = Split(strCode, vbNewLine)
  strTmp = ExpandForDetection(arrTmp(GetProcCodeLineOfRoutine(arrTmp)))
  Do Until IsInArray(LeftWord(strTmp), arrTest)
    strTmp = Trim$(Mid$(strTmp, Len(LeftWord(strTmp)) + 1))
    If LenB(strTmp) = 0 Then
      Exit Do
    End If
  Loop
  If LenB(strCode) Then
    Do While IsInArray(LeftWord(strTmp), arrTest)
      strTmp = Trim$(Mid$(strTmp, Len(LeftWord(strTmp)) + 1))
      If LenB(strTmp) = 0 Then
        Exit Do
      End If
    Loop
    GetProcNameStr = LeftWord(strTmp)
  End If

End Function

Public Function InArrayinSomeFormat(ByVal varTest As Variant, _
                                    ByVal SomeArray As Variant) As Boolean

  InArrayinSomeFormat = IsInArray(varTest, SomeArray)
  If Not InArrayinSomeFormat Then
    'test for type suffix in code and As Type in Dims
    If TypeSuffixExists(varTest) Then
      InArrayinSomeFormat = IsInArray(Left$(varTest, Len(varTest) - 1), SomeArray)
    End If
  End If

End Function

Public Function IsDimLine(ByVal varTest As Variant, _
                          Optional ByVal bHard As Boolean = True, _
                          Optional strProcHead As String) As Boolean

  Dim strRedim As String

  'tests for Routine level declarations (Dim,Static,Const,ReDim)
  If varTest <> strProcHead Then
    If bHard Then
      IsDimLine = MultiLeft(varTest, True, "Dim ", "Static ", "Const ", "ReDim ")
      If IsDimLine Then
        If Left$(varTest, 6) = "ReDim " Then
          strRedim = WordInString(ExpandForDetection(varTest), 2)
          If strRedim = "Preserve" Or IsDeclaration(strRedim) Or Left$(strRedim, 1) = "." Then
            IsDimLine = False
           ElseIf Len(strProcHead) Then
            If ContainsWholeWord(strProcHead, strRedim) Then
              IsDimLine = False
            End If
          End If
        End If
        'v3.0.3 for Const whose value is set from other constants
        If Left$(varTest, 6) = "Const " Then
          If InStr(varTest, " Or ") > InStr(varTest, " = ") Then
            IsDimLine = False
           ElseIf InStr(varTest, " And ") > InStr(varTest, " = ") Then
            IsDimLine = False
          End If
        End If
      End If
     Else
      IsDimLine = MultiLeft(varTest, True, "Dim ", "Static ", "Const ")
    End If
  End If

End Function

Public Function IsMenu(ByVal varName As Variant) As Boolean

  Dim I        As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDName = varName Then
        IsMenu = CntrlDesc(I).CDClass = "Menu"
        Exit For 'unction
      End If
    Next I
  End If

End Function

Public Function JoinPartial(ByVal tarr As Variant, _
                            ByVal LoMember As Long, _
                            ByVal HiMember As Long, _
                            Optional Delimiter As String = SngSpace) As String

  Dim I As Long

  'join section of array from member LoMember to HiMember
  'can cope with actual count is beyond the parameters by skipping missing members
  For I = LBound(tarr) To UBound(tarr)
    If I >= LoMember Then
      If I <= HiMember Then
        If LenB(JoinPartial) Then
          JoinPartial = JoinPartial & Delimiter & tarr(I)
         Else
          JoinPartial = tarr(I)
        End If
      End If
    End If
  Next I

End Function

Public Function JustACommentOrBlankOrDimLine(ByVal varSearch As Variant, _
                                             Optional bHard As Boolean = True) As Boolean

  'copright 2003 Roger Gilchrist
  'detect comments and empty strings

  JustACommentOrBlankOrDimLine = JustACommentOrBlank(varSearch)
  If Not JustACommentOrBlankOrDimLine Then
    JustACommentOrBlankOrDimLine = IsDimLine(varSearch, bHard)
  End If

End Function

Public Function ModDescMember(ByVal varName As String) As Long

  Dim I As Long

  ModDescMember = -1 ' report control does not exist with this return
  If Len(varName) Then
    If bModDescExists Then
      For I = 1 To UBound(ModDesc)
        If ModDesc(I).MDName = varName Then
          ModDescMember = I
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function ProcDescMember(varName As Variant, _
                               Optional ByVal strModule As String, _
                               Optional ByVal StartPoint As Long = 0) As Long

  Dim I As Long

  ProcDescMember = -1 ' report control does not exist with this return
  If bProcDescExists Then
    For I = StartPoint To UBound(PRocDesc)
      If LenB(strModule) = 0 Or strModule = PRocDesc(I).PrdComp Then
        If PRocDesc(I).PrDName = varName Then
          ProcDescMember = I
          Exit For
        End If
      End If
    Next I
  End If

End Function

Public Sub Safe_AsTypeAdd(varCode As Variant, _
                          strAsType As String)

  Dim strComnt As String

  If LenB(strAsType) Then
    If InStr(varCode, "Const ") Then
      'v2.2.5 thanks Mike Ulik
      'this fixes Insert Type when all fixes are set to off
      '(no you can't stop it it is good for you)
      varCode = Replace$(varCode, " = ", strAsType & " = ")
     Else
      If Left$(strAsType, 3) = "<-1" Then
        varCode = Left$(varCode, Len(varCode) - 1)
        strAsType = Mid$(strAsType, 4)
      End If
      'v2.3.6 added mainly for Redim with trailing comments
      ExtractCode varCode, strComnt
      varCode = varCode & strAsType & strComnt
    End If
  End If

End Sub

Public Function SmartMarker(varCode As Variant, _
                            ByVal TargetLine As Long, _
                            Msg As String, _
                            MPos As MarkerPos) As String

  'Generic message structures
  'place message in specified MPos
  'VarCode = is the array of a piece of code
  'TargetLine = the line of the array code you are fixing
  'Msg anything you want to tack on to the code
  'MPos where to join it
  '*If MEoL or mAfter and a Line Continuation character is on targetline
  'then return (&fix if necessary) current line and insert message
  'at first acceptable point

  Msg = CleanMsg(Msg)
  SmartMarker = varCode(TargetLine)
  Select Case MPos
   Case MBefore
    If TargetLine - 1 > 0 Then
      If HasLineCont(varCode(TargetLine - 1)) Then
        Do
          TargetLine = TargetLine - 1
        Loop While HasLineCont(varCode(TargetLine)) Or TargetLine = 0
        varCode(TargetLine) = Msg & vbNewLine & varCode(TargetLine)
       Else
        SmartMarker = Msg & vbNewLine & SmartMarker
      End If
     Else
      SmartMarker = Msg & vbNewLine & SmartMarker
    End If
   Case MSoL
    If HasLineCont(SmartMarker) Then
      TargetLine = GetSafeInsertLineArray(varCode, TargetLine)
      varCode(TargetLine) = Msg & varCode(TargetLine)
     Else
      SmartMarker = Msg & SmartMarker
    End If
   Case MAfter
    If HasLineCont(SmartMarker) Then
      TargetLine = GetSafeInsertLineArray(varCode, TargetLine)
      varCode(TargetLine) = varCode(TargetLine) & vbNewLine & Msg
     Else
      SmartMarker = SmartMarker & vbNewLine & Msg
    End If
   Case MEoL
    If HasLineCont(SmartMarker) Then
      TargetLine = GetSafeInsertLineArray(varCode, TargetLine)
      varCode(TargetLine) = varCode(TargetLine) & vbNewLine & Msg
     Else
      SmartMarker = SmartMarker & Msg
    End If
  End Select

End Function

Public Function TypeUpdate(varOld As Variant, _
                           Optional WhichFix As Long, _
                           Optional bIncludeAs As Boolean = True) As String

  Dim TPos As Long

  'ver 1.1.29 now copes with type suffixed array dims ('Dim H&(200)')
  TPos = TypeSuffixArrayPosition(varOld)
  If TPos > -1 Then
    WhichFix = TPos
    TypeUpdate = Left$(varOld, Len(varOld) - 1) & IIf(bIncludeAs, " As ", vbNullString) & AsTypeArray(TPos)
   Else
    If InStr(varOld, LBracket) > 1 Then
      TPos = QSortArrayPos(TypeSuffixArray, Mid$(varOld, InStr(varOld, LBracket) - 1, 1))
      If TPos > -1 Then
        WhichFix = TPos
        TypeUpdate = Replace$(varOld, Mid$(varOld, InStr(varOld, LBracket) - 1, 1), vbNullString) & IIf(bIncludeAs, " As ", vbNullString) & AsTypeArray(TPos)
       Else
        TypeUpdate = varOld
      End If
     Else
      TypeUpdate = varOld
    End If
  End If

End Function

Public Function WordAfter(ByVal varSource As Variant, _
                          ByVal StrPrev As String) As String

  Dim arrTmp As Variant

  If Len(varSource) Then
    If Len(StrPrev) Then
      If InStrCode(varSource, StrPrev) Then
        If ExtractCode(varSource) Then
          varSource = StripDoubleSpace(varSource)
          arrTmp = Split(varSource, StrPrev)
          If UBound(arrTmp) > 0 Then
            arrTmp = Split(Trim$(arrTmp(1)), SngSpace)
            If UBound(arrTmp) > -1 Then
              WordAfter = arrTmp(0)
            End If
          End If
        End If
      End If
    End If
  End If

End Function

Public Function WordBefore(ByVal varSource As Variant, _
                           ByVal strNext As String) As String

  Dim arrTmp As Variant

  If Len(varSource) Then
    If Len(strNext) Then
      If InStrCode(varSource, strNext) Then
        If ExtractCode(varSource) Then
          varSource = StripDoubleSpace(varSource)
          arrTmp = Split(varSource, strNext)
          If UBound(arrTmp) > 0 Then
            arrTmp = Split(Trim$(arrTmp(0)), SngSpace)
            If UBound(arrTmp) > -1 Then
              WordBefore = arrTmp(UBound(arrTmp))
            End If
          End If
        End If
      End If
    End If
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:15:22 AM) 1 + 731 = 732 Lines Thanks Ulli for inspiration and lots of code.

