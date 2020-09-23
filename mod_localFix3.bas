Attribute VB_Name = "mod_localFix3"
Option Explicit
Public Enum Dim1Use
  Dim1ForUnmarkedNext
  Dim1LocalClass
  Dim1UDT
  Dim1Dummy
  Dim1DummyAssign
  Dim1Unused
  Dim1Const
  Dim1DummyInputOutPut
  Dim2Caser
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Dim1ForUnmarkedNext, Dim1LocalClass, Dim1UDT, Dim1Dummy, Dim1DummyAssign, Dim1Unused, Dim1Const, Dim1DummyInputOutPut, Dim2Caser
#End If
Public Enum Dim2Use
  DimNoAction
  Dim2InLine
  Dim2NextCounter
  Dim2NextOuterCounterInnerDriver
  Dim2SetOnceLoopDriver
  Dim2InLoopLoopDriver
  Dim2Swapper
  Dim2LoopBreak
  Dim2TmpVarForReturn
  Dim2WithHead1
  Dim2WithHead2
  Dim2AssignIfTest
  Dim2IfTest
  Dim2LocalClass
  Dim2UDT
  Dim2UDTSetPass
  Dim2UDTParamPass
  Dim2Proc2Proc
  Dim2ProcForCalc
  Dim2LocalConst
  Dim2PreserveValue
  Dim2Array
  Dim2InputLoading
  Dim2ParamChoice
  Dim2ChangeableForRange
  Dim2AssignParam
  Dim2OnceOnlyTest
  Dim2dummyTwiceinaLine
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private DimNoAction, Dim2InLine, Dim2NextCounter, Dim2NextOuterCounterInnerDriver, Dim2SetOnceLoopDriver
Private Dim2InLoopLoopDriver, Dim2Swapper, Dim2LoopBreak, Dim2TmpVarForReturn, Dim2WithHead1
Private Dim2WithHead2, Dim2AssignIfTest, Dim2IfTest, Dim2LocalClass, Dim2UDT, Dim2UDTSetPass
Private Dim2UDTParamPass, Dim2Proc2Proc, Dim2ProcForCalc, Dim2LocalConst, Dim2PreserveValue
Private Dim2Array, Dim2InputLoading, Dim2ParamChoice, Dim2ChangeableForRange, Dim2AssignParam
Private Dim2OnceOnlyTest, Dim2dummyTwiceinaLine
#End If


Private Function CountDimUsage(ByVal strCode As String, _
                               adim As Variant) As Long

  Dim I      As Long
  Dim J      As Long
  Dim arrTmp As Variant

  arrTmp = Split(strCode)
  For I = LBound(adim) To UBound(adim)
    For J = LBound(arrTmp) To UBound(arrTmp)
      If arrTmp(J) = adim(I) Then
        CountDimUsage = CountDimUsage + 1
       ElseIf Right$(adim(I), 1) = "." Then
        ' catch dimmed as Type variations
        If SmartLeft(arrTmp(J), adim(I)) Then
          CountDimUsage = CountDimUsage + 1
        End If
      End If
    Next J
  Next I

End Function

Public Function DimUsage(ByVal ModuleNumber As Long, _
                         ByVal strA As String) As String

  
  Dim I                    As Long
  Dim J                    As Long
  Dim ArrProc              As Variant
  Dim L_CodeLine           As String
  Dim MaxFactor            As Long
  Dim strType              As String
  Dim arrDim               As Variant
  Dim UsageCount           As Long
  Dim lineUsage            As Long
  Dim strDimComment        As String
  Dim strProcHead          As String
  Const strNoActionSuggest As String = UsageSign & "Dim Usage:(no action required)"
  Dim strDimUsage          As String

  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, DetectDimUnused) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        If Not JustACommentOrBlank(ArrProc(I)) Then
          L_CodeLine = ExpandForDetection(ArrProc(I))
          If isProcHead(L_CodeLine) Then
            strProcHead = L_CodeLine
          End If
          If IsDimLine(L_CodeLine, , strProcHead) Then
            UsageCount = 0
            lineUsage = 0
            arrDim = GenerateVariableTypeArray(L_CodeLine, strType)
            For J = LBound(ArrProc) + 1 To UBound(ArrProc) - 1
              If J <> I Then
                If Not JustACommentOrBlankOrDimLine(ArrProc(J), True) Then
                  L_CodeLine = ArrProc(J)
                  If ExtractCode(L_CodeLine) Then
                    L_CodeLine = ExpandForDetection(L_CodeLine)
                    If InstrAtPositionSetArray(L_CodeLine, ipAny, True, arrDim) Then
                      UsageCount = UsageCount + CountDimUsage(L_CodeLine, arrDim)
                      lineUsage = lineUsage + 1
                    End If
                  End If
                End If
              End If
            Next J
            strDimUsage = UsageSign & "Dim Usage:(" & UsageCount & ") on (" & lineUsage & ") lines"
            If Xcheck(XUsageComments) Then
              ArrProc(I) = ArrProc(I) & strDimUsage
            End If
            On Error Resume Next
            strDimComment = vbNullString
            Select Case UsageCount
             Case 0
              strDimComment = vbNewLine & WARNING_MSG & "Unused Dim"
             Case 1
              Select Case OneUseMode(arrDim, ArrProc, strType)
               Case Dim1Unused
                strDimComment = vbNewLine & _
                 WARNING_MSG & "(1) Variable is only used once. If Constant use explicit value; Numeric Type use '0'; String use """
               Case Dim1ForUnmarkedNext
                strDimComment = vbNewLine & _
                 SUGGESTION_MSG & "(1) For Counter (add counter variable to the 'Next' statement)."
               Case Dim1LocalClass
                strDimComment = strNoActionSuggest & "(1) Setting a Class for procedure only usage."
               Case Dim1UDT
                If strType = "New" Then
                  strDimComment = strNoActionSuggest & "(1) Local Class instance."
                 Else
                  strDimComment = strNoActionSuggest & "(1) User Defined Type/Object."
                End If
               Case Dim1Dummy
                strDimComment = strNoActionSuggest & "(1) Dummy for unneeded parameter to another procedure"
               Case Dim1DummyAssign
                strDimComment = strNoActionSuggest & "(1) Dummy variable for a Function call, allows procedure to capture one of Function paramters."
               Case Dim1Const
                strDimComment = strNoActionSuggest & "(1) Local Const could made in-line."
               Case Dim1DummyInputOutPut
                strDimComment = strNoActionSuggest & "(1) Local dummy for in/output of blank line."
               Case Dim2Caser
                strDimComment = strNoActionSuggest & "(1) Local Const for Case test. Could be replaced by value."
              End Select
             Case 2
              Select Case TwoUseMode(arrDim, ArrProc, strType)
               Case Dim2InLine
                strDimComment = vbNewLine & _
                 SUGGESTION_MSG & "(2) Replace with in-line call to calculation/function which assigns it a value."
               Case Dim2NextCounter
                strDimComment = strNoActionSuggest & "(2) For Counter."
               Case Dim2NextOuterCounterInnerDriver
                strDimComment = strNoActionSuggest & "(2) For Counter and Start/End Value."
               Case Dim2SetOnceLoopDriver
                strDimComment = strNoActionSuggest & "(2) Used to avoid re-calculation inside a loop."
               Case Dim2InLoopLoopDriver
                'ArrProc(I) = ArrProc(I) & strNoActionSuggest & "(2) In loop test to break out."
                '^^This style forces the message to display
                strDimComment = strNoActionSuggest & "(2) In loop test to break out."
               Case Dim2Swapper
                strDimComment = strNoActionSuggest & "(2) Used to Swap 2 variable values."
               Case Dim2LoopBreak
                strDimComment = strNoActionSuggest & "(2) Used to break out of Loop"
               Case Dim2TmpVarForReturn
                strDimComment = vbNewLine & _
                 SUGGESTION_MSG & "(2) Temporary Variable used to set Function/Property Return. Replace with direct assignment of Function/Property Return"
               Case Dim2WithHead1, Dim2WithHead2
                strDimComment = strNoActionSuggest & "(2) Using 'With' structure to set UDT member(s)"
               Case Dim2AssignIfTest
                strDimComment = strNoActionSuggest & "(2) Sets a value then tests for further actions."
               Case Dim2IfTest
                strDimComment = vbNewLine & _
                 SUGGESTION_MSG & "(2) Replace with in-line call to calculation/function which assigns it a value."
               Case Dim2LocalClass
                strDimComment = strNoActionSuggest & "(2) Setting a Class for procedure only usage."
               Case Dim2UDT
                strDimComment = strNoActionSuggest & "(2) User Defined Type."
               Case Dim2UDTSetPass
                strDimComment = strNoActionSuggest & "(2) User Defined Type members set and passed to another procedure."
               Case Dim2UDTParamPass
                strDimComment = strNoActionSuggest & "(2) User Defined Type passed from procedure to procedure."
               Case Dim2Proc2Proc
                strDimComment = strNoActionSuggest & "(2) Used to pass values between called procedures."
               Case Dim2ProcForCalc
                strDimComment = strNoActionSuggest & "(2) Get value from procedure for later calculation."
               Case Dim2LocalConst
                strDimComment = strNoActionSuggest & "(2) Local Const could made in-line but a Const allows rapid modifications."
               Case Dim2PreserveValue
                strDimComment = strNoActionSuggest & "(2) Preserves a value from in-procedure changes."
               Case Dim2Array
                strDimComment = strNoActionSuggest & "(2) Local Array probably passed to API call or Load data."
               Case Dim2InputLoading
                strDimComment = strNoActionSuggest & "(2) Used to load data from file."
               Case Dim2ParamChoice
                strDimComment = strNoActionSuggest & "(2) Created by another function then used in later choices."
               Case Dim2ChangeableForRange
                strDimComment = strNoActionSuggest & "(2) Creates a changeable value to change range of a For structure."
               Case Dim2AssignParam
                strDimComment = strNoActionSuggest & "(2) Set a parameter which would otherwise be 0/False."
               Case Dim2OnceOnlyTest
                strDimComment = strNoActionSuggest & "(2) controls an internal test usually only hitting once per call."
               Case Dim2dummyTwiceinaLine
                strDimComment = strNoActionSuggest & "(2) Dummy variable for a Function call, allows procedure to capture one of Function paramters."
               Case Else
                strDimComment = strNoActionSuggest & "(2) Code Fixer could not determine usage of this variable"
              End Select
              'Case 3
              '
            End Select
            If Len(strDimComment) Then
              If InStr(strDimComment, "(no action required)") = 0 Or Xcheck(XUsageComments) Then
                ArrProc(I) = ArrProc(I) & strDimComment
              End If
            End If
            If UsageCount > lineUsage * 3 Then
              If Not Xcheck(XUsageComments) Then
                ArrProc(I) = ArrProc(I) & strDimUsage
              End If
              ArrProc(I) = ArrProc(I) & vbNewLine & _
               WARNING_MSG & "Unusual ratio of usage(" & UsageCount & ") to lines(" & lineUsage & ") for this variable."
            End If
            On Error GoTo 0
          End If
        End If
      Next I
    End If
  End If
  DimUsage = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Sub DimUseDetectors(VarLine As Variant, _
                            varDim As Variant, _
                            ByVal strPName As String, _
                            StrPClass As String, _
                            bAssign As Boolean, _
                            bCalc As Boolean, _
                            bForCount As Boolean, _
                            bForDrive As Boolean, _
                            bLoopBreaker As Boolean, _
                            bParam As Boolean, _
                            bNext As Boolean, _
                            bTmpReturn As Boolean, _
                            bWithHead As Boolean, _
                            bIfTest As Boolean, _
                            bInput As Boolean, _
                            bOutput As Boolean, _
                            bCaseSel As Boolean)

  Dim EqPos As Long

  On Error GoTo oops
  EqPos = InStr(VarLine, EqualInCode)
  If EqPos Then
    bAssign = SmartLeft(VarLine, varDim & EqualInCode)
    bParam = IsProcedure(WordAfter(VarLine, "="))
    If Not bParam Then
      bCalc = InStrWholeWordRX(VarLine, varDim, EqPos)
    End If
  End If
  bInput = ContainsWholeWord(VarLine, "Input")
  bOutput = ContainsWholeWord(VarLine, "Print") Or ContainsWholeWord(VarLine, "Write")
  bIfTest = SmartLeft(VarLine, "If " & varDim & EqualInCode) Or VarLine Like "If * =*" & varDim & "* Then*" Or VarLine Like "If *" & varDim & "* Then*"
  bForCount = SmartLeft(VarLine, "For " & varDim & EqualInCode)
  bForDrive = VarLine Like "For * =*" & varDim & "*"
  bLoopBreaker = VarLine Like "Do*" & varDim & "*" Or VarLine Like "Loop *" & varDim & "*"
  bNext = VarLine Like "Next " & varDim & "*"
  If EqPos = 0 Then
    bParam = IsProcedure(LeftWord(VarLine))
    'Not (bAssign Or bCalc Or bForCount Or bForDrive Or bLoopBreaker Or bNext)
    If Not bParam Then
      bParam = IsDeclareName(LeftWord(ExpandForDetection(VarLine)))
      'IsDeclaration(LeftWord(ExpandForDetection(VarLine)))
    End If
  End If
  If Not bParam Then
    bParam = IsDeclareName(WordAfter(ExpandForDetection(VarLine), "="))
    ' IsDeclaration(WordAfter(ExpandForDetection(VarLine), "="))
  End If
  If Not bParam Then
    bParam = IsProcedure(WordAfter(ExpandForDetection(VarLine), "="))
  End If
  If Not bParam Then
    'v2.6.9 test for string used to assign to navigate property
    bParam = SmartRight(WordBefore(ExpandForDetection(VarLine), varDim), "Navigate")
  End If
  If bParam And bAssign Then
    bAssign = False
  End If
  If ArrayMember(StrPClass, "Function", "Property") Then
    bTmpReturn = VarLine Like strPName & EqualInCode & varDim
  End If
  bCaseSel = VarLine Like "Select Case " & varDim & "*"
  bWithHead = VarLine Like "With " & varDim & "*"
  bIfTest = SmartLeft(VarLine, "If " & varDim & EqualInCode) Or VarLine Like "If * =*" & varDim & "* Then*" Or VarLine Like "If " & varDim & "* Then*"
  On Error GoTo 0

Exit Sub

oops:
  MsgBox Err.Number & SngSpace & Err.Description

End Sub

Private Function GenerateHardArray(ParamArray Soft() As Variant) As Variant

  GenerateHardArray = StripDuplicateArray(Split(Join(Soft, ","), ","))

End Function

Public Function GenerateVariableTypeArray(L_CodeLine As String, _
                                          strType As String) As Variant

  Dim arrTmp As Variant
  Dim tmpV1  As String
  Dim tmpV2  As String
  Dim tmpV3  As String
  Dim TPos   As Long

  arrTmp = Split(Trim$(L_CodeLine))
  tmpV1 = CStr(arrTmp(1))
  If tmpV1 = "Preserve" Then
    tmpV1 = CStr(arrTmp(2))
  End If
  tmpV2 = tmpV1
  tmpV3 = tmpV1
  If GetLeftBracketPos(tmpV1) Then
    tmpV1 = Left$(tmpV1, GetLeftBracketPos(tmpV1) - 1)
    tmpV3 = Left$(tmpV1, GetLeftBracketPos(tmpV1))
  End If
  If UBound(arrTmp) > 1 Then
    If Has_AS(L_CodeLine) Then
      strType = GetType(L_CodeLine)
      TPos = ArrayPos(arrTmp(ArrayPos("As", arrTmp) + 1), AsTypeArray)
      If TPos > -1 Then
        tmpV2 = tmpV1 & TypeSuffixArray(TPos)
      End If
    End If
  End If
  GenerateVariableTypeArray = GenerateHardArray(tmpV1, tmpV2, tmpV3, "#" & tmpV1, tmpV1 & ".")

End Function

Private Function isClassName(ByVal strTest As String) As Boolean

  Dim I As Long

  If bModDescExists Then
    For I = LBound(ModDesc) To UBound(ModDesc)
      If ModDesc(I).MDName = strTest Then
        isClassName = ModDesc(I).MDType = "Class"
        Exit For
      End If
    Next I
  End If

End Function

Private Function OneUseMode(adim As Variant, _
                            Acode As Variant, _
                            strType As String) As Dim1Use

  Dim aDimUsed       As String
  Dim L0Assign       As Boolean
  Dim L0Calc         As Boolean
  Dim L0ForDrive     As Boolean
  Dim L0ForCount     As Boolean
  Dim L0LoopBreaker  As Boolean
  Dim L0Param        As Boolean
  Dim L0Next         As Boolean
  Dim L0TmpReturn    As Boolean
  Dim L0WithHead     As Boolean
  Dim L0IfTest       As Boolean
  Dim I              As Long
  Dim lngdummy       As Long
  Dim L0Input        As Boolean
  Dim L0Output       As Boolean
  Dim L0CaseSelector As Boolean
  Dim strPName       As String
  Dim StrProcClass   As String
  Dim aline          As String

  For I = LBound(Acode) To UBound(Acode)
    If Not JustACommentOrBlankOrDimLine(Acode(I)) Then
      If isProcHead(Acode(I)) Then
        strPName = GetProcNameStr(GetWholeLineArray(Acode, I, lngdummy))
        StrProcClass = GetProcClassStr(GetWholeLineArray(Acode, I, lngdummy))
      End If
      If Not IsDimLine(Acode(I)) Then
        If InstrAtPositionSetArray(Acode(I), ipAny, True, adim) Then
          If LenB(aline) = 0 Then
            aline = Acode(I)
            aDimUsed = UsedMember(adim, Acode(I))
            DimUseDetectors aline, aDimUsed, strPName, StrProcClass, L0Assign, L0Calc, L0ForCount, L0ForDrive, L0LoopBreaker, L0Param, L0Next, L0TmpReturn, L0WithHead, L0IfTest, L0Input, L0Output, L0CaseSelector
            Exit For
          End If
        End If
      End If
    End If
  Next I
  If StrProcClass = "Const" Then
    If L0CaseSelector Then
      OneUseMode = Dim2Caser
      GoTo Done
    End If
    OneUseMode = Dim2LocalConst
    GoTo Done
  End If
  If L0Input Or L0Output Then
    OneUseMode = Dim1DummyInputOutPut
    GoTo Done
  End If
  If L0ForCount Then
    OneUseMode = Dim1ForUnmarkedNext
    GoTo Done
  End If
  If Not InQSortArray(StandardTypes, strType) Then
    If isClassName(strType) Then
      OneUseMode = Dim1LocalClass
      GoTo Done
     Else
      OneUseMode = Dim1UDT
      GoTo Done
    End If
  End If
  If L0Assign Then
    OneUseMode = Dim1DummyAssign
   Else
    If L0IfTest Or L0WithHead Or L0LoopBreaker Then
      OneUseMode = Dim1Unused
     Else
      OneUseMode = Dim1Dummy
    End If
  End If
Done:

End Function

Private Function TwoUseMode(adim As Variant, _
                            Acode As Variant, _
                            strType As String) As Dim2Use

  
  Dim strSwapTest       As String
  Dim LinesBetween      As Long
  Dim L0Deep            As Long
  Dim L1Deep            As Long
  Dim L0Assign          As Boolean
  Dim L0Calc            As Boolean
  Dim L1Assign          As Boolean
  Dim L1Calc            As Boolean
  Dim L0ForDrive        As Boolean
  Dim L0ForCount        As Boolean
  Dim L1ForDrive        As Boolean
  Dim L1ForCount        As Boolean
  Dim L0LoopBreaker     As Boolean
  Dim L1LoopBreaker     As Boolean
  Dim L0Param           As Boolean
  Dim L1Param           As Boolean
  Dim L0Next            As Boolean
  Dim L1Next            As Boolean
  Dim L0Input           As Boolean
  Dim L1Input           As Boolean
  Dim L0Output          As Boolean
  Dim L1Output          As Boolean
  Dim L0CaseSelector    As Boolean
  Dim L1CaseSelector    As Boolean
  Dim L0TmpReturn       As Boolean
  Dim L1TmpReturn       As Boolean
  Dim L0WithHead        As Boolean
  Dim L1WithHead        As Boolean
  Dim L0IfTest          As Boolean
  Dim L1IfTest          As Boolean
  Dim I                 As Long
  Dim lngdummy          As Long
  Dim strPName          As String
  Dim StrProcClass      As String
  Dim aLines(1)         As String
  Dim aLinesPos(1)      As Long
  Dim aDimUsed(1)       As String
  Dim ForSep            As Long
  Dim DoSep             As Long
  Dim NoSep             As Boolean
  Dim noLoopSep         As Boolean
  Dim EnclosedStructure As Long
  Dim Out2Inside        As Boolean

  For I = LBound(Acode) To UBound(Acode)
    If Not JustACommentOrBlankOrDimLine(Acode(I)) Then
      If isProcHead(Acode(I)) Then
        strPName = GetProcNameStr(GetWholeLineArray(Acode, I, lngdummy))
        StrProcClass = GetProcClassStr(GetWholeLineArray(Acode, I, lngdummy))
      End If
      If IsDimLine(Acode(I)) Then
        If InstrAtPositionSetArray(Acode(I), ipAny, True, adim) Then
          'v2.4.4 reconfigured to short circuit
          If InStr(Acode(I), LBracket) Then
            If InStr(Acode(I), RBracket) Then
              If InCode(Acode(I), InStr(Acode(I), LBracket)) Then
                TwoUseMode = Dim2Array
                Exit Function
              End If
            End If
          End If
        End If
       Else
        If InstrAtPositionSetArray(Acode(I), ipAny, True, adim) Then
          If LenB(aLines(0)) = 0 Then
            aLines(0) = Acode(I)
            aLinesPos(0) = I
            aDimUsed(0) = UsedMember(adim, Acode(I))
            DimUseDetectors aLines(0), aDimUsed(0), strPName, StrProcClass, L0Assign, L0Calc, L0ForCount, L0ForDrive, L0LoopBreaker, L0Param, L0Next, L0TmpReturn, L0WithHead, L0IfTest, L0Input, L0Output, L0CaseSelector
            L0Deep = GetStructureDepthLine(Acode, I)
           Else
            aLines(1) = Acode(I)
            aLinesPos(1) = I
            aDimUsed(1) = UsedMember(adim, Acode(I))
            DimUseDetectors aLines(1), aDimUsed(1), strPName, StrProcClass, L1Assign, L1Calc, L1ForCount, L1ForDrive, L1LoopBreaker, L1Param, L1Next, L1TmpReturn, L1WithHead, L1IfTest, L1Input, L1Output, L1CaseSelector
            L1Deep = GetStructureDepthLine(Acode, I)
            Exit For
          End If
        End If
      End If
    End If
  Next I
  For I = aLinesPos(0) To aLinesPos(1)
    If Not JustACommentOrBlankOrDimLine(Acode(I)) Then
      If LeftWord(Acode(I)) = "For" Then
        If ForSep = 0 Then
          If Acode(I) <> aLines(0) Then
            If Acode(I) <> aLines(1) Then
              Out2Inside = True
            End If
          End If
        End If
        ForSep = ForSep + 1
       ElseIf LeftWord(Acode(I)) = "Next" Then
        ForSep = ForSep - 1
        If ForSep = 0 Then
          Out2Inside = False
          EnclosedStructure = EnclosedStructure + 1
        End If
      End If
      If LeftWord(Acode(I)) = "Do" Then
        If DoSep = 0 Then
          If Acode(I) <> aLines(0) Then
            If Acode(I) <> aLines(1) Then
              Out2Inside = True
            End If
          End If
        End If
        DoSep = DoSep + 1
       ElseIf LeftWord(Acode(I)) = "Loop" Then
        DoSep = DoSep - 1
        If DoSep = 0 Then
          Out2Inside = False
          'In2Outside = False
          EnclosedStructure = EnclosedStructure + 1
        End If
      End If
    End If
  Next I
  For I = aLinesPos(0) + 1 To aLinesPos(1)
    If Not JustACommentOrBlankOrDimLine(Acode(I)) Then
      LinesBetween = LinesBetween + 1
    End If
  Next I
  'For I = aLinesPos(0) + 1 To aLinesPos(1) - 1
  'Next
  If Len(aLines(0)) Then
    If Len(aLines(1)) = 0 Then
      TwoUseMode = Dim2dummyTwiceinaLine
      GoTo Done
    End If
  End If
  If L0ForDrive Then
    If L1Assign Then
      TwoUseMode = Dim2SetOnceLoopDriver
      GoTo Done
    End If
  End If
  If L0Param Then
    If L1LoopBreaker Then
      TwoUseMode = Dim2InLoopLoopDriver
      GoTo Done
     ElseIf L1IfTest Then
      TwoUseMode = Dim2AssignIfTest
      GoTo Done
    End If
  End If
  If L0Input Then
    TwoUseMode = Dim2InputLoading
    GoTo Done
  End If
  If StrProcClass = "Const" Then
    TwoUseMode = Dim2LocalConst
    GoTo Done
  End If
  NoSep = (DoSep = 0 And ForSep = 0)
  noLoopSep = (DoSep = 0 And ForSep = 0)
  '---------- 1st line is outside a For/Do 2nd inside
  If L0ForCount And L1Next Then
    TwoUseMode = Dim2NextCounter
    GoTo Done
  End If
  If Out2Inside Then
    TwoUseMode = Dim2SetOnceLoopDriver
    GoTo Done
  End If
  '----------
  If L0WithHead Then
    TwoUseMode = Dim2WithHead1
  End If
  If L1WithHead Then
    TwoUseMode = Dim2WithHead2
  End If
  '----------------
  If L1ForDrive Then
    TwoUseMode = Dim2ChangeableForRange
    GoTo Done
  End If
  If L0IfTest Then
    If L1Assign Then
      TwoUseMode = Dim2OnceOnlyTest
      GoTo Done
    End If
  End If
  '-------------  c=a:a=b:b=c
  If LinesBetween > 1 Then 'must be at least 1 line between
    If L0Assign Then
      If L1IfTest Then
        TwoUseMode = Dim2AssignIfTest
        GoTo Done
      End If
      If L1Param Then
        If Abs(L0Deep - L1Deep) > 1 Then
          TwoUseMode = Dim2AssignParam
          GoTo Done
        End If
        TwoUseMode = Dim2InLine
        GoTo Done
      End If
      If L1Calc Then
        If SmartRight(aLines(1), EqualInCode & LeftWord(aLines(0))) Then
          TwoUseMode = Dim2PreserveValue
          strSwapTest = WordAfter(aLines(0), "=") & EqualInCode & LeftWord(aLines(1))
          For I = aLinesPos(0) + 1 To aLinesPos(1) - 1
            If SmartLeft(Acode(I), strSwapTest) Then
              TwoUseMode = Dim2Swapper
              Exit For 'unction
            End If
          Next I
          GoTo Done
         Else
          If NoSep Then
            If SmartLeft(aLines(0), LeftWord(aLines(0)) & EqualInCode) Then
              If Abs(L0Deep - L1Deep) < 2 Then
                TwoUseMode = Dim2InLine
                GoTo Done
              End If
            End If
          End If
        End If
      End If
    End If
  End If
  '----------
  If SmartRight(aLines(1), EqualInCode & LeftWord(aLines(0))) Then
    TwoUseMode = Dim2PreserveValue
    strSwapTest = WordAfter(aLines(0), "=") & EqualInCode & LeftWord(aLines(1))
    For I = aLinesPos(0) + 1 To aLinesPos(1) - 1
      If SmartLeft(Acode(I), strSwapTest) Then
        TwoUseMode = Dim2Swapper
        GoTo Done
      End If
    Next I
    GoTo Done
  End If
  '----------
  If L0Assign Then                      'x = y
    If L1Calc Then                        'c = x * z
      If LinesBetween = 1 Then    'following line or no looping structures
        If Abs(L0Deep - L1Deep) < 2 Then
          TwoUseMode = Dim2InLine
          GoTo Done
        End If
       ElseIf NoSep Then
        If EnclosedStructure = 0 Then
          If Abs(L0Deep - L1Deep) < 2 Then
            TwoUseMode = Dim2InLine
            GoTo Done
          End If
         Else
          TwoUseMode = Dim2SetOnceLoopDriver
          GoTo Done
        End If
       ElseIf Out2Inside Then
        TwoUseMode = Dim2SetOnceLoopDriver
        GoTo Done
       ElseIf EnclosedStructure Then
        TwoUseMode = Dim2SetOnceLoopDriver
        GoTo Done
       ElseIf L1TmpReturn Then
        TwoUseMode = Dim2TmpVarForReturn
        GoTo Done
       ElseIf L1IfTest And noLoopSep Then
        TwoUseMode = Dim2IfTest
        GoTo Done
      End If
     ElseIf L1IfTest Then
      If LinesBetween = 1 Or noLoopSep Then    'following line or no looping structures
        TwoUseMode = Dim2IfTest
        GoTo Done
      End If
     ElseIf L1Param Then
      If Abs(L0Deep - L1Deep) < 2 Then
        TwoUseMode = Dim2InLine
        GoTo Done
      End If
    End If
  End If
  '----------
  If L1ForDrive Then
    If noLoopSep Then
      If Abs(L0Deep - L1Deep) < 2 Then
        TwoUseMode = Dim2InLine
        GoTo Done
      End If
     Else
      If Out2Inside Or EnclosedStructure Then
        TwoUseMode = Dim2SetOnceLoopDriver
        GoTo Done
      End If
    End If
  End If
  '----- Used as counter in 2 For/Next structures or counter in one for and start/end in inner one
  If L0ForCount Then               'For X =......
    If L1Next Then                  'Next X
      TwoUseMode = Dim2NextCounter
      GoTo Done
     ElseIf L1ForCount Then
      'For X (2nd time NOTE Next without variable (Should not happen in CF treated code)
      TwoUseMode = Dim2NextCounter
      GoTo Done
     ElseIf L1ForDrive Then           'For y = X to J'  OR 'For y = J To X'
      TwoUseMode = Dim2NextOuterCounterInnerDriver
      GoTo Done
     Else
      If L1Calc Or L1Param Then
        ' Z =  calculation with X inside loop or sent as parameter inside loop
        TwoUseMode = Dim2NextCounter
        GoTo Done
      End If
    End If
  End If
  '-----------
  If L0LoopBreaker Then
    If L1Assign Or L1Param Then
      TwoUseMode = Dim2LoopBreak
      GoTo Done
    End If
  End If
  If L1LoopBreaker Then
    If L0Assign Or L0Param Then
      TwoUseMode = Dim2LoopBreak
      GoTo Done
    End If
  End If
  If L0Param Then
    If L1Param Then
      TwoUseMode = Dim2Proc2Proc
      GoTo Done
     ElseIf L1Calc Then
      TwoUseMode = Dim2ProcForCalc
      GoTo Done
    End If
  End If
  If Not InQSortArray(StandardTypes, strType) Then
    If isClassName(strType) Then
      TwoUseMode = Dim2LocalClass
      GoTo Done
     Else
      If L0WithHead Then
        If L1Param Then
          TwoUseMode = Dim2UDTSetPass
          GoTo Done
        End If
      End If
      If L0Param Then
        If L1Param Then
          TwoUseMode = Dim2UDTParamPass
          GoTo Done
        End If
      End If
      TwoUseMode = Dim2UDT
      GoTo Done
    End If
  End If
  If L0Param Then
    If L1CaseSelector Or L1IfTest Then
      TwoUseMode = Dim2ParamChoice
      GoTo Done
    End If
    If L1Param Then
      TwoUseMode = Dim2UDTParamPass
      GoTo Done
    End If
  End If
  If Abs(L0Deep - L1Deep) > 1 Then
    TwoUseMode = Dim2OnceOnlyTest
    GoTo Done
  End If
  If L0ForCount Then
    TwoUseMode = Dim2NextCounter
    'Exit Function
  End If
Done:

End Function

Private Function UsedMember(adim As Variant, _
                            ByVal varCode As Variant) As String

  Dim tdim As Variant

  varCode = ExpandForDetection(varCode)
  For Each tdim In adim
    If InstrAtPosition(varCode, tdim, ipAny) Then
      UsedMember = tdim
      Exit For
    End If
  Next tdim

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:15:38 AM) 53 + 777 = 830 Lines Thanks Ulli for inspiration and lots of code.

