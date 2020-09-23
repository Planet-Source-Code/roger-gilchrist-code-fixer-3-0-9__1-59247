Attribute VB_Name = "mod_Parameter"

'Copyright 2003 Roger Gilchrist
'e-mail: rojagilkrist@hotmail.com
Option Explicit
Public ArrQActiveControlClass                       As Variant
Public Enum ParamArrayType
  FullParam
  VariableOnlyParam
  VariableOnlyParam2
  UnTypedParam
  SuffixTypeParam
  OptionalParam
  NoBy_Param
  NoBy_Param2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private FullParam, VariableOnlyParam, VariableOnlyParam2, UnTypedParam, SuffixTypeParam, OptionalParam, NoBy_Param, NoBy_Param2
Private NoBy_Param, NoBy_Param2
#End If


Private Sub BadNameParameter(cMod As CodeModule)

  Dim ModuleNumber  As Long
  Dim L_CodeLine    As String
  Dim ArrMember     As Variant
  Dim Member        As Long
  Dim ArrRoutine    As Variant
  Dim UpDated       As Boolean
  Dim MUpdated      As Boolean
  Dim MemberCount   As Long
  Dim strMsg        As String
  Dim TopOfRoutine  As Long
  Dim LineContRange As Long
  Dim Rname         As String
  Dim strOld        As String
  Dim Strnew        As String
  Dim strFind       As String
  Dim ArrOld        As Variant
  Dim ArrNew        As Variant
  Dim I             As Long
  Dim J             As Long
  Dim K             As Long
  Dim A             As Long

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ParameterBadName_CXX) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        If UBound(ArrRoutine) > -1 Then
          If Not IsStub(ArrRoutine) Then
            L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, Rname)
            If Not IsVBControlRoutine(Rname) Then
              strMsg = vbNullString
              strOld = vbNullString
              Strnew = vbNullString
              BadNameParameterTest L_CodeLine, strMsg, strOld, Strnew
              If Len(strMsg) Then
                ParamLineContFix LineContRange, TopOfRoutine, ArrRoutine, L_CodeLine
                Select Case FixData(ParameterBadName_CXX).FixLevel
                 Case CommentOnly
                  ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, strMsg, MAfter, UpDated)
                  AddNfix ParameterBadName_CXX
                 Case FixAndComment
                  ArrOld = CleanArray(Split(strOld, ","))
                  ArrNew = CleanArray(Split(Strnew, ","))
                  If UBound(ArrOld) <> UBound(ArrNew) Then
                    'Safety: Didn't make replacement for all bad parameters so just comment
                    ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, strMsg, MAfter, UpDated)
                    AddNfix ParameterBadName_CXX
                   Else
                    For J = LBound(ArrOld) To UBound(ArrOld)
                      If Len(ArrOld(J)) Then
                        strFind = Mid$(ArrOld(J), 2)
                        strFind = Left$(strFind, Len(strFind) - 1)
                        For K = LBound(ArrNew) To UBound(ArrNew)
                          If ArrNew(K) <> strFind Then
                            If SmartRight(ArrNew(K), strFind, False) Then
                              For I = LBound(ArrRoutine) To UBound(ArrRoutine)
                                If Not JustACommentOrBlank(ArrRoutine(I)) Then
                                  strOld = ArrRoutine(I)
                                  If InStr(strOld, strFind) Then
                                    'v2.6.9  true parameter is modification to deal with
                                    'Func Color_X(ByVal Color)
                                    'updating only the parameter to lngColor
                                    WholeWordReplacerParameter strOld, strFind, ArrNew(K), UpDated, True
                                    'v2.8.7 because the suffix stripper hasn't occured this special sub version triggers
                                    If ArrRoutine(I) = strOld Then
                                      For A = LBound(TypeSuffixArray) To UBound(TypeSuffixArray)
                                        WholeWordReplacerParameter strOld, strFind & TypeSuffixArray(A), ArrNew(K), UpDated, True
                                        If ArrRoutine(I) <> strOld Then
                                          Exit For
                                        End If
                                      Next A
                                    End If
                                    'v2.1.1 deal with spcial case of Variable has same name as Type ie 'Picture as Picture'
                                    'v2.7.9 woops new and old were reversed
                                    'v2.8.0 woops2 that'll teach me not to read my own comments ;)What was I thinking of? Thanks Ian
                                    WholeWordReplacerParameter strOld, "As " & ArrNew(K), "As " & strFind, UpDated
                                    ArrRoutine(I) = strOld
                                    AddNfix ParameterBadName_CXX
                                  End If
                                End If
                              Next I
                            End If
                          End If
                        Next K
                      End If
                    Next J
                    L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, Rname)
                    strMsg = WARNING_MSG & "Poorly named Parameters " & Join(CleanArray(ArrOld), CommaSpace) & " renamed to " & strInSQuotes(Join(CleanArray(ArrNew), "', '"))
                    ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, strMsg, MAfter, UpDated)
                  End If
                End Select
              End If
              UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
            End If
          End If
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Sub BadNameParameterTest(ByVal TestLine As String, _
                                 strMsg As String, _
                                 strMsgOrig As String, _
                                 strMsgNew As String)

  Dim I                As Long
  Dim ArrParam         As Variant
  Dim ArrParamFull     As Variant
  Dim strParam         As String
  Dim strMsg1          As String
  Dim strMsg2          As String
  Dim strMsg3          As String
  Dim strMsg4          As String
  Dim strSuggestPrefix As String

  strMsg = vbNullString
  If ExtractCode(TestLine) Then
    If GetLeftBracketPos(TestLine) Then
      strParam = TestLine
      ArrParam = ParamArrays(strParam, VariableOnlyParam2)
      ArrParamFull = ParamArrays(strParam, FullParam)
      If ArrayHasContents(ArrParam) Then
        For I = LBound(ArrParam) To UBound(ArrParam)
          If CntrlDescMember(ArrParam(I)) > -1 Then
            strMsg1 = AccumulatorString(strMsg1, strInSQuotes(ArrParam(I)))
            strSuggestPrefix = AccumulatorString(strSuggestPrefix, PrefixFromType(ArrParam(I), ArrParamFull(I)))
           ElseIf IsControlProperty(ArrParam(I)) Then
            strMsg2 = AccumulatorString(strMsg2, strInSQuotes(ArrParam(I)))
            strSuggestPrefix = AccumulatorString(strSuggestPrefix, PrefixFromType(ArrParam(I), ArrParamFull(I)))
           ElseIf IsControlProperty(CStr(ArrParam(I))) Then
            strMsg3 = AccumulatorString(strMsg3, strInSQuotes(ArrParam(I)))
            strSuggestPrefix = AccumulatorString(strSuggestPrefix, PrefixFromType(ArrParam(I), ArrParamFull(I)))
           ElseIf IsControlEvent(CStr(ArrParam(I))) Then
            strMsg4 = AccumulatorString(strMsg4, strInSQuotes(ArrParam(I)))
            strSuggestPrefix = AccumulatorString(strSuggestPrefix, PrefixFromType(ArrParam(I), ArrParamFull(I)))
           ElseIf isRefLibVBCommands(CStr(ArrParam(I)), False) Then
            strMsg2 = AccumulatorString(strMsg2, strInSQuotes(ArrParam(I)))
            strSuggestPrefix = AccumulatorString(strSuggestPrefix, PrefixFromType(ArrParam(I), ArrParamFull(I)))
          End If
        Next I
        If Len(strMsg1) Then
          strMsg = BuildSameNameMsg(strMsg1, strSuggestPrefix, " Control.")
        End If
        If Len(strMsg2) Then
          strMsg = strMsg & IIf(Len(strMsg), vbNewLine, vbNullString)
          strMsg = strMsg & BuildSameNameMsg(strMsg2, strSuggestPrefix, " VB command.")
        End If
        If Len(strMsg3) Then
          strMsg = strMsg & IIf(Len(strMsg), vbNewLine, vbNullString)
          strMsg = strMsg & BuildSameNameMsg(strMsg3, strSuggestPrefix, " Control Property.")
        End If
        If Len(strMsg4) Then
          strMsg = strMsg & IIf(Len(strMsg), vbNewLine, vbNullString)
          strMsg = strMsg & BuildSameNameMsg(strMsg3, strSuggestPrefix, " Control Event.")
        End If
      End If
    End If
  End If
  strMsgOrig = LStrip(strMsg4 & "," & strMsg3 & "," & strMsg2 & "," & strMsg1, ",")
  strMsgOrig = Trim$(RStrip(strMsgOrig, ","))
  strMsgNew = Trim$(strSuggestPrefix)

End Sub

Private Function BuildSameNameMsg(ByVal strMsg As String, _
                                  ByVal strUpdate As String, _
                                  ByVal strMatchObj As String) As String

  Dim Plural     As Boolean

  Plural = InStr(strMsg, ",") > 0
  BuildSameNameMsg = WARNING_MSG & "Parameter" & IIf(Plural, "s ", SngSpace) & strMsg & IIf(Plural, " have the same names as a", " has the same name as a") & strMatchObj & vbNewLine & _
   "Legal but makes code hard to read. " & SUGGESTION_MSG & "Rename the parameter" & IIf(Plural, "s.", ".") & IIf(Len(strUpdate), " Suggested " & strInBrackets(strUpdate), " using type Suffixes is suggested.")

End Function

Private Sub ByValParameter(cMod As CodeModule)

  Dim ModuleNumber    As Long
  Dim CompName        As String
  Dim L_CodeLine      As String
  Dim ArrMember       As Variant
  Dim Member          As Long
  Dim ArrRoutine      As Variant
  Dim UpDated         As Boolean
  Dim MUpdated        As Boolean
  Dim MemberCount     As Long
  Dim strMsg          As String
  Dim strMsg2         As String
  Dim TopOfRoutine    As Long
  Dim LineContRange   As Long
  Dim Rname           As String
  Dim ArrParam        As Variant
  Dim I               As Long
  Dim StrParamNew     As String
  Dim lngNoTypeCount  As Long
  Dim StrParamOld     As String

  Const strVerboseMsg As String = vbNewLine & _
   "WARNING NEW FIX : This is still experimental (Testing is very conservative)." & vbNewLine & _
   " List may be incomplete or contain members in error, test carefully." & vbNewLine & _
   " User created Events can use ByVal but you must edit the Declaration as well." & vbNewLine & _
   "otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'" & vbNewLine & _
   " The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)" & vbNewLine & _
   "OR you don't want any changes returned(You have to hand code this) make the parameter ByVal."
  CompName = cMod.Parent.Name
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ParameterByVal) Then
    'v2.9.7 Don't do to implements classes Thanks Ian K
    If Not InQSortArray(ImplementsArray, CompName) Then
      ArrMember = GetMembersArray(cMod)
      MemberCount = UBound(ArrMember)
      If MemberCount > 0 Then
        For Member = 1 To MemberCount
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
          If UBound(ArrRoutine) > -1 Then
            If Not IsStub(ArrRoutine) Then
              L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, Rname)
              If Not isUserEvent(Rname, CompName) Then
                If Not IsVBControlRoutine(Rname) Then
                  If ByValParameterTest(L_CodeLine, TopOfRoutine + LineContRange, ArrRoutine, strMsg) Then
                    ParamLineContFix LineContRange, TopOfRoutine, ArrRoutine, L_CodeLine
                    Select Case FixData(ParameterByVal).FixLevel
                     Case CommentOnly
                      ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, SUGGESTION_MSG & "Insert 'ByVal ' for Parameter" & IIf(InStr(strMsg, CommaSpace), "s", vbNullString) & SngSpace & strInSQuotes(strMsg) & IIf(Xcheck(XVerbose), strVerboseMsg, vbNullString), MAfter, UpDated)
                      AddNfix ParameterByVal
                     Case FixAndComment
                      If LineContRange = 0 Then
                        ArrParam = Split(strMsg, ",")
                        StrParamNew = Mid$(L_CodeLine, InStr(L_CodeLine, LBracket))
                        StrParamOld = StrParamNew
                        strMsg2 = vbNullString
                        lngNoTypeCount = 0
                        For I = LBound(ArrParam) To UBound(ArrParam)
                          'v2.9.7 copes with a single untyped parameter
                          If InStr(ArrParam(I), " As ") Then
                            StrParamNew = Safe_Replace(StrParamNew, ArrParam(I), " ByVal " & ArrParam(I), 1, 1, False)
                           Else
                            strMsg2 = strMsg2 & IIf(Len(strMsg2), ", ", vbNullString) & ArrParam(I)
                            strMsg = Replace$(strMsg, ", " & ArrParam(I), vbNullString)
                            strMsg = Replace$(strMsg, ArrParam(I) & ",", vbNullString)
                            strMsg = Replace$(strMsg, ArrParam(I), vbNullString)
                            lngNoTypeCount = lngNoTypeCount + 1
                          End If
                        Next I
                        L_CodeLine = Replace$(L_CodeLine, StrParamOld, StrParamNew)
                        If LenB(strMsg) Then
                          ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, WARNING_MSG & "'ByVal ' inserted for Parameter" & IIf(InStr(strMsg, CommaSpace), "s", vbNullString) & SngSpace & strInSQuotes(strMsg), MAfter, UpDated)
                        End If
                        If lngNoTypeCount Then
                          L_CodeLine = ArrRoutine(TopOfRoutine)
                          ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, WARNING_MSG & "may be able to apply 'ByVal ' Parameter" & IIf(InStr(strMsg2, CommaSpace), "s", vbNullString) & SngSpace & strInSQuotes(strMsg2) & " once Type-cast", MAfter, UpDated)
                        End If
                       Else
                        ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, SUGGESTION_MSG & "Insert 'ByVal ' for Parameter" & IIf(InStr(strMsg, CommaSpace), "s", vbNullString) & SngSpace & strInSQuotes(strMsg) & IIf(Xcheck(XVerbose), strVerboseMsg, vbNullString), MAfter, UpDated)
                      End If
                    End Select
                    strMsg = vbNullString
                  End If
                  If UpDated Then
                    ArrMember(Member) = Join(ArrRoutine, vbNewLine)
                    MUpdated = True
                    UpDated = False
                  End If
                End If
              End If
            End If
          End If
        Next Member
        ReWriteMembers cMod, ArrMember, MUpdated
      End If
    End If
  End If

End Sub

Private Function ByValParameterTest(ByVal TestLine As String, _
                                    ByVal StartPos As Long, _
                                    Arr As Variant, _
                                    strMsg As String) As Boolean

  Dim I                As Long
  Dim J                As Long
  Dim K                As Long
  Dim L                As Long
  Dim ArrParam         As Variant
  Dim ArrParamTest     As Variant
  Dim ArrParamFullTest As Variant
  Dim strParam         As String
  Dim L_CodeLine       As String
  Dim ArrCode          As Variant

  strMsg = vbNullString
  If GetLeftBracketPos(TestLine) Then
    strParam = TestLine
    ArrParam = ParamArrays(strParam, VariableOnlyParam2)
    ArrParamTest = ParamArrays(strParam, NoBy_Param2)
    ArrParamFullTest = ParamArrays(strParam, FullParam)
    If ArrayHasContents(ArrParamTest) Then
      For K = 0 To UBound(ArrParamTest)
        If Len(ArrParamTest(K)) Then
          For I = StartPos + 1 To UBound(Arr)
            If Not JustACommentOrBlank(Arr(I)) Then
              L_CodeLine = ExpandForDetection(Arr(I))
              For J = 0 To UBound(ArrParam)
                If LenB(Trim$(ArrParamTest(J))) = 0 Then
                  SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                End If
                'v2.8.0' if Mid$ is used to change a part of a string then no ByVal
                If MultiLeft(L_CodeLine, True, "Mid (", "Mid$ (") Then
                  If InStr(L_CodeLine, ArrParamTest(J)) < InStr(L_CodeLine, " = ") Then
                    SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                    Exit For
                  End If
                End If
                If LenB(Trim$(ArrParam(J))) Then
                  '*If a value is assigned (=) OR A with structure is used (and probably a value is assigned in the With structure)
                  'then ByVal cannot be used. Will develop a better test for the With structure later on.
                  If MultiLeft(L_CodeLine, True, ArrParam(J) & EqualInCode, "With " & ArrParam(J), ArrParam(J) & ".") Then
                    SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                    Exit For
                  End If
                  '*If ByVal/ByRef is already set don't suggest
                  'May add a test to check that ByRef is required in future
                  If InstrArrayWholeWord(ArrParamTest(J), "ByRef", "ByVal", "Optional ByRef", "Optional ByVal") Then
                    SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                    Exit For
                  End If
                  If InStr(ArrParamTest(J), "()") Then
                    SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                    Exit For
                  End If
                  'User Defined Types cannot use the ByVal command because VB has to pass the Object/Type
                  'so that it knows what to do with the data.
                  If Get_As_Pos(ArrParamTest(J)) > 0 Then
                    If Not InQSortArray(StandardTypes, GetType(ArrParamTest(J))) Then
                      SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                      Exit For
                    End If
                  End If
                  'TEMPORARY PATCH 1
                  'until CF can recognize Array assignment this stops Arrays being suggested
                  If InStr(ArrParamTest(J), " As Variant") Then
                    SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                    Exit For
                  End If
                  'TEMPORARY PATCH 2
                  'this blocks any variable being passed to other routines
                  'eventually there will be a test to see if it is sent to a ByVal parameter
                  ' and so can be set to ByVal in current routine
                  ArrCode = Split(ExpandForDetection(L_CodeLine))
                  For L = LBound(ArrCode) To UBound(ArrCode)
                    If ArrCode(L) = ArrParam(J) Then
                      Exit For
                    End If
                    If IsProcedureName(ArrCode(L)) Then
                      If InstrAtPosition(L_CodeLine, ArrParam(J), ipAny, True) Then
                        SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                      End If
                    End If
                  Next L
                  'ver 1.1.17
                  'This stops Properties using array indexes from being given ByVal
                  'which don't work for indexing values
                  If InstrAtPosition(TestLine, "Property", ipLeftOr2nd, True) Then
                    If LenB(ArrParam(J)) Then
                      If InStr(L_CodeLine, EqualInCode) > InStr(L_CodeLine, ArrParam(J)) Then
                        If EnclosedInBrackets(L_CodeLine, InStr(L_CodeLine, ArrParam(J))) Then
                          SetToNullString ArrParam(J), ArrParamTest(J), ArrParamFullTest(J)
                        End If
                      End If
                    End If
                  End If
                End If
              Next J
            End If
          Next I
        End If
      Next K
      For I = 0 To UBound(ArrParamTest)
        'get rid of blank space holders
        If Len(Trim$(ArrParamTest(I))) = 0 Then
          ArrParamTest(I) = vbNullString
        End If
      Next I
      strMsg = safe_Join(ArrParamFullTest, CommaSpace)
      ByValParameterTest = UBound(Split(strMsg, CommaSpace)) > -1
    End If
  End If

End Function

Private Sub Function2Sub(cMod As CodeModule)

  Dim ModuleNumber As Long
  Dim FunctIsFunct As Boolean
  Dim L_CodeLine   As String
  Dim strTest      As String
  Dim FuncName     As String
  Dim ArrMember    As Variant
  Dim RLine        As Long
  Dim Member       As Long
  Dim ArrRoutine   As Variant
  Dim TmpC         As Variant
  Dim UpDated      As Boolean
  Dim MUpdated     As Boolean
  Dim MemberCount  As Long

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ConvertFunction2Sub) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
          L_CodeLine = ArrRoutine(RLine)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If InstrAtPosition(L_CodeLine, "Function", ipLeftOr2ndOr3rd) Then
              strTest = L_CodeLine
              If ExtractCode(L_CodeLine) Then
                If SmartRight(Trim$(strTest), RBracket) And Not IsArrayReturningFunction(strTest) Then
                  'this catches array returning functions Function Fred(Wilma as long) As String()
                  FuncName = Left$(strTest, GetLeftBracketPos(strTest) - 1)
                  TmpC = Split(FuncName)
                  FuncName = TmpC(UBound(TmpC))
                  FunctIsFunct = RoutineSearch(ArrRoutine, FuncName, RLine, ipAny, True)
                  'ver1.1.20
                  ' added test for Function name withType suffix
                  If Not FunctIsFunct Then
                    If IsPunct(Right$(FuncName, 1)) Then
                      FunctIsFunct = RoutineSearch(ArrRoutine, Left$(FuncName, Len(FuncName) - 1), RLine, ipAny, True)
                    End If
                  End If
                  If InQSortArray(ImplementsArray, ModDesc(ModuleNumber).MDName) Then
                    'leave implements classes alone
                    FunctIsFunct = True
                  End If
                  If Not FunctIsFunct Then
                    Select Case FixData(ConvertFunction2Sub).FixLevel
                     Case CommentOnly
                      ArrRoutine(RLine) = Marker(L_CodeLine, SUGGESTION_MSG & "Function should be changed to Sub as nothing is returned via the Function Name.", MAfter)
                     Case FixAndComment, JustFix
                      If InstrAtPosition(L_CodeLine, "Function", IpLeft) Then
                        L_CodeLine = Safe_Replace(L_CodeLine, "Function ", "Sub ", 1, 1)
                       ElseIf InstrAtPosition(L_CodeLine, "Function", ip2nd) Then
                        L_CodeLine = Safe_Replace(L_CodeLine, " Function ", " Sub ", , 1)
                      End If
                      If ArrRoutine(RLine) <> L_CodeLine Then
                        If FixData(ConvertFunction2Sub).FixLevel = FixAndComment Then
                          ArrRoutine(RLine) = Marker(L_CodeLine, WARNING_MSG & "Function changed to Sub as nothing is returned via the Function Name.", MAfter)
                         Else
                          ArrRoutine(RLine) = L_CodeLine
                        End If
                      End If
                    End Select
                    UpDated = True
                  End If
                End If
              End If
            End If
          End If
        Next RLine
        If UpDated Then
          AddNfix ConvertFunction2Sub
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Function GetCallingRoutine(cMod As CodeModule, _
                                   ByVal RoutineName As String, _
                                   strParam As String, _
                                   MissingCount As Long, _
                                   ByVal strCurCompName As String) As Variant

  Dim AsusedA     As Variant
  Dim ParamA      As Variant
  Dim K           As Long
  Dim M           As Long
  Dim N           As Long
  Dim strDummy    As String
  Dim Parameters  As String
  Dim FoundOne    As Boolean
  Dim Comp        As VBComponent
  Dim Proj        As VBProject
  Dim TmpC        As Variant
  Dim TmpD        As Variant
  Dim tmpCodArr   As Variant
  Dim I           As Long
  Dim J           As Long
  Dim strCodeDump As String
  Dim AsUsed      As String
  Dim MaxFactor   As Long

  'Collect and record names and types of Functions
  'For UnusedFunction tests
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If ModuleHasCode(Comp.CodeModule) Then
          If Comp.Name <> strCurCompName Then
            GoTo SkipComp
          End If
          tmpCodArr = GetMembersArray(cMod)
          MaxFactor = UBound(tmpCodArr)
          If MaxFactor > -1 Then
            For I = 0 To MaxFactor
              TmpC = Split(tmpCodArr(I), vbNewLine)
              strCodeDump = vbNullString
              Parameters = vbNullString
              If RoutineSearch(TmpC, RoutineName, 0, ipAny, True) Then
                If InStr(TmpC(0), "()") = 0 Then
                  For J = 0 To UBound(TmpC)
                    If J = 0 Then
                      ParamaterBreakOut CStr(TmpC(J)), strDummy, Parameters, strDummy, strDummy
                      Parameters = ReplaceArray(Parameters, "byVal ", vbNullString, "byRef ", vbNullString, "Optional ", vbNullString, CommaSpace, vbNewLine)
                      strCodeDump = AccumulatorString(strCodeDump, Parameters, vbNewLine)
                     ElseIf IsDimLine(TmpC(J)) Then
                      strCodeDump = AccumulatorString(strCodeDump, Mid$(TmpC(J), 4), vbNewLine)
                     ElseIf InStr(TmpC(J), RoutineName) Then
                      AsUsed = TmpC(J)
                    End If
                  Next J
                  strCodeDump = Safe_Replace(Trim$(strCodeDump), CommaSpace, ",")
                  AsUsed = Mid$(AsUsed, InStr(AsUsed, RoutineName) + Len(RoutineName) + 1)
                  AsUsed = Safe_Replace(Trim$(AsUsed), CommaSpace, ",")
                  TmpD = Split(strCodeDump, vbNewLine)
                  AsusedA = Split(AsUsed, ",")
                  ParamA = Split(strParam, ",")
                  If UBound(ParamA) = UBound(AsusedA) Then
                    For K = LBound(AsusedA) To UBound(AsusedA)
                      If LenB(ParamA(K)) Then
                        For M = LBound(TmpD) To UBound(TmpD)
                          If AsusedA(K) & SngSpace = Left$(TmpD(M), Len(AsusedA(K)) + 1) Then
                            For N = LBound(ParamA) To UBound(ParamA)
                              If Get_As_Pos(ParamA(K)) = 0 Then
                                If Get_As_Pos(TmpD(M)) Then
                                  ParamA(K) = ParamA(K) & " As " & GetType(TmpD(M))
                                  MissingCount = MissingCount - 1
                                  FoundOne = True
                                End If
                              End If
                            Next N
                          End If
                        Next M
                      End If
                    Next K
                  End If
                End If
              End If
              If FoundOne Then
                strParam = Join(ParamA, CommaSpace)
                Exit For
              End If
            Next I
          End If
        End If
      End If
SkipComp:
      If FoundOne Or MissingCount = 0 Then
        Exit For
      End If
    Next Comp
    If FoundOne Or MissingCount = 0 Then
      Exit For
    End If
  Next Proj
  '

End Function

Private Function GetFunctionTypeFromInternalEvidence(ByVal ModuleNumber As Long, _
                                                     arrModule As Variant, _
                                                     ByVal FName As String, _
                                                     ByVal CompName As String) As String

  Dim I                As Long
  Dim J                As Long
  Dim arrLine          As Variant
  Dim ExistingDimArray As Variant

  If UsingDefTypes Then
    GetFunctionTypeFromInternalEvidence = FromDefType(FName, ModuleNumber)
   Else
    ExistingDimArray = BuildExistingDimArray(Join(arrModule, vbNewLine))
    For I = LBound(arrModule) + 1 To UBound(arrModule) - 1
      If SmartLeft(arrModule(I), FName) Then
        arrLine = Split(arrModule(I))
        If UBound(arrLine) > 0 Then
          If arrLine(1) = "=" Then
            GetFunctionTypeFromInternalEvidence = InternalEvidenceType(arrLine(2), CompName)
            If LenB(GetFunctionTypeFromInternalEvidence) Then
              Exit For
            End If
            GetFunctionTypeFromInternalEvidence = InternalEvidenceType(JoinPartial(arrLine, 2, UBound(arrLine)), CompName)
            If LenB(GetFunctionTypeFromInternalEvidence) Then
              Exit For
            End If
            GetFunctionTypeFromInternalEvidence = InternalEvidenceType(JoinPartial(arrLine, 3, UBound(arrLine)), CompName)
            If LenB(GetFunctionTypeFromInternalEvidence) Then
              Exit For
            End If
           ElseIf UBound(arrLine) > 1 Then
            'v2.9.6 possible bug? v3.0.0 (I) changed to (J)
            If InQSortArray(ExistingDimArray, arrLine(J)) Then
              For J = LBound(arrModule) To UBound(arrModule)
                If WriteType(arrModule, J, arrLine, "Dim", GetFunctionTypeFromInternalEvidence) Then
                  Exit For
                End If
                If WriteType(arrModule, J, arrLine, "Const", GetFunctionTypeFromInternalEvidence) Then
                  Exit For
                End If
                If WriteType(arrModule, J, arrLine, "Static", GetFunctionTypeFromInternalEvidence) Then
                  Exit For
                End If
              Next J
             ElseIf IsNumeric(arrLine(2)) Then
              GetFunctionTypeFromInternalEvidence = " As " & LowestTypeFit(arrLine(2))
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function GetRoutineDeclaration(Arr As Variant, _
                                      Lnum As Long, _
                                      Range As Long, _
                                      RoutineName As String) As String

  Lnum = GetProcCodeLineOfRoutine(Arr)
  GetRoutineDeclaration = GetWholeLineArray(Arr, Lnum, Range)
  RoutineName = GetRoutineName(GetRoutineDeclaration)

End Function

Private Function InStrCodeRev(ByVal varSearch As Variant, _
                              ByVal varFind As Variant) As Long

  Dim TestPos As Long

  TestPos = InStrRev(varSearch, varFind)
  Do While TestPos
    If InCode(varSearch, TestPos) Then
      InStrCodeRev = TestPos
      Exit Do 'unction
    End If
    TestPos = InStrRev(varSearch, varFind, TestPos - 1)
  Loop
  'InStrCodeRev = 0

End Function

Public Function IsArrayEmpty(aArray As Variant) As Boolean

  'Richard Hundhausen, Boise, Idaho
  'VBPJ TechTips 8th edition

  On Error Resume Next
  'crash at this line means that error trapping is 'break on all errors'
  '  IsArrayEmpty = UBound(aArray)
  '  IsArrayEmpty = Err ' Error 9 (Subscript out of range)
  IsArrayEmpty = IsNumeric(UBound(aArray))
  On Error GoTo 0

End Function

Private Function IsArrayReturningFunction(ByVal strTest As String) As Boolean

  Dim TmpA As Variant

  TmpA = Split(strTest)
  If Right$(TmpA(UBound(TmpA)), 2) = "()" Then
    If TmpA(UBound(TmpA) - 1) = "As" Then
      IsArrayReturningFunction = CountSubString(strTest, LBracket) >= 2
    End If
  End If

End Function

Public Function IsControlEvent(ByVal strTest As String, _
                               Optional strCaseFix As String) As Boolean

  Dim I       As Long

  If bCtrlDescExists Then
    For I = LBound(ArrQActiveControlClass) To UBound(ArrQActiveControlClass)
      If TLibEventFinder(strTest, CStr(ArrQActiveControlClass(I)), strCaseFix) Then
        IsControlEvent = True
        Exit For
      End If
    Next I
  End If

End Function

Public Function IsControlProperty(ByVal strTest As String) As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    If Not IsNumeric(strTest) Then
      For I = LBound(ArrQActiveControlClass) To UBound(ArrQActiveControlClass)
        If ReferenceLibraryControlProperty(strTest, "Form") Then
          'CStr(ArrQActiveControlClass(I))) Then
          IsControlProperty = True
          Exit For
        End If
      Next I
    End If
  End If

End Function

Private Function IsStub(arrR As Variant) As Boolean

  Dim I As Long

  'Used to stop processing of stub procedures
  For I = 0 To UBound(arrR)
    If Left$(arrR(I), 7) = "'<STUB>" Then
      IsStub = True
      Exit For
    End If
  Next I

End Function

Private Function isTypeCastFunction(ByVal strCode As String) As Boolean

  Dim arrTmp As Variant

  'v2.8.9 updated the first test works for everything
  strCode = strCodeOnly(strCode)
  arrTmp = ExpandForDetection2(strCode) 'Split(ExpandForDetection(strCode))
  If UBound(arrTmp) > 1 Then
    isTypeCastFunction = arrTmp(UBound(arrTmp) - 1) = "As"
    '    If Not isTypeCastFunction Then
    '      'deal with type of format 'As Type()'
    '      If Right(strCode, 2) = RBracket & RBracket Then
    '        isTypeCastFunction = (UBound(arrTmp) - 4) = "As"
    '      End If
    '    End If
  End If

End Function

Private Function isUserEvent(ByVal strTest As String, _
                             ByVal CompName As String) As Boolean

  Dim I As Long

  If bEventDescExists Then
    For I = LBound(EventDesc) To UBound(EventDesc)
      If EventDesc(I).EScope = "Public" Then
        If SmartRight(strTest, EventDesc(I).EName) Then
          isUserEvent = True
          Exit For
        End If
       ElseIf EventDesc(I).EForm = CompName Then
        If EventDesc(I).EScope = "Private" Then
          If SmartRight(strTest, EventDesc(I).EName) Then
            isUserEvent = True
            Exit For
          End If
        End If
      End If
    Next I
  End If

End Function

Private Sub MultiLeftDelete(varSearch As Variant, _
                            ByVal CaseSensitive As Boolean, _
                            ParamArray Afind() As Variant)

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If CaseSensitive Then
      If Left$(varSearch, Len(FindIt)) = FindIt Then
        varSearch = Mid$(varSearch, Len(FindIt) + 1)
        Exit For 'Sub
      End If
     Else
      If Left$(varSearch, Len(FindIt)) = LCase$(FindIt) Then
        varSearch = Mid$(varSearch, Len(FindIt) + 1)
      End If
    End If
  Next FindIt

End Sub

Public Function ParamArrays(ByVal TestLine As String, _
                            ParamMode As ParamArrayType) As Variant

  Dim ParamStart As Long
  Dim strParam   As String
  Dim ArrParam   As Variant
  Dim I          As Long

  'convert parameters to an array of type Mode or empty array
  If ExtractCode(TestLine) Then
    ParamStart = GetLeftBracketPos(TestLine)
    If ParamStart Then
      strParam = Mid$(TestLine, ParamStart + 1)
      'v3.0.2 improved extractor
      If Right$(strParam, 2) = "()" Then
        strParam = Left$(strParam, Len(strParam) - 2)
      End If
      If InStrRev(strParam, RBracket) Then
        strParam = Left$(strParam, InStrRev(strParam, RBracket) - 1)
       Else
        strParam = vbNullString ' it was just the () of a no parameter procedure
      End If
      'v2.2.4 ignore 'Property Get X() as Long()' style which otherwise gets ') as Long(' as parameter
      If LenB(strParam) Then
        If Left$(strParam, 1) = ")" Then
          If Right$(strParam, 1) = "(" Then
            Exit Function
          End If
        End If
        ArrParam = Split(strParam, CommaSpace)
        Select Case ParamMode
         Case FullParam
          ParamArrays = ArrParam
         Case VariableOnlyParam
          For I = 0 To UBound(ArrParam)
            MultiLeftDelete ArrParam(I), True, "Optional ", "ParamArray ", "ByRef ", "ByVal "
          Next I
          '          'cope with 'Optional ByVal'
          For I = 0 To UBound(ArrParam)
            MultiLeftDelete ArrParam(I), True, "ByRef ", "ByVal ", "Optional ", "ParamArray "
          Next I
          For I = 0 To UBound(ArrParam)
            If TypeSuffixExists(ArrParam(I)) Then
              ArrParam(I) = Left$(ArrParam(I), Len(ArrParam(I)) - 1)
            End If
            ArrParam(I) = LeftWord(Trim$(ArrParam(I)))
            If Right$(ArrParam(I), 2) = "()" Then
              ArrParam(I) = Left$(ArrParam(I), Len(ArrParam(I)) - 2)
            End If
          Next I
         Case VariableOnlyParam2
          'this is a recursive call to cut down on coding
          ArrParam = ParamArrays(TestLine, VariableOnlyParam)
         Case UnTypedParam
          For I = 0 To UBound(ArrParam)
            If Get_As_Pos(ArrParam(I)) > 0 Then
              ' using type suffix so accept
              If Not TypeSuffixExists(ArrParam(I)) Then
                ArrParam(I) = vbNullString
              End If
            End If
          Next I
         Case SuffixTypeParam
          For I = 0 To UBound(ArrParam)
            If Get_As_Pos(ArrParam(I)) <> 0 Then
              ArrParam(I) = vbNullString
            End If
          Next I
          For I = 0 To UBound(ArrParam)
            If Not TypeSuffixExists(ArrParam(I)) Then
              ArrParam(I) = vbNullString
            End If
          Next I
         Case OptionalParam
          For I = 0 To UBound(ArrParam)
            If Not Left$(ArrParam(I), 9) = "Optional " Then
              ArrParam(I) = vbNullString
            End If
          Next I
         Case NoBy_Param
          For I = 0 To UBound(ArrParam)
            If InStr(ArrParam(I), "ByVal ") Or InStr(ArrParam(I), "ByRef ") Then
              ArrParam(I) = vbNullString
            End If
          Next I
         Case NoBy_Param2
          For I = 0 To UBound(ArrParam)
            If InStr(ArrParam(I), "ByVal ") Or InStr(ArrParam(I), "ByRef ") Then
              ArrParam(I) = SngSpace
            End If
          Next I
        End Select
        ParamArrays = Split(safe_Join(ArrParam, CommaSpace), CommaSpace)
      End If
    End If
  End If

End Function

Private Sub ParamaterBreakOut(ByVal strTest As String, _
                              strLeft As String, _
                              strParam As String, _
                              strRight As String, _
                              strRoutineName As String)

  Dim LeftBracket  As Long
  Dim RightBracket As Long

  LeftBracket = InStrCode(strTest, LBracket)
  RightBracket = InStrCodeRev(strTest, RBracket)
  If LeftBracket > 0 Then
    If RightBracket > LeftBracket Then
      strLeft = Left$(strTest, LeftBracket)
      strRight = Mid$(strTest, RightBracket)
      strRoutineName = Left$(strLeft, Len(strLeft) - 1)
      strRoutineName = Mid$(strRoutineName, InStrRev(strRoutineName, SngSpace) + 1)
      strParam = Mid$(strTest, LeftBracket + 1, RightBracket - LeftBracket - 1)
    End If
  End If

End Sub

Public Sub Parameter_Engine()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long

  On Error GoTo BugHit
  If Not bAborting Then
    'MOst of the following tests could be incorperated into a single code sweeper
    'but separating them out makes code clearer if slower.
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          With Comp
            If .CodeModule.CountOfLines Then
              WorkingMessage "Type Suffix Parameter", 1, 7
              TypeSuffixExpanderParameter .CodeModule
              WorkingMessage "Untyped Parameter", 2, 7
              UnTypedParameter .CodeModule
              WorkingMessage "Poor Parameter name", 3, 7
              BadNameParameter .CodeModule
              WorkingMessage "Unused Parameter", 4, 7
              UnusedParameter .CodeModule
              WorkingMessage "ByVal test", 5, 7
              ByValParameter .CodeModule
              WorkingMessage "Function to Sub", 6, 7
              Function2Sub .CodeModule
              WorkingMessage "Poorly Named Parameter", 7, 7
              TypeCastFunction .CodeModule
            End If
          End With 'Comp
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "Parameter_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Function ParameterExpandDo(cMod As CodeModule, _
                                   arrCurRoutine As Variant, _
                                   ByVal strTest As String) As String

  Dim LeftBracket  As Long
  Dim Parameters   As String
  Dim RoutineName  As String
  Dim LBit         As String
  Dim Rbit         As String
  Dim TmpA         As Variant
  Dim I            As Long
  Dim J            As Long
  Dim Guard        As String
  Dim MissingType  As Boolean
  Dim MissingCount As Long
  Dim AsType       As String
  Dim Hit          As Boolean
  Dim strOrig      As String
  Dim ModuleNumber As Long
  Dim CompName     As String

  CompName = cMod.Parent.Name
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  strOrig = strTest
  LeftBracket = InStrCode(strTest, LBracket)
  If LeftBracket > 0 Then
    If InStrCodeRev(strTest, RBracket) > LeftBracket Then
      ParamaterBreakOut strTest, LBit, Parameters, Rbit, RoutineName
      For I = 1 To Len(Parameters)
        If Mid$(Parameters, I, 1) = "," Then
          If InLiteral(Parameters, I) Then
            Parameters = Left$(Parameters, I - 1) & "~*^" & Mid$(Parameters, I + 1)
          End If
        End If
      Next I
      Guard = Parameters
      TmpA = Split(Parameters, CommaSpace)
      Parameters = Join(TmpA, ",")
      TmpA = Split(Parameters, ",")
      ParameterUnTyped TmpA, MissingType, MissingCount
      'rewrite ParamArray and Optional parameters
      If MissingType Then
        ParameterInternalEvidenceExpander TmpA, Hit, CompName
        If Hit Then
          Parameters = Join(TmpA, ",")
          ParameterUnTyped TmpA, MissingType, MissingCount
        End If
      End If
      If MissingType Then
        For J = LBound(TmpA) To UBound(TmpA)
          If Get_As_Pos(TmpA(J)) = 0 Then
            AsType = GetDimTypeFromInternalEvidence(ModuleNumber, arrCurRoutine, TmpA(J), CompName)
            If Left$(AsType, 3) = "<-1" Then
              TmpA(J) = Left$(TmpA(J), Len(TmpA(J)) - 1)
              AsType = Mid$(AsType, 4)
            End If
            If LenB(AsType) Then
              'v2.29 resturcture to work porperly
              If InStr(TmpA(J), "Optional") Then
                'catch stuff like Optional Fred = True , Optional Wilma = "Hello")
                AsType = " As " & TypeFromValue(RightWord(TmpA(J)))
                If Len(AsType) > 4 Then
                  TmpA(J) = Replace$(TmpA(J), EqualInCode, AsType & EqualInCode, 1, 1)
                  Hit = True
                End If
               Else
                TmpA(J) = TmpA(J) & AsType
                Hit = True
              End If
            End If
          End If
        Next J
        Parameters = Join(TmpA, ",")
      End If
      If Hit Then
        strTest = LBit & Parameters & Rbit
      End If
      ParameterUnTyped TmpA, MissingType, MissingCount
      If MissingType Then
        'count  missing Types
        GetCallingRoutine cMod, RoutineName, Parameters, MissingCount, cMod.Parent.Name
        Parameters = Safe_Replace(Parameters, "~*^", ", ")
        If Guard <> Parameters Then
          strTest = LBit & Parameters & Rbit & vbNewLine & _
           WARNING_MSG & "Untyped Parameters use Variants which use excessive memory." & IIf(Xcheck(XPrevCom), vbNewLine & _
           PREVIOUSCODE_MSG & Guard, vbNullString)
         Else
          strTest = strTest & vbNewLine & _
           WARNING_MSG & "Untyped Parameters use Variants which use excessive memory."
        End If
       Else
        If Hit Then
          strTest = strTest & vbNewLine & _
           WARNING_MSG & "Untyped Parameters auto-Typed." & IIf(UsingDefTypes, vbNullString, " May be not be correct.") & IIf(Xcheck(XPrevCom), vbNewLine & _
           PREVIOUSCODE_MSG & strOrig, vbNullString)
        End If
      End If
    End If
  End If
  ParameterExpandDo = strTest

End Function

Private Sub ParameterInternalEvidenceExpander(TArray As Variant, _
                                              Hit As Boolean, _
                                              ByVal CompName As String)

  Dim I      As Long
  Dim J      As Long
  Dim TmpC   As Variant
  Dim strTmp As String

  'get parameter type from internal evidence of parameters
  ''ParamArray' must be Variant,
  ''Optional' may carry clues in its DefaultValue(if any)
  For I = 0 To UBound(TArray)
    If Get_As_Pos(TArray(I)) = 0 Then
      If Left$(TArray(I), 10) = "ParamArray" Then
        TArray(I) = TArray(I) & " As Variant"
        Hit = True
       ElseIf MultiLeft(TArray(I), True, "Optional", "ByVal Optional", "Optional ByVal") Then
        TmpC = Split(ExpandForDetection(TArray(I)))
        For J = 0 To UBound(TmpC)
          If TmpC(J) = "=" Then
            strTmp = TypeSuffix2String(TmpC(J - 1))
            If LenB(strTmp) Then
              TmpC(J - 1) = Left$(TmpC(J - 1), Len(TmpC(J - 1)) - 1) & " As " & strTmp
              TArray(I) = Join(TmpC)
              Hit = True
              Exit For
            End If
          End If
        Next J
       ElseIf Get_As_Pos(TArray(I)) = 0 Then
        TmpC = Split(ExpandForDetection(TArray(I)))
        If UBound(TmpC) > 0 Then
          For J = 0 To UBound(TmpC)
            If TmpC(J) = "=" Then
              If LenB(InternalEvidenceType(TmpC(J + 1), CompName)) Then
                TmpC(J - 1) = TmpC(J - 1) & InternalEvidenceType(TmpC(J + 1), CompName)
                TArray(I) = Join(TmpC)
                Hit = True
                Exit For
              End If
            End If
          Next J
         ElseIf UBound(TmpC) = 0 Then
          strTmp = TypeSuffix2String(TmpC(0))
          If LenB(strTmp) Then
            TmpC(0) = Left$(TmpC(0), Len(TmpC(0)) - 1) & " As " & strTmp
            TArray(I) = Join(TmpC)
            Hit = True
            Exit For
          End If
        End If
      End If
    End If
  Next I

End Sub

Private Sub ParameterUnTyped(arrParameters As Variant, _
                             bMissing As Boolean, _
                             MissCount As Long)

  Dim I As Long

  bMissing = False
  MissCount = 0
  For I = 0 To UBound(arrParameters)
    If Get_As_Pos(arrParameters(I)) = 0 Then
      bMissing = True
      MissCount = MissCount + 1
    End If
  Next I

End Sub

Public Sub ParamLineContFix(ByVal Range As Long, _
                            ByVal TopPos As Long, _
                            Arr As Variant, _
                            ByVal strTarget As String)

  Dim bDummy As Boolean
  Dim I      As Long

  If Range Then
    For I = 1 To Range
      Arr(TopPos) = Left$(Arr(TopPos), Len(Arr(TopPos)) - 1) + Trim$(Arr(TopPos + I))
      Arr(TopPos + I) = vbNullString
    Next I
  End If
  LineContinuationForVeryLongLines strTarget, ContMark & vbNewLine, bDummy

End Sub

Private Function PrefixFromType(ByVal varParamName As Variant, _
                                ByVal varFullParam As Variant) As String

  Dim strType   As String
  Dim strPrefix As String
  Dim IDNo      As Long

  If Has_AS(varFullParam) Then
    varFullParam = ExpandForDetection(varFullParam)
    strType = GetType(varFullParam)
    IDNo = QSortArrayPos(StandardTypes, strType)
    If IDNo > -1 Then
      strPrefix = StandardPreFix(IDNo)
     Else
      strPrefix = FakeHungarian(strType)
    End If
    If Len(strPrefix) Then
      'v2.9.6 copes with some very rare problems with renaming parapeters
      'example FillStyle As rgnFillStyle would have become rgnFillStyle As FillStyle
      'now becomes myFillStyle As rgnFillStyle
      If strType <> strPrefix & Ucase1st(varParamName) Then
        PrefixFromType = strPrefix & Ucase1st(varParamName)
       Else
        If strType <> "my" & Ucase1st(varParamName) Then
          PrefixFromType = "my" & Ucase1st(varParamName)
         Else
          PrefixFromType = "my" & strPrefix & Ucase1st(varParamName)
        End If
      End If
    End If
  End If

End Function

Private Function safe_Join(Arr As Variant, _
                           Optional ByVal Delim As String = SngSpace) As String

  Dim I      As Long
  Dim strTmp As String

  For I = LBound(Arr) To UBound(Arr)
    If Len(Arr(I)) Then
      strTmp = strTmp & IIf(Len(strTmp), Delim, vbNullString) & Arr(I)
    End If
  Next I
  safe_Join = strTmp

End Function

Private Function SafeParamReplace(ByVal FPos As Long, _
                                  ByVal cde As String, _
                                  ByVal strF As String) As String

  If FPos > 1 Then
    If FPos + Len(strF) > Len(cde) Then
      SafeParamReplace = Mid$(cde, FPos - 1, 1) <> "." And Mid$(cde, FPos - 1, 1) <> "_" And Mid$(cde, FPos + Len(strF), 1) <> "_"
     Else
      SafeParamReplace = Mid$(cde, FPos - 1, 1) <> "." And Mid$(cde, FPos - 1, 1) <> "_"
    End If
  End If

End Function

Private Sub SetToNullString(ParamArray stuff() As Variant)

  Dim I As Long

  For I = LBound(stuff) To UBound(stuff)
    stuff(I) = vbNullString
  Next I

End Sub

Private Sub TypeCastFunction(cMod As CodeModule)

  Dim TopOfRoutine As Long
  Dim lngdummy     As Long   ' this is a junk variable
  Dim FunctIsFunct As Boolean
  Dim L_CodeLine   As String
  Dim CommentStore As String
  Dim ArrMember    As Variant
  Dim Member       As Long
  Dim ArrRoutine   As Variant
  Dim UpDated      As Boolean
  Dim Evid         As String
  Dim MemberCount  As Long
  Dim Rname        As String
  Dim Strnew       As String
  Dim I            As Long
  Dim ModuleNumber As Long
  Dim CompName     As String

  CompName = cMod.Parent.Name
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  ArrMember = GetMembersArray(cMod)
  MemberCount = UBound(ArrMember)
  If MemberCount > 0 Then
    For Member = 1 To MemberCount
      ArrRoutine = Split(ArrMember(Member), vbNewLine)
      MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
      If UBound(ArrRoutine) > -1 Then
        If Not IsStub(ArrRoutine) Then
          L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, lngdummy, Rname)
          If ExtractCode(L_CodeLine, CommentStore) Then
            If InstrAtPosition(L_CodeLine, "Function", ipLeftOr2ndOr3rd, True) Then
              If Not IsVBControlRoutine(Rname) Then
                If Not isTypeCastFunction(L_CodeLine) Then
                  Evid = vbNullString
                  Strnew = vbNullString
                  FunctIsFunct = RoutineSearch(ArrRoutine, Rname, TopOfRoutine, ipAny, False)
                  If Not FunctIsFunct Then
                    If InQSortArray(TypeSuffixArray, Right$(Rname, 1)) Then
                      'test for typesuffixed versions (v2.3.3 wrong test fixed) Thanks Will Barden
                      FunctIsFunct = RoutineSearch(ArrRoutine, Left$(Rname, Len(Rname) - 1), TopOfRoutine, ipAny, False)
                    End If
                  End If
                  If InQSortArray(ImplementsArray, ModDesc(ModuleNumber).MDName) Then
                    'leave implements classes alone
                    FunctIsFunct = True
                  End If
                  If FunctIsFunct Then
                    If InQSortArray(TypeSuffixArray, Right$(Rname, 1)) Then
                      Evid = " As " & AsTypeArray(QSortArrayPos(TypeSuffixArray, Right$(Rname, 1)))
                      If Len(Evid) Then
                        Strnew = Replace$(L_CodeLine, Rname, Left$(Rname, Len(Rname) - 1), 1) & Evid & CommentStore
                      End If
                     Else
                      Evid = GetFunctionTypeFromInternalEvidence(ModuleNumber, ArrRoutine, Rname, CompName)
                      If Len(Evid) Then
                        Strnew = L_CodeLine & Evid & CommentStore
                      End If
                    End If
                  End If
                  If LenB(Strnew) Then
                    Select Case FixData(ConvertFunction2Sub).FixLevel
                     Case CommentOnly
                      ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, SUGGESTION_MSG & "Function should be Type-cast using '" & Evid, MAfter, UpDated)
                     Case FixAndComment, JustFix ' never do it without telling user
                      ArrRoutine(TopOfRoutine) = Marker(Strnew, WARNING_MSG & "Function Typed automatically(May not be correct Type)", MAfter, UpDated)
                      If InQSortArray(TypeSuffixArray, Right$(Rname, 1)) Then
                        For I = TopOfRoutine To UBound(ArrRoutine)
                          If InStr(ArrRoutine(I), Rname) Then
                            ArrRoutine(I) = Replace$(ArrRoutine(I), Rname, Left$(Rname, Len(Rname) - 1), 1)
                          End If
                        Next I
                      End If
                    End Select
                    ArrMember(Member) = Join(ArrRoutine, vbNewLine)
                   Else
                    If Not IsArrayReturningFunction(L_CodeLine) Then
                      Select Case FixData(ConvertFunction2Sub).FixLevel
                       Case CommentOnly
                        ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, SUGGESTION_MSG & "Function should be Type-cast using '" & Evid, MAfter, UpDated)
                       Case FixAndComment, JustFix ' never do it without telling user
                        ArrRoutine(TopOfRoutine) = Marker(L_CodeLine, SUGGESTION_MSG & "Function should be Type-cast but Code Fixer cannot determine the Type to apply." & WARNING_MSG & "Function will return Variant value.", MAfter, UpDated)
                      End Select
                      ArrMember(Member) = Join(ArrRoutine, vbNewLine)
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next Member
    ReWriteMembers cMod, ArrMember, UpDated
  End If

End Sub

Private Function TypeFromValue(VarVal As Variant) As String

  If IsNumeric(VarVal) Then
    TypeFromValue = "Long"
   ElseIf VarVal = "True" Or VarVal = "False" Then
    TypeFromValue = "Boolean"
   ElseIf InStr(VarVal, DQuote) Then
    TypeFromValue = "String"
   ElseIf VarVal = "vbNullString" Then
    TypeFromValue = "String"
    'TypeFRomValue = "Variant"
  End If

End Function

Private Sub TypeSuffixExpanderParameter(cMod As CodeModule)

  Dim ModuleNumber  As Long
  Dim L_CodeLine    As String
  Dim MyStr         As String
  Dim UpDated       As Boolean
  Dim ArrMember     As Variant
  Dim ArrRoutine    As Variant
  Dim Member        As Long
  Dim MemberCount   As Long
  Dim TopOfRoutine  As Long
  Dim LineContRange As Long
  Dim strDummy      As String
  Dim ArrParam      As Variant
  Dim ArrParamX     As Variant
  Dim strComT       As String
  Dim I             As Long

  'Type cast parameters , update type suffixes
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ParameterExpandTypeSuffix) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        If UBound(ArrRoutine) > -1 Then
          L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, strDummy)
          If LenB(L_CodeLine) Then
            'v2.3.1 improved parameter updater
            ArrParam = ParamArrays(L_CodeLine, SuffixTypeParam)
            If IsArray(ArrParam) Then
              MyStr = L_CodeLine
              If UBound(ArrParam) > -1 Then
                ReDim ArrParamX(UBound(ArrParam)) As Variant
                For I = LBound(ArrParam) To UBound(ArrParam)
                  ArrParamX(I) = TypeSuffixExtender(ArrParam(I))
                Next I
                For I = LBound(ArrParam) To UBound(ArrParam)
                  If ArrParam(I) <> ArrParamX(I) Then
                    ExtractCode ArrParamX(I), strComT
                    ArrParamX(I) = Replace$(ArrParamX(I), vbNewLine, vbNullString)
                    MyStr = Replace$(MyStr, ArrParam(I), ArrParamX(I)) & vbNewLine & strComT
                  End If
                Next I
                'MyStr = TypeSuffixExtender(L_CodeLine)
                If MyStr <> L_CodeLine Then
                  ParamLineContFix LineContRange, TopOfRoutine, ArrRoutine, MyStr
                  ArrRoutine(TopOfRoutine) = MyStr & IIf(Xcheck(XPrevCom), vbNewLine & _
                   PREVIOUSCODE_MSG & L_CodeLine, vbNullString)
                  UpDated = True
                  ArrMember(Member) = Join(ArrRoutine, vbNewLine)
                End If
                AddNfix ParameterExpandTypeSuffix
                ArrMember(Member) = Join(ArrRoutine, vbNewLine)
              End If
            End If
          End If
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, UpDated
    End If
  End If

End Sub

Public Function Ucase1st(varIn As Variant) As String

  Ucase1st = UCase$(Left$(varIn, 1)) & Mid$(varIn, 2)

End Function

Private Sub UnTypedParameter(cMod As CodeModule)

  Dim ModuleNumber  As Long
  Dim L_CodeLine    As String
  Dim MyStr         As String
  Dim UpDated       As Boolean
  Dim ArrMember     As Variant
  Dim ArrRoutine    As Variant
  Dim Member        As Long
  Dim MemberCount   As Long
  Dim TopOfRoutine  As Long
  Dim LineContRange As Long
  Dim strDummy      As String

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ParameterNoType) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        If UBound(ArrRoutine) > -1 Then
          L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, strDummy)
          If LenB(L_CodeLine) Then
            If IsArray(ParamArrays(L_CodeLine, UnTypedParam)) Then
              ArrRoutine = Split(ArrMember(Member), vbNewLine) 'redo in case first part hit
              L_CodeLine = ArrRoutine(TopOfRoutine)
              MyStr = ParameterExpandDo(cMod, ArrRoutine, L_CodeLine)
              If MyStr <> L_CodeLine Then
                ParamLineContFix LineContRange, TopOfRoutine, ArrRoutine, MyStr
                ArrRoutine(TopOfRoutine) = MyStr
                UpDated = True
                AddNfix ParameterNoType
                ArrMember(Member) = Join(ArrRoutine, vbNewLine)
              End If
            End If
          End If
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, UpDated
    End If
  End If

End Sub

Private Sub UnusedParameter(cMod As CodeModule)

  Dim ModuleNumber  As Long
  Dim strCmpName    As String
  Dim L_CodeLine    As String
  Dim ArrMember     As Variant
  Dim Member        As Long
  Dim ArrRoutine    As Variant
  Dim UpDated       As Boolean
  Dim MUpdated      As Boolean
  Dim MemberCount   As Long
  Dim strMsg        As String
  Dim TopOfRoutine  As Long
  Dim LineContRange As Long
  Dim Rname         As String

  strCmpName = cMod.Parent.Name
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ParameterDead_CXX) Then
    'v2.9.7 Don't do to implements classes Thanks Ian K
    If Not InQSortArray(ImplementsArray, strCmpName) Then
      ArrMember = GetMembersArray(cMod)
      MemberCount = UBound(ArrMember)
      If MemberCount > 0 Then
        For Member = 1 To MemberCount
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
          If UBound(ArrRoutine) > -1 Then
            If Not IsStub(ArrRoutine) Then
              L_CodeLine = GetRoutineDeclaration(ArrRoutine, TopOfRoutine, LineContRange, Rname)
              If Not IsVBControlRoutine(Rname) And Not RoutineNameIsVBGenerated(Rname, cMod.Parent) Then
                UnusedParameterTest L_CodeLine, TopOfRoutine + LineContRange, ArrRoutine, strMsg
                If Len(strMsg) Then
                  Select Case FixData(ParameterDead_CXX).FixLevel
                   Case CommentOnly
                    If InQSortArray(ImplementsArray, strCmpName) Then
                      'v3.0.1 problem with long parameter sets fixed' thanks Marcel
                      ArrRoutine(TopOfRoutine + LineContRange) = Marker(ArrRoutine(TopOfRoutine + LineContRange), WARNING_MSG & "Unused Parameter" & IIf(InStr(strMsg, CommaSpace), "s", vbNullString) & strInSQuotes(strMsg, True) & vbNewLine & _
                       "This is an Interface class; Parameters must be left in place.", MAfter, UpDated)
                     Else
                      ArrRoutine(TopOfRoutine + LineContRange) = Marker(ArrRoutine(TopOfRoutine + LineContRange), WARNING_MSG & "Unused Parameter" & IIf(InStr(strMsg, CommaSpace), "s", vbNullString) & strInSQuotes(strMsg, True) & "could be removed.", MAfter, UpDated)
                    End If
                    AddNfix ParameterDead_CXX
                  End Select
                  strMsg = vbNullString
                End If
                UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
              End If
            End If
          End If
        Next Member
        ReWriteMembers cMod, ArrMember, MUpdated
      End If
    End If
  End If

End Sub

Private Sub UnusedParameterTest(ByVal TestLine As String, _
                                ByVal StartPos As Long, _
                                Arr As Variant, _
                                strMsg As String, _
                                Optional LFix As String)

  Dim I            As Long
  Dim J            As Long
  Dim ArrParam     As Variant
  Dim ArrParamFull As Variant
  Dim strParam     As String
  Dim L_CodeLine   As String
  Dim hits         As Long

  strMsg = vbNullString
  If ExtractCode(TestLine) Then
    If GetLeftBracketPos(TestLine) Then
      strParam = TestLine
      ArrParam = ParamArrays(strParam, VariableOnlyParam2)
      ArrParamFull = ParamArrays(strParam, FullParam)
      If ArrayHasContents(ArrParam) Then
        For I = StartPos + 1 To UBound(Arr)
          If Not JustACommentOrBlank(Arr(I)) Then
            L_CodeLine = ExpandForDetection(Arr(I))
            For J = 0 To UBound(ArrParam)
              If Len(ArrParam(J)) Then
                If InstrAtPosition(L_CodeLine, ArrParam(J), ipAny, True) Then
                  ArrParamFull(J) = vbNullString
                  If LenB(ArrParam(J)) Then
                    'v3.0.2 speed improvment only count if it isn't already blank
                    hits = hits + 1
                  End If
                  ArrParam(J) = vbNullString
                End If
                If SmartLeft(L_CodeLine, ArrParam(J)) Then
                  ArrParamFull(J) = vbNullString
                  If LenB(ArrParam(J)) Then
                    hits = hits + 1
                  End If
                  ArrParam(J) = vbNullString
                End If
              End If
            Next J
          End If
          If hits = UBound(ArrParam) + 1 Then
            hits = 0
            Exit For
          End If
        Next I
        strMsg = safe_Join(ArrParamFull, CommaSpace)
        If Len(strMsg) Then
          LFix = TestLine
          For I = 0 To UBound(ArrParamFull)
            If Len(ArrParamFull(I)) Then
              LFix = ReplaceArray(LFix, CommaSpace & ArrParamFull(I), vbNullString, ArrParamFull(I) & CommaSpace, vbNullString, ArrParamFull(I), vbNullString)
            End If
          Next I
          strMsg = " " & strMsg & " "
          strMsg = Replace$(strMsg, " Optional ", " ")
          strMsg = Replace$(strMsg, " ByVal ", " ")
          strMsg = Replace$(strMsg, " ByRef ", " ")
          ArrParam = Split(Trim$(strMsg), ",")
          For I = 0 To UBound(ArrParam)
            ArrParam(I) = LeftWord(Trim$(ArrParam(I)))
          Next I
          strMsg = Join(ArrParam, ", ")
        End If
      End If
    End If
  End If

End Sub

Private Sub WholeWordReplacerParameter(cde As String, _
                                       ByVal strF As String, _
                                       ByVal StrRep As String, _
                                       bUpdated As Boolean, _
                                       Optional bIgnoreUnderScore As Boolean = False)

  Dim Cmpr As VbCompareMethod
  Dim FPos As Long

  If Len(cde) Then
    If Len(strF) Then
      If InStr(cde, strF) Then
        Cmpr = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
        FPos = InStr(1, cde, strF, Cmpr)
        Do While InStr(FPos, cde, strF, Cmpr)
          If InCode(cde, FPos) Then
            If IsWholeWord(cde, strF, FPos, bIgnoreUnderScore) Then
              'v2.4.4 Thanks Aaron Spivey
              'deals with the problem of strF being part of a larger word with underscores
              'EG Hwnd => lngHwnd in this
              'SetWindowPos Hwnd, IIf(isTopMost, Hwnd_TOPMOST,
              If IsRealWord(cde, strF) Then
                If FPos > 1 Then
                  If SafeParamReplace(FPos, cde, strF) Then
                    cde = Left$(cde, FPos - 1) & StrRep & Mid$(cde, FPos + Len(strF))
                    bUpdated = True
                  End If
                 Else
                  cde = Left$(cde, FPos - 1) & StrRep & Mid$(cde, FPos + Len(strF))
                End If
              End If
            End If
          End If
          FPos = InStr(FPos + 1, cde, strF, Cmpr)
          If FPos = 0 Then
            Exit Do
          End If
        Loop
      End If
    End If
  End If

End Sub

Private Function WriteType(arrModule As Variant, _
                           lineMod As Long, _
                           arrLine As Variant, _
                           strType As String, _
                           StrReturn As String) As Boolean

  If SmartLeft(arrModule(lineMod), strType & " " & arrLine(2)) Then
    StrReturn = Mid$(arrModule(lineMod), Len(strType & " " & arrLine(2)) + 1)
    WriteType = True
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:16:48 AM) 21 + 1626 = 1647 Lines Thanks Ulli for inspiration and lots of code.

