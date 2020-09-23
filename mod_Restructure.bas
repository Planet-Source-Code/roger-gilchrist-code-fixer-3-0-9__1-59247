Attribute VB_Name = "mod_Restructure"

Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public bWithSuggested                       As Boolean
Private blnStaticInTypeDef                  As Boolean
Private blnStaticInEnumDef                  As Boolean
Private blnStaticInConditionalCompile       As Boolean
'
'an 'a' in the ID part indicates that a fix exits
Public Enum eExits
  eExitUNknown_ID0 '=0
  eExitIf_Then_Exit_EndIf_ID1
  eExitAs1stCode_ID1a
  eExitIf_Then_Exit_EndIf_Code_ID2
  eExitAs1stCodeComplex_ID2a
  eExitIf_Then_Exit_Else_Code_EndIf_ID3
  eExitErrorTrapShield_ID4
  eExitDeepStructure_ID5
  eExitProc2ExitLoop_ID8a
  eExitFromWith_ID9
  eExitFromFors_ID10
  eExitGEneric_ID11
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eExitIf_Then_Exit_EndIf_ID1, eExitAs1stCode_ID1a, eExitIf_Then_Exit_EndIf_Code_ID2, eExitAs1stCodeComplex_ID2a
Private eExitIf_Then_Exit_Else_Code_EndIf_ID3, eExitErrorTrapShield_ID4, eExitDeepStructure_ID5, eExitProc2ExitLoop_ID8a
Private eExitFromWith_ID9, eExitFromFors_ID10, eExitGEneric_ID11
#End If
Public Enum EnumStruct
  IfStruct
  SelectStruct
  WithStruct
  TypeStruct
  EnumStruct
  LoopStruct
  ForStruct
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private IfStruct, SelectStruct, WithStruct, TypeStruct, EnumStruct, LoopStruct, ForStruct
#End If
Public Const StrPlus                        As String = " + "

Public Function CallRemoval(ByVal varCode As Variant) As String

  Dim TPos       As Long
  Dim TRBPos     As Long
  Dim TLBPos     As Long
  Dim StrBalance As String
  Dim J          As Long
  Dim strComment As String

  'v2.2.9 added comment separation for safer use
  CallRemoval = varCode
  'v2.3.0 stop word ending in Call from being hit
  If InstrAtPosition(varCode, "Call", IpLeft) Then
    TPos = InStr(varCode, "Call ") + 5
    If TPos Then
      ExtractCode varCode, strComment
      TRBPos = InStrRev(varCode, ")")
      If TRBPos = Len(varCode) Then
        'v2.4.1 Thanks Paul Caton this will stop it hitting on 'Call Form(i).Show'
        For J = TRBPos To 1 Step -1 'Len(varCode)
          StrBalance = Mid$(varCode, J, 1) & StrBalance
          If Not CountSubStringImbalance(StrBalance, "(", ")") Then
            TLBPos = J
            Exit For
          End If
        Next J
        Mid$(varCode, TRBPos, 1) = " "
        Mid$(varCode, TLBPos, 1) = " "
      End If
      'v2.2.7 fixed so it gets 'Call ProcedureName' where no parameters are used
      Mid$(varCode, TPos - 5, 5) = "     "
      CallRemoval = Trim$(varCode) & strComment
    End If
  End If

End Function

Private Sub CallRemovalEngine(cMod As CodeModule)

  Dim strTest2  As String
  Dim StartLine As Long

  'v2.4.1 NEW version of old fix
  With cMod
    Do While .Find("Call ", StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), "Call ")) Then
          If Not HasLineCont(.Lines(StartLine, 1)) Then
            strTest2 = CallRemoval(.Lines(StartLine, 1))
            If strTest2 <> .Lines(StartLine, 1) Then
              Select Case FixData(UnNeededCall).FixLevel
               Case CommentOnly
                .InsertLines GetInsertionPoint(cMod, StartLine), WARNING_MSG & "Unneeded 'Call' command could be removed."
               Case FixAndComment
                .ReplaceLine StartLine, strTest2
                .InsertLines GetInsertionPoint(cMod, StartLine), WARNING_MSG & "Unneeded 'Call' command removed."
              End Select
            End If
          End If
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Sub CaseIIfEngine(cMod As CodeModule, _
                          ByVal LngMode As Long)

  Dim strTest1         As String
  Dim strTest2         As String
  Dim StartLine        As Long
  Dim StrNewCode       As String
  Dim StrCom           As String
  Dim lngCodeLineNo(5) As Long
  Dim J                As Long
  Dim strVal           As String
  Dim strVal2          As String

  'v2.5.0 NEW FIX converts
  '0 Select Case X
  '1 Case True
  '2   Y = Val
  '3 Case False
  '4   Y = Val2
  '5 End Select
  'to collapsed form
  ' y= IIF(X, Val, Val2)
  ''lngMode = 0 line 2 is Y = True
  ''lngMode = 1 line 2 is Y = False
  'This procedure is triggered by detecting line 2 then searching around it for matching structures 1,3,4,5
  'any eol comments are collected and then written after the collapsed line
  '
  Select Case LngMode
   Case 0
    strTest1 = "True"
    strTest2 = "False"
   Case 1
    strTest2 = "True"
    strTest1 = "False"
  End Select
  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find("Case " & strTest1, StartLine, 1, -1, -1)
      If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), strTest1)) Then
        MemberMessage "", StartLine, .CountOfLines
        'get the real code lines
        lngCodeLineNo(0) = GetPreviousCodeLine(StartLine, cMod)         '0 Select Case X
        lngCodeLineNo(1) = StartLine                                    '1 Case True/False
        lngCodeLineNo(2) = GetNextCodeLine(lngCodeLineNo(1), cMod)      '2 Y= Val
        lngCodeLineNo(3) = GetNextCodeLine(lngCodeLineNo(2), cMod)      '3 Case False/True/Else
        lngCodeLineNo(4) = GetNextCodeLine(lngCodeLineNo(3), cMod)      '4  Y= Val2
        lngCodeLineNo(5) = GetNextCodeLine(lngCodeLineNo(4), cMod)      '5 End Select
        'test all lines are code (0 means no line found)
        For J = 0 To 5
          If lngCodeLineNo(J) = 0 Then
            GoTo Skip
          End If
        Next J
        'test line 0 conforms to target code
        If Left$(strCodeOnly(.Lines(lngCodeLineNo(0), 1)), 12) = "Select Case " Then
          'test line 5 conforms to target code
          If strCodeOnly(.Lines(lngCodeLineNo(5), 1)) = "End Select" Then
            'test lines 2 and 4 conform to target code
            If LeftWord(.Lines(lngCodeLineNo(2), 1)) = LeftWord(.Lines(lngCodeLineNo(4), 1)) Then
              If InStr(.Lines(lngCodeLineNo(2), 1), " = ") And InStr(.Lines(lngCodeLineNo(4), 1), " = ") Then
                If strCodeOnly(.Lines(lngCodeLineNo(3), 1)) = "Case " & strTest2 Or strCodeOnly(.Lines(lngCodeLineNo(3), 1)) = "Case Else" Then
                  StrCom = vbNullString
                  StrNewCode = strCodeOnly(.Lines(lngCodeLineNo(0), 1))
                  StrNewCode = LeftWord(.Lines(lngCodeLineNo(2), 1))
                  StrNewCode = StrNewCode & " = IIf(" & Replace$(strCodeOnly(.Lines(lngCodeLineNo(0), 1)), "Select Case ", " ") & ", "
                  strVal = strCodeOnly(.Lines(lngCodeLineNo(2), 1))
                  strVal = Replace$(strVal, LeftWord(strVal) & " = ", vbNullString)
                  strVal2 = strCodeOnly(.Lines(lngCodeLineNo(4), 1))
                  strVal2 = Replace$(strVal2, LeftWord(strVal2) & " = ", vbNullString)
                  StrNewCode = StrNewCode & strVal & ", " & strVal2 & ")"
                  ' extract and comments and set them aside for later reinsertionn after the collapse
                  MultilineCommentExtractor cMod, lngCodeLineNo(0), lngCodeLineNo(5), StrCom
                  'deal with all other cases
                  InsertNewCodeComment cMod, lngCodeLineNo(0), lngCodeLineNo(5) - lngCodeLineNo(0) + 1, StrNewCode, WARNING_MSG & "Select Case structure being used to do IIf  collapsed." & IIf(InStr(StrNewCode, " = Not("), vbNewLine & _
                   SUGGESTION_MSG & "You may be able to simplify the code further by coding out the 'Not'", vbNullString) & IIf(Len(StrCom), vbNewLine & _
                   StrCom, vbNullString), True     ' Xcheck(XPrevCom)
                  StrCom = vbNullString
                End If
              End If
            End If
          End If
        End If
      End If
Skip:
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop
  End With

End Sub

Private Sub CaseIsEngine(cMod As CodeModule)

  Dim StartLine  As Long

  'v2.4.1 NEW version of old fix
  With cMod
    Do While .Find("Case Is =", StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), "Case Is =")) Then
          Select Case FixData(UneededCaseIs).FixLevel
           Case CommentOnly
            InsertNewCodeComment cMod, StartLine, 1, .Lines(StartLine, 1), WARNING_MSG & "Unneeded 'Case Is = ' could be replaced with 'Case '."
           Case FixAndComment
            InsertNewCodeComment cMod, StartLine, 1, Replace$(.Lines(StartLine, 1), "Case Is = ", "Case ", 1, 1), WARNING_MSG & "Unneeded 'Case Is = ' replaced with 'Case '."
          End Select
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Sub CaseOfRoutineFix(cMod As CodeModule)

  Dim arrMembers     As Variant
  Dim arrLine        As Variant
  Dim I              As Long
  Dim M              As Long
  Dim UpDated        As Boolean
  Dim MUpdated       As Boolean
  Dim MaxFactor      As Long
  Dim CaseUpdate     As Boolean
  Dim CaseFix        As String
  Dim TopOfRoutine   As Long
  Dim LineContRange  As Long
  Dim Rname          As String
  Dim L_CodeLine     As String
  Dim strR           As String
  Dim strL           As String
  Dim strCase        As String
  Dim ModuleNumber   As Long
  Dim strCorrect     As String
  Dim bImplementsMsg As Boolean   'v3.0.1 detect incorrect case and select message to use

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, RoutineCaseFix) Then
    arrMembers = GetMembersArray(cMod)
    MaxFactor = UBound(arrMembers)
    UpDated = False
    If MaxFactor > -1 Then
      For I = 1 To MaxFactor
        If Len(arrMembers(I)) Then
          arrLine = Split(arrMembers(I), vbNewLine)
          MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
          L_CodeLine = GetRoutineDeclaration(arrLine, TopOfRoutine, LineContRange, Rname)
          If InStr(Rname, "_") Then
            strR = strGetRightOf(Rname, "_")
            strL = strGetLeftOf(Rname, "_")
            If IsControlEvent(strR, strCase) Then
              For M = LBound(CntrlDesc) To UBound(CntrlDesc)
                If MultiLeft(Rname, False, CntrlDesc(M).CDName) Then
                  If LCase$(Rname) = LCase$(CntrlDesc(M).CDName & "_" & strR) Then
                    If Rname <> CntrlDesc(M).CDName & "_" & strCase Then
                      CaseUpdate = True
                      bImplementsMsg = False
                      CaseFix = CntrlDesc(M).CDName & "_" & strCase
                    End If
                  End If
                  Exit For
                End If
              Next M
             ElseIf WrongImplementsCase(strL, strCorrect) Then
              'v3.0.1 detect incorrect case for implements procedures. thanks Ian K
              CaseUpdate = True
              bImplementsMsg = True
              CaseFix = Replace$(Rname, strL, strCorrect)
            End If
          End If
          If CaseUpdate Then
            CaseUpdate = False
            ParamLineContFix LineContRange, TopOfRoutine, arrLine, L_CodeLine
            Select Case FixData(RoutineCaseFix).FixLevel
             Case CommentOnly
              arrLine(TopOfRoutine) = Marker(arrLine(TopOfRoutine), SUGGESTION_MSG & "Routine's case does not match that of the " & IIf(bImplementsMsg, "Implements name and should be updated.", "control and should be updated."), MAfter, UpDated)
              AddNfix RoutineCaseFix
             Case FixAndComment
              L_CodeLine = arrLine(TopOfRoutine)
              L_CodeLine = Replace$(L_CodeLine, Rname, CaseFix, 1, 1)
              arrLine(TopOfRoutine) = Marker(L_CodeLine, WARNING_MSG & "Routine name case adjusted to match " & IIf(bImplementsMsg, "Implements name", "VB standard formatting"), MAfter, UpDated)
              AddNfix RoutineCaseFix
              'Case JustFix
            End Select
          End If
        End If
        UpdateMember arrMembers(I), arrLine, UpDated, MUpdated
      Next I
    End If
    ReWriteMembers cMod, arrMembers, MUpdated
  End If

End Sub

Private Sub Chr2ConstantDo(strWork As String)

  'converts specific Chr$ to named variables for better readability
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Only has to check for the $ version because variant is fixed by UpDateStringFunctions

  UpdateStringArray strWork, arrOldChr, arrNewConst

End Sub

Private Sub Chr2ConstantUpDateFast(cMod As CodeModule)

  'v3.0.4 new faster version
  
  Dim StartLine  As Long
  Dim I          As Long
  Dim StrNewCode As String

  With cMod
    StartLine = .CountOfDeclarationLines
    For I = LBound(arrOldChr) To UBound(arrOldChr)
      StartLine = 0
      Do While .Find(arrOldChr(I), StartLine, 1, -1, -1)
        StrNewCode = .Lines(StartLine, 1)
        If InCode(StrNewCode, InStr(StrNewCode, arrOldChr(I))) Then
          Chr2ConstantDo StrNewCode
          If StrNewCode <> .Lines(StartLine, 1) Then
            InsertNewCodeComment cMod, StartLine, 1, StrNewCode, vbNullString, Xcheck(XPrevCom)
          End If
        End If
Skip:
        If StartLine < .CountOfDeclarationLines Then
          Exit Do
        End If
        StartLine = StartLine + 1
        If StartLine > .CountOfLines Then
          Exit Do
        End If
      Loop
    Next I
  End With

End Sub

Private Function CompactIndexArr(Arr As Variant) As Variant

  Dim K As Long
  Dim I As Long

  For I = LBound(Arr) To UBound(Arr)
    K = I
    If Len(Arr(I)) Then
      If CountSubStringImbalance(Arr(I), LBracket, RBracket) Then
        Do While CountSubStringImbalance(Arr(I), LBracket, RBracket)
          K = K + 1
          'ver 2.0.1 stops words being forced together
          If K < UBound(Arr) Then
            Arr(I) = Arr(I) & IIf(IsAlphaIntl(Right$(Arr(I), 1)) Or IsAlphaIntl(Left$(Arr(K), 1)), SngSpace, vbNullString) & Arr(K)
            Arr(K) = vbNullString
           Else
            Exit Do
          End If
        Loop
      End If
    End If
  Next I
  CompactIndexArr = CleanArray(Arr)

End Function

Private Function CreateLetProperty(ByVal StrTyp As String, _
                                   ByVal CompName As String) As Boolean

  If StrTyp = "Object" Then
    CreateLetProperty = False
   ElseIf InQSortArray(StandardTypes, StrTyp) Then
    CreateLetProperty = True
   ElseIf IsDeclaration(StrTyp, "Public", "Type") Then
    CreateLetProperty = True
   ElseIf IsDeclaration(StrTyp, "Public", "Enum") Then
    CreateLetProperty = True
   ElseIf IsDeclaration(StrTyp, "Private", "Type", CompName) Then
    CreateLetProperty = True
   ElseIf IsDeclaration(StrTyp, "Private", "Enum", CompName) Then
    CreateLetProperty = True
   ElseIf IsRefLibKnownVBConstantType(StrTyp) Then
    CreateLetProperty = True
   ElseIf ReferenceLibraryConstant(StrTyp) Then
    'v2.8.6 Thanks Ian K for spotting the need for this
    CreateLetProperty = True
  End If

End Function

Private Sub CreatePropertyFromPublicVariableInClass(cMod As CodeModule, _
                                                    Cmp As VBComponent)

  Dim ModuleNumber As Long
  Dim MUpdated     As Boolean
  Dim strNewProc   As String
  Dim arrDec       As Variant
  Dim MaxFactor    As Long
  Dim strName      As String
  Dim strNewName   As String
  Dim strType      As String
  Dim I            As Long
  Dim lngFixMode   As Long
  Dim NewMember    As Boolean

  'new ver 1.1.15
  'this routine converts Public Variables in Classes into Properties for better coding
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If ModDesc(ModuleNumber).MDisControlHolder Or ModDesc(ModuleNumber).MDType = "Class" Then
    'v2.7.7 Public variables in Implements classes should not be converted Thanks Jacobo Hernández Valdelamar
    If Not InQSortArray(ImplementsArray, cMod.Parent.Name) Then
      If dofix(ModuleNumber, PublicVar2Property) Then
        lngFixMode = FixData(PublicVar2Property).FixLevel
        'v 2.5.1 only safe to do this fix if the Public To Private fix has worked first.
        'Thanks Mike Wardle
        ' so reduce fix level to comment only
        If FixData(Public2Private).FixLevel <= CommentOnly Then
          lngFixMode = CommentOnly
        End If
        arrDec = GetDeclarationArray(cMod)
        MaxFactor = UBound(arrDec)
        If MaxFactor > 0 Then
          For I = 1 To MaxFactor
            If Left$(arrDec(I), 7) = "Public " Then
              If InstrArrayWholeWord(arrDec(I), "Event", "WithEvents") = 0 Then
                If IsDeclaration(WordInString(arrDec(I), 2), "Public") Then
                  If Not ProtectConnectVBIDE(arrDec(I)) Then
                    Select Case lngFixMode
                     Case CommentOnly
                      arrDec(I) = Marker(arrDec(I), SUGGESTION_MSG & "Public Variables in Forms/Class/Controls are poor " & ModuleType(cMod) & " coding, Convert to Property", MAfter, MUpdated)
                      AddNfix PublicVar2Property
                     Case FixAndComment
                      strName = WordInString(arrDec(I), 2)
                      strNewName = "CFm_" & strName
                      strType = GetType(arrDec(I))
                      If LenB(strType) = 0 Then
                        strType = "Variant"
                      End If
                      ReDim Preserve DeclarDesc(UBound(DeclarDesc) + 1) As DeclarDescriptor
                      With DeclarDesc(UBound(DeclarDesc))
                        .DDName = strNewName
                        .DDType = strType
                        .DDIndexing = -1
                      End With 'DeclarDesc(UBound(DeclarDesc))
                      arrQDeclarPresence = QuickSortArray(AppendArray(arrQDeclarPresence, strNewName))
                      replaceAllinComp cMod, strName, strNewName
                      'ver1.1.21
                      'ver 1.1.93 simplified
                      strNewProc = PropertyBuilder(IIf(CreateLetProperty(strType, Cmp.Name), "Let", "Set"), strName, strType, strNewName)
                      arrDec(I) = Replace$(arrDec(I), "Public " & strName, "Private " & strNewName)
                      arrDec(I) = Marker(arrDec(I), WARNING_MSG & "Public Variables are poor " & ModuleType(cMod) & " coding." & WARNING_MSG & "Variable has been converted to Private and a Property added to use it", MAfter, MUpdated)
                      'v2.5.4 fix for a problem if last member of Declaration needs to be converted
                      If Not NewMember Then
                        NewMember = True
                        ReDim Preserve arrDec(MaxFactor + 1)
                      End If
                      arrDec(MaxFactor + 1) = arrDec(MaxFactor + 1) & vbNewLine & strNewProc
                      MUpdated = True
                      AddNfix PublicVar2Property
                    End Select
                  End If
                End If
              End If
            End If
          Next I
          If MUpdated Then
            ReWriter cMod, arrDec, RWDeclaration
          End If
        End If
      End If
    End If
  End If

End Sub

Public Sub DoIFBracketEngine(cMod As CodeModule)

  Dim ArrWords  As Variant
  Dim K         As Long
  Dim strTest   As String
  Dim TPos      As Long
  Dim EPos      As Long
  Dim Hit       As Boolean
  Dim J         As Long
  Dim StartLine As Long
  Dim arrTest   As Variant

  arrTest = Array("=", ">", "<")
  'v2.2.3 new fix
  'v2.5.0 restructure of old fix
  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find("If (", StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If Left$(strCodeOnly(.Lines(StartLine, 1)), 4) = "If (" Then
        strTest = .Lines(StartLine, 1)
        ArrWords = Split(strTest)
        For K = LBound(ArrWords) To UBound(ArrWords)
          If InCode(strTest, InStr(strTest, ArrWords(K))) Then
            If Left$(ArrWords(K), 1) = "(" Then
              If Right$(ArrWords(K), 1) = ")" Then
                If CountSubString(ArrWords(K), "(") = 1 Then
                  If CountSubString(ArrWords(K), ")") = 1 Then
                    'v2.3.5 Thanks Stephen S Coleman
                    '       stop 'If (UBound(x) - Index) \ Y > 0 Then' becoming
                    '            ' If  UBound(x - Index) \ Y > 0 Then'
                    ArrWords(K) = Mid$(ArrWords(K), 2)
                    ArrWords(K) = Left$(ArrWords(K), Len(ArrWords(K)) - 1)
                  End If
                End If
              End If
            End If
          End If
        Next K
        If strTest <> Join(ArrWords) Then
          strTest = Join(ArrWords)
          Hit = True
          GoTo donesimple
        End If
        TPos = InStr(strTest, "If (") + 4
        EPos = InStr(TPos, strTest, ") Then")
        Hit = TPos And (EPos > TPos)
        If Hit Then
          For J = TPos To EPos
            If Mid$(strTest, J, 1) = " " Then
              If Not IsInArray(Mid$(strTest, J + 1, 1), arrTest) Then
                If Not IsInArray(Mid$(strTest, J - 1, 1), arrTest) Then
                  If (J > 4 And Mid$(strTest, J - 3, 4) <> "Not ") Or J < 4 Then
                    Hit = False
                    Exit For
                  End If
                End If
              End If
            End If
          Next J
          If Hit Then
            Mid$(strTest, TPos - 1, 1) = " "
            Mid$(strTest, EPos, 1) = " "
          End If
donesimple:
          If Hit Then
            Select Case FixData(UnNeededCall).FixLevel
             Case CommentOnly
              .ReplaceLine StartLine, .Lines(StartLine, 1) & vbNewLine & _
               WARNING_MSG & "Unneeded brackets could be removed."
             Case FixAndComment
              .ReplaceLine StartLine, Marker(strTest, WARNING_MSG & "Unneeded brackets removed" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & .Lines(StartLine, 1), vbNullString), MAfter)
            End Select
          End If
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With
  On Error GoTo 0

End Sub

Private Sub DoIfThenStructureExpanderFast(cMod As CodeModule)

  Dim PSLine   As Long
  Dim PEndLine As Long
  Dim ArrProc  As Variant
  Dim UpDated  As Boolean
  Dim Sline    As Long

  With cMod
    If dofix(ModDescMember(cMod.Parent.Name), NIfThenExpand) Then
      Do While .Find("*If * Then *", Sline, 1, -1, -1, , , True)
        If Sline > .CountOfDeclarationLines Then
          If InCode(.Lines(Sline, 1), InStr(.Lines(Sline, 1), "If ")) Then
            ArrProc = ReadProcedureCodeArray2(cMod, Sline, PSLine, PEndLine)
            MemberMessage GetProcName(cMod, Sline), Sline, .CountOfLines
            IfThenArrayExpander ArrProc, UpDated
            If UpDated Then
              ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
            End If
          End If
        End If
        Sline = Sline + IIf(PEndLine - PSLine > 1, PEndLine - PSLine, 1)
        PEndLine = 0
        PSLine = 0
        If Sline > cMod.CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

End Sub

Private Sub DoStringFunctionsCorrect(Cline As String)

  Dim J          As Long
  Dim tmpstring2 As String
  Dim Tmpstring1 As String

  'Ulli's code updated to use built-in array instead of Listbox
  'You don't get the option with me!!
  'but it is safe for code which make literal string references to these (Like the array that this uses)
  For J = LBound(ArrQStrVarFunc) To UBound(ArrQStrVarFunc)
    If Get_As_Pos(Cline) = 0 Then
      Tmpstring1 = ArrQStrVarFunc(J)
      If InstrAtPosition(Cline, Tmpstring1, ipAny, False) Then
        If isEvent("Error") And InStr(Tmpstring1, "Error") Then
          GoTo noFix
        End If
        tmpstring2 = LBracket & Tmpstring1
        Tmpstring1 = SngSpace & Tmpstring1
        Cline = Replace$(Cline, Tmpstring1 & LBracket, Tmpstring1 & "$(")
        Cline = Replace$(Cline, tmpstring2 & LBracket, tmpstring2 & "$(")
        If MultiLeft(Cline, True, Trim$(Tmpstring1) & LBracket) Then
          Cline = Replace$(Cline, Trim$(Tmpstring1) & LBracket, Trim$(Tmpstring1) & "$(", , 1)
        End If
noFix:
      End If
    End If
  Next J

End Sub

Private Function EmptyStringComparisonFix(varTest As Variant, _
                                          DoneIt As Boolean, _
                                          ImpossCode As Boolean) As String

  Dim arrTmp As Variant
  Dim I      As Long
  Dim Cycles As Long

  'ver 1.0.95 Update
  'rudz suggested the use of LenB instead of Len for this
  Cycles = CountSubString(varTest, EmptyString)
  If Cycles > 0 Then
    ImpossCode = ImpossibleEmptyStringTest(varTest)
    Do
      varTest = ConcealParameterSpaces(CStr(varTest))
      arrTmp = CompactIndexArr(Split(varTest))
      For I = LBound(arrTmp) To UBound(arrTmp)
        If arrTmp(I) = EmptyString Then
          Select Case arrTmp(I - 1)
           Case "<>", ">"
            arrTmp(I - 2) = "LenB(" & arrTmp(I - 2) & RBracket
            arrTmp(I - 1) = vbNullString
            arrTmp(I) = vbNullString
            If I + 1 < UBound(arrTmp) Then
              If I - 2 > LBound(arrTmp) Then
                If arrTmp(I + 1) = "And" Or arrTmp(I - 3) = "And" Then
                  arrTmp(I - 1) = ">"
                  arrTmp(I) = "0"
                End If
              End If
            End If
            DoneIt = True
           Case "="
            arrTmp(I - 2) = "LenB(" & arrTmp(I - 2) & RBracket
            arrTmp(I) = "0"
            DoneIt = True
            'Case Else
          End Select
          varTest = Join(arrTmp)
          Exit For
        End If
      Next I
      If InStr(varTest, EmptyString & " = ") Then
        If arrTmp(1) = EmptyString Then
          If arrTmp(2) = "=" Then
            arrTmp(1) = "LenB("
            arrTmp(2) = vbNullString
            arrTmp(3) = arrTmp(3) & ")"
            varTest = Join(arrTmp)
            DoneIt = True
          End If
        End If
      End If
      Cycles = Cycles - 1
    Loop While Cycles > 0
    varTest = Replace$(varTest, Chr$(160), SngSpace)
    If InStr(varTest, "&") Then
      For I = 1 To Len(varTest)
        If Mid$(varTest, I, 1) = "&" Then
          If InCode(varTest, I) Then
            varTest = Left$(varTest, I - 1) & " & " & Mid$(varTest, I + 1)
            'Fix to stop the same '&" being hit twice
            I = I + 2
          End If
        End If
      Next I
    End If
  End If
  If Not DoneIt Then
    ZeroStringForcedTrue varTest, DoneIt, ImpossCode
  End If

End Function

Private Function EmptyStringType(ByVal strCode As String) As Long

  'added ver 1.0.95
  'Thanks to Rudz suggesting the AssignVbNullString type
  'None               =0
  'Length_test        =1
  'Assign_vbNullString=2
  'InString double quotes=3
  'In parameter = 4

  If InStr(strCode, EmptyString) Then
    If InStr(strCode, EmptyString & EmptyString) = 0 Then
      'don't apply to strings using the """" to generate Chr(34) in literal strings
      If InStr(strCode, ".InsertLines") Then
        EmptyStringType = 0
       ElseIf InStr(strCode, EqualInCode) = 0 And InStr(strCode, " <> ") = 0 Then
        ' new stops 'SubName ""' (sub sent empty string)
        EmptyStringType = 0
       ElseIf InStr(strCode, "Optional ") Then
        EmptyStringType = 2
       ElseIf InStringDblQuote(strCode) Then
        EmptyStringType = 3
       ElseIf InStr(strCode, "= " & EmptyString) Or InStr(strCode, "<> " & EmptyString) Then
        'ver 2.0.3 updated to get the <> version as well
        If MultiLeft(strCode, True, "If", "ElseIf", "Do", "Loop", "While") Then
          EmptyStringType = 1
          'v2.5.7 thanks Roy Blanch
          If MultiLeft(strCode, True, "If", "ElseIf") Then
            If InStr(strCode, EmptyString) > InStr(strCode, "Then") Then
              EmptyStringType = 0
            End If
          End If
         Else
          If WordInString(strCode, 2) = "=" Or WordInString(strCode, 2) = "<>" Then
            'ver1.1.77 protects assignment code where left of 1st word would pass the MultiLeft test
            EmptyStringType = IIf(WordInString(strCode, 3) = EmptyString, 2, 0)
          End If
        End If
       ElseIf InStr(strCode, DQuote & DQuote & SQuote & DQuote) Or InStr(strCode, DQuote & SQuote & DQuote & DQuote) Then
        'v 2.6.1 this protects ""'" a string used to detect single and double quotes using Instr. Thanks Mike Ulik
        EmptyStringType = 0
       ElseIf InstrControlWhichCantAcceptVBNull(strCode) Then
        EmptyStringType = 0
       ElseIf InStr(strCode, CommaSpace & EmptyString) Or InStr(strCode, EmptyString & CommaSpace) Then
        EmptyStringType = 4
       ElseIf MultiLeft(strCode, True, "If", "ElseIf", "Do", "Loop", "While") Then
        EmptyStringType = 1
       ElseIf InStr(strCode, "> " & EmptyString) Or InStr(strCode, "<> " & EmptyString) Then
        EmptyStringType = 2
       ElseIf strCode = "Case " & EmptyString Then
        'added this for unlikely Case but just happened to be in first program I tested it on.
        EmptyStringType = 2
      End If
    End If
  End If

End Function

Private Function ErrErrorFix(ByVal strWork As String) As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'This routine is updates the Older Err and Error$ to the newer
  'Err.Number and Err.Description
  'v2.1.0 Thanks to Dipankar Basu who brought this to my attention

  If InStrCode(strWork, "On Error GoTo Err") Then
    If IsRealWord(strWork, "Err") Then ' ignore 'On Error GoTo ErrHandler' which matches aove
      'v2.1.1 InstrCode stops detecting this line
      ErrErrorFix = strWork & vbNewLine & _
       WARNING_MSG & "Legal but hard to read use of a VB command word as GoTo label."
      GoTo SafeExit
    End If
  End If
  'v2.8.6 ignore Error if it is an event named Error
  ' thanks
  If InStrCode(strWork, "RaiseEvent Error") Or InStrCode(strWork, "Raise Error") Then
    If IsRealWord(strWork, "Error") Then ' ignore 'On Error GoTo ErrHandler' which matches aove
      'v2.1.1 InstrCode stops detecting this line
      GoTo SafeExit
    End If
  End If
  UpdateStringArray strWork, Array("Err ", "Error$"), Array("Err.Number ", "Err.Description")
  '2.0.1 Thanks to Ken Marsh who found a bug that lead me to tighten the code up here
  If LastWord(strWork) = "Err" Then
    If Not Left$(strWork, 8) = "With Err" Then
      strWork = StrReverse(Replace$(StrReverse(strWork), "rrE", "rebmuN.rrE", , 1))
    End If
  End If
  ErrErrorFix = strWork
SafeExit:

End Function

Private Sub ErrErrorUpDate(cMod As CodeModule)

  Dim UpDated     As Boolean
  Dim MUpdated    As Boolean
  Dim ArrMember   As Variant
  Dim ArrRoutine  As Variant
  Dim Member      As Long
  Dim RLine       As Long
  Dim MemberCount As Long
  Dim strTmp      As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  If dofix(ModDescMember(cMod.Parent.Name), UpdateErrError) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        If UBound(ArrRoutine) > -1 Then
          For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
            strTmp = ArrRoutine(RLine)
            If Not JustACommentOrBlank(strTmp) Then
              '2.0.1 Thanks to Ken Marsh who found a bug that lead me to tighten the code up here
              If InStrWholeWord(strTmp, "Err") Or InStrWholeWord(strTmp, "Error$") Then
                If Not InstrAtPosition(strTmp, "Dim", IpLeft) Then
                  If isEvent("Error") And InStr(strTmp, "Error") Then
                    GoTo noFix
                  End If
                  strTmp = ErrErrorFix(strTmp)
                  If ArrRoutine(RLine) <> strTmp Then
                    ArrRoutine(RLine) = strTmp
                    UpDated = True
                  End If
noFix:
                 Else
                  ArrRoutine(RLine) = ArrRoutine(RLine) & RGSignature & "Poor coding to use VB Command as variant name"
                  UpDated = True
                End If
              End If
            End If
          Next RLine
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Sub ExitProc2ExitFor(arrR As Variant, _
                             ScanFrom As Long, _
                             success As Boolean)

  Dim I            As Long
  Dim NextCount    As Long
  Dim Exitline     As Long
  Dim strRemTarget As String

  strRemTarget = strCodeOnly(arrR(ScanFrom))
  If GetExitPointType(arrR, ScanFrom + 1, Exitline) Then
    For I = ScanFrom + 1 To Exitline - 1
      If Not JustACommentOrBlank(arrR(I)) Then
        If Not Left$(arrR(I), 6) = "End If" Then ' accept this
          'v2.9.0 improved test
          If Not IsOnErrorCode(arrR(I)) Then      ' accept this
            If LeftWord(arrR(I)) <> "Next" Then         'test this
              NextCount = 0                               ' some other code; be safe and fail test
              Exit For
             Else
              NextCount = NextCount + 1
              If NextCount > 1 Then ' too many for the fix fail test
                Exit For
              End If
            End If
          End If
        End If
      End If
    Next I
    If NextCount = 1 Then
      arrR(ScanFrom) = "Exit For" & WARNING_MSG & "'Unnecessary Exit Procedure fix' replaced '" & strRemTarget
      success = True
    End If
  End If

End Sub

Private Function ExitPRoc2ExitStruct(arrR As Variant, _
                                     ByVal RLine As Long) As Boolean

  Dim Counter As Long
  Dim I       As Long

  'v2.7.4
  For I = LBound(arrR) To RLine
    If Left$(arrR(I), 3) = "If " Then
      Counter = Counter + 1
    End If
    If Left$(arrR(I), 4) = "For " Then
      Counter = Counter + 1
    End If
    If Left$(arrR(I), 3) = "Do " Then
      Counter = Counter + 1
    End If
    If Left$(arrR(I), 6) = "End If" Then
      Counter = Counter - 1
    End If
    If LeftWord(arrR(I)) = "Loop" Then
      Counter = Counter - 1
    End If
    If LeftWord(arrR(I)) = "Next" Then
      Counter = Counter - 1
    End If
    If Counter > 2 Then
      Exit For
    End If
  Next I
  ExitPRoc2ExitStruct = Counter = 2

End Function

Private Sub ExitProc2ExtendedIfThen(arrR As Variant, _
                                    ScanFrom As Long, _
                                    ByVal ExitLineType As Long, _
                                    success As Boolean)

  Dim I            As Long
  Dim Exitline     As Long
  Dim Ifcount      As Long
  Dim strRemTarget As String

  strRemTarget = strCodeOnly(arrR(ScanFrom))
  If GetExitPointType(arrR, ScanFrom + 1, Exitline) Then
    For I = ScanFrom - 1 To LBound(arrR) Step -1
      If LeftWord(arrR(I)) = "If" Then
        If ExitLineType = eExitAs1stCode_ID1a Then
          'only negate the simple case
          If WordInString(arrR(I), 2) = "Not" Then
            arrR(I) = Replace$(arrR(I), "If Not ", "If ", , 1) & PREVIOUSCODE_MSG & arrR(I)
           Else
            arrR(I) = Replace$(arrR(I), "If ", "If Not ", , 1) & PREVIOUSCODE_MSG & arrR(I)
          End If
          arrR(I) = IfNotThenSimplify(CStr(arrR(I)))
          Exit For
        End If
      End If
    Next I
    'End If
    For I = ScanFrom To UBound(arrR)
      If LeftWord(arrR(I)) = "If" Then
        Ifcount = Ifcount + 1
      End If
      If strCodeOnly(arrR(I)) = "End If" Then
        Ifcount = Ifcount - 1
        If Ifcount <= 0 Then
          If ExitLineType = eExitAs1stCodeComplex_ID2a Then
            arrR(ScanFrom) = "Else" & WARNING_MSG & "'Unnecessary Exit Procedure fix': " & _
                        "'If structure' extended. Old code '" & strRemTarget
           Else
            arrR(ScanFrom) = WARNING_MSG & "'Unnecessary Exit Procedure fix': " & _
                              "'If structure' logic reversed and extended. Old code '" & strRemTarget
          End If
          arrR(I) = WARNING_MSG & "'Unnecessary Exit Procedure fix'  removed '" & arrR(I)
          Exit For
        End If
      End If
    Next I
    'add the necessary End If just above the Exit line
    arrR(Exitline) = "End If" & WARNING_MSG & "'Unnecessary Exit Procedure fix' " & _
                         "added this." & vbNewLine & _
                         arrR(Exitline)
    success = True
  End If

End Sub

Private Function ExplicitExitTopOfCode(arrR As Variant, _
                                       ByVal RLine As Long, _
                                       strModName As String, _
                                       Mode As Long) As Boolean

  Dim I       As Long
  Dim Sline   As Long
  Dim arrTest As Variant

  arrTest = Array("Else", "ElseIf")
  'v2.7.3
  Do Until Left$(arrR(RLine), 3) = "If "
    RLine = RLine - 1
    If Not JustACommentOrBlank(arrR(RLine)) Then
      Sline = Sline + 1
      'v2.9.6 this stops incorrect fix of exit proc after an Else. Thanks Denis Sugrue
      If IsInArray(LeftWord(arrR(RLine)), arrTest) Then
        ExplicitExitTopOfCode = False
        Exit Function
      End If
    End If
  Loop
  If RLine Then
    Mode = Sline
    ExplicitExitTopOfCode = True
    For I = 1 To RLine - 1
      If Not JustACommentOrBlank(arrR(I)) Then
        If Not IsDimLine(arrR(I), True) Then
          If Not IsGotoLabel(arrR(I), strModName) Then
            If Not IsOnErrorCode(arrR(I)) Then
              If Not SetVarForIfTest(arrR(I), arrR(RLine)) Then
                If Not isProcHead(arrR(I)) Then
                  If Left$(arrR(RLine), 3) = "If " Then  'only if a single 'If '
                    ExplicitExitTopOfCode = False
                    Exit For
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Private Function ExplicitExitType(arrR As Variant, _
                                  RLine As Long, _
                                  cMod As CodeModule) As eExits

  Dim StrDepDesc    As String
  Dim I             As Long
  Dim lngSearchBack As Long
  Dim strTest       As String
  Dim TCodeMOde     As Long

  'eExitUNknown '=0
  'eExitIf_Then_Exit_EndIf = 1
  'eExitIf_Then_Exit_EndIf_Code = 2
  'eExitIf_Then_Exit_Else_Code_EndIf = 3
  'eExitErrorTrapShield = 4
  'eExitDeepStructure = 5
  'eExitAs1stCode '= 6
  'SDeep = GetExitStructureDepthLine(arrR, RLine, StrDepDesc)
  GetExitStructureDepthLine arrR, RLine, StrDepDesc
  If Not MultiStructureExit(arrR, RLine) Then
    If MultiLeft(NextCodeLine(arrR, RLine), True, "End If", "Else") Then
      For I = RLine - 1 To LBound(arrR) Step -1
        If Not JustACommentOrBlank(arrR(I)) Then
          lngSearchBack = lngSearchBack + 1
          strTest = strCodeOnly(arrR(I))
          If Left$(strTest, 7) = "ElseIf " Then
            Exit For
          End If
          If Left$(strTest, 3) = "If " Then
            If Right$(strTest, 5) = " Then" Then
              Exit For
            End If
          End If
        End If
      Next I
      If Left$(NextCodeLine(arrR, RLine), 6) = "End If" Then
        If ExplicitExitTopOfCode(arrR, RLine, cMod.Parent.Name, TCodeMOde) Then
          ExplicitExitType = IIf(TCodeMOde > 1, eExitAs1stCodeComplex_ID2a, eExitAs1stCode_ID1a)
         Else
          If MultiStructureExit(arrR, RLine) Then
            ExplicitExitType = eExitDeepStructure_ID5 '5
           Else
            If ExitPRoc2ExitStruct(arrR, RLine) Then
              ExplicitExitType = eExitProc2ExitLoop_ID8a
             Else
              ExplicitExitType = IIf(lngSearchBack = 1, eExitIf_Then_Exit_EndIf_ID1, eExitIf_Then_Exit_EndIf_Code_ID2) '1,2
              ' test for deep embbeded
              'End If
            End If
          End If
        End If
       Else
        ExplicitExitType = IIf(lngSearchBack = 1, eExitIf_Then_Exit_EndIf_ID1, eExitIf_Then_Exit_Else_Code_EndIf_ID3) '1,3
      End If
    End If
    If ExplicitExitType = 0 Then
      For I = RLine + 1 To UBound(arrR)
        If Not JustACommentOrBlank(arrR(I)) Then
          If IsGotoLabel(arrR(I), cMod.Parent.Name) Then
            ExplicitExitType = eExitErrorTrapShield_ID4 ' 4
            Exit For
          End If
        End If
      Next I
     Else
      If MultiStructureExit(arrR, RLine) Then
        ExplicitExitType = eExitDeepStructure_ID5 '5
        'Else
        'explicitExitType = eExitUNknown = 0
      End If
    End If
   Else
    If InStr(StrDepDesc, "With") Then
      ExplicitExitType = eExitFromWith_ID9
     ElseIf InStr(StrDepDesc, "If,If") Then
      ExplicitExitType = eExitGEneric_ID11
     ElseIf CountSubString(StrDepDesc, "For") > 1 Then
      ExplicitExitType = eExitFromFors_ID10
     Else
      ExplicitExitType = eExitGEneric_ID11
    End If
  End If

End Function

Private Function GetExitPointType(arrR As Variant, _
                                  ByVal TestFrom As Long, _
                                  Exitline As Long) As Long

  Dim I         As Long
  Dim TestDepth As Long

  'v2.8.5 support for SecondryExitsHard
  'return =0 No exit? something deeply wrong
  '       =1  End Proc
  '       =2  ErrorHandler protection Exit Proc
  '
  TestDepth = GetStructureDepthLine(arrR, TestFrom)
  For I = TestFrom To UBound(arrR)
    If MultiLeft(arrR(I), True, "Exit Sub", "Exit Function", "Exit Property") Then
      If GetStructureDepthLine(arrR, I) < TestDepth Then ' ignore intervening exits
        GetExitPointType = 2
        Exitline = I
        Exit For
      End If
    End If
    If MultiLeft(arrR(I), True, "End Sub", "End Function", "End Property") Then
      GetExitPointType = 1
      Exitline = I
      Exit For
    End If
  Next I

End Function

Public Function GetExitStructureDepthLine(arrR As Variant, _
                                          ByVal LineNo As Long, _
                                          StrDepth As String) As Long

  Dim I       As Long
  Dim lngDeep As Long

  For I = 0 To LineNo
    lngDeep = StructureDeep2(lngDeep, arrR(I), StrDepth)
  Next I
  StrDepth = Replace$(StrDepth, ",,", ",")
  If Left$(StrDepth, 1) = "," Then
    StrDepth = Mid$(StrDepth, 2)
  End If
  If Right$(StrDepth, 1) = "," Then
    StrDepth = Left$(StrDepth, Len(StrDepth) - 1)
  End If
  GetExitStructureDepthLine = lngDeep

End Function

Public Function GetGoToTargetArray(ArrRoutine As Variant) As Variant

  Dim I       As Long
  Dim strGoTo As String

  'v2.4.4 Thanks Paul Caton
  'this helps fix a problem with goto target labels on same line as other code
  For I = LBound(ArrRoutine) To UBound(ArrRoutine)
    If InStr(ArrRoutine(I), "GoTo ") Then
      'v2.7.7 Thanks Evan Toder this was where the unneeded colon was being picked up
      strGoTo = AccumulatorString(strGoTo, Replace$(WordAfter(ArrRoutine(I), "GoTo"), ":", vbNullString))
    End If
  Next I
  GetGoToTargetArray = QuickSortArray(Split(strGoTo, ","))

End Function

Public Function GetInsertionPoint(cMod As CodeModule, _
                                  Lnum As Long) As String

  Dim StrLine As String

  StrLine = cMod.Lines(Lnum, 1)
  Do While Right$(StrLine, 2) = " _"
    Lnum = Lnum + 1
    StrLine = cMod.Lines(Lnum, 1)
  Loop
  GetInsertionPoint = Lnum

End Function

Private Function GetNextCodeLine(ByVal StartLine As Long, _
                                 cMod As CodeModule) As Long

  Dim I As Long

  For I = StartLine + 1 To cMod.CountOfLines
    If Not JustACommentOrBlank(cMod.Lines(I, 1)) Then
      GetNextCodeLine = I
      Exit For
    End If
  Next I

End Function

Private Function GetPreviousCodeLine(ByVal StartLine As Long, _
                                     cMod As CodeModule) As Long

  Dim I As Long

  For I = StartLine - 1 To 1 Step -1
    If Not JustACommentOrBlank(cMod.Lines(I, 1)) Then
      GetPreviousCodeLine = I
      Exit For
    End If
  Next I

End Function

Private Sub GotoLabelSeparator(cMod As CodeModule)

  Dim Member          As Long
  Dim RLine           As Long
  Dim ArrMember       As Variant
  Dim ArrRoutine      As Variant
  Dim L_CodeLine      As String
  Dim MemberCount     As Long
  Dim UpDated         As Boolean
  Dim MUpdated        As Boolean
  Dim arrGoTo         As Variant
  Dim arrGoToCount    As Long
  Dim strPoorNameType As String

  'v2.6.8 correct GoToLabel separator
  'Break up compound lines
  If dofix(ModDescMember(cMod.Parent.Name), SeperateCompounds) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        'v3.0.7 speed up test
        If InStr(ArrMember(Member), ":") Then
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          arrGoTo = GetGoToTargetArray(ArrRoutine)
          If UBound(arrGoTo) > -1 Then
            For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
              L_CodeLine = ArrRoutine(RLine)
              If Not JustACommentOrBlank(L_CodeLine) Then
                For arrGoToCount = 0 To UBound(arrGoTo)
                  If LenB(arrGoTo(arrGoToCount)) Then
                    If SmartLeft(L_CodeLine, arrGoTo(arrGoToCount) & ":") Then
                      ArrRoutine(RLine) = Replace$(L_CodeLine, ":", ":" & vbNewLine, , 1)
                      AddNfix SeperateCompounds
                      UpDated = True
                    End If
                  End If
                Next arrGoToCount
                For arrGoToCount = 0 To UBound(arrGoTo)
                  If LenB(arrGoTo(arrGoToCount)) Then
                    If PoorNameGoto(arrGoTo(arrGoToCount), cMod.Parent.Name, strPoorNameType) Then
                      ProcArrayErrTrapRenamer ArrRoutine, CStr(arrGoTo(arrGoToCount)), arrGoTo(arrGoToCount) & GetProcNameStr(ArrMember(Member)), vbNewLine & _
                       WARNING_MSG & "Poorly named GotoLabel Renamed", UpDated
                      arrGoTo(arrGoToCount) = vbNullString
                      Exit For ' once changed exit for
                    End If
                  End If
                Next arrGoToCount
              End If
            Next RLine
          End If
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Function HasUDTParameter(ByVal strTest As String) As Boolean

  Dim I As Long

  If bProcDescExists Then
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      If PRocDesc(I).PrDName = strTest Then
        HasUDTParameter = PRocDesc(I).PrDUDTParam
        Exit For
      End If
    Next I
  End If

End Function

Private Sub IfAndThenShortCircuit_Apply2(cMod As CodeModule, _
                                         ByVal LineNo As Long, _
                                         L_CodeLine As String, _
                                         Update As Boolean)

  Dim K      As Long
  Dim EndPos As Long

  If TypeOfIf2(cMod, LineNo, EndPos) = Simple Then
    IfAndThenShortCircuitSplitter L_CodeLine
    L_CodeLine = Marker(L_CodeLine, WARNING_MSG & "Short Curcuit: 'If <condition1> And <condition2> Then' expanded." & IIf(Xcheck(XVerbose), RGSignature & "'If <condition1> Then" & RGSignature & "If <condition2> Then '" & RGSignature & "Make <condition1> most likely to fail.", vbNullString), MAfter, Update)
    cMod.ReplaceLine LineNo, L_CodeLine
    EndPos = EndPos + CountSubString(L_CodeLine, vbNewLine)
    For K = 1 To CountSubString(L_CodeLine, " Then " & vbNewLine & "If ")
      cMod.InsertLines EndPos, Marker("End If", RGSignature & "Short Circuit inserted this line", MEoL, Update)
    Next K
  End If

End Sub

Public Sub IfAndThenShortCircuitEngine(cMod As CodeModule)

  Dim L_CodeLine As String
  Dim UpDated    As Boolean
  Dim StartLine  As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'IfAndThenShortCircuit
  'This routine seeks potential points for performing Logical Short Ciruiting
  'This means that when testing multiple conditions you test one at a time rather than
  'all at once in situations where any failure negates following action.
  'for example the line:
  '    If len(b$)>0 and Instr(a$,"Fred") Then
  'will fail if b$="" or "Fred" is not in a$ but as written above both tests must be performed
  'before it fails(or passes)
  'however the following is faster
  '    If len(b$) Then
  '    If Instr(a$,"Fred") Then
  'may fail after only performing one test.
  'When using this style you should place the most likely to fail test first,
  'unless the test is extremely complex, in which case place the simplest/fastest test first.
  '
  'NOTE If the line read '*If len(A$)>0 and Instr(A$,"Fred") Then' then 'If Instr(A$,"Fred") Then' would be just as good.
  '
  'This routine is not efficent in that it makes no attempt to work out which test is most efficent
  'but relies on the tendancy of people to code the simplest tests first then add more complex ones later.
  '
  'v2.4.3 new version of fix
  '
  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find("If * And * Then", StartLine, 1, -1, -1, , , True)
      MemberMessage "", StartLine, .CountOfLines
      L_CodeLine = .Lines(StartLine, 1)
      If IfAndThenShortCircuitLine(L_CodeLine) Then
        If IfAndThenShortCircuitSafeToApply(L_CodeLine) Then
          Select Case FixData(DetectShortCircuit).FixLevel
           Case CommentOnly
            .ReplaceLine StartLine, Marker(L_CodeLine, SUGGESTION_MSG & "Short Curcuit: Previous line may be able to be optimized by multiple 'If test_Condition Then' lines", MAfter)
           Case FixAndComment
            IfAndThenShortCircuit_Apply2 cMod, StartLine, L_CodeLine, UpDated
            AddNfix DetectShortCircuit
            UpDated = True
          End Select
        End If
      End If
      '              If .Lines(StartLine, 1) <> L_CodeLine Then
      '                .ReplaceLine StartLine, L_CodeLine
      '              End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Sub IfAndThenShortCircuitSplitter(strCode As String)

  Dim PossAnd As Long

  'v2.8.3 improved splitter allows multiple splits but ignores logic And's (in brackets) Thanks Joakim Schramm
  'original simple
  'strCode = Safe_Replace(strCode, " And ", " Then " & vbNewLine & "If ") ', , 1)
  'smart version
  PossAnd = InStr(strCode, " And ") ' get a possible target
  If PossAnd Then
    Do
      If Not EnclosedInBrackets(strCode, PossAnd) Then
        strCode = Safe_Replace(strCode, " And ", " Then " & vbNewLine & "If ", PossAnd, 1)
      End If
      PossAnd = InStr(PossAnd + 1, strCode, " And ") ' get next possible target
    Loop While PossAnd
  End If

End Sub

Private Function IfNotThenReducer(ByVal strCode As String, _
                                  ByVal strOld As String, _
                                  ByVal Strnew As String, _
                                  success As Boolean) As String

  IfNotThenReducer = strCode
  If InStr(strCode, strOld) > InStr(strCode, "If Not(") Then
    If InStr(strCode, strOld) < InStr(strCode, ") Then") Then
      strCode = Replace$(strCode, "If Not(", "If ", , 1)
      strCode = Replace$(strCode, ") Then", " Then", , 1)
      strCode = Replace$(strCode, strOld, Strnew, , 1)
      IfNotThenReducer = strCode
      success = True
    End If
  End If

End Function

Private Function IfNotThenSimplify(ByVal strCode As String) As String

  Dim strTemp   As String
  Dim strOldCom As String
  Dim Hit       As Boolean

  IfNotThenSimplify = strCode
  ExtractCode IfNotThenSimplify, strOldCom
  strTemp = strCodeOnly(IfNotThenSimplify)
  If CountSubString(strTemp, " And ") = 0 Then ' too complex; leave alone
    If CountSubString(strTemp, " Or ") = 0 Then
      If CountSubStringArray(strCode, " = 0) ", " <> ", " < ", " > ", " >= ", " <= ", " = ") Then
        'only hit if only a single operator
        If InStr(strTemp, "If Not(") Then
          If CountSubString(strTemp, " = 0) ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " = 0 ", vbNullString, Hit)
           ElseIf CountSubStringCode(strTemp, " <> ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " <> ", " = ", Hit)
           ElseIf CountSubStringCode(strTemp, " > ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " > ", " <= ", Hit)
           ElseIf CountSubStringCode(strTemp, " < ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " < ", " >= ", Hit)
           ElseIf CountSubStringCode(strTemp, " <= ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " <= ", " > ", Hit)
           ElseIf CountSubStringCode(strTemp, " < ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " >= ", " < ", Hit)
           ElseIf CountSubStringCode(strTemp, " = ") = 1 Then
            strTemp = IfNotThenReducer(strTemp, " = ", " <> ", Hit)
          End If
          If Hit Then
            IfNotThenSimplify = strTemp & vbNewLine & _
             WARNING_MSG & "Logic simplified from:" & vbNewLine & _
             strOldCom
          End If
        End If
      End If
    End If
  End If

End Function

Private Sub IfThen2BooleanEngine(cMod As CodeModule, _
                                 ByVal LngMode As Long)

  Dim strTest1         As String
  Dim strTest2         As String
  Dim StartLine        As Long
  Dim StrNewCode       As String
  Dim StrCom           As String
  Dim lngCodeLineNo(4) As Long
  Dim J                As Long

  'v2.4.1 NEW FIX converts
  'v2.4.3 expanded to cope with If Then struture that equates to  'x = Not X'
  '1 If X Then
  '2   Y = True
  '3 Else
  '4   Y = False
  '5 End If
  'to collapsed form
  ' Y = X
  ''lngMode = 0 line 2 is Y = True
  ''lngMode = 1 line 2 is Y = False
  'This procedure is triggered by detecting line 2 then searching around it for matching structures 1,3,4,5
  'any eol comments are collected and then written after the collapsed line
  '
  'TODO deal with the half smart version(s) of this code
  '1 If X Then
  '2   Y = True
  '3 End If
  Select Case LngMode
   Case 0
    strTest1 = " = True"
    strTest2 = " = False"
   Case 1
    strTest2 = " = True"
    strTest1 = " = False"
  End Select
  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find(strTest1, StartLine, 1, -1, -1)
      If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), strTest1)) Then
        MemberMessage "", StartLine, .CountOfLines
        'get the real code lines
        lngCodeLineNo(0) = GetPreviousCodeLine(StartLine, cMod)         'If Y Then
        lngCodeLineNo(1) = StartLine                                    ' X = T/F
        lngCodeLineNo(2) = GetNextCodeLine(lngCodeLineNo(1), cMod)      'Else
        lngCodeLineNo(3) = GetNextCodeLine(lngCodeLineNo(2), cMod)      ' X = F/T
        lngCodeLineNo(4) = GetNextCodeLine(lngCodeLineNo(3), cMod)      'End If
        'test all lines are code (0 means no line found)
        For J = 0 To 4
          If lngCodeLineNo(J) = 0 Then
            GoTo Skip
          End If
        Next J
        'test lines 2 and 4 conform to target code
        If strCodeOnly(.Lines(lngCodeLineNo(2), 1)) = "Else" Then
          If strCodeOnly(.Lines(lngCodeLineNo(4), 1)) = "End If" Then
            'check that line 3 is inverse of line 1
            If InStr(strCodeOnly(.Lines(lngCodeLineNo(3), 1)), strTest2) Then
              'check line 1 & 3 are setting the same variable
              If LeftWord(.Lines(lngCodeLineNo(3), 1)) = LeftWord(.Lines(lngCodeLineNo(1), 1)) Then
                StrCom = vbNullString
                StrNewCode = strCodeOnly(.Lines(lngCodeLineNo(0), 1))
                'check that if is in the right format
                If Left$(StrNewCode, 3) = "If " Then
                  If Right$(StrNewCode, 5) = " Then" Then
                    ' extract and comments and set them aside for later reinsertionn after the collapse
                    MultilineCommentExtractor cMod, lngCodeLineNo(0), lngCodeLineNo(4), StrCom
                    'convert line 0 to the right side of equation
                    StrNewCode = Mid$(StrNewCode, 4)
                    StrNewCode = Left$(StrNewCode, Len(StrNewCode) - 5)
                    ' deal with silly case of If X Then/X =T /Else/X = F/End If
                    If StrNewCode = LeftWord(.Lines(lngCodeLineNo(1), 1)) Then
                      If strTest1 = " = True" Then
                        StrNewCode = StrNewCode & strTest1
                        IfThen2BooleanMessage cMod, lngCodeLineNo(0), lngCodeLineNo(4) - lngCodeLineNo(0) + 1, IfThen2BooleanSimplify(StrNewCode), StrCom
                       Else
                        StrNewCode = StrNewCode & " = Not (" & StrNewCode & ")"
                        IfThen2BooleanMessage cMod, lngCodeLineNo(0), lngCodeLineNo(4) - lngCodeLineNo(0) + 1, IfThen2BooleanSimplify(StrNewCode), StrCom
                      End If
                     ElseIf StrNewCode = "Not" & LeftWord(.Lines(lngCodeLineNo(1), 1)) Then
                      If strTest1 = " = False" Then
                        StrNewCode = LeftWord(.Lines(lngCodeLineNo(1), 1)) & strTest1
                        IfThen2BooleanMessage cMod, lngCodeLineNo(0), lngCodeLineNo(4) - lngCodeLineNo(0) + 1, IfThen2BooleanSimplify(StrNewCode), StrCom
                       Else
                        StrNewCode = LeftWord(.Lines(lngCodeLineNo(1), 1)) & " = Not " & LeftWord(.Lines(lngCodeLineNo(1), 1))
                        IfThen2BooleanMessage cMod, lngCodeLineNo(0), lngCodeLineNo(4) - lngCodeLineNo(0) + 1, IfThen2BooleanSimplify(StrNewCode), StrCom
                      End If
                      'End If
                     Else
                      'deal with all other cases
                      StrNewCode = Replace$(strCodeOnly(.Lines(lngCodeLineNo(1), 1)), strTest1, " = " & StrNewCode)
                      If strTest1 = " = False" Then
                        StrNewCode = Replace$(StrNewCode, " = ", " = Not(", , 1) & ")"
                      End If
                      IfThen2BooleanMessage cMod, lngCodeLineNo(0), lngCodeLineNo(4) - lngCodeLineNo(0) + 1, IfThen2BooleanSimplify(StrNewCode), StrCom
                      StrCom = vbNullString
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
Skip:
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop
  End With

End Sub

Private Sub IfThen2BooleanMessage(cMod As CodeModule, _
                                  StartLine As Long, _
                                  DelRange As Long, _
                                  StrNewCode As String, _
                                  StrCom As String)

  InsertNewCodeComment cMod, StartLine, DelRange, StrNewCode, WARNING_MSG & "If structure being used to set Boolean value collapsed." & IIf(InStrCode(StrNewCode, " = Not("), vbNewLine & _
   SUGGESTION_MSG & "You may be able to simplify the code further by coding out the 'Not'", vbNullString) & IIf(Len(StrCom), vbNewLine & _
   StrCom, vbNullString), True   ' Xcheck(XPrevCom)

End Sub

Private Function IfThen2BooleanSimplify(ByVal strCode As String) As String

  Dim strTemp As String

  IfThen2BooleanSimplify = strCode
  strTemp = strCodeOnly(IfThen2BooleanSimplify)
  If InStr(strTemp, "= Not(") Then
    If InStr(strTemp, " = 0)") > InStr(strTemp, "= Not(") Then
      If Right$(strTemp, 5) = " = 0)" Then
        IfThen2BooleanSimplify = Replace$(IfThen2BooleanSimplify, "= Not(", " = ", , 1)
        IfThen2BooleanSimplify = Replace$(IfThen2BooleanSimplify, " = 0)", vbNullString, , 1)
        IfThen2BooleanSimplify = IfThen2BooleanSimplify & vbNewLine & _
         WARNING_MSG & "Logic simplified from:" & vbNewLine & _
         RGSignature & strCode
      End If
    End If
    'v3.0.7 removed the logic fails if the test value should only be > 0 but could be negative. Thanks Lee Chase
    'ElseIf InStr(strTemp, " = ") < InStr(strTemp, " > 0") Then
    'v2.8.9
    '    If Right$(strTemp, 4) = " > 0" Then
    '      IfThen2BooleanSimplify = Replace$(IfThen2BooleanSimplify, " > 0", vbNullString, , 1)
    '    End If
  End If

End Function

Public Sub IfThenArrayExpander(ArrRoutine As Variant, _
                               Optional UpDated As Boolean)

  Dim L_CodeLine As String
  Dim RLine      As Long

  For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
    L_CodeLine = ArrRoutine(RLine)
    If Not JustACommentOrBlank(L_CodeLine) Then
      If InstrAtPositionArray(L_CodeLine, IpLeft, False, "If", "ElseIf") Or InStr(Trim$(L_CodeLine), " If ") > 1 Then
        'thanks to Manuel Muñoz for finding the bug the Or phrase above helps to fix
        If Left$(L_CodeLine, 6) <> "End If" Then
          ' if a structure is already expanded then this stops it hitting again
          L_CodeLine = IfThenStructureExpander(L_CodeLine)
          If ArrRoutine(RLine) <> L_CodeLine Then
            ArrRoutine(RLine) = L_CodeLine
            UpDated = True
          End If
        End If
      End If
    End If
  Next RLine

End Sub

Private Function IfThenStructureExpander(ByVal varName As String) As String

  Dim TmpA         As Variant
  Dim MyStr        As String
  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim ThenIfTest   As Long
  Dim I            As Long
  Dim strPreIfCode As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Expands single line If Then (Else)  structures into multiple line version
  'Any end of line Comment is placed above the code
  On Error GoTo BadError
  MyStr = varName
  ' remove end comments for restoring after changing Type Suffixes
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    'thanks to Manuel Muñoz for finding the bug this fixes
    ' consider code      "Red = Red * F:     If Red > 255 Then Red = 255: R1 = 1"
    ' add the necessary newline for rest of the fix to work
    RemoveExcessColonSpace MyStr
    If Left$(MyStr, 3) <> "If " Then
      If Left$(MyStr, 7) <> "ElseIf " Then
        If InCode(MyStr, InStr(MyStr, "If ")) Then
          strPreIfCode = Left$(MyStr, InStr(MyStr, "If "))
          strPreIfCode = Left$(MyStr, InStr(MyStr, "If ") - 1)
          MyStr = Mid$(MyStr, Len(strPreIfCode) + 1)
        End If
      End If
    End If
    MyStr = Safe_Replace(MyStr, ": If ", vbNewLine & "If ")
    'UPDATE colon test moved here to deal with unnecessary colons in If Then Structures
    If InstrAtPosition(MyStr, "Else:", ipAny, True) Then
      MyStr = Safe_Replace(MyStr, "Else:", vbNewLine & "Else" & vbNewLine)
    End If
    If InstrAtPosition(MyStr, "Then:", ipAny, True) Then
      MyStr = Safe_Replace(MyStr, "Then:", "Then")
    End If
    'For Code of format:    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
    ' this need extra End Ifs attached to end of string
    If InstrAtPosition(MyStr, "Then If ", ipAny, True) Then
      DisguiseLiteral MyStr, " Then If ", True
      ThenIfTest = CountSubString(MyStr, " Then If ")
      DisguiseLiteral MyStr, " Then If ", False
      If ThenIfTest > 0 Then
        MyStr = Safe_Replace(MyStr, " Then If ", " Then" & vbNewLine & "If ")
        For I = 1 To ThenIfTest
          MyStr = MyStr & vbNewLine & "End If"
        Next I
      End If
    End If
    DisguiseLiteral MyStr, "Then", True
    If InStr(MyStr, " Then ") Then
      If InCode(MyStr, InStr(MyStr, " Then ")) Then
        TmpA = Split(MyStr, " Then ")
        If LeftWord(MyStr) = "If" Then
          MyStr = Join(TmpA, " Then " & vbNewLine) & RepeatedString(vbNewLine & "End If", UBound(TmpA))
         ElseIf LeftWord(MyStr) = "ElseIf" Then
          'v 2.2.3 ElseIf handler
          MyStr = Join(TmpA, " Then " & vbNewLine)
        End If
      End If
    End If
    DisguiseLiteral MyStr, "Then", False
    DisguiseLiteral MyStr, "Else", True
    If InstrAtPosition(MyStr, "Else", ipAny, False) Then
      TmpA = Split(MyStr, " Else ")
      MyStr = Join(TmpA, vbNewLine & " Else " & vbNewLine)
    End If
    'special case probably poor coding but can be done
    If InstrAtPosition(MyStr, "Else" & vbNewLine & "End If", IpRight, True) Then
      TmpA = Split(MyStr, " Else" & vbNewLine)
      MyStr = Join(TmpA, vbNewLine & "Else" & vbNewLine)
      If InStr(MyStr, vbNewLine & "Else" & vbNewLine & "End If") Then
        MyStr = Replace$(MyStr, vbNewLine & _
         "Else" & vbNewLine & _
         "End If", vbNewLine & _
         "Else" & vbNewLine & _
         RGSignature & "This 'Else' statement may be unnecessary" & vbNewLine & _
         "End If")
      End If
    End If
    DisguiseLiteral MyStr, "Else", False
    'comments are placed above the structure to leave space for my comment after
    If LenB(strPreIfCode) Then
      MyStr = strPreIfCode & vbNewLine & MyStr
    End If
    If MyStr & CommentStore <> varName Then
      Select Case FixData(NIfThenExpand).FixLevel
       Case CommentOnly
        IfThenStructureExpander = varName & vbNewLine & _
         RGSignature & "'If..Then' lines should be expanded for readability."
       Case FixAndComment
        IfThenStructureExpander = CommentStore & vbNewLine & _
         SpaceOffSet & MyStr & RGSignature & "Structure Expanded."
       Case JustFix
        IfThenStructureExpander = CommentStore & vbNewLine & SpaceOffSet & MyStr
      End Select
     Else
      IfThenStructureExpander = varName
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  IfThenStructureExpander = varName

End Function

Private Function ImpossibleEmptyStringTest(varTest As Variant) As Boolean

  Dim I      As Long
  Dim arrTmp As Variant

  arrTmp = Split(varTest)
  For I = 2 To UBound(arrTmp) - 1
    If arrTmp(I) = "=" Then
      If SmartLeft(arrTmp(I + 1), EmptyString) Then
        If MultiLeft(arrTmp(I - 2), True, "Left", "Mid", "Right") Then
          ImpossibleEmptyStringTest = True
          Exit For
        End If
      End If
    End If
  Next I

End Function

Public Function InConditionalCompile(ByVal strTest As String) As Boolean

  'v 2.2.1
  'based on InTypeDef (modified because InstrAtPosition is confused by #
  'and start and end of structure don't have shared members
  'Detect if current line is inside an Enum def
  'Used by DecIndent to allow indenting of members

  If Left$(strTest, 4) = "#If " Then
    InConditionalCompile = True
    blnStaticInConditionalCompile = True
   ElseIf Left$(strTest, 7) = "#End If" Then
    InConditionalCompile = False
    blnStaticInConditionalCompile = False
   Else
    InConditionalCompile = blnStaticInConditionalCompile
  End If

End Function

Public Function InEnumDef(ByVal strTest As String, _
                          StrEnumName As String) As Boolean

  'v 2.2.1
  'based on InTypeDef
  'Detect if current line is inside an Enum def
  'Used by DecIndent to allow indenting of members

  If InstrAtPosition(strTest, "Enum", ipLeftOr2nd) Then
    StrEnumName = WordAfter(strTest, "Enum")
    InEnumDef = True
    blnStaticInEnumDef = True
    If InstrAtPosition(strTest, "End Enum", IpLeft) Then
      InEnumDef = False
      blnStaticInEnumDef = False
    End If
   Else
    InEnumDef = blnStaticInEnumDef
  End If

End Function

Public Sub InsertNewCodeComment(cMod As CodeModule, _
                                LNo As Long, _
                                ByVal DelRange As Long, _
                                ByVal strCode As String, _
                                StrCom As String, _
                                Optional ByVal PrevC As Boolean = False)

  Dim StrPrev As String

  With cMod
    If PrevC Then
      StrPrev = .Lines(LNo, DelRange)
      StrPrev = PREVIOUSCODE_MSG & Replace$(StrPrev, vbNewLine, vbNewLine & PREVIOUSCODE_MSG)
    End If
    .DeleteLines LNo, DelRange
    .InsertLines LNo, strCode
    If PrevC Then
      .InsertLines GetInsertionPoint(cMod, LNo + 1 + CountSubString(strCode, vbNewLine)), StrPrev
    End If
    If Len(StrCom) Then
      .InsertLines GetInsertionPoint(cMod, LNo + 1 + CountSubString(strCode, vbNewLine)), StrCom
    End If
  End With

End Sub

Private Function InstrControlWhichCantAcceptVBNull(ByVal strCode As String) As Boolean

  Dim TmpA        As Variant
  Dim I           As Long
  Dim lngTmpIndex As Long

  'thanks to knormal night who found the bug that made this test necessary
  strCode = ExpandForDetection(strCode)
  strCode = Replace$(strCode, ".", SngSpace)
  TmpA = Split(strCode)
  For I = LBound(TmpA) To UBound(TmpA)
    lngTmpIndex = CntrlDescMember(TmpA(I))
    If lngTmpIndex > -1 Then
      If Not InQSortArray(ArrQNoNullControl, CntrlDesc(lngTmpIndex).CDClass) Then
        InstrControlWhichCantAcceptVBNull = True
        Exit For 'unction
      End If
    End If
  Next I

End Function

Private Function InStringDblQuote(ByVal strTest As String) As Boolean

  Dim DQCount As Long
  Dim I       As Long
  Dim J       As Long

  DQCount = CountSubString(strTest, EmptyString)
  If DQCount Then
    If InStr(1, strTest, EmptyString) Then
      For I = 1 To CountSubString(strTest, DQuote)
        For J = 1 To DQCount
          If Not IsOdd(I) Then
            If QatPosPosition(strTest, I, DQuote) = QatPosPosition(strTest, J, EmptyString) Then
              InStringDblQuote = True
              Exit For
            End If
          End If
        Next J
        If InStringDblQuote Then
          Exit For
        End If
      Next I
    End If
  End If

End Function

Private Sub Integer2Long(cMod As CodeModule)

  
  Dim strDecType      As String
  Dim bFuncUpdated    As Boolean
  Dim MsgUpGradeProc  As String
  Dim UpDated         As Boolean
  Dim MUpdated        As Boolean
  Dim bisUdProc       As Boolean
  Dim ModuleNumber    As Long
  Dim Member          As Long
  Dim RLine           As Long
  Dim TestPos         As Long
  Dim MemberCount     As Long
  Dim LineRange       As Long
  Dim L_CodeLine      As String
  Dim CommentStore    As String
  Dim CommentStore2   As String
  Dim StrTypeName     As String
  Dim ArrMember       As Variant
  Dim ArrRoutine      As Variant
  Dim strTarget       As String
  Dim StrUpdatedParam As String
  Dim strCUrProc      As String

  'strTypeDefWarning = "Type definitions Integers may be necessary for Windows/API usage" & vbNewLine & "but other User Defined Type definitions could be upgraded to Long"
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Because Longs fit DWord boundries better than Integer it is better to use Longs even when Integers are big enough to hold values
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, UpdateInteger2Long) Then
    ArrMember = GetModuleArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        InTypeDef ArrMember(Member), StrTypeName ' track type membership
        If InstrAtPosition(ArrMember(Member), "As Integer", ipAny) Then
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
            L_CodeLine = GetWholeLineArray(ArrRoutine, RLine, LineRange)
            If Not JustACommentOrBlank(L_CodeLine) Then
              'updating Integer to Long needs different approaches for
              'Procedure heads, Types and Dim/Public/Private
              If ExtractCode(L_CodeLine, CommentStore) Then
                If isProcHead(L_CodeLine) Then
                  strCUrProc = GetProcNameStr(L_CodeLine)
                  MemberMessage strCUrProc, Member, MemberCount
                End If
                'If LenB(strCUrProc) = 0 Then
                '                InTypeDef L_CodeLine, StrTypeName ' track type membership
                'Else
                ' StrTypeName = ""
                'End If
                If UpdatableInteger(L_CodeLine, cMod.Parent, bisUdProc, TestPos) Then
                  If dofix(ModuleNumber, UpdateDimType) Then
                    If InStr(L_CodeLine, "%") Then
                      L_CodeLine = TypeSuffixExtender(L_CodeLine)
                      ExtractCode L_CodeLine, CommentStore2
                      CommentStore = CommentStore & vbNewLine & CommentStore2
                    End If
                  End If
                  If isProcHead(L_CodeLine) Then
                    'Procedure heads
                    If NotVBControlCode(L_CodeLine, cMod, bisUdProc) Then
                      Select Case FixData(UpdateInteger2Long).FixLevel
                       Case CommentOnly
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Integer parameter(s) could be upgraded to Long.", MAfter)
                       Case FixAndComment
                        StrUpdatedParam = vbNullString
                        Do While InStr(L_CodeLine, " As Integer")
                          strTarget = WordBefore(ExpandForDetection(L_CodeLine), " As Integer")
                          If Right$(strTarget, 1) <> ")" Then
                            StrUpdatedParam = AccumulatorString(StrUpdatedParam, strInSQuotes(strTarget))
                            ReplaceType L_CodeLine, strTarget & " As Integer", strTarget & " As Long"
                           Else
                            bFuncUpdated = True
                            ReplaceType L_CodeLine, strTarget & " As Integer", strTarget & " As Long"
                          End If
                          UpdateIntegerTypeSuffix ArrRoutine, strTarget
                        Loop
                        L_CodeLine = L_CodeLine & CommentStore
                        UpDated = True
                        ArrRoutine(RLine) = L_CodeLine & CommentStore
                        MsgUpGradeProc = IIf(Len(StrUpdatedParam), WARNING_MSG & "Integer parameter" & IIf(InStr(StrUpdatedParam, ","), "s ", " ") & StrUpdatedParam & " upgraded to Long.", vbNullString)
                        If bFuncUpdated Then
                          MsgUpGradeProc = MsgUpGradeProc & WARNING_MSG & IIf(InStr(L_CodeLine, "Property Get"), "Property", "Function") & " upgraded to Long"
                        End If
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, MsgUpGradeProc, MAfter)
                        MsgUpGradeProc = vbNullString
                        bFuncUpdated = False
                      End Select
                    End If
                   ElseIf InTypeDef(L_CodeLine, StrTypeName) Then
                    'Types
                    If Not PassedToAPI(StrTypeName) Then
                      Select Case FixData(UpdateInteger2Long).FixLevel
                       Case CommentOnly
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Integer Type member could be upgraded to Long.", MAfter)
                       Case FixAndComment
                        Do While InStr(L_CodeLine, " As Integer")
                          strTarget = WordBefore(L_CodeLine, " As Integer")
                          'v2.8.9  no space at start of reps below becuase ReplaceTpe adds them
                          'caused a lock up on code of form '"Dim R As Integer, G As Integer, B As Integer"
                          ReplaceType L_CodeLine, "As Integer", "As Long"
                        Loop
                        L_CodeLine = L_CodeLine & CommentStore
                        UpDated = True
                        ArrRoutine(RLine) = L_CodeLine & CommentStore
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, WARNING_MSG & "Integer Type member upgraded to Long.", MAfter)
                      End Select
                     Else
                      ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, WARNING_MSG & "Integer Type members of a Type passed to API must remain as Integer.", MAfter)
                      'v 2.1.8 If the only comments generated are these then it will not update and these messages will not show
                      'Updated = True
                    End If
                   Else
                    'Scopes: Dim/Public/Private
                    strDecType = LeftWord(L_CodeLine)
                    'v2.4.4 reconfigured to short circuit
                    If strDecType <> "Const" Then
                      If InStr(L_CodeLine, " Const ") Then
                        strDecType = strDecType & " Const"
                      End If
                    End If
                    If InstrAtPositionArray(L_CodeLine, ipLeftOr2ndOr3rd, True, "Event", "WithEvents") Then
                      ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, WARNING_MSG & "Integer " & IIf(InStr(L_CodeLine, "WithEvents"), "WithEvents", "Event") & " parameters must remain as Integer.", MAfter)
                      'v 2.1.8 If the only comments generated are these then it will not update and these messages will not show
                      'Updated = True
                     Else
                      Select Case FixData(UpdateInteger2Long).FixLevel
                       Case CommentOnly
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Integer " & strDecType & " could be upgraded to Long.", MAfter)
                       Case FixAndComment
                        ReplaceType L_CodeLine, "As Integer", "As Long"
                        ArrRoutine(RLine) = L_CodeLine & CommentStore
                        UpDated = True
                        ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, WARNING_MSG & "Integer " & strDecType & " upgraded to Long.", MAfter)
                      End Select
                    End If
                  End If
                End If
              End If
            End If
          Next RLine
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      If MUpdated Then
        ReWriter cMod, ArrMember, RWModule
      End If
    End If
  End If

End Sub

Public Function InTypeDef(ByVal strTest As String, _
                          StrTypeName As String) As Boolean

  'Copyright 2003 Roger Gilchrist
  'email: rojagilkrist@hotmail.com
  'Detect if current line is inside a Type def
  'Used by Integer2Long to stop it touching TypeDef members

  If InstrAtPosition(strTest, "Type", ipLeftOr2nd) Then
    StrTypeName = WordAfter(strTest, "Type")
    InTypeDef = True
    blnStaticInTypeDef = True
    If InstrAtPosition(strTest, "End Type", IpLeft) Then
      InTypeDef = False
      blnStaticInTypeDef = False
      StrTypeName = vbNullString
    End If
   Else
    InTypeDef = blnStaticInTypeDef
  End If

End Function

Private Function IsGotoTarget(ByVal strTest As String) As Boolean

  'v3.0.3 simple test

  strTest = strCodeOnly(strTest)
  IsGotoTarget = InStr(strTest, " ") = 0 And Right$(strTest, 1) = ":"

End Function

Private Function IsNumericColon(ByVal varTest As Variant) As Boolean

  If Right$(varTest, 1) = Colon Then
    IsNumericColon = IsNumeric(Left$(varTest, Len(varTest) - 1))
  End If

End Function

Public Function IsOnErrorCode(varCode As Variant, _
                              Optional ByVal All0Goto1Resume2 As Long = 0) As Boolean

  'v2.9.0 simplifyied detection
  'v2.9.7 added On Local tests

  Select Case All0Goto1Resume2
   Case 0
    IsOnErrorCode = MultiLeft(varCode, True, "On Error GoTo", "On Error Resume", "On Local Error GoTo", "On Local Error Resume")
   Case 1
    IsOnErrorCode = MultiLeft(varCode, True, "On Error GoTo", "On Local Error GoTo")
   Case 2
    IsOnErrorCode = MultiLeft(varCode, True, "On Error Resume", "On Local Error Resume")
  End Select

End Function

Private Function IsUnneededExit(cMod As CodeModule, _
                                CurLine As Long, _
                                Optional ModeRet As Long) As Boolean

  
  Dim ArrProc        As Variant
  Dim ProcStart      As Long
  Dim ProcEnd        As Long
  Dim strProcName    As String
  Dim lngPRocKind    As Long
  Dim J              As Long
  Dim K              As Long
  Dim ProcTargetLine As Long
  Dim ElseDeep       As Long

  'v2.8.4
  ProcStart = GetProcStartLine(cMod, CurLine, strProcName, lngPRocKind)
  ProcEnd = GetProcEndLine(cMod, CurLine, strProcName, lngPRocKind)
  ArrProc = Split(cMod.Lines(ProcStart, ProcEnd - ProcStart), vbNewLine)
  ProcTargetLine = CurLine - ProcStart
  If NothingButEndIfOrLastCodeLine(ArrProc, ProcTargetLine, 0) Then
    IsUnneededExit = True
    ModeRet = 0
   ElseIf NothingButEndIfOrLastCodeLine(ArrProc, ProcTargetLine, 1) Then
    IsUnneededExit = True
    ModeRet = 1
   ElseIf NothingButEndIfOrLastCodeLine(ArrProc, ProcTargetLine, 2) Then
    IsUnneededExit = True
    ModeRet = 2
   Else
    For K = ProcTargetLine + 1 To UBound(ArrProc) - 1
      If Not JustACommentOrBlank(ArrProc(K)) Then
        If LeftWord(ArrProc(K)) = "Else" Then
          ElseDeep = GetStructureDepthLine(ArrProc, K)
          For J = K + 1 To UBound(ArrProc) - 1
            If Left$(ArrProc(J), 6) = "End If" Then
              If ElseDeep - 1 = GetStructureDepthLine(ArrProc, J) Then
                If NothingButEndIfOrLastCodeLine(ArrProc, J) Then
                  IsUnneededExit = True
                End If
                GoTo JumpOut
              End If
            End If
          Next J
         Else
          Exit For
        End If
      End If
      If IsUnneededExit Then
        Exit For
      End If
    Next K
JumpOut:
    'v2.9.6 new extends the search for code which is logically leapable thus allowing removal of an exit proc
    If Not IsUnneededExit Then
      ElseDeep = GetStructureDepthLine(ArrProc, ProcTargetLine)
      For K = ProcTargetLine + 1 To UBound(ArrProc) - 1
        If Not JustACommentOrBlank(ArrProc(K)) Then
          If Left$(ArrProc(K), 6) = "End If" Then
            For J = K + 1 To UBound(ArrProc) - 1
              If ElseDeep = GetStructureDepthLine(ArrProc, J) Then
                If NothingButEndIfOrLastCodeLine(ArrProc, J - 1) Then
                  IsUnneededExit = True
                End If
                GoTo jumpout2
              End If
            Next J
           Else
            Exit For
          End If
        End If
        If IsUnneededExit Then
          Exit For
        End If
      Next K
    End If
jumpout2:
    If Not IsUnneededExit Then
      For K = ProcTargetLine + 1 To UBound(ArrProc) - 1
        If Not JustACommentOrBlank(ArrProc(K)) Then
          If LeftWord(ArrProc(K)) = "ElseIf" Then
            ElseDeep = GetStructureDepthLine(ArrProc, K)
            For J = K + 1 To UBound(ArrProc) - 1
              If Left$(ArrProc(J), 6) = "End If" Then
                If ElseDeep - 1 = GetStructureDepthLine(ArrProc, J) Then
                  If NothingButEndIfOrLastCodeLine(ArrProc, J) Then
                    IsUnneededExit = True
                    Exit For
                  End If
                End If
              End If
            Next J
           Else
            Exit For
          End If
        End If
        If IsUnneededExit Then
          Exit For
        End If
      Next K
    End If
  End If
  'v3.0.3 nothing separates ExitProc from End Proc except a GoToTarget label:
  If Not IsUnneededExit Then
    IsUnneededExit = True
    For K = ProcTargetLine + 1 To UBound(ArrProc) - 1
      If Not JustACommentOrBlank(ArrProc(K)) Then
        If Not IsGotoTarget(ArrProc(K)) Then
          IsUnneededExit = False
          Exit For
        End If
      End If
    Next K
  End If

End Function

Private Sub LineNumberRemoval(cMod As CodeModule)

  Dim ModuleNumber           As Long
  Dim UpDated                As Boolean
  Dim MUpdated               As Boolean
  Dim ArrMember              As Variant
  Dim ArrRoutine             As Variant
  Dim TmpC                   As Variant
  Dim Member                 As Long
  Dim RLine                  As Long
  Dim MemberCount            As Long
  Dim IsNumericColonDetected As Boolean

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, RemoveLineNum) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        If UBound(ArrRoutine) > -1 Then
          For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
            TmpC = Split(ArrRoutine(RLine))
            If UBound(TmpC) > -1 Then
              If IsNumericColon(TmpC(0)) Then
                TmpC(0) = Left$(TmpC(0), Len(TmpC(0)) - 1)
                IsNumericColonDetected = True
              End If
              If IsNumeric(TmpC(0)) Then
                UpDated = True
                If Not RoutineSearch(ArrRoutine, "GoTo " & TmpC(0), , , True) Then
                  'v2.8.3 'On X GoSub/SoTo num1, num2, num3' code very old style( pre Select Case/ElseIf)
                  ' now detected
                  If Not MultiGoToGoSub(ArrRoutine, TmpC(0), UpDated) Then
                    TmpC(0) = vbNullString
                    AddNfix RemoveLineNum
                  End If
                End If
                If LenB(TmpC(0)) Then
                  If IsNumericColonDetected Then
                    TmpC(0) = TmpC(0) & Colon
                  End If
                End If
                'Ver 1.0.97 fix
                'THANKS to Dream for finding the need to insert Trim in next code line
                ' the bug doesn't actually hit until the IfThenStructureExpand routine
                'Very nasty one.
                ArrRoutine(RLine) = Trim$(Join(TmpC))
                If LenB(TmpC(0)) Then
                  TmpC(0) = TmpC(0) & vbNewLine & _
                   WARNING_MSG & "Line number is the target of at least one 'GoTo' in this routine." & vbNewLine & _
                   RGSignature & "Use Find & Replace to replace it with a named label; line numbers make little sense in VB." & vbNewLine
                  ArrRoutine(RLine) = Trim$(Join(TmpC))
                  ArrRoutine(RLine) = Replace$(ArrRoutine(RLine), vbNewLine & SngSpace, vbNewLine)
                End If
              End If
            End If
          Next RLine
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Sub LoadUnLoadErrorResumeFast(cMod As CodeModule)

  'v3.0.7 new fast version
  'test whether a routine calls Load/Unload and insert Error Checking if necessary
  
  Dim PSLine        As Long
  Dim PEndLine      As Long
  Dim TopOfRoutine  As Long
  Dim ArrProc       As Variant
  Dim L_CodeLine    As String
  Dim UpDateRoutine As Boolean
  Dim RLine         As Long
  Dim Sline         As Long
  Dim UpDated       As Long

  With cMod
    If dofix(ModDescMember(.Parent.Name), NCloseResumeLoad) Then
      Do While .Find("Load", Sline, 1, -1, -1, True, True)
        If InCode(.Lines(Sline, 1), InStr(.Lines(Sline, 1), "Load")) Then
          If InStrWholeWordRX(.Lines(Sline, 1), "Load") Then ' skip "Form_Load"
            ArrProc = ReadProcedureCodeArray2(cMod, Sline, PSLine, PEndLine, TopOfRoutine)
            MemberMessage GetProcName(cMod, Sline), Sline, .CountOfLines
            For RLine = LBound(ArrProc) To UBound(ArrProc)
              L_CodeLine = ArrProc(RLine)
              If Not JustACommentOrBlank(L_CodeLine) Then
                If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Load", "Unload") Then
                  If Not InstrAtPosition(L_CodeLine, "Unload Me", IpLeft, True) Then
                    If Not SearchRoutineArray2(ArrProc, "On Error Resume Next") Then
                      If Not SearchRoutineArray2(ArrProc, "On Error Goto") Then
                        ' ver 1.1.93 thanks to Morgan Haueisen for noticing this bug
                        'CF was adding double Error Trapping
                        '(harmless but ugly as the original Error Trap would override the CF inserted one)
                        Select Case FixData(NCloseResumeLoad).FixLevel
                         Case CommentOnly
                          ArrProc(RLine) = Marker(ArrProc(RLine), SUGGESTION_MSG & "Load/UnLoad can be made safer by using 'On Error Resume Next'", MAfter, UpDateRoutine)
                          UpDated = True
                         Case FixAndComment, JustFix 'Always comment
                          ArrProc(TopOfRoutine) = Marker(ArrProc(TopOfRoutine) & vbNewLine & _
                           "On Error Resume Next", RISK_MSG & " Load/UnLoad are safer with Error Trapping", MAfter, UpDateRoutine)
                          UpDated = True
                        End Select
                      End If
                    End If
                    Exit For ' found it or didn't once that good enough
                  End If
                End If
              End If
            Next RLine
            If UpDated Then
              ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine, False
              UpDated = False
            End If
          End If
        End If
        Sline = Sline + IIf(PEndLine - PSLine > 1, PEndLine - PSLine, 1)
        PEndLine = 0
        PSLine = 0
        If Sline > .CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

End Sub

Private Function MultiGoToGoSub(arrR As Variant, _
                                NumLabel As Variant, _
                                UpDated As Boolean) As Boolean

  Dim I      As Long
  Dim strMsg As String

  'v2.8.3 Thanks Joakim Schramm
  'this detects code of the format
  'On L - 18 GoSub 1430, 1440, 1450, 1460, 1470
  'this is also possible with GoTo so routine checks for both
  'this is very rare it was a way of doing the equivalant of Select Case or ElseIf
  'in BasicA which didn't have these structures
  'In CF this routine protects the code from line number removal hitting the line numbers
  'Also inserts a suggestion to upgrade to Select Case
  strMsg = SUGGESTION_MSG & "Obsolete Multiple Target 'On X GoSub' or 'On X GoTo' structures" & vbNewLine & _
   RGSignature & "should be replaced with a Select Case structure"
  If InStr(Join(arrR, vbNewLine), vbNewLine & "On ") Then ' most code will fail at this
    If InStr(Join(arrR), " GoSub ") Or InStr(Join(arrR), " GoTo ") Then
      For I = 0 To UBound(arrR)
        If LeftWord(arrR(I)) = "On" Then
          If InStr(arrR(I), NumLabel) Then
            MultiGoToGoSub = True
            If InStr(arrR(I), strMsg) = False Then
              arrR(I) = arrR(I) & vbNewLine & strMsg
              UpDated = True
            End If
          End If
        End If
      Next I
    End If
  End If

End Function

Private Sub MultilineCommentExtractor(cMod As CodeModule, _
                                      ByVal StartLine As Long, _
                                      ByVal EndLine As Long, _
                                      StrCom As String)

  Dim I          As Long
  Dim strComment As String

  'v2.5.0 support for the new structure collapse fixes
  StrCom = vbNullString
  For I = StartLine To EndLine
    ExtractCode cMod.Lines(I, 1), strComment
    If Len(strComment) Then
      StrCom = StrCom & strComment & vbNewLine
    End If
  Next I

End Sub

Private Function MultiStructureExit(arrR As Variant, _
                                    ByVal RLine As Long) As Long

  Dim I            As Long
  Dim DepthCounter As Long
  Dim arrTest1     As Variant
  Dim arrTest2     As Variant

  arrTest1 = Array("For", "Do", "If")
  arrTest2 = Array("Next", "Loop")
  'v2.7.2 support routine for Exit fix
  'stops the fix comenting on Exits used to get out of deeply nested code
  For I = LBound(arrR) To RLine - 1
    If IsInArray(LeftWord(arrR(I)), arrTest1) Then
      DepthCounter = DepthCounter + 1
     ElseIf IsInArray(LeftWord(arrR(I)), arrTest2) Then
      DepthCounter = DepthCounter - 1
     ElseIf Left$(arrR(I), 6) = "End If" Then
      DepthCounter = DepthCounter - 1
    End If
  Next I
  MultiStructureExit = DepthCounter > 1

End Function

Private Function NothingButEndIfOrLastCodeLine(ArrProc As Variant, _
                                               ByVal ScanFrom As Long, _
                                               Optional Modes As Long = 0) As Boolean

  Dim I           As Long
  Dim StructCount As Long

  'v2.8.4
  NothingButEndIfOrLastCodeLine = True
  For I = ScanFrom + 1 To UBound(ArrProc) - 1
    If Not JustACommentOrBlank(ArrProc(I)) Then
      '      If Not Left$(ArrProc(I), 6) = "End If" Then
      'v3.0.0 don't apply to double structures exits
      If InstrAtPositionArray(ArrProc(I), IpLeft, True, "Loop", "Next") Then
        StructCount = StructCount + 1
        If StructCount > 1 Then
          NothingButEndIfOrLastCodeLine = False
          Exit For
        End If
      End If
      'v.2.8.8 more complex tests to extend fix
      If Not InstrAtPositionArray(ArrProc(I), IpLeft, True, "End If", "Err.Clear", "On Error GoTo 0", IIf(Modes = 0, vbNullString, IIf(Modes = 1, "Loop", "Next"))) Then
        NothingButEndIfOrLastCodeLine = False
        Exit For
      End If
    End If
  Next I

End Function

Private Function NotVBControlCode(strCode As String, _
                                  cMod As CodeModule, _
                                  Optional isUDProc As Boolean) As Boolean

  Dim strProcName As String

  isUDProc = False
  If Not isProcHead(strCode) Then
    NotVBControlCode = True
   Else
    If InStr(strCode, " Declare ") = 0 Then
      strProcName = GetProcNameStr(strCode)
      If Len(strProcName) Then
        NotVBControlCode = Not RoutineNameIsVBGenerated(strProcName, cMod.Parent)
        If HasUDTParameter(strProcName) Then
          isUDProc = True
        End If
       Else
        NotVBControlCode = False
      End If
    End If
  End If

End Function

Private Sub OnErrorResumeCloseFast(cMod As CodeModule)

  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim ArrProc      As Variant
  Dim ModuleNumber As Long
  Dim Sline        As Long
  Dim Hit          As Boolean
  Dim Possible     As Boolean
  Dim UpDated      As Boolean
  Dim I            As Long
  Dim EndOfRoutine As Long

  With cMod
    ModuleNumber = ModDescMember(.Parent.Name)
    If dofix(ModuleNumber, NCloseResume) Then
      Do While .Find("On Error Resume Next", Sline, 1, -1, -1, True, True)
        If InCode(.Lines(Sline, 1), InStr(.Lines(Sline, 1), "On Error Resume Next")) Then
          ArrProc = ReadProcedureCodeArray2(cMod, Sline, PSLine, PEndLine)
          MemberMessage GetProcName(cMod, Sline), Sline, .CountOfLines
          '-----------------------------------------------
          Possible = Not SearchRoutineArray2(ArrProc, "On Error GoTo 0")
          If Possible Then
            '            ArrRoutine = Split(ArrProc, vbNewLine)
            If SearchRoutineArray2(ArrProc, "Exit ") Then
              'ver 1.1.93 another  Morgan Haueisen suggestion
              'backtrack for Exit Procedure code protecting an existing error handler
              Hit = False
              For I = UBound(ArrProc) To LBound(ArrProc) Step -1
                If Left$(ArrProc(I), 5) = "Exit " Then
                  Hit = True
                  EndOfRoutine = I - 1
                  ArrProc(EndOfRoutine) = ErrorResumeCloser(ArrProc(EndOfRoutine))
                  Exit For
                End If
              Next I
              If Not Hit Then  ' safety should never hit
                EndOfRoutine = GetEndOfRoutine(ArrProc)
                ArrProc(EndOfRoutine) = ErrorResumeCloser(ArrProc(EndOfRoutine))
              End If
             Else
              EndOfRoutine = GetEndOfRoutine(ArrProc)
              ArrProc(EndOfRoutine) = ErrorResumeCloser(ArrProc(EndOfRoutine))
            End If
            UpDated = True
            Possible = False
          End If
          '----------------------------------------
          If UpDated Then
            ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine, False
            UpDated = False
          End If
        End If
        Sline = Sline + IIf(PEndLine - PSLine > 1, PEndLine - PSLine, 1)
        PEndLine = 0
        PSLine = 0
        If Sline > .CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

End Sub

Private Function PassedToAPI(ByVal strTest As String) As Boolean

  ' ver 1.1.93 thanks to Morgan Haueisen for noticing the bug this protects against
  'v 2.1.7 updated to single function
  'v2.6.5 see RequiredIntegerType for details
  'v2.6.7 corrected test order/layout for amax efficency

  PassedToAPI = RequiredIntegerType(strTest)
  If Not PassedToAPI Then
    PassedToAPI = PassedToAPI2(strTest)
  End If
  If Not PassedToAPI Then
    PassedToAPI = TypeUsedByAPI(strTest)
  End If
  If Not PassedToAPI Then
    PassedToAPI = TypeUsedByAPI2(strTest)
  End If
  If Not PassedToAPI Then
    PassedToAPI = PassedToAPI3(strTest)
  End If
  If Not PassedToAPI Then
    PassedToAPI = PassedToGet(strTest)
  End If

End Function

Private Function PassedToAPI2(ByVal strTest As String) As Boolean

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim GuardLine  As Long
  Dim StartLine  As Long
  Dim L_CodeLine As String

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        GuardLine = 0
        With Comp
          Do While .CodeModule.Find(strTest, StartLine, 1, -1, -1, True, True)
            L_CodeLine = .CodeModule.Lines(StartLine, 1)
            'Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strTest, L_CodeLine, StartLine)
            If GuardLine > 0 Then
              If GuardLine > StartLine Then
                Exit Do
              End If
            End If
            'v2.8.1 bug caused by a module being only Declaration section and
            ' a legal but mesy variable with san name as its Type ie 'Board As Board'
            If WordBefore(L_CodeLine, "As") = WordAfter(L_CodeLine, "As") Then
              '
              Exit Do
            End If
            If Trim$(L_CodeLine) Like "*Declare * Lib *" & strTest Then
              PassedToAPI2 = True
              Exit Do
            End If
            If Trim$(L_CodeLine) Like "*As " & strTest Then
              ' recursive search for something typed as UDT
              If PassedToAPI2(WordBefore(L_CodeLine, "As")) Then
                PassedToAPI2 = True
                Exit Do
              End If
            End If
            StartLine = StartLine + 1
            GuardLine = StartLine
          Loop
        End With 'Comp
      End If
      If PassedToAPI2 Then
        Exit For
      End If
    Next Comp
    If PassedToAPI2 Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Private Function PassedToAPI3(ByVal strTest As String) As Boolean

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim GuardLine  As Long
  Dim StartLine  As Long
  Dim L_CodeLine As String
  Dim strAPIUsed As String

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        GuardLine = 0
        With Comp
          Do While .CodeModule.Find(strTest, StartLine, 1, -1, -1, True, True)
            L_CodeLine = .CodeModule.Lines(StartLine, 1)
            '          Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strTest, L_CodeLine,  StartLine)
            If GuardLine > 0 Then
              If GuardLine > StartLine Then
                Exit Do
              End If
            End If
            If IsAPIinLine(L_CodeLine, strAPIUsed) Then
              If InStr(L_CodeLine, strAPIUsed) < InStr(L_CodeLine, strTest) Then
                PassedToAPI3 = True
                Exit Do
              End If
            End If
            StartLine = StartLine + 1
            GuardLine = StartLine
          Loop
        End With 'Comp
      End If
      If PassedToAPI3 Then
        Exit For
      End If
    Next Comp
    If PassedToAPI3 Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Private Function PassedToGet(ByVal strTest As String) As Boolean

  'v3.0.6 this catches Types used to define variables filled by using the Get command Thanks Alfred Koppold for bring it to my attention
  
  Dim Proj            As VBProject
  Dim Comp            As VBComponent
  Dim GuardLine       As Long
  Dim StartLine       As Long
  Dim L_CodeLine      As String
  Dim strTestTypedVar As String

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        GuardLine = 0
        With Comp.CodeModule
          Do While .Find(strTest, StartLine, 1, -1, -1, True, True)
            If GuardLine > StartLine Then
              Exit Do
            End If
            GuardLine = StartLine
            L_CodeLine = .Lines(StartLine, 1)
            If Not JustACommentOrBlank(L_CodeLine) Then
              If L_CodeLine Like "Get * " & strTest Then ' should never hit but just in case
                PassedToGet = True
                Exit Do
              End If
              If L_CodeLine Like "* As " & strTest Then
                strTestTypedVar = WordBefore(L_CodeLine, "As")
                If InStr(strTestTypedVar, "(") Then ' in case passed as varName()
                  strTestTypedVar = Left$(strTestTypedVar, InStr(strTestTypedVar, "(") - 1)
                End If
                If FindCodeUsageLike("Get *, " & strTestTypedVar & "*") Then
                  PassedToGet = True
                  Exit Do
                End If
                'End If
              End If
            End If
            StartLine = StartLine + 1
            If StartLine > .CountOfLines Then
              Exit Do
            End If
          Loop
        End With 'Comp
      End If
      If PassedToGet Then
        Exit For
      End If
    Next Comp
    If PassedToGet Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Private Sub PleonasmCleanerEngine(cMod As CodeModule, _
                                  bTarget As Boolean)

  Dim StartLine As Long
  Dim strTmp    As String
  Dim strTF     As String

  'v2.7.8 Modifications suggested by George E's article 'Increase your game's FPS tips'
  strTF = IIf(bTarget, "True", "False")
  With cMod
    Do While .Find(" = " & strTF, StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        If InStr(.Lines(StartLine, 1), strTF) Then
          If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), strTF)) Then
            If InstrAtPositionArray(.Lines(StartLine, 1), ipAny, True, "= " & strTF & " Then", "= " & strTF & " And", "= " & strTF & " Or", "= " & strTF & ")", "= " & strTF) Then
              'don't touch parameters dealt with in PleonasmCleaner
              strTmp = PleonasmCleaner(.Lines(StartLine, 1), bTarget)
              If .Lines(StartLine, 1) <> strTmp Then
                .ReplaceLine StartLine, strTmp
              End If
            End If
          End If
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Sub ProcArrayErrTrapRenamer(ArrProc As Variant, _
                                    strOld As String, _
                                    Strnew As String, _
                                    ByVal strMsg As String, _
                                    UpDated As Boolean)

  Dim I       As Long
  Dim strTemp As String

  '
  For I = LBound(ArrProc) To UBound(ArrProc)
    If InStr(ArrProc(I), "GoTo " & strOld) Then
      strTemp = ArrProc(I)
      If IsRealWord(strTemp, strOld) Then
        WholeWordReplacer strTemp, "GoTo " & strOld, "GoTo " & Strnew
        If ArrProc(I) <> strTemp Then
          UpDated = True
          ArrProc(I) = strTemp & strMsg
        End If
      End If
    End If
    If InStr(ArrProc(I), strOld & ":") Then
      strTemp = ArrProc(I)
      If IsRealWord(strTemp, strOld & ":") Then
        WholeWordReplacer strTemp, strOld & ":", Strnew & ":"
        If ArrProc(I) <> strTemp Then
          UpDated = True
          ArrProc(I) = strTemp & strMsg
        End If
      End If
    End If
  Next I

End Sub

Public Function PropertyBuilder(LetSet As String, _
                                ProcName As String, _
                                strType As String, _
                                strPrivateVariableName As String) As String

  'support for CreatePropertyFromPublicVariableInClass
  'ver 1.1.93 simplified

  PropertyBuilder = vbNewLine & _
   "Public Property " & LetSet & SngSpace & ProcName & strInBrackets("PropVal As " & strType) & vbNewLine
  PropertyBuilder = PropertyBuilder & RGSignature & "Property created to replace Public Variable" & vbNewLine
  PropertyBuilder = PropertyBuilder & IIf(LetSet = "Set", "Set ", vbNullString) & strPrivateVariableName & " = PropVal" & vbNewLine
  PropertyBuilder = PropertyBuilder & "End Property" & vbNewLine
  PropertyBuilder = PropertyBuilder & "Public Property Get " & ProcName & "() As " & strType & vbNewLine
  PropertyBuilder = PropertyBuilder & RGSignature & "Property created to replace Public Variable" & vbNewLine
  PropertyBuilder = PropertyBuilder & IIf(LetSet = "Set", "Set ", vbNullString) & ProcName & EqualInCode & strPrivateVariableName & vbNewLine
  PropertyBuilder = PropertyBuilder & "End Property" & vbNewLine

End Function

Private Function QatPosPosition(ByVal strTest As String, _
                                ByVal PosinStr As Long, _
                                ByVal strFind As String) As Long

  Dim LQPos   As Long
  Dim LstrPos As Long

  Do
    LQPos = InStr(LQPos + 1, strTest, strFind)
    LstrPos = LstrPos + 1
  Loop Until LstrPos = PosinStr Or LQPos = 0
  QatPosPosition = LQPos

End Function

Private Sub RemoveExcessColonSpace(strCode As String)

  'thanks to Manuel Muñoz for finding the bug this helps fix
  'it should also take care of some other very rare bugs that have occasionally popped up
  'caused by unnecessary spaces blocking other parts of the IfThenStructureExpander routine

  DisguiseLiteral strCode, ":  ", True
  Do While InStr(strCode, ":  ") ' strip the unnecessary spaces
    strCode = Safe_Replace(strCode, ":  ", vbNewLine & ": ")
  Loop
  DisguiseLiteral strCode, ":  ", False

End Sub

Private Sub replaceAllinComp(cMod As CodeModule, _
                             ByVal strNme As String, _
                             ByVal strNewNme As String)

  Dim arrDec     As Variant
  Dim arrLine    As Variant
  Dim I          As Long
  Dim J          As Long
  Dim L_CodeLine As String

  arrDec = GetMembersArray(cMod)
  For I = LBound(arrDec) To UBound(arrDec)
    arrLine = Split(arrDec(I), vbNewLine)
    For J = LBound(arrLine) To UBound(arrLine)
      If InStrWholeWordRX(arrLine(J), strNme) Then
        arrLine(J) = Safe_Replace(arrLine(J), strNme, strNewNme) ', , ,  False)
        If InStr(arrLine(J), "." & strNewNme) Then
          L_CodeLine = arrLine(J)
          L_CodeLine = Replace$(L_CodeLine, "." & strNewNme, "." & strNme)
          arrLine(J) = L_CodeLine
        End If
      End If
    Next J
    arrDec(I) = Join(arrLine, vbNewLine)
  Next I
  ReWriteMembers cMod, arrDec, True

End Sub

Private Sub ReplaceType(strCode As String, _
                        strOld As String, _
                        Strnew As String)

  'ver 2.1.5 seperated out of Integer 2 Long may be useful elsewhere

  strCode = Safe_Replace(strCode, " " & strOld & ",", " " & Strnew & ",")
  If SmartRight(strCode, " " & strOld) Then
    strCode = Safe_Replace(strCode, " " & strOld, " " & Strnew)
  End If
  If InStr(strCode, " " & strOld & " =") Then
    'ver 1.1.93 this allows parameters in optional with value to update
    'no longer for Constants only
    strCode = Safe_Replace(strCode, " " & strOld & " =", " " & Strnew & " = ")
  End If
  If InStr(strCode, " " & strOld & ")") Then
    strCode = Safe_Replace(strCode, " " & strOld & ")", " " & Strnew & ")")
  End If
  If InStr(strCode, "(" & strOld) Then
    strCode = Safe_Replace(strCode, "(" & strOld, "(" & Strnew)
  End If
  'v2.9.0 deal with Function X(Y)As Integer()
  If InStr(strCode, strOld & "()") Then
    strCode = Safe_Replace(strCode, strOld & "()", Strnew & "()")
  End If
  If InStr(strOld, ") As ") > 0 Then
    If InStr(strCode, strOld) > 0 Then
      strCode = Safe_Replace(strCode, " " & strOld, " " & Strnew)
    End If
  End If
  If SmartRight(strCode, strOld) Then
    strCode = Safe_Replace(strCode, strOld, Strnew)
  End If

End Sub

Private Function RequiredIntegerType(ByVal strTest As String) As Boolean

  'v2.6.5 added safety
  ' test is case insensitive but if you use VB's API Viewer or AllAPI's APIViewer
  ' the strings will usually be all caps (a few exceptions)
  ' Thanks Aaron Spivey

  RequiredIntegerType = QSortArrayPos(ArrQIntTypes, UCase$(strTest)) > -1

End Function

Public Sub ReStructure_Engine()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long

  On Error GoTo BugHit
  'Most of the following tests could be incorperated into a single code sweeper
  'but separating them out makes code clearer if slower.
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          If Not bAborting Then
            With Comp
              If .CodeModule.CountOfLines Then
                WorkingMessage "Module Self Reference remover", 1, 21
                FormSelfReference .CodeModule
                WorkingMessage "Top Comment Move", 2, 21
                TopCommentsIntoRoutine .CodeModule
                WorkingMessage "Update String Concatenation", 2, 21
                UpDateStringConcatenationFast .CodeModule
                WorkingMessage "Update Variant to $ Functions", 3, 21
                UpDateStringFunctionsFast .CodeModule
                WorkingMessage "Update Chr to VB constant", 4, 21
                Chr2ConstantUpDateFast .CodeModule
                WorkingMessage "Update Err to Err Object", 5, 21
                ErrErrorUpDate .CodeModule
                WorkingMessage "Scope Procedures", 6, 21
                UpdateRoutineScope .CodeModule
                WorkingMessage "Line number removal", 7, 21
                LineNumberRemoval .CodeModule
                WorkingMessage "Integer to Long", 8, 21
                Integer2Long .CodeModule
                'v2.6.7 double up to cope with 'Again:    If i < (bx * per) Then'
                WorkingMessage "GoTo Expander", 9, 21
                GotoLabelSeparator .CodeModule
                WorkingMessage "If..Then.. Expander", 10, 21
                DoIfThenStructureExpanderFast .CodeModule
                'DoIfThenStructureExpander .CodeModule
                WorkingMessage "Compound Lines Expander", 11, 21
                SeparateCompoundLinesFast .CodeModule
                'SeparateCompoundLines .CodeModule
                WorkingMessage "If..Then.. Expander 2", 12, 21
                DoIfThenStructureExpanderFast .CodeModule
                'DoIfThenStructureExpander .CodeModule
                WorkingMessage "Load/UnLoad without Resume Error", 13, 21
                LoadUnLoadErrorResumeFast .CodeModule
                'LoadUnLoadErrorResume .CodeModule
                WorkingMessage "Error Resume Close", 14, 21
                OnErrorResumeCloseFast .CodeModule
                'OnErrorResumeClose .CodeModule
                WorkingMessage "Purify With Structures", 15, 21
                WithStructurePurityFast .CodeModule
                'WithStructurePurity .CodeModule
                WorkingMessage "String Length testing", 16, 21
                ZeroStringFixFast .CodeModule
                'ZeroStringFix .CodeModule
                WorkingMessage "Suggest With Structures", 17, 21
                bWithSuggested = False ' if nothing suggested the skip the apply stage
                SuggestWithStructure .CodeModule
                WorkingMessage "Public Variable to Property", 18, 21
                CreatePropertyFromPublicVariableInClass .CodeModule, Comp
                If FixData(DetectWithStructure).FixLevel > CommentOnly And bWithSuggested Then
                  WorkingMessage "Apply With Structures", 19, 21
                  SuggestWithStructure .CodeModule
                  WorkingMessage "Clear False With suggestions", 20, 21
                  StripErrantSuggestwithsFast .CodeModule
                  'StripErrantSuggestwiths .CodeModule
                End If
                WorkingMessage "Routine Case Fix", 21, 21
                CaseOfRoutineFix .CodeModule
              End If
            End With 'Comp
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For
      End If
    Next Proj
    RestructureDriver
  End If

Exit Sub

BugHit:
  BugTrapComment "Restructure_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Sub ReStructure_EngineFormat()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long

  On Error GoTo BugTrap
  'MOst of the following tests could be incorperated into a single code sweeper
  'but separating them out makes code clearer if slower.
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          If bAborting Then
            Exit For
          End If
          With Comp
            WorkingMessage "Update String Concatenation", 1, 16
            'UpDateStringConcatenation .CodeModule
            UpDateStringConcatenationFast .CodeModule
            WorkingMessage "Update Variant to $ Functions", 2, 16
            'UpDateStringFunctions .CodeModule
            UpDateStringFunctionsFast .CodeModule
            WorkingMessage "Update Chr to VB constant", 3, 16
            'Chr2ConstantUpDate .CodeModule
            Chr2ConstantUpDateFast .CodeModule
            WorkingMessage "Update Err to Err Object", 4, 16
            ErrErrorUpDate .CodeModule
            WorkingMessage "Line number removal", 5, 16
            LineNumberRemoval .CodeModule
            'v2.6.7 double up to cope with 'Again:    If i < (bx * per) Then'
            WorkingMessage "GoTo Expander", 6, 16
            GotoLabelSeparator .CodeModule
            WorkingMessage "If..Then.. Expander", 7, 16
            DoIfThenStructureExpanderFast .CodeModule
            'DoIfThenStructureExpander .CodeModule
            WorkingMessage "Compound Lines Expander", 8, 16
            SeparateCompoundLinesFast .CodeModule
            'SeparateCompoundLines .CodeModule
            WorkingMessage "If..Then.. Expander 2", 9, 16
            DoIfThenStructureExpanderFast .CodeModule
            'DoIfThenStructureExpander .CodeModule
            WorkingMessage "Load/UnLoad without Resume Error", 10, 16
            LoadUnLoadErrorResumeFast .CodeModule
            'LoadUnLoadErrorResume .CodeModule
            WorkingMessage "Purify With Structures", 11, 16
            WithStructurePurityFast .CodeModule
            'WithStructurePurity .CodeModule
            WorkingMessage "String Length testing", 12, 16
            ZeroStringFixFast .CodeModule
            'ZeroStringFix .CodeModule
            WorkingMessage "Suggest With Structures", 13, 16
            bWithSuggested = False ' if nothing suggested the skip the apply stage
            SuggestWithStructure .CodeModule
            If FixData(DetectWithStructure).FixLevel > CommentOnly And bWithSuggested Then
              WorkingMessage "Apply With Structures", 14, 16
              SuggestWithStructure .CodeModule
              WorkingMessage "Clear False With suggestions", 15, 16
              StripErrantSuggestwithsFast .CodeModule
              'StripErrantSuggestwiths .CodeModule
            End If
            WorkingMessage "Routine Case Fix", 16, 16
            CaseOfRoutineFix .CodeModule
          End With 'Comp
        End If
      Next Comp
      If bAborting Then
        Exit For
      End If
    Next Proj
    'these need to be here to occur after seperate compound lines
    RestructureDriver
  End If

Exit Sub

BugTrap:
  BugTrapComment "ReStructure_EngineFormat"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub RestructureDriver()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long
  Dim NumFixes     As Long

  NumFixes = 14
  'v2.5.0 simplified handler for search based fixes
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        ModuleMessage Comp, CurCompCount
        If dofix(CurCompCount, DetectShortCircuit) Then
          WorkingMessage "If..And..Then Short Circuit", 1, NumFixes
          IfAndThenShortCircuitEngine Comp.CodeModule
        End If
        If dofix(CurCompCount, UnNeededIfBracket) Then
          WorkingMessage "Unneeded Bracket removal", 2, NumFixes
          DoIFBracketEngine Comp.CodeModule
        End If
        If dofix(CurCompCount, NPleonasmFix) Then
          WorkingMessage "Pleonasm (False) removal", 3, NumFixes
          PleonasmCleanerEngine Comp.CodeModule, False
          WorkingMessage "Pleonasm removal", 4, NumFixes
          PleonasmCleanerEngine Comp.CodeModule, True
        End If
        If dofix(CurCompCount, CollapseIfThenBool) Then
          WorkingMessage "Collapse If Then Boolean settings 1", 5, NumFixes
          IfThen2BooleanEngine Comp.CodeModule, 0
          WorkingMessage "Collapse If Then Boolean settings 2", 6, NumFixes
          IfThen2BooleanEngine Comp.CodeModule, 1
        End If
        If dofix(CurCompCount, UneededCaseIs) Then
          WorkingMessage "Unneeded 'Case Is = ", 7, NumFixes
          CaseIsEngine Comp.CodeModule
        End If
        If dofix(CurCompCount, CollapseCaseBool) Then
          WorkingMessage "Collapse If Then Boolean settings 1", 8, NumFixes
          CaseIIfEngine Comp.CodeModule, 0
          WorkingMessage "Collapse If Then Boolean settings 2", 9, NumFixes
          CaseIIfEngine Comp.CodeModule, 1
        End If
        'v2.7.3 reordered to cut down on messages
        If dofix(CurCompCount, UnNeededExit) Then
          WorkingMessage "Unneeded Exits", 10, NumFixes
          UnNecessaryExitEngine Comp.CodeModule
        End If
        If dofix(CurCompCount, SecondryExitFix) Then
          WorkingMessage "Secondry Exits test", 11, NumFixes
          If FixData(SecondryExitFix).FixLevel = FixAndComment Then
            SecondryExitsHard Comp.CodeModule
          End If
          SecondryExits Comp.CodeModule
        End If
        If dofix(CurCompCount, UpdateWend) Then
          WorkingMessage "'While/Wend' to 'Do While/Loop'", 12, NumFixes
          WendUpdateEngine Comp.CodeModule
        End If
        If dofix(CurCompCount, UnNeededAmpersand) Then
          'v2.8.6 Thanks to Bazz
          WorkingMessage "UnNeeded Ampersand String jointfix", 13, NumFixes
          StringLiteralBlend Comp.CodeModule
        End If
        If dofix(CurCompCount, UnNeededCall) Then
          WorkingMessage "Unneeded 'Call' remover", 14, NumFixes
          CallRemovalEngine Comp.CodeModule
        End If
      End If
    Next Comp
  Next Proj

End Sub

Private Sub SecondryExits(cMod As CodeModule)

  Dim strFixHeader  As String
  Dim L_CodeLine    As String
  Dim ArrMember     As Variant
  Dim ArrRoutine    As Variant
  Dim Member        As Long
  Dim RLine         As Long
  Dim UpDateRoutine As Boolean
  Dim MUpdated      As Boolean
  Dim MemberCount   As Long

  'v2.7.1 new fix
  strFixHeader = SUGGESTION_MSG & "(EXPERIMENTAL follow advice with care )" & vbNewLine & _
   "Explict 'Exit ProcedureType' can make code flow harder to follow."
  ArrMember = GetMembersArray(cMod)
  MemberCount = UBound(ArrMember)
  If MemberCount > 0 Then
    For Member = 1 To MemberCount
      If InStr(ArrMember(Member), "Exit ") Then
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        'If UBound(ArrRoutine) > -1 Then
        '  TopOfRoutine = GetProcCodeLineOfRoutine(ArrRoutine, True)
        'End If
        For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
          L_CodeLine = ArrRoutine(RLine)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If MultiLeft(L_CodeLine, True, "Exit Sub", "Exit Function", "Exit Property") Then
              Select Case ExplicitExitType(ArrRoutine, RLine, cMod)
               Case eExitIf_Then_Exit_EndIf_ID1
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 1)" & vbNewLine & _
                 "Convert 'If..Then/Exit/End If'  to " & vbNewLine & _
                 "'If Not .. Then/Rest_Of_Code/End If'", MAfter, UpDateRoutine)
               Case eExitIf_Then_Exit_EndIf_Code_ID2
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & " (Fix ID 2)" & vbNewLine & _
                 "Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to" & vbNewLine & _
                 "'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If" & vbNewLine & _
                 "OR if Exit Code block is only the Exit Command" & vbNewLine & _
                 "'If Not ..Then/ Rest_Of_Code/End If", MAfter, UpDateRoutine)
               Case eExitAs1stCode_ID1a
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 1a Auto fixable)" & vbNewLine & _
                 "Reverse the test (EG 'X = Y' -> 'X <> Y', If <Boolean> Then -> If Not <Boolean> Then') " & vbNewLine & _
                 "Delete 'Exit <ProcedureType>' line" & vbNewLine & _
                 "Move the 'End If' to just above the 'End <ProcedureType>' Or 'Exit <ProcedureType>' if used for ErrorTrap protecting.", MAfter, UpDateRoutine)
               Case eExitAs1stCodeComplex_ID2a
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 2a Auto fixable)" & vbNewLine & _
                 "Replace 'Exit <ProcedureType>' with 'Else'" & vbNewLine & _
                 "Move the 'End If' to just above the 'End <ProcedureType>' Or 'Exit <ProcedureType>' if used for ErrorTrap protecting.", MAfter, UpDateRoutine)
               Case eExitIf_Then_Exit_Else_Code_EndIf_ID3
                '3 'If Then/Exit CODE/Else/Rest+Of_Code/End If
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 3)" & vbNewLine & _
                 "Convert 'If..Then/Code(with Explicit Exit)/Else/Rest_of_Code/End If' to" & vbNewLine & _
                 "'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If" & vbNewLine & _
                 "OR if Exit Code block is only the Exit Command" & vbNewLine & _
                 "'If Not ..Then/ Rest_Of_Code/End If", MAfter, UpDateRoutine)
                'Case eExitErrorTrapShield ' 4 ' protects ErrorTrap so no comment
                'ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader &"(Fix ID 4)"& vbNewLine & "ErrorTrap protector", MAfter, UpDateRoutine)
               Case eExitDeepStructure_ID5
                ' too deep in structures to be easily rewritten
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 5)" & vbNewLine & _
                 "too deep in structures to be easily rewritten" & vbNewLine & _
                 "but a 'GoTo' to end of procedure is a possibility.", MAfter, UpDateRoutine)
               Case eExitProc2ExitLoop_ID8a
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 8a Auto fixable)" & vbNewLine & _
                 "You may be able to replace 'Exit <ProcedureType> with 'Exit <For|Do>'", MAfter, UpDateRoutine)
               Case eExitFromWith_ID9
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 9)" & vbNewLine & _
                 "Exiting from within a With structure is a potential memory leak." & vbNewLine & _
                 "Use a GoTo to reach the 'End With' line AND a Boolean variable to reach a safe Exit point.", MAfter, UpDateRoutine)
               Case eExitFromFors_ID10
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 10)" & vbNewLine & _
                 "Exiting from nested For structures can be done by " & vbNewLine & _
                 "using 'Exit For' and a Boolean variable to test if outer For's should exit.", MAfter, UpDateRoutine)
               Case eExitGEneric_ID11
                ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID 11)" & vbNewLine & _
                 "No recommended action but consider coding around it.", MAfter, UpDateRoutine)
                'Case Else
                '''ArrRoutine(RLine) = Marker(ArrRoutine(RLine), strFixHeader & "(Fix ID ?)", MAfter, UpDateRoutine)
              End Select
            End If
          End If
        Next RLine
        UpdateMember ArrMember(Member), ArrRoutine, UpDateRoutine, MUpdated
      End If
    Next Member
    ReWriteMembers cMod, ArrMember, MUpdated
  End If
  On Error GoTo 0

End Sub

Private Sub SecondryExitsHard(cMod As CodeModule)

  Dim UpDateRoutine As Boolean
  Dim MUpdated      As Boolean
  Dim MemberCount   As Long
  Dim ExitLineType  As Long
  Dim Member        As Long
  Dim RLine         As Long
  Dim L_CodeLine    As String
  Dim ArrMember     As Variant
  Dim ArrRoutine    As Variant

  'v2.8.5 new fix
  ArrMember = GetMembersArray(cMod)
  MemberCount = UBound(ArrMember)
  If MemberCount > 0 Then
    For Member = 1 To MemberCount
      If InStr(ArrMember(Member), "Exit ") Then
        ArrRoutine = Split(ArrMember(Member), vbNewLine)
        MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
        For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
          L_CodeLine = ArrRoutine(RLine)
          If Not JustACommentOrBlank(L_CodeLine) Then
            'v2.9.9 faster test
            If InstrAtPositionSetArray(L_CodeLine, IpLeft, True, ArrExitFuncPropSub) Then
              'v2.9.9 allows system to recover avoid task manager avoid not responding message
              Safe_Sleep
              ExitLineType = ExplicitExitType(ArrRoutine, RLine, cMod)
              Select Case ExitLineType
               Case eExitAs1stCodeComplex_ID2a, eExitAs1stCode_ID1a ', eExitIf_Then_Exit_EndIf_Code
                ExitProc2ExtendedIfThen ArrRoutine, RLine, ExitLineType, UpDateRoutine
               Case eExitGEneric_ID11
                ExitProc2ExitFor ArrRoutine, RLine, UpDateRoutine
              End Select
            End If
          End If
        Next RLine
        UpdateMember ArrMember(Member), ArrRoutine, UpDateRoutine, MUpdated
      End If
    Next Member
    ReWriteMembers cMod, ArrMember, MUpdated
  End If
  On Error GoTo 0

End Sub

Private Sub SeparateCompoundLinesFast(cMod As CodeModule)

  'v3.0.7 Fast version
  
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim ArrProc      As Variant
  Dim L_CodeLine   As String
  Dim RLine        As Long
  Dim Sline        As Long
  Dim arrGoTo      As Variant
  Dim UpDated      As Boolean
  Dim ModuleNumber As Long

  With cMod
    ModuleNumber = ModDescMember(cMod.Parent.Name)
    If dofix(ModuleNumber, SeperateCompounds) Then
      Do While .Find(":", Sline, 1, -1, -1)
        If InCode(.Lines(Sline, 1), InStr(.Lines(Sline, 1), ":")) Then 'ignore GotoLabels
          If Not ((InStr(.Lines(Sline, 1), " ") = 0) And Right$(strCodeOnly(.Lines(Sline, 1)), 1) = Colon) Then
            ArrProc = ReadProcedureCodeArray2(cMod, Sline, PSLine, PEndLine)
            MemberMessage GetProcName(cMod, Sline), Sline, .CountOfLines
            arrGoTo = GetGoToTargetArray(ArrProc)
            For RLine = LBound(ArrProc) To UBound(ArrProc)
              L_CodeLine = ArrProc(RLine)
              If Not JustACommentOrBlank(L_CodeLine) Then
                'v2.8.3 improved test for less hits
                If InStr(L_CodeLine, ":") Then
                  DoSeparateCompoundLines ModuleNumber, L_CodeLine, cMod.Parent.Name, arrGoTo
                  If ArrProc(RLine) <> L_CodeLine Then
                    AddNfix SeperateCompounds
                    ArrProc(RLine) = L_CodeLine
                    UpDated = True
                  End If
                End If
              End If
            Next RLine
            If UpDated Then
              ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine, False
            End If
          End If
        End If
        Sline = Sline + IIf(PEndLine - PSLine > 1, PEndLine - PSLine, 1)
        PEndLine = 0
        PSLine = 0
        If Sline > cMod.CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With
  '

End Sub

Private Function SetVarForIfTest(varCode As Variant, _
                                 VarIfLine As Variant) As Boolean

  Dim strTestVariable As String

  If WordInString(varCode, 2) = "=" Then
    strTestVariable = LeftWord(varCode)
    SetVarForIfTest = InStrWholeWord(VarIfLine, strTestVariable)
  End If

End Function

Private Sub StringConcatenationUpDate(strWork As String)

  Dim MyStr        As String
  Dim CommentStore As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'fixes old style string concatenation '+' to use  safer '&'
  'NOTE This routine copes with line continuation if the " is before the line cont character
  'but NOT if form is:
  '
  '               non-String-Variable + '               "String Literal"
  '
  On Error GoTo BadError
  If InStr(strWork, StrPlus) Then
    If InstrAtPositionArray(strWork, ipAny, False, DQuote & StrPlus, StrPlus & DQuote, StrPlus & " vbNewLine ", " vbNewLine " & StrPlus, StrPlus & " vbTab ", " vbTab " & StrPlus) Then
      MyStr = strWork
      If ExtractCode(MyStr, CommentStore) Then
        DisguiseLiteral MyStr, StrPlus, True
        MyStr = ReplaceArray(MyStr, DQuote & StrPlus, DQuote & " & ", StrPlus & DQuote, " & " & DQuote, StrPlus & " vbNewLine ", " & " & " vbNewLine ", " vbNewLine " & StrPlus, " vbNewLine " & " & ", StrPlus & " vbTab ", " & " & " vbTab ", " vbTab " & StrPlus, " vbTab " & " & ")
        DisguiseLiteral MyStr, StrPlus, False
        If strWork <> MyStr & CommentStore Then
          strWork = MyStr & CommentStore
        End If
      End If
    End If
  End If
  On Error GoTo 0
BadError:

End Sub

Private Sub StringLiteralBlend(cMod As CodeModule)

  Dim strTest2  As String
  Dim StartLine As Long

  'v2.8.6 Thanks to Bazz for sending me code that suggested this fix being needed
  ' so
  'Open .PluginPath & .PluginName & "\files" & "\" & "ramdisk_" & .PluginName & ".cmd" For Output As #hFile
  'becomes
  'Open .PluginPath & .PluginName & "\files\ramdisk_" & .PluginName & ".cmd" For Output As #hFile
  '
  'v2.4.1 NEW version of old fix
  With cMod
    Do While .Find(DQuote & " & " & DQuote, StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        If Not JustACommentOrBlank(.Lines(StartLine, 1)) Then
          '  because you don't want to check the quotemark check the space just after it
          If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), DQuote & " & " & DQuote) + 1) Then
            If Len(.Lines(StartLine, 1)) < lngLineLength * 1.1 Then
              ' only do it to short lines you want longer ones to wrap around in later formatting
              ' the 1.1 increase in length is to cope with the extra length of the stuff targerted for removal
              If Not HasLineCont(.Lines(StartLine, 1)) Then
                strTest2 = Replace$(.Lines(StartLine, 1), DQuote & " & " & DQuote, vbNullString)
                If strTest2 <> .Lines(StartLine, 1) Then
                  Select Case FixData(UnNeededCall).FixLevel
                   Case FixAndComment
                    .ReplaceLine StartLine, strTest2 & vbNewLine & _
                     WARNING_MSG & "Unneeded String Literals joined by ampersand removed." & vbNewLine & _
                     PREVIOUSCODE_MSG & .Lines(StartLine, 1)
                  End Select
                End If
              End If
            End If
          End If
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Sub StripErrantSuggestwithsFast(cMod As CodeModule)

  Dim Sline   As Long

  With cMod
    If dofix(ModDescMember(.Parent.Name), DetectWithStructure) Then
      Do While .Find("'APPROVED(Y)", Sline, 1, -1, -1, True, True)
        If Left$(.Lines(Sline, 1), 12) = "'APPROVED(Y)" Then
          .DeleteLines Sline
        End If
        Sline = Sline + 1
        If Sline > .CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

End Sub

Private Function StructureDeep2(ByVal CurDepth As Long, _
                                varSearch As Variant, _
                                StrDeep As String) As Long

  Dim strTest As String

  strTest = LeftWord(varSearch)
  Select Case strTest
   Case "Select", "For", "Do", "While", "With", "If"
    StrDeep = AccumulatorString(StrDeep, strTest, , False)
    StructureDeep2 = CurDepth + 1
   Case "Next", "Loop", "Wend"
    Select Case strTest
     Case "Next"
      StrDeep = Replace$(StrDeep, "For", vbNullString, , 1)
     Case "Loop"
      StrDeep = Replace$(StrDeep, "Do", vbNullString, , 1)
     Case "Wend"
      StrDeep = Replace$(StrDeep, "While", vbNullString, , 1)
    End Select
    StructureDeep2 = CurDepth - 1
   Case "End"
    If MultiLeft(varSearch, True, "End If", "End With", "End Select") Then
      StrDeep = Replace$(StrDeep, WordInString(varSearch, 2), vbNullString, , 1)
      StructureDeep2 = CurDepth - 1
     Else
      StructureDeep2 = CurDepth 'do nothing
    End If
   Case Else
    StructureDeep2 = CurDepth
  End Select

End Function

Public Function TypeOfIf2(cMod As CodeModule, _
                          ByVal LineNo As Long, _
                          EndPos As Long) As IFType

  Dim I       As Long
  Dim Level   As Long
  Dim strTest As String
  Dim EndLine As Long

  EndLine = GetProcEndLine(cMod, LineNo)
  For I = LineNo To EndLine
    strTest = cMod.Lines(I, 1)
    If InstrAtPosition(strTest, "If", IpLeft, True) Then
      'ver1.1.29 this stops single line If..Then....'s from being counted
      If InStrCode(strTest, "Then") Then
        Level = Level + 1
      End If
    End If
    'v2.8.3 improved test allows If within If's   Thanks Joakim Schramm
    ' looks for Else as well as End If
    If InstrAtPosition(strTest, "End If", IpLeft) Then
      Level = Level - 1
    End If
    If Level = 0 Then
      TypeOfIf2 = Simple
      EndPos = I
      Exit For
    End If
    If Level = 1 Then
      If InstrAtPosition(strTest, "Else", IpLeft, True) Then
        TypeOfIf2 = Complex1
        EndPos = -1
        Exit For
      End If
      If InstrAtPosition(strTest, "ElseIf", IpLeft, True) Then
        TypeOfIf2 = Complex2
        EndPos = -1
        Exit For
      End If
    End If
  Next I

End Function

Private Function TypeUsedByAPI(ByVal strPar As String) As Boolean

  Dim I As Long

  If bProcDescExists Then
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      If PRocDesc(I).PrDUDTParam Then
        If PRocDesc(I).PrDAPI Then
          If InStr(PRocDesc(I).PrDArgumentTypes, strPar) Then
            TypeUsedByAPI = True
            Exit For
          End If
        End If
      End If
    Next I
  End If

End Function

Private Function TypeUsedByAPI2(ByVal strPar As String) As Boolean

  Dim I As Long

  If bDeclExists Then
    For I = LBound(DeclarDesc) To UBound(DeclarDesc)
      If InStr(DeclarDesc(I).DDArguments, strPar) Then
        TypeUsedByAPI2 = True
        Exit For
      End If
    Next I
  End If

End Function

Private Sub UnNecessaryExitEngine(cMod As CodeModule)

  'v2.9.9 improved to deal with Next/Loop  and all 'End If's
  
  Dim StartLine As Long
  Dim Mode      As Long

  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find("Exit ", StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If InCode(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), "Exit ")) Then
        'ignore exit loop commands
        If InstrAtPositionSetArray(.Lines(StartLine, 1), IpLeft, True, ArrExitFuncPropSub) Then
          'v2.8.4 improved test for unneeded Exit Procedure
          If IsUnneededExit(cMod, StartLine, Mode) Then
            'v2.9.9 added extra exits for Next and Loop
            Select Case Mode
             Case 0
              InsertNewCodeComment cMod, StartLine, 1, WARNING_MSG & "Unneeded " & strCodeOnly(.Lines(StartLine, 1)), ""
             Case 1
              InsertNewCodeComment cMod, StartLine, 1, "Exit Do " & vbNewLine & _
               WARNING_MSG & "Unneeded " & strCodeOnly(.Lines(StartLine, 1)) & " replaced with 'Exit Do'", ""
             Case 2
              InsertNewCodeComment cMod, StartLine, 1, "Exit For " & vbNewLine & _
               WARNING_MSG & "Unneeded " & strCodeOnly(.Lines(StartLine, 1)) & " replaced with 'Exit For'", ""
            End Select
          End If
        End If
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Function UpdatableInteger(ByVal strCode As String, _
                                  cMod As VBComponent, _
                                  bisUdProc As Boolean, _
                                  TestPos As Long) As Boolean

  'ver 2.1.3 centralised all the test in one

  strCode = strCodeOnly(strCode)
  If IsRealWord(strCode, "Integer") Then
    TestPos = InStr(strCode, " Integer")
    If TestPos Then
      If NotVBControlCode(strCode, cMod.CodeModule, bisUdProc) Then
        ' ver 1.1.93 thanks to Morgan Haueisen for noticing this bug
        'A function called Integer_Pos would trigger this fix
        UpdatableInteger = InCode(strCode, TestPos)
        ' End If
      End If
    End If
  End If

End Function

Private Sub UpdateIntegerTypeSuffix(arrR As Variant, _
                                    strVariable As String)

  Dim K As Long

  For K = LBound(arrR) To UBound(arrR)
    arrR(K) = Safe_Replace(arrR(K), strVariable & "%", strVariable)
  Next K

End Sub

Private Sub UpdateRoutineScope(cMod As CodeModule)

  Dim L_CodeLine  As String
  Dim ArrMember   As Variant
  Dim RLine       As Long
  Dim Member      As Long
  Dim ArrRoutine  As Variant
  Dim UpDated     As Boolean
  Dim MUpdated    As Boolean
  Dim MemberCount As Long
  Dim lngdummy    As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'This deals with the case of routines not having an explicit Scope
  If dofix(ModDescMember(cMod.Parent.Name), DimGlobal2PublicPrivate) Then
    If LenB(GetActiveProject.FileName) Then
      ArrMember = GetMembersArray(cMod)
      MemberCount = UBound(ArrMember)
      If MemberCount > 0 Then
        For Member = 1 To MemberCount
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
          For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
            L_CodeLine = ArrRoutine(RLine)
            If Not JustACommentOrBlank(L_CodeLine) Then
              If InstrAtPositionSetArray(L_CodeLine, IpLeft, True, ArrFuncPropSub) Then
                L_CodeLine = ScopeTo(cMod, L_CodeLine, lngdummy)
                If L_CodeLine <> ArrRoutine(RLine) Then
                  ArrRoutine(RLine) = L_CodeLine
                  UpDated = True
                End If
              End If
            End If
          Next RLine
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        Next Member
        ReWriteMembers cMod, ArrMember, MUpdated
      End If
    End If
  End If

End Sub

Public Sub UpdateStringArray(strWork As String, _
                             ArrayOld As Variant, _
                             ArrayNew As Variant)

  Dim MyStr        As String
  Dim StrSafe      As String
  Dim I            As Long
  Dim CommentStore As String

  'General service for updating members of ArrayOld with equivalent member of ArrayNew
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  StrSafe = strWork
  On Error GoTo BadError
  For I = LBound(ArrayOld) To UBound(ArrayOld)
    If InstrAtPosition(strWork, ArrayOld(I), ipAny, False) Then
      MyStr = strWork
      If ExtractCode(MyStr, CommentStore) Then
        DisguiseLiteral MyStr, ArrayOld(I), True
        MyStr = Safe_Replace(MyStr, ArrayOld(I), ArrayNew(I), , , False, True)
        DisguiseLiteral MyStr, ArrayOld(I), False
        'v2.4.4 replaceing vbCrLf with vbNewLine could drive the string over length
        If strWork <> MyStr & CommentStore Then
          If Len(MyStr & CommentStore) < 1023 Then
            strWork = MyStr & CommentStore
          End If
        End If
      End If
    End If
  Next I
  On Error GoTo 0

Exit Sub

BadError:
  strWork = StrSafe

End Sub

Private Sub UpDateStringConcatenationFast(cMod As CodeModule)

  'v3.0.4 new faster version
  
  Dim StartLine  As Long
  Dim StrNewCode As String

  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find(" + ", StartLine, 1, -1, -1)
      StrNewCode = .Lines(StartLine, 1)
      If InCode(StrNewCode, InStr(StrNewCode, " + ")) Then
        StringConcatenationUpDate StrNewCode
        If StrNewCode <> .Lines(StartLine, 1) Then
          InsertNewCodeComment cMod, StartLine, 1, StrNewCode, vbNullString, Xcheck(XPrevCom)
        End If
      End If
Skip:
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop
  End With

End Sub

Private Sub UpDateStringFunctionsFast(cMod As CodeModule)

  'v3.0.4 new faster version
  
  Dim StartLine  As Long
  Dim I          As Long
  Dim StrNewCode As String

  With cMod
    StartLine = .CountOfDeclarationLines
    For I = LBound(ArrQStrVarFunc) To UBound(ArrQStrVarFunc)
      StartLine = 0
      Do While .Find(ArrQStrVarFunc(I) & "(", StartLine, 1, -1, -1)
        StrNewCode = .Lines(StartLine, 1)
        If InCode(StrNewCode, InStr(StrNewCode, ArrQStrVarFunc(I) & "(")) Then
          DoStringFunctionsCorrect StrNewCode
          If StrNewCode <> .Lines(StartLine, 1) Then
            InsertNewCodeComment cMod, StartLine, 1, StrNewCode, vbNullString, Xcheck(XPrevCom)
          End If
        End If
Skip:
        If StartLine < .CountOfDeclarationLines Then
          Exit Do
        End If
        StartLine = StartLine + 1
        If StartLine > .CountOfLines Then
          Exit Do
        End If
      Loop
    Next I
  End With

End Sub

Private Sub WendUpdateEngine(cMod As CodeModule)

  Dim StartLine As Long
  Dim RLine     As Long

  'v2.4.3 NEW version of old fix
  With cMod
    Do While .Find("Wend", StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        For RLine = StartLine To 1 Step -1
          If Left$(Trim$(.Lines(RLine, 1)), 6) = "While " Then
            Select Case FixData(UpdateWend).FixLevel
             Case CommentOnly
              '.ReplaceLine StartLine, .Lines(StartLine, 1)
              .InsertLines GetInsertionPoint(cMod, RLine), SUGGESTION_MSG & "Use Do...Loop it has 'While', 'Until' and' Exit' Methods which make it clearer and easier to code."
             Case FixAndComment, JustFix
              .ReplaceLine RLine, "Do " & .Lines(RLine, 1)
              .ReplaceLine StartLine, Replace$(.Lines(StartLine, 1), "Wend", "Loop", , 1)
              .InsertLines GetInsertionPoint(cMod, RLine), IIf(FixData(UpdateWend).FixLevel = FixAndComment, vbNewLine & _
               UPDATED_MSG & "'While ..Wend' to 'Do While ..Loop'.", "")
            End Select
            Exit For
          End If
        Next RLine
      End If
      If StartLine < .CountOfDeclarationLines Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > .CountOfLines Then
        Exit Do
      End If
    Loop '
  End With

End Sub

Private Function WrongImplementsCase(strTest As String, _
                                     strC As String) As Boolean

  Dim I As Long

  If Not IsEmpty(ImplementsArray) Then
    For I = 0 To UBound(ImplementsArray)
      If LCase$(ImplementsArray(I)) = LCase$(strTest) Then
        WrongImplementsCase = ImplementsArray(I) <> strTest
        strC = ImplementsArray(I)
        Exit For
      End If
    Next I
  End If

End Function

Private Sub ZeroStringFixFast(cMod As CodeModule)

  'v3.0.7 new fast version of fix
  'updated ver 1.0.95
  'added the vbNullString substitute for emptystring. Thanks to Rudz for suggesting it
  
  Dim ModuleNumber As Long
  Dim L_CodeLine   As String
  Dim Sline        As Long
  Dim UpDated      As Boolean
  Dim CommentStore As String
  Dim OldCode      As String
  Dim FixDone      As Boolean
  Dim ImpossCode   As Boolean

  With cMod
    ModuleNumber = ModDescMember(.Parent.Name)
    If dofix(ModuleNumber, NCloseResumeLoad) Then
      Do While .Find(EmptyString, Sline, 1, -1, -1, True, True)
        L_CodeLine = .Lines(Sline, 1)
        If ExtractCode(L_CodeLine, CommentStore) Then
          OldCode = .Lines(Sline, 1)
          Select Case EmptyStringType(L_CodeLine)
            'Case 0 ' do nothing
           Case 1 'Use Length Testing
            Select Case FixData(DetectZeroLengthStringTests).FixLevel
             Case CommentOnly
              L_CodeLine = Marker(OldCode, SUGGESTION_MSG & "Using 'Not LenB(X) ', 'LenB(X) = 0' or 'LenB(X) ' is more effiecent than 'If X = " & EmptyString & " Then | If X <> " & EmptyString & " Then'" & RGSignature & "Be careful to test Len value if there is an 'And' in code line." & vbNewLine & _
                                                                                                                                              "'If LenB(X) And LenB(Y) Then' is a Logical comparison AND NOT the same as 'If LenB(X) > 0 and LenB(Y) > 0 Then'", MAfter, UpDated)
             Case FixAndComment
              'testing in next section compacts paramateres and stuffs up if parameters include "somestring" & "someother string"
              EmptyStringComparisonFix L_CodeLine, FixDone, ImpossCode
              If Len(CommentStore) Then
                If InStr(L_CodeLine, CommentStore) = 0 Then
                  L_CodeLine = L_CodeLine & CommentStore
                End If
              End If
              If FixDone Then
                L_CodeLine = Marker(L_CodeLine, WARNING_MSG & "Empty String comparison updated to use LenB() " & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & OldCode, vbNullString), MAfter)
                FixDone = False
               Else
                If Not ImpossCode Then
                  L_CodeLine = Marker(OldCode, SUGGESTION_MSG & "Using 'Not LenB(X)', 'LenB(X) = 0' or 'Len(X)' is more effiecent than 'If X = " & EmptyString & " Then | If X <> " & EmptyString & " Then'" & vbNewLine & _
                                                                                                                                               "Be careful to test Len value if there is an 'And' in code line." & vbNewLine & _
                                                                                                                                               "'If LenB(X) And LenB(Y) Then' is a Logical comparison AND NOT the same as 'If LenB(X) > 0 and LenB(Y) > 0 Then'", MAfter, UpDated)
                 Else
                  ImpossCode = False ' turn it off for next sweep
                  L_CodeLine = Marker(OldCode, WARNING_MSG & "Left/Mid/Right should not be applied to empty strings, the answer is always False" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & OldCode, vbNullString), MAfter, UpDated)
                End If
              End If
              'Case JustFix
            End Select
            UpDated = True
            AddNfix DetectZeroLengthStringTests
           Case 2, 4 'Assign vbNullString
            Select Case FixData(DetectZeroLengthStringAssign).FixLevel
             Case CommentOnly
              If Len(CommentStore) Then
                If InStr(L_CodeLine, CommentStore) = 0 Then
                  L_CodeLine = L_CodeLine & CommentStore
                End If
              End If
              L_CodeLine = Marker(L_CodeLine, SUGGESTION_MSG & "Using ' = vbNullString' is more effiecent than ' = " & EmptyString & " '", MAfter, UpDated)
             Case FixAndComment
              L_CodeLine = Replace$(L_CodeLine, EmptyString, "vbNullString")
              If OldCode <> L_CodeLine Then
                If Len(CommentStore) Then
                  If InStr(L_CodeLine, CommentStore) = 0 Then
                    L_CodeLine = L_CodeLine & CommentStore
                  End If
                End If
                'arrproc(RLine) = Marker(arrproc(RLine), WARNING_MSG & "Empty String assignment updated to use vbNullString" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & OldCode, vbNullString), MAfter, UpDated)
                L_CodeLine = Marker(L_CodeLine, WARNING_MSG & "Empty String assignment updated to use vbNullString" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & OldCode, vbNullString), MAfter)
                UpDated = True
              End If
              'Case JustFix
            End Select
            UpDated = True
            AddNfix DetectZeroLengthStringTests
          End Select
        End If
        If UpDated Then
          .ReplaceLine Sline, L_CodeLine
          Sline = Sline + CountSubString(L_CodeLine, vbNewLine)
          UpDated = False
         Else
          Sline = Sline + 1
        End If
        If Sline > .CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

End Sub

Private Sub ZeroStringForcedTrue(varTest As Variant, _
                                 DoneIt As Boolean, _
                                 ImpossCode As Boolean)

  Dim I As Long

  'Dealwith code of the format
  '*If (strVariable <> "") Then'
  'OR
  '*If (strVariable = "") Then'
  If Left$(varTest, 4) = "If (" Then
    If Right$(varTest, 13) = "<> """") Then" Then
      If InStr(varTest, "(") = 1 Then
        varTest = "If LenB" & Mid$(varTest, 4)
        varTest = Left$(varTest, Len(varTest) - 11) & ") Then"
        DoneIt = True
        ImpossCode = False
       Else
        DoneIt = False
        'ImpossCode = True
      End If
     ElseIf Right$(varTest, 12) = "= """") Then" Then
      'v2.8.3 Improved to deal with multilpe contidions
      For I = InStr(varTest, "= """") Then") To 1 Step -1
        If Mid$(varTest, I, 1) = "(" Then
          varTest = Left$(varTest, I - 1) & "LenB" & Mid$(varTest, I)
          varTest = Left$(varTest, Len(varTest) - 11) & ") = 0 Then"
          DoneIt = True
          ImpossCode = False
          Exit For
        End If
      Next I
    End If
  End If
  If Left$(varTest, 8) = "ElseIf (" Then
    If Right$(varTest, 13) = "<> """") Then" Then
      varTest = "If LenB" & Mid$(varTest, 4)
      varTest = Left$(varTest, Len(varTest) - 11) & ") Then"
      DoneIt = True
      ImpossCode = False
     ElseIf Right$(varTest, 12) = "= """") Then" Then
      varTest = "If LenB" & Mid$(varTest, 4)
      varTest = Left$(varTest, Len(varTest) - 11) & ") = 0 Then"
      DoneIt = True
      ImpossCode = False
    End If
  End If

End Sub

'Private Sub OuterFastLoopFast(cMod As CodeModule)
'Dim PSLine        As Long
'Dim PEndLine      As Long
'Dim TopOfRoutine  As Long
'Dim Ptype         As Long
'Dim PName         As String
'Dim ProcKind      As String
'Dim ArrProc       As Variant
'Dim ModuleNumber  As Long
'Dim CurPos        As Long
'Dim L_CodeLine    As String
'Dim UpDateRoutine As Boolean
'Dim RLine         As Long
'Dim Sline         As Long
'  With cMod
'    ModuleNumber = ModDescMember(cMod.Parent.Name)
'    If dofix(ModuleNumber, XXXXXXXXXXXXXXXXXXXXXXX) Then
'      Do While .Find("XXXXXXXXX", Sline, 1, -1, -1, True, True)
'        If InCode(.Lines(Sline, 1), InStr(.Lines(Sline, 1), "XXXXXXXXXX")) Then
'          ArrProc = ReadProcedureCodeArray2(Cmod, Sline, PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind)
'          MemberMessage Rname, Sline, .CountOfLines
'          If PName <> "(Declarations)" Then
'
'
'
'
'
'            ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
'          End If
'        End If
'        Sline = Sline + IIf(UBound(ArrProc) > 1, UBound(ArrProc), 1)
'        If Sline > cMod.CountOfLines Then
'          Exit Do
'        End If
'      Loop
'    End If
'  End With
'
'End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:18:53 AM) 44 + 3926 = 3970 Lines Thanks Ulli for inspiration and lots of code.

