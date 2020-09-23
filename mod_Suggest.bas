Attribute VB_Name = "mod_Suggest"
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Option Explicit
Public Enum MarkerPos
  MBefore
  MAfter
  MEoL
  MSoL
  MEmbed
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private MBefore, MAfter, MEoL, MSoL, MEmbed
#End If
Private Enum StructViolation
  VNone
  VInto
  VOutof
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private VNone, VInto, VOutof
#End If
'*****************
Private Type GoToLocs
  Name                    As String
  LineNo                  As Long
  GoToTarget              As Boolean
  Orphen                  As Boolean
  Targets                 As String
  ErrorGoto               As Boolean
  Error0                  As Boolean
  IndentLevel             As Long
  IfLevel                 As Long
  RsumGo2                 As Boolean
  G0Sub2                  As Boolean
End Type
Public Const RISK_MSG     As String = RGSignature & ":RISK: "
Private GoToData()        As GoToLocs

Private Function AnyDangerousCodeTest(ByVal ModuleNumber As Long) As Boolean

  'ver 1.0.94
  'NEW
  'this allows the whole test cycle to be skipped if required

  If dofix(ModuleNumber, DetectDangerousCode_CXX) Then
    AnyDangerousCodeTest = True
   ElseIf dofix(ModuleNumber, DetectDangerousAPICode_CXX) Then
    AnyDangerousCodeTest = True
   ElseIf dofix(ModuleNumber, DetectDangerousReference_CXX) Then
    AnyDangerousCodeTest = True
   ElseIf dofix(ModuleNumber, DetectDangerousString_CXX) Then
    AnyDangerousCodeTest = True
   ElseIf dofix(ModuleNumber, DetectHardPath_CXX) Then
    AnyDangerousCodeTest = True
   ElseIf UsingScripting Then
    AnyDangerousCodeTest = True
   Else
    AnyDangerousCodeTest = False
  End If

End Function

Private Function AnyHardPaths(ByVal varTest As Variant, _
                              Optional ByVal bWrongSlash As Boolean) As Boolean

  Dim I    As Long
  Dim TPos As Long

  If InStr(varTest, ":\") Or InStr(varTest, ":/") Then
    '2nd is wrong but works in WinXP because the bad char / is deleted
    'also only works if target was root folder
    'don't do if basic path structure isn't present
    If ExtractCode(varTest) Then
      varTest = UCase$(varTest)
      For I = 65 To 90 'A to Z
        TPos = InStr(varTest, DQuote & Chr$(I) & ":\")
        If TPos = 0 Then
          TPos = InStr(varTest, DQuote & Chr$(I) & ":/")
          If TPos Then
            bWrongSlash = True
          End If
        End If
        If TPos Then
          If Not InComment(varTest, TPos) Then
            AnyHardPaths = True
            Exit For
          End If
        End If
      Next I
    End If
  End If

End Function

Private Sub BadControlName(cMod As CodeModule, _
                           Cmp As VBComponent)

  Dim strBadMsg    As String
  Dim strName      As String
  Dim strDoneGuard As String
  Dim strReason    As String
  Dim arrDec       As Variant
  Dim I            As Long

  If dofix(ModDescMember(cMod.Parent.Name), DetectBadCtrlName_CXX) Then
    If AnyBadNames Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        MemberMessage "", I, UBound(CntrlDesc)
        With CntrlDesc(I)
          If .CDBadType <> 0 Then
            strName = .CDName
            If .CDForm = Cmp.Name Then
              If InStr(strDoneGuard, strName) = 0 Then
                'stops whole arrays of controls getting a comment
                strDoneGuard = strDoneGuard & "****" & strName
                strReason = BadNameMsg(strName)
                strBadMsg = strBadMsg & SUGGESTION_MSG & "Poor Name:(" & strReason & ")" & strInSQuotes(strName, True) & .CDClass & IIf(.CDIndex <> -1, " Control Array. ", ".") & IIf(Xcheck(XVerbose), vbNewLine & _
                 "Legal but unclear and stops some Code Fixer actions working. Can make code hard to read.", vbNullString)
              End If
            End If
          End If
        End With 'CntrlDesc(I)
      Next I
      If Len(strBadMsg) Then
        arrDec = GetDeclarationArray(cMod)
        If UBound(arrDec) = -1 Then
          ReDim arrDec(0) As Variant
        End If
        arrDec(0) = Marker(arrDec(0), strBadMsg, MBefore)
        ReWriter cMod, arrDec, RWDeclaration
      End If
    End If
  End If

End Sub

Private Function CountArrayCodeLines(arrC As Variant) As Long

  Dim I As Long

  For I = LBound(arrC) To UBound(arrC)
    If Not JustACommentOrBlank(arrC(I)) Then
      CountArrayCodeLines = CountArrayCodeLines + 1
    End If
  Next I

End Function

'Private Sub StaticToPrivate(cMod As CodeModule)
'
'  Dim L_CodeLine   As String
'  Dim arrMembers   As Variant
'  Dim ArrRoutine   As Variant
'  Dim Member       As Long
'  Dim RLine        As Long
'  Dim UpDated      As Boolean
'  Dim MUpdated     As Boolean
'  Dim MaxFactor    As Long
'  Dim ModuleNumber As Long
'
'  'Copyright 2003 Roger Gilchrist
'  'e-mail: rojagilkrist@hotmail.com
'  ModuleNumber = ModDescMember(cMod.Parent.Name)
'  If dofix(ModuleNumber, DetectStatic2Private_CXX) Then
'    arrMembers = GetMembersArray(cMod)
'    MaxFactor = UBound(arrMembers)
'    If MaxFactor > 0 Then
'      For Member = 0 To MaxFactor
'        If InStr(arrMembers(Member), "Static ") Then
'          MemberMessage GetProcNameStr(arrMembers(Member)), Member, MaxFactor
'          ArrRoutine = Split(arrMembers(Member), vbNewLine)
'          For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
'            L_CodeLine = ArrRoutine(RLine)
'            If Not JustACommentOrBlank(L_CodeLine) Then
'              'v2.3.3 Thanks Roy Blanch. This was a spin off from first patch
'              If Not isProcHead(L_CodeLine) Then
'                If InstrAtPosition(L_CodeLine, "Static", IpLeft) Then
'                  Select Case FixData(DetectStatic2Private_CXX).FixLevel
'                   Case CommentOnly
'                    ArrRoutine(RLine) = Marker(ArrRoutine(RLine), SUGGESTION_MSG & "Static is very memory hungry; try using a Private Module level variable instead", MAfter, UpDated)
'                  End Select
'                  AddNfix DetectStatic2Private_CXX
'                End If
'              End If
'            End If
'          Next RLine
'          UpdateMember arrMembers(Member), ArrRoutine, UpDated, MUpdated
'        End If
'      Next Member
'    End If
'    ReWriteMembers cMod, arrMembers, MUpdated
'  End If
'
'End Sub
''
Private Sub DangerousCoding(cMod As CodeModule)

  Dim L_CodeLine   As String
  Dim arrMembers   As Variant
  Dim ArrRoutine   As Variant
  Dim Member       As Long
  Dim J            As Long
  Dim RLine        As Long
  Dim UpDated      As Boolean
  Dim MUpdated     As Boolean
  Dim MaxFactor    As Long
  Dim ModuleNumber As Long
  Dim bWrongSlash  As Boolean

  'ver 1.0.94
  ' this consolidates 5 routines into one should add to speed of CF
  'NOte individual tests can still be turned on/off
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'v3.0.3 reduced by shifting several tests to DangerousCodingFast
  'for even greater speed the remaining test are too complex for the fast way (so far)
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If AnyDangerousCodeTest(ModuleNumber) Then
    arrMembers = GetMembersArray(cMod)
    MaxFactor = UBound(arrMembers)
    If MaxFactor > 0 Then
      For Member = 0 To MaxFactor
        MemberMessage GetProcNameStr(arrMembers(Member)), Member, MaxFactor
        ArrRoutine = Split(arrMembers(Member), vbNewLine)
        For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
          L_CodeLine = ArrRoutine(RLine)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If dofix(ModuleNumber, DetectHardPath_CXX) Then
              If AnyHardPaths(L_CodeLine, bWrongSlash) Then
                If bWrongSlash Then
                  ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Hard-Coded Path with forward slash only works if you are writing to root folder(Bad idea)," & " Use 'App.Path & " & DQuote & "\FileName" & DQuote & "' or SpecialFolders API instead as you can be sure(or test)that the path will exist", MAfter)
                 Else
                  If InStr(L_CodeLine, "ShellExecute") Then
                    'ver 2.0.4
                    ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Hard-Coded URLs may make go out of date be sure to provide a human readable contact point.", MAfter)
                   Else
                    ArrRoutine(RLine) = SmartMarker(ArrRoutine, RLine, SUGGESTION_MSG & "Hard-Coded Paths are generally unwise, " & " Use 'App.Path & " & DQuote & "\FileName" & DQuote & "' or SpecialFolders API instead as you can be sure(or test)that the path will exist", MAfter)
                  End If
                End If
                UpDated = True
                AddNfix DetectHardPath_CXX
              End If
            End If
            If dofix(ModuleNumber, DetectDangerousAPICode_CXX) Then
              For J = LBound(DangerousAPIArray) To UBound(DangerousAPIArray)
                'If InstrAtPosition(L_Codeline, DangerousAPIArray(J), IpLeft) Then
                If WordInString(L_CodeLine, 1) = DangerousAPIArray(J) Then
                  If IsDeclareName(DangerousAPIArray(J)) Then
                    'IsDeclaration(DangerousAPIArray(J)) Then
                    ArrRoutine(RLine) = Marker(ArrRoutine(RLine), RISK_MSG & " API Code " & DangerousAPIArray(J) & RiskMessage(DangerousAPILevelArray(J)), MAfter, UpDated)
                    AddNfix DetectDangerousAPICode_CXX
                   Else
                    ArrRoutine(RLine) = Marker(ArrRoutine(RLine), RISK_MSG & " A user defined function has same name as an API procedure." & vbNewLine & _
                     RGSignature & "May make code harder to read", MAfter, UpDated)
                  End If
                End If
              Next J
            End If
            'v3.0.3 removed to DangerousCodingFast
          End If
        Next RLine
        UpdateMember arrMembers(Member), ArrRoutine, UpDated, MUpdated
      Next Member
    End If
    'v 2.2.2 added safety message (only appears if no other risks are found
    'v 2.2.3 removed it mis-hits
    '    If UsingScripting Then
    '      If Not MUpdated Then
    '        arrMembers(LBound(arrMembers)) = RISK_MSG & "This project uses Scripting. This may have potential to be very dangerous." & vbNewLine &
    ' '         arrMembers(LBound(arrMembers))
    '        MUpdated = True
    '      End If
    '    End If
    ReWriteMembers cMod, arrMembers, MUpdated
  End If

End Sub

''
Private Sub DangerousCodingFast()

  Dim J            As Long

  'v3.0.3 much faster version of DangerousCoding
  ' takes advantage of the FindAndComment fat approach
  ' may cause some doubling up of comments
  If FixData(DetectDangerousString_CXX).FixLevel > Off Then
    For J = LBound(DangerousStringArray) To UBound(DangerousStringArray)
      FindAndComment DangerousStringArray(J), RISK_MSG & "'" & DangerousStringArray(J) & "' could be used dangerously. Check that the code is safe.", , , True
    Next J
  End If
  If FixData(DetectDangerousAPICode_CXX).FixLevel > Off Then
    For J = LBound(DangerousAPIArray) To UBound(DangerousAPIArray)
      FindAndComment DangerousAPIArray(J), RISK_MSG & " API Code " & DangerousAPIArray(J) & RiskMessage(DangerousAPILevelArray(J)), , True
    Next J
  End If
  If FixData(DetectDangerousCode_CXX).FixLevel > Off Then
    For J = LBound(DangerousCodeArray) To UBound(DangerousCodeArray)
      FindAndComment DangerousCodeArray(J), RISK_MSG & " Risky Code " & DangerousCodeArray(J) & RiskMessage(DangerousCodeLevelArray(J)), , True
    Next J
    For J = LBound(DangerousScriptArray) To UBound(DangerousScriptArray)
      FindAndComment DangerousScriptArray(J), RISK_MSG & " Risky Code " & DangerousScriptArray(J) & RiskMessage(DangerousScriptLevelArray(J)), , True
    Next J
  End If
  If FixData(DetectDangerousReference_CXX).FixLevel > Off Then
    For J = LBound(DangerousReferenceArray) To UBound(DangerousReferenceArray)
      FindAndComment DangerousReferenceArray(J), RISK_MSG & " Reference to " & DangerousReferenceArray(J) & RiskMessage(DangerousReferenceArray(J)), , True
    Next J
  End If

End Sub

Private Sub DuplicatePublicProcedures(cMod As CodeModule, _
                                      Cmp As VBComponent)

  Dim CurCompCount As Long
  Dim arrMembers   As Variant
  Dim ArrProc      As Variant
  Dim I            As Long
  Dim UpDated      As Boolean
  Dim MUpdated     As Boolean
  Dim MaxFactor    As Long
  Dim TopOfRoutine As Long
  Dim lngdummy     As Long
  Dim Rname        As String
  Dim bGotOne      As Boolean

  If dofix(CurCompCount, DuplicatePublicProc_CXX) Then
    If bProcDescExists Then
      For I = LBound(PRocDesc) To UBound(PRocDesc)
        If PRocDesc(I).PrDDuplicate Then ' this is a rare condition so test it first
          bGotOne = True
          Exit For
        End If
      Next I
      If bGotOne Then
        arrMembers = GetMembersArray(cMod)
        MaxFactor = UBound(arrMembers)
        UpDated = False
        If MaxFactor > -1 Then
          For I = 1 To MaxFactor
            If Len(arrMembers(I)) Then
              ArrProc = Split(arrMembers(I), vbNewLine)
              MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
              GetRoutineDeclaration ArrProc, TopOfRoutine, lngdummy, Rname
              'v2.4.4 reconfigured to short circuit
              If Not IsComponent_User_Class(Cmp) Then
                If IsDuplicateProcName(Rname) Then
                  Select Case FixData(DuplicatePublicProc_CXX).FixLevel
                   Case CommentOnly
                    ArrProc(TopOfRoutine) = Marker(ArrProc(TopOfRoutine), SUGGESTION_MSG & "Routine's name is Duplicated in another module." & vbNewLine & _
                     "If they are identical, pick one and delete the other, otherwise rename one of them.", MAfter, UpDated)
                    AddNfix DuplicatePublicProc_CXX
                  End Select
                End If
              End If
            End If
            UpdateMember arrMembers(I), ArrProc, UpDated, MUpdated
          Next I
        End If
        ReWriteMembers cMod, arrMembers, MUpdated
      End If
    End If
  End If

End Sub

Private Sub ElseIfEvaluator(cMod As CodeModule)

  Dim Member      As Long
  Dim ArrMember   As Variant
  Dim ArrRoutine  As Variant
  Dim UpDated     As Boolean
  Dim MUpdated    As Boolean
  Dim MemberCount As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  If dofix(ModDescMember(cMod.Parent.Name), DoElseIfEval_CXX) Then
    ArrMember = GetMembersArray(cMod)
    MemberCount = UBound(ArrMember)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        If InStr(ArrMember(Member), "ElseIf ") Then
          MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
          ArrRoutine = Split(ArrMember(Member), vbNewLine)
          ElseIfEvaluatorTest ArrRoutine, UpDated
          UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
        End If
      Next Member
      ReWriteMembers cMod, ArrMember, MUpdated
    End If
  End If

End Sub

Private Sub ElseIfEvaluatorTest(arrR As Variant, _
                                UpDated As Boolean)

  Dim I             As Long
  Dim arrTmp        As Variant
  Dim J             As Long
  Dim K             As Long
  Dim HasSafetyElse As Boolean
  Dim Convert2Case  As Boolean
  Dim EndLine       As Long
  Dim Level         As Long
  Dim ElseIfCount   As Long

  For I = 0 To UBound(arrR)
    EndLine = I
    If TypeOfIf(arrR, I, 0) = Complex2 Then
      K = -1
      Level = 0
      ElseIfCount = 0
      ReDim arrTmp(UBound(arrR)) As Variant
      For J = I To UBound(arrR)
        If MultiLeft(arrR(J), True, "If ", "ElseIf ", "Else") Then
          If Left$(arrR(J), 3) = "If " Then
            Level = Level + 1
          End If
          If Left$(arrR(J), 7) = "ElseIf " Then
            ElseIfCount = ElseIfCount + 1
          End If
          K = K + 1
          EndLine = J
          arrTmp(K) = arrR(J)
        End If
        If Left$(arrR(J), 6) = "End If" Then
          Level = Level - 1
          K = K + 1
          EndLine = J
          arrTmp(K) = arrR(J)
          If Level = 0 Then
            Exit For
          End If
        End If
      Next J
      ReDim Preserve arrTmp(K) As Variant
      ' less than 3 is not worth considering
      If K > -1 Then
        HasSafetyElse = LeftWord(arrTmp(UBound(arrTmp) - 1)) = "Else" And Level = 0
        'And ElseIfCount > 2
        If Not HasSafetyElse Then
          If ElseIfCount > 2 Then
            arrR(EndLine) = Marker(arrR(EndLine), SUGGESTION_MSG & "Default 'Else' section of code is a good safety net at end of 'If..ElseIf..End If' structures", MBefore, UpDated)
          End If
        End If
        Convert2Case = ElseIFToCase(arrTmp, HasSafetyElse) And ElseIfCount > 2
        If Convert2Case Then
          arrR(I) = Marker(arrR(I), SUGGESTION_MSG & "'ElseIf' structure may be more efficient if converted to a 'Case' structure.", MAfter, UpDated)
        End If
      End If
    End If
    I = EndLine + 1
  Next I

End Sub

Private Function ElseIFToCase(arrTest As Variant, _
                              ByVal FinalElse As Boolean) As Boolean

  Dim CommonCount As Long
  Dim arrTest2    As Variant
  Dim I           As Long

  'ver1116 made test more rigerous ot avoid so many miss hits
  arrTest2 = Split(arrTest(0))
  For I = 1 To UBound(arrTest2) - 1 'skip the terminal 'End If' as it can't have shared ref
    CommonCount = 0
    If InStr(arrTest(2), arrTest2(I)) Then
      CommonCount = CommonCount + 1
    End If
    If CommonCount = UBound(arrTest) + IIf(FinalElse, -1, 0) Then
      ElseIFToCase = True
      Exit For
    End If
  Next I

End Function

Public Sub FindAndComment(ByVal strFind As String, _
                          ByVal strComment As String, _
                          Optional ByVal bPattern As Boolean = True, _
                          Optional ByVal bWhole As Boolean = False, _
                          Optional ByVal bInStrLiteral As Boolean = False)

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim strFound     As String
  Dim StartLine    As Long
  Dim SCol         As Long

  'v3.0.3 added bWhole and bInStrLiteral
  'bInStrLiteral to find string literal text (for DangerousStringArray)
  'v2.7.2 new fix
  'allow you to detect and comment on any code
  If bPattern Then
    If bWhole Then
      bPattern = False
    End If
  End If
  If Not MultiLeft(strComment, True, SUGGESTION_MSG, WARNING_MSG, UPDATED_MSG, PREVIOUSCODE_MSG) Then
    strComment = SUGGESTION_MSG & strComment
  End If
  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        ModuleMessage Comp, CurCompCount
        StartLine = 1
        SCol = 1
        With Comp
          Do While .CodeModule.Find(strFind, StartLine, SCol, -1, -1, bWhole, False, bPattern)
            strFound = .CodeModule.Lines(StartLine, 1)
            MemberMessage "", StartLine, .CodeModule.CountOfLines
            If bInStrLiteral Then
              If InLiteral(strFound, SCol) Then
                SafeInsertModule .CodeModule, StartLine, strComment
              End If
             ElseIf InCode(strFound, SCol) Then
              SafeInsertModule .CodeModule, StartLine, strComment
            End If
            StartLine = StartLine + 1
            If StartLine > .CodeModule.CountOfLines Then
              Exit Do
            End If
            SCol = 1
          Loop
        End With 'Comp
      End If
    Next Comp
  Next Proj
  On Error GoTo 0

End Sub

Private Sub Generate_GotoArray(arrR As Variant, _
                               ByVal CompName As String)

  
  Dim ILevel     As Long
  Dim FLevel     As Long
  Dim strTargets As String
  Dim arrTargets As Variant
  Dim I          As Long
  Dim J          As Long
  Dim strTmp     As String

  For I = LBound(arrR) To UBound(arrR)
    If Not JustACommentOrBlank(arrR(I)) Then
      If InStrCode(arrR(I), "GoTo ") Then
        If IsRealWord(arrR(I), "GoTo") Then
          strTargets = AccumulatorString(strTargets, WordAfter(arrR(I), "GoTo ") & "," & I)
        End If
      End If
      If InStrCode(arrR(I), " Resume ") Or Left$(arrR(I), 7) = "Resume " Then
        If IsRealWord(arrR(I), "Resume") Then
          If WordAfter(arrR(I), "Resume") <> "Next" Then
            strTargets = AccumulatorString(strTargets, "Resume " & WordAfter(arrR(I), "Resume ") & "," & I)
          End If
        End If
      End If
      If InStrCode(arrR(I), " GoSub ") Or SmartLeft(arrR(I), "GoSub ") Then
        strTargets = AccumulatorString(strTargets, "GoSub " & WordAfter(arrR(I), "GoSub") & "," & I)
      End If
    End If
  Next I
  If LenB(strTargets) = 0 Then
    ReDim GoToData(0) As GoToLocs
   Else
    arrTargets = Split(strTargets, ",")
    strTargets = vbNullString
    ReDim GoToData(UBound(arrTargets) \ 2) As GoToLocs
    For I = LBound(arrTargets) To UBound(arrTargets) Step 2
      With GoToData(J)
        If LeftWord(arrTargets(I)) = "Resume" Then
          arrTargets(I) = WordInString(arrTargets(I), 2)
          .RsumGo2 = True
        End If
        If LeftWord(arrTargets(I)) = "GoSub" Then
          arrTargets(I) = WordInString(arrTargets(I), 2)
          .G0Sub2 = True
        End If
        .Name = arrTargets(I)
        'FIXME this is a temp fix for a bug caused by 'On Num Goto 1,2,3,4.....
        If I < UBound(arrTargets) Then
          .LineNo = IIf(Len(arrTargets(I + 1)), arrTargets(I + 1), 0)
        End If
        .GoToTarget = False
      End With 'GoToData(J)
      J = J + 1
    Next I
    strTargets = vbNullString
    For I = LBound(arrR) To UBound(arrR)
      For J = LBound(GoToData) To UBound(GoToData)
        If InstrAtPositionArray(arrR(I), IpLeft, True, GoToData(J).Name, GoToData(J).Name & ":") Then
          strTargets = AccumulatorString(strTargets, GoToData(J).Name & CommaSpace & I)
        End If
      Next J
    Next I
    If Len(strTargets) Then
      J = UBound(GoToData) + 1
      arrTargets = Split(strTargets, ",")
      strTargets = vbNullString
      ReDim Preserve GoToData(J + (UBound(arrTargets) \ 2)) As GoToLocs
      For I = LBound(arrTargets) To UBound(arrTargets) Step 2
        With GoToData(J)
          .Name = arrTargets(I)
          .LineNo = arrTargets(I + 1)
          .GoToTarget = True
        End With 'GoToData(J)
        J = J + 1
      Next I
    End If
    For I = LBound(arrR) To UBound(arrR)
      strTmp = WordInString(arrR(I), 1)
      If IsGotoLabel(strTmp, CompName) Then
        If Right$(strTmp, 1) = ":" Then
          strTmp = Left$(strTmp, Len(strTmp) - 1)
        End If
        strTargets = AccumulatorString(strTargets, strTmp & CommaSpace & I)
      End If
    Next I
    If Len(strTargets) Then
      arrTargets = Split(strTargets, ",")
      For I = LBound(arrTargets) To UBound(arrTargets) Step 2
        For J = LBound(GoToData) To UBound(GoToData)
          If GoToData(J).Name = arrTargets(I) Then
            arrTargets(I) = vbNullString
            arrTargets(I + 1) = vbNullString
            Exit For
          End If
        Next J
      Next I
      strTmp = Trim$(Replace$(Join(arrTargets, ","), ",  , ", vbNullString))
      strTmp = RStrip(strTmp, ",")
      strTmp = LStrip(strTmp, ",")
      arrTargets = Split(strTmp, ",")
      If UBound(arrTargets) > -1 Then
        If LenB(arrTargets(0)) Then
          J = UBound(GoToData) + 1
          ReDim Preserve GoToData(J + (UBound(arrTargets) \ 2)) As GoToLocs
          For I = LBound(arrTargets) To UBound(arrTargets) Step 2
            With GoToData(J)
              .Name = arrTargets(I)
              .LineNo = Val(arrTargets(I + 1))
              .GoToTarget = True
              .Orphen = True
            End With 'GoToData(J)
            J = J + 1
          Next I
        End If
      End If
    End If
    For I = LBound(GoToData) To UBound(GoToData)
      For J = LBound(GoToData) To UBound(GoToData)
        If I <> J Then
          With GoToData(I)
            If .Name = GoToData(J).Name Then
              'v2.4.4 reconfigured to short circuit
              If Not .GoToTarget Then
                If GoToData(J).GoToTarget Then
                  .Targets = AccumulatorString(.Targets, GoToData(J).LineNo)
                End If
              End If
            End If
          End With 'GoToData(I)
        End If
      Next J
    Next I
    For I = LBound(GoToData) To UBound(GoToData)
      If InstrAtPosition(arrR(GoToData(I).LineNo), "Error GoTo", ipAny) Then
        GoToData(I).ErrorGoto = True
        GoToData(I).Error0 = GoToData(I).Name = "0"
      End If
    Next I
    For I = LBound(arrR) To UBound(arrR)
      GetIndentLevel arrR(I), ILevel
      GetStructureLevel arrR(I), FLevel, "If", "End If"
      For J = LBound(GoToData) To UBound(GoToData)
        If GoToData(J).LineNo = I Then
          GoToData(J).IndentLevel = ILevel
          GoToData(J).IfLevel = FLevel
          Exit For
        End If
      Next J
    Next I
    For I = LBound(arrR) To UBound(arrR)
      For J = LBound(GoToData) To UBound(GoToData)
        If GoToData(J).LineNo = I Then
          GoToData(J).IndentLevel = ILevel
          GoToData(J).IfLevel = FLevel
          Exit For
        End If
      Next J
    Next I
  End If

End Sub

Private Sub GetIndentLevel(ByVal VarLine As Variant, _
                           ByVal IndentLvl As Long)

  'must be called form inside a loop which is stepping through a procedure
  'assumes that code is not using colon seperators

  GetStructureLevel VarLine, IndentLvl, "With", "End With"
  GetStructureLevel VarLine, IndentLvl, "For", "Next"
  GetStructureLevel VarLine, IndentLvl, "Do", "Loop"
  GetStructureLevel VarLine, IndentLvl, "While", "Wend"
  GetStructureLevel VarLine, IndentLvl, "Select", "End Select"
  GetStructureLevel VarLine, IndentLvl, "If", "End If"
  GetStructureLevel VarLine, IndentLvl, "ElseIf", "ElseIf,Else,End If"
  GetStructureLevel VarLine, IndentLvl, "Case", "Case,End Select"

End Sub

Private Sub GetStructureLevel(ByVal VarLine As Variant, _
                              Slvl As Long, _
                              ByVal strStart As String, _
                              ByVal strEnd As String)

  Dim arrLine       As Variant
  Dim arrEnd        As Variant
  Dim FakeDelimiter As String
  Dim I             As Long
  Dim J             As Long
  Dim strTest       As String

  'must be called form inside a loop which is stepping through a procedure
  'assumes that code is not using colon seperators
  'lines with the same sLvl are in same structrue
  'while the slvl stays above 0
  Do
    FakeDelimiter = RandomString(48, 122, 3, 6)
  Loop While InStr(VarLine, FakeDelimiter)
  If InStrCode(VarLine, ":") Then
    VarLine = Safe_Replace(VarLine, ":", FakeDelimiter, , , True)
    arrLine = Split(VarLine, FakeDelimiter)
   Else
    'single line insterted into array
    arrLine = Split(VarLine, FakeDelimiter)
  End If
  'this copes with the cases of ElseIf (end factors=ElseIf,Else,End If) and Case (Case,End Select)
  arrEnd = Split(strEnd, ",")
  For I = LBound(arrLine) To UBound(arrLine)
    strTest = strCodeOnly(arrLine(I))
    If InstrAtPosition(strTest, strStart, IpLeft) Then
      If strStart = "If" Then
        If Left$(strTest, 3) = "If " Then
          If Right$(strTest, 5) = " Then" Then
            Slvl = Slvl + 1
          End If
        End If
       Else
        Slvl = Slvl + 1
        If UBound(arrEnd) > 0 Then
          Slvl = Slvl - 1
        End If
      End If
    End If
    For J = LBound(arrEnd) To UBound(arrEnd)
      If InstrAtPosition(strTest, arrEnd(J), IpLeft) Then
        Slvl = Slvl - 1
        If UBound(arrEnd) > 0 Then
          Slvl = Slvl + 1
        End If
      End If
    Next J
  Next I

End Sub

Private Function GoTo2Exit(ByVal arrR As Variant, _
                           GTData As GoToLocs, _
                           ByVal strTrigger As String, _
                           TLine As Long) As Boolean

  Dim I            As Long
  Dim J            As Long
  Dim TopScan      As Long
  Dim EndScan      As Long
  Dim arrTarget    As Variant
  Dim Depth        As Long
  Dim TriggerCount As Long

  arrTarget = Split(GTData.Targets, ",")
  For I = LBound(arrTarget) To UBound(arrTarget)
    TriggerCount = 0
    Depth = 0
    If arrTarget(I) > GTData.LineNo Then
      TopScan = GTData.LineNo
      EndScan = arrTarget(I)
      'scan down only needed for this
      For J = TopScan To EndScan
        GetIndentLevel arrR(J), Depth
        If SmartLeft(arrR(J), strTrigger) Then
          TriggerCount = TriggerCount + 1
        End If
        If TriggerCount > 1 Then ' to many to be safe
          Exit For
        End If
      Next J
    End If
    If Depth = -1 Then
      If TriggerCount = 1 Then
        '-1 means that there are no other structures intervening between GoTo and target label
        GoTo2Exit = True
        TLine = arrTarget(I)
        Exit For
      End If
    End If
  Next I

End Function

Private Sub GoToCommentry(cMod As CodeModule)

  
  Dim M                 As Long
  Dim NoInteveningTests As Boolean
  Dim arrMembers        As Variant
  Dim ArrRoutine        As Variant
  Dim Member            As Long
  Dim UpDated           As Boolean
  Dim MUpdated          As Boolean
  Dim MaxFactor         As Long
  Dim I                 As Long
  Dim J                 As Long
  Dim K                 As Long
  Dim TLine             As Long
  Dim ModuleNumber      As Long

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, GoToComments_CXX) Then
    arrMembers = GetMembersArray(cMod)
    MaxFactor = UBound(arrMembers)
    If MaxFactor > 0 Then
      For Member = 0 To MaxFactor
        'v2.4.4 simplifed test goto in code or any orphaned target label(rare)
        If InStr(arrMembers(Member), "GoTo ") Or InStr(arrMembers(Member), ":" & vbNewLine) Then
          MemberMessage GetProcNameStr(arrMembers(Member)), Member, MaxFactor
          ArrRoutine = Split(arrMembers(Member), vbNewLine)
          Generate_GotoArray ArrRoutine, cMod.Parent.Name
          If UBound(GoToData) > -1 Then
            If LenB(GoToData(0).Name) Then
              For I = LBound(GoToData) To UBound(GoToData)
                If GoToData(I).Orphen Then
                  'label without GoTo
                  Select Case FixData(GoToComments_CXX).FixLevel
                   Case CommentOnly
                    ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo label is orphaned or a VB structural word which does not require a colon. Delete the label or colon.", MAfter, UpDated)
                  End Select
                End If
                If Not GoToData(I).GoToTarget Then
                  'then it is a GoTo command
                  If Len(GoToData(I).Targets) = 0 Then
                    If GoToData(I).Name <> "0" Then 'ignore error off switch
                      Select Case FixData(GoToComments_CXX).FixLevel
                       Case CommentOnly
                        ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo command has no target label. You must delete it.", MAfter, UpDated)
                      End Select
                      AddNfix GoToComments_CXX
                     Else ' is a GoTo 0
                      For J = I To UBound(GoToData) - 1
                        If GoToData(J).Error0 Then
                          For K = J + 1 To UBound(GoToData)
                            If GoToData(K).ErrorGoto Then
                              If Not GoToData(K).Error0 Then
                                Exit For ' error checking was turned back on
                               ElseIf GoToData(K).ErrorGoto And GoToData(K).Error0 And True Then
                                'second error off found
                                'v2.3.4 skip if there is some other On Error between them that reactivates the error trap
                                If GoToData(J).IfLevel = GoToData(K).IfLevel Then
                                  NoInteveningTests = True
                                  For M = GoToData(J).LineNo + 1 To GoToData(K).LineNo - 1
                                    If IsOnErrorCode(ArrRoutine(M)) Then
                                      NoInteveningTests = False
                                      Exit For
                                    End If
                                  Next M
                                  If NoInteveningTests Then
                                    Select Case FixData(GoToComments_CXX).FixLevel
                                     Case CommentOnly
                                      If InStr(ArrRoutine(GoToData(J).LineNo), "following one may be unnecessary.") = 0 Then
                                        ArrRoutine(GoToData(J).LineNo) = Marker(ArrRoutine(GoToData(J).LineNo), WARNING_MSG & "This 'GoTo 0' or the following one may be unnecessary.", MAfter, UpDated)
                                      End If
                                      If InStr(ArrRoutine(GoToData(K).LineNo), "previous one may be unnecessary.") = 0 Then
                                        ArrRoutine(GoToData(K).LineNo) = Marker(ArrRoutine(GoToData(K).LineNo), WARNING_MSG & "This 'GoTo 0' or the previous one may be unnecessary.", MAfter, UpDated)
                                      End If
                                    End Select
                                    AddNfix GoToComments_CXX
                                    Exit For
                                  End If
                                End If
                              End If
                            End If
                          Next K
                        End If
                      Next J
                    End If
                  End If
                  Select Case ViolatesStructure(ArrRoutine, GoToData(I), "With", "End With")
                   Case VInto
                    Select Case FixData(GoToComments_CXX).FixLevel
                     Case CommentOnly
                      ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo jumps your code into a With structure. Invalid code.", MAfter, UpDated)
                    End Select
                    AddNfix GoToComments_CXX
                   Case VOutof
                    'dealt with in illegalWithExit
                    If Not GoToData(I).G0Sub2 Then
                      'gosub always returns to within structure so safe from this warning
                      Select Case FixData(GoToComments_CXX).FixLevel
                       Case CommentOnly
                        ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo jumps your code out of a With structure. May cause memory leaks. Check if Error Handler returns code flow to the 'With' structure.", MAfter, UpDated)
                      End Select
                      AddNfix GoToComments_CXX
                    End If
                  End Select
                  Select Case ViolatesStructure(ArrRoutine, GoToData(I), "For", "Next")
                   Case VOutof
                    If GoTo2Exit(ArrRoutine, GoToData(I), "Next", TLine) Then
                      Select Case FixData(GoToComments_CXX).FixLevel
                       Case CommentOnly
                        ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo might be replaced with an 'Exit For'", MAfter)
                        ArrRoutine(TLine) = Marker(ArrRoutine(TLine), WARNING_MSG & "This GoTo label is the target of a GoTo which could be replace with an 'Exit For'. Delete It.", MAfter, UpDated)
                       Case FixAndComment
                        UpDated = True
                      End Select
                      AddNfix GoToComments_CXX
                    End If
                   Case VInto
                    Select Case FixData(GoToComments_CXX).FixLevel
                     Case CommentOnly
                      If Not GoToData(I).ErrorGoto Then
                        ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo jumps your code into a For structure. Invalid code.", MAfter, UpDated)
                      End If
                     Case FixAndComment
                      UpDated = True
                    End Select
                  End Select
                  Select Case ViolatesStructure(ArrRoutine, GoToData(I), "Do", "Loop")
                   Case VInto
                    If GoTo2Exit(ArrRoutine, GoToData(I), "Loop", TLine) Then
                      Select Case FixData(GoToComments_CXX).FixLevel
                       Case CommentOnly
                        ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo might be replaced with an 'Exit Do'", MAfter, UpDated)
                        ArrRoutine(TLine) = Marker(ArrRoutine(TLine), WARNING_MSG & "This GoTo label is the target of a GoTo which could be replace with an 'Exit Do'. Delete it.", MAfter, UpDated)
                       Case FixAndComment
                        UpDated = True
                      End Select
                      AddNfix GoToComments_CXX
                    End If
                   Case VOutof
                    Select Case FixData(GoToComments_CXX).FixLevel
                     Case CommentOnly
                      ArrRoutine(GoToData(I).LineNo) = Marker(ArrRoutine(GoToData(I).LineNo), WARNING_MSG & "This GoTo jumps your code out of a Do..Loop structure.", MAfter, UpDated)
                     Case FixAndComment
                      UpDated = True
                    End Select
                  End Select
                End If
              Next I
            End If
          End If
          UpdateMember arrMembers(Member), ArrRoutine, UpDated, MUpdated
          'End If
        End If
      Next Member
    End If
    ReWriteMembers cMod, arrMembers, MUpdated
  End If

End Sub

Private Sub IllegalWithExit(cMod As CodeModule)

  
  Dim UpDated     As Boolean
  Dim Member      As Long
  Dim RLine       As Long
  Dim RLine2      As Long
  Dim RLine3      As Long
  Dim MemberCount As Long
  Dim strTest     As String
  Dim L_CodeLine  As String
  Dim L_Codeline2 As String
  Dim L_Codeline3 As String
  Dim arrMembers  As Variant
  Dim ArrRoutine  As Variant
  Dim SDeepMaster As Long
  Dim SDeep       As Long
  Dim lngdummy    As Long

  If dofix(ModDescMember(cMod.Parent.Name), DetectIllegalWithExit_CXX) Then
    arrMembers = GetMembersArray(cMod)
    MemberCount = UBound(arrMembers)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        If InStr(arrMembers(Member), "Exit ") Then
          MemberMessage GetProcNameStr(arrMembers(Member)), Member, MemberCount
          ArrRoutine = Split(arrMembers(Member), vbNewLine)
          If UBound(ArrRoutine) > 0 Then
            RLine = 0
            Do
              L_CodeLine = ArrRoutine(RLine)
              If Not JustACommentOrBlank(L_CodeLine) Then
                SDeepMaster = 1
                If InstrAtPosition(L_CodeLine, "With", IpLeft) Then
                  If LeftWord(L_CodeLine) = "With" Then
                    SDeepMaster = SDeepMaster + 1
                   ElseIf InstrAtPosition(L_CodeLine, "End With", IpLeft) Then
                    SDeepMaster = SDeepMaster - 1
                  End If
                  RLine2 = RLine
                  Do
                    RLine2 = RLine2 + 1
                    If RLine2 > UBound(ArrRoutine) Then
                      Exit Do
                    End If
                    L_Codeline2 = ArrRoutine(RLine2)
                    If Not JustACommentOrBlank(L_Codeline2) Then
                      If InstrAtPositionSetArray(L_Codeline2, IpLeft, True, ArrExitFuncPropSub) Then
                        If InStructure(WithStruct, ArrRoutine, RLine2, lngdummy, lngdummy) Then
                          Select Case FixData(DetectIllegalWithExit_CXX).FixLevel
                           Case CommentOnly
                            ArrRoutine(RLine2) = Marker(ArrRoutine(RLine2), WARNING_MSG & "Exiting a procedure from within a With Structure can lead to memory leaks" & vbNewLine & _
                             "It is advised that you re-structure the code around this line.", MAfter, UpDated)
                          End Select
                        End If
                      End If
                      If InstrAtPositionArray(L_Codeline2, IpLeft, True, "Exit Do", "Exit For", "GoTo") Then
                        If SmartLeft(L_Codeline2, "Exit Do") Then
                          strTest = "Loop"
                         ElseIf SmartLeft(L_Codeline2, "Exit For") Then
                          strTest = "Next"
                         Else
                          strTest = WordInString(L_Codeline2, 2)
                        End If
                        RLine3 = RLine2
                        SDeep = 1
                        Do
                          RLine3 = RLine3 + 1
                          L_Codeline3 = ArrRoutine(RLine3)
                          If SmartLeft(L_Codeline3, strTest) Then
                            strTest = vbNullString
                            Exit Do
                          End If
                          If LeftWord(L_Codeline3) = "With" Then
                            SDeep = SDeep + 1
                           ElseIf InstrAtPosition(L_Codeline3, "End With", IpLeft) Then
                            SDeep = SDeep - 1
                          End If
                        Loop Until InstrAtPosition(L_Codeline3, "End With", IpLeft) And SDeep = 0
                        If InstrAtPosition(L_Codeline2, "GoTo", IpLeft, True) Then
                          'scan up  becuase GoTo can go that way.
                          If Len(strTest) Then
                            RLine3 = RLine2
                            Do
                              RLine3 = RLine3 - 1
                              L_Codeline3 = ArrRoutine(RLine3)
                              If SmartLeft(L_Codeline3, strTest) Then
                                strTest = vbNullString
                                Exit Do
                              End If
                            Loop Until RLine3 = RLine
                          End If
                        End If
                        If Len(strTest) Then
                          Select Case FixData(DetectIllegalWithExit_CXX).FixLevel
                           Case CommentOnly
                            ArrRoutine(RLine2) = Marker(ArrRoutine(RLine2), WARNING_MSG & "This line jumps your code out of a With Structure and can lead to memory leaks." & vbNewLine & _
                             "It is advised that you re-structure the code around this line.", MAfter, UpDated)
                          End Select
                        End If
                      End If
                      If L_Codeline2 = "End" Then
                        Select Case FixData(DetectIllegalWithExit_CXX).FixLevel
                         Case CommentOnly
                          ArrRoutine(RLine2) = Marker(ArrRoutine(RLine2), WARNING_MSG & "'End' command inside a With Structure may cause memory leaks." & vbNewLine & _
                           "It is advised that you re-structure the code around this line.", MAfter, UpDated)
                        End Select
                      End If
                    End If
                    'v 2.1.5 added to stop overflow and wrong reproting of "End"
                    If InstrAtPosition(L_Codeline2, "End With", IpLeft) Then
                      SDeepMaster = SDeepMaster - 1
                    End If
                  Loop Until InstrAtPosition(L_Codeline2, "End With", IpLeft) And SDeepMaster = 1
                End If
              End If
              RLine = RLine + 1
            Loop Until RLine >= UBound(ArrRoutine)
            arrMembers(Member) = Join(ArrRoutine, vbNewLine)
          End If
        End If
      Next Member
    End If
    ReWriteMembers cMod, arrMembers, UpDated
  End If

End Sub

Private Function IsControlProc(Arr As Variant, _
                               Cmp As VBComponent, _
                               Optional strProcName As String) As Boolean

  GetRoutineDeclaration Arr, 0, 0, strProcName
  IsControlProc = RoutineNameIsVBGenerated(strProcName, Cmp, False)

End Function

Private Sub ObsoleteCode(cMod As CodeModule)

  Dim L_CodeLine   As String
  Dim arrMembers   As Variant
  Dim ArrRoutine   As Variant
  Dim Member       As Long
  Dim J            As Long
  Dim K            As Long
  Dim RLine        As Long
  Dim UpDated      As Boolean
  Dim MemberCount  As Long
  Dim ObsoleteMsg  As String
  Dim ModuleNumber As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, DetectObsoleteCodeStructure_CXX) Then
    arrMembers = GetMembersArray(cMod)
    MemberCount = UBound(arrMembers)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        MemberMessage GetProcNameStr(arrMembers(Member)), Member, MemberCount
        ArrRoutine = Split(arrMembers(Member), vbNewLine)
        For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
          L_CodeLine = ArrRoutine(RLine)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If WordInString(L_CodeLine, 1) = "Let" Then
              Select Case FixData(DetectObsoleteCodeStructure_CXX).FixLevel
               Case CommentOnly, FixAndComment
                ArrRoutine(RLine) = Mid$(L_CodeLine, 4)
                L_CodeLine = ArrRoutine(RLine)
                UpDated = True
              End Select
            End If
            For J = LBound(ObsoleteCodeArray) To UBound(ObsoleteCodeArray)
              If InstrAtPosition(L_CodeLine, ObsoleteCodeArray(J), IpLeft, True) Then
                'ver 1.1.81 deal with the word being part of the joined to something else by an underscore
                If WordInString(L_CodeLine, 1) = ObsoleteCodeArray(J) Then
                  Select Case FixData(DetectObsoleteCodeStructure_CXX).FixLevel
                   Case CommentOnly, FixAndComment
                    Select Case ObsoleteCodeArray(J)
                     Case "IsMissing"
                      If Not Xcheck(XIgnoreCom) Then
                        ObsoleteMsg = " 'IsMissing' code can be replaced by using the extended Optional parameter format 'Optional VarName[As Type][= DefaultValue ]'." & vbNewLine & _
                                      "Where 'DefaultValue' is whatever you assign to the variable in the 'IsMissing' code strucutre."
                      End If
                     Case "GoSub", "Return"
                      If Not Xcheck(XIgnoreCom) Then
                        ObsoleteMsg = " Create a real Sub instead, it is much safer as you have better control over variables."
                      End If
                     Case "Switch", "Choose"
                      If Not Xcheck(XIgnoreCom) Then
                        ObsoleteMsg = " Use Select Case.. End Select structure it is clearer to read and less error prone."
                      End If
                     Case "Goto"
                      If Not Xcheck(XIgnoreCom) Then
                        ObsoleteMsg = " GoTo is generally considered poor code. Often you can restructure code to avoid it. (BUT if it's the answer, use it:))"
                      End If
                     Case "While", "Wend"
                      Select Case FixData(DetectObsoleteCodeStructure_CXX).FixLevel
                       Case CommentOnly
                        If Not Xcheck(XIgnoreCom) Then
                          ObsoleteMsg = " Use Do...Loop it has 'While', 'Until' and' Exit' Methods which make it clearer and easier to code."
                        End If
                       Case FixAndComment
                        ArrRoutine(RLine) = "Do " & ArrRoutine(Member)
                        For K = Member To UBound(ArrRoutine)
                          If SmartLeft(ArrRoutine(K), "Wend ") Then
                            ArrRoutine(K) = Safe_Replace(ArrRoutine(K), "Wend", "Loop", , 1)
                            Exit For
                          End If
                        Next K
                        ObsoleteMsg = " 'Do...Loop' substituted for 'While ...Wend'"
                       Case JustFix
                        ArrRoutine(RLine) = "Do " & ArrRoutine(Member)
                        For K = Member To UBound(ArrRoutine)
                          If SmartLeft(ArrRoutine(K), "Wend ") Then
                            ArrRoutine(K) = Safe_Replace(ArrRoutine(K), "Wend", "Loop", , 1)
                            Exit For
                          End If
                        Next K
                      End Select
                    End Select
                    If Not Xcheck(XIgnoreCom) Then
                      ArrRoutine(RLine) = Marker(ArrRoutine(RLine), SUGGESTION_MSG & "Obsolete Code " & strInSQuotes(ObsoleteCodeArray(J)) & ObsoleteMsg, MAfter, UpDated)
                    End If
                    'Case JustFix
                  End Select
                  AddNfix DetectObsoleteCodeStructure_CXX
                End If
              End If
            Next J
          End If
        Next RLine
        arrMembers(Member) = Join(ArrRoutine, vbNewLine)
      Next Member
    End If
    ReWriteMembers cMod, arrMembers, UpDated
  End If

End Sub

Private Function RiskMessage(DLevel As Variant) As String

  Select Case DLevel
   Case 1
    RiskMessage = " Low Risk: Has potential to damage files/system but is usually safe."
   Case 2
    RiskMessage = " Medium Risk: Has potential to damage files/system but is generally useful."
   Case 3
    RiskMessage = " High Risk: Has high potential to damage files/system make sure it is safe."
   Case 4
    RiskMessage = " High Danger: Make sure that you understand what this bit of code does before running."
   Case 5
    RiskMessage = " Extreme Danger: Should probably not be allowed to run."
  End Select

End Function

Private Sub StaticToPrivateFast(cMod As CodeModule)

  'v3.0.4 faster version
  
  Dim StartLine  As Long
  Dim StrNewCode As String

  With cMod
    StartLine = .CountOfDeclarationLines
    Do While .Find("Static", StartLine, 1, -1, -1, True, True)
      StrNewCode = .Lines(StartLine, 1)
      If InCode(StrNewCode, InStr(StrNewCode, "Static ")) Then
        If Not isProcHead(StrNewCode) Then
          InsertNewCodeComment cMod, StartLine, 1, StrNewCode, SUGGESTION_MSG & "Static is very memory hungry; try using a Private Module level variable instead"
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

Public Sub Suggest_Engine()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long
  Dim NumFixes     As Long

  NumFixes = 13
  If Not bAborting Then
    On Error GoTo BugHit
    WorkingMessage "Comments #Win16", 1, NumFixes
    FindAndComment "#If Win16 Then", "Is 16-Bit support necessary?", False
    WorkingMessage "Comments #Win32", 2, NumFixes
    FindAndComment "#If Win32 Then", "Is 16-Bit support necessary?", False
    WorkingMessage "Comments 16-bit API", 3, NumFixes
    FindAndComment "Declare * " & DQuote & "User" & DQuote, "This is a 16-Bit API. " & DQuote & "User32" & DQuote & " version would give you 32-bit"
    WorkingMessage "Comments Pi", 4, NumFixes
    FindAndComment "Const * = 3.1415", "It is better to use a Variable and 'Pi = 4 * Atn(1)' in Form_Load/Class_Initialize rather than a hard coded Constant"
    'FindAndComment "If (*) Then", "Check outer brackets; may not be necessary"
    'FindAndComment "If (*) = (*) Then", "Check brackets; may not be necessary"
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          With Comp
            If .CodeModule.CountOfLines Then
              WorkingMessage "Poor Control name", 6, NumFixes
              BadControlName .CodeModule, Comp
              WorkingMessage "Detect Obsolete Code Structures", 7, NumFixes
              ObsoleteCode .CodeModule
              WorkingMessage "Risky API/Code/Ref/Str/Hard-Path", 8, NumFixes
              DangerousCoding .CodeModule
              WorkingMessage "Static To Private", 9, NumFixes
              'StaticToPrivate .CodeModule
              StaticToPrivateFast .CodeModule
              WorkingMessage "Duplicate Public Procedures", 10, NumFixes
              DuplicatePublicProcedures .CodeModule, Comp
              WorkingMessage "Illegal With Exits", 11, NumFixes
              IllegalWithExit .CodeModule
              WorkingMessage "GoTo evaluation", 12, NumFixes
              GoToCommentry .CodeModule
              WorkingMessage "ElseIf Evaluate", 13, NumFixes
              ElseIfEvaluator .CodeModule
              WorkingMessage "Large Procedures", 14, NumFixes
              VeryLargeProcedure .CodeModule
            End If
          End With 'Comp
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
    DangerousCodingFast
    On Error GoTo 0
  End If

Exit Sub

BugHit:
  BugTrapComment "Suggest_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub VeryLargeProcedure(cMod As CodeModule)

  Dim Msg              As String
  Dim Member           As Long
  Dim arrMembers       As Variant
  Dim ArrRoutine       As Variant
  Dim MUpdated         As Boolean
  Dim MemberCount      As Long
  Dim LngCodeLineCount As Long
  Dim strProcName      As String
  Dim UpDated          As Boolean

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'this routine uses 50 lines of code as the definition of a long procedure
  'change the 50 if you want it to be less fussy
  If dofix(ModDescMember(cMod.Parent.Name), VeryBigProc_CXX) Then
    arrMembers = GetMembersArray(cMod)
    MemberCount = UBound(arrMembers)
    If MemberCount > 0 Then
      For Member = 1 To MemberCount
        MemberMessage GetProcNameStr(arrMembers(Member)), Member, MemberCount
        ArrRoutine = Split(arrMembers(Member), vbNewLine)
        LngCodeLineCount = CountArrayCodeLines(ArrRoutine)
        If LngCodeLineCount > lngBigProcLines Then
          If IsControlProc(ArrRoutine, cMod.Parent, strProcName) Then
            Msg = WARNING_MSG & "Large Control procedure (" & LngCodeLineCount & " lines of code)"
            If RoutineNameIsUserInteractive(strProcName, cMod.Parent) Then
              Msg = Msg & vbNewLine & _
               "This is a User Interactive Control Procedures, it is good practice to create procedures to seperate funtionality."
            End If
           Else
            Msg = WARNING_MSG & "Large Code procedure (" & LngCodeLineCount & " lines of code)" & IIf(Xcheck(XVerbose), vbNewLine & _
             "It is recommended that you try to break it into smaller procedures", vbNullString)
          End If
          SafeInsertArrayMarker ArrRoutine, GetProcCodeLineOfRoutine(ArrRoutine), Msg
          UpDated = True
        End If
        UpdateMember arrMembers(Member), ArrRoutine, UpDated, MUpdated
      Next Member
      ReWriteMembers cMod, arrMembers, MUpdated
    End If
  End If

End Sub

Private Function ViolatesStructure(arrR As Variant, _
                                   GTData As GoToLocs, _
                                   ByVal strBegin As String, _
                                   ByVal strEnd As String) As StructViolation

  Dim I            As Long
  Dim J            As Long
  Dim TopScan      As Long
  Dim EndScan      As Long
  Dim arrTarget    As Variant
  Dim Depth        As Long
  Dim ViolatesDown As Boolean

  'returns
  '0 no violation
  '1 violates up
  '2 violates down
  arrTarget = Split(GTData.Targets, ",")
  If InStr(GTData.Name, "GoSub") Then
    ViolatesStructure = VNone
   Else
    For I = LBound(arrTarget) To UBound(arrTarget)
      If CLng(arrTarget(I)) > GTData.LineNo Then
        TopScan = GTData.LineNo
        EndScan = arrTarget(I)
        ViolatesDown = True
        'scan down
       Else
        'scanup
        TopScan = GTData.LineNo
        EndScan = arrTarget(I)
      End If
      Depth = 0
      For J = TopScan To EndScan
        If LeftWord(arrR(J)) = strBegin Then
          If InStrCode(arrR(J), strEnd) = 0 Then
            Depth = Depth + 1
          End If
        End If
        If strBegin = "With" Then
          If SmartLeft(arrR(J), strEnd) Then
            Depth = Depth - 1
          End If
         Else
          If LeftWord(arrR(J)) = strEnd Then
            Depth = Depth - 1
          End If
        End If
      Next J
      Select Case Depth
       Case 0 'no structural violation
        ViolatesStructure = VNone
       Case Is < 0 'violates into structure
        ViolatesStructure = IIf(ViolatesDown, VOutof, VInto)
        Exit For
       Case Is > 0 'violates out of structure
        ViolatesStructure = IIf(ViolatesDown, VInto, VOutof)
        Exit For
      End Select
    Next I
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:21:54 AM) 37 + 1382 = 1419 Lines Thanks Ulli for inspiration and lots of code.

