Attribute VB_Name = "mod_UnusedFix"

Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public Const MoveableComment      As String = "'~~~'"
Public ArrQCheckedVariables       As Variant
Public IsCollectionClass()        As Boolean
Private CallByNameIsUsed          As Boolean

Private Sub ActiveDebugDriver()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long

  'v2.4.2 outer driver for IfThen2BooleanEngine
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        If dofix(ModDescMember(Comp.Name), ActiveDebugStop) Then
          ModuleMessage Comp, CurCompCount
          WorkingMessage "Active Stop search", 1, 2
          ActiveDebugEngine Comp.CodeModule, "Stop"
          WorkingMessage "Active Debug search", 2, 2
          ActiveDebugEngine Comp.CodeModule, "Debug"
        End If
      End If
    Next Comp
  Next Proj

End Sub

Private Sub ActiveDebugEngine(cMod As CodeModule, _
                              ByVal strTarget As String)

  Dim CommentStore       As String
  Dim strTest            As String
  Dim StartLine          As Long
  Const ACTIVEDEBUG_MSG1 As String = SUGGESTION_MSG & "Code used to generate data only for Debug should be removed from final code."
  Const ACTIVEDEBUG_MSG2 As String = SUGGESTION_MSG & "Code used to reach a Stop command should usually be removed from final code."

  'v2.4.4 new version of fix
  With cMod
    Do While .Find(strTarget, StartLine, 1, -1, -1)
      MemberMessage "", StartLine, .CountOfLines
      If StartLine > .CountOfDeclarationLines Then
        strTest = .Lines(StartLine, 1)
        If ExtractCode(strTest, CommentStore) Then
          If MultiLeft(strTest, True, "Debug.Print", "Debug.Assert") Then
            .ReplaceLine StartLine, Marker(strTest & CommentStore, ACTIVEDEBUG_MSG1, MAfter)
          End If
          If InstrAtPosition(strTest, "Stop", ipAny) Then
            If IsRealWord(strTest, "Stop") Then
              .ReplaceLine StartLine, Marker(strTest & CommentStore, ACTIVEDEBUG_MSG2, MAfter)
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

Public Function CountCodeUsage(ByVal strFind As String, _
                               ByVal OrigLine As String, _
                               Optional ByVal CompName As String = vbNullString, _
                               Optional ByVal DeclarationOnly As Boolean = False, _
                               Optional ByVal CurrentOnly As Boolean = False, _
                               Optional ByVal SkipCurrent As Boolean = False, _
                               Optional ByVal PartialIsOK As Boolean = False, _
                               Optional ByVal bIgnoreClassUserDocument As Boolean = False) As Double

  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim CompMod        As CodeModule
  Dim code           As String
  Dim prevcode       As String
  Dim TestPos        As Long
  Dim CodeLineNo     As Long
  Dim GuardLine      As Long
  Dim PrevCodeLineNo As Long

  'ver 1.1.1
  'new checks existence of variables
  '(replaces several old routines)
  'uses Find for greater speed
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If CurrentOnly Then
          If CompName <> Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If SkipCurrent Then
          If CompName = Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If bIgnoreClassUserDocument Then
          If Comp.Type = vbext_ct_ClassModule Or Comp.Type = vbext_ct_UserControl Then
            GoTo SkipTest
          End If
        End If
        Set CompMod = Comp.CodeModule
        If CompMod.Find(strFind, 1, 1, -1, -1, Not PartialIsOK, True) Then
          '*if exits at all, then look for the line(s)
          PrevCodeLineNo = 0
          CodeLineNo = 1
          GuardLine = 0
          prevcode = vbNullString
          'Do loop allows CodeLineNo to jump quickly through the code
          Do While CompMod.Find(strFind, CodeLineNo, 1, -1, -1, Not PartialIsOK, True)     ' Then
            If GuardLine > 0 Then
              If GuardLine > CodeLineNo Then
                Exit Do
              End If
            End If
            'found match
            code = CompMod.Lines(CodeLineNo, 1)
            If Not JustACommentOrBlank(code) Then
              If PrevCodeLineNo = CodeLineNo Then 'safety for loop getting jammed
                If prevcode = code Then ' double safety
                  Exit Do
                End If
              End If
              prevcode = code
              PrevCodeLineNo = CodeLineNo
              If PartialIsOK Then
                If InCode(code, TestPos) Then
                  CountCodeUsage = CountCodeUsage + 1
                End If
              End If
              If DeclarationOnly And CodeLineNo > CompMod.CountOfDeclarationLines Then
                'quit if CodeLineNo means that it is not in Declaration
                Exit For
              End If
              If OrigLine <> code Then '*If original line is the code line then ignore it
                TestPos = InStrWholeWordRX(code, strFind)
                Do While TestPos
                  'v2.3.6 Thanks Ulli and 'Array Info' upload
                  ' this deals with problem of code of format
                  ' strA = "ProcName "& ProcName(SomeValue)
                  ' If the proc was always called in this format CF missed it as it ignored the
                  ' string literal but didn't scan on for the second occurance as code
                  If InCode(code, TestPos) Then
                    If InstrAtPosition(OrigLine, "Function", ipLeftOr2ndOr3rd, True) Then
                      If FunctionRefInFunction(Comp, CodeLineNo, strFind) Then
                        Exit Do
                      End If
                    End If
                    If InCode(code, TestPos) Then 'avoid comments and string literals
                      CountCodeUsage = CountCodeUsage + 1
                    End If
                  End If
                  TestPos = InStrWholeWordRX(code, strFind, TestPos + 1)
                Loop
              End If
            End If
            CodeLineNo = CodeLineNo + 1
            GuardLine = CodeLineNo
          Loop
        End If
      End If
SkipTest:
    Next Comp
  Next Proj

End Function

Private Sub DeadControlCode()

  Dim Proj          As VBProject
  Dim Comp          As VBComponent
  Dim CurCompCount  As Long
  Dim arrMembers    As Variant
  Dim ArrProc       As Variant
  Dim I             As Long
  Dim J             As Long
  Dim UpDated       As Boolean
  Dim MUpdated      As Boolean
  Dim MaxFactor     As Long
  Dim Hit           As Boolean
  Dim TopOfRoutine  As Long
  Dim LineContRange As Long
  Dim Rname         As String
  Dim L_CodeLine    As String

  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          'v2.9.5 don't delete control-like stuff in modules/classes
          If IsComponent_ControlHolder(Comp) Then
            If dofix(CurCompCount, DeletedControlCode) Then
              ModuleMessage Comp, CurCompCount
              DisplayCodePane Comp
              arrMembers = GetMembersArray(Comp.CodeModule)
              MaxFactor = UBound(arrMembers)
              UpDated = False
              If MaxFactor > -1 Then
                For I = 1 To MaxFactor
                  MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
                  If Len(arrMembers(I)) Then
                    ArrProc = Split(arrMembers(I), vbNewLine)
                    L_CodeLine = GetRoutineDeclaration(ArrProc, TopOfRoutine, LineContRange, Rname)
                    If InStr(Rname, "_") Then ' has basic format of control code
                      'control with name exists
                      Hit = False
                      Hit = Not LegalControlProcedure(Rname, Comp.Name)
                      If Hit Then
                        Hit = IsControlEvent(strGetRightOf(Rname, "_"))
                      End If
                      If Hit Then
                        Hit = Not RoutineNameIsVBGenerated(Rname, Comp)
                      End If
                      If Hit Then
                        Hit = Not IsUserControlEvent(Rname, Comp)
                      End If
                      'next line takes user defined events and implements code out of detected set
                      DeadControlEventSkip Rname, Hit
                      If Hit Then
                        'test that it is not just something with a control style name
                        Hit = Not FindCodeUsage(Rname, CStr(ArrProc(TopOfRoutine)), Comp.Name, False, False, False)
                      End If
                      'v2.9.5 test that prevents Paul Caton subclass from being commented out
                      If Hit Then
                        If bPaulCatonSubClasUsed Then
                          Hit = Not (Rname = strPaulCatonSubClasProcName)
                        End If
                      End If
                      If Hit Then
                        ParamLineContFix LineContRange, TopOfRoutine, ArrProc, L_CodeLine
                        Select Case FixData(DeletedControlCode).FixLevel
                         Case CommentOnly
                          ArrProc(TopOfRoutine) = Marker(ArrProc(TopOfRoutine), SUGGESTION_MSG & "Unused Control Code should be removed.", MAfter)
                          UpDated = True
                          AddNfix DeletedControlCode
                         Case FixAndComment
                          ArrProc(TopOfRoutine) = Marker(ArrProc(TopOfRoutine), MoveableComment & WARNING_MSG & "Unused Control Code has been commented out.", MAfter, UpDated)
                          ArrProc(TopOfRoutine) = Replace$(ArrProc(TopOfRoutine), MoveableComment & vbNewLine & _
                           WARNING_MSG, MoveableComment & WARNING_MSG)
                          For J = LBound(ArrProc) To UBound(ArrProc)
                            ArrProc(J) = MoveableComment & ArrProc(J)
                          Next J
                          ArrProc(UBound(ArrProc)) = ArrProc(UBound(ArrProc)) & vbNewLine & MoveableComment
                          AddNfix DeletedControlCode
                          'Case JustFix
                        End Select
                      End If
                    End If
                  End If
                  UpdateMember arrMembers(I), ArrProc, UpDated, MUpdated
                Next I
              End If
              If MUpdated Then
                ReWriteMembers Comp.CodeModule, arrMembers, MUpdated
              End If
              If bAborting Then
                Exit For 'Sub
              End If
            End If
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If

End Sub

Private Sub DeadControlEventSkip(ByVal Rname As String, _
                                 Hit As Boolean)

  Dim N As Long

  If Hit Then
    If bEventDescExists Then
      For N = LBound(EventDesc) To UBound(EventDesc)
        If SmartLeft(Rname, EventDesc(N).EName) Then
          Hit = False
          Exit For
        End If
      Next N
    End If
  End If
  If Hit Then
    If ArrayHasContents(ImplementsArray) Then
      On Error Resume Next
      For N = 0 To UBound(ImplementsArray)
        If SmartLeft(Rname, ImplementsArray(N)) Then
          Hit = False
          Exit For ' to turn off the error trap
        End If
      Next N
      On Error GoTo 0
    End If
  End If

End Sub

Public Sub DisplayCodePane(Cmp As VBComponent, _
                           Optional ByVal ForceIt As Boolean = False)

  If Xcheck(XVisScan) Or ForceIt Then
    With Cmp.CodeModule.CodePane
      .Show
      '      If VBInstance.MainWindow.VBE.DisplayModel = vbext_dm_MDI Then
      '        .Window.WindowState = vbext_ws_Maximize
      '      End If
    End With
    Safe_Sleep
  End If

End Sub

Private Sub EmptyCaseCheck(arrR As Variant, _
                           ByVal StructTop As Long, _
                           Update As Boolean)

  Dim I          As Long
  Dim SDeep      As Long
  Dim T_Codeline As String

  If Left$(arrR(StructTop), 11) = "Select Case" Then
    I = StructTop ' - 1
    Do
      'v2.3.3 restructure by moving x = x + 1 to end of loop
      If Left$(arrR(I), 11) = "Select Case" Then
        SDeep = SDeep + 1
      End If
      If Left$(arrR(I), 10) = "End Select" Then
        SDeep = SDeep - 1
        If SDeep = 0 Then
          Exit Do
        End If
      End If
      If SDeep = 1 Then
        If Left$(arrR(I), 5) = "Case " Then
          T_Codeline = NextCodeLine(arrR, I)
          If Left$(T_Codeline, 5) = "Case " Or Left$(T_Codeline, 10) = "End Select" Then
            If EmptyToAvoidDefaultCaseElse(arrR, I) Then
              If InStr(arrR(I), WARNING_MSG & "Empty 'Case") = 0 Then
                If Left$(arrR(I), 9) = "Case Else" Then
                  arrR(I) = Marker(arrR(I), WARNING_MSG & "Empty 'Case Else' structure could be removed", MAfter)
                 Else
                  arrR(I) = Marker(arrR(I), WARNING_MSG & "Empty 'Case' structure used to avoid a default 'Case Else'", MAfter)
                  Update = True
                End If
              End If
             Else
              If InStr(arrR(I), WARNING_MSG & "Empty 'Case") = 0 Then
                If Left$(arrR(I), 9) = "Case Else" Then
                  arrR(I) = Marker(arrR(I), WARNING_MSG & "Empty 'Case Else' structure could be removed", MAfter)
                 Else
                  If EmptyToAvoidDefaultCaseElse(arrR, I) Then
                    arrR(I) = Marker(arrR(I), WARNING_MSG & "Empty 'Case' structure used to avoid a default 'Case Else'", MAfter)
                   Else
                    arrR(I) = Marker(arrR(I), WARNING_MSG & "Empty 'Case' structure could be removed", MAfter)
                  End If
                End If
                Update = True
              End If
            End If
          End If
        End If
      End If
      I = I + 1
    Loop Until I > UBound(arrR)
  End If

End Sub

Private Sub EmptyElseCheck(arrR As Variant, _
                           StructTop As Long, _
                           Update As Boolean)

  Dim I          As Long
  Dim SDeep      As Long
  Dim Msg        As String
  Dim T_Codeline As String
  Dim arrTest1   As Variant
  Dim arrTest2   As Variant

  arrTest1 = Array("If", "Else", "ElseIf")
  arrTest2 = Array("Else", "ElseIf")
  If LeftWord(arrR(StructTop)) = "If" Then
    I = StructTop
    Do
      I = I + 1
      If I > UBound(arrR) Then
        Exit Do
      End If
      If LeftWord(arrR(I)) = "If" Then
        SDeep = SDeep + 1
        If LastCodeWord(arrR(I)) <> "Then" Then
          SDeep = SDeep - 1
        End If
      End If
      If Left$(arrR(I), 6) = "End If" Then
        SDeep = SDeep - 1
        If SDeep = 0 Then
          Exit Do
        End If
      End If
      If IsInArray(LeftWord(arrR(I)), arrTest1) Then
        T_Codeline = NextCodeLine(arrR, I)
        If IsInArray(LeftWord(T_Codeline), arrTest2) Or Left$(T_Codeline, 6) = "End If" Then
          Update = True
          Msg = WARNING_MSG & "Empty" & strInSQuotes(LeftWord(arrR(I)), True)
          If LeftWord(arrR(I)) = "If" Then
            If Left$(T_Codeline, 6) = "End If" Then
              If InStr(arrR(I), Msg) = 0 Then
                arrR(I) = Marker(arrR(I), Msg & "structure could be removed", MAfter)
              End If
             Else
              If InStr(arrR(I), Msg) = 0 Then
                If InStr(arrR(I), "If Not ") Then
                  'v2.5.9 enhanced to cope with the rare empty 'If Not ... Then'
                  'v.2.7.9 stopped duplications and reinserted missing SUGGESTION_MSG
                  If InStr(arrR(I), SUGGESTION_MSG) = 0 Then
                    arrR(I) = Marker(arrR(I), SUGGESTION_MSG & "Empty 'If Not X Then' structure could be Replaced with 'If X Then' and 'Else' removed.", MAfter)
                  End If
                 Else
                  If InStr(arrR(I), SUGGESTION_MSG) = 0 Then
                    arrR(I) = Marker(arrR(I), SUGGESTION_MSG & "Empty 'If X Then' structure could be Replaced with 'If Not X Then' and 'Else' removed.", MAfter)
                  End If
                End If
              End If
            End If
           Else
            If EmptyToAvoidDefaultElse(arrR, I) Then
              If InStr(arrR(I), Msg) = 0 Then
                If Left$(NextCodeLine(arrR, I), 6) = "End If" Then
                  'v2.8.3 there is a bug in EmptyToAvoidDefaultElse that occasionally misfires
                  arrR(I) = Marker(arrR(I), Msg & "could be removed.", MAfter)
                 Else
                  arrR(I) = Marker(arrR(I), Msg & "structure used to avoid a default 'Else'", MAfter)
                End If
              End If
             Else
              If InStr(arrR(I), Msg) = 0 Then
                arrR(I) = Marker(arrR(I), Msg & "structure could be removed", MAfter)
              End If
            End If
          End If
        End If
      End If
    Loop Until I > UBound(arrR)
  End If

End Sub

Private Sub EmptyRoutineTest()

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim arrMembers   As Variant
  Dim arrLine      As Variant
  Dim I            As Long
  Dim M            As Long
  Dim UpDated      As Boolean
  Dim MUpdated     As Boolean
  Dim MaxFactor    As Long
  Dim TopOfRoutine As Long
  Dim lineCount    As Long
  Dim strProcName  As String

  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          If dofix(CurCompCount, EmptyRoutine_CXX) Then
            ModuleMessage Comp, CurCompCount
            ' DisplayCodePane Comp
            arrMembers = GetMembersArray(Comp.CodeModule)
            MaxFactor = UBound(arrMembers)
            If MaxFactor > -1 Then
              For I = 1 To MaxFactor
                MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
                If Len(arrMembers(I)) Then
                  arrLine = Split(arrMembers(I), vbNewLine)
                  UpDated = False
                  strProcName = GetProcNameStr(arrMembers(I))
                  MemberMessage strProcName, I, MaxFactor
                  lineCount = UBound(arrLine) + 1
                  For M = 0 To UBound(arrLine)
                    If JustACommentOrBlank(arrLine(M)) Then
                      lineCount = lineCount - 1
                      If Left$(arrLine(M), 15) = "'<STUB> Reason:" Then
                        lineCount = 3
                        Exit For
                      End If
                     Else
                      If HasLineCont(arrLine(M)) Then
                        lineCount = lineCount - 1
                      End If
                    End If
                  Next M
                  If lineCount <= 2 Then
                    Select Case FixData(EmptyRoutine_CXX).FixLevel
                     Case CommentOnly
                      TopOfRoutine = GetProcCodeLineOfRoutine(arrLine, True)
                      'v3.0.5 catches ISubclass members
                      If LeaveImplementsProcsAlone(strProcName, Comp.Name) Then
                        arrLine(TopOfRoutine) = Marker(arrLine(TopOfRoutine), SUGGESTION_MSG & "Empty Routine. Following comment will stop Code Fixer detecting it next time.", MAfter, UpDated)
                        arrLine(TopOfRoutine) = arrLine(TopOfRoutine) & vbNewLine & _
                         "'<STUB> Reason: Interface procedure used by a class used with 'Implements'"
                       Else
                        arrLine(TopOfRoutine) = Marker(arrLine(TopOfRoutine), SUGGESTION_MSG & "Empty Routine." & vbNewLine & _
                         " If Empty Routine is required(for Add-In Designers for example)" & vbNewLine & _
                         " just add a comment starting with '<STUB> Reason:', fill out any reason you like, and Code Fixer will ignore it in future. ", MAfter, UpDated)
                      End If
                      AddNfix EmptyRoutine_CXX
                      'Case FixAndComment
                      'Case JustFix
                    End Select
                  End If
                End If
                UpdateMember arrMembers(I), arrLine, UpDated, MUpdated
              Next I
            End If
            If MUpdated Then
              ReWriteMembers Comp.CodeModule, arrMembers, MUpdated
            End If
            If bAborting Then
              Exit For 'Sub
            End If
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If

End Sub

Private Sub EmptyStructures()

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim RLine        As Long
  Dim L_CodeLine   As String
  Dim Member       As Long
  Dim ArrMember    As Variant
  Dim ArrRoutine   As Variant
  Dim UpDated      As Boolean
  Dim MUpdated     As Boolean
  Dim MemberCount  As Long
  Dim T_Codeline   As String

  On Error GoTo BugHit
  'ver 2.0.3 new fix
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Identify potential locations for applying With...End With stuctures
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        If dofix(CurCompCount, REmptyStruct) Then
          ArrMember = GetMembersArray(Comp.CodeModule)
          MemberCount = UBound(ArrMember)
          If MemberCount > 0 Then
            For Member = 1 To MemberCount
              MemberMessage GetProcNameStr(ArrMember(Member)), Member, MemberCount
              ArrRoutine = Split(ArrMember(Member), vbNewLine)
              For RLine = LBound(ArrRoutine) To UBound(ArrRoutine)
                L_CodeLine = ArrRoutine(RLine)
                If Not JustACommentOrBlank(L_CodeLine) Then
                  If ArrayMember(LeftWord(L_CodeLine), "With", "If", "Select", "Do") Then
                    EmptyCaseCheck ArrRoutine, RLine, UpDated
                    EmptyElseCheck ArrRoutine, RLine, UpDated
                    If LeftWord(L_CodeLine) = "With" Then
                      T_Codeline = NextCodeLine(ArrRoutine, RLine)
                      If Left$(T_Codeline, 8) = "End With" Then
                        ArrRoutine(RLine) = Marker(ArrRoutine(RLine), WARNING_MSG & "Empty 'With' structure could be removed", MAfter, UpDated)
                      End If
                    End If
                    If LeftWord(L_CodeLine) = "For" Then
                      If LeftWord(NextCodeLine(ArrRoutine, RLine)) = "Next" Then
                        ArrRoutine(RLine) = Marker(ArrRoutine(RLine), WARNING_MSG & "Empty 'For' structure could be removed." & "If being used to create a delay, see 'Better Coding Suggestions' 'Creating a Pause' in the Help file", MAfter, UpDated)
                      End If
                    End If
                    If LeftWord(L_CodeLine) = "Do" Then
                      If LeftWord(NextCodeLine(ArrRoutine, RLine)) = "Loop" Then
                        ArrRoutine(RLine) = Marker(ArrRoutine(RLine), WARNING_MSG & "Empty 'Do' structure could be removed." & "If being used to create a delay, Insert a 'DoEvents' so the machine is not locked while waiting" & vbNewLine & _
                         "See 'Better Coding Suggestions' 'Creating a Pause' in the Help file", MAfter, UpDated)
                      End If
                    End If
                  End If
                End If
              Next RLine
              UpdateMember ArrMember(Member), ArrRoutine, UpDated, MUpdated
            Next Member
            ReWriteMembers Comp.CodeModule, ArrMember, MUpdated
            If bAborting Then
              Exit For 'Sub
            End If
          End If
        End If
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj

Exit Sub

BugHit:
  BugTrapComment "Empty_Structures"
  If Not bAborting Then
    If RunningInIDE Then
      Resume
     Else
      Resume Next
    End If
  End If

End Sub

Public Function EmptyToAvoidDefaultCaseElse(arrR As Variant, _
                                            ByVal RLine As Long) As Boolean

  Dim StructTop As Long
  Dim StructBot As Long
  Dim I         As Long

  For I = RLine To 0 Step -1
    If Left$(arrR(I), 11) = "Select Case" Then
      StructTop = I
      Exit For
    End If
  Next I
  For I = RLine To UBound(arrR)
    If Left$(arrR(I), 10) = "End Select" Then
      StructBot = I
      Exit For
    End If
  Next I
  For I = StructTop To StructBot
    If I <> RLine Then
      If Left$(arrR(I), 9) = "Case Else" Then
        EmptyToAvoidDefaultCaseElse = True
        Exit For
      End If
    End If
  Next I

End Function

Private Function EmptyToAvoidDefaultElse(arrR As Variant, _
                                         RLine As Long) As Boolean

  Dim StructTop As Long
  Dim StructBot As Long
  Dim I         As Long

  If InStructure(IfStruct, arrR, RLine, StructTop, StructBot) Then
    For I = StructBot To StructTop
      If I <> RLine Then
        If LeftWord(arrR(I)) = "Else" Then
          EmptyToAvoidDefaultElse = True
          Exit For
        End If
      End If
    Next I
  End If

End Function

Private Function ExtractVarType(ByVal strCode As String) As String

  Dim I        As Long
  Dim arrTmp   As Variant
  Dim arrTest1 As Variant
  Dim arrTest2 As Variant

  arrTest1 = Array("Public", "Private", "Static", "Friend", "Dim")
  arrTest2 = Array("Type", "Enum", "Declare", "Sub", "Function", "Property", "Event", "WithEvents", "Const")
  strCode = ExpandForDetection(strCode)
  arrTmp = Split(strCode)
  For I = LBound(arrTmp) To UBound(arrTmp)
    If Not IsInArray(arrTmp(I), arrTest1) Then
      ExtractVarType = arrTmp(I)
      If Not IsInArray(ExtractVarType, arrTest2) Then
        ExtractVarType = "Variable"
      End If
      Exit For
    End If
  Next I

End Function

Public Function FindCodeUsage(ByVal strFind As String, _
                              ByVal OrigLine As String, _
                              Optional ByVal CompName As String = vbNullString, _
                              Optional ByVal DeclarationOnly As Boolean = False, _
                              Optional ByVal CurrentOnly As Boolean = False, _
                              Optional ByVal SkipCurrent As Boolean = False, _
                              Optional ByVal PartialIsOK As Boolean = False, _
                              Optional ByVal bIgnoreClassUserDocument As Boolean = False) As Boolean

  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim CompMod        As CodeModule
  Dim code           As String
  Dim prevcode       As String
  Dim TestPos        As Long
  Dim CodeLineNo     As Long
  Dim GuardLine      As Long
  Dim PrevCodeLineNo As Long

  'ver 1.1.1
  'new checks existence of variables
  '(replaces several old routines)
  'uses Find for greater speed
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If CurrentOnly Then
          If CompName <> Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If SkipCurrent Then
          If CompName = Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If bIgnoreClassUserDocument Then
          If Comp.Type = vbext_ct_ClassModule Or Comp.Type = vbext_ct_UserControl Then
            GoTo SkipTest
          End If
        End If
        Set CompMod = Comp.CodeModule
        If CompMod.Find(strFind, 1, 1, -1, -1, Not PartialIsOK, True) Then
          '*if exits at all, then look for the line(s)
          PrevCodeLineNo = 0
          CodeLineNo = 1
          GuardLine = 0
          prevcode = vbNullString
          'Do loop allows CodeLineNo to jump quickly through the code
          Do While CompMod.Find(strFind, CodeLineNo, 1, -1, -1, Not PartialIsOK, True)     ' Then
            If GuardLine > 0 Then
              If GuardLine > CodeLineNo Then
                Exit Do
              End If
            End If
            'found match
            code = strCodeOnly(CompMod.Lines(CodeLineNo, 1))
            If InStr(code, strFind) Then
              If Not JustACommentOrBlank(code) Then
                If PrevCodeLineNo = CodeLineNo Then 'safety for loop getting jammed
                  If prevcode = code Then ' double safety
                    Exit Do
                  End If
                End If
                prevcode = code
                PrevCodeLineNo = CodeLineNo
                If PartialIsOK Then
                  If InCode(code, TestPos) Then
                    FindCodeUsage = True
                    Exit For
                  End If
                End If
                If DeclarationOnly And CodeLineNo > CompMod.CountOfDeclarationLines Then
                  'quit if CodeLineNo means that it is not in Declaration
                  Exit For
                End If
                If OrigLine <> code Then '*If original line is the code line then ignore it
                  TestPos = InStrWholeWordRX(code, strFind)
                  Do While TestPos
                    'v2.3.6 Thanks Ulli and 'Array Info' upload
                    ' this deals with problem of code of format
                    ' strA = "ProcName "& ProcName(SomeValue)
                    ' If the proc was always called in this format CF missed it as it ignored the
                    ' string literal but didn't scan on for the second occurance as code
                    If InCode(code, TestPos) Then
                      'v2.9.2 modified to let Implements function to be recognised
                      If InstrAtPosition(OrigLine, "Function " & strFind, ipLeftOr2ndOr3rd, True) Then
                        If FunctionRefInFunction(Comp, CodeLineNo, strFind) Then
                          Exit Do
                        End If
                      End If
                      If InCode(code, TestPos) Then 'avoid comments and string literals
                        FindCodeUsage = True
                        Exit For 'unction
                      End If
                    End If
                    TestPos = InStrWholeWordRX(code, strFind, TestPos + 1)
                  Loop
                End If
              End If
            End If
            CodeLineNo = CodeLineNo + 1
            GuardLine = CodeLineNo
          Loop
        End If
      End If
SkipTest:
    Next Comp
    If FindCodeUsage Then
      Exit For
    End If
  Next Proj

End Function

Public Function FindCodeUsageLike(ByVal strFind As String, _
                                  Optional ByVal CompName As String = vbNullString, _
                                  Optional ByVal DeclarationOnly As Boolean = False, _
                                  Optional ByVal CurrentOnly As Boolean = False, _
                                  Optional ByVal SkipCurrent As Boolean = False, _
                                  Optional ByVal bIgnoreClassUserDocument As Boolean = False) As Boolean

  'v3.0.6 this is the support routine for catching Types used to define variables filled by using the Get command Thanks Alfred Koppold for bring it to my attention
  'I'm pretty sure this routine will find many more uses in Code Fixer very soon ;)
  
  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim CompMod        As CodeModule
  Dim code           As String
  Dim prevcode       As String
  Dim CodeLineNo     As Long
  Dim GuardLine      As Long
  Dim PrevCodeLineNo As Long

  'ver 1.1.1
  'new checks existence of variables
  '(replaces several old routines)
  'uses Find for greater speed
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If CurrentOnly Then
          If CompName <> Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If SkipCurrent Then
          If CompName = Comp.Name Then
            GoTo SkipTest
          End If
        End If
        If bIgnoreClassUserDocument Then
          If Comp.Type = vbext_ct_ClassModule Or Comp.Type = vbext_ct_UserControl Then
            GoTo SkipTest
          End If
        End If
        Set CompMod = Comp.CodeModule
        If CompMod.Find(strFind, 1, 1, -1, -1, False, False, True) Then
          '*if exits at all, then look for the line(s)
          PrevCodeLineNo = 0
          CodeLineNo = 1
          GuardLine = 0
          prevcode = vbNullString
          'Do loop allows CodeLineNo to jump quickly through the code
          Do While CompMod.Find(strFind, CodeLineNo, 1, -1, -1, False, False, True)    ' Then
            If GuardLine > 0 Then
              If GuardLine > CodeLineNo Then
                Exit Do
              End If
            End If
            'found match
            code = strCodeOnly(CompMod.Lines(CodeLineNo, 1))
            If LenB(code) Then
              If code Like strFind Then ' just incase the match was in a comment
                If PrevCodeLineNo = CodeLineNo Then 'safety for loop getting jammed
                  If prevcode = code Then ' double safety
                    Exit Do
                  End If
                End If
                prevcode = code
                PrevCodeLineNo = CodeLineNo
                FindCodeUsageLike = True
                Exit For
                If DeclarationOnly And CodeLineNo > CompMod.CountOfDeclarationLines Then
                  'quit if CodeLineNo means that it is not in Declaration
                  Exit For
                End If
              End If
            End If
            CodeLineNo = CodeLineNo + 1
            GuardLine = CodeLineNo
          Loop
        End If
      End If
SkipTest:
    Next Comp
    If FindCodeUsageLike Then
      Exit For
    End If
  Next Proj

End Function

Private Function FunctionRefInFunction(Cmp As VBComponent, _
                                       ByVal CurrentLine As Long, _
                                       ByVal strTestWord As String) As Boolean

  'used to ignore self reference in functions
  'v 2.0.5 fixed

  On Error GoTo oops
  If CurrentLine <> 0 Then
    If Cmp.CodeModule.ProcOfLine(CurrentLine, vbext_pk_Proc) = strTestWord Then
      FunctionRefInFunction = True
    End If
  End If
oops:

End Function

Private Function GetCodeArray(cMod As CodeModule) As Variant

  Dim Tmpstring1 As String
  Dim I          As Long
  Dim iStart     As Long

  'Gets code (without Declarations) each member is one line of code
  With cMod
    iStart = .CountOfDeclarationLines + 1
    'skip over end of declaration code
    Do While SmartLeft(.Lines(iStart, 1), "#End If") Or Len(Trim$(.Lines(iStart, 1))) = 0
      iStart = iStart + 1
      If iStart > .CountOfLines Then
        Exit Do
      End If
    Loop
    For I = iStart To .CountOfLines
      If LenB(Trim$(.Lines(I, 1))) Then
        Tmpstring1 = Tmpstring1 & Trim$(.Lines(I, 1)) & vbNewLine
      End If
    Next I
  End With
  If LenB(Tmpstring1) Then
    GetCodeArray = Split(Left$(Tmpstring1, Len(Tmpstring1) - Len(vbNewLine)), vbNewLine)
   Else
    GetCodeArray = Split("")
  End If
  Tmpstring1 = vbNullString

End Function

Private Sub GetTestName(varA As Variant, _
                        strName As String, _
                        StrTyp As String)

  strName = ExtractName(Join(varA))
  StrTyp = ExtractVarType(Join(varA)) 'varA(1)
  Select Case StrTyp
   Case "Const"
    strName = IIf(FixData(UnusedDecConst).FixLevel > Off, strName, vbNullString)
   Case "Declare"
    strName = IIf(FixData(UnusedDecAPI).FixLevel > Off, strName, vbNullString)
   Case "Sub"
    strName = IIf(FixData(UnusedSub).FixLevel > Off, strName, vbNullString)
   Case "Function"
    strName = IIf(FixData(UnusedFunction).FixLevel > Off, strName, vbNullString)
   Case "Property"
    strName = IIf(FixData(UnusedProperty).FixLevel > Off, strName, vbNullString)
    '        these are always public
   Case "Events"
    strName = IIf(FixData(UnusedEvents).FixLevel > Off, strName, vbNullString)
    'Case "WithEvents"
    '        strName = IIf(FixData(ModuleNumber, UnusedWithEvents).FixLevel > Off, strName, "")
   Case "Variable" 'Else
    strName = IIf(FixData(UnusedDecVariable).FixLevel > Off, strName, vbNullString)
  End Select

End Sub

Public Function InEnumCapProtection(cMod As CodeModule, _
                                    ByVal arrTmp As Variant, _
                                    ByVal LineNo As Long) As Boolean

  Dim LIndex       As Long
  Dim StartECPLine As Long
  Dim Possible     As Boolean
  Dim L_CodeLine   As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Protect Enum Capitalisation Protection from Declaration Formatter
  'by detecting that line is inside an Enum Capitalisation Protection structure
  With cMod
    LIndex = 1
    Do
      L_CodeLine = arrTmp(LIndex)
      If SmartLeft(L_CodeLine, Hash_If_False_Then) Then
        Possible = True
        StartECPLine = LIndex
      End If
      If SmartLeft(L_CodeLine, Hash_End_If) Then
        If Possible Then
          If LineNo > StartECPLine Then
            If LineNo < LIndex Then
              InEnumCapProtection = True
              Exit Do
             Else
              Possible = False
            End If
          End If
        End If
      End If
      LIndex = LIndex + 1
      If LIndex > LineNo Then
        If Not Possible Then
          Exit Do
        End If
      End If
    Loop Until LIndex > UBound(arrTmp)
  End With

End Function

Private Function IsANewEnum(ByVal TmpA As Variant, _
                            ByVal TestName As String) As Boolean

  Dim I        As Long
  Dim J        As Long
  Dim K        As Long
  Dim TmpD     As Variant
  Dim TestLine As String

  '"[_NewEnum]" is a special property of collection classes
  'which is never called from anywhere in program but acts like a compile directive
  'This routine protects it from being detected as unused by CodeFixer
  'by testing for the one absolutly essential element of such I property
  For I = LBound(TmpA) + 1 To UBound(TmpA)
    If InStr(TmpA(I), TestName) Then
      If InStr(TmpA(I), "[_NewEnum]") Then
        TmpD = Split(TmpA(I), vbNewLine)
        For J = LBound(TmpD) To UBound(TmpD)
          TestLine = TmpD(J)
          If ExtractCode(TestLine) Then
            TestLine = ExpandForDetection(TestLine)
            If InStr(TestLine, TestName) Then
              For K = J To UBound(TmpD)
                TestLine = TmpD(K)
                If InstrAtPosition(TestLine, "[_NewEnum]", ipAny, False) Then
                  IsANewEnum = True
                  Exit For
                End If
              Next K
            End If
          End If
        Next J
      End If
    End If
    If IsANewEnum Then
      Exit For
    End If
  Next I

End Function

Public Function isCodeFixProtected(ByVal TestName As String) As Boolean

  Dim I As Long

  For I = LBound(CodeFixProtectedArray) To UBound(CodeFixProtectedArray)
    If SmartLeft(CodeFixProtectedArray(I), "Private " & TestName) Then
      isCodeFixProtected = True
      Exit For
    End If
  Next I

End Function

Private Function IsComponent_UserControl_Class(ByVal c As VBComponent) As Boolean

  IsComponent_UserControl_Class = c.Type = vbext_ct_UserControl Or c.Type = vbext_ct_ClassModule

End Function

Private Function IsInParameter(TmpA As Variant, _
                               ByVal strTest As String, _
                               ByVal Skipline As Long) As Boolean

  Dim I        As Long
  Dim J        As Long
  Dim TmpD     As Variant
  Dim TestLine As String

  For I = LBound(TmpA) + 1 To UBound(TmpA)
    If Skipline <> I Then
      If InStr(TmpA(I), strTest) Then
        TmpD = Split(TmpA(I), vbNewLine)
        For J = LBound(TmpD) To UBound(TmpD)
          TestLine = TmpD(J)
          If ExtractCode(TestLine) Then
            TestLine = ExpandForDetection(TestLine)
            If InstrAtPosition(TestLine, strTest, ipAny) Then
              If EnclosedInBrackets(TestLine, InStr(TestLine, strTest)) Or MultiRight(TestLine, True, "As " & strTest) Then
                IsInParameter = True
                Exit For
              End If
            End If
          End If
        Next J
      End If
    End If
    If IsInParameter Then
      Exit For
    End If
  Next I

End Function

Public Function IsRealWord(ByVal strCode As String, _
                           ByVal strWord As String) As Boolean

  Dim strLeft1  As String
  Dim strRight1 As String
  Dim TestPos   As Long
  Dim arrTest1  As Variant
  Dim arrTest2  As Variant

  arrTest1 = Array("(", "!", ".")
  arrTest2 = Array(",", ")", "(", "!", ".", ":")
  '  because of the way CF extract words it cannot tell the difference between
  '  word, _word _word_ and word_. 'The advantage is that word. and .word are found
  '  This routine helps it differentiate when necessary
  '
  'v2.3.6 added to avoid if variable is only in comment
  ExtractCode strCode
  strCode = ExpandForDetection(strCode)
  TestPos = InStr(strCode, strWord)
  Do While TestPos
    '  v2.3.6 added to cope with 2 variable with same spelling on the left
    '  Const LB_FINDSTRINGEXACT & LB_FINDSTRING
    'the do loop findes the 2nd in code where it is used after 1st (See PosInList for example)
    If InCode(strCode, TestPos) Then
      If TestPos > 1 Then
        strLeft1 = Mid$(strCode, TestPos - 1, 1)
        If IsPunct(strLeft1) Then
          strLeft1 = Mid$(strCode, TestPos - 1, 1)
          If strLeft1 <> " " Then
            If IsPunct(strLeft1) Then
              If TestPos < Len(strCode) - 1 Then
                strRight1 = Mid$(strCode, TestPos + Len(strWord), 1)
              End If
              If Not IsInArray(strRight1, arrTest1) Then
                '  v2.1.8 added test legal attached to start of a word in VB
                GoTo skipit
              End If
            End If
          End If
         ElseIf IsAlphaIntl(strLeft1) Then
          'v2.1.8 added test because it was passing if strWord was left or right end of larger words
          'v2.2.8 fixed If stRucutre so that it hit properly
          GoTo skipit
        End If
      End If
      If TestPos < Len(strCode) - 1 Then
        strRight1 = Mid$(strCode, TestPos + Len(strWord), 1)
        If strRight1 <> " " Then
          If IsPunct(strRight1) Then
            'v 2.7.7 added colon detection for poorly named goto target labels
            If Not IsInArray(strRight1, arrTest2) Then
              'v2.1.8 added test legal attached to end of a word in VB
              GoTo skipit
            End If
           ElseIf IsAlphaIntl(strRight1) Then
            GoTo skipit
          End If
        End If
      End If
      IsRealWord = True
      Exit Do
    End If
skipit:
    TestPos = InStr(TestPos + 1, strCode, strWord)
  Loop

End Function

Public Function IsUserControlEvent(ByVal Rname As String, _
                                   ByVal Comp As VBComponent) As Boolean

  Dim I As Long

  If Comp.Type = vbext_ct_UserControl Then
    If Left$(Rname, 11) = "UserControl" Then
      For I = LBound(UserCtrlEventArray) To UBound(UserCtrlEventArray)
        If SmartRight(Rname, UserCtrlEventArray(I)) Then
          IsUserControlEvent = True
          Exit For
        End If
      Next I
    End If
  End If

End Function

Private Function LastCodeWord(ByVal varChop As Variant) As String

  Dim TmpA As Variant

  If ExtractCode(varChop) Then
    TmpA = Split(varChop)
    LastCodeWord = TmpA(UBound(TmpA))
  End If

End Function

Private Function LeaveImplementsProcsAlone(ByVal strTest As String, _
                                           strCmp As String) As Boolean

  Dim strTestMe As String

  'v2.9.2 Thanks Bazz deals with Implements procs which are not called directly in code but used in subclassing
  'test 1 are we checking in the module?
  LeaveImplementsProcsAlone = InQSortArray(ImplementsArray, strCmp)
  If Not LeaveImplementsProcsAlone Then
    strTestMe = Split(strTest, "_")(0)
    'test 2 out of module calls to the Implements module
    If InQSortArray(ImplementsArray, strTestMe) Then
      'leave Implements using Procedures alone
      LeaveImplementsProcsAlone = True
      'FindCodeUsage(Mid$(strTest, Len(strTestMe) + 2), varCode, strTestMe, , True)
    End If
  End If

End Function

Public Function LegalControlProcedure(strName As String, _
                                      strCompName As String) As Boolean

  Dim strL As String

  'test if the left side is a control and the right an event
  If CountSubString(strName, "_") = 1 Then
    strL = strGetLeftOf(strName, "_")
    If CntrlDescMember(strL) <> -1 Then
      If CntrlDesc(CntrlDescMember(strL)).CDForm = strCompName Then
        LegalControlProcedure = IsControlEvent(strGetRightOf(strName, "_"))
      End If
    End If
  End If

End Function

Public Function NextCodeLine(Arr As Variant, _
                             ByVal Pos As Long, _
                             Optional ByVal LinesForward As Long = 1, _
                             Optional NextPos As Long) As String

  Dim LCount As Long

  If Pos < UBound(Arr) Then
    Do
      Do
        Pos = Pos + 1
        If Pos = UBound(Arr) Then
          Exit Do
        End If
      Loop While JustACommentOrBlank(Arr(Pos))
      LCount = LCount + 1
      If LinesForward = LCount Then
        NextCodeLine = Arr(Pos)
        NextPos = Pos
        Exit Do
      End If
      If Pos = UBound(Arr) Then
        Exit Do
      End If
    Loop
  End If

End Function

Private Function PartialCntrlDescMember(ByVal varTest As Variant, _
                                        ByVal LeftTRightF As Boolean, _
                                        ByVal strCmpName As String) As Boolean

  'v3.0.0 added strCmpName for greater accuracy
  
  Dim I           As Long

  If bCtrlDescExists Then
    If CntrlDescMember(varTest) > -1 Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If LeftTRightF Then
          If SmartLeft(varTest, CntrlDesc(I).CDName) Then
            PartialCntrlDescMember = True
            Exit For 'unction
          End If
         Else
          If SmartRight(varTest, CntrlDesc(I).CDName) Then
            PartialCntrlDescMember = True
            Exit For 'unction
          End If
        End If
      Next I
    End If
    If Not PartialCntrlDescMember Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If CntrlDesc(I).CDForm = strCmpName Then
          If LeftTRightF Then
            'v2.9.8 improved test reason some unused were not being eliminated
            If SmartLeft(varTest, CntrlDesc(I).CDName & "_", False) Then
              PartialCntrlDescMember = True
              varTest = Replace$(varTest, Left$(varTest, Len(CntrlDesc(I).CDName)), CntrlDesc(I).CDName)
              Exit For 'unction
            End If
           Else
            If SmartRight(varTest, CntrlDesc(I).CDName) Then
              PartialCntrlDescMember = True
              varTest = Replace$(varTest, Right$(varTest, Len(CntrlDesc(I).CDName)), CntrlDesc(I).CDName)
              Exit For 'unction
            End If
          End If
        End If
      Next I
    End If
  End If

End Function

Private Function PartialEventDescMember(ByVal varTest As Variant, _
                                        ByVal LeftTRightF As Boolean, _
                                        Optional ByVal StrScope As String, _
                                        Optional ByVal strForm As String) As Boolean

  Dim I           As Long
  Dim strExtScope As String

  If StrScope = "Public" Then
    strExtScope = "Friend"
   ElseIf StrScope = "Private" Then
    strExtScope = "Static"
  End If
  If bEventDescExists Then
    For I = LBound(EventDesc) To UBound(EventDesc)
      If LenB(strForm) = 0 Or EventDesc(I).EForm = strForm Then
        If LenB(StrScope) = 0 Or EventDesc(I).EScope = StrScope Or EventDesc(I).EScope = strExtScope Then
          If LeftTRightF Then
            If SmartLeft(varTest, EventDesc(I).EName) Then
              PartialEventDescMember = True
              Exit For 'unction
            End If
          End If
         Else
          If SmartRight(varTest, EventDesc(I).EName) Then
            PartialEventDescMember = True
            Exit For 'unction
          End If
        End If
      End If
    Next I
  End If

End Function

Private Function ProtectUnusedClassOrControl(Cmp As VBComponent) As Boolean

  If FixData(CommentOutClassmembers).FixLevel = Off Then
    If IsComponent_UserControl_Class(Cmp) Then
      ProtectUnusedClassOrControl = True
    End If
  End If

End Function

Private Function RoutineIsCalledByCallByName(ByVal strTest As String, _
                                             Optional ByVal CurrentModuleOnly As Boolean = False, _
                                             Optional ByVal IgnoreCurrentModule As Boolean = False, _
                                             Optional ByVal strCurCompName As String) As Boolean

  Dim Comp      As VBComponent
  Dim Proj      As VBProject
  Dim TestLine  As String
  Dim tmpCodArr As Variant
  Dim I         As Long
  Dim MaxFactor As Long
  Dim TPos      As Long

  'Detect if a routine is called only by call CallByName method
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If ModuleHasCode(Comp.CodeModule) Then
          If CurrentModuleOnly Then
            If Comp.Name <> strCurCompName Then
              GoTo SkipComp
            End If
          End If
          If IgnoreCurrentModule Then
            If Comp.Name = strCurCompName Then
              GoTo SkipComp
            End If
          End If
          tmpCodArr = GetCodeArray(Comp.CodeModule)
          If InStr(Join(tmpCodArr, vbNewLine), strTest) Then
            MaxFactor = UBound(tmpCodArr)
            If MaxFactor > -1 Then
              For I = 0 To MaxFactor
                TestLine = tmpCodArr(I)
                If Not JustACommentOrBlank(TestLine) Then
                  TestLine = ExpandForDetection(TestLine)
                  TPos = InStr(TestLine, strTest)
                  If TPos Then
                    If InLiteral(TestLine, TPos) Then
                      If InstrAtPosition(TestLine, "CallByName", ipAny, True) Then
                        RoutineIsCalledByCallByName = True
                      End If
                      Exit For
                    End If
                  End If
                End If
              Next I
            End If
          End If
        End If
      End If
      If RoutineIsCalledByCallByName Then
        Exit For
      End If
SkipComp:
    Next Comp
    If RoutineIsCalledByCallByName Then
      Exit For
    End If
  Next Proj

End Function

Private Function RoutineNameIsEvent(ByVal strTest As String, _
                                    ByVal Comp As VBComponent) As Boolean

  'All Controls generate automatic subs
  'This routine test if first half of routine name
  'is controlname followed by underscore
  'This stops Unused_Private and Unused_Public detecting these Subs as unused.
  'These are the endings of the standard VB created functions for the standard controls

  strTest = Trim$(strTest)
  If isEvent(strTest, "Public") Then
    If PartialEventDescMember(strTest, True, "Public") Then
      RoutineNameIsEvent = True
     ElseIf PartialEventDescMember(strTest, True, "Private") Then
      RoutineNameIsEvent = True
     ElseIf PartialEventDescMember(strTest, True, "Static", Comp.Name) Then
      RoutineNameIsEvent = True
     Else
      If PartialEventDescMember(strTest, True, "Friend") Then
        RoutineNameIsEvent = True
      End If
    End If
   Else
    RoutineNameIsEvent = False
  End If

End Function

Public Function RoutineNameIsGenericToModule(ByVal strTest As String, _
                                             ByVal Comp As VBComponent) As Boolean

  'All Controls generate automatic subs
  'This routine test if first half of routine name
  'is controlname followed by underscore
  'This stops Unused_Private and Unused_Public detecting these Subs as unused.
  'These are the endings of the standard VB created functions for the standard controls

  If ArrayMember(Comp.Type, vbext_ct_MSForm, vbext_ct_UserControl, vbext_ct_VBMDIForm, vbext_ct_VBForm) Then
    RoutineNameIsGenericToModule = MultiLeft(Trim$(strTest), False, "Form_", "UserControl_", "MDIForm_")
  End If

End Function

Public Function RoutineNameIsInternalToModule(ByVal strTest As String, _
                                              ByVal Comp As VBComponent) As Boolean

  'All Controls generate automatic subs
  'This routine test if first half of routine name
  'is controlname followed by underscore
  'This stops Unused_Private and Unused_Public detecting these Subs as unused.
  'These are the endings of the standard VB created functions for the standard controls

  strTest = Trim$(strTest)
  If IsComponent_ControlHolder(Comp) Then
    '.Type = vbext_ct_ActiveXDesigner Or Comp.Type = vbext_ct_ClassModule Then
    'these are standard end of routine names for designers and classes( rarely called explicitly but not unused)
    If RoutineNameIsGenericToModule(strTest, Comp) Or RoutineNameIsVBGenerated(strTest, Comp) Then
      RoutineNameIsInternalToModule = Not RoutineNameIsUserInteractive(strTest, Comp)
    End If
  End If

End Function

Public Function RoutineNameIsUserInteractive(ByVal strTest As String, _
                                             ByVal Comp As VBComponent) As Boolean

  'All Controls generate automatic subs
  'This routine test if first half of routine name
  'is controlname followed by underscore
  'This stops Unused_Private and Unused_Public detecting these Subs as unused.
  'These are the endings of the standard VB created functions for the standard controls

  strTest = Trim$(strTest)
  If IsComponent_ControlHolder(Comp) Then
    'these are standard end of routine names for designers and classes( rarely called explicitly but not unused)
    If RoutineNameIsGenericToModule(strTest, Comp) Or RoutineNameIsVBGenerated(strTest, Comp) Then
      RoutineNameIsUserInteractive = MultiRight(strTest, False, "_Click", "_DblClick", "_MouseUp", "_MouseDown", "_MouseMove", "_KeyUp", "_KeyDown", "_KeyPress", "_DragDrop", "_DragOver", "_OLECompleteDrag", "_OLEDragDrop", "_OLEStartDrag")
    End If
  End If

End Function

Public Function RoutineNameIsVBGenerated(ByVal strTest As String, _
                                         ByVal Comp As VBComponent, _
                                         Optional ByVal bSoft As Boolean = True) As Boolean

  'All Controls generate automatic subs
  'This routine test if first half of routine name
  'is controlname followed by underscore
  'This stops Unused_Private and Unused_Public detecting these Subs as unused.
  'These are the endings of the standard VB created functions for the standard controls

  strTest = Trim$(strTest)
  If Comp.Type = vbext_ct_ActiveXDesigner Or Comp.Type = vbext_ct_ClassModule Then
    'these are standard end of routine names for designers and classes( rarely called explicitly but not unused)
    RoutineNameIsVBGenerated = MultiRight(strTest, False, "_Initialize", "_OnBeginShutdown", "_OnConnection", "_OnDisconnection", "_OnStartupComplete", "_OnAddinsUpdate", "_Terminate", "_Resize")
  End If
  If Not RoutineNameIsVBGenerated Then
    If MultiRight(strTest, False, "_Initialize", "_OnBeginShutdown", "_OnConnection", "_OnDisconnection", "_OnStartupComplete", "_OnAddinsUpdate", "_Terminate", "_Resize", "_WriteProperties", "_ReadProperties") Then
      'these are just standard/default routines Code Fixer assumes are real
      RoutineNameIsVBGenerated = True
     ElseIf MultiLeft(strTest, False, "Main", "Form_", "UserControl_", "MDIForm_") Then
      RoutineNameIsVBGenerated = True
     ElseIf bSoft And PartialCntrlDescMember(strTest, True, Comp.Name) Then
      RoutineNameIsVBGenerated = True
     ElseIf bSoft And PartialEventDescMember(strTest, True, "Private", Comp.Name) Then
      RoutineNameIsVBGenerated = True
     ElseIf IsDeclareName(strTest) Then 'IsDeclaration(strTest) Then
      RoutineNameIsVBGenerated = True
     Else
      RoutineNameIsVBGenerated = False
    End If
  End If

End Function

Public Function SafeCompToProcess(ByVal Cmp As VBComponent, _
                                  CmpCounter As Long, _
                                  Optional ByVal DontTouchTest As Boolean = True) As Boolean

  'returns True if the component is anything that can/should be processed by program
  'test that the component is one you can edit at all
  'SafeCompToProcess = cmp.Type <> vbext_ct_ResFile And cmp.Type <> vbext_ct_RelatedDocument

  SafeCompToProcess = LenB(Cmp.Name)
  ' this routine is called at start of all rewrite code so
  ' this is a good spot to make sure that suspend is off
  'the counter has to increase whether or not it is true
  '*If the call uses 'Dummy' to call function then the number
  'is not needed for that routine it is just thrown away
  'only test these if first test is pased
  If SafeCompToProcess Then
    CmpCounter = CmpCounter + 1
    If DontTouchTest Then
      If bModDescExists Then
        If ModDesc(CmpCounter).MDDontTouch Then
          'causes routine to skip processing section of code
          SafeCompToProcess = False
        End If
      End If
    End If
  End If
  Safe_Sleep

End Function

Private Sub ScopeConst(testTarget As Variant, _
                       ByVal TestName As String, _
                       ByVal CurLine As Long, _
                       bUsed As Boolean, _
                       bPrivOnly As Boolean)

  If testTarget = "Const" Then
    bUsed = FindCodeUsage(TestName, vbNullString, vbNullString, True, False, True)
    If Not bUsed Then
      bUsed = FindCodeUsage(TestName, vbNullString, vbNullString, True, True, False)
      If bUsed Then
        If MultiLeft(SearchDeclarationsStrMatch(TestName, True, False, CurLine), True, "Private ") Then
          bPrivOnly = True
        End If
      End If
    End If
  End If

End Sub

Private Sub ScopeTooLargeMsg(varTyp As Variant, _
                             ByVal strName As String, _
                             VarArray As Variant, _
                             ByVal CurLine As Long)

  Dim MsgPos            As Long

  'This changes codefixer comment attachment point if necessary.
  'Due to possible line continuation characters the correct spot
  'for the comment may not be just after the current line
  'although it usually will be.
  MsgPos = CurLine
  ' because this treats variables as well as named structures (and varTyp of variables is variable name)
  Select Case FixData(UnusedFixType(varTyp)).FixLevel
   Case CommentOnly
    SafeInsertArrayMarker VarArray, MsgPos, WARNING_MSG & "Scope Too Large " & IIf(varTyp <> strName, varTyp & SngSpace & strName, strName) & " is only ever called from this module, Change Scope to Private." & IIf(Xcheck(XVerbose), vbNewLine & _
     RGSignature & "If it was not auto-created by a non-standard VB control you can safely delete it." & vbNewLine & _
     RGSignature & "Check that it is not a prototype you have not yet implimented and should not be made Public for other modules in your code.", vbNullString)
   Case FixAndComment
    VarArray(CurLine) = "Private " & Mid$(VarArray(CurLine), 7)
    SafeInsertArrayMarker VarArray, MsgPos, WARNING_MSG & "Scope Too Large. Reduced to Private " & IIf(Xcheck(XVerbose), RGSignature & "May be a prototype you have not yet implimented or left over from a deleted Control.", vbNullString)
    'Case JustFix
    'Much too dangerous to activate this
  End Select

End Sub

Private Sub ScopeTypeEnum(cMod As CodeModule, _
                          testTarget As Variant, _
                          ByVal TestName As String, _
                          arrCurModule As Variant, _
                          ByVal CurLineMember As Long, _
                          bUsed As Boolean, _
                          bPrivOnly As Boolean)

  Dim MemberNo      As Long
  Dim StrMemberTest As String

  If ArrayMember(testTarget, "Type", "Enum") Then
    bUsed = FindCodeUsage(TestName, vbNullString, vbNullString, True, False, True)
    If Not bUsed Then
      bUsed = FindCodeUsage(TestName, vbNullString, vbNullString, True, True, False)
      If bUsed Then
        If MultiLeft(SearchDeclarationsStrMatch(TestName, True, False, CurLineMember), True, "Private ") Then
          bPrivOnly = True
        End If
      End If
    End If
    If Not bUsed Then
      MemberNo = CurLineMember
      Do
        MemberNo = MemberNo + 1
        If MultiLeft(arrCurModule(MemberNo), True, "End Enum", "End Type") Then
          Exit Do
        End If
        If MemberNo = UBound(arrCurModule) Then
          Exit Do
        End If
        StrMemberTest = LeftWord(arrCurModule(MemberNo))
        If LenB(StrMemberTest) Then
          bUsed = FindCodeUsage(StrMemberTest, CStr(arrCurModule(MemberNo)), cMod.Name, False, False, False)
          If bUsed Then
            Exit Do
          End If
          bUsed = FindCodeUsage(StrMemberTest, CStr(arrCurModule(MemberNo)), cMod.Name, True, False, False)
          If bUsed Then
            Exit Do
          End If
        End If
      Loop
    End If
  End If

End Sub

Private Function SearchDeclarationsStrMatch(ByVal strRoutineName As String, _
                                            Optional ByVal CurrentModuleOnly As Boolean = False, _
                                            Optional ByVal IgnoreCurrentModule As Boolean = False, _
                                            Optional ByVal IgnoreSelf As Long = -1) As String

  Dim Comp      As VBComponent
  Dim Proj      As VBProject
  Dim tmpCodArr As Variant
  Dim I         As Long

  'Collect and record names and types of Functions
  'For MissingDim, UntypedDim  and UnusedFunction tests
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If ModuleHasCode(Comp.CodeModule) Then
          If Comp.Name <> CurrentModuleOnly Then
            GoTo SkipComp
          End If
        End If
        If IgnoreCurrentModule Then
          If Comp.Name = CurrentModuleOnly Then
            GoTo SkipComp
          End If
        End If
        tmpCodArr = GetDeclarationArray(Comp.CodeModule) 'AllModDecArray(CurCompCount)
        For I = LBound(tmpCodArr) To UBound(tmpCodArr)
          If IgnoreSelf = -1 Then
            If InstrAtPosition(tmpCodArr(I), strRoutineName, ipAny) Then
              SearchDeclarationsStrMatch = tmpCodArr(I)
              Exit For
            End If
           Else
            If IgnoreSelf <> I Then
              If InstrAtPosition(tmpCodArr(I), strRoutineName, ipAny) Then
                SearchDeclarationsStrMatch = tmpCodArr(I)
                Exit For
              End If
            End If
          End If
        Next I
      End If
      If LenB(SearchDeclarationsStrMatch) Then
        Exit For
      End If
SkipComp:
    Next Comp
    If LenB(SearchDeclarationsStrMatch) Then
      Exit For
    End If
  Next Proj

End Function

Public Sub Unused_Engine()

  Dim NumFixes As Long

  On Error GoTo BugHit
  NumFixes = IIf(Xcheck(XUsageComments), 14, 6)
  If Not bAborting Then
    WorkingMessage "Empty Structure", 1, NumFixes
    EmptyStructures ' placed here so that this can run on unsaved/ no project code
    If LenB(GetActiveProject.FileName) Then
      ArrQCheckedVariables = Split("")
      'clear the array generated in previous run
      'call by name changes the way a routine can be detected as unused
      CallByNameIsUsed = FindCodeUsage("CallByName", vbNullString, vbNullString, False, False, False)
      WorkingMessage "Empty Routines", 2, NumFixes
      EmptyRoutineTest
      WorkingMessage "Unused Controls Code", 3, NumFixes
      DeadControlCode
      WorkingMessage "Unused Public", 4, NumFixes
      Unused_Public
      WorkingMessage "Unused Private", 5, NumFixes
      Unused_Private
      WorkingMessage "Active Debug", 6, NumFixes
      ActiveDebugDriver
      If Xcheck(XUsageComments) Then
        WorkingMessage "Public Declaration Usage", 7, NumFixes
        UsageCountDelarations "Public"
        WorkingMessage "Private Declaration Usage", 8, NumFixes
        UsageCountDelarations "Private"
        WorkingMessage "Static Declaration Usage", 9, NumFixes
        UsageCountDelarations "Static"
        WorkingMessage "Friend Declaration Usage", 10, NumFixes
        UsageCountDelarations "Friend"
        WorkingMessage "Public Procedure Usage", 11, NumFixes
        UsageCountProcedures "Public"
        WorkingMessage "Private Procedure Usage", 12, NumFixes
        UsageCountProcedures "Private"
        WorkingMessage "Static Procedure Usage", 13, NumFixes
        UsageCountProcedures "Static"
        WorkingMessage "Friend Procedure Usage", 14, NumFixes
        UsageCountProcedures "Friend"
      End If
    End If
  End If
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "Unused_Engine"
  If Not bAborting Then
    If RunningInIDE Then
      Resume
     Else
      Resume Next
    End If
  End If

End Sub

Private Sub Unused_Private()

  
  Dim N            As Long
  Dim M            As Long
  Dim CurCompCount As Long
  Dim TestName     As String
  Dim TestType     As String
  Dim arrMembers   As Variant
  Dim ArrProc      As Variant
  Dim arrLine      As Variant
  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim UpDated      As Boolean
  Dim isUsed       As Boolean
  Dim I            As Long
  Dim J            As Long
  Dim K            As Long
  Dim TmpD         As Variant
  Dim bJunk        As Boolean
  Dim arrTest      As Variant

  arrTest = Array("Type", "Enum", "Const")
  On Error GoTo BugTrap
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          arrMembers = GetMembersArray(Comp.CodeModule)
          ArrQCheckedVariables = Array()
          If UBound(arrMembers) > -1 Then
            For I = 0 To UBound(arrMembers)
              MemberMessage GetProcNameStr(arrMembers(I)), I, UBound(arrMembers)
              ArrProc = Split(arrMembers(I), vbNewLine)
              For M = LBound(ArrProc) To UBound(ArrProc)
                If I <> 0 Then
                  'v2.4.9 don't go line to line if not Declarations
                  M = GetProcCodeLineOfRoutine(ArrProc)
                End If
                isUsed = False
                'v2.5.1 error next line tested for 'Public' Thanks Mike Wardle
                If LeftWord(ArrProc(M)) = "Private" Then
                  arrLine = Split(ExpandForDetection(ArrProc(M)))
                  If UBound(arrLine) > 2 Then
                    GetTestName arrLine, TestName, TestType
                    MemberMessage TestName, M, UBound(ArrProc)
                    If LenB(TestName) Then
                      'v2.9.5 test that prevents Paul Caton subclass from being commented out
                      If bPaulCatonSubClasUsed Then
                        isUsed = TestName = strPaulCatonSubClasProcName
                      End If
                      If Not isUsed Then
                        If LeaveImplementsProcsAlone(TestName, Comp.Name) Then
                          isUsed = True 'leave Implements classes alone
                        End If
                      End If
                      If Not isUsed Then
                        UnusedWithEventsTest TestType, TestName, ArrProc(M), Comp.Name, isUsed, bJunk
                      End If
                      If Not isUsed Then
                        UnusedEventsTest TestType, TestName, ArrProc(M), Comp.Name, isUsed, bJunk
                      End If
                      If Not isUsed Then
                        If Not InQSortArray(ArrQCheckedVariables, TestName) Then
                          UsedTests TestType, TestName, Comp, isUsed
                          UsedTestsPrivate TestName, Comp, ArrProc, M, isUsed
                          If Not isUsed Then
                            For J = LBound(arrMembers) + 1 To UBound(arrMembers)
                              If I <> J Then
                                If InStr(arrMembers(J), TestName) Then
                                  TmpD = Split(arrMembers(J), vbNewLine)
                                  For K = LBound(TmpD) To UBound(TmpD)
                                    If InstrAtPosition(ExpandForDetection(TmpD(K)), TestName, ipAny) Then
                                      isUsed = True
                                      Exit For
                                    End If
                                  Next K
                                End If
                              End If
                              If isUsed Then
                                Exit For
                              End If
                            Next J
                          End If
                          If Not isUsed Then
                            If IsInArray(TestType, arrTest) Then
                              isUsed = FindCodeUsage(TestName, ArrProc(M), vbNullString, True, False, False)
                              'v2.8.3
                              If Not isUsed Then
                                If TestType = "Enum" Then
                                  For N = M To UBound(ArrProc)
                                    If LeftWord(ArrProc(N)) = "End" Then
                                      Exit For
                                    End If
                                    isUsed = FindCodeUsage(LeftWord(ArrProc(N)), ArrProc(N), vbNullString, True, False, False)
                                    If isUsed Then
                                      Exit For
                                    End If
                                  Next N
                                End If
                              End If
                            End If
                          End If
                          If Not isUsed Then
                            UnusedMsg TestType, TestName, ArrProc, M, Comp
                            arrMembers(I) = Join(ArrProc, vbNewLine)
                            UpDated = True
                            UnusedRecord TestType
                            ArrQCheckedVariables = QuickSortArray(AppendArray(ArrQCheckedVariables, TestName))
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
                'End If
                If I <> 0 Then
                  'v2.4.9 don't go line to line if not Declarations
                  Exit For
                End If
              Next M
            Next I
          End If
          If UpDated Then
            ReWriteMembers Comp.CodeModule, arrMembers, UpDated
          End If
          If bAborting Then
            Exit For 'Sub
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If

Exit Sub

BugTrap:
  BugTrapComment "Unused_Private"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub Unused_Public()

  
  Dim UpDated      As Boolean
  Dim isUsed       As Boolean
  Dim NotPublic    As Boolean
  Dim I            As Long
  Dim M            As Long
  Dim CurCompCount As Long
  Dim MaxFactor    As Long
  Dim TestName     As String
  Dim TestType     As String
  Dim arrMembers   As Variant
  Dim ArrProc      As Variant
  Dim arrLine      As Variant
  Dim CompMod      As CodeModule
  Dim Proj         As VBProject
  Dim Comp         As VBComponent

  On Error GoTo BugTrap
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          'DisplayCodePane Comp
          Set CompMod = Comp.CodeModule
          arrMembers = GetMembersArray(CompMod)
          MaxFactor = UBound(arrMembers)
          IsCollectionClass(CurCompCount) = False
          If MaxFactor > -1 Then
            For I = 0 To MaxFactor
              isUsed = False
              MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
              ArrProc = Split(arrMembers(I), vbNewLine)
              For M = LBound(ArrProc) To UBound(ArrProc)
                If I <> 0 Then
                  'v2.4.9 don't go line to line if not Declarations
                  M = GetProcCodeLineOfRoutine(ArrProc)
                End If
                NotPublic = False
                isUsed = False
                If LeftWord(ArrProc(M)) = "Public" Then
                  arrLine = Split(ExpandForDetection(ArrProc(M)))
                  GetTestName arrLine, TestName, TestType
                  MemberMessage TestName, M, UBound(ArrProc)
                  If LenB(TestName) Then
                    'v2.9.5 test that prevents Paul Caton subclass from being commented out
                    If bPaulCatonSubClasUsed Then
                      isUsed = TestName = strPaulCatonSubClasProcName
                    End If
                    If LeaveImplementsProcsAlone(TestName, Comp.Name) Then
                      isUsed = True 'leave Implements classes alone
                    End If
                    '                    If InQSortArray(ImplementsArray, ModDesc(CurCompCount).MDName) Then
                    '                      isUsed = True
                    '                    End If
                    If Not isUsed Then
                      UnusedWithEventsTest TestType, TestName, ArrProc(M), Comp.Name, isUsed, NotPublic
                    End If
                    If Not isUsed Then
                      UnusedEventsTest TestType, TestName, ArrProc(M), Comp.Name, isUsed, NotPublic
                    End If
                    If Not isUsed Then
                      If Not InQSortArray(ArrQCheckedVariables, TestName) Then
                        UsedTests TestType, TestName, Comp, isUsed
                        If Not isUsed Then
                          ' this searches all routines except current one to let the
                          'later 'Scope Too Large' test to operate and suggest reducing scope value
                          isUsed = FindCodeUsage(TestName, ArrProc(M), Comp.Name, False, False, False)
                        End If
                        If Not isUsed Then
                          isUsed = IsANewEnum(arrMembers, TestName)
                          If isUsed Then
                            If Not IsCollectionClass(CurCompCount) Then
                              IsCollectionClass(CurCompCount) = True
                            End If
                          End If
                        End If
                        If Not isUsed Then
                          ScopeTypeEnum CompMod, TestType, TestName, ArrProc, M, isUsed, NotPublic
                        End If
                        If Not isUsed Then
                          ScopeConst TestType, TestName, M, isUsed, NotPublic
                        End If
                        If Not isUsed Then
                          Select Case TestType
                           Case "Enum"
                            If Not IsInParameter(arrMembers, TestName, I) Then
                              '*if is in paramater then Enum references are public even if
                              'not references in any other module as the members may be referenced
                              NotPublic = True
                            End If
                           Case "Property"
                            'don't mark unused Class or UserControl Properties as too large
                            If IsComponent_UserControl_Class(Comp) Then
                              NotPublic = False
                            End If
                           Case "Event"
                            'events are always Public so never suggest Scope Too Large
                            NotPublic = False
                          End Select
                        End If
                        If Not isUsed Then
                          If CallByNameIsUsed Then
                            isUsed = RoutineIsCalledByCallByName(TestName, False, False, Comp.Name)
                          End If
                        End If
                        If Not isUsed Then
                          If CheckDefaultAttribute(CompMod, TestName) Then
                            'new detects default member of class based tools
                            isUsed = True
                          End If
                        End If
                        If Not isUsed Then
                          If TestType = "Function" Or TestType = "Sub" Or TestType = "Property" Then
                            isUsed = FindCodeUsage(TestName, ArrProc(M), Comp.Name, False, False, False)
                            If isUsed Then
                              arrLine(0) = "Private"
                              NotPublic = True
                            End If
                          End If
                        End If
                        If Not isUsed Then
                          isUsed = ProtectConnectVBIDE(ArrProc(M))
                        End If
                        If Not isUsed Then
                          UnusedMsg TestType, TestName, ArrProc, M, Comp
                          If TestType = "Function" Or TestType = "Sub" Or TestType = "Property" Then
                            arrMembers(I) = vbNullString
                            arrMembers(UBound(arrMembers)) = arrMembers(UBound(arrMembers)) & vbNewLine & _
                             Join(ArrProc, vbNewLine)
                           Else
                            arrMembers(I) = Join(ArrProc, vbNewLine)
                          End If
                          UpDated = True
                          UnusedRecord TestType
                         ElseIf NotPublic Then
                          If Not IsComponent_UserControl_Class(Comp) Then
                            'Property in class very likely to be a prototype so this ignores them
                            ScopeTooLargeMsg TestType, TestName, ArrProc, M
                            arrMembers(I) = Join(ArrProc, vbNewLine)
                            UpDated = True
                            'whatever was done don't recheck
                            UnusedRecord TestType
                          End If
                        End If
                        ArrQCheckedVariables = QuickSortArray(AppendArray(ArrQCheckedVariables, TestName))
                      End If
                    End If
                  End If
                End If
                'End If
                If I <> 0 Then
                  'v2.4.9 don't go line to line if not Declarations
                  Exit For
                End If
              Next M
              'End If
            Next I
          End If
          If UpDated Then
            ReWriteMembers CompMod, arrMembers, UpDated
          End If
          If bAborting Then
            Exit For 'Sub
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If

Exit Sub

BugTrap:
  BugTrapComment "Unused_Public"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub UnusedEventsTest(varTrig As Variant, _
                             ByVal TestName As String, _
                             varCode As Variant, _
                             ByVal CompName As String, _
                             isUsed As Boolean, _
                             NotPublic As Boolean)

  If varTrig = "Event" Then
    'WIthEvents are always Public so never suggest Scope Too Large
    isUsed = FindCodeUsage("_" & TestName, varCode, CompName, False, True, False, True)
    If Not isUsed Then
      isUsed = FindCodeUsage("RaiseEvent " & TestName, varCode, CompName, False, True, False, True)
    End If
    If isUsed Then
      NotPublic = False
    End If
  End If

End Sub

Private Function UnusedFixType(varTyp As Variant) As Long

  Select Case varTyp
   Case "Const"
    UnusedFixType = UnusedDecConst
   Case "Declare"
    UnusedFixType = UnusedDecVariable
   Case "Function"
    UnusedFixType = UnusedFunction
   Case "Sub"
    UnusedFixType = UnusedSub
   Case "Property"
    UnusedFixType = UnusedProperty
   Case "Events"
    UnusedFixType = UnusedEvents
   Case "WithEvents"
    UnusedFixType = UnusedWithEvents
   Case Else
    UnusedFixType = UnusedDecVariable
  End Select

End Function

Private Sub UnusedMsg(varTyp As Variant, _
                      ByVal strName As String, _
                      VarArray As Variant, _
                      ByVal CurLine As Long, _
                      Comp As VBComponent)

  Dim ArrTmp2      As Variant
  Dim J            As Long
  Dim MsgPos       As Long
  Dim I            As Long
  Dim StrUnusedMsg As String
  Dim UCUnused     As Boolean

  'This changes codefixer comment attachment point if necessary.
  'Due to possible line continuation characters the correct spot
  'for the comment may not be just after the current line
  'although it usually will be.
  UCUnused = ProtectUnusedClassOrControl(Comp)
  MsgPos = CurLine
  If UCUnused Then
    StrUnusedMsg = WARNING_MSG & "This " & ModuleType(Comp.CodeModule) & " member is not used in the current project," & vbNewLine & _
     "and may be removed for purposes of this program."
   Else
    strName = strInSQuotes(strName, True)
    StrUnusedMsg = WARNING_MSG & "Unused " & IIf(varTyp <> strName, varTyp & strName, "Variable" & strName) & IIf(Xcheck(XVerbose), vbNewLine & _
     "May be a prototype " & varTyp & " you have not yet implimented or left over from a deleted Control.", vbNullString)
    If strName = "Item" Then
      StrUnusedMsg = StrUnusedMsg & vbNewLine & _
       "Item in a collection class may be marked unused if it is called as a default Method."
    End If
  End If
  'because this treats variables as well as named structures (and varTyp of variables is variable name)
  MsgPos = GetSafeInsertLineArray(VarArray, MsgPos)
  Select Case FixData(UnusedFixType(varTyp)).FixLevel
   Case CommentOnly
    SafeInsertArrayMarker VarArray, MsgPos, StrUnusedMsg
   Case FixAndComment
    'Designed to allow commenting out of Declare, Const and Variable lines NOT function,sub,property,event,withevents
    If Not UCUnused Then
      If ArrayMember(varTyp, "Declare", "Const") Or varTyp = "Variable" Or varTyp = strName Then
        VarArray(CurLine) = "''" & VarArray(CurLine)
        If MsgPos <> CurLine Then
          For I = CurLine To MsgPos
            VarArray(I) = "''" & VarArray(I)
          Next I
        End If
      End If
    End If
    SafeInsertArrayMarker VarArray, MsgPos, StrUnusedMsg
    If Not UCUnused Then
      If IsInArray(varTyp, ArrFuncPropSub) Then
        For I = LBound(VarArray) To UBound(VarArray)
          ArrTmp2 = Split(VarArray(I), vbNewLine)
          For J = LBound(ArrTmp2) To UBound(ArrTmp2)
            If InStr(Left$(ArrTmp2(J), 6), MoveableComment) = 0 Then
              'v 2.2.2 stop multiples applied if previous proc has become a pre-proc comment due to being commented out
              ArrTmp2(J) = MoveableComment & ArrTmp2(J)
            End If
          Next J
          VarArray(I) = Join(ArrTmp2, vbNewLine)
        Next I
        VarArray(LBound(VarArray)) = MoveableComment & vbNewLine & VarArray(LBound(VarArray))
        VarArray(UBound(VarArray)) = VarArray(UBound(VarArray)) & vbNewLine & MoveableComment
      End If
    End If
    'Case JustFix
    'Much too dangerous to activate this
  End Select

End Sub

Private Sub UnusedRecord(varTyp As Variant)

  Select Case varTyp
   Case "Const"
    AddNfix UnusedDecConst
   Case "Declare"
    AddNfix UnusedDecVariable
   Case "Function"
    AddNfix UnusedFunction
   Case "Sub"
    AddNfix UnusedSub
   Case "Property"
    AddNfix UnusedProperty
   Case "Events"
    AddNfix UnusedEvents
   Case "WithEvents"
    AddNfix UnusedWithEvents
   Case Else
    AddNfix UnusedDecVariable
  End Select

End Sub

Private Sub UnusedWithEventsTest(varTrig As Variant, _
                                 ByVal TestName As String, _
                                 varCode As Variant, _
                                 ByVal CompName As String, _
                                 isUsed As Boolean, _
                                 NotPublic As Boolean)

  If varTrig = "WithEvents" Then
    'WIthEvents are always Public so never suggest Scope Too Large
    isUsed = FindCodeUsage(TestName & "_", varCode, CompName, False, True, False, True)
    If Not isUsed Then
      isUsed = FindCodeUsage("." & TestName, varCode, CompName, False, True, False, True)
    End If
    If Not isUsed Then
      isUsed = FindCodeUsage(TestName, varCode, CompName, False, True, False, True)
    End If
    If isUsed Then
      NotPublic = False
    End If
  End If

End Sub

Private Sub UsageCountDelarations(ByVal StrScope As String)

  
  Dim Comp          As VBComponent
  Dim Proj          As VBProject
  Dim L_CodeLine    As String
  Dim CurCompCount  As Long
  Dim StartLine     As Long
  Dim strTmp        As String
  Dim InitFindLine  As Long
  Dim lngLastHit    As Long
  Dim J             As Long
  Dim bLastLine     As Boolean
  Dim CompMod       As CodeModule
  Dim L_Origline    As String
  Dim bSkipDeadEnum As Boolean

  On Error GoTo BugTrap
  If Not bAborting Then
    '
    'Copyright 2004 Roger Gilchrist
    'e-mail: rojagilkrist@hotmail.com
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount, False) Then
          ModuleMessage Comp, CurCompCount
          Set CompMod = Comp.CodeModule
          With CompMod
            InitFindLine = 1
            If .Find(StrScope, InitFindLine, 1, .CountOfDeclarationLines, 1, True, True, False) Then
              'if exits at all, then look for the line(s)
              If InitFindLine <= .CountOfDeclarationLines Then
                StartLine = 1
                bLastLine = False
                Do While .Find(StrScope, StartLine, 1, -1, -1, True, True)
                  L_CodeLine = .Lines(StartLine, 1)
                  'Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, StrScope, L_CodeLine, StartLine)
                  bLastLine = StartLine = .CountOfDeclarationLines
                  If bLastLine Then
                    If StartLine < .CountOfDeclarationLines Then
                      Exit Do
                    End If
                  End If
                  If StartLine > .CountOfDeclarationLines Then
                    Exit Do
                  End If
                  MemberMessage "", StartLine, .CountOfDeclarationLines
                  If Not IsECPCode(L_CodeLine, StartLine, CompMod) Then
                    L_Origline = L_CodeLine
                    If ExtractCode(L_CodeLine) Then
                      If StrScope = "Private" Then
                        For J = StartLine To lngLastHit Step -1
                          If J > 0 Then
                            If SmartLeft(.Lines(J, 1), Hash_If_False_Then) Then
                              If Left$(.Lines(J, 1), 8) = "End Type" Then
                                GoTo ECPSkip
                              End If
                            End If
                          End If
                        Next J
                      End If
                      If InCode(L_CodeLine, InStr(L_CodeLine, StrScope & SngSpace)) Then
                        strTmp = GetDeclarationCount(L_CodeLine, L_Origline, StrScope, Comp.Name, bSkipDeadEnum)
                        If LenB(strTmp) Then
                          SafeInsertModule CompMod, StartLine, strTmp
                          lngLastHit = StartLine
                        End If
                      End If
ECPSkip:
                    End If
                  End If
                  If InStr(L_CodeLine, "Enum ") Then
                    If InCode(L_CodeLine, InStr(L_CodeLine, "Enum ")) Then
                      StartLine = StartLine + 1
                      Do Until InStr(.Lines(StartLine, 1), "End Enum")
                        If Not bSkipDeadEnum Then
                          L_CodeLine = .Lines(StartLine, 1)
                          If Not JustACommentOrBlank(L_CodeLine) Then
                            If SmartLeft(L_CodeLine, "#If False Then") Then
                              Do
                                StartLine = StartLine + 1
                              Loop Until SmartLeft(L_CodeLine, "#End If")
                            End If
                            L_CodeLine = .Lines(StartLine, 1)
                            If Not JustACommentOrBlank(L_CodeLine) Then
                              strTmp = GetDeclarationCount(L_CodeLine, L_CodeLine, StrScope, Comp.Name, bSkipDeadEnum)
                              If LenB(strTmp) Then
                                SafeInsertModule CompMod, StartLine, strTmp
                              End If
                            End If
                          End If
                          If bSkipDeadEnum Then
                            bSkipDeadEnum = False
                            Exit Do
                          End If
                        End If
                        StartLine = StartLine + 1
                      Loop
                      bSkipDeadEnum = False
                    End If
                  End If
                  StartLine = StartLine + 1
                  If StartLine > .CountOfDeclarationLines Then
                    Exit Do
                  End If
                Loop
              End If
            End If
          End With
        End If
      Next Comp
    Next Proj
  End If

Exit Sub

BugTrap:
  BugTrapComment "UsageCountDelarations"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub UsedTests(ByVal TestType As String, _
                      TestName As String, _
                      Comp As VBComponent, _
                      isUsed As Boolean)

  If TestType = "Sub" Then
    If Not Right$(TestName, 33) = "DELETE_ME_INDEXED_VERSION_CREATED" Then
      isUsed = RoutineNameIsVBGenerated(TestName, Comp)
    End If
  End If
  If Not isUsed Then
    isUsed = isCodeFixProtected(TestName)
  End If

End Sub

Private Sub UsedTestsPrivate(TestName As String, _
                             Comp As VBComponent, _
                             tmpB As Variant, _
                             M As Long, _
                             isUsed As Boolean)

  If Not isUsed Then
    isUsed = RoutineNameIsEvent(TestName, Comp)
  End If
  If Not isUsed Then
    isUsed = isEvent(TestName, "Private", Comp.Name)
  End If
  If Not isUsed Then
    isUsed = InEnumCapProtection(Comp.CodeModule, tmpB, M)
  End If
  If Not isUsed Then
    isUsed = FindCodeUsage(TestName, CStr(tmpB(M)), Comp.Name, False, True, False)
  End If

End Sub

''Public Function IsRealWord(strCode As String,
'''                      strTest As String) As Boolean
'''v2.4.4 replaced older version
'''this is Ulli's IsIn function only change is the name
''' and the added inCode detector at the end
''  Const Delims As String = " ,.:'()[]"""
''  Dim J        As Long
'''returns true if Text is in strCode as word
''
''  If Len(strCode) Then
''    J = InStr(strCode, strTest)
''    If J Then
''    Do
''      IsRealWord = InStr(Delims, Mid$(strCode, J + Len(strTest), 1))
''      If J > 1 Then
''        IsRealWord = IsRealWord And InStr(Delims, Mid$(strCode, J - 1, 1))
''      End If
''      J = InStr(J + 1, strCode, strTest)
''   Loop Until J = 0 Or IsRealWord
''    End If
''  End If
''  If IsRealWord Then
''  'safety make sure it is in code not comment/string literal
''  'but only spend the time if it is necessary
''    IsRealWord = InCode(strCode, J)
''  End If
''End Function

':)Code Fixer V3.0.9 (25/03/2005 4:23:25 AM) 10 + 2379 = 2389 Lines Thanks Ulli for inspiration and lots of code.

