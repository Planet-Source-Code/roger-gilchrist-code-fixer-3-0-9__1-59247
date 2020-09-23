Attribute VB_Name = "mod_LocalFix"

Option Explicit
Public bCtrlDefFix     As Boolean

Private Function AlreadyDimmed(ByVal CompName As String, _
                               ByVal varTest As Variant, _
                               ByVal ExistingDimArray As Variant) As Boolean

  'test if already dimmed
  'Put any words/code you don't want hit here if they are nowhere else
  'These test are in order of best chance of getting a hit

ReTry:
  AlreadyDimmed = (varTest = "Err")
  If Not AlreadyDimmed Then
    AlreadyDimmed = InArrayinSomeFormat(varTest, ExistingDimArray)
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = InQSortArray(ArrQKnowVBWord, varTest)
  End If
  If Not AlreadyDimmed Then
    ' test for being in any of the know user generated code
    AlreadyDimmed = IsDeclaration(CStr(varTest), "Private", , CompName)
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = IsDeclareName(varTest)
    'IsDeclaration(CStr(varTest)) 'InArrayinSomeFormat(varTest, PublicDecArray)
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = InTypeSomeFormat(varTest)
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = IsProcedure(CStr(varTest), "Public")
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = IsProcedure(CStr(varTest), "Friend")
  End If
  If Not AlreadyDimmed Then
    'v2.3.5 incorrect parameter setting fixed
    AlreadyDimmed = IsProcedure(CStr(varTest), "Private", , CompName)
  End If
  If Not AlreadyDimmed Then
    'v2.3.5 incorrect parameter setting fixed
    AlreadyDimmed = IsProcedure(CStr(varTest), "Static", , CompName)
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = IsControlProperty(CStr(varTest))
  End If
  If Not AlreadyDimmed Then
    AlreadyDimmed = isEvent(CStr(varTest))
  End If
  If Not AlreadyDimmed Then
    'v2.1.0 Thanks to Evan Toder who showed me this potenetial bug
    ' retest it there is a Type suffix.
    If InQSortArray(TypeSuffixArray, Right$(varTest, 1)) Then
      varTest = Left$(varTest, Len(varTest) - 1)
      GoTo ReTry
    End If
  End If

End Function

Public Sub CompactIndexes(Arr As Variant)

  Dim I           As Long
  Dim J           As Long
  Dim lngTmpIndex As Long

  For I = LBound(Arr) To UBound(Arr)
    lngTmpIndex = CntrlDescMember(Arr(I))
    If lngTmpIndex > -1 Then
      If CntrlDesc(lngTmpIndex).CDIndex > -1 Then
        If I < UBound(Arr) Then
          If Arr(I + 1) = LBracket Then
            J = 0
            Do While CountSubStringImbalance(Arr(I + 1), LBracket, RBracket)
              J = J + 1
              If J = UBound(Arr) Then
                Exit Do
              End If
              Arr(I + 1) = Arr(I + 1) & SngSpace & Arr(I + 1 + J)
              Arr(I + 1 + J) = vbNullString
            Loop
            Arr(I + 1) = Replace$(Replace$(Arr(I + 1), "( ", LBracket), " )", RBracket)
          End If
        End If
      End If
    End If
  Next I
  Arr = CleanArray(Arr)

End Sub

Private Function ControlDescription(ByVal IDNo As Long) As String

  If IDNo > -1 Then
    With CntrlDesc(IDNo)
      ControlDescription = strInSQuotes(.CDName, True) & strInSQuotes(.CDClass) & " on " & strInSQuotes(.CDForm)
    End With
  End If

End Function

Public Sub CtrlDefPropFast()

  Dim Proj    As VBProject
  Dim Comp    As VBComponent
  Dim I       As Long
  Dim arrDone As Variant

  'v2.7.2 new fix to replace DimControlDefaultProp
  'This quickly inserts default properties using
  ' a complete sweep of code for each control
  'v2.7.6 added to stop multiple calls to control array (1st call should do all)
  On Error Resume Next
  WorkingMessage "Default Property of Controls", 0, 1
  For I = LBound(CntrlDesc) To UBound(CntrlDesc)
    MemberMessage CntrlDesc(I).CDName, I, UBound(CntrlDesc)
    If Not IsInArray(CntrlDesc(I).CDName, arrDone) Then
      arrDone = AppendArray(arrDone, CntrlDesc(I).CDName)
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If LenB(Comp.Name) Then
            If CntrlDesc(I).CDUsage = 1 Then ' has code presence(ignores forms/classes/fakes etc)
              If LenB(CntrlDesc(I).CDDefProp) Then ' has a defprop (most UC don't)
                DefPropAdd Comp, I
              End If
            End If
          End If
        Next Comp
      Next Proj
    End If
  Next I
  On Error GoTo 0

End Sub

Private Sub DefPropAdd(Cmp As VBComponent, _
                       CntrlNo As Long)

  
  Dim TmpA             As Variant
  Dim UpDated          As Boolean
  Dim DefProp          As String
  Dim CommentUpdated   As Boolean
  Dim IsIndexed        As Boolean
  Dim CtrlAffected     As String
  Dim TmpBSize         As Long
  Dim CFMsg            As String  'report multiple controls default property added
  Dim CFMsg2           As String  'report multiple controls Form Reference added
  Dim CFMsg3           As String  'report multiple controls Me/UserControl added
  Dim RefUpdated       As Boolean
  Dim LParam           As Variant
  Dim ProcName         As String
  Dim DimLine          As Long
  Dim lngJunkMe        As Long
  Dim strComentId      As String
  Dim StrFullCtrlID    As String
  Dim Sline            As Long
  Dim ELine            As Long
  Dim SCol             As Long
  Dim Ecol             As Long
  Dim L_CodeLine       As String
  Dim StrSearchFor     As String
  Dim CommentStore     As String
  Dim lngTmpIndex      As Long
  Dim tmpB             As Variant
  Dim I                As Long
  Dim J                As Long
  Dim K                As Long
  Dim ExistingDimArray As Variant

  Sline = 1
  SCol = 1
  ELine = -1
  Ecol = -1
  StrSearchFor = CntrlDesc(CntrlNo).CDName
  '  If Len(StrSearchFor) <= 2 Then '65=AA 67=Cd 68=Th
  '  'Stop
  '  End If
  '  If CntrlNo >= 400 Then
  '  Stop
  '  End If
  '
  '  If StrSearchFor = "txtpmi" Then ') <= 2 Then '65=AA 67=Cd 68=Th
  '  Stop
  '  End If
  If CntrlDesc(CntrlNo).CDIndex > -1 Then
    StrSearchFor = CntrlDesc(CntrlNo).CDName & "(*)"
  End If
  Do While Cmp.CodeModule.Find(StrSearchFor, Sline, SCol, ELine, Ecol, False, False, True)
    L_CodeLine = Cmp.CodeModule.Lines(Sline, 1)
    'v3.0.0 new test should speed up for shortname controls
    If Not JustACommentOrBlankOrDimLine(L_CodeLine) Then
      If IsRealWord(L_CodeLine, StrSearchFor) Then
        If Not IsDimLine(L_CodeLine) Then
          If isProcHead(L_CodeLine) Then
            'fixed ver1.1.40 crashed if prochead used line continuation
            ProcName = GetProcName(Cmp.CodeModule, Sline) ' GetProcNameStr(L_CodeLine)
            TmpA = ReadProcedureCodeArray
            LParam = ParamArrays(GetWholeLineArray(TmpA, I, lngJunkMe), VariableOnlyParam)
            GoTo NoSuspectCode  ' no suspect code; so get next line
          End If
          'CommentStore = CommentClip(L_Codeline)
          ExtractCode L_CodeLine, CommentStore
          'TmpB = ExpandForDetection2(L_Codeline)
          tmpB = Split(ExpandForDetection(L_CodeLine))
          If Not PotentialControlDefault(Cmp, tmpB, StrSearchFor) Then
            GoTo NoSuspectCode  ' no suspect code; so get next line
          End If
          CompactIndexes tmpB
          TmpBSize = UBound(tmpB)
          If TmpBSize Then
            If Not IsEmpty(TmpA) Then
              ExistingDimArray = BuildExistingDimArray(CStr(Join(TmpA, vbNewLine)))
            End If
            J = ArrayPos(CntrlDesc(CntrlNo).CDName, tmpB) '((J + 1
            'ver 1.1.25 because J may be reset within the For loop to excess value
            If J > -1 Then
              lngTmpIndex = CntrlDescMember(tmpB(J))
              lngTmpIndex = CntrlDescMember2(tmpB(J), Cmp.Name)
              'Deal with Properties of Forms/UserControls being used without reference to the Form/UserControl
              '-------------------------------
              If lngTmpIndex = -1 Then
                'Not a control test for FreeStandingProperty
                If IsControlProperty(tmpB(J)) Then
                  If IsComponent_ControlHolder(Cmp) Then
                    FreeStandingProperty Cmp, tmpB, J, L_CodeLine, CFMsg3, ExistingDimArray, TmpBSize
                    GoTo DoMessage
                  End If
                End If
               Else
                '-------------------------------
                'skip 'Is' tests
                If J > 1 Then
                  If lngTmpIndex > -1 Then
                    If tmpB(J - 1) = "Is" Then
                      GoTo AlreadyHasOne
                    End If
                  End If
                End If
                'ver 1.1.31 avoid ifthere is an Is test
                If tmpB(J) <> ProcName Then
                  ' it is the Get Property with the same name so thats OK
                  If InQSortArray(ExistingDimArray, tmpB(J)) Then
                    'ver 1.1.31 Control and Local Dim have same name
                    Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(TmpA(I), CommentStore)
                    SafeInsertModule Cmp.CodeModule, Sline, WARNING_MSG & "The variable '" & tmpB(J) & "' has the same name as a Control." & vbNewLine & _
                     "While legal this makes code harder to read and may prevent some Code Fixer actions." & SuggestNewName(tmpB(J), TmpA(DimLine))
                    DimLine = FindDimLine(tmpB(J), TmpA)
                   ElseIf IsInArray(tmpB(J), LParam) Then
                    Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(TmpA(I), CommentStore)
                    SafeInsertModule Cmp.CodeModule, Sline, WARNING_MSG & "Control and a Parameter '" & tmpB(J) & "' have the same name." & vbNewLine & _
                     "While legal this makes code harder to read and may prevent some Code Fixer actions." & SuggestNewName(tmpB(J), TmpA(DimLine))
                  End If
                 ElseIf IsControlProperty(tmpB(J)) And Not IsVBControlRoutine(ProcName) Then
                  Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(TmpA(I), CommentStore)
                  SafeInsertModule Cmp.CodeModule, Sline, WARNING_MSG & "The variable '" & tmpB(J) & "' has the same name as a Control Property." & vbNewLine & _
                   "While legal this makes code harder to read and may prevent some Code Fixer actions." & SuggestNewName(tmpB(J), TmpA(DimLine))
                  DimLine = FindDimLine(tmpB(J), TmpA)
                 ElseIf isVBReservedWord(ProcName) Then
                  Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(TmpA(I), CommentStore)
                  SafeInsertModule Cmp.CodeModule, Sline, WARNING_MSG & "The variable '" & tmpB(J) & "' has the same name as a VB Reserved WOrd." & vbNewLine & _
                   "While legal this makes code harder to read and may prevent some Code Fixer actions." & SuggestNewName(tmpB(J), TmpA(DimLine))
                  DimLine = FindDimLine(tmpB(J), TmpA)
                End If
              End If
              'Deal with Default properties
              If lngTmpIndex > -1 Then
                DefProp = CntrlDesc(lngTmpIndex).CDDefProp
                IsIndexed = CntrlDesc(lngTmpIndex).CDIndex > -1
                If IsIndexed Then
                  If J + 2 <= UBound(tmpB) Then
                    If (Left$(tmpB(J + 2), 1) = ".") Then
                      GoTo AlreadyHasOne
                    End If
                  End If
                 Else
                  If J + 1 <= UBound(tmpB) Then
                    If (Left$(tmpB(J + 1), 1) = ".") Then
                      GoTo AlreadyHasOne
                    End If
                  End If
                End If
                If IsIndexed Then
                  'collect the index elements
                  ' as ExpandForDetection separates brackets from everything
                  ' this section of code restores the index parameter to the
                  ' control before adding properties
                  ' but only if it is  NOT a reference to the whole collection of controls
                  ' either last element
                  If J <> UBound(tmpB) Then
                    'or without index reference
                    If Left$(tmpB(J + 1), 1) = "(" Then
                      If Right$(tmpB(J + 1), 1) <> ")" Then
                        K = J + 1
                        Do Until tmpB(K) = ")"
                          tmpB(J) = tmpB(J) & tmpB(K)
                          tmpB(K) = vbNullString
                          K = K + 1
                          If K > UBound(tmpB) Then
                            'safety
                            K = UBound(tmpB)
                            Exit Do
                          End If
                        Loop
                        tmpB(J) = tmpB(J) & tmpB(K)
                        tmpB(K) = vbNullString
                       Else
                        tmpB(J) = tmpB(J) & tmpB(J + 1)
                        tmpB(J + 1) = vbNullString
                      End If
                     Else
                      'indexed but no index present means that it is being passed as an array
                      'so don't touch it
                      GoTo AlreadyHasOne
                    End If
                  End If
                End If
              End If
              'v2.3.7 FIXME control/module level varaiable
              If notBadNameOrFixableBad(tmpB(J)) Then
                'ver 1.1.42 tests for control names that are unsafe (for Code Fixer)
                With Cmp
                  If Not IsParameter2(tmpB, J, .Name, False) Then
                    RefUpdated = ControlNeedsFormRef(tmpB(J), .Name)
                    InsertDefProp L_CodeLine, tmpB(J), DefProp, UpDated, CommentUpdated, RefUpdated
                    If CommentUpdated Then
                      strComentId = tmpB(J)
                      StrFullCtrlID = ControlDescription(lngTmpIndex)
                      If LenB(StrFullCtrlID) = 0 Then
                        'for some reason the control was not fully identified
                        StrFullCtrlID = ControlDescription(CntrlDescMember(tmpB(J))) & IIf(lngTmpIndex, "(May be incorrect)", vbNullString)
                      End If
                    End If
                    If UpDated Then
                      CtrlAffected = tmpB(J)
                      If RefUpdated Then
                        CFMsg2 = CFMsg2 & WARNING_MSG & "Full Form Reference for '" & CtrlAffected & "' inserted."
                        RefUpdated = False
                      End If
                      CFMsg = CFMsg & WARNING_MSG & "Default Property of Control '" & CtrlAffected & "' inserted."
                      tmpB = ExpandForDetection2(L_CodeLine)
                      CompactIndexes tmpB
                      TmpBSize = UBound(tmpB)
                      UpDated = False
                    End If
                  End If
                End With 'cmp
              End If
            End If
AlreadyHasOne:
          End If
          '-------------------------
DoMessage:
          If LenB(CFMsg) Then
            Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(L_CodeLine, CommentStore)
            SafeInsertModule Cmp.CodeModule, Sline, CFMsg
            CFMsg = vbNullString
           ElseIf CommentUpdated Then
            'v2.9.5 this was the wrong 'fix' when it should have just reported difficulties Thanks Ian K
            'Cmp.CodeModule.ReplaceLine Sline, L_CodeLine & SafeReattachComment(L_CodeLine, CommentStore)
            SafeInsertModule Cmp.CodeModule, Sline, WARNING_MSG & "Problem identifying " & strInSQuotes(strComentId) & "." & vbNewLine & _
             RGSignature & IIf(LenB(StrFullCtrlID), "Variable has same name as Control " & StrFullCtrlID & ".", "Variable has same name as a Control/VB command") & IIf(LenB(StrFullCtrlID), vbNullString, vbNewLine & _
             RGSignature & "or a Control without a detectable default Property.")
            CommentUpdated = False
            strComentId = vbNullString
            StrFullCtrlID = vbNullString
          End If
          If LenB(CFMsg2) Then
            SafeInsertModule Cmp.CodeModule, Sline, CFMsg2
            CFMsg2 = vbNullString
          End If
          If LenB(CFMsg3) Then
            L_CodeLine = L_CodeLine & SafeReattachComment(TmpA(I), CommentStore)
            If InStr(L_CodeLine, ") .") Then
              L_CodeLine = Safe_Replace(L_CodeLine, ") .", ").")
            End If
            Cmp.CodeModule.ReplaceLine Sline, L_CodeLine
            SafeInsertModule Cmp.CodeModule, Sline, CFMsg3
            CFMsg3 = vbNullString
          End If
        End If
      End If
    End If
NoSuspectCode:
    Sline = Sline + 1
    SCol = 1
    ELine = -1
    Ecol = -1
    If Sline = 1 Or Sline > Cmp.CodeModule.CountOfLines Then
      Exit Do
    End If
  Loop

End Sub

Private Function DetectErrDetectionFunction(ArrCode As Variant, _
                                            strProcHead As String) As Boolean

  Dim I        As Long
  Dim bOnError As Boolean
  Dim strProc  As String

  'Detect functions which use the success or failure of assignment to generate a return
  If InStr(strProcHead, "Function ") Then
    strProc = WordAfter(strProcHead, "Function")
    If UBound(ArrCode) - LBound(ArrCode) < 15 Then ' only small Functions are likely to do this
      For I = LBound(ArrCode) + 1 To UBound(ArrCode) - 1
        'v2.9.0 corrected error in test
        If IsOnErrorCode(ArrCode(I)) Then
          bOnError = True
        End If
        If bOnError Then
          If InStr(ArrCode(I), " = True") Then
            If InCode(ArrCode(I), InStr(ArrCode(I), " = True")) Then
              If WordBefore(ArrCode(I), "=") = strProc Then
                DetectErrDetectionFunction = True
              End If
            End If
          End If
        End If
      Next I
    End If
  End If

End Function

Public Sub Dim_Engine()

  Dim I            As Long
  Dim strRoutine   As String
  Dim arrMembers   As Variant
  Dim UpDated      As Boolean
  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim CompMod      As CodeModule
  Dim NumFixes     As Long
  Dim arrTest      As Variant

  arrTest = Array("Dim ", "Static ", "Const ")
  NumFixes = 13
  On Error GoTo BugHit
  If Not bAborting Then
    If bCtrlDefFix Then
      CtrlDefPropFast
    End If
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          Set CompMod = Comp.CodeModule
          If CompMod.Members.Count Then
            arrMembers = GetMembersArray(CompMod)
            DeclarationDefTypeDetector CurCompCount, Split(arrMembers(0), vbNewLine), True, UpDated, arrMembers(0)
            For I = 1 To UBound(arrMembers)
              MemberMessage GetProcNameStr(arrMembers(I)), I, UBound(arrMembers)
              If LenB(arrMembers(I)) Then
                strRoutine = arrMembers(I)
                WorkingMessage "Move Dims to top of routine", 1, NumFixes
                strRoutine = DimMoveToTop(strRoutine)
                'v2.3.4 new fix mainly to cope with the duplicates created by creating control arrays
                If InstrArrayHard(strRoutine, arrTest) Then
                  WorkingMessage "Double Dims", 2, NumFixes
                  strRoutine = DimDoubleUp(strRoutine)
                  WorkingMessage "Multiple Dims with single Type", 3, NumFixes
                  strRoutine = DimMultiSingleTypeing(CurCompCount, strRoutine)
                  WorkingMessage "Expand Multiple Dims", 4, NumFixes
                  strRoutine = DimExpandMulti(CurCompCount, strRoutine)
                  WorkingMessage "Update Type Affix", 5, NumFixes
                  strRoutine = DimTypeUpdate(CurCompCount, strRoutine)
                  WorkingMessage "Untyped Dims", 6, NumFixes
                  strRoutine = DimUntyped(CurCompCount, strRoutine, Comp.Name)
                  WorkingMessage "Unused Dims", 7, NumFixes
                  strRoutine = DimDead(CurCompCount, strRoutine)
                End If
                WorkingMessage "Seek Missing Dims", 8, NumFixes
                strRoutine = DimMissing(CurCompCount, strRoutine, Comp.Name)
                If InstrArrayHard(strRoutine, arrTest) Then
                  WorkingMessage "Duplicate Dims", 9, NumFixes
                  strRoutine = DimDuplicate(CurCompCount, strRoutine, Comp.Type, Comp.Name)
                  WorkingMessage "Uneeded Initialisation", 10, NumFixes
                  strRoutine = DimUnneededInitialize(CurCompCount, strRoutine)
                  WorkingMessage "Format Dims", 11, NumFixes
                  strRoutine = DimFormat(strRoutine)
                  WorkingMessage "Dim Usage", 12, NumFixes
                  strRoutine = DimUsage(CurCompCount, strRoutine)
                End If
                WorkingMessage "Type Suffix removal", 13, NumFixes
                strRoutine = TypeSuffixStrip(CurCompCount, strRoutine)
                If Not UpDated Then
                  'test only once; its easier
                  UpDated = Not (arrMembers(I) = strRoutine)
                End If
                arrMembers(I) = strRoutine
              End If
              If bAborting Then
                Exit For
              End If
            Next I
            ReWriteMembers CompMod, arrMembers, UpDated
            If bAborting Then
              Exit For 'Sub
            End If
          End If
          'v3.0.4 fix allows OptionExplicit to be applied to UserControls only after missing Variables have been declared
          If Comp.CodeModule.Parent.Type = vbext_ct_UserControl Then
            InsertOptionExplicit Comp.CodeModule
          End If
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
  BugTrapComment "Dim_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Sub Dim_EngineFormat()

  Dim I            As Long
  Dim strRoutine   As String
  Dim arrMembers   As Variant
  Dim UpDated      As Boolean
  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim CompMod      As CodeModule

  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          Set CompMod = Comp.CodeModule
          If CompMod.Members.Count Then
            arrMembers = GetMembersArray(CompMod)
            DeclarationDefTypeDetector CurCompCount, Split(arrMembers(0), vbNewLine), True, UpDated, arrMembers(0)
            For I = 1 To UBound(arrMembers)
              MemberMessage GetProcNameStr(arrMembers(I)), I, UBound(arrMembers)
              If LenB(arrMembers(I)) Then
                strRoutine = arrMembers(I)
                strRoutine = DimMoveToTop(strRoutine)
                strRoutine = DimMultiSingleTypeing(CurCompCount, strRoutine)
                strRoutine = DimExpandMulti(CurCompCount, strRoutine)
                strRoutine = DimFormat(strRoutine)
                strRoutine = TypeSuffixStrip(CurCompCount, strRoutine)
                If Not UpDated Then
                  'test only once; its easier
                  UpDated = Not (arrMembers(I) = strRoutine)
                End If
                arrMembers(I) = strRoutine
              End If
              If bAborting Then
                Exit For
              End If
            Next I
            ReWriteMembers CompMod, arrMembers, UpDated
            If bAborting Then
              Exit For 'Sub
            End If
          End If
          'v3.0.4 fix allows OptionExplicit to be applied to UserControls only after missing Variables have been declared
          If Comp.CodeModule.Parent.Type = vbext_ct_UserControl Then
            InsertOptionExplicit Comp.CodeModule
          End If
        End If
      Next Comp
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If

End Sub

Private Function DimDead(ByVal ModuleNumber As Long, _
                         ByVal strA As String) As String

  
  Dim K                          As Long
  Dim L_CodeLine                 As String
  Dim MaxFactor                  As Long
  Dim AssignedOnly               As Boolean
  Dim UnusedAPIReturn            As Boolean
  Dim UnusedFunctionReturn       As Boolean
  Dim UnusedFunctionNoParameters As Boolean
  Dim SuspicoiusSelfRefererence  As Boolean
  Dim UnusedVBFunction           As Boolean
  Dim UnusedMsgBoxVariable       As Boolean
  Dim strTmp                     As String
  Dim arrTmp                     As Variant
  Dim ScanStart                  As Long
  Dim strProcHead                As String
  Dim arrDim                     As Variant
  Dim I                          As Long
  Dim J                          As Long
  Dim DimDeadList                As String
  Dim arrLine                    As Variant
  Dim strDummy                   As String
  Dim FoundOne                   As Boolean
  Dim ArrProc                    As Variant
  Dim strTestThis                As String

  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, DetectDimUnused) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      AssignedOnly = True
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        If Not JustACommentOrBlank(ArrProc(I)) Then
          L_CodeLine = ExpandForDetection(ArrProc(I))
          If isProcHead(L_CodeLine) Then
            'v 2.3.3 Thanks Roy Blanch. CF was damaging Static Procedures by
            'commenting out the Header assuming it was a Dim line
            strProcHead = L_CodeLine
           Else
            If IsDimLine(L_CodeLine, , strProcHead) Then
              arrDim = GenerateVariableTypeArray(L_CodeLine, strDummy)
              arrLine = Split(Trim$(L_CodeLine))
              FoundOne = False
              For J = LBound(ArrProc) + 1 To UBound(ArrProc) - 1
                If J <> I Then
                  'v3.0.0 rare bug where the GoTo Label is Dimmed ignores  label only lines line
                  'v3.0.3 improved test Thanks Jim Jose
                  If Not (InStr(ArrProc(J), " ") = 0 And Right$(strCodeOnly(ArrProc(J)), 1) = ":") Then
                    If Not JustACommentOrBlank(ArrProc(J)) Then
                      'v3.0.0 rare bug where the GoTo Label is Dimmed ignore line
                      If Not IsOnErrorCode(ArrProc(J)) Then
                        L_CodeLine = ExpandForDetection(ArrProc(J))
                        'ver1.0.93
                        'added to detect the #<numeral> of open/print/close commands
                        If InstrAtPositionSetArray(L_CodeLine, ipAny, True, arrDim) Then
                          'v2.3.4 added test to solve problem with underscore embedded words
                          For K = LBound(arrDim) To UBound(arrDim)
                            If IsRealWord(L_CodeLine, CStr(arrDim(K))) Then
                              'v3.0.0 rare bug where the GoTo Label is Dimmed ignore line
                              If Not InstrAtPosition(L_CodeLine, "GoTo " & arrDim(K), IpLeft) Then
                                FoundOne = True
                                ScanStart = J
                                Exit For
                              End If
                            End If
                          Next K
                          If FoundOne Then
                            Exit For
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              Next J
              If FoundOne Then ' it is present in code but is it actually used?
                AssignedOnly = True
                ' assume not then search for it in use ('Set Var = ' and 'Var = ' are assign only)
                UnusedAPIReturn = False
                UnusedFunctionReturn = False
                UnusedFunctionNoParameters = False
                UnusedVBFunction = False
                SuspicoiusSelfRefererence = False
                strTmp = vbNullString
                If DetectErrDetectionFunction(ArrProc, strProcHead) Then
                  AssignedOnly = False
                  GoTo SkipErrDetectionFunction
                End If
                For J = ScanStart To UBound(ArrProc) - 1
                  If J <> I Then
                    If Not JustACommentOrBlank(ArrProc(J)) Then
                      L_CodeLine = ExpandForDetection(ArrProc(J))
                      If InstrAtPositionSetArray(L_CodeLine, ipAny, True, arrDim) Then
                        If InStr(L_CodeLine, EqualInCode) = 0 Then
                          AssignedOnly = False 'no equal sign no assignment
                          SuspicoiusSelfRefererence = False
                          Exit For
                         Else
                          If Not MultiLeft2(L_CodeLine, True, arrDim) Then
                            AssignedOnly = False
                            SuspicoiusSelfRefererence = False
                            Exit For
                           Else
                            If MoreThanOnceInLine(L_CodeLine, arrDim) Then
                              SuspicoiusSelfRefererence = True
                              AssignedOnly = False
                              Exit For
                            End If
                            If InstrAtPositionSetArray(L_CodeLine, ip3rdorGreater, True, arrDim) Then
                              If Not InstrAtPositionSetArray(L_CodeLine, IpLeft, True, arrDim) Then
                                AssignedOnly = False
                                Exit For
                               ElseIf InstrAtPositionSetArray(L_CodeLine, IpLeft, True, arrDim) Then
                                SuspicoiusSelfRefererence = True
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                Next J
                If AssignedOnly Then
                  'test for API/Function 'shadowing'? (just a dummy to allow a function to operate)
                  For J = ScanStart To UBound(ArrProc) - 1
                    If J <> I Then
                      If Not JustACommentOrBlankOrDimLine(ArrProc(J)) Then
                        L_CodeLine = ExpandForDetection(ArrProc(J))
                        If MultiLeft2(L_CodeLine, True, arrDim) Then
                          strTestThis = WordAfter(L_CodeLine, "=")
                          If IsAPI(strTestThis) Then
                            UnusedAPIReturn = True
                            Exit For
                           ElseIf strTestThis = "Shell" Then
                            UnusedAPIReturn = True
                            Exit For
                           ElseIf IsAPI(strTestThis) Then
                            UnusedAPIReturn = True
                            Exit For
                           ElseIf WordAfter(L_CodeLine, strTestThis) <> LBracket Then
                            UnusedFunctionNoParameters = True
                          End If
                          Exit For
                         Else
                          If WordAfter(L_CodeLine, strTestThis) = LBracket Then
                            UnusedFunctionReturn = True
                            Exit For
                          End If
                        End If
                      End If
                    End If
                  Next J
                  If Not UnusedAPIReturn Then
                    For J = ScanStart To UBound(ArrProc) - 1
                      If J <> I Then
                        If Not JustACommentOrBlankOrDimLine(ArrProc(J)) Then
                          L_CodeLine = ExpandForDetection(ArrProc(J))
                          If MultiLeft2(L_CodeLine, True, arrDim) Then
                            strTestThis = WordAfter(L_CodeLine, "=")
                            If IsVBFunction(strTestThis) Then
                              If strTestThis = "MsgBox" Then
                                UnusedMsgBoxVariable = True
                               Else
                                UnusedVBFunction = True
                              End If
                             ElseIf IsProcedure(strTestThis, , "Function") Then
                              UnusedFunctionReturn = True
                             ElseIf IsFormProcCall(strTestThis) Then
                              UnusedFunctionReturn = True
                              Exit For
                             ElseIf WordAfter(L_CodeLine, (strTestThis)) <> LBracket Then
                              UnusedFunctionNoParameters = True
                              Exit For
                            End If
                          End If
                        End If
                      End If
                    Next J
                  End If
                End If
                If AssignedOnly Then
                  For J = LBound(ArrProc) + 1 To UBound(ArrProc) - 1
                    If J <> I Then
                      If Not JustACommentOrBlankOrDimLine(ArrProc(J)) Then
                        L_CodeLine = ArrProc(J)
                        If InstrAtPositionSetArray(ArrProc(J), ipLeftOr2nd, True, arrDim) Then
                          strTmp = AccumulatorString(strTmp, J)
                        End If
                      End If
                    End If
                  Next J
                  arrTmp = Split(strTmp, ",")
                End If
              End If
              If Not FoundOne Then
                DimDeadList = AccumulatorString(DimDeadList, arrLine(1), CommaSpace)
              End If
              If LenB(DimDeadList) Then
                DimDeadList = "Unused Dim" & IIf(CountSubString(DimDeadList, CommaSpace) > 1, "s:", ":") & DimDeadList
                AddNfix DetectDimUnused
                AssignedOnly = False
                Select Case FixData(DetectDimUnused).FixLevel
                 Case CommentOnly
                  ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & DimDeadList & " detected", MAfter)
                 Case FixAndComment
                  ArrProc(I) = SQuote & ArrProc(I)
                  ArrProc(I) = SmartMarker(ArrProc, I, WARNING_MSG & "Unused Dim commented out", MAfter)
                 Case JustFix
                  ArrProc(I) = vbNullString
                End Select
                DimDeadList = vbNullString
              End If
SkipErrDetectionFunction:
              If AssignedOnly Then
                Select Case FixData(DetectDimUnused).FixLevel
                 Case CommentOnly
                  If UnusedVBFunction Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 1
                   ElseIf UnusedAPIReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 0
                   ElseIf UnusedMsgBoxVariable Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 2, , True
                   ElseIf UnusedFunctionReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 2
                   Else
                    ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Variable is assigned a value but never used in code." & vbNewLine & _
                     "You can safely delete it and all references to it from the procedure.", MAfter)
                  End If
                 Case FixAndComment, JustFix
                  If UnusedVBFunction Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 1
                    UnUsedDimCommentOut ArrProc, arrTmp, I
                   ElseIf UnusedAPIReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 0, True
                    UnUsedAPIVariable2Call ArrProc, arrTmp, I, arrDim         'tmpV1, tmpV2, tmpV3
                   ElseIf UnusedMsgBoxVariable Then
                    UnUsedAPIVariable2Call ArrProc, arrTmp, I, arrDim         'tmpV1, tmpV2, tmpV3
                   ElseIf UnusedFunctionReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 2, True
                    UnUsedAPIVariable2Call ArrProc, arrTmp, I, arrDim         'tmpV1, tmpV2, tmpV3
                   Else
                    ArrProc(I) = SQuote & ArrProc(I)
                    ArrProc(I) = SmartMarker(ArrProc, I, WARNING_MSG & "Variable was assigned a value but never used in code.", MAfter)
                    UnUsedDimCommentOut ArrProc, arrTmp, I
                  End If
                End Select
              End If
              If SuspicoiusSelfRefererence And Not UnusedAPIReturn And Not UnusedFunctionReturn And Not UnusedFunctionNoParameters And Not UnusedVBFunction Then
                If Not UsedAnywhereLowerIn(ArrProc, I, arrDim) Then
                  If UnusedVBFunction Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 1
                   ElseIf UnusedAPIReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 0
                   ElseIf UnusedFunctionReturn Then
                    FunctionParameterMessage ArrProc, arrLine(1), I, UnusedFunctionNoParameters, 2
                   Else
                    If IsAPI(WordAfter(L_CodeLine, "=")) Then
                      'v2.2.6 Thanks Tom Law new message for rare possible fix
                      ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Variable is assigned a value by an API call but also set as a parameter." & vbNewLine & _
                       "You may be able to replace the parameter with a '0' or '0&' and the assign only fix may be applicable.", MAfter)
                     Else
                      ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Variable is assigned a value but only by self-reference." & vbNewLine & _
                       "You may be able to delete it and all references to it from the procedure.", MAfter)
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimDead = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Function DimDoubleUp(ByVal strA As String) As String

  Dim MaxFactor    As Long
  Dim ArrProc      As Variant
  Dim I            As Long
  Dim J            As Long
  Dim TopOfRoutine As Long

  ArrProc = Split(strA, vbNewLine)
  MaxFactor = UBound(ArrProc)
  If MaxFactor > -1 Then
    TopOfRoutine = FirstCodeLineInProc(ArrProc)
    For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
      If Left$(ArrProc(I), 7) = "#End If" Then
        If I > TopOfRoutine Then
          TopOfRoutine = I
        End If
      End If
    Next I
    For I = TopOfRoutine To MaxFactor
      MemberMessage "", (MaxFactor / 3) * 2 + I / 3, MaxFactor
      If IsDimLine(ArrProc(I), False) Then
        ' don't move dims around each other
        For J = I + 1 To MaxFactor
          If ArrProc(I) = ArrProc(J) Then
            ArrProc(J) = "'" & ArrProc(J) & vbNewLine & WARNING_MSG & "Duplicate dim removed"
          End If
        Next J
      End If
    Next I
  End If
  DimDoubleUp = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Function DimDuplicate(ByVal ModuleNumber As Long, _
                              ByVal strA As String, _
                              ByVal CompType As Long, _
                              ByVal CompName As String) As String

  Dim I                     As Long
  Dim ArrProc               As Variant
  Dim arrLine               As Variant
  Dim MaxFactor             As Long
  Dim L_CodeLine            As String
  Dim IsPublicDec           As Boolean
  Dim IsPrivate             As Boolean
  Dim IsIndexed             As Boolean
  Dim WarnDupNameButIndexed As Boolean

  'This routine assumes that the Dims are in the one per line format.
  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, DetectDimDuplicate) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = ArrProc(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If IsDimLine(L_CodeLine) Then
            arrLine = Split(L_CodeLine)
            IsPrivate = IsDeclaration(CStr(arrLine(1)), "Private", CompName)
            If IsPrivate Then
              If UBound(arrLine) > 1 Then
                IsIndexed = GetDeclarationIndexing(CStr(arrLine(1)), "Private", CompName) = -1
                If GetLeftBracketPos(arrLine(2)) Then
                  IsPrivate = IsIndexed
                 ElseIf GetLeftBracketPos(arrLine(2)) = 0 Then
                  IsPrivate = False
                  WarnDupNameButIndexed = True
                End If
              End If
            End If
            If CompType <> vbext_ct_ActiveXDesigner Then
              If CompType <> vbext_ct_ClassModule Then
                'These modules do not have access to public variables so ignore then
                IsPublicDec = IsDeclaration(CStr(arrLine(1)), "Public")
                If IsPublicDec Then
                  If UBound(arrLine) > 1 Then
                    IsIndexed = GetDeclarationIndexing(CStr(arrLine(1))) = -1
                    If GetLeftBracketPos(arrLine(2)) Then
                      IsPublicDec = IsIndexed
                     ElseIf GetLeftBracketPos(arrLine(2)) = 0 Then
                      IsPrivate = False
                      WarnDupNameButIndexed = True
                    End If
                  End If
                End If
              End If
            End If
            If IsPublicDec Or IsPrivate Then
              AddNfix DetectDimDuplicate
              Select Case FixData(DetectDimDuplicate).FixLevel
               Case CommentOnly
                ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Duplicate Variable" & strInSQuotes(arrLine(1), True) & "already exists as a " & IIf(IsPublicDec, "Public", "Private") & " variable, so should be deleted OR " & SuggestNewName(arrLine(1), ArrProc(I)), MAfter)
               Case FixAndComment, JustFix
                If FixData(DetectDimDuplicate).FixLevel = FixAndComment Then
                  ArrProc(I) = SmartMarker(ArrProc, I, WARNING_MSG & "Duplicate Variable" & strInSQuotes(arrLine(1), True) & "already exists as " & IIf(IsPublicDec, "Public", "Private") & " variable, so should be deleted OR " & SuggestNewName(arrLine(1), ArrProc(I)), MAfter)
                End If
              End Select
              WarnDupNameButIndexed = False
             ElseIf WarnDupNameButIndexed Then
              Select Case FixData(DetectDimDuplicate).FixLevel
               Case CommentOnly, FixAndComment, JustFix
                If Not Xcheck(XIgnoreCom) Then
                  ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Duplicate Variable." & IIf(IsPublicDec, "Public", "Private") & " Indexed Variable" & strInSQuotes(arrLine(1), True) & "already exists." & vbNewLine & _
                   "This is legal but makes code less readable." & SuggestNewName(arrLine(1), ArrProc(I)), MAfter)
                  WarnDupNameButIndexed = False
                End If
              End Select
            End If
            If Not Xcheck(XIgnoreCom) Then
              If CntrlDescMember(arrLine(1)) > -1 Then
                If Len(CntrlDesc(CntrlDescMember(arrLine(1))).CDForm) Then
                  ArrProc(I) = Marker(ArrProc(I), WARNING_MSG & "Dims with same name as a Control(" & ControlDescription(CntrlDescMember(arrLine(1))) & ")" & vbNewLine & _
                   WARNING_MSG & " make code difficult to read." & SuggestNewName(arrLine(1), ArrProc(I)) & DoNotSandRMsg("Control"), MAfter)
                End If
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimDuplicate = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Function DimExpandMulti(ByVal ModuleNumber As Long, _
                               ByVal strA As String) As String

  Dim Cline        As String
  Dim CommentStore As String
  Dim CommaPos     As Long
  Dim I            As Long
  Dim ArrProc      As Variant
  Dim MaxFactor    As Long
  Dim DimStatic    As String
  Dim Guard        As String
  Dim colonPos     As Long

  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, ReWriteDimExpandMultiDim) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > 0 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        If IsDimLine(ArrProc(I)) Then
          DimStatic = LeftWord(ArrProc(I)) & SngSpace
          If InstrAtPositionArray(ArrProc(I), ipAny, False, ",", ":") Then
            Cline = ArrProc(I)
            If ExtractCode(Cline, CommentStore) Then
              Cline = ConcealParameterCommas(Cline)
              Guard = Cline
              CommaPos = 0
              colonPos = 0
              Do
                'Very rare but possible Dim X: Dim Y
                colonPos = InStr(colonPos + 1, Cline, ": ")
                If colonPos Then
                  If Not EnclosedInBrackets(Cline, colonPos) Then
                    Cline = Left$(Cline, colonPos - 1) & vbNewLine & DimStatic & Mid$(Cline, CommaPos + 2)
                  End If
                End If
              Loop While colonPos
              Do
                CommaPos = GetCommaSpacePos(Cline, CommaPos + 1)
                If CommaPos Then
                  If Not EnclosedInBrackets(Cline, CommaPos) Then
                    Cline = Left$(Cline, CommaPos - 1) & vbNewLine & DimStatic & Mid$(Cline, CommaPos + 2)
                  End If
                End If
              Loop While CommaPos
              If LenB(CommentStore) Then
                Cline = Safe_Replace(Cline, vbNewLine, CommentStore & vbNewLine, , 1)
              End If
              If Guard <> Cline Then
                AddNfix ReWriteDimExpandMultiDim
                Select Case FixData(ReWriteDimExpandMultiDim).FixLevel
                 Case CommentOnly
                  ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Multiple Dim lines are more difficult to read", MAfter)
                 Case FixAndComment
                  ArrProc(I) = Cline
                  ArrProc(I) = SmartMarker(ArrProc, I, UPDATED_MSG & "Multiple Dim line separated" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & Guard, vbNullString), MAfter)
                 Case JustFix
                  ArrProc(I) = Cline
                End Select
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimExpandMulti = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Function DimFormat(ByVal strA As String) As String

  Dim ArrProc       As Variant
  Dim I             As Long
  Dim DimCount      As Long
  Dim L_CodeLine    As String
  Dim CommentStore  As String
  Dim EOLOffSet     As Long
  Dim TmpLen        As Long
  Dim AsOffSet      As Long
  Dim MaxFactor     As Long
  Dim TPos          As Long
  Dim lngLongestDim As Long

  EOLOffSet = 0
  AsOffSet = 0
  ArrProc = Split(strA, vbNewLine)
  'v3.0.4 added test to skip complex testing
  If InStr(strA, " As ") Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > 0 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = ArrProc(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If IsDimLine(L_CodeLine, False) Then
            ExtractCode L_CodeLine, CommentStore
            'v2.2.9 stop extra spaces driving format mad
            Do While InStr(L_CodeLine, "  As ")
              L_CodeLine = Replace$(L_CodeLine, "  As ", " As ")
            Loop
            DimCount = DimCount + 1
            TmpLen = Get_As_Pos(L_CodeLine)
            If InCode(L_CodeLine, TmpLen) Then
              If TmpLen > AsOffSet Then
                AsOffSet = TmpLen
              End If
            End If
            'v2.2.9 get comments lined up properly
            If lngLongestDim < Len(L_CodeLine) Then
              lngLongestDim = Len(L_CodeLine) + 3
            End If
            If LenB(CommentStore) Then
              EOLOffSet = lngLongestDim
            End If
          End If
        End If
      Next I
      If DimCount > 1 Then
        If AsOffSet Then
          For I = LBound(ArrProc) + 1 To UBound(ArrProc)
            L_CodeLine = ArrProc(I)
            'v2.2.9 stop extra spaces driving format mad
            If IsDimLine(L_CodeLine, False) Then
              Do While InStr(L_CodeLine, "  As ")
                L_CodeLine = Replace$(L_CodeLine, "  As ", " As ")
              Loop
              TPos = Get_As_Pos(L_CodeLine)
              If TPos <> AsOffSet Then
                If InCode(L_CodeLine, TPos) Then
                  L_CodeLine = Safe_Replace(L_CodeLine, " As ", Space$(Abs(1 + AsOffSet - TPos)) & "As ", , 1)
                End If
              End If
              ArrProc(I) = L_CodeLine
            End If
          Next I
        End If
        If EOLOffSet Then
          For I = LBound(ArrProc) + 1 To UBound(ArrProc)
            L_CodeLine = ArrProc(I)
            If IsDimLine(L_CodeLine, False) Then
              ExtractCode L_CodeLine, CommentStore
              If LenB(CommentStore) Then
                L_CodeLine = L_CodeLine & Space$(Abs(EOLOffSet - Len(L_CodeLine))) & Trim$(CommentStore)
                ArrProc(I) = L_CodeLine
              End If
            End If
          Next I
        End If
      End If
    End If
  End If
  DimFormat = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Function DimMissing(ByVal ModuleNumber As Long, _
                            ByVal strProc As String, _
                            ByVal CompName As String) As String

  Dim MaxFactor        As Long
  Dim arrLine          As Variant
  Dim I                As Long
  Dim J                As Long
  Dim MissingDims      As String
  Dim ExistingDimArray As Variant
  Dim L_CodeLine       As String
  Dim ArrProc          As Variant
  Dim LTopLine         As Long

  ExistingDimArray = BuildExistingDimArray(strProc)
  ArrProc = Split(strProc, vbNewLine)
  MaxFactor = UBound(ArrProc)
  If MaxFactor > 0 Then
    LTopLine = GetProcCodeLineOfRoutine(ArrProc, True)
    For I = LTopLine To MaxFactor - 1
      MemberMessage "", I, MaxFactor
      L_CodeLine = strCodeOnly(ArrProc(I))
      If Not JustACommentOrBlankOrDimLine(L_CodeLine) Then
        If Not isProcHead(L_CodeLine) Then
          arrLine = Split(Trim$(ExpandForDetection(L_CodeLine)))
          If UBound(arrLine) > 1 Then
            If arrLine(0) = "For" Then
              If arrLine(1) <> "Each" Then
                DimMissingListGenerator arrLine(1), strProc, CompName, ExistingDimArray, ModuleNumber, MissingDims
               Else
                DimMissingListGenerator arrLine(2), strProc, CompName, ExistingDimArray, ModuleNumber, MissingDims
              End If
             ElseIf arrLine(0) = "Set" Then
              If arrLine(3) <> "Nothing" Then
                DimMissingListGenerator arrLine(1), strProc, CompName, ExistingDimArray, ModuleNumber, MissingDims
              End If
             ElseIf UBound(arrLine) > 0 Then
              If arrLine(1) = "=" Then
                If Not Left$(arrLine(0), 1) = DQuote Then
                  If Not InQSortArray(ArrQVBReservedWords, arrLine(0)) Then
                    DimMissingListGenerator arrLine(0), strProc, CompName, ExistingDimArray, ModuleNumber, MissingDims
                  End If
                End If
               ElseIf (arrLine(0) = "Line" And arrLine(1) = "Input") Or arrLine(0) = "Input" Then
                For J = IIf(arrLine(0) = "Input", 3, 4) To UBound(arrLine)
                  If InStr(arrLine(J), SQuote) Then
                    Exit For
                  End If
                  DimMissingListGenerator arrLine(J), strProc, CompName, ExistingDimArray, ModuleNumber, MissingDims
                Next J
              End If
            End If
          End If
        End If
      End If
    Next I
    If LenB(MissingDims) Then
      MissingDims = "Dim " & MissingDims
      ArrProc(LTopLine) = ArrProc(LTopLine) & vbNewLine & _
       MissingDims & IIf(Xcheck(XVerbose), vbNewLine & _
       RGSignature & "May be Control name using default property or a control(probably Form) Property being used with default assignment (not a good idea; be explicit)", vbNullString)
      AddNfix DetectDimMissing
      MissingDims = vbNullString
    End If
  End If
  DimMissing = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Sub DimMissingListGenerator(VarTrigger As Variant, _
                                    strProc As String, _
                                    CompName As String, _
                                    ExistingDimArray As Variant, _
                                    ModuleNumber As Long, _
                                    MissingDims As String)

  Dim AsType           As String

  'ver 1.1.93
  'support procedure for DimMissing
  If InStr(MissingDims, SpacePad(VarTrigger)) = 0 Then  'ignore repeats
    If Not IsNumeric(VarTrigger) Then                      'ignore numbers
      If Not IsPunct(CStr(VarTrigger)) Then                  'ignore punctuation
        If InstrArray(VarTrigger, ".", "!") = 0 Then
          'ignore object references
          If Not AlreadyDimmed(CompName, VarTrigger, ExistingDimArray) Then
            AsType = GetDimTypeFromInternalEvidence(ModuleNumber, Split(strProc, vbNewLine), VarTrigger, CompName)
            If dofix(ModuleNumber, UpdateInteger2Long) Then
              AsType = Replace$(AsType, "Integer", "Long")
            End If
            Safe_AsTypeAdd VarTrigger, AsType
            MissingDims = AccumulatorString(MissingDims, VarTrigger & SngSpace & RGSignature & "Missing Dim Auto-inserted" & LBracket & IIf(Len(AsType), IIf(UsingDefTypes, "Auto-TypedfromDefType", "Auto-Type may not be correct"), "Unable to Auto-Type") & RBracket, vbNewLine & _
             "Dim ")
          End If
        End If
      End If
    End If
  End If

End Sub

Public Function DimMoveToTop(ByVal strA As String) As String

  Dim strDecType   As String
  Dim MaxFactor    As Long
  Dim ArrProc      As Variant
  Dim I            As Long
  Dim TopOfRoutine As Long
  Dim bConstFound  As Boolean

  ArrProc = Split(strA, vbNewLine)
  MaxFactor = UBound(ArrProc)
  If MaxFactor > -1 Then
    TopOfRoutine = FirstCodeLineInProc(ArrProc)
    For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
      MemberMessage "", MaxFactor / 3 + I / 3, MaxFactor
      If Left$(ArrProc(I), 7) = "#End If" Then
        If I > TopOfRoutine Then
          TopOfRoutine = I
        End If
      End If
    Next I
    bConstFound = False
    For I = TopOfRoutine To MaxFactor
      MemberMessage "", (MaxFactor / 3) * 2 + I / 3, MaxFactor
      If IsDimLine(ArrProc(I), False) Then
        ' don't move dims around each other
        If I - 1 >= TopOfRoutine Then
          If Not IsDimLine(ArrProc(I - 1), False) Then
            strDecType = LeftWord(ArrProc(I))
            'v2.4.4 reconfigured to short circuit
            If strDecType <> "Const" Then
              If InStr(ArrProc(I), " Const ") Then
                strDecType = strDecType & " Const"
              End If
            End If
            If InStr(ArrProc(I), "Const ") Then
              bConstFound = True
            End If
            ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I)
            ArrProc(I) = vbNullString
            'v 2.0.8 move upgrade integer comments with the Dim line
            '"Integer " & strDecType & " could
            If ArrProc(I + 1) = SUGGESTION_MSG & "Integer " & strDecType & " could be upgraded to Long." Then
              ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I + 1)
              ArrProc(I + 1) = vbNullString
            End If
            If ArrProc(I + 1) = WARNING_MSG & "Integer " & strDecType & " upgraded to Long." Then
              ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I + 1)
              ArrProc(I + 1) = vbNullString
            End If
            AddNfix MoveDim2Top
          End If
        End If
      End If
    Next I
    'v2.9.6 move Consts to top of Dim section incase they are used to set 'Dim Aray(Const)
    If bConstFound Then ' only hits if Cosnts exist (fairly rare)
      For I = TopOfRoutine To MaxFactor
        MemberMessage "", (MaxFactor / 3) * 2 + I / 3, MaxFactor
        If IsDimLine(ArrProc(I), False) Then
          If InStr(ArrProc(I), "Const ") Then
            ' don't move dims around each other
            If I - 1 > TopOfRoutine Then
              If Not IsDimLine(ArrProc(I - 1), False) Then
                strDecType = LeftWord(ArrProc(I))
                'v2.4.4 reconfigured to short circuit
                If strDecType <> "Const" Then
                  If InStr(ArrProc(I), " Const ") Then
                    strDecType = strDecType & " Const"
                  End If
                End If
                If InStr(ArrProc(I), " Const ") Then
                  bConstFound = True
                End If
                ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I)
                ArrProc(I) = vbNullString
                'v 2.0.8 move upgrade integer comments with the Dim line
                '"Integer " & strDecType & " could
                If ArrProc(I + 1) = SUGGESTION_MSG & "Integer " & strDecType & " could be upgraded to Long." Then
                  ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I + 1)
                  ArrProc(I + 1) = vbNullString
                End If
                If ArrProc(I + 1) = WARNING_MSG & "Integer " & strDecType & " upgraded to Long." Then
                  ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & ArrProc(I + 1)
                  ArrProc(I + 1) = vbNullString
                End If
                AddNfix MoveDim2Top
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimMoveToTop = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Function DimMultiSingleTypeing(ByVal ModuleNumber As Long, _
                                      ByVal strA As String) As String

  Dim arrLine      As Variant
  Dim ArrProc      As Variant
  Dim TypeDef      As String
  Dim I            As Long
  Dim J            As Long
  Dim MaxFactor    As Long
  Dim CommentStore As String
  Dim L_CodeLine   As String

  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, ReWriteDimMultiSingleType) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        If Not JustACommentOrBlank(ArrProc(I)) Then
          If IsDimLine(ArrProc(I)) Then
            L_CodeLine = ArrProc(I)
            If ExtractCode(L_CodeLine, CommentStore) Then
              If InStr(L_CodeLine, ",") Then
                If CountSubString(L_CodeLine, " As ") = 1 Then
                  L_CodeLine = ConcealParameterCommas(L_CodeLine)
                  arrLine = Split(L_CodeLine, CommaSpace)
                  If UBound(arrLine) > 0 Then
                    TypeDef = arrLine(UBound(arrLine))
                    If Get_As_Pos(TypeDef) > 0 Then
                      TypeDef = GetType(TypeDef)
                      For J = LBound(arrLine) To UBound(arrLine) - 1
                        If Get_As_Pos(arrLine(J)) = 0 Then
                          If Not TypeSuffixExists(arrLine(J)) Then
                            'v 2.1.9 Thanks Mike Ulik: the " As " was missing, d'oh
                            arrLine(J) = arrLine(J) & " As " & TypeDef
                          End If
                        End If
                      Next J
                      AddNfix ReWriteDimMultiSingleType
                      Select Case FixData(ReWriteDimMultiSingleType).FixLevel
                       Case CommentOnly
                        ArrProc(I) = L_CodeLine & CommentStore
                        ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Declaring a whole line with single As Type is no longer supported. Type each Item separately.", MAfter)
                       Case FixAndComment
                        ArrProc(I) = Join(arrLine, CommaSpace) & CommentStore
                        ArrProc(I) = SmartMarker(ArrProc, I, UPDATED_MSG & "A whole line with single As Type is no longer supported." & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & ArrProc(I) & vbNewLine, vbNullString), MAfter)
                       Case JustFix
                        ArrProc(I) = Join(arrLine, CommaSpace)
                      End Select
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimMultiSingleTypeing = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Sub DimTypeApply(ByVal ModuleNumber As Long, _
                         VarRoutine As Variant, _
                         ByVal DimLine As Long, _
                         ByVal strVarname As String, _
                         ByVal ReturnCannotDoItMsg As Boolean, _
                         FoundIt As Boolean, _
                         ByVal CompName As String)

  Dim AsType       As String
  Dim CommentStore As String
  Dim strTmp       As String

  strTmp = VarRoutine(DimLine)
  If ExtractCode(strTmp, CommentStore) Then
    If Left$(strTmp, 6) = "Const " Then
      VarRoutine(DimLine) = ConstantExpander(strTmp, False) & CommentStore
      FoundIt = True
     ElseIf Left$(strTmp, 6) = "ReDim " Then
      If getReDimType(WordInString(ExpandForDetection(strTmp), 2), VarRoutine, AsType) Then
        Safe_AsTypeAdd VarRoutine(DimLine), " As " & AsType
        VarRoutine(DimLine) = SmartMarker(VarRoutine, DimLine, WARNING_MSG & "ReDim explicitly Type-cast for easier reading.", MAfter)
        FoundIt = True
       Else
        VarRoutine(DimLine) = SmartMarker(VarRoutine, DimLine, SUGGESTION_MSG & "ReDim is easier to read if you explicitly Type-cast it", MAfter)
        FoundIt = True
      End If
     Else
      AsType = GetDimTypeFromInternalEvidence(ModuleNumber, VarRoutine, strVarname, CompName)
      If ReturnCannotDoItMsg Or LenB(AsType) > 0 Then
        FoundIt = True
        AddNfix DetectReWriteDimUntyped
        Select Case FixData(DetectReWriteDimUntyped).FixLevel
         Case CommentOnly
          VarRoutine(DimLine) = SmartMarker(VarRoutine, DimLine, SUGGESTION_MSG & "UnTyped Dim " & IIf(Len(AsType), VarRoutine(DimLine) & AsType & " is suggested", " detected"), MAfter)
         Case FixAndComment, JustFix
          'v2.3.1 insert " As " properly This time for sure
          Safe_AsTypeAdd VarRoutine(DimLine), AsType
          VarRoutine(DimLine) = SmartMarker(VarRoutine, DimLine, WARNING_MSG & "Untyped Dim. " & IIf(Len(AsType), IIf(UsingDefTypes, vbNullString, "Auto-Type may not be correct"), "Type could not be determined"), MEoL)
         Case JustFix
          If Len(AsType) Then
            Safe_AsTypeAdd VarRoutine(DimLine), AsType
           Else
            VarRoutine(DimLine) = SmartMarker(VarRoutine, DimLine, WARNING_MSG & "Untyped Dim. " & IIf(Len(AsType), IIf(UsingDefTypes, vbNullString, "Auto-Type may not be correct"), "Type could not be determined"), MEoL)
          End If
        End Select
      End If
    End If
  End If

End Sub

Public Function DimTypeUpdate(ByVal ModuleNumber As Long, _
                              ByVal strA As String) As String

  Dim I          As Long
  Dim ArrProc    As Variant
  Dim arrLine    As Variant
  Dim FixNo      As Long
  Dim MaxFactor  As Long
  Dim L_CodeLine As String

  'This routine assumes that the Dims are in the one per line format.
  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, UpdateDimType) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = ArrProc(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If IsDimLine(L_CodeLine) Then
            If Get_As_Pos(L_CodeLine) = 0 Then
              L_CodeLine = TypeSuffixExtender(L_CodeLine, FixNo)
              If ArrProc(I) <> L_CodeLine Then
                Select Case FixData(UpdateDimType).FixLevel
                 Case CommentOnly
                  ArrProc(I) = SmartMarker(ArrProc, I, SUGGESTION_MSG & "Obsolete Type Suffix (" & TypeSuffixArray(FixNo) & ") should be replaced with " & AsTypeArray(FixNo), MAfter)
                 Case FixAndComment
                  ArrProc(I) = Left$(L_CodeLine, InStr(L_CodeLine, vbNewLine) - 1)
                  ArrProc(I) = SmartMarker(ArrProc, I, UPDATED_MSG & "Obsolete Type Suffix (" & TypeSuffixArray(FixNo) & ") replaced with " & AsTypeArray(FixNo), MAfter)
                 Case JustFix
                  ArrProc(I) = Join(arrLine)
                End Select
              End If
            End If
          End If
        End If
      Next I
    End If
  End If
  DimTypeUpdate = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Function DimUnneededInitialize(ByVal ModuleNumber As Long, _
                                      ByVal strA As String) As String

  Dim L_CodeLine   As String
  Dim CommentStore As String
  Dim MaxFactor    As Long
  Dim strTest      As String
  Dim strType      As String
  Dim strInitVal   As String
  Dim ArrProc      As Variant
  Dim I            As Long
  Dim J            As Long
  Dim lngJunk      As Long

  'v2.9.0 new fix
  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, NoDimInitialise) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > 0 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = ArrProc(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If IsDimLine(L_CodeLine, False) Then
            ExtractCode L_CodeLine, CommentStore
            If InStr(L_CodeLine, " As ") Then
              strTest = WordInString(L_CodeLine, 2)
              strType = WordInString(L_CodeLine, 4)
              For J = I + 1 To MaxFactor
                If InStr(strCodeOnly(ArrProc(J)), " ") = 0 And Right$(ArrProc(J), 1) = ":" Or IsNumeric(ArrProc(J)) Then
                  'looks like a GoTo target if code might goto back before the null initialization don't fix
                  Exit For
                 Else
                  If InStrWholeWord(ArrProc(J), strTest) Then
                    If InStructure(LoopStruct, ArrProc, J, lngJunk, lngJunk) Or InStructure(ForStruct, ArrProc, J, lngJunk, lngJunk) Then
                      'if code Loops past the null initialization don't fix
                      Exit For
                     Else
                      If SmartLeft(ArrProc(J), strTest & " = ") Then
                        strInitVal = WordInString(ArrProc(J), 3)
                        'v2.9.9 added safety in case code has '= 0 + variable' or '= "" & StringVar '
                        ' a trick used to force Null database fields to fit Typed variables for instance
                        If strInitVal = LastWord(strCodeOnly(ArrProc(J))) Then
                          Select Case strInitVal
                           Case "0"
                            Select Case strType
                             Case "Integer", "Byte", "Long", "Single", "Double"
                              ArrProc(J) = WARNING_MSG & "Unneeded Dim initialization removed '" & ArrProc(J)
                              Exit For
                            End Select
                           Case EmptyString
                            If strType = "String" Then
                              ArrProc(J) = WARNING_MSG & "Unneeded Dim initialization removed '" & ArrProc(J)
                              Exit For
                            End If
                           Case "False"
                            If strType = "Boolean" Then
                              ArrProc(J) = WARNING_MSG & "Unneeded Dim initialization removed '" & ArrProc(J)
                              Exit For
                            End If
                           Case Else
                            Exit For
                          End Select
                         Else
                          Exit For
                        End If
                      End If
                    End If
                  End If
                End If
              Next J
            End If
          End If
        End If
      Next I
    End If
  End If
  DimUnneededInitialize = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Function DimUntyped(ByVal ModuleNumber As Long, _
                            ByVal strA As String, _
                            ByVal CompName As String) As String

  Dim strVarname   As String
  Dim FoundOne     As Boolean
  Dim ArrProc      As Variant
  Dim MaxFactor    As Long
  Dim firstHitLine As Long
  Dim I            As Long
  Dim J            As Long

  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, DetectReWriteDimUntyped) Then
    MaxFactor = UBound(ArrProc)
    If MaxFactor > -1 Then
      For I = GetProcCodeLineOfRoutine(ArrProc) To MaxFactor
        MemberMessage "", I, MaxFactor
        If IsDimLine(ArrProc(I)) Then
          If Get_As_Pos(ArrProc(I)) = 0 Then
            strVarname = CStr(Split(Trim$(ExpandForDetection(ArrProc(I))))(1))
            FoundOne = False
            For J = 1 To MaxFactor - 1
              If J <> I Then
                If Not JustACommentOrBlankOrDimLine(ArrProc(J)) Then
                  If InstrAtPositionArray(ArrProc(J), ipAny, True, strVarname, strVarname & RBracket, strVarname & LBracket, LBracket & strVarname) Or InStr(ArrProc(J), SngSpace & strVarname) Or InStr(ArrProc(J), strVarname & ".") Then
                    'this should allow it to keep searching if it cannot work it out on first try
                    DimTypeApply ModuleNumber, ArrProc, I, strVarname, False, FoundOne, CompName
                    If FoundOne Then
                      Exit For
                     Else
                      firstHitLine = J
                    End If
                  End If
                End If
              End If
            Next J
            If Not FoundOne Then
              If firstHitLine > 0 Then
                DimTypeApply ModuleNumber, ArrProc, I, strVarname, True, FoundOne, CompName
                'Else It is unused
                'Don't worry the DimDead routine will get it
              End If
            End If
          End If
        End If
        'End If
      Next I
    End If
  End If
  DimUntyped = Join(CleanArray(ArrProc), vbNewLine)

End Function

Public Sub DimUsageDeleteNonProblem()

  Dim strBase     As String
  Dim I           As Long
  Dim MyhourGlass As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  strBase = RGSignature & "|Dim Usage:("
  For I = 4 To 39
    'effectively deletes all with count above 3 ( 4 also gets 40,41 etc
    mObjDoc.KillComments strBase & I, , True
  Next I
  mObjDoc.KillComments RGSignature & "|Dim Usage:(3) on (3)", , True

End Sub

Private Sub FunctionParameterMessage(TmpA As Variant, _
                                     Vname As Variant, _
                                     ByVal I As Long, _
                                     ByVal UnusedFunctionNoParameters As Boolean, _
                                     ByVal FuncType As Long, _
                                     Optional ByVal FixMsg As Boolean = False, _
                                     Optional ByVal BMsgBoxProblem As Boolean = False)

  Dim strFuncType As String
  Dim strSuggest  As String

  Select Case FuncType
   Case 0, 2
    strFuncType = IIf(FuncType = 0, "an API Function", "a User Defined Function")
    If Not UnusedFunctionNoParameters Then
      If FixMsg Then
        If FixData(UnNeededCall).FixLevel = FixAndComment Then
          strSuggest = SUGGESTION_MSG & "The Function may be setting one of its parameters for use in code." & WARNING_MSG & "Code Fixer has replaced assignment with a direct call to the Function"
         Else
          strSuggest = SUGGESTION_MSG & "The Function may be setting one of its parameters for use in code." & WARNING_MSG & "Code Fixer has replaced assignment with a 'Call' command."
        End If
       Else
        strSuggest = SUGGESTION_MSG & "The Function may be setting one of its parameters for use in code." & SUGGESTION_MSG & "Remove the assignment ('" & Vname & " = ') and the brackets around the parameters." & vbNewLine & _
         "OR replace assignment with 'Call' command."
      End If
     Else
      strSuggest = SUGGESTION_MSG & "The Function has no parameters that may be being set so the calling line can safely be deleted."
    End If
   Case 1
    If BMsgBoxProblem Then
      strFuncType = "a MsgBox "
      strSuggest = SUGGESTION_MSG & "The MsgBox could be called with a 'Call' command or remove the assignment and brackets"
     Else
      strFuncType = "a VB Function"
      strSuggest = SUGGESTION_MSG & "The Function has no parameters that may be being set so the calling line can safely be deleted."
    End If
  End Select
  TmpA(I) = SmartMarker(TmpA, I, WARNING_MSG & "Variable is assigned a Return value from " & strFuncType & " call but never used in code." & strSuggest, MAfter)

End Sub

Public Function GetDeclarationID(ByVal strTest As String, _
                                 Optional ByVal StrScope As String, _
                                 Optional ByVal strClass As String, _
                                 Optional ByVal strOnForm As String) As Long

  Dim I            As Long
  Dim strScopePlus As String

  If bDeclExists Then
    If StrScope = "Public" Then
      strScopePlus = "Friend"
     ElseIf StrScope = "Private" Then
      strScopePlus = "Static"
    End If
    For I = LBound(DeclarDesc) To UBound(DeclarDesc)
      If LenB(strOnForm) = 0 Or strOnForm = DeclarDesc(I).DDComp Then
        If LenB(strClass) = 0 Or strClass = DeclarDesc(I).DDType Then
          If LenB(StrScope) = 0 Or StrScope = DeclarDesc(I).DDScope Or strScopePlus = DeclarDesc(I).DDScope Then
            If strTest = DeclarDesc(I).DDName Then
              GetDeclarationID = I
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function GetDeclarationIndexing(ByVal strTest As String, _
                                       Optional ByVal StrScope As String, _
                                       Optional ByVal strClass As String, _
                                       Optional ByVal strOnForm As String) As Long

  Dim I            As Long
  Dim strScopePlus As String

  If bDeclExists Then
    If StrScope = "Public" Then
      strScopePlus = "Friend"
     ElseIf StrScope = "Private" Then
      strScopePlus = "Static"
    End If
    For I = LBound(DeclarDesc) To UBound(DeclarDesc)
      If LenB(strOnForm) = 0 Or strOnForm = DeclarDesc(I).DDComp Then
        If LenB(strClass) = 0 Or strClass = DeclarDesc(I).DDType Then
          If LenB(StrScope) = 0 Or StrScope = DeclarDesc(I).DDScope Or strScopePlus = DeclarDesc(I).DDScope Then
            If strTest = DeclarDesc(I).DDName Then
              GetDeclarationIndexing = DeclarDesc(I).DDIndexing
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Private Function getReDimType(ByVal strFind As String, _
                              VarRoutine As Variant, _
                              strType As String) As Boolean

  Dim I        As Long
  Dim J        As Long
  Dim ParArray As Variant

  'v2.2.9 type-cast ReDim if it is declared at procedure level
  'v2.3.2 type-cast ReDim if it is declared in Procedure paramaters
  'Procedure paramaters
  For I = LBound(VarRoutine) To UBound(VarRoutine)
    If isProcHead(VarRoutine(I)) Then
      ParArray = ParamArrays(VarRoutine(I), FullParam)
      If Not IsEmpty(ParArray) Then
        For J = LBound(ParArray) To UBound(ParArray)
          If InstrAtPosition(ParArray(J), strFind, ipAny) Then
            strType = WordAfter(ParArray(J), " As ")
            If Len(strType) Then
              getReDimType = True
              Exit For
            End If
          End If
        Next J
      End If
      If Len(strType) Then
        getReDimType = True
        Exit For
      End If
    End If
    'Local Dim line
    If SmartLeft(VarRoutine(I), "Dim " & strFind) Then
      'v2.6.4 improved test to avoid partial matching
      If WordInString(ExpandForDetection(VarRoutine(I)), 2) = strFind Then
        '      If WordAfter(ExpandForDetection(VarRoutine(I)), "Dim") = strFind Then
        strType = WordAfter(VarRoutine(I), " As ")
        'v2.6.4 safety for objects being dimmed(may not be necessary
        ' I think it was an artifact of the partial match error
        If strType = "New" Then
          strType = WordAfter(VarRoutine(I), " New ")
        End If
        If Len(strType) Then
          getReDimType = True
          Exit For
        End If
      End If
    End If
  Next I

End Function

Private Function InTypeSomeFormat(ByVal varTest As Variant) As Boolean

  InTypeSomeFormat = CntrlDescMember(varTest) > -1
  If Not InTypeSomeFormat Then
    'test for type suffix in code and As Type in Dims
    If TypeSuffixExists(varTest) Then
      InTypeSomeFormat = CntrlDescMember(Left$(varTest, Len(varTest) - 1)) > -1
    End If
  End If

End Function

Private Function isConnectorBang(vCodePart As Variant, _
                                 ByVal strSuf As String) As Boolean

  Dim Pos     As Long
  Dim strNext As String

  'ver 2.0.1
  'Thanks Tom Law
  'bug fix for bang (!) connectors
  'in DB references; RecordSet![FieldName]
  'and  those who still use it fo control referencing; Form!Control
  If strSuf <> "!" Then
    isConnectorBang = False
   Else
    Pos = InStr(vCodePart, strSuf)
    If Pos = Len(vCodePart) Then
      isConnectorBang = False
     Else
      strNext = Mid$(vCodePart, Pos + 1, 1)
      If IsAlphaIntl(strNext) Or strNext = "[" Then
        isConnectorBang = True
      End If
    End If
  End If

End Function

Public Function IsControlName(varTest As Variant) As Boolean

  IsControlName = InQSortArray(arrQCtrlPresence, varTest)

End Function

Public Function IsDeclareName(varTest As Variant) As Boolean

  IsDeclareName = InQSortArray(arrQDeclarPresence, varTest)

End Function

Private Function IsEmbeddedPunct(strCode As String, _
                                 VarPunct As Variant) As Boolean

  Dim Pos As Long

  Pos = InStr(strCode, VarPunct)
  If Pos > 1 Then
    If Pos < Len(strCode) Then
      If Not IsPunct(Mid$(strCode, Pos - 1)) Then
        If Not IsPunct(Mid$(strCode, Pos + 1, 1)) Then
          IsEmbeddedPunct = True
        End If
      End If
    End If
  End If

End Function

Private Function IsFormProcCall(ByVal strTest As String) As Boolean

  Dim arrParts As Variant

  If InStr(strTest, ".") Then
    arrParts = Split(strTest, ".")
    IsFormProcCall = ProcDescMember(arrParts(1), arrParts(0)) > -1
  End If

End Function

Private Function IsHexValue(strCode As Variant) As Boolean

  If Left$(strCode, 2) = "&H" Then
    IsHexValue = True
  End If

End Function

Private Function MoreThanOnceInLine(strCode As String, _
                                    arrDim As Variant) As Boolean

  Dim I       As Long
  Dim J       As Long
  Dim Hit     As Long
  Dim arrTest As Variant

  'v2.2.6 Thanks Tom Law
  'stops the Assign only fix hitting if the targeted Dim occurs twice in the line
  'NOTE if the Dim is used in an API call in the format 'X = APICall(a,b,X)'
  'and the value of X has not been set previously then you can probably replace the X with '0" or '0&'
  'and then the Assign only fix could be applied (CF throw a message to this effect if an API is detected)
  arrTest = Split(ExpandForDetection(strCode))
  For I = LBound(arrDim) To UBound(arrDim)
    For J = LBound(arrTest) To UBound(arrTest)
      If arrTest(J) = arrDim(I) Then
        Hit = Hit + 1
        If Hit > 1 Then
          MoreThanOnceInLine = True
          Exit For
        End If
      End If
    Next J
    If MoreThanOnceInLine Then
      Exit For
    End If
  Next I

End Function

Private Function NeedsTypeSuffixStrip(ByVal strA As String) As Boolean

  'v3.0.4 simple test before going into complex fix
  
  Dim I As Long

  For I = LBound(TypeSuffixArray) To UBound(TypeSuffixArray)
    If InStr(strA, TypeSuffixArray(I)) Then
      NeedsTypeSuffixStrip = True
      Exit For
    End If
  Next I

End Function

Public Function PotentialControlDefault(Comp As VBComponent, _
                                        ByVal arrTmp As Variant, _
                                        ByVal strFind As String) As Boolean

  Dim J As Long

  If ArrayMember(arrTmp(0), "With", "Load", "Set", "Unload", "For", "Next", "RaiseEvent", "End") Then
    'Lines starting with any of these can/should be ignored.
    'ver 2.0.3 move from DimControlDefaultProp (added ver 1.1.20)
    'repaired ignore Next and Unload (Corrected capitalization)
    PotentialControlDefault = False
   ElseIf arrTmp(0) = "Name" Then
    'ver 2.0.3 Thanks to Mike Ulik. He found a bug caused by the fact that
    'Name is a Command and a Property
    'however Property Name is read-only so can never be first word of a code line
    'while Name Command is a Sub so can only be the first word of a line
    'This test assumes that you are not using any default control property values
    'to use the Name command (technically possible but pretty unlikely).
    'I could have just added "Name" to the text above but figured I'd create
    'a seperate test in case there are other such PRoperty/Command clashes I haven't thought of
    PotentialControlDefault = False
   Else
    If Not InStr(strFind, "*") Then
      For J = LBound(arrTmp) To UBound(arrTmp)
        If arrTmp(J) = strFind Then
          PotentialControlDefault = True
          Exit For 'unction
        End If
      Next J
    End If
    If Not PotentialControlDefault Then
      For J = LBound(arrTmp) To UBound(arrTmp)
        'v3.0.0 has keyword
        If InStr(arrTmp(J), strFind) Then
          'v3.0.0 added ignore numbers
          If InStr(arrTmp(J), DQuote) = 0 Then
            'v3.0.0 already has a property
            If InStr(arrTmp(J), strFind & ".") = 0 Then
              If Not InQSortArray(ArrQVBReservedWords, arrTmp(J)) Then
                If Not isRefLibVBCommands(arrTmp(J)) Then
                  If IsControlName(arrTmp(J)) Then
                    PotentialControlDefault = True
                    Exit For 'unction
                  End If
                  'deal with 'Form.Control' code
                  If InStr(arrTmp(J), ".") Then
                    If IsControlName(Mid$(arrTmp(J), InStr(arrTmp(J), ".") + 1)) Then
                      PotentialControlDefault = True
                      Exit For 'unction
                    End If
                  End If
                  If IsControlProperty(arrTmp(J)) Then
                    If IsComponent_ControlHolder(Comp) Then
                      'property without control
                      PotentialControlDefault = True
                      Exit For 'unction
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Next J
    End If
  End If

End Function

Private Sub RemoveTypSuffix(strCode As String, _
                            ByVal strSuf As String)

  Dim arrTmp   As Variant
  Dim I        As Long
  Dim bRemoved As Long

  DisguiseLiteral strCode, strSuf, True
  'v2.8.3 Thanks Joakim Schramm makes clean up work properly
  'arrTmp = ExpandForDetection2(strCode)' this wrongly expanded brackets and blocked detection
  '                                     ' of the suffix inside brackets
  arrTmp = Split(ExpandForDetection(strCode))
  For I = LBound(arrTmp) To UBound(arrTmp)
    If Right$(arrTmp(I), Len(strSuf)) = strSuf Or InStr(arrTmp(I), strSuf) > 1 Then
      ' >1 protects initial # of FreeFile/Input filenumbers
      'ver 2.0.1 Bang fix
      If Not isConnectorBang(arrTmp(I), strSuf) Then
        If Not IsNumeric(Left$(arrTmp(I), Len(arrTmp(I)) - 1)) Then
          If Not IsHexValue(arrTmp(I)) Then
            If InCode(strCode, InStr(strCode, arrTmp(I))) Then 'protect literals
              If Not IsEmbeddedPunct(strCode, strSuf) Then 'ArrTmp(I)) Then
                arrTmp(I) = Replace$(arrTmp(I), strSuf, vbNullString)
                bRemoved = True
              End If
            End If
          End If
        End If
      End If
    End If
  Next I
  If bRemoved Then
    strCode = Join(arrTmp)
  End If
  DisguiseLiteral strCode, strSuf, False

End Sub

Private Function SplitForDetection(ByVal VarStr As Variant) As Variant

  'v2.8.8 thanks to Anele Mbanga whose code let me track down a bug in 'ExpandForDetection'
  'in conjunction with another bug (not recognising Replace as a Str/Var function)
  'the code that now calls this would incorrectly insert extra spaces into string literals
  'which were brackets if the VB command 'Replace$ was in the codeline
  'Thanks also to Joakim who alerted me to this
  'I made it a separate procedure because I use similar code in several places

  SplitForDetection = Split(ExpandForDetection3(VarStr))

End Function

Private Function TypeSuffixStrip(ByVal ModuleNumber As Long, _
                                 ByVal strA As String) As String

  Dim I          As Long
  Dim J          As Long
  Dim L_CodeLine As String
  Dim MaxFactor  As Long
  Dim arrLine    As Variant
  Dim ArrProc    As Variant
  Dim Hit        As Boolean
  Dim strComment As String
  Dim arrTest    As Variant

  arrTest = Array("$", "&")
  ArrProc = Split(strA, vbNewLine)
  If dofix(ModuleNumber, CodeTypeSuffixStrip) Then
    'v3.0.4 added test
    If NeedsTypeSuffixStrip(strA) Then
      If dofix(ModuleNumber, UpdateDimType) And dofix(ModuleNumber, UpdateDecTypeSuffix) Then
        MaxFactor = UBound(ArrProc)
        If MaxFactor > 0 Then
          For I = 1 To MaxFactor
            MemberMessage "", I, MaxFactor
            L_CodeLine = ArrProc(I)
            If Not JustACommentOrBlankOrDimLine(L_CodeLine) Then
              If ExtractCode(L_CodeLine, strComment) Then
                For J = LBound(TypeSuffixArray) To UBound(TypeSuffixArray)
                  If Not IsInArray(TypeSuffixArray(J), arrTest) Then
                    'deal with the 4 TypeSuffixes that you can be sure are type suffixes
                    'end of word
                    If InStr(L_CodeLine, TypeSuffixArray(J)) Then
                      If InCode(L_CodeLine, InStr(L_CodeLine, TypeSuffixArray(J))) Then
                        RemoveTypSuffix L_CodeLine, TypeSuffixArray(J) '& " "
                        ' parameter member
                        RemoveTypSuffix L_CodeLine, TypeSuffixArray(J) & ","
                        'last member of function parameters
                        RemoveTypSuffix L_CodeLine, TypeSuffixArray(J) & RBracket
                        'rarer typecast function calls or explict arrays ie Doit%(Index)
                        RemoveTypSuffix L_CodeLine, TypeSuffixArray(J) & LBracket
                      End If
                    End If
                  End If
                Next J
                If InStr(L_CodeLine, "&") Then
                  If InCode(L_CodeLine, InStr(L_CodeLine, "&")) Then
                    arrLine = SplitForDetection(L_CodeLine)
                    For J = LBound(arrLine) To UBound(arrLine)
                      If arrLine(J) <> "&" Then
                        If Right$(arrLine(J), 1) = "&" Then
                          If Not IsNumeric(Mid$(arrLine(J), Len(arrLine(J)) - 1, 1)) Then
                            If Not IsHexValue(arrLine(J)) Then
                              If InCode(L_CodeLine, InStr(L_CodeLine, arrLine(J))) Then
                                'ignore in literals
                                arrLine(J) = Left$(arrLine(J), Len(arrLine(J)) - 1)
                                Hit = True
                              End If
                            End If
                          End If
                        End If
                      End If
                    Next J
                    If Hit Then
                      L_CodeLine = Join(arrLine)
                      Hit = False
                    End If
                  End If
                End If
                If InStr(L_CodeLine, "$") Then
                  If InCode(L_CodeLine, InStr(L_CodeLine, "$")) Then
                    arrLine = SplitForDetection(L_CodeLine)
                    For J = LBound(arrLine) To UBound(arrLine)
                      If Right$(arrLine(J), 1) = "$" Then
                        If InCode(L_CodeLine, InStr(L_CodeLine, arrLine(J))) Then
                          'ignore in literals
                          If Not InQSortArray(ArrQStrVarFunc, Left$(arrLine(J), Len(arrLine(J)) - 1)) Then
                            arrLine(J) = Left$(arrLine(J), Len(arrLine(J)) - 1)
                            Hit = True
                          End If
                        End If
                      End If
                    Next J
                    If Hit Then
                      L_CodeLine = Join(arrLine)
                      Hit = False
                    End If
                  End If
                End If
                Do While InStrCode(L_CodeLine, ") .")
                  L_CodeLine = CharReplace(L_CodeLine, ").", InStrCode(L_CodeLine, ") ."), 3)
                Loop
                ArrProc(I) = L_CodeLine & strComment
              End If
            End If
          Next I
        End If
      End If
    End If
  End If
  TypeSuffixStrip = Join(CleanArray(ArrProc), vbNewLine)

End Function

Private Sub UnUsedAPIVariable2Call(ArrRoutine As Variant, _
                                   arrTargets As Variant, _
                                   ByVal TriggerLine As Long, _
                                   arrVarName As Variant)

  Dim varPosVar As Variant
  Dim J         As Long

  For J = UBound(ArrRoutine) - 1 To LBound(ArrRoutine) + 1 Step -1
    If IsInArray(J, arrTargets) Then
      If Not IsDimLine(ArrRoutine(J)) Then
        For Each varPosVar In arrVarName
          If Len(varPosVar) Then
            If InstrAtPosition(ArrRoutine(J), varPosVar, ipLeftOr2nd, True) Then
              'ver 1.1.81 cope with TypeSuffix( not deleted when this fix hits)
              'in previous version it took 2 runs to complete the fix
              If InStr(ArrRoutine(J), varPosVar & EqualInCode) Then
                ArrRoutine(J) = Replace$(ArrRoutine(J), varPosVar & EqualInCode, "Call ")
                If FixData(UnNeededCall).FixLevel = FixAndComment Then
                  ArrRoutine(J) = CallRemoval(ArrRoutine(J))
                  ArrRoutine(J) = SmartMarker(ArrRoutine, J, WARNING_MSG & "assigned only variable" & strInSQuotes(varPosVar, True) & "removed.", MAfter)
                 Else
                  ArrRoutine(J) = SmartMarker(ArrRoutine, J, WARNING_MSG & "assigned only variable" & strInSQuotes(varPosVar, True) & "converted to 'Call'.", MAfter)
                End If
                Exit For
              End If
            End If
          End If
        Next varPosVar
       Else
        ArrRoutine(J) = SQuote & ArrRoutine(J)
        ArrRoutine(J) = SmartMarker(ArrRoutine, J, WARNING_MSG & "assigned only variable commented out", MAfter)
      End If
    End If
  Next J
  ArrRoutine(TriggerLine) = SQuote & ArrRoutine(TriggerLine)

End Sub

Private Sub UnUsedDimCommentOut(ArrRoutine As Variant, _
                                arrTargets As Variant, _
                                ByVal TriggerLine As Long)

  Dim J As Long

  For J = UBound(ArrRoutine) - 1 To LBound(ArrRoutine) + 1 Step -1
    If IsInArray(J, arrTargets) Then
      ArrRoutine(J) = SQuote & ArrRoutine(J)
      ArrRoutine(J) = SmartMarker(ArrRoutine, J, WARNING_MSG & "assigned only variable commented out", MAfter)
    End If
  Next J
  ArrRoutine(TriggerLine) = SQuote & ArrRoutine(TriggerLine)

End Sub

Private Function UsedAnywhereLowerIn(ArrProc As Variant, _
                                     ByVal TestPos As Long, _
                                     arrDim As Variant) As Boolean

  Dim I As Long
  Dim J As Long

  For I = TestPos + 1 To UBound(ArrProc)
    For J = LBound(arrDim) To UBound(arrDim)
      If InstrAtPosition(ArrProc(I), arrDim(J), ipAny) Then
        UsedAnywhereLowerIn = True
        Exit For
      End If
    Next J
    If UsedAnywhereLowerIn Then
      Exit For
    End If
  Next I

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:16:04 AM) 5 + 2180 = 2185 Lines Thanks Ulli for inspiration and lots of code.

