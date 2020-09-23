Attribute VB_Name = "mod_UsageComment"
Option Explicit

Public Function ExtractName(ByVal strCode As String) As String

  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTest As Variant

  arrTest = Array("Private", "Public", "Friend", "Type", "Static", "Enum", "Declare", "Let", "Set", "Get", "Sub", "Function", _
                  "Property", "Event", "WithEvents", "Const")
  strCode = ExpandForDetection(strCode)
  arrTmp = Split(strCode)
  For I = LBound(arrTmp) To UBound(arrTmp)
    'v2.3.1 a mis-type damaged the Declare element of the following line
    If Not IsInArray(arrTmp(I), arrTest) Then
      ExtractName = arrTmp(I)
      Exit For
    End If 'Select
  Next I

End Function

Private Function ExtractVariableType(ByVal strCodeOnly As String, _
                                     ByVal strWord As String, _
                                     strOrigCode As String) As String

  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTest As Variant

  arrTest = Array("Private", "Public", "Friend", "Static", "Let", "Set", "Get")
  arrTmp = Split(strCodeOnly, strWord)
  arrTmp = Split(arrTmp(0))
  For I = LBound(arrTmp) To UBound(arrTmp)
    If Not IsInArray(arrTmp(I), arrTest) Then
      ExtractVariableType = arrTmp(I)
      Exit For
    End If 'Select
  Next I
  If Len(ExtractVariableType) = 0 Then
    If LeftWord(strOrigCode) = strWord Then
      ExtractVariableType = "Enum Member"
     Else
      ExtractVariableType = "Variable"
    End If
  End If

End Function

Private Function GenerateUsageMsg(ByVal strHomeMod As String, _
                                  ByVal strComType As String, _
                                  ByVal strWord As String, _
                                  ByVal strInput As String, _
                                  ByVal strProc As String, _
                                  ByVal bNoRefInHomeMod As Boolean, _
                                  ByVal ModCount As Long, _
                                  TotalCount As Long, _
                                  ByVal StrHeadType As String, _
                                  ByVal StrScope As String, _
                                  ByVal bIsReset As Boolean, _
                                  ByVal bIsFormModuleInterface As Boolean, _
                                  ByVal bMultiForm As Boolean, _
                                  bSkipDeadEnum As Boolean) As String

  
  Dim bSuggestedIsForm      As Boolean
  Dim I                     As Long
  Dim strHomeValue          As Long
  Dim strSuggestDim         As String
  Dim strLocations          As String
  Dim strTmp                As String
  Dim strFirstSuggestion    As String
  Dim strControlEvent       As String
  Dim strStarupComment      As String
  Const StrSafeControlEvent As String = UsageSign & "Control Event."
  Dim arrQUse               As Variant
  Dim ArrProc               As Variant
  Dim arrTest               As Variant

  arrTest = Array("Declare", "Function", "Sub", "Property", "Const")
  On Error GoTo BugTrap
  bSkipDeadEnum = False
  ArrProc = Split(strProc, "|")
  If StrHeadType = "Declare" Then
    StrHeadType = " API Declare"
  End If
  'Control Procedure
  '0 OK if Control Exists
  '  Called from code advise seperate proc
  '  Large Procedure  advise seperate proc
  '
  'Form Procedure
  ' 0 delete
  '  Large Procedure  advise seperate proc/module
  '
  'Form/Form Procedure
  ' Mult Forms call must be in a module
  '
  'Form/Form [Variable]
  ' move to module easier than form references
  '
  'Form/Module[Procedure|Variable]
  '
  'Module/Module[Procedure|Variable]
  ' If strWord = "Form_Load" Then
  If StartUpComponent = strHomeMod Then
    If "Sub " & strWord = StartUpProcedure Then
      strStarupComment = UsageSign & "Project StartUp Object"
    End If
  End If
  For I = LBound(ArrProc) To UBound(ArrProc)
    If ContainsWholeWord(ArrProc(I), strHomeMod) Then
      ArrProc(I) = ArrProc(I) & " *"
    End If
  Next I
  If UBound(ArrProc) > -1 Then
    strLocations = UsageSign & vbTab & " (Count) " & IIf(StrScope = "Private", vbNullString, "Moudule-") & "Procedure(s)" & IIf(StrScope <> "Private", " [*] Indicates Home Module", vbNullString) & UsageSign & vbTab & Join(ArrProc, UsageSign & vbTab)
  End If
  strHomeValue = "0"
  If Len(strInput) Then
    arrQUse = QuickSortArray(Split(strInput, ","), False)
    For I = LBound(arrQUse) To UBound(arrQUse)
      arrQUse(I) = LStrip(CStr(arrQUse(I)), "0")
      If Left$(arrQUse(I), 3) = "  a" Or Left$(arrQUse(I), 3) = "  Z" Then
        arrQUse(I) = "0 " & arrQUse(I)
      End If
      arrQUse(I) = String$(5 - Len(Trim$(Left$(arrQUse(I), InStr(arrQUse(I), SngSpace)))), 32) & arrQUse(I)
      If InStr(arrQUse(I), " a ") Then
        arrQUse(I) = Replace$(arrQUse(I), " a ", "<") & ">"
        strHomeValue = WordInString(Trim$(arrQUse(I)), 1)
       Else
        arrQUse(I) = Replace$(arrQUse(I), " Z ", vbNullString)
      End If
    Next I
    strFirstSuggestion = RStrip(LStrip(WordInString(arrQUse(0), 2), "<"), ">")
    bSuggestedIsForm = IsControlHolderFromName(strFirstSuggestion)
    If IsVBControlRoutine(strWord) Then
      If TotalCount Then
        strControlEvent = UsageSign & "Control Event accessed from code." & UsageSign & " Suggestion: Makes code difficult to read." & UsageSign & "Create a Sub from Event code and call that."
       Else
        strControlEvent = StrSafeControlEvent
      End If
      If RoutineNameIsUserInteractive(strWord, GetComponent("", strHomeMod)) Then
        strControlEvent = strControlEvent & UsageSign & "User-Object Interaction"
       ElseIf RoutineNameIsInternalToModule(strWord, GetComponent("", strHomeMod)) Then
        strControlEvent = strControlEvent & UsageSign & "Internal Object " & strComType
      End If
    End If
    If bNoRefInHomeMod Then
      If ModCount = 1 Then
        If bSuggestedIsForm Then
          strTmp = UsageSign & "Code Entry point from " & strFirstSuggestion
         Else
          If StrScope = "Public" Then
            If bIsFormModuleInterface Then
              strTmp = UsageSign & "Form/Module " & strComType & "(No action required)." & UsageSign & "SUGGESTION: Scope to Private by moving to form if program is very simple."
             Else
              If bSuggestedIsForm Then
                strTmp = UsageSign & "Form Only " & strComType & UsageSign & "RECOMMENDED: Scope to Private by moving to form below."
               Else
                If bNoRefInHomeMod Then
                  strTmp = UsageSign & "Form/Module Interface (No action required)." & UsageSign & "SUGGESTION: Scope to Private by moving to form if program is very simple."
                 Else
                  strTmp = UsageSign & "Form/Module " & strComType & " NOT USED IN THIS MODULE" & UsageSign & "RECOMMENDED: Scope to Private by moving to module below."
                End If
              End If
            End If
          End If
        End If
       Else
        If bSuggestedIsForm Then
          strTmp = UsageSign & "Code Entry point from " & strFirstSuggestion & UsageSign & "NOT USED IN THIS MODULE" & UsageSign & "Could be moved to following module(s):"
         Else
          If StrScope = "Public" Then
            If bMultiForm Then
              strTmp = UsageSign & "Multi-Module " & strComType & UsageSign & "Must be in a bas module "
             Else
              strTmp = UsageSign & "NOT USED IN THIS MODULE" & UsageSign & "Could be moved to following module(s):"
            End If
          End If
        End If
      End If
    End If
    If StrScope = "Public" Then
      GenerateUsageMsg = UsageSign & "Usage " & strComType & " Local/Total: (" & strHomeValue & "/" & TotalCount & ")" & strInSQuotes(strWord, True) & strControlEvent & strStarupComment & strTmp & strLocations
     Else
      If UBound(ArrProc) > 0 Then
        GenerateUsageMsg = UsageSign & "Usage " & strComType & " Total: (" & TotalCount & ")" & strInSQuotes(strWord, True) & strControlEvent & strStarupComment & strTmp & strLocations
       Else
        If IsInArray(strComType, arrTest) Or ((TotalCount = 1) And InStr(strLocations, "(Declarations)")) Then
          strSuggestDim = vbNullString
         Else
          If bIsReset Then
            strSuggestDim = UsageSign & "RECOMMENDED: Could be Dim in following procedure"
          End If
        End If
        If strControlEvent = StrSafeControlEvent Then
          'Control event not called by code
          GenerateUsageMsg = UsageSign & "Usage " & strComType & strInSQuotes(strWord, True) & strControlEvent & strStarupComment
         Else
          GenerateUsageMsg = UsageSign & "Usage " & strComType & " Total: (" & TotalCount & ")" & strInSQuotes(strWord) & strControlEvent & strStarupComment & strTmp & strSuggestDim & strLocations
        End If
      End If
    End If
   Else
    'safety should never hit
    If Not bNoRefInHomeMod Then
      If ModCount = 0 Then
        GenerateUsageMsg = UsageSign & "Usage " & strComType & " NOT USED" & strInSQuotes(strWord, True)
      End If
     Else
      If strWord = "Main" Then
        GenerateUsageMsg = UsageSign & "Usage " & strComType & strInSQuotes(strWord, True) & UsageSign & IIf(StartUpProcedure = "Sub Main", "Project StartUp Object", "NOT USED IN THIS PROGRAM: StartUp Object set to " & StartUpComponent)
       Else
        If IsControlEvent(strWord) Then
          GenerateUsageMsg = UsageSign & "Usage " & strComType & strInSQuotes(strWord, True) & IIf(LenB(strStarupComment), strStarupComment, UsageSign & "Control Event accessed only from Object or User Interface")
         Else
          If Not IsVBControlRoutine(strWord) Then
            GenerateUsageMsg = UsageSign & "Usage " & strComType & " NOT USED." & strInSQuotes(strWord, True)
            If strComType = "Const" Then
              GenerateUsageMsg = GenerateUsageMsg & "(Probably used to create a composite Const which was later marked unused)"
            End If
            If strComType = "Enum" Then
              If IsArray(DeclarDesc(GetDeclarationID(strWord, StrScope, "Enum", strHomeMod)).DDUsage) Then
                GenerateUsageMsg = UsageSign & "Usage " & strComType & " Not used directly but member(s) are used." & strInSQuotes(strWord, True)
               Else
                bSkipDeadEnum = True
              End If
            End If
          End If
        End If
      End If
    End If
  End If
  Do While SmartRight(GenerateUsageMsg, UsageSign & vbTab)
    GenerateUsageMsg = Left$(GenerateUsageMsg, Len(GenerateUsageMsg) - Len(UsageSign & vbTab))
  Loop

Exit Function

BugTrap:
  BugTrapComment "GenerateUsageMsg"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Function

Public Function GetDeclarationCount(ByVal strCodeOnly As String, _
                                    ByVal OrigLin As String, _
                                    ByVal StrScope As String, _
                                    ByVal CName As String, _
                                    Optional bSkipDeadEnum As Boolean = False) As String

  
  Dim Proj                   As VBProject
  Dim Comp                   As VBComponent
  Dim StartLine              As Long
  Dim CompFindCount          As Long
  Dim TotalCount             As Long
  Dim strHomeMod             As String
  Dim L_CodeLine             As String
  Dim bNoRefInHomeMod        As Boolean
  Dim bIsFormModuleInterface As Boolean
  Dim FormCount              As Long
  Dim bMultiForm             As Boolean
  Dim ModCount               As Long
  Dim strTmp                 As String
  Dim strProcTmp             As String
  Dim strProc                As String
  Dim strComType             As String
  Dim GuardLine              As Long
  Dim strWord                As String
  Dim lngdummy               As Long
  Dim IsReset                As Boolean
  Dim strOldProc             As String
  On Error GoTo BugTrap
  strWord = ExtractName(strCodeOnly)
  If Not isCodeFixProtected(strWord) Then
    strComType = ExtractVariableType(strCodeOnly, strWord, strCodeOnly)
    On Error Resume Next
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, lngdummy, False) Then
          If StrScope = "Private" Then
            If CName <> Comp.Name Then
              GoTo SkipComp
            End If
          End If
          StartLine = 1
          GuardLine = 0
          With Comp
            Do While .CodeModule.Find(strWord, StartLine, 1, -1, -1, True, True)
              L_CodeLine = .CodeModule.Lines(StartLine, 1)
              'Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strWord, L_CodeLine, StartLine)
              If GuardLine > 0 Then
                If GuardLine > StartLine Then
                  Exit Do
                End If
              End If
              If OrigLin = L_CodeLine Then
                strHomeMod = .Name
               Else
                L_CodeLine = Trim$(L_CodeLine)     '.CodeModule.Lines(StartLine, 1))
                If InCode(L_CodeLine, InStr(L_CodeLine, strWord)) Then
                  If GetProcName(.CodeModule, StartLine) <> strOldProc Then
                    strOldProc = GetProcName(.CodeModule, StartLine)
                    'if first occurance in any proc is assignment, set or For StartLine then
                    'the variable need not be module level
                    IsReset = MultiLeft(Trim$(L_CodeLine), True, strWord & EqualInCode, "Set " & strWord & EqualInCode, "For " & strWord & EqualInCode)
                    If IsReset Then
                      'if variable is first used as  var = var + StartLine then it needs to be at module level
                      IsReset = Not MultiLeft(Trim$(L_CodeLine), True, , strWord & EqualInCode & strWord)
                    End If
                  End If
                  'Enum Cap block for Enums
                  If strComType = "Enum Member" Then
                    If strOldProc = "(Declarations)" Then
                      If ((L_CodeLine Like "Private *" & strWord & "*") And InStr(L_CodeLine, " = ") = 0) Or L_CodeLine Like "*," & strWord & "*" Then
                        GoTo Skip
                      End If
                    End If
                  End If
                  If OrigLin <> L_CodeLine Then
                    CompFindCount = CompFindCount + 1
                    TotalCount = TotalCount + 1
                    strProcTmp = AccumulatorString(strProcTmp, .Name & "-" & GetProcName(.CodeModule, StartLine), ",", False)
                  End If
Skip:
                End If
              End If
              StartLine = StartLine + 1
              GuardLine = StartLine
              If StartLine > .CodeModule.CountOfLines Then
                'assumes you never want the last line 'End something almost certainly
                Exit Do
              End If
            Loop
          End With 'Comp
          If CompFindCount Then
            If Not bMultiForm Then
              If IsComponent_ControlHolder(Comp) Then
                FormCount = FormCount + 1
                bMultiForm = FormCount > 1
              End If
            End If
            strTmp = AccumulatorString(strTmp, Format$(CompFindCount, "00000000") & SngSpace & IIf(strHomeMod = Comp.Name, " a ", " Z ") & Comp.Name)
            strProc = AccumulatorString(strProc, UsageData(strProcTmp, StrScope, strHomeMod), "|", False)
            ModCount = ModCount + 1
            CompFindCount = 0
           Else
            If Comp.Name = strHomeMod Then
              bNoRefInHomeMod = True
             Else
              If IsComponent_ControlHolder(Comp) And strHomeMod <> Comp.Name Then
                bIsFormModuleInterface = True
              End If
            End If
          End If
SkipComp:
        End If
      Next Comp
    Next Proj
    GetDeclarationCount = GenerateUsageMsg(strHomeMod, strComType, strWord, strTmp, strProc, bNoRefInHomeMod, ModCount, TotalCount, "Declaration", StrScope, IsReset, bIsFormModuleInterface, bMultiForm, bSkipDeadEnum)
  End If
  On Error GoTo 0

Exit Function

BugTrap:
  BugTrapComment "GetDeclarationCount"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Function

Private Function GetProcedureCount(ByVal OrigLin As String, _
                                   ByVal StrScope As String, _
                                   ByVal CName As String) As String

  Dim strTmp        As String
  Dim Proj          As VBProject
  Dim Comp          As VBComponent
  Dim StartLine     As Long
  Dim CompFindCount As Long
  Dim TotalCount    As Long
  Dim strHomeMod    As String
  Dim L_CodeLine    As String
  Dim FormCount     As Long
  Dim bMultiForm    As Boolean
  Dim ModCount      As Long
  Dim strProc       As String
  Dim strProcTmp    As String
  Dim strComType    As String
  Dim GuardLine     As Long
  Dim strWord       As String
  Dim lngdummy      As Long

  On Error GoTo BugTrap
  strWord = GetRoutineName(OrigLin)
  strComType = ExtractVariableType(OrigLin, strWord, OrigLin)
  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, lngdummy, False) Then
        If StrScope = "Private" Then
          If CName <> Comp.Name Then
            GoTo SkipComp
          End If
        End If
        StartLine = Comp.CodeModule.CountOfDeclarationLines + 1
        GuardLine = 0
        With Comp
          Do While .CodeModule.Find(strWord, StartLine, 1, -1, -1, True, True)
            L_CodeLine = .CodeModule.Lines(StartLine, 1)
            '          Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strWord, L_CodeLine, StartLine)
            If GuardLine > 0 Then
              If GuardLine > StartLine Then
                Exit Do
              End If
            End If
            If OrigLin = L_CodeLine Then
              strHomeMod = .Name
              'CompFindCount = CompFindCount + 1
              'TotalCount = TotalCount + 1
              'strProcTmp = AccumulatorString(strProcTmp, Comp.Name & "-" & GetProcName(.CodeModule, StartLine), ",", False)
             Else
              L_CodeLine = Trim$(L_CodeLine)
              If InCode(L_CodeLine, InStr(L_CodeLine, strWord)) Then
                If Not InProcedure(strWord, Comp, StartLine) Then
                  CompFindCount = CompFindCount + 1
                  TotalCount = TotalCount + 1
                  strProcTmp = AccumulatorString(strProcTmp, .Name & "-" & GetProcName(.CodeModule, StartLine), ",", False)
                End If
              End If
            End If
            StartLine = StartLine + 1
            GuardLine = StartLine
          Loop
        End With 'Comp
        If CompFindCount Then
          If Not bMultiForm Then
            If IsComponent_ControlHolder(Comp) Then
              FormCount = FormCount + 1
              bMultiForm = FormCount > 1
            End If
          End If
          strTmp = AccumulatorString(strTmp, Format$(CompFindCount, "00000000") & SngSpace & IIf(strHomeMod = Comp.Name, " a ", " Z ") & Comp.Name, ",", False)
          ' " a ","Z" forces the sort rouitne to place this at top of its number range
          strProc = AccumulatorString(strProc, UsageData(strProcTmp, StrScope, strHomeMod), "|", False)
          ModCount = ModCount + 1
          CompFindCount = 0
         Else
          If Comp.Name = strHomeMod Then
            strTmp = AccumulatorString(strTmp, Format$(CompFindCount, "00000000") & SngSpace & IIf(strHomeMod = Comp.Name, " a ", " Z ") & Comp.Name, ",", False)
            ' " a ","Z" forces the sort rouitne to place this at top of its number range
            strProc = AccumulatorString(strProc, UsageData(strProcTmp, StrScope, strHomeMod), "|", False)
           Else
            If IsComponent_ControlHolder(Comp) And strHomeMod <> Comp.Name Then
              strTmp = AccumulatorString(strTmp, Format$(CompFindCount, "00000000") & SngSpace & IIf(strHomeMod = Comp.Name, " a ", " Z ") & Comp.Name, ",", False)
              ' " a ","Z" forces the sort rouitne to place this at top of its number range
              strProc = AccumulatorString(strProc, UsageData(strProcTmp, StrScope, strHomeMod), "|", False)
            End If
          End If
        End If
SkipComp:
      End If
    Next Comp
  Next Proj
  GetProcedureCount = GenerateUsageMsg(strHomeMod, strComType, strWord, strTmp, strProc, True, ModCount, TotalCount, "Procedure", StrScope, False, True, bMultiForm, False)
  On Error GoTo 0

Exit Function

BugTrap:
  BugTrapComment "GetProcedureCount"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Function

Public Function GetProcName(cMod As CodeModule, _
                            ByVal Sline As Long) As String

  On Error Resume Next
  GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Proc)
  If LenB(GetProcName) = 0 Then
    GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Let)
  End If
  If LenB(GetProcName) = 0 Then
    GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Get)
  End If
  If LenB(GetProcName) = 0 Then
    GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Set)
  End If
  If LenB(GetProcName) = 0 Then
    'dummy for detecting that item is in Declaration section
    GetProcName = "(Declarations)"
  End If
  On Error GoTo 0

End Function

Public Function GetSafeInsertLine(cMod As CodeModule, _
                                  ByVal StartLine As Long) As Long

  Do While HasLineCont(cMod.Lines(StartLine, 1))
    StartLine = StartLine + 1
  Loop
  'v 2.0.8 doesn't need the last line test of GetSafeInsertLineArray bcause it can create new member
  GetSafeInsertLine = StartLine + 1

End Function

Public Function GetSafeInsertLineArray(arrR As Variant, _
                                       ByVal StartLine As Long) As Long

  If StartLine < 0 Then
    StartLine = 0
    'GetSafeInsertLineArray = 0
    'Exit Function
  End If
  Do While HasLineCont(arrR(StartLine))
    StartLine = StartLine + 1
  Loop
  'v 2.0.8 added trap for lastline of array
  GetSafeInsertLineArray = StartLine '+ IIf(StartLine < UBound(arrR), 1, 0)

End Function

Private Function IsControlHolderFromName(ByVal strTest As String) As Boolean

  Dim Proj As VBProject
  Dim Comp As VBComponent

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If Comp.Name = strTest Then
          IsControlHolderFromName = IsComponent_ControlHolder(Comp)
          Exit For
        End If
      End If
    Next Comp
    If IsControlHolderFromName Then
      Exit For
    End If
  Next Proj

End Function

Public Function IsECPCode(strCode As String, _
                          ByVal CodeLineNo As Long, _
                          CompMod As CodeModule) As Boolean

  Dim LineNo As Long

  If Left$(strCode, 8) = "Private " Then
    If InStr(strCode, ",") Then
      If Not Has_AS(strCode) Then
        LineNo = CodeLineNo
        Do While Left$(CompMod.Lines(LineNo, 1), 8) = "Private "
          LineNo = LineNo - 1
          If LineNo = 0 Then
            Exit Do 'Function
          End If
        Loop
        If LineNo Then
          If Left$(CompMod.Lines(LineNo, 1), 36) = "#If False Then 'Trick preserves Case" Then
            IsECPCode = True
           ElseIf Left$(CompMod.Lines(LineNo - 1, 1), 8) = "End Enum" Then
            IsECPCode = True
          End If
        End If
      End If
    End If
  End If

End Function

Public Function IsVBControlRoutine(ByVal strTest As String) As Boolean

  Dim strR As String

  'NEW ver 1.1.40 simplifies the tests
  'v 2.1.2 simplified further
  'v 2.3.1 added extractor which copes with multi underscores in name
  If CountSubString(strTest, "_") Then
    strR = StrReverse(strGetRightOf(StrReverse(strTest), "_"))
    If IsControlEvent(strR) Then
      IsVBControlRoutine = True
     Else
      If CntrlDescMember(strR) > -1 Then
        IsVBControlRoutine = True
       Else
        If Len(strR) Then
          IsVBControlRoutine = isEvent(strR)
        End If
        If IsControlEvent(strR) Then
          IsVBControlRoutine = True
         Else
          If Len(strR) Then
            IsVBControlRoutine = isEvent(strR)
          End If
        End If
      End If
    End If
  End If

End Function

Private Sub QuicksortA(A As Variant, _
                       ByVal Lo As Long, _
                       ByVal Hi As Long, _
                       Optional Ascending As Boolean = True)

  Dim HiIndex    As Long
  Dim LoIndex    As Long
  Dim CurElement As Variant

  HiIndex = Lo
  LoIndex = Hi
  CurElement = A((Lo + Hi) / 2)
  Do While (HiIndex <= LoIndex)
    Do While IIf(Ascending, (A(HiIndex) < CurElement), (A(HiIndex) > CurElement)) And HiIndex < Hi
      HiIndex = HiIndex + 1
    Loop
    Do While IIf(Ascending, (CurElement < A(LoIndex)), (CurElement > A(LoIndex))) And LoIndex > Lo
      LoIndex = LoIndex - 1
    Loop
    If HiIndex <= LoIndex Then
      SwapVariants A(HiIndex), A(LoIndex)
      HiIndex = HiIndex + 1
      LoIndex = LoIndex - 1
    End If
  Loop
  If Lo < LoIndex Then
    QuicksortA A, Lo, LoIndex, Ascending
  End If
  If HiIndex < Hi Then
    QuicksortA A, HiIndex, Hi, Ascending
  End If

End Sub

Public Sub SafeInsertModule(cMod As CodeModule, _
                            ByVal lngLineNo As Long, _
                            ByVal StrInsert As String)

  'ver 1.184
  'Scan until you find end of line continuation codeline
  'or just the next line and insert the string

  If Len(StrInsert) Then
    cMod.InsertLines GetSafeInsertLine(cMod, lngLineNo), StrInsert
  End If

End Sub

Public Sub SwapVariants(Var1 As Variant, _
                        Var2 As Variant)

  Dim Var3 As Variant

  Var3 = Var1
  Var1 = Var2
  Var2 = Var3

End Sub

Public Sub UsageCountProcedures(ByVal StrScope As String)

  Dim strComment   As String
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim L_CodeLine   As String
  Dim CurCompCount As Long
  Dim StartLine    As Long
  Dim GuardLine    As Long

  On Error GoTo BugTrap
  If Not bAborting Then
    'Copyright 2004 Roger Gilchrist
    'e-mail: rojagilkrist@hotmail.com
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount, False) Then
          If Not IsComponent_ClassMod(Comp.Type) Then
            ModuleMessage Comp, CurCompCount
            With Comp.CodeModule
              StartLine = .CountOfDeclarationLines + 1
              'if exits at all, then look for the line(s)
              Do While .Find(StrScope, StartLine, 1, -1, -1, True, True)
                L_CodeLine = .Lines(StartLine, 1)
                'Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, StrScope, L_CodeLine, StartLine)
                If GuardLine > 0 Then
                  If GuardLine > StartLine Then
                    Exit Do
                  End If
                End If
                MemberMessage "", StartLine, .CountOfLines
                If ExtractCode(L_CodeLine) Then
                  If InCode(L_CodeLine, InStr(L_CodeLine, StrScope & SngSpace)) Then
                    '                  If LenB(ExtractName(L_CodeLine)) Then
                    '                    z = ProcDescMember(ExtractName(L_CodeLine), Comp.Name)
                    '                    If z > -1 Then
                    '                      SafeInsertModule Comp.CodeModule, StartLine, PRocDesc(z).PrDComment
                    '                    End If
                    '
                    '
                    '                  End If
                    If LenB(ExtractName(L_CodeLine)) Then
                      strComment = GetProcedureCount(L_CodeLine, StrScope, Comp.Name)
                      If Len(strComment) Then
                        SafeInsertModule Comp.CodeModule, StartLine, strComment
                        'xPRocDesc(z).PrDComment
                      End If
                    End If
                  End If
                End If
                StartLine = StartLine + 1
                GuardLine = StartLine
                If StartLine >= .CountOfLines Then
                  Exit Do
                End If
              Loop
            End With
          End If
        End If
      Next Comp
    Next Proj
  End If

Exit Sub

BugTrap:
  BugTrapComment "UsageCountProcedures"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Function UsageData(strProcTmp As String, _
                           ByVal StrScope As String, _
                           ByVal strHomeMod As String) As String

  Dim strVal    As String
  Dim ArrProc   As Variant
  Dim curProc   As String
  Dim ProcCount As Long
  Dim strProc   As String
  Dim I         As Long

  ArrProc = Split(strProcTmp, ",")
  strProcTmp = vbNullString
  curProc = ArrProc(0)
  For I = LBound(ArrProc) To UBound(ArrProc)
    If curProc <> ArrProc(I) Then
      strProc = AccumulatorString(strProc, Format$(ProcCount, "00000000") & IIf(MultiLeft(ArrProc(I), True, strHomeMod), " a ", " Z ") & curProc)
      ProcCount = 1
      curProc = ArrProc(I)
     Else
      ProcCount = ProcCount + 1
    End If
  Next I
  'get the last one
  ' depends on the little known (and never used 'cause its so dubious) fact that a FOr value is actual 1 above last val on exiting the structure
  UsageData = AccumulatorString(strProc, Format$(ProcCount, "00000000") & IIf(MultiLeft(ArrProc(I - 1), True, strHomeMod), " a ", " Z ") & curProc)
  ' " a "/" Z " used to force the home module to top of its count range (rest of count range is alphabetical)
  ArrProc = QuickSortArray(Split(UsageData, ","), False)
  For I = LBound(ArrProc) To UBound(ArrProc)
    ArrProc(I) = Replace$(ArrProc(I), " Z ", SngSpace) ' strip a/Z range setters
    ArrProc(I) = Replace$(ArrProc(I), " a ", SngSpace)
    ArrProc(I) = LStrip(CStr(ArrProc(I)), "0")
    ' strip the excess 0's that make sure list sorts properly
    strVal = Trim$(Left$(ArrProc(I), InStr(ArrProc(I), SngSpace)))
    'take value off temporarily
    ArrProc(I) = Mid$(ArrProc(I), InStr(ArrProc(I), SngSpace) + 1)
    'get the naming piece
    If StrScope = "Private" Then
      ArrProc(I) = Mid$(ArrProc(I), InStr(ArrProc(I), " - ") + 1)
     Else
      If LenB(strHomeMod) Then
        If SmartLeft(ArrProc(I), strHomeMod) Then   'add brackets to indicate home module
          ArrProc(I) = "<" & ArrProc(I)
          ArrProc(I) = Replace$(ArrProc(I), "-", ">-")
        End If
      End If
    End If
    ArrProc(I) = String$(5 - Len(strVal), 32) & strInBrackets(strVal) & SngSpace & ArrProc(I)
    ' format the line
  Next I
  UsageData = Join(ArrProc, "|")

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:24:24 AM) 1 + 780 = 781 Lines Thanks Ulli for inspiration and lots of code.

