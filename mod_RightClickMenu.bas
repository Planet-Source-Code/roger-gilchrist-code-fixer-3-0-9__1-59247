Attribute VB_Name = "mod_RightClickMenu"
Option Explicit
Private Type FormatUndo
  FUHasData                     As Boolean
  FUProcName                    As String
  FUmodule                      As CodeModule
  FUStart                       As Long
  FUEnd                         As Long
  FUArray                       As Variant
End Type
Private UnDoProcFormatData      As FormatUndo
'making them public allows enabling/disabling using their names rather than riskier
''VBInstance.CommandBars(15).Controls.Item(16).Controls.Item(3).Enabled = True which would be
'thrown off by any other menu manipulations
'Private MyRightUnDoIndent       As Office.CommandBarButton
''Private MyRightFormMenu         As CommandBarButton
Public MyRightControlsMenu      As CommandBarButton
Public MyRightCodeMenu          As CommandBarButton

Private Sub ArrayLineContCollapse(arrData As Variant)

  Dim I As Long
  Dim J As Long

  'v2.2.6
  'strip out the line continuations
  'used by DecIndent, ProcIndent, ProcNoIndent
  For I = LBound(arrData) To UBound(arrData)
    If HasLineCont(arrData(I)) Then
      J = I
      Do While HasLineCont(arrData(I))
        J = J + 1
        arrData(I) = Replace$(arrData(I), ContMark, " ") & arrData(J)
        arrData(J) = vbNullString
      Loop
    End If
  Next I

End Sub

Private Sub DecIndent()

  Dim StrTypeName  As String
  Dim StrEnumName  As String
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant
  Dim CPane        As CodePane
  Dim cMod         As CodeModule
  Dim bDummy       As Boolean
  Dim I            As Long
  Dim CurPos       As Long
  Dim MyhourGlass  As cls_HourGlass

  'v2.2.0
  Set MyhourGlass = New cls_HourGlass
  Set CPane = GetActiveCodePane
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind, CurPos)
  For I = TopOfRoutine + 1 To UBound(ArrProc)
    ArrProc(I) = Trim$(ArrProc(I))
    If SmartLeft(ArrProc(I), WARNING_MSG & "Incomplete ") Then
      ArrProc(I) = vbNullString
    End If
  Next I
  ArrayLineContCollapse ArrProc
  SeparateCompoundDeclarationLines -1, ArrProc, cMod.Parent.Name
  Assign_Type_To_Constants -1, ArrProc
  Create_Enum_Capitalisation_Protection -1, ArrProc
  DeclarationOffSetAsTypeEOL -1, ArrProc
  LineContinuationFix ArrProc, bDummy, -1, True
  ProtectEnumCap ArrProc
  For I = LBound(ArrProc) To UBound(ArrProc)
    If InTypeDef(ArrProc(I), StrTypeName) Then
      'v 2.2.1 tighter test 'don't indent start of type
      If Not InstrAtPosition(ArrProc(I), "Type", ipAny) Then
        ArrProc(I) = Space$(1 * IndentSize) & ArrProc(I)
      End If
    End If
    If InEnumDef(ArrProc(I), StrEnumName) Then
      'v 2.2.1 tighter test'don't indent start of enum
      If Not InstrAtPosition(ArrProc(I), "Enum", ipAny) Then
        ArrProc(I) = Space$(1 * IndentSize) & ArrProc(I)
      End If
    End If
    If InConditionalCompile(ArrProc(I)) Then
      'v 2.2.1 ' added cuteness
      If Not Left$(ArrProc(I), 4) = "#If " Then
        If Not Left$(ArrProc(I), 6) = "#Else " Then
          ArrProc(I) = Space$(1 * IndentSize) & ArrProc(I)
        End If
      End If
    End If
  Next I
  ReplaceProcedureCode cMod, ArrProc, 1, PEndLine

End Sub

Private Sub DoMatchTrigger(strData As String, _
                           strTest As String, _
                           bResult As Boolean)

  Dim RLen As Long

  'v 2.2.6 Support routine for MatchTrigger
  If SmartRight(strData, strTest) Then
    bResult = True
    RLen = Len(strTest)
    strData = Left$(strData, Len(strData) - IIf(InStr(strData, ","), RLen + 1, RLen))
  End If

End Sub

Public Sub DoRightCase(ByVal LngMode As Long)

  Select Case LngMode
   Case 0
    ConvertSelectedText myeLowerCase
   Case 1
    ConvertSelectedText myeUpperCase
   Case 2
    ConvertSelectedText myeProperCase
   Case 3
    ConvertSelectedText myeSimpleSentence
  End Select

End Sub

Public Sub DoRightMenuFind(ByVal Level As Long)

  If Len(GetSelectedText) Then
    Select Case Level
     Case 0
      iRange = AllCode
     Case 1
      iRange = ModCode
     Case 2
      iRange = ProcCode
    End Select
    mObjDoc.InitiateSearch GetSelectedText()
  End If

End Sub

Private Function GetProcPositionFromName(cMod As CodeModule, _
                                         ByVal ProcName As String, _
                                         SPos As Long, _
                                         EPos As Long) As Boolean

  Dim Sline  As Long
  Dim lDummy As Long

  'v2.2.0 support routine for UnDoProcFormat
  Do While cMod.Find(ProcName, Sline, lDummy, lDummy, lDummy)
    If isProcHead(cMod.Lines(Sline, 1)) Then
      SPos = Sline
      EPos = GetProcEndLine(cMod, Sline)
      GetProcPositionFromName = True
      Exit Do
    End If
  Loop

End Function

Private Function MatchTrigger(strData As String, _
                              ByVal strTrigger As String, _
                              strCode As String, _
                              Optional bExccessEndError As Boolean = False) As Boolean

  Dim ErrStruct As Variant

  'v2.2.0
  Select Case strTrigger
   Case "Case"
    DoMatchTrigger strData, "Case", MatchTrigger
   Case "Else"
    DoMatchTrigger strData, "Else", MatchTrigger
   Case "ElseIf"
    DoMatchTrigger strData, "ElseIf", MatchTrigger
    If Not MatchTrigger Then
      DoMatchTrigger strData, "Else", MatchTrigger
    End If
   Case "Loop"
    DoMatchTrigger strData, "Do", MatchTrigger
   Case "Next"
    DoMatchTrigger strData, "For", MatchTrigger
   Case "Close"
    DoMatchTrigger strData, "Open", MatchTrigger
   Case "Wend"
    DoMatchTrigger strData, "While", MatchTrigger
   Case "End"
    ErrStruct = Split(strData, ",")
    'v2.7.3 stops crash in single proc formatting
    If UBound(ErrStruct) > -1 Then
      If SmartLeft(strCode, "End " & ErrStruct(UBound(ErrStruct))) Then
        MatchTrigger = True
        strData = Left$(strData, Len(strData) - (Len(ErrStruct(UBound(ErrStruct))) + IIf(InStr(strData, ","), 1, 0)))
      End If
     Else
      If Left$(strCode, 6) = "End If" Then
        MatchTrigger = True
        bExccessEndError = True
      End If
    End If
  End Select

End Function

Private Sub ProcDeleteCFComments(ArrProc As Variant)

  Dim I    As Long
  Dim cpos As Long

  For I = LBound(ArrProc) To UBound(ArrProc)
    cpos = InStr(ArrProc(I), RGSignature)
    If cpos Then
      ArrProc(I) = Left$(ArrProc(I), cpos - 1)
    End If
  Next I

End Sub

Private Sub ProcIndent()

  
  Dim strF                  As String
  Dim PSLine                As Long
  Dim PEndLine              As Long
  Dim TopOfRoutine          As Long
  Dim Ptype                 As Long
  Dim PName                 As String
  Dim ProcKind              As String
  Dim ArrProc               As Variant
  Dim CPane                 As CodePane
  Dim cMod                  As CodeModule
  Dim bDummy                As Boolean
  Dim IndentLevel           As Double
  Dim StrBalance            As String
  Dim StrBalanceSubElements As String
  Dim strTrigger            As String
  Dim ErrStruct             As Variant
  Dim I                     As Long
  Dim CurPos                As Long
  Dim MyhourGlass           As cls_HourGlass
  Dim arrGoTo               As Variant
  Dim ExccessEndError       As Boolean

  'v2.2.0
  Set MyhourGlass = New cls_HourGlass
  Set CPane = GetActiveCodePane
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind, CurPos)
  IndentLevel = 1
  'v2.6.2 fixed format of multiline headers
  For I = LBound(ArrProc) To UBound(ArrProc)
    ArrProc(I) = Trim$(ArrProc(I))
    If SmartLeft(ArrProc(I), WARNING_MSG & "Incomplete ") Then
      ArrProc(I) = vbNullString
    End If
  Next I
  ArrayLineContCollapse ArrProc
  strF = Join(ArrProc, vbNewLine)
  IfThenArrayExpander ArrProc
  ArrProc = Split(DimMoveToTop(Join(ArrProc, vbNewLine)), vbNewLine)
  ArrProc = Split(DimMultiSingleTypeing(-1, Join(ArrProc, vbNewLine)), vbNewLine)
  ArrProc = Split(DimExpandMulti(-1, Join(ArrProc, vbNewLine)), vbNewLine)
  ArrProc = Split(DimTypeUpdate(-1, Join(ArrProc, vbNewLine)), vbNewLine)
  ArrProc = Split(DimUsage(-1, Join(ArrProc, vbNewLine)), vbNewLine)
  ArrProc = Split(DimFormat(Join(ArrProc, vbNewLine)), vbNewLine)
  arrGoTo = GetGoToTargetArray(ArrProc)
  For I = LBound(ArrProc) To UBound(ArrProc)
    If Not JustACommentOrBlank(ArrProc(I)) Then
      strF = ArrProc(I)
      DoSeparateCompoundLines -1, strF, cMod.Parent.Name, arrGoTo
      ArrProc(I) = strF
    End If
  Next I
  'v2.6.2 fixed format of multiline headers
  For I = 1 To UBound(ArrProc)
    strTrigger = LeftWord(ArrProc(I))
    Select Case strTrigger
      'v 2.2.1 added #If support
     Case "If", "Select", "Do", "With", "For", "While", "Open", "#If"
      ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
      StrBalance = AccumulatorString(StrBalance, strTrigger, , False)
      IndentLevel = IndentLevel + 1
     Case "Loop", "Next", "Wend", "Close", "#End"
      If MatchTrigger(StrBalance, strTrigger, CStr(ArrProc(I)), ExccessEndError) Then
        IndentLevel = IndentLevel - 1
        ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
       Else
        If ExccessEndError Then
          'v2.7.3 new fix detects excess Strutural ends
          ArrProc(I) = ArrProc(I) & vbNewLine & WARNING_MSG & "Excess End of Structure code."
          ExccessEndError = False
         Else
          ErrStruct = Split(StrBalance, ",")
          If UBound(ErrStruct) > -1 Then
            ArrProc(I) = ArrProc(I) & WARNING_MSG & "Incomplete " & ErrStruct(UBound(ErrStruct)) & " Structure"
           Else
            ArrProc(I) = ArrProc(I) & vbNewLine & WARNING_MSG & "Excess End of Structure code."
          End If
          Exit For
        End If
      End If
     Case "End"
      If Not strCodeOnly(ArrProc(I)) = "End" Then ' ignore 'End' command
        If strCodeOnly(ArrProc(I)) = "End " & ProcKind Then  ' special for end of Procedure
          If Len(StrBalance) Then
            ArrProc(I) = WARNING_MSG & "Incomplete " & StrBalance & " Structure(s)" & vbNewLine & ArrProc(I)
          End If
         Else
          If MatchTrigger(StrBalance, strTrigger, CStr(ArrProc(I)), ExccessEndError) Then
            If MatchTrigger(StrBalanceSubElements, "Case", CStr(ArrProc(I))) Then
              IndentLevel = IndentLevel - 1
              ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
             ElseIf MatchTrigger(StrBalanceSubElements, "Else", CStr(ArrProc(I))) Then
              IndentLevel = IndentLevel - 1
              ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
             ElseIf MatchTrigger(StrBalanceSubElements, "Elseif", CStr(ArrProc(I))) Then
              IndentLevel = IndentLevel - 1
              ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
             Else
              If ExccessEndError Then
                'v2.7.3 new fix detects excess Strutural ends
                ArrProc(I) = ArrProc(I) & vbNewLine & WARNING_MSG & "Excess End of Structure code."
                ExccessEndError = False
               Else
                IndentLevel = IndentLevel - 1
                ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
              End If
            End If
           Else
            ErrStruct = Split(StrBalance, ",")
            'v2.7.3 stops crash in single proc formatting
            If UBound(ErrStruct) > -1 Then
              ArrProc(I) = ArrProc(I) & vbNewLine & _
               WARNING_MSG & "Incomplete " & ErrStruct(UBound(ErrStruct)) & " Structure"
              Exit For
             Else
              'ArrProc(I) = ArrProc(I) & vbNewLine &
              'WARNING_MSG & "Incomplete " & ErrStruct(UBound(ErrStruct)) & " Structure"
              ArrProc(I) = ArrProc(I) & vbNewLine & WARNING_MSG & "Excess End of Structure code."
            End If
          End If
        End If
      End If
     Case "Case", "Else", "ElseIf", "#Else"
      If Not MatchTrigger(StrBalanceSubElements, strTrigger, CStr(ArrProc(I))) Then
        StrBalanceSubElements = AccumulatorString(StrBalanceSubElements, strTrigger, , False)
        IndentLevel = IndentLevel - 1
        ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
        IndentLevel = IndentLevel + 1
       Else
        StrBalanceSubElements = AccumulatorString(StrBalanceSubElements, strTrigger, , False)
        IndentLevel = IndentLevel - 1
        ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
        IndentLevel = IndentLevel + 1
      End If
     Case "Private", "Public", "Static", "Friend"
      'v3.0.4 stops proc header from indenting
     Case Else
      If LenB(ArrProc(I)) Then ' don't indent blanks
        If Not IsDimLine(ArrProc(I), False) Then
          ArrProc(I) = Space$(IndentLevel * IndentSize) & ArrProc(I)
        End If
      End If
    End Select
  Next I
  LineContinuationFix ArrProc, bDummy, -1, True
  ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine

End Sub

Public Sub ProcNoIndent()

  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant
  Dim CPane        As CodePane
  Dim cMod         As CodeModule
  Dim I            As Long

  'v2.2.0
  'If a procedure has BugTrap delete it else add it
  Set CPane = GetActiveCodePane
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind)
  'Trim and delete any Proc indent messages
  For I = TopOfRoutine + 1 To UBound(ArrProc)
    ArrProc(I) = Trim$(ArrProc(I))
    If SmartLeft(ArrProc(I), WARNING_MSG & "Incomplete ") Then
      ArrProc(I) = vbNullString
    End If
  Next I
  ArrayLineContCollapse ArrProc
  For I = LBound(ArrProc) To UBound(ArrProc)
    ArrProc(I) = Trim$(ArrProc(I))
  Next I
  ArrProc = CleanArray(ArrProc)
  ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
  ' End If

End Sub

Public Function ReadProcedureCodeArray(Optional ProcStart As Long, _
                                       Optional ProcEnd As Long, _
                                       Optional strProcName As String, _
                                       Optional lngPRocKind As Long, _
                                       Optional FirstCodeLine As Long, _
                                       Optional ProcKind As String, _
                                       Optional CursorPos As Long) As Variant

  Dim CurLine  As Long
  Dim startCol As Long
  Dim EndLine  As Long
  Dim endCol   As Long
  Dim CPane    As CodePane
  Dim cMod     As CodeModule
  Dim arrTmp   As Variant

  'v2.2.0
  'If a procedure has BugTrap delete it else add it
  Set CPane = GetActiveCodePane
  Set cMod = GetActiveCodeModule
  CPane.GetSelection CurLine, startCol, EndLine, endCol
  CursorPos = CurLine
  ProcStart = GetProcStartLine(cMod, CurLine, strProcName, lngPRocKind)
  If strProcName = "(Declarations)" Then
    ProcEnd = cMod.CountOfDeclarationLines + 1
    FirstCodeLine = 1
    ProcStart = 1
    ProcKind = "(Declarations)"
    arrTmp = Split(cMod.Lines(1, ProcEnd), vbNewLine)
   Else
    ' InsertSupportCode
    ProcEnd = GetProcEndLine(cMod, CurLine)
    arrTmp = Split(cMod.Lines(ProcStart, ProcEnd - ProcStart), vbNewLine)
    FirstCodeLine = GetProcCodeLineOfRoutine(arrTmp, True) 'v 2.2.3 corrects for line continuation
    ProcKind = GetProcClassStr(arrTmp(GetProcCodeLineOfRoutine(arrTmp)))
  End If
  '  End If
  ReadProcedureCodeArray = arrTmp
  With UnDoProcFormatData
    .FUHasData = True
    .FUProcName = strProcName
    Set .FUmodule = cMod
    .FUArray = arrTmp
  End With
  frm_RCMenus.mnuRCFormatOpt(2).Enabled = True

End Function

Public Function ReadProcedureCodeArray2(cMod As CodeModule, _
                                        CurLine As Long, _
                                        Optional ProcStart As Long, _
                                        Optional ProcEnd As Long, _
                                        Optional FirstCodeLine As Long) As Variant

  'v3.0.7 part of the new fast fixes
  
  Dim arrTmp      As Variant
  Dim strProcName As String

  'v2.2.0
  'If a procedure has BugTrap delete it else add it
  ProcStart = GetProcStartLine(cMod, CurLine, strProcName) ', lngPRocKind)
  If strProcName = "(Declarations)" Then
    ProcEnd = cMod.CountOfDeclarationLines + 1
    FirstCodeLine = 1
    ProcStart = 1
    arrTmp = Split(cMod.Lines(1, ProcEnd), vbNewLine)
   Else
    ' InsertSupportCode
    ProcEnd = GetProcEndLine(cMod, CurLine)
    arrTmp = Split(cMod.Lines(ProcStart, ProcEnd - ProcStart), vbNewLine)
    FirstCodeLine = GetProcCodeLineOfRoutine(arrTmp, True) 'v 2.2.3 corrects for line continuation
  End If
  '  End If
  ReadProcedureCodeArray2 = arrTmp
  With UnDoProcFormatData
    .FUHasData = True
    .FUProcName = strProcName
    Set .FUmodule = cMod
    .FUArray = arrTmp
  End With
  frm_RCMenus.mnuRCFormatOpt(2).Enabled = True

End Function

Public Sub ReplaceProcedureCode(cMod As CodeModule, _
                                ArrProc As Variant, _
                                ByVal ProcStart As Long, _
                                ByVal ProcEnd As Long, _
                                Optional ByVal bUseCodePane As Boolean = True)

  'v3.0.7 added bUseCodePane , the new fast fixes don't need it only the Context menu fixes
  
  Dim CurStartLine As Long
  Dim CurStartCol  As Long
  Dim CurEndLine   As Long
  Dim CurEndCol    As Long

  'v2.2.0
  With cMod
    If bUseCodePane Then
      .CodePane.GetSelection CurStartLine, CurStartCol, CurEndLine, CurEndCol 'store cursor pos
    End If
    .DeleteLines ProcStart, ProcEnd - ProcStart                             'destroy old code
    .InsertLines ProcStart, Join(ArrProc, vbNewLine)                        'insert new code
    If bUseCodePane Then
      .CodePane.SetSelection CurStartLine, CurStartCol, CurEndLine, CurEndCol 'restore cursor pos
    End If
  End With 'Cmod

End Sub

Public Sub RightClickIndent()

  Dim StartLine As Long
  Dim startCol  As Long
  Dim EndLine   As Long
  Dim endCol    As Long
  Dim cMod      As CodeModule

  'v2.2.0
  RefreshUserSettingsFromString
  'If a procedure has BugTrap delete it else add it
  Set cMod = GetActiveCodeModule
  cMod.CodePane.GetSelection StartLine, startCol, EndLine, endCol
  If GetProcName(cMod, StartLine) = "(Declarations)" Then
    mObjDoc.ShowWorking True, "Indenting Current Declaration Section", , False
    DecIndent
   Else
    mObjDoc.ShowWorking True, "Indenting Current Procedure", , False
    ProcIndent
  End If
  mObjDoc.ShowWorking False

End Sub

Public Sub RightMenuWithCreate()

  Dim UpDated      As Boolean
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant
  Dim cMod         As CodeModule
  Dim CurPos       As Long
  Dim oldFixLevel  As Long
  Dim strProc      As String
  Dim MyhourGlass  As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  RefreshUserSettingsFromString
  Set cMod = GetActiveCodeModule
  ProcNoIndent
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind, CurPos)
  If PName <> "(Declarations)" Then
    mObjDoc.ShowWorking True, "Creating With Structure in Current Procedure", , False
    ProcDeleteCFComments ArrProc
    oldFixLevel = FixData(DetectWithStructure).FixLevel
    FixData(DetectWithStructure).FixLevel = FixAndComment
    strProc = Join(ArrProc, vbNewLine)
    SuggestWithStructureProcedure strProc, cMod.Parent.Name, UpDated
    If UpDated Then
      ArrProc = Split(strProc, vbNewLine)
      strProc = Join(Split(strProc, vbNewLine), vbNewLine)
      SuggestWithStructureProcedure strProc, cMod.Parent.Name, UpDated
    End If
    If UpDated Then
      ReplaceProcedureCode cMod, Split(strProc, vbNewLine), PSLine, PEndLine
    End If
    FixData(DetectWithStructure).FixLevel = oldFixLevel
    ProcIndent
    mObjDoc.ShowWorking False
  End If

End Sub

Public Sub RightMenuWithPurify()

  Dim UpDated      As Boolean
  Dim MissingWith  As Boolean
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant
  Dim cMod         As CodeModule
  Dim CurPos       As Long
  Dim MyhourGlass  As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  RefreshUserSettingsFromString
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind, CurPos)
  If PName <> "(Declarations)" Then
    mObjDoc.ShowWorking True, "Purifying With Structure in Current Procedure", , False
    ProcDeleteCFComments ArrProc
    ProcWithStructurePurify ArrProc, UpDated, MissingWith
    If Not MissingWith Then
      If UpDated Then
        ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
      End If
     Else
      mObjDoc.Safe_MsgBox "The With Structures in this Procedure are incomplete", vbCritical
    End If
    mObjDoc.ShowWorking False
  End If

End Sub

Public Sub RightMenuWithRemove()

  Dim UpDated      As Boolean
  Dim MissingWith  As Boolean
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant
  Dim cMod         As CodeModule
  Dim CurPos       As Long
  Dim MyhourGlass  As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind, CurPos)
  If PName <> "(Declarations)" Then
    mObjDoc.ShowWorking True, "Removing With Structure in Current Procedure", , False
    ProcDeleteCFComments ArrProc
    ProcWithStructureRemove ArrProc, UpDated, MissingWith
    If Not MissingWith Then
      If UpDated Then
        ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
      End If
     Else
      mObjDoc.Safe_MsgBox "The With Structures in this Procedure are incomplete", vbCritical
    End If
    mObjDoc.ShowWorking False
  End If

End Sub

Public Sub UnDoProcFormat()

  Dim StartLine   As Long
  Dim EndLine     As Long
  Dim MyhourGlass As cls_HourGlass

  'v2.2.0
  Set MyhourGlass = New cls_HourGlass
  With UnDoProcFormatData
    If .FUHasData Then
      If GetProcPositionFromName(.FUmodule, .FUProcName, StartLine, EndLine) Then
        ReplaceProcedureCode .FUmodule, .FUArray, StartLine, EndLine
      End If
      .FUHasData = False
      Set .FUmodule = Nothing
      .FUArray = Array("")
    End If
    frm_RCMenus.mnuRCFormatOpt(2).Enabled = False
    '
  End With

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:27:36 AM) 17 + 647 = 664 Lines Thanks Ulli for inspiration and lots of code.

