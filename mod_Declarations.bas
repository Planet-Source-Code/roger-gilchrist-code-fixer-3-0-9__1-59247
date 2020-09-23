Attribute VB_Name = "mod_Declarations"

Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public bPaulCatonSubClasUsed             As Boolean
Public strPaulCatonSubClasProcName       As String
Public strPaulCatonSubClasCompName       As String
Public UsingDefTypes                     As Boolean
Public Const OptExplicitMsgVerbose       As String = WARNING_MSG & "Option Explicit makes coding much safer but will cause difficulties until you declare all variables." & vbNewLine & _
 RGSignature & "(Code Fixer identifies most of them as Missing Dims)" & vbNewLine & _
 RGSignature & "Run code using [Ctrl]+[F5] to find any others yourself." & vbNewLine & _
 RGSignature & "To auto-insert 'Option Explicit' in all new code: Tools|Options...|Editor tab, check 'Require Variable Declaration'." & vbNewLine
Public Const OptExplicitMsg              As String = WARNING_MSG & "All variables must now be declared." & vbNewLine & _
 RGSignature & "Run code using [Ctrl]+[F5] to find undeclared variables that Code Fixer misses." & vbNewLine
Private Const EnumProtDesc               As String = "'Trick preserves Case of Enums when typing in IDE"
Private Const EnumSuggestion             As String = SUGGESTION_MSG & "Inserted by Code Fixer. (Must be placed after Enum Declaration for Code Fixer to recognize it properly)"
Private Const EnumMultiCantUpdate        As String = WARNING_MSG & "Multiple line Enum Capitalization Protection can not be updated, just delete the protection and it will be rebuilt."
Private Const EnumProtEndofDecMsg        As String = WARNING_MSG & "Due to a problem with detecting '#End If' as a part of the Declarations if it is last item in declarations" & vbNewLine & _
 RGSignature & "Code Fixer inserts the following very silly Declaration for safety purposes." & vbNewLine & _
 "Private Fake_To_Protect_ECP_at_end_of_Declarations as Boolean ' Move_Enum_Away_from_End_ofDeclarations_And_Delete_Me"

Private Function AutoCodeSize(ByVal FName As String) As Single

  Dim frefil    As Long
  Dim sngSize   As Single
  Dim strRecord As String
  Dim bNearEnd  As Boolean
  Dim strTmp    As String

  'based on Kenneth Ives' AutoGen_Code routine
  'in his upload 'Count Lines of Code v1.0' 'Count_Loc.zip' 01-NOV-2000
  'his version counted the lines, mine calculates the size of the
  'hidden control description part of controls
  'Called from Function ModuleSize
  frefil = FreeFile
  Open FName For Input Access Read As frefil
  Do
    'rad line
    Line Input #frefil, strRecord
    If Len(strRecord) Then
      strTmp = Left$(strRecord, 10)
      'looking for end of section
      If Not bNearEnd Then
        If strTmp = "Attribute " Then
          bNearEnd = True ' found lines which are near the end of section
        End If
       Else
        'now near the end so check still in hidden section
        If strTmp <> "Attribute " Then
          'reached code or comments that will appear in IDE so stop looking
          Exit Do
        End If
      End If
      'add line size to sngSize
      sngSize = sngSize + Len(strRecord) + 1
      ' just add one for newline (bit dubious but i think correct)
     Else
      sngSize = sngSize + 1 ' just add one for blank lines (bit dubious but i think correct)
    End If
  Loop Until EOF(frefil)
  Close #frefil
  AutoCodeSize = sngSize

End Function

Private Sub BadNameVariableTest(ByVal ModuleNumber As Long, _
                                dArray As Variant)

  Dim strTestMe  As String
  Dim L_CodeLine As String
  Dim UpDated    As Boolean
  Dim TmpDecArr  As Variant
  Dim I          As Long
  Dim MaxFactor  As Long

  'This routine tests for variables with poor naming conventions
  'Thanks to Dream whose bug report lead me to think of this
  If dofix(ModuleNumber, BadNameVariableWarning_CXX) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      For I = 1 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Dim", "Global", "Public", "Private") Then
            If Not InstrAtPositionArray(L_CodeLine, ipAny, True, "WithEvents", "Declare", "Event") Then
              strTestMe = GetName(Split(L_CodeLine))
              'v2.8.3 speed improvement
              'NB all tests needed there could be more than one error in naming
              If Len(strTestMe) = 1 Then
                If InStr("xy", LCase$(strTestMe)) Then
                  If NotUsedOutSideMouseEvents(ModuleNumber, L_CodeLine, strTestMe) Then
                    'v3.0.7 new fix gets MOdule level X, Y variables that are not actually used
                    TmpDecArr(I) = Marker("''" & TmpDecArr(I), WARNING_MSG & "Single letter Variable '" & strTestMe & "' not used except in VB native Mouse Events.", MAfter, UpDated)
                   Else
                    TmpDecArr(I) = Marker(TmpDecArr(I), WARNING_MSG & "Single letter Variables 'X' or 'Y' make code difficult to read as VB uses them in Mouse Events." & SuggestNewName(strTestMe, L_CodeLine, True) & RGSignature & "If you are only using it as a For structures counter use a Dim instead" & IIf(Xcheck(XVerbose), RGSignature & "(may cause local Dims to be marked as duplicates)", vbNullString), MAfter, UpDated)
                  End If
                 Else
                  TmpDecArr(I) = Marker(TmpDecArr(I), WARNING_MSG & "Single letter Variables make code difficult to read, use a meaningful name." & SuggestNewName(strTestMe, L_CodeLine, True) & RGSignature & "If you are only using it as a For structures counter use a Dim instead" & IIf(Xcheck(XVerbose), RGSignature & "(may cause local Dims to be marked as duplicates)", vbNullString), MAfter, UpDated)
                End If
                AddNfix BadNameVariableWarning_CXX
              End If
              If BadVBName(strTestMe) Then
                TmpDecArr(I) = Marker(TmpDecArr(I), WARNING_MSG & "Variables with same name as a VB Command/Property make code difficult to read." & SuggestNewName(strTestMe, L_CodeLine) & DoNotSandRMsg("Command/Property"), MAfter, UpDated)
                AddNfix BadNameVariableWarning_CXX
              End If
              If CntrlDescMember(strTestMe) > -1 Then
                TmpDecArr(I) = Marker(TmpDecArr(I), WARNING_MSG & "Variables with same name as a control make code difficult to read." & SuggestNewName(strTestMe, L_CodeLine) & DoNotSandRMsg("Control"), MAfter, UpDated)
                AddNfix BadNameVariableWarning_CXX
              End If
            End If
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Function BadVBName(ByVal strTest As String) As Boolean

  If IsControlProperty(strTest) Then
    BadVBName = True
   ElseIf isRefLibVBCommands(strTest, False) Then
    BadVBName = True
   ElseIf InQSortArray(ArrQVBStructureWords, strTest) Then
    BadVBName = True
   ElseIf InQSortArray(ArrQVBReservedWords, strTest) Then
    BadVBName = True
  End If

End Function

Public Sub CleanUp_Engine()

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim J            As Long
  Dim NumFixes     As Long
  Dim Sline        As Long

  NumFixes = 7
  On Error GoTo BugHit
  If Not bAborting Then
    'This routine is a wrapper for the following actions
    'separated out so that the XP Frame tester can leave a message
    mObjDoc.KillComments
    WorkingMessage "Remove Old Code Fixer Comments", 1, NumFixes
    KillContaining Smiley & "Code Fixer V"
    WorkingMessage "Replace Rems", 2, NumFixes
    ReplaceRem
    WorkingMessage "Remove Old Code Fixer Comments", 3, NumFixes
    KillContaining "'APPROVED(Y )"
    WorkingMessage "Remove Old Code Fixer Comments", 4, NumFixes
    For J = 2 To UBound(CodeFixProtectedArray)
      MemberMessage "", 1 + J, 4
      KillContaining Left$(CodeFixProtectedArray(J), 31)
    Next J
    bModDescExists = bModDescExists
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          With Comp
            SaveMemberAttributes CurCompCount, .CodeModule.Members
            'v2.9.5 This was the uncommented insertion point for Option Explict (not needed). thanks Ian K
            'GetDeclarationArray .CodeModule
            'First 2 affect whole module
            CleanupPreviousRun .CodeModule, CurCompCount
          End With 'Comp
        End If
      Next Comp
      If bAborting Then
        Exit For '      Sub
      End If
    Next Proj
  End If
  'v3.0.4 Disable Active Timers on User Controls fires before clean up
  'so uses a special EARLYWARNING_MSG which survives the CF message removal
  'this changes it into the standard WARNING_MSG
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      With Comp
        Sline = 1
        If LenB(.Name) Then
          Do While .CodeModule.Find(EARLYWARNING_MSG, Sline, 1, -1, -1, False)
            .CodeModule.ReplaceLine Sline, Replace$(.CodeModule.Lines(Sline, 1), EARLYWARNING_MSG, WARNING_MSG)
          Loop
        End If
      End With 'Comp
    Next Comp
  Next Proj
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "CleanUp_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub CleanUpendOfDecOptCompile(cMod As CodeModule, _
                                      dArray As Variant)

  If UBound(dArray) > -1 Then
    If dArray(UBound(dArray)) = CodeFixProtectedArray(endDec) Then
      'this protects end of declarations "#End If' from being duplicated
      If cMod.Lines(cMod.CountOfDeclarationLines + 1, 1) = "#End If" Then
        cMod.DeleteLines cMod.CountOfDeclarationLines + 1, 1
      End If
    End If
  End If

End Sub

Private Sub CleanupPreviousRun(cMod As CodeModule, _
                               ByVal ModuleNumber As Long)

  
  Dim I       As Long
  Dim LinC    As Long
  Dim strT    As String
  Dim lngJunk As Long
 ' Dim doAgain As Boolean
  ' used to patch the error caused by trimming offsets too often in a line continuation line
  'Remove all indenting
  'Collapse line continuation code to single lines
  'v2.3.8 faster
  'v2.3.9 actions reordered to deal with a problem of
  'meanless line continuations at end of a code line followed by a blank line
  'Thanks Ian J. Keiller who brought Evan Toder's CalmMent program
  'which contained this problem to my attention
  With cMod
    If .CountOfLines > 1 Then
      'trim all lines
      If Not ModDesc(ModuleNumber).MDDontTouch Then
        'v2.6.4'  'Changed to use DO because the inserted messages can change CountOfLines
        'and cause clearner to miss end of modules
        'TRIM ALL LINES IN CODE
        Do
          I = I + 1
          'v2.4.4 deal with very long lines thanks Ulli
          On Error GoTo LineTooLong ' Errors always Resume to here
          DoEvents
          strT = Trim$(.Lines(I, 1))
          .ReplaceLine I, strT
        Loop Until I >= .CountOfLines
        On Error GoTo 0
        'remove line continuation
        If .Find(ContMark, lngJunk, lngJunk, lngJunk, lngJunk) Then
          LinC = .CountOfLines
          'v2.6.4'  'Changed to use DO because the inserted messages can change CountOfLines
          I = 0
          WorkingMessage "Reduce Line Continuation Comments", 4, 6
          Do
            I = I + 1
            'For I = 1 To .CountOfLines 'Line Cont comments
            MemberMessage "", I, LinC
            If Left$(.Lines(I, 1), 1) = "'" Then
              If Right$(.Lines(I, 1), 2) = ContMark Then
                .ReplaceLine I, Left$(.Lines(I, 1), Len(.Lines(I, 1)) - 2)
                .ReplaceLine I + 1, "'" & .Lines(I + 1, 1)
              End If
            End If
          Loop Until I >= .CountOfLines
          'Next I
          LinC = .CountOfLines
          'v2.6.4'  'Changed to use DO because the inserted messages can change CountOfLines
          I = 0
          WorkingMessage "Reduce Line Continuation Code", 5, 6
          Do
            I = I + 1
            'For I = 1 To .CountOfLines ' To 1 Step -1
            MemberMessage "", I, LinC
            If Right$(.Lines(I, 1), 2) = ContMark Then
              If Len(Left$(.Lines(I, 1), Len(.Lines(I, 1)) - 1) & Trim$(.Lines(I + 1, 1))) < 800 Then
                .ReplaceLine I, Left$(.Lines(I, 1), Len(.Lines(I, 1)) - 1) & Trim$(.Lines(I + 1, 1))
                .DeleteLines I + 1
                I = I - 1
              ' Else
              '  doAgain = True
              End If
            End If
          Loop Until I >= .CountOfLines
'          If doAgain Then
'            I = 0
'            Do
'              I = I + 1
'              'For I = 1 To .CountOfLines ' To 1 Step -1
'              MemberMessage "", I, LinC
'              If Right$(.Lines(I, 1), 2) = ContMark Then
'                If Len(Left$(.Lines(I, 1), Len(.Lines(I, 1)) - 1) & Trim$(.Lines(I + 1, 1))) < 800 Then
'                  .ReplaceLine I, Left$(.Lines(I, 1), Len(.Lines(I, 1)) - 1) & Trim$(.Lines(I + 1, 1))
'                  .DeleteLines I + 1
'                  I = I - 1
'                End If
'              End If
'            Loop Until I >= .CountOfLines
'          End If
          'Next I
        End If
        'preserve or delete blanks
        LinC = .CountOfLines
        WorkingMessage "Clean Blanks", 6, 6
        For I = .CountOfLines To 1 Step -1
          MemberMessage "", LinC - I, LinC
          If LenB(Trim$(.Lines(I, 1))) = 0 Then
            If Xcheck(XBlankPreserve) Then
              .ReplaceLine I, REMBLANK
             Else
              .DeleteLines I
            End If
          End If
        Next I
        'only allow single blanks
        If .Find(REMBLANK, lngJunk, lngJunk, lngJunk, lngJunk) Then
          If Xcheck(XBlankPreserve) Then
            For I = .CountOfLines To 2 Step -1
              If .Lines(I, 1) = REMBLANK Then
                If .Lines(I - 1, 1) = REMBLANK Then
                  .DeleteLines I
                End If
              End If
            Next I
          End If
        End If
      End If
    End If
  End With

Exit Sub

  'v2.4.4 deal with very long lines thanks Ulli
LineTooLong:
  Select Case Err.Number
   Case 40192
    'inserting the comment then doing the fix seems to clear up the error ????
    With cMod
      .InsertLines I + 1, ""
      .InsertLines I + 1, WARNING_MSG & "Code Fixer had some difficulty with the previous line," & vbNewLine & _
       RGSignature & "A fully formatted version may not be possible (too long or too many Line continuations)." & vbNewLine & _
       RGSignature & "If possible you should break the line into 2 lines."
      strT = Trim$(.Lines(I, 1))
      .ReplaceLine I, strT
    End With 'cMod
   Case Else
    With cMod
      .InsertLines I + 1, WARNING_MSG & "Error " & Err.Number & " occurred in CleanupPreviousRun"
      strT = Trim$(.Lines(I, 1))
      .ReplaceLine I, strT
    End With
  End Select
  Resume Next

End Sub

Private Sub CorrectIncompleteProperty(cMod As CodeModule)

  Dim Sline       As Long
  Dim SLine2      As Long
  Dim strCode     As String
  Dim strPropName As String

  'v2.6.9 Thanks to Mike Ulik(for finding the bug) and Ulli Muehlenweg (for pushing me to fix it)
  'The DoSorting routine deletes incomplete Properties (Read only or Write only Properties)
  'if they are Typed with a same name Enum (or Type?)
  'this is legal in VB but disrupts VB's ability to recognise them properly as members
  'This routine builds a dummy compliment to the existing Get or Let and inserts a message
  'explaining why and how to fix it
  'This routine will NOT complete Property structures unless they also have the Enum naming problem
  'That may come later
  '
  'v2.7.0 better test in case the PropName is a substring of some other part of the procedure header
  ' included Test for Property Set
  'for speed I dont't figure out if Let/Set is the correct complimentry part of Get
  'If you get this response checking the dummy and message should be enough
  With cMod
    Do While .Find("Property", Sline, 1, -1, -1, True)
      strCode = .Lines(Sline, 1)
      'v2.7.5 reset SLine2 otherwise it doesn't look back and floods the for with duplicates
      SLine2 = 0
      If Not JustACommentOrBlank(strCode) Then
        strPropName = GetProcName(cMod, Sline)
        If FindCodeUsage("Enum " & strPropName, "") Then
          If InStr(strCode, "Property Let") Or InStr(strCode, "Property Set") Then
            If InCode(strCode, InStr(strCode, "Property Let")) Or InCode(strCode, InStr(strCode, "Property Set")) Then
              If Not .Find("Property Get " & strPropName, SLine2, 1, -1, -1) Then
                If CountSubStringWhole(strCode, strPropName) > 1 Then
                  If strPropName <> "Font" Then 'v2.7.4
                    .AddFromString CorrectIncompletePropertyDummy("Get", strPropName, strPropName)
                    Sline = Sline + 1
                  End If
                End If
              End If
            End If
          End If
          If InStr(strCode, "Property Get") Then
            If InCode(strCode, InStr(strCode, "Property Get")) Then
              If Not .Find("Property Let " & strPropName, SLine2, 1, -1, -1) Then
                If Not .Find("Property Set " & strPropName, SLine2, 1, -1, -1) Then
                  If CountSubStringWhole(strCode, strPropName) > 1 Then
                    If strPropName <> "Font" Then
                      .AddFromString CorrectIncompletePropertyDummy("Let", strPropName, strPropName)
                      Sline = Sline + 1
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      Sline = Sline + 1
      If Sline = 0 Or Sline > .CountOfLines Then
        Exit Do
      End If
    Loop
  End With

End Sub

Private Function CorrectIncompletePropertyDummy(GetLetSet As String, _
                                                ProcName As String, _
                                                strType As String) As String

  'v2.6.9 support routine for CorrectIncompleteProperty

  Select Case GetLetSet
   Case "Set", "Let"
    CorrectIncompletePropertyDummy = "Property " & GetLetSet & " " & ProcName & "(PropVal As " & strType & ")"
   Case "Get"
    CorrectIncompletePropertyDummy = "Property Get " & ProcName & "() As " & strType
  End Select
  CorrectIncompletePropertyDummy = vbNewLine & _
   CorrectIncompletePropertyDummy & vbNewLine & _
   WARNING_MSG & "Dummy Property inserted to avoid a problem caused by an incomplete Property with an Enum/type with same name." & vbNewLine & _
   WARNING_MSG & "It is strongly recommended that you change the name of the Property or the Enum/Type" & vbNewLine & _
   "End Property"

End Function

Public Sub Create_Enum_Capitalisation_Protection(ByVal ModuleNumber As Long, _
                                                 dArray As Variant)

  
  Dim L_CodeLine              As String
  Dim MultilineRepairDetected As Boolean
  Dim strTemp                 As String
  Dim TmpDecArr               As Variant
  Dim I                       As Long
  Dim J                       As Long
  Dim UpDated                 As Boolean
  Dim MaxFactor               As Long
  Dim TmpA                    As Variant
  Dim InsComma                As String
  Dim LenCount                As Long
  Dim lngRealCodeline         As Long
  Dim EnumError               As Boolean
  Dim EnumStart               As Long

  'v 2.2.2 Thanks Dipankar Basu updated detectors so that commets/spaces between Enum and Fix
  'are recognised to stop 2nd fix being added.
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'This routine inserts the Enum Capitalisation Protection Trick
  'into code immediately after any Enum
  'This is a bit ineffecient because it collects the data then
  'test if it is needed BUT this allows it to update if necessary
  If dofix(ModuleNumber, CreateEnumCapProtect) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > -1 Then
      For I = 0 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If InstrAtPosition(L_CodeLine, "Enum", ipLeftOr2nd) Then
            'generate the trick's code
            EnumStart = I
            I = I + 1
            L_CodeLine = ExpandForDetection(TmpDecArr(I))
            strTemp = vbNullString
            Do
              If Not JustACommentOrBlank(L_CodeLine) Then
                If InStr(L_CodeLine, " =") Then
                  L_CodeLine = Left$(L_CodeLine, InStr(L_CodeLine, " =") - 1)
                End If
                If Left$(L_CodeLine, 1) = "[" And Right$(L_CodeLine, 1) = "]" Then
                  If Mid$(L_CodeLine, 2, Len(L_CodeLine) - 2) <> StripPunctuation(Mid$(L_CodeLine, 2, Len(L_CodeLine) - 2), , True) Then
                    L_CodeLine = vbNullString
                   Else
                    L_CodeLine = Mid$(L_CodeLine, 2, Len(L_CodeLine) - 2)
                  End If
                 Else
                  If L_CodeLine <> StripPunctuation(L_CodeLine, , True) Then
                    L_CodeLine = vbNullString
                  End If
                End If
                If IsNumeric(Left$(L_CodeLine, 1)) Then
                  L_CodeLine = vbNullString
                End If
                'v2.9.0 extra ignore words
                If InQSortArray(ArrQVBCommands, StripPunctuation(L_CodeLine)) Then
                  L_CodeLine = vbNullString
                End If
                If InQSortArray(ArrQVBReservedWords, StripPunctuation(L_CodeLine)) Then
                  L_CodeLine = vbNullString
                End If
                If Left$(L_CodeLine, 1) = "_" Then
                  'v2.3.5 fixed
                  L_CodeLine = vbNullString
                End If
                If LenB(L_CodeLine) Then
                  strTemp = AccumulatorString(strTemp, Trim$(L_CodeLine), CommaSpace)
                End If
              End If
              I = I + 1
              'v2.3.8 if Enum code is incomplete
              If I > UBound(TmpDecArr) Then
                I = UBound(TmpDecArr)
                EnumError = True
                Exit Do
              End If
              L_CodeLine = ExpandForDetection(TmpDecArr(I))
            Loop Until Left$(L_CodeLine, 8) = "End Enum"
            'i is now the position of End Enum so this is where we will attach the ECP code
            If LenB(strTemp) Then
              'create code line(s)
              ' cope with very long lines by creating multiple lines
              If Not EnumError Then
                LenCount = 0
                InsComma = vbNullString
                TmpA = Split(strTemp, CommaSpace)
                strTemp = "Private "
                For J = LBound(TmpA) To UBound(TmpA)
                  If LenB(TmpA(J)) Then
                    strTemp = strTemp & InsComma & TmpA(J)
                    InsComma = CommaSpace
                    LenCount = LenCount + Len(TmpA(J))
                    If LenCount >= lngLineLength Then
                      If J < UBound(TmpA) Then
                        MultilineRepairDetected = True
                        strTemp = strTemp & vbNewLine & "Private "
                        InsComma = vbNullString
                        LenCount = 0
                      End If
                    End If
                  End If
                Next J
              End If
              'test that enum does not already contain an ECP
              If Not EnumError Then
                If Not TestNextLineOfDeclaration(I, TmpDecArr, Hash_If_False_Then, lngRealCodeline) Then
                  TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                   Hash_If_False_Then & EnumProtDesc & vbNewLine & _
                   strTemp & vbNewLine & _
                   Hash_End_If & vbNewLine & _
                   EnumSuggestion
                  If MultilineRepairDetected Then
                    TmpDecArr(I) = TmpDecArr(I) & vbNewLine & EnumMultiCantUpdate
                  End If
                  If lngRealCodeline = UBound(TmpDecArr) Then
                    TmpDecArr(I) = TmpDecArr(I) & EnumProtEndofDecMsg
                  End If
                  UpDated = True
                  AddNfix CreateEnumCapProtect
                 Else
                  If Not MultilineRepairDetected Then
                    If Left$(NextCodeLine(TmpDecArr, I + 1, 1), Len(strTemp)) <> strTemp Then
                      TmpDecArr(I + 2) = strTemp
                      UpDated = True
                      AddNfix CreateEnumCapProtect
                    End If
                    'Else
                    'Unwritten multiline updater
                  End If
                End If
               Else
                TmpDecArr(EnumStart) = TmpDecArr(EnumStart) & vbNewLine & _
                 WARNING_MSG & " Enum Structure incomplete"
                Exit For
              End If
            End If
          End If
        End If
      Next I
    End If
    dArray = CleanArray(TmpDecArr, UpDated)
  End If

End Sub

Public Sub Declaration_Engine()

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CompMod      As CodeModule
  Dim DecArray     As Variant
  Dim CurCompCount As Long
  Dim NumFixes     As Long

  NumFixes = 16
  'This routine is a wrapper for the following actions
  On Error GoTo BugHit
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount, False) Then
          'each fix conducts its own DontTouch Test
          ModuleMessage Comp, CurCompCount
          DisplayCodePane Comp
          Set CompMod = Comp.CodeModule
          DecArray = GetDeclarationArray(CompMod)
          WorkingMessage "SubClassing check", 1, NumFixes
          SubClassDetectorInitialise CompMod, DecArray
          With Comp
            CleanUpendOfDecOptCompile .CodeModule, DecArray
            If FileSize(.FileNames(1)) > 64 Then
              InsertLargeFileWarning CurCompCount, DecArray, .FileNames(1)
            End If
          End With 'Comp
          WorkingMessage "Expand Compound Lines Declaration", 2, NumFixes
          SeparateCompoundDeclarationLines CurCompCount, DecArray, Comp.Name
          WorkingMessage "API Declares to End of Declarations", 3, NumFixes
          Move_API_Declares_To_End CurCompCount, DecArray
          'v2.7.3 find based updater
          WorkingMessage "Update Dim/Global", 4, NumFixes
          DeclareDimGlobalUpdate CurCompCount, Comp, DecArray
          UpDate_Dim_Global_Declarations CurCompCount, DecArray, Comp
          WorkingMessage "Expand Multiple Single Type Declarations", 5, NumFixes
          Expand_SingleLine_SingleType_Declaration CompMod, DecArray
          WorkingMessage "Expand Multiple Declarations", 6, NumFixes
          Expand_SingleLine_Multiple_Declaration CompMod, DecArray
          WorkingMessage "Typing Declaration Constants", 7, NumFixes
          Assign_Type_To_Constants CurCompCount, DecArray
          WorkingMessage "Type Suffix Declarations", 8, NumFixes
          UpDate_TypeSuffix_Declarations CurCompCount, DecArray
          WorkingMessage "Type-Cast DefType", 9, NumFixes
          Update_DefType_to_AsType_Declaration CurCompCount, DecArray
          WorkingMessage "Type-Cast Declaration Variables", 10, NumFixes
          TypeCastUnTypedDeclarations CurCompCount, DecArray, Comp
          WorkingMessage "Scope Dim/Global", 11, NumFixes
          UpDate_Excess_Public_To_Private_Declarations CurCompCount, DecArray, Comp
          WorkingMessage "Form Public", 12, NumFixes
          Fix_Form_Public_Declarations CurCompCount, DecArray, Comp
          WorkingMessage "Enum Capitalisation Protection", 13, NumFixes
          Create_Enum_Capitalisation_Protection CurCompCount, DecArray
          WorkingMessage "Poorly named Variable", 14, NumFixes
          BadNameVariableTest CurCompCount, DecArray
          WorkingMessage "Offseting Dec Type & EOL comments", 15, NumFixes
          DeclarationOffSetAsTypeEOL CurCompCount, DecArray
          WorkingMessage "Protect Enum Capitaisation", 16, NumFixes
          ProtectEnumCap DecArray
          WorkingMessage "Reduce Excess Scope", 17, NumFixes
          ExcessiveScopeForVariables CurCompCount, DecArray
          WorkingMessage "Long Line Fix", 18, NumFixes
          LongLineFix CurCompCount, DecArray
          WorkingMessage "", 0, 1
          If Not ModDesc(CurCompCount).MDDontTouch Then
            ReWriter CompMod, DecArray, RWDeclaration
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
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "Declaration_Engine"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Sub DeclarationDefTypeDetector(ByVal ModuleNumber As Long, _
                                      dArray As Variant, _
                                      ByVal AddComment As Boolean, _
                                      Optional UpDated As Boolean, _
                                      Optional ArrReturn As Variant)

  Dim L_CodeLine As String
  Dim strTmp     As String
  Dim TArray     As Variant
  Dim I          As Long
  Dim J          As Long
  Dim arrTest1   As Variant

  arrTest1 = Array("DefLng", "DefSng", "DefInt", "DefBool", "DefByte", "DefCur", "DefDbl", "DefDate", "DefStr", "DefUbj", _
                   "DefVar")
  'ver 1.0.88 update AddComment stops double message benig added
  'Declaration Section doesntt AddComment, Dim section does
  'NOTE DefType is module level so need to be reset as each module is entered
  UsingDefTypes = False
  If dofix(ModuleNumber, UpdateDefType2AsType) Then
    For I = LBound(dArray) To UBound(dArray)
      L_CodeLine = dArray(I)
      If Not JustACommentOrBlank(L_CodeLine) Then
        If MultiLeft(Trim$(L_CodeLine), True, "DefLng", "DefSng", "DefInt", "DefBool", "DefByte", "DefCur", "DefDbl", "DefDate", "DefStr", "DefUbj", "DefVar") Then
          TArray = Split(ExpandForDetection(L_CodeLine))
          If IsInArray(TArray(0), arrTest1) Then
            TArray = GenerateDefTypeTargetArray(TArray)
            strTmp = vbNullString
            For J = 1 To UBound(TArray)
              strTmp = AccumulatorString(strTmp, UCase$(TArray(J)))
            Next J
            If LenB(strTmp) Then
              UsingDefTypes = True
              SetDefArrays TArray, strTmp, ModuleNumber
            End If
          End If
          If AddComment Then
            UpDated = True
            Select Case FixData(UpdateDefType2AsType).FixLevel
             Case CommentOnly
              dArray(I) = dArray(I) & vbNewLine & RGSignature & "DefType no longer needed."
             Case FixAndComment
              dArray(I) = SQuote & dArray(I) & RGSignature & "DefType no longer needed."
             Case JustFix
              dArray(I) = vbNullString
            End Select
          End If
        End If
      End If
    Next I
    If AddComment Then
      ArrReturn = Join(dArray, vbNewLine)
    End If
  End If

End Sub

Public Sub DeclarationOffSetAsTypeEOL(ByVal ModuleNumber As Long, _
                                      dArray As Variant)

  Dim L_CodeLine   As String
  Dim TestPos      As Long
  Dim TmpDecArr    As Variant
  Dim AsOffSet     As Long
  Dim EOLOffSet    As Long
  Dim CommentStore As String
  Dim UpDated      As Boolean
  Dim I            As Long
  Dim MaxFactor    As Long
  Dim TypeOffSet   As Long
  Dim StrTypeName  As String

  'This routine was suggested by
  'Coding Genius's article at PSC 'Code Formatting Techniques'
  'Submitted on: 9/24/2002 1:07:56 PM in the section Code Formatting – Variables
  'It identifies the longest Declaration Variable with an "As Type"
  'then offsets all "As Type" to form a single column
  'NO coding value but improves readablity
  If dofix(ModuleNumber, LayOutAtTypeEOLComment) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > -1 Then
      For I = 0 To MaxFactor
        MemberMessage "", I / 2, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If ExtractCode(L_CodeLine, CommentStore) Then
            L_CodeLine = StripDoubleSpace(L_CodeLine)
            If Has_AS(L_CodeLine) Then
              'v2.2.1 tighter test
              If Not InstrAtPosition(L_CodeLine, "Declare", ipLeftOr2ndOr3rd) Then
                TestPos = Get_As_Pos(L_CodeLine) + IndentSize
                If InCode(L_CodeLine, TestPos) Then
                  If TestPos > AsOffSet Then
                    AsOffSet = TestPos
                  End If
                End If
              End If
            End If
          End If
        End If
      Next I
      AsOffSet = AsOffSet * 1.1
      If AsOffSet Then
        For I = LBound(TmpDecArr) To UBound(TmpDecArr)
          L_CodeLine = TmpDecArr(I)
          If Not JustACommentOrBlank(L_CodeLine) Then
            'v2.4.4 Thanks Mike Ulik fixed to cope with proper offset for type members
            'test was in wrong place
            If InTypeDef(L_CodeLine, StrTypeName) Then
              TypeOffSet = IndentSize
             Else
              TypeOffSet = 0
            End If
            L_CodeLine = Trim$(StripDoubleSpace(L_CodeLine))
            If Has_AS(L_CodeLine) Then
              'v2.2.1 tighter test
              If Not InstrAtPosition(L_CodeLine, "Declare", ipLeftOr2ndOr3rd) Then
                TestPos = Get_As_Pos(L_CodeLine)
                If InCode(L_CodeLine, TestPos) Then
                  TmpDecArr(I) = Left$(L_CodeLine, TestPos - 1) & Space$(AsOffSet - TypeOffSet - TestPos) & Mid$(L_CodeLine, TestPos)
                  UpDated = True
                End If
              End If
            End If
          End If
        Next I
      End If
      For I = 0 To MaxFactor
        MemberMessage "", MaxFactor / 2 + I / 2, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If ExtractCode(L_CodeLine, CommentStore) Then
            If Has_AS(L_CodeLine) Then
              If LenB(CommentStore) Then
                If LenB(L_CodeLine) > EOLOffSet Then
                  EOLOffSet = Len(L_CodeLine)
                End If
              End If
            End If
          End If
        End If
      Next I
      EOLOffSet = EOLOffSet + 2
      If EOLOffSet Then
        For I = LBound(TmpDecArr) To UBound(TmpDecArr)
          L_CodeLine = TmpDecArr(I)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If ExtractCode(L_CodeLine, CommentStore) Then
              If Has_AS(L_CodeLine) Then
                If LenB(CommentStore) Then
                  If EOLOffSet - Len(L_CodeLine) > 1 Then
                    TmpDecArr(I) = L_CodeLine & Space$(EOLOffSet - Len(L_CodeLine) + 1) & CommentStore
                  End If
                  UpDated = True
                End If
              End If
            End If
          End If
        Next I
      End If
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Function DEFXXXUsed(cMod As CodeModule) As Boolean

  Dim I      As Long
  Dim arrTmp As Variant

  'v2.8.3 Thanks Joakim Schramm
  'this stops 'Expand_SingleLine_SingleType_Declaration' from running
  'if DefXXX is being used (otherwise untyped are updated incorrectly)
  arrTmp = GetDeclarationArray(cMod)
  If UBound(arrTmp) > 0 Then
    For I = 1 To UBound(arrTmp)
      If MultiLeft(Trim$(arrTmp(I)), True, "DefLng", "DefSng", "DefInt", "DefBool", "DefByte", "DefCur", "DefDbl", "DefDate", "DefStr", "DefUbj", "DefVar") Then
        DEFXXXUsed = True
        Exit For
      End If
    Next I
  End If

End Function

Public Function DoNotSandRMsg(ByVal strA As String) As String

  DoNotSandRMsg = RGSignature & "Do NOT use Search and Replace; you will have to check each occurance " & RGSignature & "to determine whether it is the " & strA & " or Variable being referenced."

End Function

Public Sub DoSorting(cMod As CodeModule)

  'v3.0.9 simplified
  Dim I          As Long
  Dim Tmpstring1 As String
  Dim Cline      As String
  Dim MaxFactor  As Long

  'Moved this out of FormatCode while working out how it worked
  'didn't bother to put it back
  'order modules alphabetically
  'VER 1.1.46 MOVED TO EARLY IN CYCLE so that orphaned comments no longer occur
  With cMod
    If FixData(NSortModules).FixLevel Then
      'v2.8.6 protects PAul Caton's very useful sub-classing which requires the Sub named to be first in file)
      If Not SubClassingDetected(cMod) Then
        If .Members.Count > 1 Then
          'v3.6.9 hi Ulli this is the fix, so far
          CorrectIncompleteProperty cMod
          SortElems = SortingTagExtraction(cMod)
          QuickSort 1, UBound(SortElems), 0
          'build sorted component
          MaxFactor = UBound(SortElems)
          For I = 1 To MaxFactor
            MemberMessage "", I, MaxFactor
            Select Case I
             Case 1
              If SortElems(I)(1) > .CountOfDeclarationLines Then
                Tmpstring1 = Tmpstring1 & .Lines(SortElems(I)(1), SortElems(I)(2)) & vbNewLine
              End If
             Case Else
              'there's a quirk in VB: it returns Events as methods and if an
              'Event has the same name as a Sub/Function then this results in
              'duplicates, so here duplicates are filtered out
              If SortElems(I)(1) <> SortElems(I - 1)(1) Then
                If SortElems(I)(1) > .CountOfDeclarationLines Then
                  Tmpstring1 = Tmpstring1 & .Lines(SortElems(I)(1), SortElems(I)(2)) & vbNewLine
                End If
              End If
            End Select
          Next I
          ReWriter cMod, Split(Tmpstring1, vbNewLine), RWCode
          ' this line is because some UserControls generate duplicates of some Properties
          'This strips them out again
          ReWriter cMod, StripDuplicateArray(GetMembersArray(cMod)), RWMembers
          'remove trailing blank lines if any
          Do
            Cline = Trim$(.Lines(.CountOfLines, 1))
            If LenB(Cline) = 0 Then
              .DeleteLines .CountOfLines
            End If
          Loop Until Len(Cline)
          KillSelection
        End If
      End If
      DoEvents
    End If
  End With

End Sub

Public Sub ExcessiveScopeForVariables(ByVal ModuleNumber As Long, _
                                      dArray As Variant)

  Dim L_CodeLine As String
  Dim UpDated    As Boolean
  Dim TmpDecArr  As Variant
  Dim I          As Long
  Dim MaxFactor  As Long

  If dofix(ModuleNumber, UpdateDecTypeSuffix) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      For I = 1 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If Has_AS(L_CodeLine) Then
            If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Dim", "Global", "Public", "Private") Then
              If Not InstrAtPositionArray(L_CodeLine, ipAny, True, "Declare", "Type", "Enum") Then
                If ModuleLevelToDim(WordBefore(ExpandForDetection(L_CodeLine), " As"), L_CodeLine) Then
                  TmpDecArr(I) = Marker(TmpDecArr(I), SUGGESTION_MSG & "Variable could be changed to procedure level Dim", MAfter, UpDated)
                End If
              End If
            End If
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Sub Expand_SingleLine_Multiple_Declaration(cMod As CodeModule, _
                                                   dArray As Variant)

  Dim CommentStore As String
  Dim arrLine      As Variant
  Dim L_CodeLine   As String
  Dim CommaPos     As Long
  Dim UpDated      As Boolean
  Dim TmpDecArr    As Variant
  Dim I            As Long
  Dim LineUpdate   As Boolean
  Dim MaxFactor    As Long
  Dim ModuleNumber As Long

  ModuleNumber = ModDescMember(cMod.Parent.Name)
  'This routine places each Declaration variable on a separate line
  '(Except if line is in Enum Capitalisation Protection Structure)
  '*If there is an EOL comment then the comment is placed next to first member
  'NO coding value but improves readablity
  If dofix(ModuleNumber, ExpandDecSingleLine) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      For I = 1 To MaxFactor
        LineUpdate = False
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Dim", "Global", "Public", "Private", "Const") Then
            If Not InEnumCapProtection(cMod, TmpDecArr, I) Then
              If ExtractCode(L_CodeLine, CommentStore) Then
                arrLine = Split(L_CodeLine)
                CommaPos = 0
                L_CodeLine = ConcealParameterCommas(L_CodeLine)
                Do
                  CommaPos = GetCommaSpacePos(L_CodeLine, CommaPos + 1)
                  If CommaPos Then
                    If InCode(L_CodeLine, CommaPos) Then
                      If Not EnclosedInBrackets(L_CodeLine, CommaPos) Then
                        If arrLine(1) = "Const" Then
                          L_CodeLine = Left$(L_CodeLine, CommaPos - 1) & vbNewLine & _
                           arrLine(0) & " Const " & Mid$(L_CodeLine, CommaPos + 1)
                          LineUpdate = True
                         Else
                          L_CodeLine = Left$(L_CodeLine, CommaPos - 1) & vbNewLine & _
                           arrLine(0) & Mid$(L_CodeLine, CommaPos + 1)
                          LineUpdate = True
                        End If
                      End If
                    End If
                  End If
                Loop While CommaPos
              End If
            End If
            If LineUpdate Then
              'insert comment if any after first member of new layout
              If LenB(CommentStore) Then
                L_CodeLine = Safe_Replace(L_CodeLine, vbNewLine, CommentStore & vbNewLine, , 1)
              End If
              TmpDecArr(I) = L_CodeLine & vbNewLine & _
               RGSignature & "Multiple Declarations on one line expanded to one per line"
              AddNfix ExpandDecSingleLine
              UpDated = True
            End If
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Sub Expand_SingleLine_SingleType_Declaration(cMod As CodeModule, _
                                                     dArray As Variant)

  Dim ModuleNumber As Long
  Dim OrigLine     As String
  Dim L_CodeLine   As String
  Dim arrLine      As Variant
  Dim TmpDecArr    As Variant
  Dim UpDated      As Boolean
  Dim TypeDef      As String
  Dim I            As Long
  Dim J            As Long
  Dim RGsign       As String
  Dim FixMissing   As Boolean
  Dim LType        As String
  Dim LTypeDEf     As String
  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim MaxFactor    As Long

  'Convert 'Dim X ,Y ,Z As Type' to 'Dim X As Type, Y  As Type, Z As Type'
  On Error GoTo Opps
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, ExpandDecSingleLineSingleType) Then
    If Not DEFXXXUsed(cMod) Then
      TmpDecArr = CleanArray(dArray)
      MaxFactor = UBound(TmpDecArr)
      If MaxFactor > -1 Then
        For I = 0 To MaxFactor
          MemberMessage "", I, MaxFactor
          L_CodeLine = TmpDecArr(I)
          If Not JustACommentOrBlank(L_CodeLine) Then
            'detect firstline of code vs top of comments
            If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Dim", "Global", "Public", "Private", "Friend") Then
              If Not InEnumCapProtection(cMod, TmpDecArr, 0) Then
                If ExtractCode(L_CodeLine, CommentStore, SpaceOffSet) Then
                  L_CodeLine = ConcealParameterCommas(L_CodeLine, True)
                  OrigLine = L_CodeLine
                  If GetCommaSpacePos(L_CodeLine) > 0 Then
                    If CountSubString(L_CodeLine, " As ") = 1 Then
                      arrLine = Split(L_CodeLine, CommaSpace)
                      If CountSubString(arrLine(UBound(arrLine)), " As ") = 1 Then
                        TypeDef = GetType(arrLine(UBound(arrLine)))
                        For J = LBound(arrLine) To UBound(arrLine) - 1
                          If Get_As_Pos(arrLine(J)) = 0 Then
                            If Not TypeSuffixExists(arrLine(J)) Then
                              arrLine(J) = arrLine(J) & " As " & TypeDef
                            End If
                          End If
                        Next J
                        L_CodeLine = Join(arrLine, CommaSpace)
                        RGsign = vbNewLine & _
                         WARNING_MSG & "Declaring a whole line with single As Type is no longer supported." & IIf(Xcheck(XPrevCom), vbNewLine & _
                         PREVIOUSCODE_MSG & OrigLine, vbNullString)
                      End If
                     ElseIf CountSubString(L_CodeLine, " As ") > 1 Then
                      'This fixes any untyped variables if the rest of the variables are Typed the same
                      FixMissing = True
                      arrLine = Split(L_CodeLine, CommaSpace)
                      For J = LBound(arrLine) To UBound(arrLine)
                        If Get_As_Pos(arrLine(J)) > 0 Then
                          LType = GetType(arrLine(J))
                          If LenB(LTypeDEf) = 0 Then
                            LTypeDEf = LType
                          End If
                          If LTypeDEf <> LType Then
                            FixMissing = False
                            Exit For
                          End If
                        End If
                      Next J
                      If FixMissing Then
                        For J = LBound(arrLine) To UBound(arrLine)
                          If InStr(arrLine(J), LTypeDEf) = 0 Then
                            If Not TypeSuffixExists(arrLine(J)) Then
                              arrLine(J) = arrLine(J) & LTypeDEf
                            End If
                          End If
                        Next J
                        L_CodeLine = Join(arrLine, CommaSpace)
                        If SpaceOffSet & L_CodeLine <> OrigLine Then
                          RGsign = vbNewLine & _
                           WARNING_MSG & "Dimmed un-Typed variable in line where all others had same Type." & IIf(Xcheck(XPrevCom), vbNewLine & _
                           PREVIOUSCODE_MSG & OrigLine, vbNullString)
                        End If
                      End If
                    End If
                  End If
                End If
                If LenB(RGsign) Or OrigLine <> L_CodeLine Then
                  TmpDecArr(I) = L_CodeLine & CommentStore & RGsign
                  RGsign = vbNullString
                  AddNfix ExpandDecSingleLineSingleType
                  UpDated = True
                End If
              End If
            End If
          End If
        Next I
        dArray = CleanArray(TmpDecArr, UpDated)
      End If
    End If
  End If
  On Error GoTo 0

Exit Sub

Opps:
  mObjDoc.Safe_MsgBox Err.Description, vbCritical
  Resume Next

End Sub

Private Function GetDimTypeFromModule(ByVal ModuleNumber As Long, _
                                      ByVal strVariableName As String, _
                                      ByVal CompName As String, _
                                      Optional ByVal almodules As Boolean = True) As String

  Dim Proj         As VBProject
  Dim LModule      As CodeModule
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim TmpLinArr(2) As Variant
  Dim StartLine    As Long

  TmpLinArr(0) = vbNullString
  TmpLinArr(2) = vbNullString
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        If Not almodules Then
          If ModuleNumber <> CurCompCount Then
            GoTo skipmodule
          End If
        End If
        Set LModule = Comp.CodeModule
        StartLine = 1
        Do While LModule.Find(strVariableName, StartLine, 1, -1, -1, True, True)
          'Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strVariableName, L_CodeLine, StartLine)
          TmpLinArr(1) = LModule.Lines(StartLine, 1)
          GetDimTypeFromModule = GetDimTypeFromInternalEvidence(CurCompCount, TmpLinArr, strVariableName, CompName)
          If Len(GetDimTypeFromModule) Then
            Exit Do
          End If
          StartLine = StartLine + 1
          If StartLine >= LModule.CountOfLines Then
            Exit Do
          End If
        Loop
        If Len(GetDimTypeFromModule) Then
          Exit For
        End If
      End If
skipmodule:
    Next Comp
    If Len(GetDimTypeFromModule) Then
      Exit For
    End If
  Next Proj

End Function

Public Function GetFirstPublicProcedureName(cMod As CodeModule) As String

  'v3.0.8 Improvement to Subclass detection
  'Paul Caton's sbubclassing sub need not be first in code but must be first PUBLIC sub
  'Thanks Carles P.V. for making this mnecessary
  
  Dim I As Long

  With cMod
    For I = 1 To .Members.Count
      If .Members(I).Type = vbext_mt_Method Then
        If .Members(I).Scope = vbext_Public Then
          GetFirstPublicProcedureName = .Members(I).Name
          Exit For
        End If
      End If
    Next I
  End With

End Function

Public Function Has_AS(varCode As Variant) As Boolean

  Has_AS = InstrAtPosition(varCode, "As", ipAny)

End Function

Private Function HungarianDim(VarDimName As Variant, _
                              ByVal varCode As Variant, _
                              Optional ByVal TreatSingles As Boolean = False) As String

  Dim strType   As String
  Dim strPrefix As String
  Dim IDNoSType As Long

  'TreatSingles allows this routine to deal with single letter variables
  'by default it ignores them
  'However the declaration level test for poorlu named variables picks them up
  If Len(VarDimName) > 1 Or TreatSingles Then
    varCode = ExpandForDetection(varCode)
    strType = GetType(varCode)
    IDNoSType = QSortArrayPos(StandardTypes, strType)
    If IDNoSType > -1 Then
      strPrefix = StandardPreFix(IDNoSType) 'ArrayPos(strType, StandardTypes))
     ElseIf ArrayPos(strType, StandardControl) > -1 Then
      strPrefix = StandardCtrPrefix(ArrayPos(strType, StandardControl))
     Else
      'v3.0.5 improved naming conventions
      If WordBefore(varCode, VarDimName) = "Type" Then
        strPrefix = "typ"
       ElseIf WordBefore(varCode, VarDimName) = "Enum" Then
        strPrefix = "enm"
       Else
        strPrefix = FakeHungarian(strType)
      End If
    End If
    If Len(strPrefix) Then
      HungarianDim = strPrefix & Ucase1st(VarDimName)
    End If
  End If

End Function

Private Sub InsertLargeFileWarning(ByVal ModuleNumber As Long, _
                                   dArray As Variant, _
                                   ByVal FName As String)

  Dim TmpDecArr As Variant
  Dim strMsg    As String

  'ver 1.1.26 rblanch pointed out a code that had this problem
  If dofix(ModuleNumber, LargeSourceWarning_CXX) Then
    TmpDecArr = dArray
    If UBound(TmpDecArr) > 0 Then
      strMsg = SUGGESTION_MSG & strInSQuotes(FileNameOnly(FName)) & " Size " & ModuleSize(FName) & vbNewLine & _
       RGSignature & "exceeds VB's recommended maximum size of 64KB for a single source file."
      If MultiRight(FName, False, ".frm", ".ctl", ".dob") Then
        strMsg = strMsg & vbNewLine & _
         RGSignature & "For control bearing modules this limit normally applys only to the code value."
      End If
      TmpDecArr(0) = Marker(TmpDecArr(0), strMsg, MBefore)
      dArray = CleanArray(TmpDecArr)
    End If
  End If

End Sub

Public Function isPaulCatonSubClassing(cMod As CodeModule) As Boolean

  ',v2.9.7 improved to avoid self detection
  
  Dim lngJunk As Long

  With cMod
    If .Find(StrReverse("08EE0BE8F5498CF54980C13758F4C385E9855"), lngJunk, lngJunk, -1, -1) Then
      'StrReverse stops self detection
      'use Find because FindCodeUsage doesn't see strings
      'this is part of the Machine code that drives the technique,
      'looking for it rather than the renameable sub is probably safer
      isPaulCatonSubClassing = True
     ElseIf FindCodeUsage("Sub zSubclass_Proc", "", .Parent.Name, , True) Then
      'this is the default name just in case someone hides the Machine code (Encrypted, res file etc)
      'BUT doesn't rename the sub
      isPaulCatonSubClassing = True
    End If
  End With

End Function

Private Sub LongLineFix(ByVal ModuleNumber As Long, _
                        dArray As Variant)

  Dim L_Updated  As Boolean
  Dim J          As Long
  Dim TmpDecArr  As Variant
  Dim L_CodeLine As String
  Dim MaxFactor  As Long
  Dim UpDated    As Boolean

  'very long lines cannot be written properly to IDE pages so this recreates line seperators if necessary
  'Break up compound lines
  If Not ModDesc(ModuleNumber).MDDontTouch Then
    TmpDecArr = dArray
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      'For I = 1 To MaxFactor
      For J = LBound(TmpDecArr) To UBound(TmpDecArr)
        MemberMessage "", J, MaxFactor
        If LenB(TmpDecArr(J)) > 1023 Then
          L_CodeLine = TmpDecArr(J)
          If Not JustACommentOrBlank(L_CodeLine) Then
            LineContinuationForVeryLongLines L_CodeLine, ContMark & vbNewLine, L_Updated
            If L_Updated Then
              TmpDecArr(J) = L_CodeLine
              UpDated = True
              L_Updated = False
            End If
          End If
        End If
      Next J
      'Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Public Function ModuleSize(ByVal strF As String) As String

  Dim dblFilsize As Double
  Dim dblCntrl   As Double

  dblFilsize = FileSize(strF)
  If MultiRight(strF, False, ".frm", ".ctl", ".dob") Then
    dblCntrl = Round(AutoCodeSize(strF) / 1024, 1)
    ModuleSize = "(Code:" & Round(dblFilsize - dblCntrl, 1) & "KB; Control Description:" & Round(dblCntrl, 1) & "KB) "
   Else
    ModuleSize = "(Code:" & Round(dblFilsize, 1) & "KB) "
  End If

End Function

Private Sub Move_API_Declares_To_End(ByVal ModuleNumber As Long, _
                                     dArray As Variant)

  Dim StrDeclareCollector As String
  Dim L_CodeLine          As String
  Dim UpDated             As Boolean
  Dim TmpDecArr           As Variant
  Dim I                   As Long
  Dim InHashIf            As Boolean
  Dim MaxFactor           As Long
  Dim EndDecs             As Long

  'Move API calls to bottom of Declarations
  'Safety formatting; makes sure that any required Types are declared first.
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  TmpDecArr = CleanArray(dArray)
  MaxFactor = UBound(TmpDecArr)
  If dofix(ModuleNumber, MoveAPIDown) Then
    If MaxFactor > -1 Then
      EndDecs = MaxFactor
      'avoid moving Delcates which are already at end of section
      Do While InstrAtPosition(TmpDecArr(EndDecs), "Declare", ipLeftOr2nd)
        EndDecs = EndDecs - 1
      Loop
      If EndDecs Then
        For I = 0 To EndDecs
          MemberMessage "", I, EndDecs
          L_CodeLine = TmpDecArr(I)
          If Not JustACommentOrBlank(L_CodeLine) Then
            If Not InHashIf Then
              '*if False the test to set it
              InHashIf = Left$(L_CodeLine, 4) = "#If "
             Else
              '*if True then leave True until it is turned off
              If Left$(L_CodeLine, 7) = "#End If" Then
                InHashIf = False
              End If
            End If
            If Not InHashIf Then
              If I < EndDecs Then
                If InstrAtPosition(TmpDecArr(I), "Declare", ipLeftOr2nd) Then
                  StrDeclareCollector = StrDeclareCollector & vbNewLine & TmpDecArr(I)
                  UpDated = True
                  AddNfix MoveAPIDown
                  TmpDecArr(I) = vbNullString
                End If
              End If
            End If
          End If
        Next I
        If UpDated Then
          TmpDecArr(UBound(TmpDecArr)) = TmpDecArr(UBound(TmpDecArr)) & StrDeclareCollector
          dArray = CleanArray(TmpDecArr)
        End If
      End If
    End If
  End If

End Sub

Private Function NotUsedOutSideMouseEvents(ByVal ModuleNumber As Long, _
                                           ByVal strCode As String, _
                                           strTest As String) As Boolean

  Dim bPrivOnly    As Boolean
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long
  Dim CompMod      As CodeModule
  Dim Sline        As Long

  NotUsedOutSideMouseEvents = True
  bPrivOnly = InStr(strCode, "Private")
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        If ModuleNumber <> CurCompCount Then
          If bPrivOnly Then
            GoTo Skip
          End If
        End If
        ModuleMessage Comp, CurCompCount
        Set CompMod = Comp.CodeModule
        With CompMod
          Sline = 1
          Do While .Find(strTest, Sline, 1, -1, -1, True, False, False)
            'if exits at all, then look for the line(s)
            If IsRealWord(.Lines(Sline, 1), strTest) Then
              If .Lines(Sline, 1) <> strCode Then
                If InStr(GetProcName(CompMod, Sline), "_Mouse") = 0 Then
                  NotUsedOutSideMouseEvents = False
                  Exit Do
                End If
              End If
            End If
            Sline = Sline + 1
          Loop
        End With
      End If
Skip:
      If Not NotUsedOutSideMouseEvents Then
        Exit For
      End If
    Next Comp
    If Not NotUsedOutSideMouseEvents Then
      Exit For
    End If
  Next Proj

End Function

Public Sub ProtectEnumCap(dArray As Variant)

  Dim I As Long

  If Not IsEmpty(dArray) Then
    If UBound(dArray) > -1 Then
      'ver1.1.30 update
      'copes with last element(s) being blank
      For I = UBound(dArray) To 1 Step -1
        If Len(Trim$(dArray(I))) Then
          'count back to first non-blank
          If InStr(SQuote, Left$(dArray(I), 1)) Then
            '*if it is a comment (inserted by Code Fixer) or a compilation directive (#End If)
            'Insert a temporary comment to keep it attached to Declaration
            'and keep it from being moved into the first procedure( in Sort procedures fix)
            dArray(I) = dArray(I) & vbNewLine & CodeFixProtectedArray(endDec) & vbNewLine
          End If
          Exit For
        End If
      Next I
    End If
  End If

End Sub

Private Sub ReplaceRem()

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim strKill      As String
  Dim StartLine    As Long

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        ModuleMessage Comp, CurCompCount
        StartLine = 1
        With Comp
          Do While .CodeModule.Find("Rem ", StartLine, 1, -1, -1, False, False, True)
            MemberMessage "", StartLine, .CodeModule.CountOfLines
            strKill = Trim$(.CodeModule.Lines(StartLine, 1))
            If Left$(strKill, 4) = "Rem " Then
              .CodeModule.ReplaceLine StartLine, SQuote & Mid$(strKill, 5)
            End If
            StartLine = StartLine + 1
            If StartLine > .CodeModule.CountOfLines Then
              Exit Do
            End If
          Loop
        End With 'Comp
      End If
    Next Comp
  Next Proj
  On Error GoTo 0

End Sub

Public Sub SaveMemberAttributes(ByVal ModuleNumber As Long, _
                                membs As Members)

  Dim Member             As Member
  Dim I                  As Long
  Dim MemberAttributes() As Variant   'vbArray of 13 values

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  I = 0
  ReDim MemberAttributes(0 To membs.Count) As Variant
  On Error Resume Next
  For Each Member In membs
    I = I + 1
    ' Something strange happens here:
    'Getting the Description Attribute sometimes fails
    'after the source under examination has run in the IDE.
    'It never fails when the source has not been running yet.
    'Maybe it's a VB-Bug(?)
    Err.Clear
    With Member
      MemberAttributes(I) = Array(.Name, .Bindable, .Browsable, .Category, .DefaultBind, .Description, .DisplayBind, .HelpContextID, _
                                  .Hidden, .PropertyPage, .RequestEdit, .StandardMethod, .UIDefault)
      If Err.Number Then
        I = I - 1
      End If
    End With
  Next Member
  On Error GoTo 0
  ReDim Preserve MemberAttributes(0 To I)
  Attributes(ModuleNumber) = MemberAttributes

End Sub

Public Sub SeparateCompoundDeclarationLines(ByVal ModuleNumber As Long, _
                                            dArray As Variant, _
                                            ByVal CompName As String)

  Dim I          As Long
  Dim J          As Long
  Dim TmpDecArr  As Variant
  Dim arrLine    As Variant
  Dim L_CodeLine As String
  Dim MaxFactor  As Long
  Dim UpDated    As Boolean

  'This deals with the legal but extremely rare format of Enum and Type using Colon seperation
  'Thanks :)  to CentralWare OSD Control for using that
  'Break up compound lines
  If dofix(ModuleNumber, SeperateCompounds) Then
    TmpDecArr = dArray
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      For I = 1 To MaxFactor
        MemberMessage "", I, MaxFactor
        arrLine = Split(TmpDecArr(I), vbNewLine)
        For J = LBound(arrLine) To UBound(arrLine)
          L_CodeLine = arrLine(J)
          If Not JustACommentOrBlank(L_CodeLine) Then
            DoSeparateCompoundLines ModuleNumber, L_CodeLine, CompName
            If arrLine(J) <> L_CodeLine Then
              AddNfix SeperateCompounds
              arrLine(J) = L_CodeLine
              UpDated = True
            End If
          End If
        Next J
        TmpDecArr(I) = Join(arrLine, vbNewLine)
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Sub SubClassDetectorInitialise(cMod As CodeModule, _
                                       Optional dArray As Variant)

  If SubClassingDetected(cMod) Then
    If IsArray(dArray) Then
      If isPaulCatonSubClassing(cMod) Then
        bPaulCatonSubClasUsed = True
        strPaulCatonSubClasProcName = GetFirstPublicProcedureName(cMod)
        strPaulCatonSubClasCompName = cMod.Parent.Name
        dArray(0) = WARNING_MSG & "SUB-CLASSING DETECTED" & vbNewLine & _
         WARNING_MSG & "Code Fixer may move the first procedure ('" & strPaulCatonSubClasProcName & "') in this module." & vbNewLine & _
         WARNING_MSG & "Code Fixer should reposition it in the correct place but please check before running the code." & vbNewLine & _
         dArray(0)
        'v2.9.5 new variable used to stop unused commenting the proc out
       Else
        dArray(0) = WARNING_MSG & "SUB-CLASSING DETECTED" & vbNewLine & _
         WARNING_MSG & "Code Fixer may move a procedure which this method requires at a specific place in the code." & vbNewLine & _
         dArray(0)
      End If
      '    ReWriter cMod, ArrCode, RWModule
      dArray = CleanArray(dArray)
    End If
  End If

End Sub

Public Function SubClassingDetected(cMod As CodeModule) As Boolean

  ',v2.9.7 improved to avoid self detection
  
  Dim lngJunk As Long

  With cMod
    'The emptystring is to stop self detection of this procedure
    If .Find(StrReverse("85E9855"), lngJunk, lngJunk, -1, -1) Then
      'StrReverse stops self detection
      'use Find because FindCodeUsage doesn't see strings
      'this is part of the Machine code that drives at least 2 techniques of subclassing,
      'looking for it rather than the renameable sub is probably safer
      SubClassingDetected = True
     ElseIf FindCodeUsage("zSubclass_Proc", "", .Parent.Name, , True) Then
      'only if Paul Caton subclassing is used
      'this is the default name just in case someone hides the Machine code (Encrypted, res file etc)
      'BUT doesn't rename the sub
      SubClassingDetected = True
    End If
  End With

End Function

Public Function SuggestNewName(varOld As Variant, _
                               varCode As Variant, _
                               Optional TreatSingles As Boolean = False) As String

  Dim strTmp As String

  strTmp = HungarianDim(varOld, varCode, TreatSingles)
  If LenB(strTmp) Then
    SuggestNewName = SUGGESTION_MSG & "Change the variable name to (" & strTmp & ")."
  End If

End Function

Private Function TestNextLineOfDeclaration(ByVal LinNo As Long, _
                                           ByVal TmpA As Variant, _
                                           ByVal strTest As String, _
                                           lngRealCodeline As Long) As Boolean

  If LinNo = UBound(TmpA) Then
    TestNextLineOfDeclaration = False
   ElseIf InstrAtPosition(NextCodeLine(TmpA, LinNo, , lngRealCodeline), strTest, IpLeft) Then
    TestNextLineOfDeclaration = True
  End If

End Function

Private Sub TypeCastUnTypedDeclarations(ByVal ModuleNumber As Long, _
                                        dArray As Variant, _
                                        Comp As VBComponent)

  Dim TargetMember As Long
  Dim OrigLine     As String
  Dim L_CodeLine   As String
  Dim CommentStore As String
  Dim TmpDecArr    As Variant
  Dim UpDated      As Boolean
  Dim I            As Long
  Dim MaxFactor    As Long
  Dim arrLine      As Variant
  Dim AsType       As String

  'Use Array Generated in DeclarationDefTypeDetector
  'to set each Declaration to 'As Type style
  If Not ModDesc(ModuleNumber).MDDontTouch Then
    DeclarationDefTypeDetector ModuleNumber, dArray, False
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > -1 Then
      For I = 0 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Public", "Private", "Static") Then
            If Not InstrAtPositionArray(L_CodeLine, ip2nd, True, "Enum", "Type", "Declare", "Event") Then
              If Not InEnumCapProtection(Comp.CodeModule, dArray, I) Then
                OrigLine = L_CodeLine
                ExtractCode L_CodeLine, CommentStore
                If Not Has_AS(L_CodeLine) Then
                  L_CodeLine = ExpandForDetection(L_CodeLine)
                  L_CodeLine = ConcealParameterCommas(L_CodeLine, True)
                  OrigLine = L_CodeLine
                  arrLine = Split(L_CodeLine)
                  If arrLine(1) = "Const" Then
                    TargetMember = 2
                   Else
                    TargetMember = 1
                  End If
RetryPublic2PrivateChange:
                  AsType = GetDimTypeFromInternalEvidence(ModuleNumber, TmpDecArr, CStr(arrLine(TargetMember)), Comp.Name)
                  If LenB(AsType) Then
                    If Left$(AsType, 2) = "<-" Then
                      L_CodeLine = Left$(L_CodeLine, Len(L_CodeLine) - 1)
                      AsType = Mid$(AsType, 4)
                    End If
                    Safe_AsTypeAdd L_CodeLine, AsType
                    TmpDecArr(I) = L_CodeLine & CommentStore
                    TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Untyped Variable " & IIf(Len(AsType), IIf(UsingDefTypes, vbNullString, ". Auto-Type may not be correct"), ". Type could not be determined"), MAfter)
                    UpDated = True
                   Else
                    If arrLine(0) = "Private" Then
                      If Not FindCodeUsage(CStr(arrLine(1)), OrigLine, Comp.Name, False, True, False) Then
                        TmpDecArr(I) = SQuote & TmpDecArr(I)
                        TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Unused & Untyped Variable is not used in code", MAfter)
                        UpDated = True
                       Else '1.0.87 bug fix
                        AsType = GetDimTypeFromModule(ModuleNumber, CStr(arrLine(TargetMember)), False)
                        If LenB(AsType) Then
                          Safe_AsTypeAdd L_CodeLine, AsType
                          TmpDecArr(I) = L_CodeLine & CommentStore
                          TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Untyped Variable " & IIf(Len(AsType), IIf(UsingDefTypes, vbNullString, ". Auto-Type may not be correct"), ". Type could not be determined"), MAfter)
                         Else
                          TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Untyped Variable will behave as Variant", MAfter)
                        End If
                        UpDated = True
                      End If
                     ElseIf arrLine(0) = "Public" Then
                      If Not FindCodeUsage(CStr(arrLine(TargetMember)), OrigLine, Comp.Name, , False, False) Then
                        If Not FindCodeUsage(CStr(arrLine(TargetMember)), OrigLine, Comp.Name, , True, False) Then
                          arrLine(0) = "Private"
                          UpDated = True
                          GoTo RetryPublic2PrivateChange
                         Else
                          TmpDecArr(I) = SQuote & TmpDecArr(I)
                          TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Unused Variable is not used in code", MAfter)
                          UpDated = True
                        End If
                       Else ' 1.1.70 bug fix
                        AsType = GetDimTypeFromModule(ModuleNumber, CStr(arrLine(TargetMember)), Comp.Name)
                        If LenB(AsType) Then
                          Safe_AsTypeAdd L_CodeLine, AsType
                          TmpDecArr(I) = L_CodeLine & CommentStore
                          TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Untyped Variable " & IIf(Len(AsType), IIf(UsingDefTypes, vbNullString, ". Auto-Type may not be correct"), ". Type could not be determined"), MAfter)
                         Else
                          TmpDecArr(I) = SmartMarker(TmpDecArr, I, WARNING_MSG & "Untyped Variable will behave as Variant", MAfter)
                        End If
                        UpDated = True
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Private Sub Update_DefType_to_AsType_Declaration(ByVal ModuleNumber As Long, _
                                                 dArray As Variant)

  Dim OrigLine     As String
  Dim L_CodeLine   As String
  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim TmpDecArr    As Variant
  Dim UpDated      As Boolean
  Dim I            As Long
  Dim MaxFactor    As Long

  'Use Array Generated in DeclarationDefTypeDetector
  'to set each Declaration to 'As Type style
  If dofix(ModuleNumber, UpdateDefType2AsType) Then
    DeclarationDefTypeDetector ModuleNumber, dArray, False
    If UsingDefTypes Then
      TmpDecArr = CleanArray(dArray)
      MaxFactor = UBound(TmpDecArr)
      If MaxFactor > -1 Then
        For I = 0 To MaxFactor
          MemberMessage "", I, MaxFactor
          L_CodeLine = TmpDecArr(I)
          OrigLine = L_CodeLine
          If ExtractCode(L_CodeLine, CommentStore, SpaceOffSet) Then
            If InstrAtPositionArray(L_CodeLine, IpLeft, True, "Dim", "Public", "Private", "Static") Then
              If Not InstrAtPositionArray(L_CodeLine, ip2nd, True, "Enum", "Type", "Declare") Then
                If Not Has_AS(L_CodeLine) Then
                  L_CodeLine = ConcealParameterCommas(L_CodeLine, True)
                  OrigLine = L_CodeLine
                  If Get_As_Pos(L_CodeLine) = 0 Then
                    If UsingDefTypes Then
                      L_CodeLine = L_CodeLine & FromDefType(Split(L_CodeLine)(1), ModuleNumber)
                    End If
                    If OrigLine <> L_CodeLine Then
                      L_CodeLine = SpaceOffSet & L_CodeLine & CommentStore
                      TmpDecArr(I) = L_CodeLine
                      AddNfix UpdateDefType2AsType
                      UpDated = True
                    End If
                  End If
                End If
              End If
            End If
          End If
        Next I
        dArray = CleanArray(TmpDecArr, UpDated)
      End If
    End If
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:14:13 AM) 16 + 1790 = 1806 Lines Thanks Ulli for inspiration and lots of code.

