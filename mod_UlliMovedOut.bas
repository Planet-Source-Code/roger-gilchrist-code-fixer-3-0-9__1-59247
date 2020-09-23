Attribute VB_Name = "mod_UlliMovedOut"

Option Explicit
Public bNoOptExp                  As Boolean
'v3.0.0 disable Option Explict insertion for format only thanks Ian K
Public bAbortComplete             As Boolean
Public bFixing                    As Boolean
'© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
'with extensive modifications
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Private Const IntegerLimit        As Long = 32767
Private StructuralERRORFOUND      As Boolean
Private FullTabWidth              As Long
Private BreakLoop                 As Boolean
Public Enum MemAttrPtrs
  MemName = 0
  MemBind = 1
  MemBrws = 2
  MemCate = 3
  MemDfbd = 4
  MemDesc = 5
  MemDbnd = 6
  MemHelp = 7
  MemHidd = 8
  MemProp = 9
  MemRqed = 10
  MemStme = 11
  MemUide = 12
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private MemName, MemBind, MemBrws, MemCate, MemDfbd, MemDesc, MemDbnd, MemHelp, MemHidd, MemProp, MemRqed, MemStme, MemUide
#End If
Private Words()                   As String
Private ErrLineFrom               As Long
Private ErrLineTo                 As Long
Private curTopLine                As Long
Private Type SelectionCoords
  StartLine                       As Long
  EndLine                         As Long
  startCol                        As Long
  endCol                          As Long
End Type
'Constants
Public Const REMBLANK             As String = "'REM BLANK"
Public Const Smiley               As String = "':)"
Public Const Grumpy               As String = "':("
Private Const StackUnderflow      As String = RGSignature & "[None]"
Private Const Quote               As String = """"
Public Const Colon                As String = ":"
Private Const AlmostInfinity      As Long = 999999
Public Attributes()               As Variant
Public IndentSize                 As Long
Private Sel                       As SelectionCoords
Private CodeIsDead                As Long
Private LastWordIndex             As Long
Private VarNames                  As Collection
Private wordstack()               As String
Private StrucStack()              As String
'Options and other stuff
Private DefConstType              As Boolean
Private DuplNameDetected          As Boolean
Private TypeSuffixFound           As Boolean
'Program States
'Complaints
Private SlIf                      As Boolean
'Assortment of Strings
Private HalfIndent                As String
Private DefTypeChars              As String
Private complain                  As Boolean
Private ToDo1                     As String
Private ToDo2                     As String
Private ToDo2a                    As String
Private ToDo3                     As String
Private ToDo4                     As String
Private ToDo5                     As String
Private ToDo6                     As String
Private ToDo7                     As String
Private ToDo8                     As String
Private ToDo9                     As String
Private ToDo10                    As String
Private ToDo11                    As String
Private ToDo12                    As String
Private ToDo14                    As String
Private ToDo15                    As String

Private Sub AbortFormatting()

  Dim CurCompCount As Long
  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim cMod         As CodeModule

  ' allows last one to exit outside the For Loops and be used in report
  With frm_CodeFixer
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          ModuleMessage Comp, CurCompCount
          Set cMod = Comp.CodeModule
          MoveUnusedRoutines cMod
          CleanDummyLine cMod
          ReDoLineContinuation cMod, True
          RestoreMemberAttributes CurCompCount, cMod.Members
        End If
        If BreakLoop Then
          Exit For
        End If
      Next Comp
      If BreakLoop Then
        Exit For
      End If
    Next Proj
  End With

End Sub

Private Sub AutoFix(cMod As CodeModule, _
                    ByVal LIndex As Long, _
                    MarkText As String, _
                    ByVal indent As Long, _
                    Optional ResetForAddedLine As Boolean)

  Dim ReWrite      As String
  Dim strTarget    As String
  Dim ModuleNumber As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'switchboard program for various ways of rewriting codelines
  'This makes it easy to develop new sub routines to fix other Detectable problems
  'if AutoFix can't do anything then Ulli's original comment (MarkText )
  'Any routine that inserts vbNewLine needs to set the ResetForAddedLine to true if it hits
  'This also allows lines with multiple Ulli comments to get hit multiple times
  'RGSignature = StrReverse(")-:<'")
  With cMod
    ModuleNumber = ModDescMember(.Parent.Name)
    strTarget = .Lines(LIndex, 1)
    'because some MarkTexts ("Move line to top of current ","Repeat For-Variable:" and "Expand Structure")
    ' have variable members Instr rather than Select Case must be used
    If InStr(MarkText, "Move line to top of current ") Then
      Select Case FixData(MoveDim2Top).FixLevel
       Case 0
        MarkText = vbNullString
       Case 1
        ReWrite = strTarget
        '       Case 2, 3
      End Select
      'does not use ReWrite and writing section so needs to test here
     ElseIf InStr(MarkText, "Expand Structure") Then
      Select Case FixData(NIfThenExpand).FixLevel
       Case 0
        MarkText = vbNullString
       Case 1
        ReWrite = strTarget
       Case 2, 3
        If ReWrite <> strTarget Then
          AddNfix NIfThenExpand
        End If
        ResetForAddedLine = True
      End Select
     ElseIf InStr(MarkText, "Repeat For-Variable:") Then
      Select Case FixData(NForNextVar).FixLevel
       Case 0
        MarkText = vbNullString
       Case 1
        ReWrite = strTarget
       Case 2, 3, 4
        ReWrite = ForVariableInsert(strTarget, MarkText)
      End Select
     ElseIf InStr(MarkText, "On Error Resume still active") And dofix(ModuleNumber, NCloseResume) Then
      ReWrite = ErrorResumeCloser(strTarget)
     ElseIf InStr(MarkText, "Remove Pleonasm") And dofix(ModuleNumber, NPleonasmFix) Then
      ReWrite = PleonasmCleaner(strTarget, InStr(strTarget, "= True") > 0)
     ElseIf InStr(MarkText, ": UnTyped Variable. Will behave as Variant") Then
      If dofix(ModuleNumber, UpdateDimType) Then
        ReWrite = AsVariantFix(cMod, LIndex, strTarget)
       Else
        'This is a specialised test to stop Ulli's code detecting Enum Capitalisation Protector
        If DuplicateEnumCapitalizationTest(cMod, LIndex) Then
          .ReplaceLine LIndex, strTarget
        End If
      End If
     ElseIf InStr(MarkText, UPDATED_MSG & "Obsolete Type Suffix") And dofix(ModuleNumber, UpdateDimType) Then
      ReWrite = TypeSuffixExtender(strTarget)
     ElseIf InStr(MarkText, "Missing Scope") Then
      ReWrite = ScopeTo(cMod, strTarget, indent)
     ElseIf InStr(MarkText, "Duplicated Name") Then
      ReWrite = DuplicateEnumCapitalization(cMod, LIndex, strTarget)
     ElseIf InStr(MarkText, "Structure Error") Then
      'Stop Formatting at structural errors
      '
      mObjDoc.Safe_MsgBox " A Structural Error has stopped Code Fixer " & vbNewLine & _
                    "Search the Routine above the comment for Layout errors." & vbNewLine & _
                    "Error caused by an incorrect AutoFix OR errors in original code.", vbCritical
      'then stop prosess as soon as possible
      BreakLoop = True
      StructuralERRORFOUND = True
      ReWrite = Marker(strTarget & MarkText, RGSignature & "Search the Routine above this comment for Layout errors." & vbNewLine & _
       "Check Code logic. Error usually caused by an incorrect AutoFix OR errors in original code.", MAfter)
     Else
      '*If Fix_XXX was not True then Ulli's original message is added here
      'Also the following are not used or treated by AutoFix so always reach here
      '"Modifies active For-Variable"
      '"Remove Line Number"
      '"Possible Structure Violation"
      ' "Dead Code"
      'So just add Ulli's comment and fall past the Len(ReWrite) test
      ReWrite = strTarget
    End If
  End With
  If LenB(ReWrite) Then
    AutoFixWriter cMod, MarkText, ReWrite, strTarget, LIndex, ResetForAddedLine
  End If

End Sub

Private Function CheckVardefs(ByVal WordIndex As Long, _
                              ByVal LocalDef As Boolean, _
                              ByVal IsConst As Boolean, _
                              ByVal bIsFunction As Boolean) As Boolean

  Dim I          As Long
  Dim Tmpstring1 As String

  'Returns True if line containes untyped variable
  'May also set the DuplNameDetected, TypeSuffixFound, and/or DefConstType flags
  DuplNameDetected = False
  DefConstType = False
  TypeSuffixFound = False
  Do
    If LenB(Words(WordIndex)) Then
      Tmpstring1 = Replace$(Words(WordIndex), CommaSpace, vbNullString)
      If Not bIsFunction Then
        I = GetLeftBracketPos(Tmpstring1)
        If I Then
          Tmpstring1 = Left$(Tmpstring1, I - 1)
          Do Until GetRightBracketPos(Words(WordIndex))
            WordIndex = WordIndex + 1
            If WordIndex > UBound(Words) Then
              WordIndex = UBound(Words)
              Exit Do
            End If
          Loop
        End If
       Else
        'roger G added this; allows function parameters using type suffixes to be fixed
        If TypeIsUndefined(Tmpstring1, bIsFunction, True) Then
          TypeSuffixFound = False
        End If
      End If
      'TmpString1 has variable/const name
      If TypeIsUndefined(Tmpstring1, IsConst) Then
        If IsConst Then
          If WordIndex = LastWordIndex Then
            DefConstType = True
           Else
            If Words(WordIndex + 1) = "=" Then
              If Val(Words(WordIndex + 2)) >= -IntegerLimit - 1 Then
                If Val(Words(WordIndex + 2)) <= IntegerLimit Then
                  If Val(Words(WordIndex + 2)) = Int(Val(Words(WordIndex + 2))) Then
                    DefConstType = Not (TypeSuffixExists(Words(WordIndex + 2)) Or SmartLeft(Words(WordIndex + 2), Quote) Or (SmartLeft(Words(WordIndex + 2), "&") And Len(Words(WordIndex + 2)) > 6))
                  End If
                End If
              End If
            End If
          End If
         Else
          If WordIndex >= LastWordIndex Or Right$(Words(WordIndex), 1) = "," Then
            'no DefType and no type suffix and no 'As'
            CheckVardefs = True
          End If
        End If
      End If
      If TypeSuffixExists(Tmpstring1) Then
        Tmpstring1 = Left$(Tmpstring1, Len(Tmpstring1) - 1)
        TypeSuffixFound = True
      End If
      With VarNames
        On Error Resume Next
        If LocalDef Then
          Tmpstring1 = .Item(Tmpstring1)
          DuplNameDetected = DuplNameDetected Or (Err.Number = 0)
         Else
          .Add True, Tmpstring1
          DuplNameDetected = DuplNameDetected Or (Err.Number <> 0)
        End If
        On Error GoTo 0
      End With
      Do Until Right$(Words(WordIndex), 1) = ","
        If WordIndex = LastWordIndex Then
          Exit Do
        End If
        WordIndex = WordIndex + 1
      Loop
    End If
    WordIndex = WordIndex + 1
  Loop Until WordIndex > LastWordIndex Or bIsFunction

End Function

Private Sub CleanDummyLine(cMod As CodeModule)

  Dim DecArray As Variant
  Dim I        As Long

  DecArray = GetDeclarationArray(cMod)
  For I = LBound(DecArray) To UBound(DecArray)
    If MultiLeft(Trim$(DecArray(I)), True, Left$(CodeFixProtectedArray(endDec), 31)) Then
      DecArray(I) = vbNullString
    End If
  Next I
  ReWriter cMod, DecArray, RWDeclaration

End Sub

Private Function CleanUpCode(cMod As CodeModule) As Boolean

  Dim Cline       As String
  Const cElipsis  As String = "…"
  Dim SrcFileName As String

  'Moved this out of FormatCode while working out how it worked
  'didn't bother to put it back
  'converted to Function so that it can abort FormatCode if form/module contains no code
  CleanUpCode = True
  With cMod
    SrcFileName = .Parent.FileNames(1)
    If LenB(SrcFileName) > 40 Then
      SrcFileName = Left$(SrcFileName, 20) & cElipsis & LTrim$(Right$(SrcFileName, 20))
    End If
    SrcFileName = SrcFileName & IIf(.Parent.IsDirty, "(unsaved)", vbNullString)
    If LenB(Trim$(.Lines(.CountOfDeclarationLines + 1, 1))) Then
      If Trim$(.Lines(.CountOfDeclarationLines + 1, 1)) <> REMBLANK Then
        'insert a blank line between declarations and code
        .InsertLines .CountOfDeclarationLines + 1, ""
        KillSelection
      End If
    End If
    'remove bottom empty lines
    Do Until .CountOfLines = 0
      Cline = Trim$(.Lines(.CountOfLines, 1))
      If LenB(Cline) = 0 Or Left$(Cline & SngSpace, Len(Smiley)) = Smiley Then
        .DeleteLines .CountOfLines
       Else
        Exit Do
      End If
    Loop
    If InsufficentCode(cMod) Then
      NotEnoughCode cMod
      CleanUpCode = False
    End If
  End With

End Function

Private Sub ClearOldData()

  ReDim DeclarDesc(0) As DeclarDescriptor
  bDeclExists = False
  ReDim PRocDesc(0) As ProcedureDescriptor
  bProcDescExists = False
  ReDim CntrlDesc(0) As ControlDescriptor
  bCtrlDescExists = False
  ReDim EventDesc(0) As EventDescriptor
  bEventDescExists = False
  bPaulCatonSubClasUsed = False

End Sub

Private Sub ConcatentForFormatCode(cMod As CodeModule, _
                                   ByRef Cline As String, _
                                   ByVal LLine As Long)

  'Moved out of FormatCode to make it more readable
  'RG This needs to be here even if my ConcatentLineContinuation is on
  'because very long lines might fall through it

  With cMod
    Do
      'concat continuation lines
      Cline = Cline & LCase$(Trim$(.Lines(LLine, 1)))
      If HasLineCont(Cline) Then
        Cline = Left$(Cline, Len(Cline) - 1)
        LLine = LLine + 1
       Else
        Exit Do
      End If
    Loop
  End With

End Sub

Private Sub ConcatentRangeForFormatCode(Cline As String)

  Dim I As Long

  'Moved out of FormatCode to make it more readable
  'FromLine and ToLine define the concatted range
  For I = 1 To Len(Cline)
    Select Case Mid$(Cline, I, 1)
     Case SQuote
      If I > 1 Then
        Cline = Left$(Cline, I - 1)
      End If
      Exit For
     Case "r"
      If I > 1 Then
        If Mid$(Cline, I - 1, 1) = SngSpace Then
          If Mid$(Cline, I + 1, 3) = "em " Then
            Cline = Left$(Cline, I - 1)
            Exit For
          End If
        End If
       Else
        If Mid$(Cline, I + 1, 3) = "em " Then
          Exit For
        End If
      End If
     Case Quote
      Do
        I = I + 1
        Select Case Mid$(Cline, I, 1)
         Case Quote
          Exit Do
         Case SngSpace
          Mid$(Cline, I, 1) = vbNullChar
        End Select
      Loop While I <= Len(Cline)
      'v2.3.5 special case if len line is exactly the max length thie causes a crash
     Case "["
      Do
        I = I + 1
        Select Case Mid$(Cline, I, 1)
         Case "]"
          Exit Do
         Case SngSpace
          Mid$(Cline, I, 1) = vbNullChar
        End Select
      Loop While I <= Len(Cline)
      'v2.3.5 special case if len line is exactly the max length thie causes a crash
    End Select
  Next I

End Sub

Public Sub D0_Abort()

  Dim MyhourGlass As cls_HourGlass

  DoEvents
  If Not bAddinTerminate Then
    If MsgBox("Abort Code Fixer?" & vbNewLine & _
          "Selecting 'Yes' will start a shut down process." & vbNewLine & _
          "Some further processing will occur to clean up after Code Fixer." & vbNewLine & _
          "Please wait....", vbInformation + vbYesNo, AppDetails) = vbYes Then
      Set MyhourGlass = New cls_HourGlass
      bAborting = True
      bAbortComplete = False
    End If
  End If
  'Unload Me

End Sub

Public Sub DoABackUp()

  If frm_FindSettings.chkUser(3) = vbChecked Then
    If LenB(GetActiveProject.FileName) Then
      BackUpDeleteEngine "CodeFixBackUp"
      BackUpMakeOne
    End If
  End If

End Sub

Private Sub EmitChars()

  Dim I As Long
  Dim J As Long
  Dim K As Long

  'Emit deftype range chars
  For I = 1 To LastWordIndex
    J = Asc(Left$(Words(I), 1))
    If Len(Words(I)) < 3 Then
      K = J
     Else
      K = Asc(Mid$(Words(I), 3, 1))
    End If
    If J = Asc("a") Then
      If K = Asc("z") Then
        For J = 0 To 255
          DefTypeChars = DefTypeChars & Chr$(J)
        Next J
       Else
        Do
          DefTypeChars = DefTypeChars & Chr$(J)
          J = J + 1
        Loop Until J > K
      End If
    End If
  Next I

End Sub

Public Sub FormatAll()

  Dim SectionTotal   As Long

  On Error Resume Next
  SetupDataArrays
  RefreshUserSettingsFromString
  StructuralERRORFOUND = False
  ClearOldData
  With frm_CodeFixer
    .Visible = False
    If VBInstance.VBProjects.VBE.ActiveVBProject.VBComponents.Count Then
      If frm_FindSettings.chkUser(3).Value = vbChecked Then
        DoABackUp
      End If
      bFixing = True
      BreakLoop = False
      NFixDataClearAll
      SectionTotal = 12
      mObjDoc.GraphCaption "Fix && Format"
      'v3.0.9 new fixes
     SectionMessage "Repair faulty UserControls", 1 / SectionTotal
      WorkingMessage "Check UserControl Active Timer", 1, 2
      UserControlActiveTimerFix
      WorkingMessage "Check UserControl Font settings", 2, 2
      If UserControlFontRefFix Then
        bAbortComplete = True
        bFixing = False
        GoTo SafeExit
      End If
      SectionMessage "Build Control Arrays", 2 / SectionTotal
      If Not Control_Engine Then
        bAbortComplete = True
        bFixing = False
        GoTo SafeExit
      End If
      SectionMessage "Clean Up", 3 / SectionTotal
      CleanUp_Engine
      SectionMessage "XP && Frame test", 4 / SectionTotal
      XPManifestFrameWarning
      SectionMessage "Declaration Section", 5 / SectionTotal
      Declaration_Engine
      SectionMessage "Build Data Arrays", 6 / SectionTotal
      Generate_Publics
      SectionMessage "Restructure", 7 / SectionTotal
      ReStructure_Engine
      SectionMessage "Parameters", 8 / SectionTotal
      Parameter_Engine
      SectionMessage "Local variables", 9 / SectionTotal
      Dim_Engine
      SectionMessage "Suggest", 10 / SectionTotal
      Suggest_Engine
      SectionMessage "Unused", 11 / SectionTotal
      Unused_Engine
      SectionMessage "Layout", 12 / SectionTotal
      FormatDo
      If bAborting Then
        AbortFormatting
      End If
      OptionalCompileReEnable
      If Not StructuralERRORFOUND Then
        mObjDoc.ForceFind RGSignature
      End If
      bAbortComplete = True
      bFixing = False
      .Caption = AppDetails
     Else
      mObjDoc.Safe_MsgBox "Cannot see any code - you must open one or more Code Panels first.", vbInformation
    End If
SafeExit:
  End With
  On Error GoTo 0

End Sub

Private Sub FormatCode(cMod As CodeModule)

  
  Dim PleonasmDet     As Long
  Dim NumDeclLines    As Long
  Dim NumCodeLines    As Long
  Dim LineIndex       As Long
  Dim Codeline        As String
  Dim indent          As Long
  Dim CurrProcType    As String
  Dim Tmpstring1      As String
  Dim LineReset       As Boolean
  Dim ToLine          As Long
  Dim FromLine        As Long
  Dim FirstCodeLine   As Long
  Dim I               As Long
  Dim ActiveForVars   As String
  Dim OEIndent        As String
  Dim OEIndentAfter   As String
  Dim MaxIndent       As Long
  Dim IndentAfter     As Long
  Dim InEnum          As Boolean
  Dim LocalDim        As Boolean
  Dim NewProcStarting As Boolean
  Dim InProcHeader    As Boolean
  Dim MissingScope    As Boolean
  Dim NonProc         As Boolean
  Dim VaCoFuDcl       As Boolean
  Dim EmptyNext       As Boolean
  Dim NoOEClose       As Boolean
  Dim DeadCode        As Boolean
  Dim PossVio         As Boolean
  Dim ForVarMod       As Boolean
  Dim StrucErr        As Boolean
  Dim Skipped         As Long
  Dim bDummy          As Boolean

  'These are only used here so moved into routine to avoid Module level memory overhead
  'and modified the location of some of the code just to make it easier to read the code
  'Words = Split(VBInstance.ActiveWindow.Caption)
  'Module level moved to local level
  If Not bAborting Then
    If InsufficentCode(cMod) Then
      NotEnoughCode cMod
     Else
      If CleanUpCode(cMod) Then
        MoveUnusedRoutines cMod
        CleanDummyLine cMod
        If InsufficentCode(cMod) Then
          NotEnoughCode cMod
          DoEvents
         Else
          With cMod
            'prepare for scan
            Set VarNames = New Collection
            ReDim wordstack(0)
            ReDim StrucStack(0)
            indent = 0
            CodeIsDead = AlmostInfinity
            MaxIndent = 0
            IndentAfter = 0
            OEIndent = vbNullString
            DefTypeChars = vbNullString
            ActiveForVars = "="
            SetToDo Xcheck(XStructCom)
            FirstCodeLine = 0
            complain = False
            NonProc = False
            VaCoFuDcl = False
            SlIf = False
            ForVarMod = False
            MissingScope = False
            EmptyNext = False
            InProcHeader = False
            Dupl = False
            LocalDim = False
            Skipped = 0
            Inserted = 0
            NoOEClose = False
            StrucErr = False
            PossVio = False
            DeadCode = False
            'Here we go
            LineIndex = 1
            Do
              If LineIndex Mod 10 = 0 Or LineIndex = .CountOfLines Then
                MemberMessage "", LineIndex, .CountOfLines
              End If
              Do
LinesAdded:
                Codeline = vbNullString
                FromLine = LineIndex
                ConcatentForFormatCode cMod, Codeline, LineIndex
                ToLine = LineIndex
                ConcatentRangeForFormatCode Codeline
                Codeline = Trim$(Replace$(Codeline, Colon & SngSpace, SngSpace))
                If LenB(Codeline) = 0 Then
                  LineIndex = LineIndex + 1
                 Else
                  Exit Do
                End If
                If LineIndex > .CountOfLines Then
                  Exit Do
                End If
              Loop
              PleonasmDet = 0
              If LenB(Trim$(Codeline)) Then
                Words = Split(Codeline, SngSpace)
                LastWordIndex = UBound(Words)
                If LastWordIndex < 2 Then
                  ReDim Preserve Words(2)
                  Words(2) = SngSpace
                End If
                If Left$(Codeline, 1) = SQuote Then
                  Words(1) = Words(0)
                  Words(0) = "rem"
                End If
                If Words(0) <> "rem" Then
                  If FirstCodeLine = 0 Then
                    FirstCodeLine = FromLine
                  End If
                End If
                If Words(0) = "static" Then
                  If Words(1) <> "sub" Then
                    If Words(1) <> "function" Then
                      If Words(1) <> "property" Then
                        Words(0) = "dim"
                      End If
                    End If
                  End If
                End If
                Select Case Words(0)
                 Case "rem"
                  If InProcHeader Then
                    If Not LocalDim Then
                      IndentAfter = indent
                      indent = 0
                      HalfIndent = Space$(IndentSize)
                      NewProcStarting = True
                    End If
                  End If
                 Case "declare"
                  If Words(1) = "function" Then
                    CheckVardefs 2, False, False, True
                    If TypeSuffixFound Then
                      MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                     ElseIf Words(LastWordIndex - 1) <> "as" Then
                      MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                    End If
                    MarkLine cMod, LineIndex, ToDo12, MissingScope, FromLine, ToLine, indent
                  End If
                 Case "public", "private", "friend", "global", "static"
                  Select Case Words(0)
                   Case "global"
                    .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), "Global ", "Public ", , 1)
                   Case "static"
                    MarkLine cMod, LineIndex, ToDo12, MissingScope, FromLine, ToLine, indent
                  End Select
                  If indent Then
                    .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ToDo6
                    GoTo StructErrDetected
                  End If
                  indent = 0
                  Select Case Words(1)
                   Case "static"
                    Select Case Words(2)
                     Case "sub", "function", "property"
                      IndentAfter = 1
                      CurrProcType = UCase$(Left$(Words(2), 1)) & Mid$(Words(2), 2)
                      If LenB(Trim$(.Lines(ToLine + 1, 1))) Then
                        If Trim$(.Lines(ToLine + 1, 1)) <> REMBLANK Then
                          'v3.0.1 stops the multiline false break artifact
                          If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                            .InsertLines ToLine + 1, ""
                            KillSelection
                          End If
                        End If
                      End If
                      NewProcStarting = True
                      CodeIsDead = AlmostInfinity
                      LocalDim = False
                      If Words(2) = "function" Then
                        If CheckVardefs(3, False, False, True) Then
                          MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                        End If
                        If TypeSuffixFound Then
                          MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                        End If
                      End If
                    End Select
                   Case "sub", "function", "property", "enum", "type"
                    IndentAfter = 1
                    InEnum = (Words(1) = "enum")
                    CurrProcType = UCase$(Left$(Words(1), 1)) & Mid$(Words(1), 2)
                    If Words(1) <> "enum" Then
                      If Words(1) <> "type" Then
                        If LenB(Trim$(.Lines(ToLine + 1, 1))) Then
                          If Trim$(.Lines(ToLine + 1, 1)) <> REMBLANK Then
                            'v3.0.1 stops the multiline false break artifact
                            If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                              .InsertLines ToLine + 1, ""
                              KillSelection
                            End If
                          End If
                        End If
                        NewProcStarting = True
                        CodeIsDead = AlmostInfinity
                        LocalDim = False
                        If Words(1) = "function" Then
                          If CheckVardefs(2, False, False, True) Then
                            MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                          End If
                          If TypeSuffixFound Then
                            MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                          End If
                        End If
                      End If
                    End If
                   Case "declare"
                    If Words(2) = "function" Then
                      CheckVardefs 3, False, False, True
                      If TypeSuffixFound Then
                        MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                       ElseIf Words(LastWordIndex - 1) <> "as" Then
                        MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                      End If
                    End If
                    'Case "event", "withevents" ' stop 'Case Else' case hitting these
                   Case Else
                    If Not ArrayMember(Words(1), "event", "withevents") Then
                      IndentAfter = 0
                      If CheckVardefs(IIf(Words(1) = "const", 2, 1), False, Words(1) = "const", False) Then
                        MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                      End If
                      If DuplNameDetected Then
                        MarkLine cMod, LineIndex, ToDo8, Dupl, FromLine, ToLine, indent
                      End If
                      If DefConstType Then
                        MarkLine cMod, LineIndex, ToDo9, VaCoFuDcl, FromLine, ToLine, indent
                      End If
                      If TypeSuffixFound Then
                        MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                      End If
                    End If
                  End Select
                 Case "dim", "const"
                  If CheckVardefs(1 + Abs(Words(1) = "withevents"), indent <> 0, Words(0) = "const", False) Then
                    MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                  End If
                  If DuplNameDetected Then
                    MarkLine cMod, LineIndex, ToDo8, Dupl, FromLine, ToLine, indent
                  End If
                  If DefConstType Then
                    MarkLine cMod, LineIndex, ToDo9, VaCoFuDcl, FromLine, ToLine, indent
                  End If
                  If TypeSuffixFound Then
                    MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                  End If
                  If indent Then
                    If Not LocalDim Then
                      If InProcHeader Then
                        If LenB(Trim$(.Lines(FromLine - 1, 1))) Then
                          If Trim$(.Lines(FromLine, 1)) <> REMBLANK Then
                            'v3.0.1 stops the multiline false break artifact
                            If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                              .InsertLines FromLine, ""
                              KillSelection
                            End If
                          End If
                        End If
                      End If
                    End If
                    IndentAfter = indent
                    indent = 0
                    HalfIndent = Space$(IndentSize)
                    If InProcHeader Then
                      NewProcStarting = True
                     Else
                      MarkLine cMod, LineIndex, ToDo1 & IIf(Len(ToDo1), CurrProcType, vbNullString), NonProc, FromLine, ToLine, indent
                    End If
                    LocalDim = True
                   Else
                    MarkLine cMod, LineIndex, ToDo12, MissingScope, FromLine, ToLine, indent
                  End If
                 Case "event"
                  MarkLine cMod, LineIndex, ToDo12, MissingScope, FromLine, ToLine, indent
                 Case "exit"
                  SubsequentCodeIsDead indent
                  Select Case Words(1)
                   Case "sub", "function", "property"
                    If indent = 1 Then
                      indent = 0
                      IndentAfter = 1
                      If LenB(Trim$(.Lines(ToLine + 1, 1))) Then
                        If Trim$(.Lines(ToLine + 1, 1)) <> REMBLANK Then
                          'v3.0.1 stops the multiline false break artifact
                          If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                            .InsertLines ToLine + 1, ""
                            KillSelection
                            LineIndex = ToLine + 1
                          End If
                        End If
                      End If
                      If LenB(Trim$(.Lines(FromLine - 1, 1))) Then
                        If Trim$(.Lines(FromLine, 1)) <> REMBLANK Then
                          'v3.0.1 stops the multiline false break artifact
                          If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                            .InsertLines FromLine, ""
                            KillSelection
                          End If
                        End If
                      End If
                    End If
                   Case "for", "do"
                    If GetStruc(False) <> Words(1) Then
                      MarkLine cMod, LineIndex, ToDo14, PossVio, FromLine, ToLine, indent
                    End If
                  End Select
                 Case "sub", "function", "property", "enum", "type"
                  indent = 0
                  CodeIsDead = AlmostInfinity
                  IndentAfter = 1
                  InEnum = (Words(0) = "enum")
                  CurrProcType = UCase$(Left$(Words(0), 1)) & Mid$(Words(0), 2)
                  If Words(0) <> "enum" Then
                    If Words(0) <> "type" Then
                      If LenB(Trim$(.Lines(ToLine + 1, 1))) Then
                        If Trim$(.Lines(ToLine + 1, 1)) <> REMBLANK Then
                          'v3.0.1 stops the multiline false break artifact
                          If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                            .InsertLines ToLine + 1, ""
                            KillSelection
                          End If
                        End If
                      End If
                      NewProcStarting = True
                      LocalDim = False
                      If Words(0) = "function" Then
                        If CheckVardefs(1, False, False, True) Then
                          MarkLine cMod, LineIndex, ToDo7, VaCoFuDcl, FromLine, ToLine, indent
                        End If
                        If TypeSuffixFound Then
                          MarkLine cMod, LineIndex, ToDo10, VaCoFuDcl, FromLine, ToLine, indent
                        End If
                      End If
                    End If
                  End If
                  If Words(0) <> "enum" Then
                    If Words(0) <> "type" Then
                      MarkLine cMod, LineIndex, ToDo12, MissingScope, FromLine, ToLine, indent
                    End If
                  End If
                 Case "if"
                  If Words(LastWordIndex) = "then" Then
                    IndentAfter = indent + 1
                    PleonasmDet = InStr(Codeline, "= true ") + InStr(Codeline, " true =") + InStr(Codeline, "= true)") + InStr(Codeline, "(true =")
                    PushWord SQuote & IIf(Words(2) = "then", vbNullString, "NOT ") & Words(1) & IIf(Words(2) = "then", " = FALSE/0", "...")
                   Else
                    MarkLine cMod, LineIndex, ToDo2, SlIf, FromLine, ToLine, indent, LineReset
                    If LineReset Then
                      LineReset = False
                    End If
                    Select Case Words(LastWordIndex)
                     Case "sub", "function", "property", "for", "do"
                      .ReplaceLine LineIndex, .Lines(LineIndex, 1) & IIf(SmartRight(.Lines(LineIndex, 1), "Expand Structure"), ToDo2a, vbNullString)
                    End Select
                  End If
                 Case "#if", "#else", "#elseif", "#end", "#const"
                  IndentAfter = indent
                  indent = 0
                  OEIndentAfter = OEIndent
                  OEIndent = vbNullString
                 Case "do", "for", "while", "select", "with"
                  IndentAfter = indent + 1
                  If Words(0) <> "select" Then
                    StrucStack(UBound(StrucStack)) = Words(0)
                    ReDim Preserve StrucStack(UBound(StrucStack) + 1)
                  End If
                  Select Case Words(0)
                   Case "do"
                    If InStr(Codeline & SngSpace, " loop ") Then
                      IndentAfter = indent
                      MarkLine cMod, LineIndex, ToDo2, StrucErr, FromLine, ToLine, indent
                      If BreakLoop Then
                        GoTo StructErrDetected
                      End If
                    End If
                   Case "for"
                    If InStr(Codeline & SngSpace, " next ") Then
                      IndentAfter = indent
                      MarkLine cMod, LineIndex, ToDo2, StrucErr, FromLine, ToLine, indent
                      If BreakLoop Then
                        GoTo StructErrDetected
                      End If
                     Else
                      If Words(1) = "each" Then
                        PushWord Words(2)
                       Else
                        If TypeSuffixExists(Words(1)) Then
                          Words(1) = Left$(Words(1), Len(Words(1)) - 1)
                        End If
                        PushWord Words(1)
                        ActiveForVars = ActiveForVars & Words(1) & "="
                      End If
                    End If
                   Case "while"
                    If InStr(Codeline & SngSpace, " wend ") Then
                      IndentAfter = indent
                      MarkLine cMod, LineIndex, ToDo2, StrucErr, FromLine, ToLine, indent
                      If BreakLoop Then
                        GoTo StructErrDetected
                      End If
                    End If
                   Case "with"
                    If InStr(Codeline & SngSpace, " end with ") Then
                      IndentAfter = indent
                      MarkLine cMod, LineIndex, ToDo2, StrucErr, FromLine, ToLine, indent
                      GoTo StructErrDetected
                     Else
                      PushWord Words(1)
                    End If
                  End Select
                  PleonasmDet = InStr(Codeline & SngSpace, "= true ") + InStr(Codeline, " true =") + InStr(Codeline, "= true)") + InStr(Codeline, "(true =")
                 Case "on"
                  If Words(1) = "error" Then
                    Select Case Words(2)
                     Case "resume"
                      OEIndentAfter = Space$(FullTabWidth)
                     Case "goto"
                      OEIndentAfter = vbNullString
                      OEIndent = vbNullString
                    End Select
                   ElseIf Words(2) = "error" Then
                    Select Case Words(3)
                     Case "resume"
                      OEIndentAfter = Space$(FullTabWidth)
                     Case "goto"
                      OEIndentAfter = vbNullString
                      OEIndent = vbNullString
                    End Select
                  End If
                 Case "case" 'v3.0.0 Separate case and Else indenting
                  'v2.0.7 Paul Caton suggested this
                  Select Case lngCaseIndent
                    'Case 0' no indent(beyond the If/Case level already used
                   Case 1
                    HalfIndent = Space$(IndentSize / 2)
                   Case 2
                    HalfIndent = Space$(IndentSize)
                  End Select
                  IndentAfter = indent
                  indent = indent - 1
                  If indent < 0 Then
                    indent = 0
                    MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                    GoTo StructErrDetected
                  End If
                  PleonasmDet = InStr(Codeline & SngSpace, "= true ") + InStr(Codeline, " true =") + InStr(Codeline, "= true)") + InStr(Codeline, "(true =")
                  'Ulli's structural comments
                  Select Case Words(0)
                   Case "else"
                    PopPush cMod, Codeline = "else", FromLine, vbNullString, vbNullString, FromLine, ToLine, indent
                   Case "elseif"
                    PopPush cMod, Words(LastWordIndex) = "then", FromLine, Words(1), Words(2), FromLine, ToLine, indent
                  End Select
                 Case "else", "elseif"
                  'v2.0.7 Paul Caton suggested this
                  Select Case lngElseIndent
                    'Case 0' no indent(beyond the If/Case level already used
                   Case 1
                    HalfIndent = Space$(IndentSize / 2)
                   Case 2
                    HalfIndent = Space$(IndentSize)
                  End Select
                  IndentAfter = indent
                  indent = indent - 1
                  If indent < 0 Then
                    indent = 0
                    MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                    GoTo StructErrDetected
                  End If
                  PleonasmDet = InStr(Codeline & SngSpace, "= true ") + InStr(Codeline, " true =") + InStr(Codeline, "= true)") + InStr(Codeline, "(true =")
                  'Ulli's structural comments
                  Select Case Words(0)
                   Case "else"
                    PopPush cMod, Codeline = "else", FromLine, vbNullString, vbNullString, FromLine, ToLine, indent
                   Case "elseif"
                    PopPush cMod, Words(LastWordIndex) = "then", FromLine, Words(1), Words(2), FromLine, ToLine, indent
                  End Select
                 Case "end"
                  If Codeline = "end" Then
                    Words(0) = "END"
                    SubsequentCodeIsDead indent
                   Else
                    'this line is a special trap for the case of a Type member named 'End'
                    'Very rare but causes a Structural Error if it hits
                    If Not InStr(Codeline, " as ") <> 0 Then
                      indent = indent - 1
                      InEnum = False
                      If indent < 0 Then
                        indent = 0
                        If Codeline <> "end sub" Then
                          If Codeline <> "end function" Then
                            If Codeline <> "end property" Then
                              MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                              GoTo StructErrDetected
                            End If
                          End If
                        End If
                      End If
                    End If
                    IndentAfter = indent
                    Select Case Words(1)
                     Case "sub", "function", "property"
                      If indent Then
                        MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                        GoTo StructErrDetected
                       ElseIf LenB(OEIndent) Then
                        MarkLine cMod, LineIndex, ToDo4, NoOEClose, FromLine, ToLine, indent, LineReset
                        If LineReset Then
                          LineReset = False
                          GoTo LinesAdded
                        End If
                      End If
                      If LenB(Trim$(.Lines(ToLine + 1, 1))) Then
                        If Left$(Trim$(.Lines(ToLine + 1, 1)) & SngSpace, 2) <> "#E" Then
                          'the next line is not blank and not #EndIf nor #ElseIf nor End Copy
                          If Trim$(.Lines(ToLine + 1, 1)) <> REMBLANK Then
                            'v3.0.1 stops the multiline false break artifact
                            If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                              .InsertLines ToLine + 1, ""
                              KillSelection
                            End If
                          End If
                        End If
                      End If
                      If LenB(Trim$(.Lines(FromLine - 1, 1))) Then
                        If Trim$(.Lines(FromLine - 1, 1)) <> REMBLANK Then
                          'v3.0.1 stops the multiline false break artifact
                          If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                            .InsertLines FromLine, ""
                            KillSelection
                          End If
                          LineIndex = ToLine + 1
                        End If
                      End If
                      OEIndentAfter = vbNullString
                      OEIndent = vbNullString
                      ReDim wordstack(0)
                      ReDim StrucStack(0)
                     Case "with"
                      GetStruc True
                      Tmpstring1 = SQuote & PopWord()
                      If Xcheck(XStructCom) Then
                        If InStr(1, .Lines(LineIndex, 1), Tmpstring1, vbTextCompare) = 0 Then
                          .ReplaceLine LineIndex, .Lines(LineIndex, 1) & Tmpstring1
                        End If
                       Else
                        'v2.7.8 oddity if Strutural Code message has extra comment tagged to the end
                        'then only the ' and structural method are deleted leaving rest of comment in code
                        If SmartRight(.Lines(LineIndex, 1), Tmpstring1) = False Then
                          .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), Tmpstring1, "'")
                         Else
                          .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), Tmpstring1, vbNullString)
                        End If
                      End If
                     Case "if"
                      PopWord
                    End Select
                  End If
                 Case "loop", "wend"
                  indent = indent - 1
                  GetStruc True
                  If indent < 0 Then
                    indent = 0
                    MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                    GoTo StructErrDetected
                  End If
                  IndentAfter = indent
                  PleonasmDet = InStr(Codeline & SngSpace, "= true ") + InStr(Codeline, " true =") + InStr(Codeline, "= true)") + InStr(Codeline, "(true =")
                 Case "next"
                  indent = indent - 1
                  GetStruc True
                  I = 4
                  Do
                    I = InStr(I + 1, Codeline, ",")
                    If I Then
                      indent = indent - 1
                      GetStruc True
                      ActiveForVars = Replace$(ActiveForVars, "=" & PopWord & "=", "=", , , vbTextCompare)
                     Else
                      Exit Do
                    End If
                  Loop
                  If indent < 0 Then
                    indent = 0
                    MarkLine cMod, LineIndex, ToDo6, StrucErr, FromLine, ToLine, indent
                    GoTo StructErrDetected
                  End If
                  IndentAfter = indent
                  'Ulli's structural comments
                  Tmpstring1 = PopWord
                  ActiveForVars = Replace$(ActiveForVars, "=" & Tmpstring1 & "=", "=", , , vbTextCompare)
                  If Codeline = "next" Then
                    MarkLine cMod, LineIndex, ToDo3 & SngSpace & Tmpstring1, EmptyNext, FromLine, ToLine, indent
                  End If
                 Case "goto"
                  SubsequentCodeIsDead indent
                  complain = SetComplain(FromLine, ToLine)
                 Case "return"
                  SubsequentCodeIsDead indent
                 Case "let"
                  If TypeSuffixExists(Words(1)) Then
                    Words(1) = Left$(Words(1), Len(Words(1)) - 1)
                  End If
                  If InStr(ActiveForVars, "=" & Words(1) & Words(2)) Then
                    MarkLine cMod, LineIndex, ToDo11, ForVarMod, FromLine, ToLine, indent
                  End If
                 Case "deflng", "defbool", "defbyte", "defint", "defcur", "defsng", "defdbl", "defdec", "defdate", "defstr", "defobj"
                  EmitChars
                 Case Else
                  If InEnum Then
                    CheckVardefs 0, False, False, False
                    If DuplNameDetected Then
                      MarkLine cMod, LineIndex, ToDo8, Dupl, FromLine, ToLine, indent
                    End If
                   Else
                    If TypeSuffixExists(Words(0)) Then
                      Words(0) = Left$(Words(0), Len(Words(0)) - 1)
                    End If
                    If InStr(ActiveForVars, "=" & Words(0) & Words(1)) Then
                      MarkLine cMod, LineIndex, ToDo11, ForVarMod, FromLine, ToLine, indent
                    End If
                    If Right$(Words(0), 1) = Colon Then
                      CodeIsDead = AlmostInfinity
                    End If
                  End If
                End Select
                If indent < CodeIsDead Then
                  CodeIsDead = AlmostInfinity
                 Else
                  If CodeIsDead >= 0 Then
                    If Words(0) <> "rem" Then
                      MarkLine cMod, LineIndex, ToDo15, DeadCode, FromLine, ToLine, indent
                    End If
                  End If
                End If
                If InProcHeader Then
                  If Not NewProcStarting Then
                    If LenB(Trim$(.Lines(FromLine - 1, 1))) Then
                      If Trim$(.Lines(FromLine - 1, 1)) <> REMBLANK Then
                        'v3.0.1 stops the multiline false break artifact
                        If Not HasLineCont(Trim$(.Lines(ToLine, 1))) Then
                          .InsertLines FromLine, ""
                        End If
                        FromLine = FromLine + 1
                        ToLine = ToLine + 1
                        LineIndex = LineIndex + 1
                      End If
                    End If
                  End If
                End If
                If PleonasmDet Then
                  MarkLine cMod, LineIndex, ToDo5, bDummy, FromLine, ToLine, indent
                End If
                For I = FromLine To ToLine
                  .ReplaceLine I, String$(indent, vbTab) & HalfIndent & OEIndent & Trim$(.Lines(I, 1))
                  If indent > MaxIndent Then
                    MaxIndent = indent
                  End If
                  'indentation for continuation lines
                  HalfIndentSet I, 0, FromLine, HalfIndent
                Next I
                indent = IndentAfter
                OEIndent = OEIndentAfter
                HalfIndent = vbNullString
                InProcHeader = NewProcStarting
                NewProcStarting = False
                CodeIsDead = Abs(CodeIsDead)
              End If
StructErrDetected:
              If BreakLoop Then
                Exit Do
              End If
              If Xcheck(XVisScan) Then
                .CodePane.SetSelection LineIndex, 1, LineIndex, 1
              End If
              LineIndex = LineIndex + 1
            Loop Until LineIndex > .CountOfLines
            If LenB(.Lines(.CountOfDeclarationLines + 1, 1)) Then
              If Trim$(.Lines(.CountOfDeclarationLines + 1, 1)) <> REMBLANK Then
                .InsertLines .CountOfDeclarationLines + 1, ""
              End If
            End If
            '---End of formatting code-----------------------------------------------
            NumDeclLines = .CountOfDeclarationLines
            NumCodeLines = .CountOfLines - NumDeclLines
            If NumCodeLines = 0 Then
              NumDeclLines = NumDeclLines + 2
             Else
              NumCodeLines = NumCodeLines + 2
            End If
            'Sign off line
            .InsertLines .CountOfLines + 1, Smiley & AppDetails & " (" & Now & ") " & NumDeclLines & StrPlus & NumCodeLines & EqualInCode & NumDeclLines + NumCodeLines & " Lines" & IIf(Skipped, strInBrackets("Skipped " & Skipped), "") & " Thanks Ulli for inspiration and lots of code."
            Set VarNames = Nothing
          End With
          If Xcheck(XVisScan) Then
            With cMod.CodePane
              Select Case True
               Case complain
                .SetSelection ErrLineFrom, 1, ErrLineTo, Len(.CodeModule.Lines(ErrLineTo, 1)) + 1
               Case Sel.StartLine <> 0
                .SetSelection Sel.StartLine, Sel.startCol, Sel.EndLine, Sel.endCol
               Case Else
                .SetSelection 1, 1, 1, 1
              End Select
              .TopLine = IIf(curTopLine > 0, curTopLine, 1)
            End With
          End If
        End If
      End If
    End If
  End If

End Sub

Private Sub FormatDo()

  Dim strProc      As String
  Dim ELine        As Long
  Dim arrTmp       As Variant
  Dim ArrTmp2()    As Variant
  Dim I            As Long
  Dim Sline        As Long
  Dim StartLine    As Long
  Dim CurCompCount As Long
  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim cMod         As CodeModule
  Dim NumFixes     As Long

  NumFixes = 3
  ' allows last one to exit outside the For Loops and be used in report
  On Error GoTo BugHit
  If Not bAborting Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          'v3.0.5 moved location for better effects; unused Procedures moved to end, bragline at end of code and
          ' a bug if last Declaration line was '#End If'.
          'v3.0.4 sorting moved here to make it work properly (some fixes reorder so don't sort until everything else is fixed
          If Not ModDesc(CurCompCount).MDDontTouch Then
            'ProtectEnumCap DecArray
            'ReWriter CompMod, DecArray, RWDeclaration
            WorkingMessage "Sorting Routines", 1, NumFixes
            DoSorting Comp.CodeModule
          End If
          ModuleMessage Comp, CurCompCount
          Set cMod = Comp.CodeModule
          DisplayCodePane Comp
          WorkingMessage "Indenting", 2, NumFixes
          FormatCode cMod
          WorkingMessage "Line continuation Layout", 3, NumFixes
          ReDoLineContinuation cMod
          'v2.9.5 Reposition Paul Caton required procedure
          If bPaulCatonSubClasUsed Then
            If strPaulCatonSubClasCompName = cMod.Parent.Name Then
              If strPaulCatonSubClasProcName = GetFirstPublicProcedureName(cMod) Then
                bPaulCatonSubClasUsed = False ' it is in the right place
               Else
                If cMod.Find(strPaulCatonSubClasProcName, Sline, 1, -1, -1) Then
                  'it is in this module so move it
                  arrTmp = ReadProcedureCodeArray(Sline, ELine, strPaulCatonSubClasProcName)
                  'in case there are commented out procedures below the target procedure
                  For I = UBound(arrTmp) To 0 Step -1
                    If JustACommentOrBlank(arrTmp(I)) Then
                      ELine = ELine - 1
                      arrTmp(I) = vbNullString
                     Else
                      Exit For
                    End If
                  Next I
                  ReDim ArrTmp2(ELine - Sline - 1) As Variant
                  For I = LBound(ArrTmp2) To UBound(ArrTmp2)
                    ArrTmp2(I) = arrTmp(I)
                  Next I
                  strProc = Join(ArrTmp2, vbNewLine)
                  cMod.DeleteLines Sline, ELine - Sline
                  cMod.AddFromString strProc
                  bPaulCatonSubClasUsed = False
                End If
              End If
            End If
          End If
          'v2.3.8 faster
          'v.2.4.4 Thanks Aaron Spivey correctly remove the blank line preserving text
          If Xcheck(XBlankPreserve) Then
            With cMod
              'v2.7.2 thanks Fabricio Ferreira' Not resetting StartLine meant that this fix only hit in first module
              StartLine = 0
              Do While .Find(REMBLANK, StartLine, -1, 1, -1)
                If Trim$(.Lines(StartLine, 1)) = REMBLANK Then
                  .ReplaceLine StartLine, ""
                End If
                StartLine = StartLine + 1
              Loop
            End With
          End If
          RestoreMemberAttributes CurCompCount, cMod.Members
        End If
        If BreakLoop Then
          Exit For
        End If
      Next Comp
      If BreakLoop Then
        Exit For
      End If
      If bAborting Then
        Exit For 'Sub
      End If
    Next Proj
  End If
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "FormatDo"
  If Not bAborting Then
    If RunningInIDE Then
      Resume
     Else
      Resume Next
    End If
  End If

End Sub

Public Sub FormatOnly()

  Dim SectionTotal As Long

  On Error Resume Next
  SetupDataArrays
  SettingCollector
  StructuralERRORFOUND = False
  ClearOldData
  bNoOptExp = True
  With frm_CodeFixer
    If VBInstance.VBProjects.VBE.ActiveVBProject.VBComponents.Count Then
      If frm_FindSettings.chkUser(3).Value = vbChecked Then
        DoABackUp
      End If
      mObjDoc.GraphCaption "Format only"
      bFixing = True
      BreakLoop = False
      NFixDataClearAll
      SectionTotal = 5
      SectionMessage "Clean Up", 1 / SectionTotal
      CleanUp_Engine
      SectionMessage "Declarations", 2 / SectionTotal
      Declaration_Engine
      SectionMessage "Restructure", 3 / SectionTotal
      ReStructure_EngineFormat
      SectionMessage "Local", 4 / SectionTotal
      Dim_EngineFormat
      SectionMessage "Layout", 5 / SectionTotal
      FormatDo
      If bAborting Then
        AbortFormatting
      End If
      OptionalCompileReEnable
      If Not StructuralERRORFOUND Then
        mObjDoc.ForceFind RGSignature
      End If
      bAbortComplete = True
      bFixing = False
      .Caption = AppDetails
     Else
      mObjDoc.Safe_MsgBox "Cannot see any code - you must open one or more Code Panels first.", vbInformation
    End If
  End With
  OptionalCompileReEnable
  bNoOptExp = False
  On Error GoTo 0

End Sub

Private Function GetStruc(ByVal PopIt As Boolean) As String

  Dim K As Long

  'Stack keeps track of the current stucture level type, ie
  'whether we are inside a for, do, while, with &c structure bracket
  'This stack is used to detect possible structure violations by exit statements
  K = UBound(StrucStack) - 1
  If K < 0 Then
    GetStruc = StackUnderflow 'Mid$(StackUnderflow, 2)
   Else
    If PopIt Then
      ReDim Preserve StrucStack(K)
    End If
    GetStruc = StrucStack(K)
  End If

End Function

Private Sub HalfIndentSet(ByVal I As Long, _
                          J As Long, _
                          ByVal FromLine As Long, _
                          HalfIndent As String)

  'Moved out of FormatCode to make it more readable

  If I = FromLine Then
    Select Case Words(0)
     Case "public", "global", "private", "dim", "friend"
      J = 1
      HalfIndent = HalfIndent & Space$(Len(Words(0)) + 1)
      'adjusted so that the params are aligned
      Select Case Words(1)
       Case "sub", "function"
        HalfIndent = HalfIndent & Space$(Len(Words(1)) + 1)
        J = 2
       Case "property"
        HalfIndent = HalfIndent & Space$(Len(Words(1)) + 5)
        J = 3
       Case "declare"
        HalfIndent = HalfIndent & Space$(Len(Words(1)) + 3)
        J = 3
      End Select
     Case "sub", "function"
      HalfIndent = HalfIndent & Space$(Len(Words(0)) + 1)
      J = 2
     Case Else
      J = 0
    End Select
    J = GetLeftBracketPos(Words(J))
    If J Then
      HalfIndent = HalfIndent & Space$(J)
     Else
      HalfIndent = HalfIndent & Space$(Len(Words(0)) + IIf(Words(1) = "=", 3, 1))
    End If
  End If

End Sub

Private Function InsufficentCode(cMod As CodeModule) As Boolean

  Dim TmpCount As Long
  Dim I        As Long

  InsufficentCode = True
  For I = 1 To cMod.CountOfLines
    If ExtractCode(cMod.Lines(I, 1)) Then
      TmpCount = TmpCount + 1
      If TmpCount > 2 Then
        InsufficentCode = False
        Exit For 'unction
      End If
    End If
  Next I

End Function

Public Sub KillSelection()

  Sel.StartLine = 0
  curTopLine = 1

End Sub

Private Sub MarkLine(cMod As CodeModule, _
                     ByVal LIndex As Long, _
                     ByRef MarkText As String, _
                     ByRef ErrFlag As Boolean, _
                     ByRef FromLine As Long, _
                     ByRef ToLine As Long, _
                     ByVal indent As Long, _
                     Optional ByVal ResetForAddedLine As Boolean)

  'This was the first version of Code Fixer and is still here as a final guard
  'ReWrites codeline either
  'by attaching Ulli's original comments
  'OR
  'rewriting the code itself and attaching my comments
  '
  'Optional ResetForAddedLine resets the Scan Loop if any AutoFix routine  adds vbNewLine to text
  '(and therefore changes the linecount) so that Indenting etc can work properly

  ErrFlag = True
  If Len(MarkText) > 1 Then
    AutoFix cMod, LIndex, MarkText, indent, ResetForAddedLine
  End If 'Select
ReWriteDone:
  complain = SetComplain(FromLine, ToLine)

End Sub

Private Function ModuleIsEmpty(Wmod As CodeModule) As Boolean

  Dim I As Long

  If ModuleHasCode(Wmod) Then
    For I = 1 To Wmod.CountOfLines
      If Left$(Trim$(Wmod.Lines(I, 1)), 1) <> SQuote Then
        If Not Left$(Trim$(Wmod.Lines(I, 1)), 7) = "Option " Then
          'any code other than 'Option XXXX' which is module level only anyway
          'will declare the module has code
          GoTo SafeExit
        End If
      End If
    Next I
    ModuleIsEmpty = True
   Else
    For I = 1 To Wmod.CountOfLines
      If Left$(Trim$(Wmod.Lines(I, 1)), 1) <> SQuote Then
        If Not Left$(Trim$(Wmod.Lines(I, 1)), 7) = "Option " Then
          'any code other than 'Option XXXX' which is module level only anyway
          'will declare the module has code
          GoTo SafeExit
        End If
      End If
    Next I
    ModuleIsEmpty = Wmod.CountOfDeclarationLines > 0
  End If
SafeExit:

End Function

Private Sub MoveUnusedRoutines(cMod As CodeModule)

  Dim I         As Long
  Dim MaxFactor As Long
  Dim ArrCode   As Variant
  Dim strMoving As String

  'this solves problem of routine sorting moving commented out routines away from end of page
  With cMod
    ArrCode = GetModuleArray(cMod, False)
    MaxFactor = UBound(ArrCode)
    For I = 1 To MaxFactor
      If SmartLeft(ArrCode(I), MoveableComment) Then
        strMoving = strMoving & vbNewLine & ArrCode(I)
        ArrCode(I) = vbNullString
      End If
    Next I
    Do While InStr(strMoving, MoveableComment)
      'v 2.2.2 remove the MoveableComment marker
      strMoving = Replace$(strMoving, MoveableComment, "''")
    Loop
    ArrCode(MaxFactor) = ArrCode(MaxFactor) & vbNewLine & strMoving
    ReWriter cMod, ArrCode, RWModule
    'DoEvents
  End With

End Sub

Private Sub NFixDataClearAll()

  Dim I As Long

  For I = 0 To DummyFixEnd - 1 'UBound(FixData)
    FixData(I).FixModCount = 0
    FixData(I).FixTotalCount = 0
  Next I

End Sub

Private Sub NotEnoughCode(codeMod As CodeModule)

  If ModuleType(codeMod) = "Module" Then
    If ModuleIsEmpty(codeMod) Then
      codeMod.InsertLines 1, WARNING_MSG & "This module contains no code and can be deleted from project"
    End If
  End If

End Sub

Private Sub PopPush(cMod As CodeModule, _
                    ByVal Condition As Boolean, _
                    ByVal LLine As Long, _
                    ByVal Word1 As String, _
                    ByVal Word2 As String, _
                    ByRef FromLine As Long, _
                    ByRef ToLine As Long, _
                    ByVal indent As Long)

  Dim Tmpstring1 As String

  With cMod
    If Condition Then
      Tmpstring1 = PopWord()
      If Xcheck(XStructCom) Then
        'v2.2.0 maybe fixes error with multiple messages being attached
        If SmartRight(.Lines(LLine, 1), Tmpstring1) = False Then
          .ReplaceLine LLine, .Lines(LLine, 1) & Tmpstring1
        End If
       Else
        'v2.7.8 oddity if Strutural Code message has extra comment tagged to the end
        'then only the ' and structural method are deleted leaving rest of comment in code
        If SmartRight(.Lines(LLine, 1), Tmpstring1) = False Then
          .ReplaceLine LLine, Replace$(.Lines(LLine, 1), Tmpstring1, "'")
         Else
          .ReplaceLine LLine, Replace$(.Lines(LLine, 1), Tmpstring1, vbNullString)
        End If
      End If
      PushWord SQuote & IIf(Word2 = "then", vbNullString, "NOT ") & Word1 & IIf(Word2 = "then", " = FALSE/0", "...")
     Else
      MarkLine cMod, LLine, ToDo2, SlIf, FromLine, ToLine, indent
      Select Case Words(LastWordIndex)
       Case "sub", "function", "property", "for", "do"
        .ReplaceLine LLine, .Lines(LLine, 1) & ToDo2a
      End Select
    End If
  End With

End Sub

Private Function PopWord() As String

  Dim K As Long

  'Retrieve Word from Stack
  K = UBound(wordstack) - 1
  If K < 0 Then
    PopWord = StackUnderflow 'Mid$(StackUnderflow, 2)
   Else
    ReDim Preserve wordstack(K)
    PopWord = UCase$(wordstack(K))
  End If

End Function

Private Sub PushWord(ByVal Word As String)

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'Save (modified) Word on Stack

  wordstack(UBound(wordstack)) = Replace$(Word, vbNullChar, SngSpace)
  ReDim Preserve wordstack(UBound(wordstack) + 1)

End Sub

Public Sub RestoreMemberAttributes(ByVal ModuleNumber As Long, _
                                   membs As Members)

  Dim I                  As Long
  Dim MemberAttributes() As Variant   'vbArray of 13 values

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'restore the member attributes
  MemberAttributes = Attributes(ModuleNumber)
  For I = 1 To UBound(MemberAttributes)
    Err.Clear
    On Error Resume Next
    With membs(MemberAttributes(I)(MemName))
      'may produce an error on undo when member attributes cannot be restored
      'because a new member was created after the last format scan and thats
      'now missing in the undo buffer but it's attributes have been saved
      If Err.Number = 0 Then
        If LenB(MemberAttributes(I)(MemCate)) Then
          .Category = MemberAttributes(I)(MemCate)
        End If
        If LenB(MemberAttributes(I)(MemDesc)) Then
          .Description = MemberAttributes(I)(MemDesc)
        End If
        If MemberAttributes(I)(MemHelp) Then
          .HelpContextID = MemberAttributes(I)(MemHelp)
        End If
        If LenB(MemberAttributes(I)(MemProp)) Then
          .PropertyPage = MemberAttributes(I)(MemProp)
        End If
        If MemberAttributes(I)(MemStme) <= 0 Then
          .StandardMethod = MemberAttributes(I)(MemStme)
        End If
        If MemberAttributes(I)(MemBind) Then
          .Bindable = True
        End If
        If MemberAttributes(I)(MemBrws) Then
          .Browsable = True
        End If
        If MemberAttributes(I)(MemDfbd) Then
          .DefaultBind = True
        End If
        If MemberAttributes(I)(MemDbnd) Then
          .DisplayBind = True
        End If
        If MemberAttributes(I)(MemHidd) Then
          .Hidden = True
        End If
        If MemberAttributes(I)(MemRqed) Then
          .RequestEdit = True
        End If
        If MemberAttributes(I)(MemUide) Then
          .UIDefault = True
        End If
      End If
    End With
    On Error GoTo 0
  Next I

End Sub

Private Function SetComplain(ByRef FromLine As Long, _
                             ByRef ToLine As Long) As Boolean

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)

  SetComplain = True
  ErrLineFrom = FromLine
  curTopLine = FromLine
  ErrLineTo = ToLine

End Function

Private Sub SetToDo(ByVal bActivate As Boolean)

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)

  ToDo1 = IIf(bActivate, RGSignature & "Move line to top of current ", vbNullString)
  ToDo2 = IIf(bActivate, RGSignature & "Expand Structure", vbNullString)
  ToDo2a = IIf(bActivate, " or consider reversing Condition", vbNullString)
  ToDo3 = IIf(bActivate, RGSignature & "Repeat For-Variable:", vbNullString)
  ToDo4 = IIf(bActivate, RGSignature & "On Error Resume still active", vbNullString)
  ToDo5 = IIf(bActivate, RGSignature & "Remove Pleonasm", vbNullString)
  ToDo6 = IIf(bActivate, RGSignature & "Structure Error", vbNullString)
  ToDo7 = IIf(bActivate, RGSignature & ": UnTyped Variable. Will behave as Variant", vbNullString)
  ToDo8 = IIf(bActivate, WARNING_MSG & "Duplicated Name. It is also declared at Module or Program wide level.", vbNullString)
  ToDo9 = IIf(bActivate, RGSignature & "As 16-bit Integer ?", vbNullString)
  ToDo10 = IIf(bActivate, RGSignature & "Obsolete Type Suffix", vbNullString)
  ToDo11 = IIf(bActivate, WARNING_MSG & "Modifies active For-Variable", vbNullString)
  ToDo12 = IIf(bActivate, RGSignature & "Missing Scope", vbNullString)
  ToDo14 = IIf(bActivate, WARNING_MSG & "Possible Structure Violation", vbNullString)
  ToDo15 = IIf(bActivate, WARNING_MSG & "Dead Code", vbNullString)

End Sub

Private Sub SubsequentCodeIsDead(ByVal IndentVal As Long)

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'Prepares for Dead Code Recognition

  If CodeIsDead = AlmostInfinity Then
    CodeIsDead = -IndentVal
  End If

End Sub

Private Function TypeIsUndefined(ByVal varName As String, _
                                 ByVal IsConst As Boolean, _
                                 Optional ByVal AnyPosition As Boolean = False) As Boolean

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'Returns True if type of variable or constant is undefined

  If Not AnyPosition Then
    If Not TypeSuffixExists(varName) Then
      TypeIsUndefined = IsConst Or (InStr(DefTypeChars, Left$(varName, 1)) = 0)
    End If
   Else
    TypeIsUndefined = TypeSuffixExistsAnywhere(varName)
  End If

End Function

Public Function TypeSuffixExistsAnywhere(ByVal strSearch As String) As Boolean

  Dim I        As Long

  'Modification to catch these in parameter lists
  'Returns True if StrSearch has type suffix: int% long& single! double# currency@ string$
  For I = 0 To 5 '1 To 6
    If InCode(strSearch, InStr(strSearch, TypeSuffixArray(I))) Then
      TypeSuffixExistsAnywhere = True
      Exit For
    End If
  Next I

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:22:34 AM) 86 + 1735 = 1821 Lines Thanks Ulli for inspiration and lots of code.

