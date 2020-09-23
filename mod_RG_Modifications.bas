Attribute VB_Name = "mod_RG_Modifications"
Option Explicit
'Copyright 2003 Roger Gilchrist
'email: rojagilkrist@hotmail.com
Public BadUserControlTestConducted         As Boolean
Public Const ContMark                      As String = " _"
Public bVeryLargeMsgShow                   As Boolean
Public bInitializing                       As Boolean
Public Enum ETabPage
  TPNone
  TPModules
  TPControls
  TPFile
  TPHelp
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TPNone, TPModules, TPControls, TPFile, TPHelp
#End If
Public Enum CFParray
  FakeProtect
  endDec
  CommentDummy
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private FakeProtect, endDec, CommentDummy
#End If
Public Enum Levels
  LevelOff
  LevelComment
  LevelCommentFix
  LevelFix
  LevelUser
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private LevelOff, LevelComment, LevelCommentFix, LevelFix, LevelUser
#End If
Public Inserted                            As Long
Public Dupl                                As Boolean
Public Const RGSignature                   As String = "'<" & ":-" & ") "
Public Const UsageSign                     As String = vbNewLine & RGSignature & "|"
Public Const SUGGESTION_MSG                As String = RGSignature & ":SUGGESTION: "
Public Const WARNING_MSG                   As String = RGSignature & ":WARNING: "
Public Const EARLYWARNING_MSG              As String = "'<~>" & "<~><~>"
Public Const UPDATED_MSG                   As String = RGSignature & ":UPDATED: "
Public Const PREVIOUSCODE_MSG              As String = RGSignature & ":PREVIOUS CODE : "
Public Const LineContinuationWarning       As String = SUGGESTION_MSG & "Line continuation prevents Auto-Repair. Edit to remove Line continuation characters."
Private Const StrFuncitontoSub             As String = SUGGESTION_MSG & "Function may be changed to Sub if nothing is returned via Function Name."
Public Const strTurnOffMsg                 As String = vbNewLine & _
 "(Stop this message appearing from Settings tab)"
'StrFuncitontoSub is used so that its presence can be tested in AutoFix
Public Const Hash_If_False_Then            As String = "#If False Then"
Public Const Hash_End_If                   As String = "#End If"
Public Const DQuote                        As String = """"
Public Const SQuote                        As String = "'"
Public Const EmptyString                   As String = DQuote & DQuote
Public Const EqualInCode                   As String = " = "
Public Const RBracket                      As String = ")"
Public Const LBracket                      As String = "("
Public Xcheck                              As New cls_ByteCheckBox

Private Sub AMPJoinedStrings(varTest As Variant, _
                             UpDated As Boolean)

  Dim strOffset  As String
  Dim arrTmp     As Variant
  Dim I          As Long
  Dim EqPos      As Long
  Dim strComment As String
  Dim strSpace   As String

  'ver 2.1.3 added lngLineLength test to stop short lines wrapping
  'v2.7.2 reversed test as Len is faster
  If Len(varTest) >= lngLineLength Then
    'v2.6.4test moved here for more speed
    If InStr(varTest, DQuote & " & " & DQuote) Then
      ExtractCode varTest, strComment, strSpace
      If Len(varTest) < lngLineLength Then
        varTest = strSpace & varTest & strComment
       Else
        If InStr(varTest, DQuote & " & " & DQuote) Then
          arrTmp = ExpandForDetection2(varTest)
          For I = LBound(arrTmp) To UBound(arrTmp) - 1
            'If Updated = False Then
            If LenB(strOffset) = 0 Then
              If InStr(varTest, EqualInCode & DQuote) Then
                strOffset = String$(InStr(varTest, EqualInCode & DQuote) - 4 + Len(strSpace), 32)
               Else
                EqPos = InStr(varTest, EqualInCode)
                If EqPos Then
                  If InCode(varTest, EqPos) Then
                    strOffset = String$(EqPos + 2 + Len(strSpace), 32)
                  End If
                End If
              End If
            End If
            If arrTmp(I) = "&" Then
              If Right$(arrTmp(I - 1), 1) = DQuote Then
                If Left$(arrTmp(I + 1), 1) = DQuote Then
                  arrTmp(I) = arrTmp(I) & ContMark & vbNewLine & strOffset
                  UpDated = True
                End If
              End If
            End If
          Next I
          If UpDated Then
            varTest = strSpace & Join(arrTmp) & strComment
          End If
         Else
          varTest = strSpace & varTest & strComment
        End If
      End If
    End If
  End If

End Sub

Public Function AsVariantFix(cMod As CodeModule, _
                             ByVal LIndex As Long, _
                             ByVal varName As String) As String

  Dim FixMissing   As Boolean
  Dim LTypeDEf     As String
  Dim LType        As String
  Dim TypeDef      As String
  Dim RGsign       As String
  Dim MyStr        As String
  Dim I            As Long
  Dim arrVariable  As Variant
  Dim CommentStore As String
  Dim SpaceOffSet  As String
  Dim Guard        As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Rewrites old style Dim x, y, z As Integer which Dimmed all members as integer
  'VB4 and above do not support this. You now get 2 Variants(x,y) and one Integer (z)
  'This is medium risky to use it is triggered by the " As Variant ?' comment
  'It has multiple safety requirements and tends not to hit
  'Message added includes old code for safety
  'NEW also detects and fixes untyped variables in single line multiple Dims where all others have same Type
  On Error GoTo BadError
  MyStr = varName
  MyStr = DuplicateEnumCapitalization(cMod, LIndex, MyStr)
  If varName = MyStr Then
    If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
      If GetLeftBracketPos(MyStr) = 0 Then
        'Safety dont try if any member is dimensioned i.e. 'Square (8,8)' would become 'Square(8 As Type, 8) As Type' if it went into next lines of code
        If InStr(MyStr, ",") Then
          MyStr = ConcealParameterCommas(MyStr)
          If CountSubString(MyStr, " As ") = 1 Then
            arrVariable = Split(MyStr, CommaSpace)
            If CountSubString(arrVariable(UBound(arrVariable)), " As ") = 1 Then
              TypeDef = GetType(arrVariable(UBound(arrVariable)))
              'Mid$(arrVariable(UBound(arrVariable)), Get_As_Pos(TypeDef))
              For I = LBound(arrVariable) To UBound(arrVariable) - 1
                If Get_As_Pos(arrVariable(I)) = 0 Then
                  If Not TypeSuffixExists(arrVariable(I)) Then
                    arrVariable(I) = arrVariable(I) & TypeDef
                  End If
                End If
              Next I
              MyStr = Join(arrVariable, ",")
              RGsign = vbNewLine & _
               WARNING_MSG & "Declaring a whole line with single As Type is no longer supported." & IIf(Xcheck(XPrevCom), vbNewLine & _
               PREVIOUSCODE_MSG & varName & vbNewLine, vbNullString)
            End If
           ElseIf CountSubString(MyStr, " As ") > 1 Then
            'THis fixes any untyped variables if the rest of the variables are Typed the same
            FixMissing = True
            arrVariable = Split(MyStr, ",")
            For I = LBound(arrVariable) To UBound(arrVariable)
              If Get_As_Pos(arrVariable(I)) > 0 Then
                LType = GetType(arrVariable(I))
                If LenB(LTypeDEf) = 0 Then
                  LTypeDEf = LType
                End If
                If LTypeDEf <> LType Then
                  FixMissing = False
                  Exit For
                End If
              End If
            Next I
            If FixMissing Then
              For I = LBound(arrVariable) To UBound(arrVariable)
                If InStr(arrVariable(I), LTypeDEf) = 0 Then
                  arrVariable(I) = arrVariable(I) & LTypeDEf
                End If
              Next I
              MyStr = Join(arrVariable, ",")
              If SpaceOffSet & MyStr <> varName Then
                RGsign = vbNewLine & _
                 WARNING_MSG & "Dimmed un-Typed variable(s) in line where all others had same Type." & IIf(Xcheck(XPrevCom), vbNewLine & _
                 PREVIOUSCODE_MSG & varName & vbNewLine, vbNullString)
               Else
                RGsign = vbNewLine & LineContinuationWarning
              End If
            End If
          End If
        End If
       ElseIf InStr(MyStr, " Function ") Then
        Guard = MyStr
        MyStr = TypeSuffixExtender(MyStr)
        'As Variant ?' also appears on untyped Functions which may not operate as Funtions but Subs
        If Guard <> MyStr Then
          RGsign = vbNewLine & RGSignature & "Function call Typed by AutoFix."
        End If
      End If
      AsVariantFix = SpaceOffSet & MyStr & CommentStore & RGsign
      On Error GoTo 0
    End If
  End If

Exit Function

BadError:
  AsVariantFix = varName

End Function

Public Sub AutoFixWriter(cMod As CodeModule, _
                         ByVal MarkText As String, _
                         ByVal ReWrite As String, _
                         ByVal strTarget As String, _
                         ByVal LIndex As Long, _
                         ResetForAddedLine As Boolean)

  With cMod
    If ReWrite <> strTarget Then
      If InStr(ReWrite, LineContinuationWarning) Or InStr(ReWrite, StrFuncitontoSub) Then
        '(NEW This probably wont hit very often now that concatentLineContinuation is available
        'NEW this allows an additional message on summary notice
        If InStr(ReWrite, LineContinuationWarning) Then
          ResetForAddedLine = False
        End If
        'trap for a problem AutoFixer routines have dealing with
        'line Continuation Character so it displays the
        'routine's message AND Ulli's Message as well and makes no change to code.
        Inserted = Inserted + 1
        .ReplaceLine LIndex, ReWrite & MarkText
       Else
        .ReplaceLine LIndex, ReWrite
      End If
     Else
      'for If AuotFIx routine failed to change ReWrite then Ulli's error message is preserved
      'NEW this allows an additional message on summary notice
      If InStr(ReWrite, LineContinuationWarning) Then
        ResetForAddedLine = False
      End If
      Inserted = Inserted + 1
      .ReplaceLine LIndex, ReWrite & MarkText
    End If
  End With

End Sub

Private Function BadUserControlTest() As Boolean

  Dim I      As Long
  Dim strBad As String

  On Error Resume Next
  'v 2.1.5 added usercontrol test
  'This will block all tool launches
  If Not BadUserControlTestConducted Then
    For I = LBound(ModDesc) To UBound(ModDesc)
      If ModDesc(I).MDType = "UserControl" Then
        strBad = BadInitialization(ModDesc(I).MDProj, ModDesc(I).MDName)
        If InStr(strBad, SQuote & " or " & SQuote) = 0 Then
          BadUserControlTest = True
          BadUserControlTestConducted = True
          HideInitiliser
          mObjDoc.Safe_MsgBox "Code Fixer Tools have been disabled because of the following error" & IIf(CountSubString(strBad, vbNewLine) < 1, "s", "") & " in the UserControl" & strInSQuotes(ModDesc(I).MDName, True) & "." & vbNewLine & _
                    strBad & vbNewLine & _
                    "Please correct " & IIf(CountSubString(strBad, vbNewLine) < 1, "these problems", "this problem") & ". Moving the marked code to 'UserControl_ReadProperties' will usually resolve the problem." & vbNewLine & _
                    "You will have to close and restart VB before Code Fixer can be used.", vbCritical
        End If
      End If
    Next I
   Else
    mObjDoc.Safe_MsgBox "Code Fixer encountered a problem with a UserControl initialising." & vbNewLine & _
                    "You will have to close VB and restart the project to fully re-initialise the UserControls.", vbInformation
    BadUserControlTest = True
  End If
  On Error GoTo 0

End Function

Private Function CodeSplit(ByVal varTest As Variant, _
                           ByVal strSep As String) As Variant

  Dim StrDisguise As String

  Do
    StrDisguise = RandomString(48, 122, 3, 6)
  Loop While InStr(varTest, StrDisguise)
  DisguiseLiteral varTest, strSep, True
  CodeSplit = Split(varTest, ",")
  varTest = Join(CodeSplit, StrDisguise)
  DisguiseLiteral varTest, strSep, False
  CodeSplit = Split(varTest, StrDisguise)

End Function

Private Sub DoBlankPreserveCleanUp(cMod As CodeModule, _
                                   arrModule As Variant, _
                                   UpDated As Boolean)

  Dim I As Long

  For I = 0 To UBound(arrModule)
    If Xcheck(XSpaceSep) Then
      If Trim$(arrModule(I)) = Trim$(RGSignature) Then
        If I < UBound(arrModule) Then
          If arrModule(I + 1) = vbNullString Then
            arrModule(I) = vbNullString
            UpDated = True
          End If
        End If
       ElseIf Trim$(arrModule(I)) = vbNullString Then
        arrModule(I) = RGSignature
      End If
     Else '
      If Trim$(arrModule(I)) = Trim$(RGSignature) Then
        arrModule(I) = vbNullString
        UpDated = True
      End If
    End If
  Next I
  If Xcheck(XSpaceSep) Then
    ReWriter cMod, arrModule, RWModule, True
    arrModule = GetModuleArray(cMod, True, False)
    For I = 0 To UBound(arrModule)
      If Trim$(arrModule(I)) = Trim$(RGSignature) Then
        arrModule(I) = vbNullString
        UpDated = True
      End If
    Next I
  End If

End Sub

Public Function DuplicateEnumCapitalization(cMod As CodeModule, _
                                            ByVal LIndex As Long, _
                                            ByVal varName As String) As String

  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim MyStr        As String
  Dim I            As Long
  Dim UlliCount    As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ' This is just to to preserve the Enum Capitalization Trick
  ' The routine removes the As Variant comment that The trick gets from Ulli's code
  ' and decreases the Inserted count to reflect this
  On Error GoTo BadError
  MyStr = varName
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    If DuplicateEnumCapitalizationTest(cMod, LIndex) Then
      Select Case LeftWord(MyStr)
       Case "Dim"
        MyStr = Safe_Replace(MyStr, "Dim ", "Private ", , 1)
       Case "Public"
        MyStr = Safe_Replace(MyStr, "Public ", "Private ", , 1)
      End Select
      '*if it is the trick it may have picked up another Ulli comment so delete that
      UlliCount = CountSubString(CommentStore, Grumpy & Chr$(160) & "As Variant ?")
      If UlliCount Then
        For I = 1 To UlliCount
          Inserted = Inserted - 1
        Next I
      End If
      CommentStore = Replace$(CommentStore, Grumpy & Chr$(160) & "As Variant ?", vbNullString)
      CommentStore = Replace$(CommentStore, Grumpy & Chr$(160) & "Duplicate Name", vbNullString)
      DuplicateEnumCapitalization = SpaceOffSet & MyStr & CommentStore & SngSpace
     Else
      DuplicateEnumCapitalization = SpaceOffSet & MyStr & CommentStore
    End If
  End If

Exit Function

BadError:
  DuplicateEnumCapitalization = varName

End Function

Public Function DuplicateEnumCapitalizationTest(cMod As CodeModule, _
                                                ByVal LIndex As Long) As Boolean

  Dim strNext     As String
  Dim StrPrev     As String
  Dim L_Startline As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ' support code for DuplicateEnumCapitalization
  With cMod
    L_Startline = LIndex
    If L_Startline > 2 Then
      Do
        StrPrev = Trim$(.Lines(L_Startline, 1))
        L_Startline = L_Startline - 1
      Loop While Left$(StrPrev, 8) = "Private " And LIndex > 2
      L_Startline = LIndex
      Do
        strNext = Trim$(.Lines(L_Startline, 1))
        L_Startline = L_Startline + 1
      Loop While Left$(strNext, 8) = "Private " And LIndex < .CountOfDeclarationLines + 2
      If SmartLeft(StrPrev, Hash_If_False_Then) Then
        If SmartLeft(strNext, Hash_End_If) Then
          DuplicateEnumCapitalizationTest = True
          Dupl = False
        End If
      End If
    End If
  End With

End Function

Public Function FileSize(ByVal filespec As String) As Single

  'returns file size in KB

  If LenB(filespec) Then
    FileSize = Round(FSO.GetFile(filespec).Size / 1024, 1)
  End If

End Function

Private Sub FillLsvModNames()

  Dim Smsg     As String
  Dim strClass As String
  Dim XX       As Long
  Dim strMod   As String
  Dim strBas   As String

  HideColumnsReturnPosition frm_FindSettings.lsvModNames
  DoEvents
  Generate_ModuleArray False
  Generate_ProjectArray False
  DoFrontPageListView Smsg
  For XX = LBound(ModDesc) To UBound(ModDesc)
    With ModDesc(XX)
      If .MDType = "Class" Then
        'MIGHT REMOVE THIS ^ AT SOME TIME
        strClass = AccumulatorString(strClass, .MDName, , True)
      End If
      If .MDType = "Module" Then
        strBas = AccumulatorString(strBas, .MDName, , True)
      End If
      strMod = AccumulatorString(strMod, .MDName, , True)
    End With 'ModDesc(XX)
  Next XX
  FillArray QSortModNameArray, strMod, , , True
  FillArray QSortModClassArray, strClass, , , True
  FillArray QSortModBasArray, strBas, , , True
  If bVeryLargeMsgShow Then
    If Not Xcheck(XNoLargeFileMsg) Then
      If Len(Smsg) Then
        'ver 1.1.27 thanks to rblanch who pointed out the problem that lead to this
        Smsg = Smsg & IIf(CountSubString(Smsg, vbNewLine) > 1, "are", "is") & " above the recommended size for a single VB source file(64KB) " & vbNewLine & _
         "If you get 'Out of Memory (7)' errors it is recommended that you turn off some sections" & vbNewLine & _
         "(Restructure, Parameters, Suggest & Unused) on first run of Code Fixer." & vbNewLine & _
         "NOTE: Size of Forms and other control bearing sources may safely exceed this limit in most cases." & strTurnOffMsg
        mObjDoc.Safe_MsgBox Smsg, vbInformation
        bVeryLargeMsgShow = False
      End If
    End If
  End If
  'ver 1.1.67 this sets or hides unneeded columns on the find tool

End Sub

Private Sub FormatArrayStructures(varTest As Variant, _
                                  UpDated As Boolean)

  Dim arrTmp            As Variant
  Dim strFull           As String
  Dim I                 As Long
  Dim strTmp            As String
  Dim strOffset         As String
  Dim strTrigger        As String
  Dim strEmptyMember    As String
  Dim strComment        As String
  Dim strSpace          As String
  Dim minMembersPerLine As Long
  Dim memCount          As Long

  'Ver 1.1.84 replace old slower version
  'ver 2.1.3 added lngLineLength test to stop short lines wrapping
  'v2.6.3 faster test
  If Len(varTest) >= lngLineLength Then
    If InStr(varTest, "= Array(") Then
      ExtractCode varTest, strComment, strSpace
      If Len(varTest) < lngLineLength Then
        varTest = strSpace & varTest & strComment
       Else
        If InstrAtPosition(varTest, "= Array(", ipAny, False) Then
          strTrigger = CommaSpace
          'arrTmp = CodeSplit(varTest, ",")
          arrTmp = CodeSplit(varTest, ",")
          'v2.4.7 Thanks Mohammad Alian Nejadi
          'this helps with very large arrays which may generate too many line continuations
          'if length alone is used to break lines
          minMembersPerLine = (UBound(arrTmp) - 1) / 24
          For I = LBound(arrTmp) To UBound(arrTmp)
            If LenB(Trim$(arrTmp(I))) = 0 Then
              strEmptyMember = WARNING_MSG & " There is an empty element in this Array code."
              Exit For
            End If
          Next I
          strOffset = String$(InStr(arrTmp(0), "= Array(") + 6 + Len(strSpace), 32)
          For I = LBound(arrTmp) To UBound(arrTmp)
            memCount = memCount + 1
            If Len(strTmp & strTrigger & arrTmp(I)) < lngLineLength + IIf(Len(strFull), 0, Len(arrTmp(0))) Or memCount <= minMembersPerLine Then
              ''v2.4.7 added Len(strFull) bit to keep whole array formatted more neatly and memCount for very large arrays
              'in small arrays the memCount test will be passed early so the length test will determine the length
              'in Large arrays the length test will be passed but line will keep building until memCount kicks in
              strTmp = strTmp & IIf(Len(strTmp), strTrigger, vbNullString) & Trim$(arrTmp(I))
             Else
              If Len(strTmp) >= lngLineLength Then
                strFull = strFull & IIf(Len(strFull), ContMark & vbNewLine & _
                 strOffset, vbNullString) & strTmp & strTrigger
                strTmp = vbNullString
                memCount = 0
              End If
              strTmp = strTmp & IIf(Len(strTmp), strTrigger, vbNullString) & arrTmp(I)
            End If
          Next I
          If Len(strTmp) Then
            strFull = strFull & IIf(Len(strFull), ContMark & vbNewLine & strOffset, vbNullString) & strTmp
            strTmp = vbNullString
            memCount = 0
          End If
          If LenB(strEmptyMember) Then
            strFull = strFull & vbNewLine & strEmptyMember
          End If
          varTest = strSpace & strFull & strComment
          UpDated = True
         Else
          varTest = strSpace & varTest & strComment
        End If
      End If
    End If
  End If

End Sub

Private Sub FormatStringStructures(varTest As Variant, _
                                   ByVal strTrigger As String, _
                                   UpDated As Boolean)

  Dim arrTmp     As Variant
  Dim strFull    As String
  Dim I          As Long
  Dim strOffset  As String
  Dim EqPos      As Long
  Dim strComment As String
  Dim strSpace   As String

  'Ver 1.1.84 replace old slower version
  'ver 2.0.1 Thanks to Randy Giese
  'complete rewrite to make it behave properly
  'when there is more than one trigger in a row
  'ver 2.0.3 stops it touching minor strings using standard VB procedures)
  'ver 2.1.3 added lngLineLength test to stop short lines wrapping
  If Len(varTest) > -lngLineLength Then
    ExtractCode varTest, strComment, strSpace
    If Len(varTest) < lngLineLength Then
      varTest = strSpace & varTest & strComment
     Else
      If Not InstrAtPosition(varTest, strTrigger, ipAny, False) Then
        varTest = strSpace & varTest & strComment
        GoTo SafeExit
       Else
        EqPos = InStr(varTest, EqualInCode)
        If EqPos Then
          If InCode(varTest, EqPos) Then
            If isRefLibVBCommands(WordAfter(ExpandForDetection(varTest), "=")) Then
              GoTo SafeExit
            End If
          End If
        End If
      End If
      If InstrAtPosition(varTest, strTrigger, ipAny, False) Then
        DisguiseLiteral varTest, strTrigger, True
        arrTmp = Split(varTest, strTrigger)
        For I = LBound(arrTmp) To UBound(arrTmp) - 1
          arrTmp(I) = arrTmp(I) & strTrigger
        Next I
        'Gather Tab offsets for use on new lines
        strOffset = strSpace
        'for code of form ' x = "some string" & vbnewline & "more string" '
        'this lines up the initial quote of each line with the first quote mark
        If InStr(arrTmp(0), EqualInCode & DQuote) Then
          strOffset = strOffset & String$(InStr(arrTmp(0), EqualInCode & DQuote) + 1, 32)
        End If
        If InStr(arrTmp(0), "MsgBox") Then
          If InCode(varTest, InStr(varTest, "MsgBox(")) Then
            strOffset = String$(InStr(arrTmp(0), "MsgBox") + 5, 32)
           ElseIf InCode(varTest, InStr(varTest, "MsgBox")) Then
            strOffset = String$(InStr(arrTmp(0), "MsgBox") + 5, 32)
          End If
        End If
        'v 2.2.6 deal with embedded spaces
        For I = LBound(arrTmp) To UBound(arrTmp)
          If InStr(arrTmp(I), ContMark & vbNewLine & " ") Then
            Do While InStr(arrTmp(I), ContMark & vbNewLine & " ")
              arrTmp(I) = Replace$(arrTmp(I), ContMark & vbNewLine & " ", ContMark & vbNewLine)
            Loop
            arrTmp(I) = Replace$(arrTmp(I), ContMark & vbNewLine, ContMark & vbNewLine & strOffset & " ")
          End If
        Next I
        'Insert line cont marks
        If UBound(arrTmp) > 0 Then
          For I = LBound(arrTmp) To UBound(arrTmp) - 1
            arrTmp(I) = arrTmp(I) & ContMark & vbNewLine
          Next I
        End If
        For I = 1 To UBound(arrTmp)
          If Len(arrTmp(I)) Then
            'v 2.2.6 LTrim copes with lines that have internal spacing
            arrTmp(I) = strOffset & arrTmp(I)
          End If
        Next I
        For I = LBound(arrTmp) To UBound(arrTmp)
          strFull = strFull & arrTmp(I)
        Next I
        UpDated = InStr(strFull, ContMark)
        varTest = strSpace & strFull & strComment
        DisguiseLiteral varTest, strTrigger, False
      End If
    End If
  End If
SafeExit:

End Sub

Private Sub FormatVBStructures(varTest As Variant, _
                               ByVal strTrigger As String, _
                               UpDated As Boolean)

  Dim SpaceOffSet As Long
  Dim CommaPos    As Long

  'This routine can cope with (avoid) routine headers and Delcare lines embedded as literal strings.
  'Thanks Ulli for proving this was needed with the CodeProfiler code
  If InstrAtPosition(varTest, strTrigger, IpLeft, True) Then
    If Not InStr(varTest, ContMark) Then
      If CountSubString(varTest, CommaSpace) Then
        SpaceOffSet = GetLeftBracketPos(varTest)
        CommaPos = GetCommaSpacePos(varTest)
        Do While CommaPos
          If InCode(varTest, CommaPos) Then
            varTest = Left$(varTest, CommaPos) & ContMark & vbNewLine & _
             Space$(SpaceOffSet) & Mid$(varTest, CommaPos + 2)
            UpDated = True
          End If
          CommaPos = GetCommaSpacePos(varTest, CommaPos + 2 + SpaceOffSet)
        Loop
      End If
    End If
  End If

End Sub

Public Sub FormCaptionDisplay(Optional StrCap As String = vbNullString)

  frm_CodeFixer.Caption = AppDetails & IIf(Len(StrCap), "...", vbNullString) & StrCap

End Sub

Public Function ForVariableInsert(ByVal varName As String, _
                                  ByVal MarkText As String) As String

  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim MyStr        As String
  Dim NewVariable  As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Apply correct variable to Next in For...Next structures
  'NEW if there is a comment which contains the correct variable don't add it
  'because this is actually slightly faster usage while still leaving the code easier to read
  On Error GoTo BadError
  MyStr = varName
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    NewVariable = Mid$(MarkText, InStr(MarkText, "Variable:") + Len("Variable:"))
    If InstrAtPosition(LCase$(CommentStore), LCase$(LTrim$(NewVariable)), ipAny) = 0 Then
      Select Case FixData(NForNextVar).FixLevel
       Case CommentOnly
        ForVariableInsert = SpaceOffSet & MyStr & "' " & NewVariable & CommentStore & RGSignature & "Insert For-Variable?"
        AddNfix NForNextVar
       Case FixAndComment
        ForVariableInsert = SpaceOffSet & MyStr & SngSpace & NewVariable & CommentStore & RGSignature & "For-Variable Inserted"
        AddNfix NForNextVar
       Case JustFix
        ForVariableInsert = SpaceOffSet & MyStr & "' " & NewVariable & CommentStore
        AddNfix NForNextVar
      End Select
     Else
      ForVariableInsert = SpaceOffSet & MyStr & CommentStore & SngSpace
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  ForVariableInsert = varName

End Function

Public Function GetComponentCount() As Long

  Dim I       As Long

  On Error Resume Next
  For I = 1 To VBInstance.VBProjects.Count
    GetComponentCount = GetComponentCount + VBInstance.VBProjects(I).VBComponents.Count
  Next I
  On Error GoTo 0

End Function

Public Function GetModuleArray(cMod As CodeModule, _
                               Optional ByVal DeleteBlanks As Boolean = True, _
                               Optional ByVal TrimSpace As Boolean = True) As Variant

  Dim Tmpstring1 As String
  Dim I          As Long
  Dim Cline      As String

  'Gets code (including Declarations) each member is one line of code
  ' if DeleteBlanks then don't include blank lines
  'This makes it easier for most routines in tis program to deal with code
  'set to False only if a routine calls after Indenting( which includes line spacing) has been done
  With cMod
    For I = 1 To .CountOfLines
      If TrimSpace Then
        Cline = Trim$(.Lines(I, 1))
       Else
        Cline = .Lines(I, 1)
      End If
      If LenB(Cline) Or Not DeleteBlanks Then
        Tmpstring1 = Tmpstring1 & Cline & vbNewLine
      End If
    Next I
  End With
  If LenB(Tmpstring1) Then
    GetModuleArray = Split(Left$(Tmpstring1, Len(Tmpstring1) - Len(vbNewLine)), vbNewLine)
   Else
    GetModuleArray = Split("")
  End If
  Tmpstring1 = vbNullString

End Function

Private Function GetSafeCutPoint(ByVal strA As String, _
                                 ByVal BasePos As Long) As Long

  Do
    GetSafeCutPoint = GetSpacePos(strA, BasePos)
    BasePos = BasePos + 1
    'v2.64
  Loop Until InCode(strA, GetSafeCutPoint) Or BasePos >= Len(strA) Or GetSafeCutPoint = 0

End Function

Public Sub HideFraPage()

  Dim I As Long

  With frm_CodeFixer
    .Visible = False
    For I = 1 To .frapage.Count
      If I <> TPControls Then
        .frapage(I).BorderStyle = 0
        .frapage(I).Visible = False
        .frapage(I).Caption = vbNullString
      End If
    Next I
  End With

End Sub

Public Function JustACommentOrBlank(ByVal varSearch As Variant) As Boolean

  'copright 2003 Roger Gilchrist
  'detect comments and empty strings
  'TestLineSuspension varSearch
  'v2.8.3 speed up with inline testing

  varSearch = Trim$(varSearch)
  If LenB(varSearch) = 0 Then
    JustACommentOrBlank = True
   ElseIf Left$(varSearch, 1) = SQuote Then
    JustACommentOrBlank = True
   ElseIf Left$(varSearch, 4) = "Rem " Then
    JustACommentOrBlank = True
  End If
  '  safe_sleep

End Function

Public Function LaunchTool(Optional bShowForm As Boolean = True) As Boolean

  Dim MyhourGlass     As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  'v 2.2.1 Thanks rblanch. this is a refinement of CF behaviour due to the bug with damaged frx files
  'also applies to UserControl errors. This traps any problems on previous run that require VB restart
  'before CF has gone too far into processing code
  If bForceReload Then
    mObjDoc.Safe_MsgBox "Apologies." & vbNewLine & _
                    "Due to a problem encountered on previous run Code Fixer requires that VB be closed and restarted in order to operate.", vbCritical
   bForceReload = False
   LaunchTool = False
   Else
    SetUpPreLoad
    'Force all modules to open so that Code Fixer can get at them
    FillLsvModNames
    If Not BadUserControlTest Then
      LaunchTool = True
      With frm_CodeFixer
        '.Visible = bShowForm
        .Caption = AppDetails & " Initialising..."
        mObjDoc.GraphCaption "Initialising..."
        PageLaunch bShowForm
      End With
      With frm_CodeFixer
        SetTopMost frm_CodeFixer, True
        SetupDataArrays
        If Not Xcheck(XNoPrjWarning) Then
          If LenB(GetActiveProject.FileName) = 0 Then
            mObjDoc.Safe_MsgBox "You have loaded a VB file but not a project OR you have not saved the project." & vbNewLine & _
                    "The Declaration and Procedure Scope tests and All Unused Tools will not function." & vbNewLine & _
                    "This protects your code form unnecessary changes in Scope while allowing you to apply all other fixes." & vbNewLine & _
                    "Check that Missing Dims inserted are not Public Variables form other parts of the module's home project before accepting them." & vbNewLine & _
                    "If you have not saved the project you cannot Reload, Backup or Recover", vbInformation + vbOKOnly
          End If
        End If
        .Caption = AppDetails
      End With
    End If
  End If

End Function

Public Sub LineContinuationFix(ModArray As Variant, _
                               UpDated As Boolean, _
                               Optional ModuleNumber As Long = 0, _
                               Optional bForce As Boolean = False)

  Dim arrTmp    As Variant
  Dim J         As Long
  Dim OldLen    As Long
  Dim I         As Long
  Dim Jump      As Long
  Dim MaxFactor As Long

  MaxFactor = UBound(ModArray)
  For I = 0 To MaxFactor
    Jump = 1
    If I Mod 10 = 0 Or I = MaxFactor Then
      MemberMessage "", I, MaxFactor
    End If
    If Not Xcheck(XNoIndentCom) Then
      'Hi Randy thanks for suggesting this
      'it was disgustingly easy once I looked for it
      'ver 2.0.1 Thanks to Randy Giese (again) this was the cause of the
      'blank line comments not deleting after I updated the comment marker to include a space
      If JustACommentOrBlank(Trim$(ModArray(I))) Then
        ModArray(I) = Trim$(ModArray(I))
      End If
    End If
    'collect extremely long line continuation strings which were not fully de-line continued previously
    If JustACommentOrBlank(ModArray(I)) Then
      'FIXME check following is valid
      'because a line continuation comment may have been wrapped up into a too long line
      If InStr(ModArray(I), ContMark) = 0 Then
        LineContinuationForVeryLongLines ModArray(I), ContMark & vbNewLine, UpDated
      End If
     Else
      If HasLineCont(ModArray(I)) Then
        Do While HasLineCont(ModArray(I))
          If LenB(ModArray(I + Jump)) Then
            'Remove any offset spaces
            ModArray(I) = Left$(ModArray(I), Len(ModArray(I)) - 1) & Trim$(ModArray(I + Jump))
            ModArray(I + Jump) = vbNullString
          End If
          Jump = Jump + 1
        Loop
      End If
      'Do Array formatting
      If dofix(ModuleNumber, FormatArrayLineContinuation) Or bForce Then
        FormatArrayStructures ModArray(I), UpDated
      End If
      'Do string formatting
      If dofix(ModuleNumber, FormatStringLineContinuation) Or bForce Then
        AMPJoinedStrings ModArray(I), UpDated
        ReDoLineStringFormating ModArray(I), UpDated
      End If
      ReDoLineParameters ModuleNumber, ModArray(I), UpDated
      If Jump > 1 Then
        'v2.6.4
        If InStr(ModArray(I), ContMark) = 0 Then
          LineContinuationForVeryLongLines ModArray(I), ContMark & vbNewLine, UpDated
        End If
      End If
      I = I + Jump - 1
      LongLineWithEndComment ModArray(I), UpDated
      'v2.4.4 cope with line too long for VB after indenting
      ' Debug.Print ModArray(7)
      If Len(ModArray(I)) > 1032 Then
        'v.2.4.7 better version
        arrTmp = Split(ModArray(I), vbNewLine)
        For J = LBound(arrTmp) To UBound(arrTmp)
          If Len(arrTmp(J)) > 1032 Then
            UpDated = True
            OldLen = Len(arrTmp(J))
            arrTmp(J) = Trim$(arrTmp(J))
            If Len(arrTmp(J)) < OldLen Then
              Do While Len(arrTmp(J)) < 1023
                arrTmp(J) = " " & arrTmp(J)
              Loop
            End If
          End If
        Next J
        ModArray(I) = Join(arrTmp, vbNewLine)
        'ModArray(I) = WARNING_MSG & "Following line is too long to be indented properly" & vbNewLine & SUGGESTION_MSG & " you should break the line if at all possible" & vbNewLine & ModArray(I)
      End If
    End If
  Next I
  'v2.5.8 added to deal with too mmany line continuation characters
  'thanks to Alberto Torres Klinger for using a code line that triggered this bug
  'v2.6.4 CountSubStringCode and Safe_Replace used to avoid code in strings
  For I = 0 To MaxFactor
    If LenB(ModArray(I)) Then
      'v2.9.6 problem with fancy header comments this just ignores them. Thanks Ian K
      If Not JustACommentOrBlank(Trim$(ModArray(I))) Then
        If CountSubStringCode(ModArray(I), " _") > 25 Then
          Do Until CountSubStringCode(ModArray(I), ContMark) < 25
            ModArray(I) = Safe_Replace(ModArray(I), ContMark & vbNewLine, " ", , 1)
          Loop
        End If
      End If
    End If
  Next I

End Sub

Public Sub LineContinuationForVeryLongLines(VarLongLine As Variant, _
                                            ByVal sep As String, _
                                            UpDated As Boolean)

  Dim strComment   As String
  Dim strTemp      As String
  Dim strVeryLong  As String
  Dim LngCutPoint  As Long
  Dim LinContCount As Long
  Dim strSpace     As String

  'Copes with very long long lines by inserting designated separators( either newline or line continuation characters
  'with line continuation characters there is a VB limit of 25
  '*if this is reached you're in trouble but this is very unlikely.
  If LenB(VarLongLine) > 1023 Then
    'get initial cut point
    'v2.1.3 detact comment part of long line, if any
    ExtractCode VarLongLine, strComment, strSpace
    ' initCut = lngLineLength
    strTemp = VarLongLine
    LngCutPoint = GetSafeCutPoint(strTemp, lngLineLength)
    If LngCutPoint > 0 Then
      strVeryLong = Mid$(strTemp, LngCutPoint + 1)
      strTemp = Left$(strTemp, LngCutPoint - 1)
      Do While LenB(strVeryLong)
        LngCutPoint = GetSafeCutPoint(strVeryLong, lngLineLength)
        If LngCutPoint = 0 Then
          If LenB(strVeryLong) > 0 Then
            strTemp = strTemp & sep & strVeryLong
            UpDated = True
            Exit Do
          End If
         Else
          'v2.8.9 this solved a problem with very long string Const
          strTemp = strTemp & sep & Left$(strVeryLong, LngCutPoint) ' - 1)
          strVeryLong = Mid$(strVeryLong, LngCutPoint + 1)
          UpDated = True
          'ver 2.0.2 Thanks to Ing. Miguel Angel Guzman Robles
          'Bug damaged strings if the very last piece of strVeryLong was a single quote mark
          'this eliminates that
          If Len(strVeryLong) < 10 Then
            strTemp = strTemp & strVeryLong
            strVeryLong = vbNullString
          End If
          If InStr(sep, ContMark) Then
            LinContCount = LinContCount + 1
            If LinContCount = 24 Then
              'upperlimit is 25 .
              'This is set to 24 so that the last one can be used to clean up and exit.
              'The last line may  be excessively long and cause a Structural Failure.
              mObjDoc.Safe_MsgBox "Too many Line Continuation Characters needed by current line", vbCritical
              strTemp = strTemp & sep & strVeryLong
              Exit Do
            End If
          End If
        End If
      Loop
    End If
    If strTemp <> strComment Then
      VarLongLine = strSpace & strTemp & IIf(LenB(strComment), vbNewLine & strComment, vbNullString)
     Else
      VarLongLine = strSpace & strTemp
    End If
  End If

End Sub

Private Sub LongLineWithEndComment(varTest As Variant, _
                                   UpDated As Boolean)

  Dim strComment As String
  Dim strSpace   As String

  'v2.1.3 if line is only long because of an eol comment move comment to next line
  If Len(varTest) >= lngLineLength Then
    ExtractCode varTest, strComment, strSpace
    If Len(strComment) + Len(strSpace) > 0 And UCase$(strComment) <> strComment Then
      'don't detach the strucutral comments that ulli's code inserts
      If Len(varTest) < lngLineLength Then
        varTest = strSpace & varTest & vbNewLine & _
         IIf(Xcheck(XNoIndentCom), strSpace, vbNullString) & strComment
        UpDated = True
       Else
        varTest = strSpace & varTest & strComment
      End If
     Else
      varTest = strSpace & varTest & strComment
    End If
  End If

End Sub

Public Sub MakeReadable(ByVal strFileName As String)

  Dim F As File

  'this routine resets the Read-Only Attribute to Normal
  Set F = FSO.GetFile(strFileName)
  If F.Attributes And ReadOnly Then
    F.Attributes = F.Attributes - vbReadOnly
  End If
  If F.Attributes And Hidden Then
    F.Attributes = F.Attributes - Hidden
  End If

End Sub

Public Sub MakeVBFilesReadable(ByVal lngIndex As Long)

  Dim Comp As VBComponent
  Dim I    As Long

  ' a component file may have support files with the same
  ' filename and diferrent extentions eg(FRM has FRX)
  ' The IDE directly editable file is always .FileNames(1)
  ' .FileCount is the number of supporting files
  Set Comp = VBInstance.VBProjects(ModDesc(lngIndex).MDProj).VBComponents(ModDesc(lngIndex).MDName)
  For I = 1 To Comp.FileCount
    MakeReadable Comp.FileNames(I)
  Next I

End Sub

Public Sub ReDoLineContinuation(cMod As CodeModule, _
                                Optional ByVal ForceAbort As Boolean = False)

  Dim ModuleNumber As Long
  Dim ModArray     As Variant
  Dim UpDated      As Boolean

  On Error GoTo BugHit
  If Not bAborting Then
    If Not ForceAbort Then
      'Most of this program relies on there being no line continuation characters for easy of operations
      'but for readability they are often useful so this routine (re)inserts the ' _' character in
      'several places as described above each section
      ModuleNumber = ModDescMember(cMod.Parent.Name)
      ModArray = GetModuleArray(cMod, False, False)
      'Offset literal strings with '&' joins one per line
      If UBound(ModArray) > -1 Then
        LineContinuationFix ModArray, UpDated, ModuleNumber
        If Not Xcheck(XSpaceSep) Then
          'this deletes Code Fixer inserted spaces
          'and leaves only the original spaces
          'ver 2.0.1 Thanks to Randy Giese who pointed out that this wasn't working properly
          ReWriter cMod, ModArray, RWModule, True
          ModArray = GetModuleArray(cMod, False, False)
        End If
        If Xcheck(XBlankPreserve) Then
          DoBlankPreserveCleanUp cMod, ModArray, UpDated
        End If
        If UpDated Then
          ReWriter cMod, ModArray, RWModule, False
        End If
      End If
    End If
  End If
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "ReDoLineContinuation"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub ReDoLineParameters(ByVal ModuleNumber As Long, _
                               varCode As Variant, _
                               UpDated As Boolean)

  Dim J As Long
  Dim K As Long

  ' Do declare parameters one per line
  If dofix(ModuleNumber, FormatDeclareLineContinuation) Then
    'separate trigger so seperate run
    If InStr(varCode, "Declare ") Then
      If InCode(varCode, InStr(varCode, "Declare ")) Then
        For J = LBound(ArrAllScopes) To UBound(ArrAllScopes)
          FormatVBStructures varCode, ArrAllScopes(J) & " Declare", UpDated
        Next J
      End If
    End If
  End If
  'Do routine parameters one per line
  If dofix(ModuleNumber, FormatRoutineLineContinuation) Then
    '*if you saw earlier version this structure protects
    'this from modifying itself because the detect tests in FormatVBStructures
    'are not robust enough
    If isProcHead(varCode) Then
      For J = LBound(ArrAllScopes) To UBound(ArrAllScopes)
        For K = LBound(ArrFuncPropSub) To UBound(ArrFuncPropSub)
          FormatVBStructures varCode, ArrAllScopes(J) & SngSpace & ArrFuncPropSub(K), UpDated
        Next K
      Next J
    End If
  End If

End Sub

Private Sub ReDoLineStringFormating(varCode As Variant, _
                                    UpDated As Boolean)

  Dim OldVar As String

  OldVar = varCode
  'ver 2.1.3 added lngLineLength test to stop short lines wrapping
  If Len(varCode) >= lngLineLength Then
    'V2.2.6 added test to skip if not applicable
    If InstrArray(varCode, " vbNewLine &", " vbCrLf &", " vbCr &") Then
      If CountSubString(varCode, " vbNewLine &") < CountSubString(varCode, " vbCrLf &") Then
        FormatStringStructures varCode, " vbCrlf &", UpDated
       Else
        FormatStringStructures varCode, " vbNewLine &", UpDated
      End If
      If OldVar = varCode Then
        If CountSubString(varCode, " vbNewLine &") < CountSubString(varCode, " vbCrLf &") Then
          FormatStringStructures varCode, " vbNewLine &", UpDated
         Else
          FormatStringStructures varCode, " vbCrlf &", UpDated
        End If
      End If
      If OldVar = varCode Then
        FormatStringStructures varCode, " vbCr &", UpDated
      End If
    End If
  End If

End Sub

Public Function RepeatedString(ByVal strBase As String, _
                               reps As Long) As String

  Dim I As Long

  If reps < 1 Then
    reps = 1
  End If
  For I = 1 To reps
    RepeatedString = RepeatedString & strBase
  Next I

End Function

Public Function ReservedWordAs_Type_or_Enum_Member(ByVal strSearch As String) As Boolean

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ' At present this is only used to protect the MicroSoft approved and
  'VB legal but dangerous to the Code formatter 'Type as Long' members of Type and Enum definitions.
  'But could be expanded to issue warnings about any other similar structures
  'which the IDE allows simply by adding them to the InstrAtPositionArray line
  'Because VB IDE recognises this sort of Type member it may be in a variety of cases so uses LCase
  'and LCase versions of Reserved words for safety (ByRef stops this feeding back

  If InstrAtPositionArray(strSearch, IpMiddle, True, "As", vbTab & "As") Then
    If InstrAtPosition(LCase$(strSearch), "type", IpLeft) Then
      ReservedWordAs_Type_or_Enum_Member = True
    End If
  End If

End Function

Public Sub SetupDataArrays()

  ComponentCount = GetComponentCount
  ReDim PrivateTypeArray(ComponentCount) As Variant
  ReDim IsCollectionClass(ComponentCount + 1) As Boolean
  ReDim Attributes(ComponentCount + 1) As Variant
  ReDim ArrQDefLng(ComponentCount + 1) As Variant
  ReDim ArrQDefSng(ComponentCount + 1) As Variant
  ReDim ArrQDefInt(ComponentCount + 1) As Variant
  ReDim ArrQDefBool(ComponentCount + 1) As Variant
  ReDim ArrQDefByte(ComponentCount + 1) As Variant
  ReDim ArrQDefCur(ComponentCount + 1) As Variant
  ReDim ArrQDefDbl(ComponentCount + 1) As Variant
  ReDim ArrQDefDate(ComponentCount + 1) As Variant
  ReDim ArrQDefStr(ComponentCount + 1) As Variant
  ReDim ArrQDefObj(ComponentCount + 1) As Variant
  ReDim ArrQDefVar(ComponentCount + 1) As Variant
  ArrQCheckedVariables = Array(vbNullString)

End Sub

Private Sub SetUpPreLoad()

  With frm_CodeFixer
    HideFraPage
  End With
  SetUpSettingFrame
  On Error GoTo 0

End Sub

Public Sub SizeToFrame()

  With frm_CodeFixer
    'v2.7.3 if you minimize then bring VB back up this would crash without the error trap
    On Error Resume Next
    .Width = .frapage(FrameActive).Width
    .Height = .frapage(FrameActive).Height
    .frapage(FrameActive).Move 0, 0
    .frapage(FrameActive).Visible = True
    .Visible = True
    .Refresh
  End With
  DoEvents
  On Error GoTo 0

End Sub

Public Function TypeSuffixExtender(ByVal varName As String, _
                                   Optional SufNo As Long = -1) As String

  
  Dim BracketPos   As Long
  Dim TSPos        As Long
  Dim I            As Long
  Dim ConstValue   As String
  Dim MyStr        As String
  Dim CommentStore As String
  Dim SpaceOffSet  As String
  Dim numcut       As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'UPDATE 13 Jan 2003 total rewrite; no longer needs supporting routines
  On Error GoTo BadError
  MyStr = varName
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    '*If constant then detach the value so that it can be reattached later
    'otherwise & in string values gets 'updated and destroyed
    If InStr(MyStr, "Const") Then
      If InCode(MyStr, InStr(MyStr, "Const")) Then
        ConstValue = Trim$(Mid$(MyStr, InStr(MyStr, " =")))
        MyStr = Left$(MyStr, InStr(MyStr, " =") - 1)
        'v2.9.9 Dont double up if the type suffix must remain
        'E.G. Private Const IniPart2       As Currency = 4023233417@"
        'where @ has to be there to stop it changing to #
        If InStr(MyStr, " As ") = 0 Then
          For I = 0 To 5
            TSPos = InStr(ConstValue, TypeSuffixArray(I))
            If TSPos = Len(ConstValue) Then
              MyStr = MyStr & " As " & AsTypeArray(I) & Left$(ConstValue, Len(ConstValue) - 1)
              If InStr(MyStr, " As " & AsTypeArray(I) & " As " & AsTypeArray(I)) Then
                'v2.2.2 Thanks Mike Ulik
                'trap for updating 'Const HWND_BOTTOM As Long = 1&' to 'Const HWND_BOTTOM As Long As Long = 1'
                MyStr = Replace$(MyStr, " As " & AsTypeArray(I) & " As " & AsTypeArray(I), " As " & AsTypeArray(I))
                SufNo = I
              End If
              GoTo ConstWithSufNumber
            End If
          Next I
        End If
      End If
    End If
    MyStr = ConcealParameterSpaces(MyStr)
    For I = 0 To 5
      TSPos = InStr(MyStr, TypeSuffixArray(I))
      Do While TSPos
        If InCode(MyStr, TSPos) Then
          numcut = InStrRev(MyStr, SngSpace, TSPos)
          If numcut Then
            If IsNumeric(Mid$(MyStr, numcut, TSPos - numcut)) Then
              Mid$(MyStr, TSPos, 1) = SngSpace
            End If
          End If
        End If
        TSPos = InStr(TSPos + 1, MyStr, TypeSuffixArray(I))
      Loop
    Next I
    If InStr(MyStr, " Lib ") Then
      For I = 0 To 5
        TSPos = InStr(MyStr, TypeSuffixArray(I) & " Lib ")
        Do While TSPos
          If InCode(MyStr, TSPos) Then
            Mid$(MyStr, TSPos, 1) = SngSpace
            MyStr = MyStr & " As " & AsTypeArray(I)
            SufNo = I
            Exit For
          End If
          TSPos = InStr(TSPos + 1, MyStr, TypeSuffixArray(I) & " Lib ")
        Loop
      Next I
    End If
    For I = 0 To 5
      TSPos = InStr(MyStr, TypeSuffixArray(I))
      If TSPos = Len(MyStr) Or InStr(",()", Mid$(MyStr, TSPos + 1, 1)) Then
        If InCode(MyStr, TSPos) Then
          'Last character or followed by a space, comma or left bracket
          'will not attack older DB referencing style of DB!Table
          Do
            If TSPos Then
              'v2.2.4
              If Not IsNumeric(Left$(MyStr, TSPos - 1)) Then
                If TSPos < Len(MyStr) Then
                  If Mid$(MyStr, TSPos + 1, 1) = LBracket Then
                    BracketPos = 0
                    If GetLeftBracketPos(Left$(MyStr, TSPos)) = 0 Then
                      BracketPos = InStrRev(MyStr, RBracket)
                     Else
                      BracketPos = GetRightBracketPos(MyStr)
                    End If
                    MyStr = Left$(MyStr, BracketPos) & " As " & AsTypeArray(I) & Mid$(MyStr, BracketPos + 1)
                    SufNo = I
                    Mid$(MyStr, TSPos) = SngSpace
                   ElseIf Mid$(MyStr, TSPos + 1, 4) = " Lib" Then
                    BracketPos = InStrRev(MyStr, RBracket)
                    MyStr = Left$(MyStr, BracketPos) & " As " & AsTypeArray(I) & Mid$(MyStr, BracketPos + 1)
                    SufNo = I
                    Mid$(MyStr, TSPos) = SngSpace
                   ElseIf Not Mid$(MyStr, TSPos - 1, 1) = SngSpace Then
                    MyStr = Left$(MyStr, TSPos - 1) & " As " & AsTypeArray(I) & Mid$(MyStr, TSPos + 1)
                    SufNo = I
                  End If
                 Else
                  MyStr = Left$(MyStr, TSPos - 1) & " As " & AsTypeArray(I) & Mid$(MyStr, TSPos + 1)
                  SufNo = I
                End If
              End If
            End If
            TSPos = InStr(TSPos + 1, MyStr, TypeSuffixArray(I))
          Loop While TSPos
        End If
      End If
    Next I
    MyStr = Replace$(MyStr, Chr$(160), SngSpace)
    If LenB(ConstValue) Then
      MyStr = MyStr & SngSpace & ConstValue
    End If
ConstWithSufNumber:
    If InStr(varName, MyStr) = 0 Then
      MyStr = MyStr & CommentStore & vbNewLine & UPDATED_MSG & "Obsolete Type Suffix replaced."
     Else
      MyStr = MyStr & CommentStore
    End If
    TypeSuffixExtender = SpaceOffSet & MyStr
  End If
  On Error GoTo 0

Exit Function

BadError:
  TypeSuffixExtender = varName

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:20:20 AM) 57 + 1296 = 1353 Lines Thanks Ulli for inspiration and lots of code.

