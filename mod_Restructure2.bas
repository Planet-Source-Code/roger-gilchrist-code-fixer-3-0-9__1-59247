Attribute VB_Name = "mod_Restructure2"
Option Explicit
Public Enum IFType
  None
  Simple
  Complex1
  Complex2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private None, Simple, Complex1, Complex2
#End If


Private Function AndIsNotBitWise(ByVal strTest As String) As Boolean

  Dim lngTmpIndex As Long
  Dim TmpA        As Variant
  Dim I           As Long
  Dim IsBool      As Long
  Dim arrTest     As Variant

  arrTest = Array("True", "False")
  'anything containing '=><' evaluates to a Boolean value and its 'And' will be a Logical Operator
  'anything which is numeric is using 'And' as a bitwise operator
  'Bitwise cannot be short circuited but can evalate as a whole to Boolean
  'This can be made clearer by enclosing the operation in brackets
  'This distinction also explains why '*If Len(x) And Len(y) then' (bitwise)
  'is not quite the same as           '*If Len(x) > 0 And Len(y) > 0 Then'
  '*if x="Fred" and y="Wilma" the 1st equates to '*If (4 And 5) Then' => 'If 4 Then' and as 4 is not 0 VB say it's True
  '                              2nd equates to '*If True And True Then' => If True Then
  'in this case either way will work but depending on the values of len(x) and len(y) bitwise can return unexpected False values.
  '*If Len(x)= 94 and Len(y)=32 the bitwise operation will return 0 and that's False so test will fail
  '
  strTest = Mid$(strTest, 4, Len(strTest) - 8)
  DisguiseLiteral strTest, " And ", True
  TmpA = Split(strTest, " And ")
  For I = LBound(TmpA) To UBound(TmpA)
    If InStr(TmpA(I), " Or ") Then
      IsBool = IsBool - 1
    End If
    If InstrAtPositionArray(TmpA(I), ipAny, False, EqualInCode, " > ", " < ", " <> ", " <= ", " >= ") Then
      IsBool = IsBool + 1
     ElseIf IsInArray(TmpA(I), arrTest) Then
      IsBool = IsBool + 1
     ElseIf Right$(TmpA(I), 6) = ".Value" Then
      'This is a bit weak but you are usually testing an OptionBox or CheckBox
      'using its default Value (this assumes that the Fix  to insert Default Properties has worked)
      lngTmpIndex = CntrlDescMember(Left$(TmpA(I), Len(TmpA(I)) - 6))
      If lngTmpIndex > -1 Then
        Select Case CntrlDesc(lngTmpIndex).CDClass
          'Select Case ControlClassArray(ArrayPos(Left$(TmpA(I), Len(TmpA(I)) - 6), ControlNamesArray))
         Case "OptionBox", "CheckBox"
          IsBool = IsBool + 1
        End Select
      End If
     ElseIf IsNumeric(TmpA(I)) Then
      IsBool = IsBool - 1
    End If
  Next I
  AndIsNotBitWise = (IsBool = UBound(TmpA) + 1)

End Function

Public Function CheckDefaultAttribute(cMod As CodeModule, _
                                      ByVal strProcName As String) As Boolean

  Dim I                  As Long
  Dim MemberAttributes() As Variant   'vbArray of 13 values
  Dim Hit                As Boolean

  ' NEW identifies default properties of user created timers
  MemberAttributes = Attributes(ModDescMember(cMod.Parent.Name))
  For I = 1 To UBound(MemberAttributes)
    Err.Clear
    On Error Resume Next
    With cMod.Members(MemberAttributes(I)(MemName))
      If strProcName = MemberAttributes(I)(MemName) Then
        If MemberAttributes(I)(MemStme) = 0 Then
          CheckDefaultAttribute = True
          Hit = True
        End If
      End If
    End With
    On Error GoTo 0
    If Hit Then
      Exit For
    End If
  Next I

End Function

Public Sub DoSeparateCompoundLines(ByVal ModuleNumber As Long, _
                                   strWork As String, _
                                   ByVal CompName As String, _
                                   Optional arrGoTo As Variant)

  Dim CommentStore   As String
  Dim strFix         As String
  Dim GoToOnCodeLine As String
  Dim StrWordOne     As String

  'v2.4.4 improved support for GoTo labels on same line as other code
  'named parameter ' := ' is safe because we replace colonspace not just colon
  'Time literals are protected by Safe_Replace
  If dofix(ModuleNumber, SeperateCompounds) Then
    'only if colon used
    'v3.0.7 just a safetey thing found while building fast fix
    If strWork = Colon Then
      strWork = vbNullString
    End If
    'c2.7.5 New FIx 'Next X, Y'
    If Left$(strWork, 5) = "Next " Then
      ExtractCode strWork, CommentStore
      Do While InStr(strWork, ", ")
        strWork = Replace$(strWork, ", ", vbNewLine & "Next ")
      Loop
      strWork = strWork & CommentStore
    End If
    If InstrAtPosition(strWork, ":", ipAny, False) Then
      ExtractCode strWork, CommentStore
      If InStr(strWork, ":") Then
        If InstrAtPosition(strWork, ":", ipAny, False) Then
          'and colon is in Code
          'GoTo targets need to retain their colon to keep label status
          If IsGotoLabel(strWork, CompName, strFix, arrGoTo) Then
            'Other parts of code fail if Goto Target has comment on same line
            'so separate line but retain the colon
            'ver 2 also comments poorly named GoTo labels
            If Len(strFix) Then
              strWork = strFix
            End If
            strWork = strWork & vbNewLine & Trim$(CommentStore)
            GoTo Done
          End If
          'extremely rare but legal label on same line as other code
          'v2.4.4 arrGoTo improved support for GoTo labels on same line as other code
          If IsGotoLabel(LeftWord(strWork), CompName, , arrGoTo) Then
            If Not WordIsVBSingleWordCommand(LeftWord(strWork)) Then
              GoToOnCodeLine = LeftWord(strWork)
              strWork = Trim$(Mid$(strWork, Len(GoToOnCodeLine) + 1))
              CommentStore = Trim$(CommentStore)
            End If
           Else
            'a label is also the name of a routine
            'extremely rare but legal
            'v 2.2.7 fixed to seperate off non-GoTo first word with colon
            StrWordOne = Left$(LeftWord(strWork), Len(LeftWord(strWork)) - 1)
            If IsProcedure(StrWordOne) Then
              If Not inProcOfLine(ModuleNumber, strWork, "GoTo " & StrWordOne) Then
                'if not then split the line
                strWork = Replace$(strWork, ":", vbNewLine, , 1) & Trim$(CommentStore)
                GoTo Done
              End If
            End If
          End If
          'deal with legal but unnecessary colons
          If InstrAtPosition(strWork, "Else:", ipAny) Then
            '*If X then DoBarney Else: DoFred
            If InStr(strWork, "Case Else:") = 0 Then
              strWork = Safe_Replace(strWork, " Else: ", " Else ")
            End If
          End If
          strWork = Safe_Replace(strWork, " Then: ", " Then ")
          '*If X Then: DoBarney Else DoFred
          '
          'All other code colons can be replaced with new line
          strWork = Safe_Replace(strWork, ": ", vbNewLine)
          If Right$(strWork, 1) = ":" Then
            'Just in case someone put an unnecessary colon on end of a line
            If InStr(strWork, SngSpace) Then
              If InCode(strWork, Len(strWork)) Then
                strWork = Left$(strWork, Len(strWork) - 1)
              End If
            End If
          End If
          NewLineTrim strWork
          strWork = strWork & CommentStore
         Else
          'colon only appeared in comment so restore comment and exit
          strWork = strWork & CommentStore
        End If
        If LenB(GoToOnCodeLine) Then
          strWork = GoToOnCodeLine & vbNewLine & strWork
        End If
      End If
    End If
  End If
Done:

End Sub

Public Function ErrorResumeCloser(ByVal varName As String) As String

  Dim MyStr        As String
  Dim CommentStore As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Apply correct variable to Next in For...Next structures
  'Because this error can occur at last line of code the comment is inserted above the last End Sub/Function/Property
  On Error GoTo BadError
  MyStr = varName
  If ExtractCode(MyStr, CommentStore) Then
    Select Case FixData(NCloseResume).FixLevel
     Case CommentOnly
      ErrorResumeCloser = Marker(MyStr & CommentStore, SUGGESTION_MSG & "Insert On Error Goto 0 to close the 'On Error Resume Next' error trap.", MBefore)
     Case FixAndComment
      ErrorResumeCloser = Marker(MyStr & CommentStore, "On Error Goto 0" & RISK_MSG & "Turns off 'On Error Resume Next' in routine(Good coding but may not be what you want)", MBefore)
     Case JustFix
      ErrorResumeCloser = "On Error Goto 0" & vbNewLine & MyStr & CommentStore
    End Select
    AddNfix NCloseResume
    On Error GoTo 0
  End If

Exit Function

BadError:
  ErrorResumeCloser = varName

End Function

Public Sub FormSelfReference(cMod As CodeModule)

  Dim UpDated    As Boolean
  Dim arrMembers As Variant
  Dim ArrRoutine As Variant
  Dim I          As Long
  Dim J          As Long
  Dim MaxFactor  As Long
  Dim TPos       As Long
  Dim strName    As String
  Dim StrNameDot As String
  Dim strW       As String
  Dim strWork    As String

  'Thanks to Richard Brisley who used the unusual code style
  'of using form referencing in a one form program
  'Code Fixer messed up by converting procedures to Private
  'and that made the procedure invisible to FormName.ProcName style calls
  'which can only see public procedures
  'This procedure strips FormName off all procedure and control calls
  'where the FormName is same as component name.
  'Form Properties are left with the FormName as it is easier to read
  'calls to other forms are not touched.
  strName = cMod.Parent.Name
  StrNameDot = strName & "."
  arrMembers = GetMembersArray(cMod)
  MaxFactor = UBound(arrMembers)
  If MaxFactor > -1 Then
    For I = 1 To MaxFactor
      MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
      ArrRoutine = Split(arrMembers(I), vbNewLine)
      If UBound(ArrRoutine) > -1 Then
        For J = LBound(ArrRoutine) To UBound(ArrRoutine)
          '2.0.1 fix for form name = end of control name
          TPos = InStrWholeWord(ArrRoutine(J), strName)
          If TPos Then
            If MultiLeft(ArrRoutine(J), True, "With " & strName) Then
              ArrRoutine(J) = ArrRoutine(J) & WARNING_MSG & "With Structures for Module's own name are not usually necessary unless setting Form Properties"
             Else
              If InCode(ArrRoutine(J), TPos) Then
                'v2.9.3 Thanks Ian K this is missing fix
                'it is safe to remove self-reference because even if you have 2 same name procedures one on form and on in a module
                'VB opts for the local one unless you use the Module.Proc format.
                'this is rough will tighten it up soonish
                'idea is to get the next word to check that it is suitable for this fix
                strW = Mid$(ArrRoutine(J), TPos + Len(StrNameDot))
                strW = LeftWord(ExpandForDetection(strW))
                If InStr(strW, ".") Then
                  strW = Left$(strW, InStr(strW, ".") - 1)
                End If
                If InStr(strW, ")") Then
                  strW = Left$(strW, InStr(strW, ")") - 1)
                End If
                If IsProcedure(strW) Or IsControlName(strW) Or IsDeclaration(strW) Then
                  'this test stops property setting code (FormName.Property = X) from being reduced
                  strWork = ArrRoutine(J)
                  Do
                    strWork = Safe_Replace(strWork, StrNameDot, vbNullString)
                  Loop While InStr(strWork, StrNameDot)
                  ArrRoutine(J) = strWork & WARNING_MSG & "Unneeded module self-reference removed" & IIf(Xcheck(XPrevCom), PREVIOUSCODE_MSG & ArrRoutine(J), vbNullString)
                  UpDated = True
                End If
              End If
            End If
          End If
        Next J
        arrMembers(I) = Join(CleanArray(ArrRoutine), vbNewLine)
      End If
    Next I
    ReWriteMembers cMod, arrMembers, UpDated
  End If

End Sub

Public Function GetMembersArray(cMod As CodeModule) As Variant

  Dim MemberArray() As Variant
  Dim arrTmp        As Variant
  Dim I             As Long
  Dim J             As Long

  'ver 1.0.93 simplified coding
  ' gets an array where each member is either a routine or the whole declaration section
  'MEmber(0) is the Declarations section (Maybe empty)
  'each other member contains one routine
  With cMod
    If Xcheck(XVisScan) Then
      .CodePane.TopLine = 1
    End If
    arrTmp = FullMemberExtraction(cMod)
    If UBound(arrTmp) > -1 Then
      ReDim MemberArray(UBound(arrTmp)) As Variant
      MemberArray(0) = Join(GetDeclarationArray(cMod), vbNewLine)
     Else
      ReDim MemberArray(0) As Variant
      MemberArray(0) = Join(GetDeclarationArray(cMod), vbNewLine)
    End If
    For I = 0 To UBound(arrTmp)
      If Len(arrTmp(I)) Then
        MemberArray(1 + J) = arrTmp(I)
        J = J + 1
      End If
    Next I
    Safe_Sleep
    For I = LBound(MemberArray) To UBound(MemberArray)
      If Left$(MemberArray(I), 7) = "#End If" Then
        MemberArray(I) = Trim$(MemberArray(I))
        MemberArray(I) = Mid$(MemberArray(I), 8)
      End If
    Next I
    GetMembersArray = MemberArray
    Erase MemberArray
    Erase arrTmp
  End With

End Function

Public Function GetWholeLineArray(Arr As Variant, _
                                  ByVal Lnum As Long, _
                                  Range As Long) As String

  Range = 0
  GetWholeLineArray = Arr(Lnum)
  If UBound(Arr) > Lnum Then
    Do While HasLineCont(GetWholeLineArray)
      Range = Range + 1
      GetWholeLineArray = Left$(GetWholeLineArray, Len(GetWholeLineArray) - 1) & LTrim$(Arr(Lnum + Range))
    Loop
  End If

End Function

Public Function IfAndThenShortCircuitLine(ByVal strTest As String) As Boolean

  'find potential line for IfAndThenShortCircuit_Apply

  strTest = strCodeOnly(strTest)
  If InstrAtPosition(strTest, "And", ipAny) Then
    If InstrAtPosition(strTest, "If", IpLeft) Then
      If InstrAtPosition(strTest, "Then", IpRight) Then
        IfAndThenShortCircuitLine = True
      End If
    End If
  End If

End Function

Public Function IfAndThenShortCircuitSafeToApply(ByVal strTest As String) As Boolean

  'CHeck potential line for IfAndThenShortCircuit_Apply is not too complex to apply

  If Not EnclosedInBrackets(strTest, InStr(strTest, " And ")) Then
    IfAndThenShortCircuitSafeToApply = AndIsNotBitWise(strTest)
  End If

End Function

Private Function inProcOfLine(ByVal ModuleNumber As Long, _
                              ByVal strTest1 As String, _
                              ByVal strTest2 As String) As Boolean

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim CurCompCount As Long
  Dim TmpDecArr    As Variant
  Dim I            As Long

  'MOst of the following tests could be incorperated into a single code sweeper
  'but separating them out makes code clearer if slower.
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        If ModuleNumber = CurCompCount Then
          TmpDecArr = GetMembersArray(Comp.CodeModule)
          For I = LBound(TmpDecArr) To UBound(TmpDecArr)
            If InStr(TmpDecArr(I), strTest1) Then
              If InStr(TmpDecArr(I), strTest2) Then
                inProcOfLine = True
                Exit For
              End If
            End If
          Next I
          '
          If inProcOfLine Then
            Exit For
          End If
        End If
      End If
      If inProcOfLine Then
        Exit For
      End If
    Next Comp
    If inProcOfLine Then
      Exit For
    End If
  Next Proj

End Function

Private Sub NewLineTrim(strCode As String)

  Dim TmpA As Variant
  Dim I    As Long

  'suggested by 'Form_Split_&_Fly' which used odd formatting
  'With f1:    .Move .Left - SPEED, .Top - SPEED:    End With
  'which didn't break properly due to extra blanks
  'so failed the With purity test
  TmpA = Split(strCode, vbNewLine)
  For I = LBound(TmpA) To UBound(TmpA)
    TmpA(I) = Trim$(TmpA(I))
  Next I
  strCode = Join(TmpA, vbNewLine)

End Sub

Private Function OkToDo(ByVal PropReplace As Boolean, _
                        ByVal Tval As String, _
                        ByVal PtargetPos As Long, _
                        ByVal varFind As Variant, _
                        Optional ByVal ControlAware As Boolean = False) As Boolean

  Dim Tlen As Long

  'This routine protects default properties from being over-applied
  'depending on what's being replaced this routine conducts different tests
  '*If PropReplace is True then it looks for a space after the target word  or being at end of string
  'this stops it hitting the admittedly rare case of two controls having near identical names
  'ie Fred and Fred2
  '*if False then it uses the other test which works for all other instances
  'ver 1.1.00 second test restructured as it was misfiring aand added Error Trap
  '
  Tlen = Len(varFind)
  On Error Resume Next
  If PropReplace Then
    If ControlAware Then
      OkToDo = IsPunctExcept(Mid$(Tval, PtargetPos + Tlen, 1), "_.") Or PtargetPos = Len(Tval) - Tlen + 1
     Else
      OkToDo = IsPunct(Mid$(Tval, PtargetPos + Tlen, 1)) Or PtargetPos = Len(Tval) - Tlen + 1
    End If
   Else
    If ControlAware Then
      OkToDo = Mid$(Tval, PtargetPos + Tlen, 1) <> "." Or Mid$(Tval, PtargetPos + Tlen - 1, 2) = " ." Or PtargetPos = Len(Tval) - Tlen + 1 Or (InStrWholeWordRX(Tval, varFind) And InCode(Tval, InStr(Tval, varFind)))
     Else
      OkToDo = Mid$(Tval, PtargetPos + Tlen, 1) <> "." Or Mid$(Tval, PtargetPos + Tlen - 1, 2) = " ." Or PtargetPos = Len(Tval) - Tlen + 1
    End If
  End If
  On Error GoTo 0

End Function

Public Function PleonasmCleaner(ByVal strCode As String, _
                                bTarget As Boolean) As String

  Dim MyStr        As String
  Dim CommentStore As String
  Dim SpaceOffSet  As String

  'v2.7.8 updated to inccorperate an False fixer as well
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Remove unnecessary '= True' from code
  On Error GoTo BadError
  ' remove end comments for restoring after changing Type Suffixes
  MyStr = strCode
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    '
    If ProtectBoolean(MyStr, bTarget) Then
      PleonasmCleaner = MyStr & CommentStore
     Else
      UpDateTF MyStr, SpaceOffSet, bTarget
      '---------------------------
      If MyStr & CommentStore <> strCode Then
        AddNfix NPleonasmFix
        Select Case FixData(NPleonasmFix).FixLevel
         Case CommentOnly
          If bTarget Then
            PleonasmCleaner = strCode & vbNewLine & _
             RGSignature & "Pleonasm (unnecessary '= True') should be removed."
           Else
            PleonasmCleaner = strCode & vbNewLine & _
             SUGGESTION_MSG & "'X = False' could be converted to faster 'Not X '."
            'above forces Previous code
          End If
         Case FixAndComment
          If bTarget Then
            PleonasmCleaner = MyStr & CommentStore & vbNewLine & _
             RGSignature & "Pleonasm Removed" & IIf(Xcheck(XPrevCom), vbNewLine & _
             PREVIOUSCODE_MSG & strCode, vbNullString)
           Else
            PleonasmCleaner = MyStr & CommentStore & vbNewLine & _
             WARNING_MSG & "'X = False' converted to faster 'Not X '" & IIf(Xcheck(XPrevCom), vbNewLine & _
             PREVIOUSCODE_MSG & strCode, vbNullString)
          End If
         Case JustFix
          PleonasmCleaner = MyStr & CommentStore
        End Select
       Else
        PleonasmCleaner = strCode
      End If
    End If
    On Error GoTo 0
  End If

Exit Function

BadError:
  PleonasmCleaner = strCode

End Function

Private Function ProtectBoolean(strCode As String, _
                                bTarget As Boolean) As Boolean

  Dim arrTmp As Variant
  Dim strTF  As String

  'v2.7.8 support for PleonasmCleaner
  strTF = IIf(bTarget, "True", "False")
  If SmartRight(strCode, "= " & strTF) Then
    If Not MultiLeft(strCode, True, "While", "Do Until", "Do While", "Loop Until", "Loop While") Then
      'it's just an assignment of False to a variable so leave it.
      ProtectBoolean = True
      ' GoTo SingleExit
    End If
  End If
  If Not ProtectBoolean Then
    If isProcHead(strCode) Then
      'it maybe an Optional parameter defaulting to False
      ProtectBoolean = True
      ' GoTo SingleExit
    End If
  End If
  If Not ProtectBoolean Then
    arrTmp = Split(strCode)
    'rare case of 'A = True And <some Condition>'
    'being used to generate a bitwise calculation
    If UBound(arrTmp) > 2 Then
      If arrTmp(1) & arrTmp(2) & arrTmp(3) = "=" & bTarget & "And" Then
        ProtectBoolean = True
        'GoTo SingleExit
      End If
    End If
  End If
  'Rare case
  '*If X Then Control.property = True Else ......'
  If Not ProtectBoolean Then
    If Left$(strCode, 3) = "If " Then
      If Between(InStr(strCode, " Then "), InStr(strCode, " = " & bTarget & " "), InStr(strCode, " Else")) Then
        ProtectBoolean = True
        ' GoTo SingleExit
      End If
    End If
  End If
  If Not ProtectBoolean Then
    'v2.9.2
    If InStr(strCode, "IsWindow") Then ' simple trap
      If IsAPI("IsWindow") Then          ' might be a user defined proc which will be OK
        If IsWholeWord(strCode, "IsWindow", InStr(1, strCode, "IsWindow")) Then
          'just incase it is a partial word
          ProtectBoolean = True
        End If
      End If
    End If
  End If
  'SingleExit:

End Function

Public Sub QuickSort(ByVal ixFrom As Long, _
                     ByVal ixThru As Long, _
                     ByVal KeyIsIn As Long)

  Dim ixLeft   As Long
  Dim ixRite   As Long
  Dim SortElem As Variant

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'Sorts a table of vbArrays
  If ixFrom < ixThru Then
    ixLeft = ixFrom
    ixRite = ixThru
    'Get ref element and make room
    SortElem = SortElems(ixLeft)
    Do
      Do Until ixRite = ixLeft
        If LCase$(SortElems(ixRite)(KeyIsIn)) >= LCase$(SortElem(KeyIsIn)) Then
          ixRite = ixRite - 1
         Else
          SortElems(ixLeft) = SortElems(ixRite)
          '...and leave the item just moved alone for now
          ixLeft = ixLeft + 1
          Exit Do
        End If
      Loop
      Do Until ixLeft = ixRite
        If LCase$(SortElems(ixLeft)(KeyIsIn)) <= LCase$(SortElem(KeyIsIn)) Then
          ixLeft = ixLeft + 1
         Else
          SortElems(ixRite) = SortElems(ixLeft)
          '...and leave the item just moved alone for now
          ixRite = ixRite - 1
          Exit Do
        End If
      Loop
    Loop Until ixLeft = ixRite
    'now the indexes have met and all bigger items are to the right and all smaller items are left
    SortElems(ixRite) = SortElem
    If ixLeft - ixFrom > ixThru - ixRite Then
      QuickSort ixRite + 1, ixThru, KeyIsIn
      QuickSort ixFrom, ixLeft - 1, KeyIsIn
     Else
      QuickSort ixFrom, ixLeft - 1, KeyIsIn
      QuickSort ixRite + 1, ixThru, KeyIsIn
    End If
  End If

End Sub

Public Sub ReWriteMembers(cMod As CodeModule, _
                          arrM As Variant, _
                          bUpdate As Boolean)

  If bUpdate Then
    ReWriter cMod, arrM, RWMembers
    bUpdate = False
  End If

End Sub

Public Function RoutineSearch(ByVal tarr As Variant, _
                              ByVal strFind As String, _
                              Optional ByVal IgnoreLine As Long = -1, _
                              Optional ByVal FindAtPos As InstrLocations = ipAny, _
                              Optional IncludeComments As Boolean = True) As Boolean

  Dim I As Long

  For I = LBound(tarr) To UBound(tarr)
    If I <> IgnoreLine Then
      If InstrAtPosition(tarr(I), strFind, FindAtPos) Then
        If InCode(tarr(I), FindAtPos) Or IncludeComments Then
          'v2.5.5 tightened test to ignore 'FuncName_Error' type labels being detected
          'v2.5.6 moved to correct place in logic
          If IsRealWord(tarr(I), strFind) Then
            RoutineSearch = True
            Exit For 'unction
          End If
        End If
      End If
    End If
  Next I

End Function

Private Function safe_InStr(ByVal StartPos As Long, _
                            ByVal varSearch As Variant, _
                            ByVal varFind As Variant, _
                            ByVal Standard As Boolean) As Long

  'extends Instr so that it only finds real words

  safe_InStr = InStr(StartPos, varSearch, varFind)
  If Not Standard Then
    If safe_InStr Then
      Select Case safe_InStr
       Case 1
        'left edge on string
        If Not IsPunct(Mid$(varSearch, safe_InStr + Len(varFind), 1)) Or Len(varSearch) = Len(varFind) Then
          safe_InStr = 0
        End If
       Case Len(varSearch) - Len(varFind) + 1
        'right edge of string
        If Not IsPunct(Mid$(varSearch, safe_InStr - 1, 1)) Or Len(varSearch) = Len(varFind) Then
          safe_InStr = 0
        End If
       Case Else
        'anywhere else in string
        If Not (IsPunct(Mid$(varSearch, safe_InStr - 1, 1)) And IsPunct(Mid$(varSearch, safe_InStr + Len(varFind), 1))) Then
          safe_InStr = 0
        End If
      End Select
    End If
  End If

End Function

Private Function safe_InStr2(ByVal StartPos As Long, _
                             ByVal varSearch As Variant, _
                             ByVal varFind As Variant, _
                             ByVal Standard As Boolean) As Long

  Dim InstrLoop As Long

  'extends Instr so that it only finds real words
  'modified for testing  default property
  'ver 1.1.30
  If InstrAtPosition(varSearch, varFind, ipAny, True) Then
    safe_InStr2 = InStr(StartPos, varSearch, varFind)
    If safe_InStr2 > 0 Then
fred:
      InstrLoop = InStr(safe_InStr2 + 1, varSearch, varFind)
      If Not Standard Then
        If safe_InStr2 Then
          Select Case safe_InStr2
           Case 1
            'left edge on string
            If Not IsPunctExcept(Mid$(varSearch, safe_InStr2 + Len(varFind), 1), "_.(") Or Len(varSearch) = Len(varFind) Then
              safe_InStr2 = 0
            End If
           Case Len(varSearch) - Len(varFind) + 1
            'right edge of string
            If Not IsPunctExcept(Mid$(varSearch, safe_InStr2 - 1, 1), "_.") Or Len(varSearch) = Len(varFind) Then
              safe_InStr2 = 0
            End If
           Case Else
            'anywhere else in string
            If Not IsPunctExcept(Mid$(varSearch, safe_InStr2 - 1, 1), "_.(") And IsPunctExcept(Mid$(varSearch, safe_InStr2 + Len(varFind), 1), "_.") And True Then
              safe_InStr2 = 0
            End If
          End Select
        End If
      End If
      If InstrLoop > 0 Then
        If safe_InStr2 = 0 Then
          If InstrAtPosition(Mid$(varSearch, InstrLoop), varFind, ipAny, True) Then
            safe_InStr2 = InStr(InstrLoop, varSearch, varFind)
            If safe_InStr2 > 0 Then
              GoTo fred
            End If
          End If
        End If
      End If
    End If
  End If

End Function

Public Function Safe_Replace(Expression As Variant, _
                             Find As Variant, _
                             VarReplace As Variant, _
                             Optional Start As Long = 1, _
                             Optional ByVal lngCount As Long = -1, _
                             Optional Standard As Boolean = True, _
                             Optional PropReplace As Boolean = False, _
                             Optional ControlAware As Boolean = False) As String

  Dim PossibleTarget As Long
  Dim LocalCount     As Long
  Dim LOffset        As Long

  'update PropReplace causes a different test to be conducted see OkToDo for details
  'Safe_Replace is designed to replace only CODE, Comments, Literal Strings and date literals cannot be touched
  Safe_Replace = Expression
  If LenB(Safe_Replace) > 0 Then
    'Ver 1.1.00 Speed up. Skips idiot case were the find and replace are the same
    If Find <> VarReplace Then
      'found coming from Dimformat (Fixed there but could be from other places so added safety here too)
      If ControlAware Then
        PossibleTarget = InStrWholeWordRX(Safe_Replace, Find)
       Else
        PossibleTarget = safe_InStr(Start, Safe_Replace, Find, Standard)
      End If
      If PossibleTarget Then
        LOffset = Len(Find)
        If LOffset < Len(VarReplace) Then
          LOffset = Len(VarReplace)
        End If
        DisguiseLiteral Expression, Find, True
        Do While PossibleTarget
          If OkToDo(PropReplace, Safe_Replace, PossibleTarget, Find, ControlAware) Then
            If InCode(Safe_Replace, PossibleTarget) Then
              Safe_Replace = Left$(Safe_Replace, PossibleTarget - 1) & VarReplace & Mid$(Safe_Replace, PossibleTarget + Len(Find))
              LocalCount = LocalCount + 1
            End If
          End If
          If ControlAware Then
            PossibleTarget = safe_InStr2(PossibleTarget + LOffset + 1, Safe_Replace, Find, False)
           Else
            PossibleTarget = safe_InStr(PossibleTarget + LOffset + 1, Safe_Replace, Find, Standard)
          End If
          If lngCount > 0 Then
            If LocalCount >= lngCount Then
              Exit Do
            End If
          End If
        Loop
        DisguiseLiteral Expression, Find, False
      End If
    End If
  End If

End Function

Public Function ScopeTo(cMod As CodeModule, _
                        ByVal varName As String, _
                        indent As Long) As String

  Dim MyStr        As String
  Dim CommentStore As String
  Dim SpaceOffSet  As String
  Dim NewScope     As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Set missing scope to Private
  'is not always correct but if a Private should be Public it will show up
  ' as soon as you try to start, then just search for the Declaration, Sub, Function or Variable involved
  'Setting to Public would make the code use memory unnecessarily
  'UPDATE: 14 Jan 2003 Dim Enums and Dim Declares are set to Public (which they usually will be)
  '                    All other Dims are set to Private
  ScopeTo = varName
  On Error GoTo BadError
  MyStr = varName
  If ExtractCode(MyStr, CommentStore, SpaceOffSet) Then
    '*if its an enum or Declare it is almost certainly going to be public
    ' the (Type AND As AND len()>4 ) test is because VB actually suggests using Type as a member of Enums and Types
    'UPDATE: 22 Jan 2003 simplified the routine with this safety trap
    'for Type and Enum members of the form 'Type As Long'
    If ReservedWordAs_Type_or_Enum_Member(MyStr) Then
      ScopeTo = SpaceOffSet & MyStr & CommentStore & SUGGESTION_MSG & "Legal but ill-advised to use VB Reserved words as Variable names"
      indent = indent + 1
     Else
      If FixData(DimGlobal2PublicPrivate).FixLevel <> Off Then
        NewScope = IIf(InstrAtPositionArray(MyStr, IpLeft, True, "Function", "Sub", "Enum", "Event", "Property", "Declare", "WithEvents") Or (InstrAtPosition(MyStr, "Type", IpLeft, True) And InstrAtPosition(MyStr, "As", IpNone, True) And Len(Trim$(MyStr)) > 4), "Public ", "Private ")
        If ArrayMember(LeftWord(MyStr), "Function", "Sub", "Enum", "Event", "Property", "Declare", "WithEvents") Then
          If LeftWord(MyStr) = "Property" Then
            'v 2.3.0 a property can be named 'Property'
            'This stops CF from applying Scope to the Assignment line of the 'Let' Routine
            If Not ArrayMember(WordAfter(MyStr, "Property"), "Let", "Set", "Get") Then
              Exit Function
            End If
          End If
          If CheckDefaultAttribute(cMod, WordInString(ExpandForDetection(MyStr), 2)) Then
            ScopeTo = SpaceOffSet & Replace$(MyStr, "Dim ", "Public ", , 1) & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to " & NewScope
            If WordInString(ScopeTo, 1) <> "Public" Then
              ScopeTo = "Public " & ScopeTo
            End If
            AddNfix DimGlobal2PublicPrivate
           ElseIf Left$(Trim$(MyStr), 4) = "Dim " Then
            ScopeTo = SpaceOffSet & Replace$(MyStr, "Dim ", NewScope, , 1) & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to " & NewScope
            AddNfix DimGlobal2PublicPrivate
           ElseIf InstrAtPositionArray(MyStr, IpLeft, True, "Enum", "Type") Then
            NewScope = ScopeGenerator(WordInString(ExpandForDetection(MyStr), 2), True)
            ScopeTo = SpaceOffSet & NewScope & MyStr & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to " & NewScope
           ElseIf InstrAtPositionArray(MyStr, IpLeft, True, "Declare", "Property") Then
            NewScope = ScopeGenerator(WordInString(ExpandForDetection(MyStr), 3))
            ScopeTo = SpaceOffSet & NewScope & MyStr & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to " & NewScope
           ElseIf InstrAtPosition(MyStr, "Event", IpLeft) Then
            ScopeTo = SpaceOffSet & "Public " & MyStr & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to Public"
           ElseIf InstrAtPositionArray(MyStr, IpLeft, True, "Event", "Const", "Sub", "Function") Then
            NewScope = ScopeGenerator(WordInString(ExpandForDetection(MyStr), 2))
            ScopeTo = SpaceOffSet & NewScope & MyStr & CommentStore & vbNewLine & _
             WARNING_MSG & "Scope Changed to " & NewScope
           Else
            '(NEW This probably wont hit now that concatentLineContinuation is available
            ScopeTo = SpaceOffSet & MyStr & CommentStore & vbNewLine & LineContinuationWarning & "5"
            AddNfix DimGlobal2PublicPrivate
          End If
        End If
      End If
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  ScopeTo = varName

End Function

Public Function SortingTagExtraction(cMod As CodeModule) As Variant

  Dim Tmpstring1        As String
  Dim I                 As Long
  Dim J                 As Long
  Dim K                 As Long
  Dim CleanElems()      As Variant
  Dim CleanElem         As Variant
  Dim TmpA              As Variant
  Dim FakeTopofRoutine  As Long
  Dim FakeEndofRoutine  As Long
  Dim FakeNameofRoutine As String

  With cMod
    If Xcheck(XVisScan) Then
      .CodePane.TopLine = 1
    End If
    ReDim CleanElems(0) As Variant
    'collect module descriptions -> (Name, StartingLine, Length)
    If .Members.Count Then
      For J = 1 To .Members.Count
        With .Members(J)
          Tmpstring1 = .Name
          I = (.Type = vbext_mt_Property Or .Type = vbext_mt_Method)
        End With
        If I Then
          For I = 1 To 4
            K = Choose(I, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set, vbext_pk_Proc)
            CleanElem = Null
            On Error Resume Next
            '*If you crash here first check that Error Trapping is not ON
            CleanElem = Array(Tmpstring1, .PRocStartLine(Tmpstring1, K), .ProcCountLines(Tmpstring1, K), K)
            On Error GoTo 0
            If Not IsNull(CleanElem) Then
              ReDim Preserve CleanElems(UBound(CleanElems) + 1)
              CleanElems(UBound(CleanElems)) = CleanElem
              'End If
            End If
          Next I
        End If
      Next J
      If UBound(CleanElems) > -1 Then
        If Not IsEmpty(CleanElems(UBound(CleanElems))) Then
          If CleanElems(UBound(CleanElems))(1) + CleanElems(UBound(CleanElems))(2) < .CountOfLines - .CountOfDeclarationLines Then
            CleanElems(UBound(CleanElems))(2) = CleanElems(UBound(CleanElems))(1) + CleanElems(UBound(CleanElems))(2) + .CountOfLines - CleanElems(UBound(CleanElems))(1) - .CountOfDeclarationLines
            CleanElems(UBound(CleanElems))(2) = CleanElems(UBound(CleanElems))(2) + 1
          End If
        End If
      End If
     Else
      'This is a slow way of doing the above for the special case that all the code
      'is enclosed by optional compilation structure '#If <var> Then' and '#End If'
      If .CountOfLines - .CountOfDeclarationLines Then
        For I = .CountOfDeclarationLines + 1 To .CountOfLines
          Tmpstring1 = .Lines(I, 1)
          If I = .CountOfLines Then
            If Left$(Tmpstring1, 7) = "#End If" Then
              CleanElems(UBound(CleanElems)) = Array(FakeNameofRoutine, FakeTopofRoutine, I)
              Exit For
            End If
          End If
          If InstrAtPositionSetArray(Tmpstring1, ipLeftOr2ndOr3rd, True, ArrFuncPropSub) Then
            FakeTopofRoutine = I
            TmpA = Split(ExpandForDetection(Tmpstring1))
            If InstrAtPositionSetArray(Tmpstring1, IpLeft, True, ArrFuncPropSub) Then
              FakeNameofRoutine = TmpA(1)
             Else
              FakeNameofRoutine = TmpA(2)
            End If
            FakeEndofRoutine = FakeTopofRoutine
            Do
              FakeEndofRoutine = FakeEndofRoutine + 1
            Loop Until MultiLeft(.Lines(FakeEndofRoutine, 1), True, "End Sub", "End Function", "End Property")
            CleanElem = Array(FakeNameofRoutine, FakeTopofRoutine, FakeEndofRoutine - FakeTopofRoutine)
            If Not IsNull(CleanElem) Then
              ReDim Preserve CleanElems(UBound(CleanElems) + 1)
              CleanElems(UBound(CleanElems)) = CleanElem
              I = FakeEndofRoutine
            End If
          End If
        Next I
        If FakeTopofRoutine + FakeEndofRoutine - FakeTopofRoutine < .CountOfLines - .CountOfDeclarationLines Then
          CleanElems(UBound(CleanElems))(2) = FakeEndofRoutine - FakeTopofRoutine + .CountOfLines - FakeTopofRoutine - .CountOfDeclarationLines
        End If
      End If
    End If
  End With
  SortingTagExtraction = CleanElems

End Function

Public Function TypeOfIf(ByVal tmpB As Variant, _
                         ByVal LineNo As Long, _
                         EndPos As Long) As IFType

  Dim I       As Long
  Dim Level   As Long
  Dim strTest As String

  For I = LineNo To GetEndOfRoutine(tmpB)
    strTest = tmpB(I)
    If InstrAtPosition(strTest, "If", IpLeft, True) Then
      'ver1.1.29 this stops single line If..Then....'s from being counted
      If InStrCode(strTest, " Then ") = 0 Then
        Level = Level + 1
      End If
    End If
    If InstrAtPosition(strTest, "End If", IpLeft, True) Then
      Level = Level - 1
    End If
    If Level = 0 Then
      TypeOfIf = Simple
      EndPos = I
      Exit For
    End If
    If Level = 1 Then
      If InstrAtPosition(strTest, "Else", IpLeft, True) Then
        TypeOfIf = Complex1
        EndPos = -1
        Exit For
      End If
      If InstrAtPosition(strTest, "ElseIf", IpLeft, True) Then
        TypeOfIf = Complex2
        EndPos = -1
        Exit For
      End If
    End If
  Next I

End Function

Public Function TypeSuffix2String(ByVal varSuffix As Variant) As String

  Dim I      As Long

  I = TypeSuffixArrayPosition(varSuffix)
  If I > -1 Then
    TypeSuffix2String = AsTypeArray(I)
  End If

End Function

Public Function TypeSuffixArrayPosition(varA As Variant) As Long

  'TypeSuffixArrayPosition = ArrayPos(Right$(varA, 1), TypeSuffixArray)

  TypeSuffixArrayPosition = QSortArrayPos(TypeSuffixArray, Right$(varA, 1))

End Function

Public Function TypeSuffixExists(ByVal varA As Variant) As Boolean

  'original code

  TypeSuffixExists = InQSortArray(TypeSuffixArray, Right$(varA, 1))
  'thanks to Manuel Muñoz for finding the bug this fixes
  'consider code
  'Private tipotit$, P$(), Tipos() As Single
  'Expand_SingleLine_SingleType_Declaration expanded it to
  'Private tipotit$ '<<<NOTE this is saved by original code for this Function
  'Private P$() As Single '<<<ERROR 1 as right of line is not a Type Suffix it added As Single
  'Private Tipos() As Single
  'then UpDate_TypeSuffix_Declarations would fix the $ resulting in the code
  'Private P() As String As Single '<<<ERROR 2
  'Thanks to this fix Error 1 doesn't occur so Error 2 doesn't either
  If Not TypeSuffixExists Then
    If InStr(varA, LBracket) > 1 Then
      TypeSuffixExists = InQSortArray(TypeSuffixArray, Mid$(varA, InStr(varA, LBracket) - 1, 1))
    End If
  End If

End Function

Private Function Update2Not(strCode As String) As Boolean

  'v2.7.8 support routine for UpDateTF

  If WordInString(strCode, 3) = "=" Then
    If WordInString(strCode, 4) = "False" Then
      Select Case LeftWord(strCode)
       Case "If"
        strCode = Safe_Replace(strCode, "If ", "If Not ", , 1)
        Update2Not = True
       Case "ElseIf"
        strCode = Safe_Replace(strCode, "ElseIf ", "ElseIf Not ", , 1)
        Update2Not = True
       Case "While"
        strCode = Safe_Replace(strCode, "While ", "While Not ", , 1)
        Update2Not = True
       Case "Do", "Loop"
        If InStr(strCode, " While ") Then
          strCode = Safe_Replace(strCode, " While ", " While Not ", , 1)
          Update2Not = True
         Else
          strCode = Safe_Replace(strCode, " Until ", " Until Not ", , 1)
          Update2Not = True
        End If
      End Select
    End If
  End If

End Function

Private Sub UpDateTF(strCode As String, _
                     ByVal StrSpaceOffeset As String, _
                     ByVal bTarget As Boolean)

  Dim strTF As String

  'v2.7.8 support routine for  PleonasmCleaner
  strTF = IIf(bTarget, "True", "False")
  If Not bTarget Then
    If Not Update2Not(strCode) Then
      Exit Sub
    End If
  End If
  If InStr(strCode, "As Boolean = " & strTF) Then
    'protect optional parameters
    strCode = Replace$(strCode, "As Boolean = " & strTF, "As Boolean=" & strTF)
  End If
  If InStr(strCode, "= " & strTF) Then
    strCode = Safe_Replace(strCode, "= " & strTF & ") ", ") ")
    strCode = StrSpaceOffeset & Safe_Replace(strCode, "= " & strTF & " ", SngSpace)
    strCode = StrSpaceOffeset & Safe_Replace(strCode, "= " & strTF, vbNullString)
   ElseIf InStr(strCode, " " & strTF & " = ") Then
    strCode = Safe_Replace(strCode, "(" & strTF & " = ", "( ")
    strCode = StrSpaceOffeset & Safe_Replace(strCode, " " & strTF & " = ", SngSpace)
  End If
  If InStr(strCode, "As Boolean=" & strTF) Then
    'protect optional parameters
    strCode = Replace$(strCode, "As Boolean=" & strTF, "As Boolean = " & strTF)
  End If

End Sub

Public Function WordInString(ByVal varChop As Variant, _
                             ByVal WordNo As Long) As String

  Dim TmpA As Variant

  If LenB(varChop) Then
    TmpA = Split(varChop)
    If WordNo < 0 Then
      'return last member whatever it is
      If UBound(TmpA) >= Abs(WordNo - 1) Then
        WordInString = TmpA(UBound(TmpA) + WordNo + 1)
      End If
     Else
      If UBound(TmpA) >= WordNo - 1 Then
        WordInString = TmpA(WordNo - 1)
      End If
    End If
  End If

End Function

Private Function WordIsVBSingleWordCommand(ByVal strTest As String) As Boolean

  'This is a guard routine to stop certain VB commands being detected as GoTo Targets
  'It is used by DoSeparateCompoundLines to decide whether to leave or remove the colon
  'v2.8.3 added safety

  strTest = Replace$(strTest, ":", ".")
  WordIsVBSingleWordCommand = ArrayMember(strTest, "Do", "While", "Loop", "Wend", "Else", "Beep")
  'VB can in fact distinguish between most of these in the format 'X:' but Beep: for some reason
  'can be used as either and VB defaults to label.
  'The code 'Beep: Beep: Beep'  only sounds twice the first one is treated as a label.
  '

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:17:41 AM) 10 + 1113 = 1123 Lines Thanks Ulli for inspiration and lots of code.

