Attribute VB_Name = "mod_GeneralServices"

Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public Const strTurnOffMsg             As String = vbNewLine & _
 "(Stop this message appearing from Settings tab)"
Public bAddinTerminate                 As Boolean
Public bAborting                       As Boolean
Public FrameActive                     As ETabPage
Public Const SngSpace                  As String = " "
Public Const DblSpace                  As String = SngSpace & SngSpace
Public Const CommaSpace                As String = "," & SngSpace
Public lngBigProcLines                 As Long
Public lngLineLength                   As Long
Public lngElseIndent                   As Long
Public lngCaseIndent                   As Long
Public Enum InstrLocations
  IpNone
  IpExact
  IpMiddle
  IpLeft
  IpRight
  ip2nd
  ip3rd
  ipLeftOr2nd
  ip2ndOr3rd
  ipLeftOr2ndOr3rd
  ipAny
  ip3rdorGreater
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private IpNone, IpExact, IpMiddle, IpLeft, IpRight, ip2nd, ip3rd, ipLeftOr2nd, ip2ndOr3rd, ipLeftOr2ndOr3rd, ipAny, ip3rdorGreater
#End If
Private PrevSettings                   As Variant
Public Enum SCMode
  SCAnd
  SCOr
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private SCAnd, SCOr
#End If
Private DisguiseStack                  As New Cls_StackVB2TheMax
'Private Const strNotLetterFilter       As String = "[!a-z!A-Z!À-Ö!Ø-ß!à-ö!ø-ÿ!0-9]"
'v3.0.3 realword test now ignores underscores
Private Const strNotLetterFilter       As String = "[!_!a-z!A-Z!À-Ö!Ø-ß!à-ö!ø-ÿ!0-9]"

Public Sub AddNfix(CounterID As NfixDesc)

  FixData(CounterID).FixModCount = FixData(CounterID).FixModCount + 1

End Sub

Public Function AppDetails() As String

  'VB tab width and font

  With App
    AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
  End With

End Function

Public Function ArrayHasContents(arrTest As Variant) As Boolean

  If IsArray(arrTest) Then
    On Error GoTo oops
    If UBound(arrTest) > -1 Then
      ArrayHasContents = True
    End If
  End If
oops:

End Function

Public Function ArrayPos(FindValue As Variant, _
                         arrSearch As Variant, _
                         Optional ByVal SkipPos As Long = 0) As Long

  Dim I As Long

  'VERY FAST function to find a value in an unsorted Array
  'By Brian Gillham
  'Versions:  VB5, VB6
  'Posted:  4/18/2001
  If IsInArray(FindValue, arrSearch) Then
    For I = LBound(arrSearch) + SkipPos To UBound(arrSearch)
      If arrSearch(I) = FindValue Then
        ArrayPos = I
        Exit For
      End If
    Next I
   Else
    ArrayPos = -1
  End If
LocalError:
  'Justin (just in case)

End Function

Public Function Between(ByVal Lo As Long, _
                        ByVal Tval As Long, _
                        ByVal Hi As Long, _
                        Optional ByVal Exclusive As Boolean = False) As Boolean

  Between = Hi >= Tval And Lo <= Tval
  If Exclusive Then
    Between = Hi > Tval And Lo < Tval
  End If

End Function

Public Function CleanArray(arrT As Variant, _
                           Optional ByVal ForceClean As Boolean = True) As Variant

  'enlarge array any code inserted delimiters
  'are used to create a new larger array
  'This allows the next routine to treat the routine as if it was newly aquired

  If ForceClean Then
    CleanArray = Split(BlankLineStripper(arrT), vbNewLine)
   Else
    CleanArray = arrT
  End If

End Function

Public Function CleanMsg(ByVal strMsg As String) As String

  'routine formats multiline messages for SmartMarker routine

  CleanMsg = ReplaceArray(strMsg, vbNewLine & _
   vbNewLine, vbNewLine, vbNewLine, RGSignature, RGSignature & RGSignature, RGSignature, RGSignature, vbNewLine & _
   RGSignature)

End Function

Private Function CommentClip(varSearch As Variant) As String

  Dim MyStr       As String
  Dim CommentPos  As Long
  Dim SpaceOffSet As Long

  'This code clips end comments from VarSearch
  'NOTE also Modifies VarSearch
  'UPDATE now copes with literal embedded '
  On Error GoTo BadError
  MyStr = varSearch
  CommentPos = InStr(1, MyStr, SQuote)
  If CommentPos > 0 Then
    Do While InLiteral(MyStr, CommentPos)
      CommentPos = InStr(CommentPos + 1, MyStr, SQuote)
      If CommentPos = 0 Then
        Exit Do
      End If
    Loop
    If CommentPos > 0 Then
      CommentClip = Mid$(MyStr, CommentPos)
      MyStr = Left$(MyStr, CommentPos - 1)
      'Preserve spaces with comment if comment is offset with them
      SpaceOffSet = Len(MyStr) - Len(RTrim$(MyStr))
      CommentClip = String$(SpaceOffSet, 32) & CommentClip
      varSearch = Left$(MyStr, Len(MyStr) - SpaceOffSet)
      'v 2.4.7 Thanks Evan Toder and Lawrence Miller
      ' special case of Long2Int where code is of form 'Function Wally(X as Integer) As Integer'
      ' the X is updated first then a comment with a newline is added
      ' this stops the Function Type updating and goes into endless loop
      If Len(varSearch) Then
        'v2.5.0
        Do While InStr(vbNewLine, Right$(varSearch, 1))
          CommentClip = CommentClip & Right$(varSearch, 1)
          varSearch = Left$(varSearch, Len(varSearch) - 1)
        Loop
      End If
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  CommentClip = vbNullString

End Function

Public Function ConcealParameterCommas(ByVal varCode As Variant, _
                                       Optional ByVal bStripDouble As Boolean = False) As String

  Dim CommaSpacePos As Long

  'Replace any CommaSpace with comma in strInBrackets parameters
  'This allows CommaSpace delimited Dim,Public, Private to be safely detected without cutting in parameters
  'VB will automatically restore them
  CommaSpacePos = GetCommaSpacePos(varCode)
  If CommaSpacePos Then
    If bStripDouble Then
      varCode = StripDoubleSpace(varCode)
    End If
    Do
      If InCode(varCode, CommaSpacePos) Then
        If EnclosedInBrackets(varCode, CommaSpacePos) Then
          varCode = Left$(varCode, CommaSpacePos - 1) & "," & Mid$(varCode, CommaSpacePos + 2)
        End If
      End If
      CommaSpacePos = GetCommaSpacePos(varCode, CommaSpacePos + 1)
    Loop While CommaSpacePos > 0
  End If
  ConcealParameterCommas = varCode

End Function

Public Function ConcealParameterSpaces(ByVal strCode As String) As String

  Dim SpacePos As Long

  'Replace any CommaSpace with comma in strInBrackets parameters
  'This allows CommaSpace delimited Dim,Public, Private to be safely detected without cutting in parameters
  'VB will automatically restore them
  SpacePos = GetSpacePos(strCode)
  Do
    If InCode(strCode, SpacePos) Then
      If EnclosedInBrackets(strCode, SpacePos) Then
        If Not SpaceBetweenReservedWords(strCode, SpacePos) Then
          strCode = Left$(strCode, SpacePos - 1) & Chr$(160) & Mid$(strCode, SpacePos + 1)
        End If
      End If
    End If
    SpacePos = GetSpacePos(strCode, SpacePos + 1)
  Loop While SpacePos > 0
  ConcealParameterSpaces = strCode

End Function

Public Function ContainsWholeWord(ByVal strSearch As String, _
                                  ByVal strFind As String, _
                                  Optional ByVal Start As Long = 1, _
                                  Optional CaseSensitive As VbCompareMethod = vbBinaryCompare) As Boolean

  'How it works
  'strSearch =string to search in
  'strFind   = string to look for
  'Start Optional Default =1; If you want to ignore earlier parts of strSearch set Start
  '                           (Could be used to scan through a string)
  'CaseSensitive Optional Default = vbBinaryCompare;
  '                          Set to vbTextCompare to find word in any case.
  'NOTE: MS have cheated by using something in InStr which they didn't provide in VB;
  'an optional 1st parameter. To use the optional Compare argument you also have to provide the Start value.
  'my Function places the Compare argument after the Find parameter and you only need to provide it if it is not default = 1
  '
  'Pad strSearch so you can detect first/last word
  '* accept anything (including nothing) before 1st * and after last *.
  '[!a-z!A-Z!À-Ö!Ø-ß!à-ö!ø-ÿ] accept any 1 character which is NOT(!) between a-z or A-Z or any of the high ASCII letter characters
  '
  'Small potential bug: if strFind = '*ry" then containsWholeWord will hit (* messes with the Like search)
  'could be exploited to build an end of word matcher if you want to try building a rhyming dictionary ;)

  strSearch = Mid$(strSearch, Start)
  If CaseSensitive = vbBinaryCompare Then
    ContainsWholeWord = SngSpace & strSearch & SngSpace Like "*" & strNotLetterFilter & strFind & strNotLetterFilter & "*"
    If Not ContainsWholeWord Then
      ContainsWholeWord = Left$(strSearch, Len(strFind) + 1) = strFind & " " And Left$(strSearch, 1) <> "'"
    End If
   Else
    ContainsWholeWord = LCase$(SngSpace & strSearch & SngSpace) Like "*" & strNotLetterFilter & LCase$(strFind) & strNotLetterFilter & "*"
  End If

End Function

Public Function CountSubString(ByVal varSearch As Variant, _
                               ByVal varFind As Variant) As Long

  CountSubString = UBound(Split(varSearch, varFind))

End Function

Public Function CountSubStringArray(ByVal varSearch As Variant, _
                                    ParamArray varFind() As Variant) As Long

  Dim F As Variant

  For Each F In varFind
    CountSubStringArray = CountSubStringArray + CountSubStringCode(varSearch, F)
  Next F

End Function

Public Function CountSubStringCode(ByVal varSearch As Variant, _
                                   ByVal varFind As Variant) As Long

  'v2.6.4 needed for line cont removal
  'so it doesn't count code in strings

  DisguiseLiteral varSearch, varFind, True
  CountSubStringCode = UBound(Split(varSearch, varFind))

End Function

Public Function CountSubStringImbalance(varTest As Variant, _
                                        strL As String, _
                                        strR As String) As Boolean

  CountSubStringImbalance = CountSubString(varTest, strL) <> CountSubString(varTest, strR)

End Function

Public Function CountSubStringWhole(ByVal varSearch As Variant, _
                                    ByVal varFind As Variant) As Long

  Dim I      As Long
  Dim arrStr As Variant

  'v2.7.0 improved and specialized CountSubString
  If InStr(varSearch, varFind) Then
    ExtractCode varSearch
    If LenB(varSearch) Then
      arrStr = Split(ExpandForDetection(strCodeOnly(varSearch)))
      For I = 0 To UBound(arrStr)
        If arrStr(I) = varFind Then
          CountSubStringWhole = CountSubStringWhole + 1
        End If
      Next I
    End If
  End If

End Function

Public Function DefNameNumber(ByVal strInput As String) As String

  Dim I As Long

  For I = 1 To Len(strInput)
    If IsNumeric(Mid$(strInput, I, 1)) Then
      DefNameNumber = DefNameNumber & Mid$(strInput, I, 1)
    End If
  Next I

End Function

Public Sub DisguiseLiteral(strSearch As Variant, _
                           ByVal HideMe As Variant, _
                           ByVal HideTShowF As Boolean)

  Dim LocalFind    As String
  Dim LocalReplace As String
  Dim FindPos      As Long
  Dim LDisguise    As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Replaces literal test words with rubbish so that Split can't find them
  'Call second time with HideME and Disguise reversed
  'after the string has been reassembled by Join to restore literals
  'Disguise is regenerated each time you call the routine to hide something( HideTShowF = True)
  '*If you change Disguise make sure it uses a '&' so formatting this code doesn't touch it
  If HideTShowF Then
    ' create a masking value for string literals
    Do
      LDisguise = RandomString(48, 122, 3, 6)
    Loop While InStr(LDisguise, strSearch) Or InStr(LDisguise, HideMe)
    DisguiseStack.Push HideMe
    DisguiseStack.Push LDisguise
    LocalFind = HideMe
    LocalReplace = LDisguise
   Else
    LocalFind = DisguiseStack.Pop
    LocalReplace = DisguiseStack.Pop
  End If
  FindPos = InStr(strSearch, LocalFind)
  If FindPos Then
    Do
      If FindPos Then
        If InLiteral(strSearch, FindPos) Then
          strSearch = Left$(strSearch, FindPos - 1) & LocalReplace & Mid$(strSearch, FindPos + Len(LocalFind))
        End If
      End If
      FindPos = InStr(FindPos + 1, strSearch, LocalFind)
    Loop While FindPos
  End If

End Sub

Public Sub Do_tbsFileToolClick()

  Dim I As Long

  With frm_CodeFixer.tbsFileTool
    .Move 0, 0
    For I = 0 To 2
      frm_CodeFixer.fraUndoTab(I).Caption = vbNullString
      frm_CodeFixer.fraUndoTab(I).Visible = False
    Next I
    frm_CodeFixer.fraUndoTab(.SelectedItem.Index - 1).Move .ClientLeft, .ClientTop
    frm_CodeFixer.fraUndoTab(.SelectedItem.Index - 1).Visible = True
  End With 'TABSTRIPUNDO

End Sub

Public Sub DoAllSelect()

  With frm_CodeFixer
    If .lsvAllControls.ListItems.Count Then
      WarningLabel
      GetCtrlDataLSV
      .lblOldName(1).Caption = CntrlDesc(LngCurrentControl).CDName
      SetNewName CntrlDesc(LngCurrentControl).CDName
      .cmdAutoLabel(5).Enabled = IsThisControlDeletable(LngCurrentControl)
    End If
  End With

End Sub

Public Function dofix(ByVal ModNum As Long, _
                      ByVal FixNum As Long) As Boolean

  If ModNum > -1 Then
    If Not ModDesc(ModNum).MDDontTouch Then
      dofix = FixData(FixNum).FixLevel > Off
      If Xcheck(XIgnoreCom) Then
        If FixData(FixNum).FixLevel = CommentOnly Then
          dofix = False
        End If
      End If
    End If
   Else
    dofix = True
  End If
  ' DoFix = True

End Function

Public Function EnclosedInBrackets(ByVal strSearch As String, _
                                   ByVal ChrPos As Long) As Boolean

  Dim LBracketCount As Long
  Dim RBracketCount As Long
  Dim MyStr         As String
  Dim LBit          As String
  Dim Rbit          As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Detect whether a Chracter Position is between brackets
  If ChrPos > 0 Then
    If Not JustACommentOrBlank(strSearch) Then
      If GetRightBracketPos(strSearch) > 0 Then
        If GetLeftBracketPos(strSearch) > 0 Then
          MyStr = strSearch
          If ExtractCode(MyStr) Then
            DisguiseLiteral MyStr, LBracket, True
            DisguiseLiteral MyStr, RBracket, True
            LBracketCount = CountSubString(MyStr, LBracket)
            RBracketCount = CountSubString(MyStr, RBracket)
            If LBracketCount = LBracketCount Then
              If LBracketCount > 0 Then
                LBit = Left$(MyStr, ChrPos)
                Rbit = Mid$(MyStr, ChrPos)
                LBracketCount = Abs(CountSubString(LBit, LBracket) - CountSubString(LBit, RBracket))
                RBracketCount = Abs(CountSubString(Rbit, LBracket) - CountSubString(Rbit, RBracket))
                If RBracketCount = LBracketCount Then
                  If LBracketCount > 0 Then
                    EnclosedInBrackets = True
                  End If
                End If
              End If
            End If
            DisguiseLiteral MyStr, LBracket, False
            DisguiseLiteral MyStr, RBracket, False
          End If
        End If
      End If
    End If
  End If

End Function

Public Function ExtractCode(varCode As Variant, _
                            Optional StrCom As String = vbNullString, _
                            Optional strSpace As String = vbNullString) As Boolean

  Dim strOrig As String

  'ver 1.1.51
  'extracts code only from soucre code
  'optionally returns or dumps the left padding spaces and comments
  'if there is no code it returns False and resets Varcode to original content
  strOrig = varCode
  If Len(varCode) Then
    StrCom = CommentClip(varCode)
    strSpace = SpaceOffsetClip(varCode)
    If Len(varCode) Then
      ExtractCode = True
     Else
      varCode = strOrig
    End If
  End If

End Function

Public Function FakeHungarian(ByVal strType As String) As String

  Dim strT As String

  strT = strType
  strT = Replace$(strT, "a", vbNullString)
  strT = Replace$(strT, "e", vbNullString)
  strT = Replace$(strT, "i", vbNullString)
  strT = Replace$(strT, "o", vbNullString)
  strT = Replace$(strT, "u", vbNullString)
  strT = LCase$(strT)
  Select Case Len(strT)
   Case 0
    strT = LCase$(strType)
    If Len(strT) > 2 Then
      FakeHungarian = Left$(strT, 3)
     Else
      FakeHungarian = strT & String$(3 - Len(strT), 95)
    End If
   Case Else
    If Len(strT) > 2 Then
      FakeHungarian = Left$(strT, 3)
     Else
      FakeHungarian = strT & String$(3 - Len(strT), 95)
    End If
  End Select
  If InStr(FakeHungarian, "___") Then
    FakeHungarian = Replace$(FakeHungarian, "___", "var")
  End If

End Function

Public Function Get_As_Pos(varSearch As Variant) As Long

  'gives a name to action making reading code more readable

  Get_As_Pos = InStr(1, varSearch, " As ")
  If Not InCode(varSearch, Get_As_Pos) Then
    Get_As_Pos = 0
  End If

End Function

Public Function GetCommaSpacePos(varSearch As Variant, _
                                 Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

  GetCommaSpacePos = InStr(StartAt, varSearch, CommaSpace)

End Function

Public Sub GetCtrlDataLSV()

  Dim Comp               As VBComponent
  Dim vbf                As VBForm
  Dim bSelectedArrayNAme As Boolean
  Dim strGuard           As String

  LngCurrentControl = GetTag(frm_CodeFixer.lsvAllControls)
  frm_CodeFixer.lblOldName(3).Caption = vbNullString
  If CntrlDesc(LngCurrentControl).CDBadType > 0 Then
    WarningLabel
    Select Case CntrlDesc(LngCurrentControl).CDBadType
     Case BNReserve, BNKnown, BNCommand
      WarningLabel "WARNING" & vbNewLine & _
       "Rename will also affect calls to this Resereved Word/Command. You will have to fix by hand."
     Case BNMultiForm
      WarningLabel "WARNING" & vbNewLine & _
       "May be errors due to name on 2+ forms. Run code (Ctrl+F5) to find and change back."
    End Select
  End If
  On Error GoTo CutFail
  Set Comp = GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm)
  If IsComponent_ControlHolder(Comp) Then
    ActivateDesigner Comp, vbf, True ' False
    On Error GoTo CutFail
    SelectOneControlSTR vbf, CntrlDesc(LngCurrentControl).CDName, CntrlDesc(LngCurrentControl).CDIndex
    'v2.4.9 don't add old name to list (may be generated by some of later suggest engines
    strGuard = "*" & CntrlDesc(LngCurrentControl).CDName & "*"
    SuggestHungarianName bSelectedArrayNAme, strGuard
    SuggestControlName strGuard
    With frm_CodeFixer
      '.cmdEditSuggest.Enabled = Len(.lblOldName(3).Caption) And AutoReName = False
      If Not bSelectedArrayNAme Then
        If Len(.lblOldName(3).Caption) Then
          .txtCtrlNewName = .lblOldName(3).Caption
        End If
      End If
    End With 'fCodeFix
  End If
CutFail:

End Sub

Public Function GetLeftBracketPos(varSearch As Variant, _
                                  Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

  GetLeftBracketPos = InStr(StartAt, varSearch, LBracket)
  'ver 1.1.21
  'v2.3.8 Thanks Paul Caton 'stops brackets in literal strings from being counted
  Do While Not InCode(varSearch, GetLeftBracketPos) And GetLeftBracketPos > 0
    GetLeftBracketPos = InStr(GetLeftBracketPos + 1, varSearch, RBracket) '0
  Loop

End Function

Public Function GetRightBracketPos(varSearch As Variant, _
                                   Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

  GetRightBracketPos = InStr(StartAt, varSearch, RBracket)
  'ver 1.1.21
  'v2.3.8 Thanks Paul Caton 'stops brackets in literal strings from being counted
  Do While Not InCode(varSearch, GetRightBracketPos) And GetRightBracketPos > 0
    GetRightBracketPos = InStr(GetRightBracketPos + 1, varSearch, RBracket) '0
  Loop

End Function

Public Function GetRoutineName(ByVal strTest As String) As String

  If MultiLeft(strTest, True, "Private ", "Public ", "Friend ", "Static ") Then
    strTest = Mid$(strTest, GetSpacePos(strTest) + 1)
  End If
  If MultiLeft(strTest, True, "Sub ", "Function ", "Property ") Then
    If LeftWord(strTest) = "Property" Then
      strTest = Mid$(strTest, GetSpacePos(strTest) + 1)
      strTest = Mid$(strTest, GetSpacePos(strTest) + 1)
     Else
      strTest = Mid$(strTest, GetSpacePos(strTest) + 1)
    End If
    If GetLeftBracketPos(strTest) Then
      strTest = Replace$(strTest, LBracket, SpacePad(LBracket))
    End If
  End If
  GetRoutineName = LeftWord(strTest)

End Function

Public Function GetSpacePos(varSearch As Variant, _
                            Optional StartAt As Long = 1) As Long

  GetSpacePos = InStr(StartAt, varSearch, SngSpace)

End Function

Public Function HasLineCont(ByVal varTest As Variant) As Boolean

  HasLineCont = SmartRight(varTest, ContMark)

End Function

Public Sub HideInitiliser()

  On Error Resume Next
  Unload frm_CodeFixer
  On Error GoTo 0

End Sub

Public Function InCode(ByVal varSearch As Variant, _
                       ByVal TestPos As Long) As Boolean

  If TestPos Then
    If InComment(varSearch, TestPos) Then
      InCode = False
     ElseIf InLiteral(varSearch, TestPos) Then
      InCode = False
     ElseIf InTimeLiteral(varSearch) Then
      InCode = False
     Else
      InCode = True
    End If
  End If

End Function

Public Function InComment(ByVal varSearch As Variant, _
                          ByVal TPos As Long) As Boolean

  Dim Possible As Long
  Dim arrTmp   As Variant
  Dim OPos     As Long
  Dim NPos     As Long
  Dim I        As Long

  'v2.0.5 fixed it was not hitting properly
  Possible = InStr(varSearch, SQuote)
  If Possible Then
    Do
      If Possible > TPos Then ' the test point is les than the possibel point
        Possible = 0
        Exit Do
      End If
      If TPos > Len(varSearch) Then ' the test point is beyond the len of string
        Possible = 0
        Exit Do
      End If
      If InLiteral(varSearch, Possible, False) Then
        Possible = InStr(Possible + 1, varSearch, SQuote)
      End If
      If Possible = 0 Or Possible < TPos Then
        Exit Do
      End If
    Loop While InLiteral(varSearch, Possible, False) And Possible > 0
    Possible = InStr(varSearch, SQuote) < TPos
    If Possible Then
      arrTmp = Split(varSearch, SQuote)
      For I = LBound(arrTmp) To UBound(arrTmp)
        NPos = NPos + 1 + Len(arrTmp(I))
        If Between(OPos, TPos, NPos) Then
          InComment = Not InLiteral(varSearch, OPos, False)
          Exit For
        End If
        OPos = NPos
        If OPos >= TPos Then
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function InLiteral(ByVal varSearch As Variant, _
                          ByVal TPos As Long, _
                          Optional ByVal CommentTest As Boolean = True) As Boolean

  Dim Possible As Long
  Dim arrTest  As Variant
  Dim I        As Long
  Dim OPos     As Long
  Dim NPos     As Long

  Possible = InStr(varSearch, DQuote)
  If Possible Then
    If Possible = TPos Then
      InLiteral = Not InComment(varSearch, TPos)
     Else
      arrTest = Split(varSearch, DQuote)
      For I = LBound(arrTest) To UBound(arrTest)
        NPos = NPos + 1 + Len(arrTest(I))
        If NPos > TPos Then
          If IsOdd(I) Then
            If Between(OPos, TPos, NPos) Then
              If CommentTest Then
                InLiteral = Not InComment(varSearch, TPos)
               Else
                ' this is only to stop nocomment creating recursive overflow
                InLiteral = True
              End If
              Exit For
            End If
          End If
        End If
        OPos = NPos
        If OPos > TPos Then
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function InstrArray(varSearch As Variant, _
                           ParamArray varFind() As Variant) As Long

  Dim VarTmp As Variant

  For Each VarTmp In varFind
    If InStr(varSearch, VarTmp) Then
      InstrArray = True
      Exit For 'unction
    End If
  Next VarTmp

End Function

Public Function InstrArrayHard(varSearch As Variant, _
                               arrFind As Variant) As Long

  Dim I As Long

  For I = LBound(arrFind) To UBound(arrFind)
    If InStr(varSearch, arrFind(I)) Then
      InstrArrayHard = True
      Exit For 'unction
    End If
  Next I

End Function

Public Function InstrArrayWholeWord(varSearch As Variant, _
                                    ParamArray varFind() As Variant) As Long

  Dim VarTmp As Variant

  For Each VarTmp In varFind
    If ContainsWholeWord(varSearch, VarTmp) Then
      InstrArrayWholeWord = True
      Exit For 'unction
    End If
  Next VarTmp

End Function

Public Function InstrAtPosition(ByVal varSearch As Variant, _
                                ByVal varFind As Variant, _
                                ByVal AtLocation As InstrLocations, _
                                Optional WholeWord As Boolean = True) As Boolean

  
  Dim SizeOfSearch As Long
  Dim I            As Long
  Dim TmpA         As Variant

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Return True or False
  'Parameteres:
  'varSearch, search in
  'varFInd, search for
  'AtLocation, test that varFInd is in varSearch at this position
  '               Left, Right, LeftOr2nd, Middle(=exists but not at Left or Right)
  '               Exact(same as varSearch=varFInd), None and Any present in no/any position
  'WholeWord, only if space delimited. While not safe for string literals with punctuation this is always true for code
  '
  'This routine only searches and finds in code all literals are excluded
  If LenB(varSearch) Then
    If LenB(varFind) Then
      varSearch = strCodeOnly(varSearch)
      If LenB(varSearch) Then
        If LenB(varFind) Then
          TmpA = Split(varSearch)
          SizeOfSearch = UBound(TmpA)
          ' make this optional trigger if you need to search comments (unlikely)
          DisguiseLiteral varSearch, varFind, True
          varSearch = Trim$(varSearch)
          If LenB(varSearch) Then
            Select Case AtLocation
             Case IpExact
              InstrAtPosition = varSearch = varFind
             Case IpLeft
              If varSearch = varFind Then
                InstrAtPosition = True
               ElseIf SmartLeft(varSearch, varFind) Then
                If WholeWord Then
                  InstrAtPosition = PosInStringWholeWord(varSearch, varFind) = 1
                 Else
                  InstrAtPosition = True
                End If
              End If
             Case IpMiddle
              InstrAtPosition = InStr(varSearch, varFind) > 0
              If WholeWord Then
                InstrAtPosition = Between(1, InStrWholeWordRX(varSearch, varFind), Len(varSearch) - Len(varFind) - 1)
               Else
                InstrAtPosition = True
              End If
             Case IpRight
              If varSearch = varFind Then
                InstrAtPosition = True
               ElseIf SmartRight(varSearch, varFind) Then
                If WholeWord Then
                  InstrAtPosition = InStrWholeWordRX(varSearch, varFind, Len(varSearch) - Len(varFind) - 1) < Len(varSearch) - Len(varFind) - 1
                 Else
                  InstrAtPosition = True
                End If
                InstrAtPosition = True
              End If
             Case ip2nd
              InstrAtPosition = PosInStringWholeWord(varSearch, varFind) = 2
              If WholeWord Then
                InstrAtPosition = PosInStringWholeWord(varSearch, varFind) = 2
               Else
                If SizeOfSearch > 0 Then
                  TmpA(0) = vbNullString
                  InstrAtPosition = SmartLeft(Trim$(Join(TmpA)), varFind)
                End If
              End If
             Case ip3rd
              If WholeWord Then
                InstrAtPosition = PosInStringWholeWord(varSearch, varFind) = 3
               Else
                If SizeOfSearch > 1 Then
                  TmpA(0) = vbNullString
                  TmpA(1) = vbNullString
                  InstrAtPosition = SmartLeft(Trim$(Join(TmpA)), varFind)
                End If
              End If
             Case ip3rdorGreater
              If WholeWord Then
                InstrAtPosition = PosInStringWholeWord(varSearch, varFind) >= 3
               Else
                If SizeOfSearch > 1 Then
                  TmpA(0) = vbNullString
                  TmpA(1) = vbNullString
                  For I = 2 To UBound(TmpA)
                    InstrAtPosition = SmartLeft(Trim$(Join(TmpA)), varFind)
                    If InstrAtPosition Then
                      Exit For
                     Else
                      TmpA(I) = vbNullString
                    End If
                  Next I
                End If
              End If
             Case ipLeftOr2nd
              InstrAtPosition = InstrAtPosition(varSearch, varFind, IpLeft, WholeWord)
              If Not InstrAtPosition Then
                InstrAtPosition = InstrAtPosition(varSearch, varFind, ip2nd, WholeWord)
              End If
             Case ip2ndOr3rd
              InstrAtPosition = InstrAtPosition(varSearch, varFind, ip2nd, WholeWord)
              If Not InstrAtPosition Then
                InstrAtPosition = InstrAtPosition(varSearch, varFind, ip3rd, WholeWord)
              End If
             Case ipLeftOr2ndOr3rd
              InstrAtPosition = InstrAtPosition(varSearch, varFind, ipLeftOr2nd, WholeWord)
              If Not InstrAtPosition Then
                InstrAtPosition = InstrAtPosition(varSearch, varFind, ip3rd, WholeWord)
              End If
             Case IpNone
              If varSearch = varFind Then
                InstrAtPosition = False
               ElseIf InStr(varSearch, varFind) > 0 Then
                InstrAtPosition = InStrWholeWordRX(varSearch, varFind) = 0
               ElseIf SmartLeft(varSearch, varFind) Then
                InstrAtPosition = InStrWholeWordRX(varSearch, varFind) = 0
               ElseIf SmartRight(varSearch, varFind) Then
                InstrAtPosition = InStrWholeWordRX(varSearch, varFind) = 0
               Else
                InstrAtPosition = True
              End If
             Case ipAny
              If varSearch = varFind Then
                InstrAtPosition = True
               ElseIf InStr(varSearch, varFind) > 0 Then
                If WholeWord Then
                  InstrAtPosition = PosInStringWholeWord(varSearch, varFind) <> 0
                 Else
                  InstrAtPosition = True
                End If
               ElseIf SmartLeft(varSearch, varFind) Then
                If WholeWord Then
                  InstrAtPosition = PosInStringWholeWord(varSearch, varFind) <> 0
                 Else
                  InstrAtPosition = True
                End If
               ElseIf SmartRight(varSearch, varFind) Then
                If WholeWord Then
                  InstrAtPosition = PosInStringWholeWord(varSearch, varFind) <> 0
                 Else
                  InstrAtPosition = True
                End If
               Else
                InstrAtPosition = False
              End If
            End Select
          End If
          DisguiseLiteral varSearch, varFind, False
        End If
      End If
    End If
  End If

End Function

Public Function InstrAtPositionArray(ByVal strSearch As String, _
                                     ByVal AtLocation As InstrLocations, _
                                     ByVal WholeWord As Boolean, _
                                     ParamArray FindA() As Variant) As Boolean

  Dim findMember As Variant

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'check that any member of FindA is space delimited part of StrSearch at position AtLocation
  'See InstrAtPosition for parameter details
  For Each findMember In FindA
    If LenB(findMember) Then
      If InstrAtPosition(strSearch, findMember, AtLocation, WholeWord) Then
        InstrAtPositionArray = True
        Exit For
      End If
    End If
  Next findMember

End Function

Public Function InstrAtPositionSetArray(ByVal strSearch As String, _
                                        ByVal AtLocation As InstrLocations, _
                                        ByVal WholeWord As Boolean, _
                                        arrFind As Variant) As Boolean

  Dim findMember As Variant

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'check that any member of FindA is space delimited part of StrSearch at position AtLocation
  'See InstrAtPosition for parameter details
  For Each findMember In arrFind
    If LenB(findMember) Then
      If InStr(strSearch, findMember) Then
        If InstrAtPosition(strSearch, findMember, AtLocation, WholeWord) Then
          If WholeWord Then
            If InStrWholeWord(strSearch, findMember) Then
              'v2.3.2 safer test
              'v2.3.1 updted for better testing
              InstrAtPositionSetArray = True
            End If
           Else
            InstrAtPositionSetArray = True
          End If
          Exit For
        End If
      End If
    End If
  Next findMember

End Function

Public Function InStrCode(ByVal varSearch As Variant, _
                          ByVal varFind As Variant, _
                          Optional ByVal StartPos As Long) As Long

  Dim TestPos  As Long
  Dim TestPos1 As Long

  If StartPos > 0 Then
    TestPos1 = StartPos
   Else
    TestPos1 = 1
  End If
  varSearch = strCodeOnly(varSearch)
  TestPos = InStr(TestPos1, varSearch, varFind)
  Do While TestPos
    If InCode(varSearch, TestPos) Then
      InStrCode = TestPos
      Exit Do 'Function
    End If
    TestPos = InStr(TestPos + 1, varSearch, varFind)
  Loop
  'InStrCode = 0

End Function

Public Function InStrWholeWord(varSearch As Variant, _
                               varFind As Variant, _
                               Optional ByVal WholeWord As Boolean = True) As Long

  Dim bFound As Boolean
  Dim TPos   As Long

  If InStr(varSearch, varFind) Then
    If WholeWord Then
      If varSearch = varFind Then
        bFound = True
        InStrWholeWord = 1
       ElseIf SmartLeft(varSearch, varFind) And IsPunct(Mid$(varSearch, 1 + Len(varFind), 1)) Then
        InStrWholeWord = IIf(IsPunct(Mid$(varSearch, InStr(varSearch, varFind) + Len(varFind), 1)), 1, 0)
       ElseIf SmartRight(varSearch, varFind) And IsPunct(Mid$(varSearch, Len(varSearch) - Len(varFind), 1)) Then
        InStrWholeWord = IIf(IsPunct(Mid$(varSearch, Len(varSearch) - Len(varFind), 1)), Len(varSearch) - Len(varFind) + 1, 0)
       Else
        TPos = InStr(varSearch, varFind)
        If TPos = 1 Then
          TPos = InStr(TPos + 1, varSearch, varFind)
        End If
        Do While TPos
          bFound = IsPunct(Mid$(varSearch, TPos + Len(varFind), 1)) And IsPunct(Mid$(varSearch, TPos - 1, 1))
          If Not bFound Then
            TPos = InStr(TPos + 1, varSearch, varFind)
           Else
            InStrWholeWord = TPos
            Exit Do
          End If
        Loop
      End If
     Else
      InStrWholeWord = InStr(varSearch, varFind)
    End If
  End If

End Function

Public Function InStrWholeWordRX(ByVal strSearch As String, _
                                 ByVal strFind As String, _
                                 Optional Start As Long = 1, _
                                 Optional CaseSensitive As VbCompareMethod = vbBinaryCompare) As Long

  Dim TPos As Long

  If Start > 1 Then
    If Start < Len(strSearch) Then
      strSearch = Mid$(strSearch, Start)
     Else
      Exit Function
    End If
  End If
  If ContainsWholeWord(strSearch, strFind, 1, CaseSensitive) Then
    If Not CaseSensitive = vbBinaryCompare Then
      strSearch = LCase$(strSearch)
      strFind = LCase$(strFind)
    End If
    TPos = InStr(strSearch, strFind)
    'Get inital test point then
    If TPos Then
      Do
        If TPos = 1 Then
          If strSearch = strFind Then
            InStrWholeWordRX = 1
            Exit Do
          End If
          If Mid$(strSearch, Len(strFind) + 1, 1) Like strNotLetterFilter Then
            InStrWholeWordRX = 1
            Exit Do
          End If
         ElseIf Mid$(strSearch, TPos - 1, 1) Like strNotLetterFilter Then
          If TPos + Len(strFind) - 1 = Len(strSearch) Then
            InStrWholeWordRX = TPos
            Exit Do
           ElseIf Mid$(strSearch, TPos + Len(strFind), 1) Like strNotLetterFilter Then
            InStrWholeWordRX = TPos
            Exit Do
          End If
        End If
        TPos = InStr(TPos + 1, strSearch, strFind)
      Loop While TPos
    End If
  End If
  If strFind = strSearch Then
    InStrWholeWordRX = 1
  End If
  If InStrWholeWordRX Then
    InStrWholeWordRX = InStrWholeWordRX + Start - 1
  End If

End Function

Private Function InTimeLiteral(ByVal varSearch As Variant) As Boolean

  Dim P1 As Long
  Dim P2 As Long
  Dim Ps As Long

  If CountSubString(varSearch, "#") > 1 Then
    Ps = InStr(varSearch, "#")
    Do
      Do
        P1 = InStr(Ps, varSearch, "#")
        P2 = InStr(P1 + 1, varSearch, "#")
        Ps = P2
        If Ps = 0 Then
          Exit Do
        End If
      Loop While InLiteral(varSearch, P1)
      If P1 > 0 Then
        If Not InComment(varSearch, P1) Then
          If P2 > P1 Then
            If Not InComment(varSearch, P2) Then
              InTimeLiteral = IsDate(Mid$(varSearch, P1, P2))
            End If
          End If
        End If
      End If
      If Ps = 0 Then
        Exit Do
      End If
    Loop While P1 > 0
  End If

End Function

Public Function IsInArray(ByVal FindValue As Variant, _
                          ByVal arrSearch As Variant) As Boolean

  'VERY FAST function to find a value in an unsorted Array
  'By Brian Gillham
  'Versions:  VB5, VB6
  'Posted:  4/18/2001

  On Error GoTo LocalError
  If ArrayHasContents(arrSearch) Then
    IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
    On Error GoTo 0
  End If
LocalError:

End Function

Public Function IsOdd(ByVal N As Variant) As Boolean

  'Here's a efficient IsEven function
  'By Sam Hills
  'shills@bbll.com
  '*If you want an IsOdd function, just omit the Not.
  '        IsEven =not -(n And 1)

  IsOdd = -(N And 1)

End Function

Public Function IsPunctExcept(ByVal strTest As String, _
                              ByVal strExcept As String) As Boolean

  'Detect punctuation

  If InStr(strExcept, strTest) = 0 Then
    If IsNumeral(strTest) Then
      IsPunctExcept = False
     Else
      IsPunctExcept = Not IsAlphaIntl(strTest)
    End If
  End If

End Function

Public Sub KeepOnScreen()

  'Prevents setting Code Fixer to appear off screen (If you change screen resolution for example)
  'You can still move the tool off the edge of the screen but when you change Tabs the tool will jump back to Default position

  With frm_CodeFixer
    If .WindowState = vbNormal Then
      If .Height + .Top > Screen.Height Then
        .Top = Screen.Height - .Height
      End If
      If .Left + .Width > Screen.Width Then
        .Left = Screen.Width - .Width
      End If
      If .Top < 0 Then
        .Top = 0
      End If
      If .Left < 0 Then
        .Left = 0
      End If
    End If
  End With
  DoEvents

End Sub

Public Function LeftWord(ByVal varChop As Variant) As String

  If LenB(varChop) Then
    LeftWord = Split(varChop)(0)
  End If

End Function

Public Sub LoadUserSettings()

  Dim TurnOn                          As Long

  UserSettings = GetSetting(AppDetails, "Options", "UserSet", DefAvgSettings)
  'v2.5.1 Crash can destroy the UserSettings
  If Len(UserSettings) = 0 Then
    UserSettings = GetSetting(AppDetails, "Options", "UserSetSafe", DefAvgSettings)
   Else
    SaveSetting AppDetails, "Options", "UserSetSafe", UserSettings
  End If
  OldSettingsTest
  lngBigProcLines = GetSetting(AppDetails, "Options", "BigProc", 50)
  lngLineLength = GetSetting(AppDetails, "Options", "LngLine", 100)
  lngElseIndent = GetSetting(AppDetails, "Options", "ElseIndnt", 1)
  lngCaseIndent = GetSetting(AppDetails, "Options", "CaseIndnt", 1)
  Select Case UserSettings
   Case DefNoSettings
    TurnOn = 0
   Case DefMinSettings
    TurnOn = 1
   Case DefAvgSettings
    TurnOn = 2
   Case DefMaxSettings
    TurnOn = 3
   Case Else
    TurnOn = 4
  End Select
  With frm_FindSettings 'CodeFixer
    '.sldLargeProc.Value = lngBigProcLines
    .UpdLongLine.Value = lngLineLength
    .UpDLongProc.Value = lngBigProcLines
    'v2.3.8 Thanks Paul Caton, This should fix indent setting
    .OptElseIndent(lngElseIndent) = True
    .OptCaseIndent(lngCaseIndent) = True
    '.lbBigProc.Caption = "Proc code lines(40-100) " & lngBigProcLines
    'UserTop = Val(GetSetting(AppDetails, "Options", "Top", CStr((Screen.Height - .Height) / 2)))
    'UserLeft = Val(GetSetting(AppDetails, "Options", "Left", CStr((Screen.Width - .Width) / 2)))
    StandardSettings TurnOn
    IndentSize = GetFullTabWidth '/ 2
  End With

End Sub

Public Function LStrip(ByVal strInput As String, _
                       ByVal strStrip As String) As String

  'Strip strInput specified character from start of string

  If Left$(strInput, 1) = strStrip Then
    Do
      strInput = Mid$(strInput, 2)
    Loop While Left$(strInput, 1) = strStrip
  End If
  LStrip = strInput

End Function

Public Function ModuleHasCode(Wmod As CodeModule) As Boolean

  ModuleHasCode = Wmod.CountOfLines - Wmod.CountOfDeclarationLines > 0

End Function

Public Function ModuleLevelToDim(strTest As String, _
                                 ByVal strOrig As String) As Boolean

  Dim Proj            As VBProject
  Dim Comp            As VBComponent
  Dim GuardLine       As Long
  Dim strCodeLine     As String
  Dim strProcName     As String
  Dim strPrevProcName As String
  Dim StartLine       As Long
  Dim Hit             As Boolean
  Dim HomeModule      As String

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        GuardLine = 0
        strPrevProcName = vbNullString
        With Comp
          Do While .CodeModule.Find(strTest, StartLine, 1, -1, -1, True, True)
            strCodeLine = .CodeModule.Lines(StartLine, 1)
            'Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strTest, strCodeLine, StartLine)
            If GuardLine > 0 Then
              If GuardLine > StartLine Then
                Exit Do
              End If
            End If
            strCodeLine = Trim$(strCodeLine)
            If strCodeLine = strOrig Then
              HomeModule = .Name
             Else 'End If
              If strCodeLine <> strOrig Then
                Hit = True
                strProcName = GetProcName(.CodeModule, StartLine)
                If strPrevProcName = strProcName Then
                  GoTo Skip
                 Else
                  If LenB(strPrevProcName) Then
                    ModuleLevelToDim = False
                    Exit Do
                  End If
                End If
               Else
                strPrevProcName = strProcName
                ModuleLevelToDim = strCodeLine Like "For " & strTest & " = *" Or (strCodeLine Like strTest & " = *" And Not (strCodeLine Like strTest & EqualInCode & strTest & " *"))
                If Not ModuleLevelToDim Then
                  Exit Do
                End If
              End If
            End If
Skip:
            StartLine = StartLine + 1
            GuardLine = StartLine
          Loop
        End With 'Comp
      End If
      If Not ModuleLevelToDim And Hit Then
        If LeftWord(strOrig) = "Private" Then
          Exit For
         ElseIf LeftWord(strOrig) = "Public" Then
          If HomeModule = Comp.Name Then
            Hit = False
          End If
          Exit For
        End If
      End If
    Next Comp
    If Not ModuleLevelToDim And Hit Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Public Function MultiLeft(ByVal varSearch As Variant, _
                          ByVal CaseSensitive As Boolean, _
                          ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If Not CaseSensitive Then
      FindIt = LCase$(FindIt)
    End If
    If Left$(varSearch, Len(FindIt)) = FindIt Then
      MultiLeft = True
      Exit For 'unction
    End If
  Next FindIt

End Function

Public Function MultiLeft2(ByVal varSearch As Variant, _
                           ByVal CaseSensitive As Boolean, _
                           Afind As Variant) As Boolean

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If CaseSensitive Then
      If Left$(varSearch, Len(FindIt)) = FindIt Then
        MultiLeft2 = True
        Exit For 'unction
      End If
     Else
      If Left$(varSearch, Len(FindIt)) = LCase$(FindIt) Then
        MultiLeft2 = True
        Exit For 'unction
      End If
    End If
  Next FindIt

End Function

Public Function MultiRight(ByVal varSearch As Variant, _
                           ByVal CaseSensitive As Boolean, _
                           ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  'CaseSensitive was added to solve a problem with hand coding of standard VB routines with wrong case
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If Not CaseSensitive Then
      FindIt = LCase$(FindIt)
    End If
    If Right$(varSearch, Len(FindIt)) = FindIt Then
      MultiRight = True
      Exit For 'unction
    End If
  Next FindIt

End Function

Public Sub OldSettingsKill()

  Dim I  As Long
  Dim J  As Long
  Dim K  As Long
  Dim UJ As Long
  Dim UK As Long

  'ver 1.1.77 thanks to Randy Giese who suggested this
  With App
    UJ = .Minor
    UK = .Revision
  End With
  For I = 0 To App.Major
    For J = 0 To UJ
      For K = 0 To UK  ' don't delete newest
        'v2.5.1 fixe to get all previous versions
        If App.ProductName & " V" & I & "." & J & "." & K <> AppDetails Then
          On Error GoTo NoSuch
          'X = GetAllSettings(App.ProductName & " V" & I & "." & J & "." & K, "Options")
          If ArrayHasContents(GetAllSettings(App.ProductName & " V" & I & "." & J & "." & K, "Options")) Then
            ' only delete if it exists
            DeleteSetting App.ProductName & " V" & I & "." & J & "." & K
          End If
        End If
NoSuch:
        Err.Number = 0
      Next K
    Next J
  Next I
  OldSettingsTest

End Sub

Private Sub OldSettingsTest()

  Dim I  As Long
  Dim J  As Long
  Dim K  As Long
  Dim UJ As Long
  Dim UK As Long
  Dim X  As Variant

  'ver 1.1.77 thanks to Randy Giese who suggested this
  With App
    UJ = .Minor
    UK = .Revision
  End With
  PrevSettings = Split("") 'safety
  'ver 2.0.3 changed to track .0 minor and .0 revisions
  For I = 0 To App.Major
    For J = 0 To UJ
      For K = 0 To UK - 1 'don't detect newest
        On Error GoTo NoSuch
        X = GetAllSettings(App.ProductName & " V" & I & "." & J & "." & K, "Options")
        If ArrayHasContents(X) Then
          PrevSettings = X ' collect settings in case you need them for updating
        End If
NoSuch:
        Err.Number = 0
      Next K
    Next J
  Next I
  frm_FindSettings.fraOldReg.Visible = ArrayHasContents(PrevSettings)

End Sub

Public Sub OldSettingsUpdate()

  Dim I As Long

  'ver 1.1.77 thanks to Randy Giese who suggested this
  For I = LBound(PrevSettings) To UBound(PrevSettings)
    SaveSetting AppDetails, "Options", PrevSettings(I, 0), PrevSettings(I, 1)
  Next I
  mObjDoc.Safe_MsgBox "WARNING" & vbNewLine & _
                    "Settings from Previous version have been applied." & vbNewLine & _
                    "If the layout or number of settings has changed there may be errors so check and reset where necessary.", vbExclamation
  LoadUserSettings
  Xcheck.LoadCheck
  'LoadToolPosStyle
  PlaceTool
  OldSettingsKill
  OldSettingsTest

End Sub

Public Sub PageLaunch(Optional ByVal bShowForm As Boolean = True)

  'v 2.1.5 added usercontrol test
  'v2.3.7 fixed Control tool not launching properly

  TabPreActivate
  If bShowForm Then
    frm_CodeFixer.WindowState = vbNormal
    SizeToFrame
    With frm_CodeFixer.frapage(FrameActive)
      .Move 0, 0
      frm_CodeFixer.Height = .Height
      .Visible = True
    End With
    PlaceTool
    frm_CodeFixer.Show
    frm_CodeFixer.Refresh
  End If
  TabPostActivate

End Sub

Public Sub PlaceTool()

  KeepOnScreen

End Sub

Public Function PosInStringWholeWord(ByVal strSearch As String, _
                                     ByVal strFind As String, _
                                     Optional Start As Long = 1, _
                                     Optional CaseSensitive As VbCompareMethod = vbBinaryCompare) As Long

  Dim TPos As Long

  TPos = InStrWholeWordRX(strSearch, strFind, Start, CaseSensitive)
  If TPos Then
    PosInStringWholeWord = UBound(Split(Left$(strSearch, TPos))) + 1
  End If

End Function

Public Function RandomString(ByVal iLowerBoundAscii As Long, _
                             ByVal iUpperBoundAscii As Long, _
                             ByVal lLowerBoundLength As Long, _
                             ByVal lUpperBoundLength As Long) As String

  Dim sHoldString As String
  Dim LCount      As Long

  '      --Eric Lynn, Ballwin, Missouri
  '        VBPJ TechTips 7th Edition
  'Verify boundaries
  If iLowerBoundAscii < 0 Then
    iLowerBoundAscii = 0
  End If
  If iLowerBoundAscii > 255 Then
    iLowerBoundAscii = 255
  End If
  If iUpperBoundAscii < 0 Then
    iUpperBoundAscii = 0
  End If
  If iUpperBoundAscii > 255 Then
    iUpperBoundAscii = 255
  End If
  If lLowerBoundLength < 0 Then
    lLowerBoundLength = 0
  End If
  'Set a random length
  'Create the random string
  For LCount = 1 To Int((CDbl(lUpperBoundLength) - CDbl(lLowerBoundLength) + 1) * Rnd + lLowerBoundLength)
    sHoldString = sHoldString & Chr$(Int((iUpperBoundAscii - iLowerBoundAscii + 1) * Rnd + iLowerBoundAscii))
  Next LCount
  RandomString = sHoldString

End Function

Public Function ReplaceArray(ByVal strInput As String, _
                             ParamArray VRep() As Variant) As String

  Dim I As Long

  'this function replaces multiple calls to Replace
  ReplaceArray = strInput
  For I = LBound(VRep) To UBound(VRep) Step 2
    If Len(VRep(I)) Then
      'v2.4.4 improved for multiple hits
      ReplaceArray = Replace$(ReplaceArray, VRep(I), VRep(I + 1))
    End If
  Next I

End Function

Public Sub ReplaceName(ByVal strOld As String, _
                       ByVal Strnew As String)

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim strKill      As String
  Dim StartLine    As Long
  Dim bDummy       As Boolean

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        ModuleMessage Comp, CurCompCount
        StartLine = 1
        With Comp
          Do While .CodeModule.Find(strOld, StartLine, 1, -1, -1, True, True)
            strKill = .CodeModule.Lines(StartLine, 1)
            '          Do While GetWholeCaseMatchCodeLine(Proj.Name, .Name, strOld, strKill, StartLine)
            'strKill = .CodeModule.Lines(StartLine, 1)
            'strKill = Replace(strKill, StrOld, strnew)
            ' .CodeModule.ReplaceLine StartLine, strKill
            WholeWordReplacer strKill, strOld, Strnew, bDummy
            .CodeModule.ReplaceLine StartLine, strKill
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

Public Function RSpacePad(varA As Variant) As String

  RSpacePad = varA & SngSpace

End Function

Public Function RStrip(ByVal strInput As String, _
                       ByVal strStrip As String) As String

  'Strip strInput specified character from end of string

  If Right$(strInput, 1) = strStrip Then
    Do
      strInput = Left$(strInput, Len(strInput) - 1)
    Loop While Right$(strInput, 1) = strStrip
  End If
  RStrip = strInput

End Function

Public Function Safe_ReplaceArray(ByVal strInput As String, _
                                  ParamArray VRep() As Variant) As String

  Dim I As Long

  'v2.8.8 thanks to Anele Mbanga whose code let me track down a bug in this
  'replaced REplace command with Safe_Replace so it ignores string literals
  'this is a specialized version of ReplaceArray only called by 'ExpandForDetection
  'this function replaces multiple calls to Replace
  Safe_ReplaceArray = strInput
  For I = LBound(VRep) To UBound(VRep) Step 2
    If Len(VRep(I)) Then
      'v2.4.4 improved for multiple hits
      Safe_ReplaceArray = Safe_Replace(Safe_ReplaceArray, VRep(I), VRep(I + 1))
    End If
  Next I

End Function

Private Sub SelectOneControlSTR(vbfrm As VBForm, _
                                ByVal strCName As String, _
                                ByVal lngIndex As Long)

  Dim Ctrl As Variant

  With vbfrm
    DeselectAll .VBControls
    For Each Ctrl In .VBControls
      If Ctrl.Properties("name").Value = strCName Then
        If Ctrl.Properties("index").Value = lngIndex Then
          Ctrl.InSelection = True
          Exit For
        End If
      End If
    Next Ctrl
  End With

End Sub

Public Function SmartLeft(ByVal varSearch As Variant, _
                          varFind As Variant, _
                          Optional ByVal CaseSensitive As Boolean = True) As Boolean

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance

  If Len(varSearch) Then
    If Len(varFind) Then
      If Not CaseSensitive Then
        varSearch = LCase$(varSearch)
        varFind = LCase$(varFind)
      End If
      SmartLeft = Left$(varSearch, Len(varFind)) = varFind
    End If
  End If

End Function

Public Function SmartRight(ByVal varSearch As Variant, _
                           ByVal varFind As Variant, _
                           Optional ByVal CaseSensitive As Boolean = True) As Boolean

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  'CaseSensitive was added to solve a problem with hand coding of standard VB routines with wrong case

  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
    varFind = LCase$(varFind)
  End If
  SmartRight = Right$(varSearch, Len(varFind)) = varFind

End Function

Private Function SpaceBetweenReservedWords(ByVal strSearch As String, _
                                           ByVal ChrPos As Long) As Boolean

  Dim I     As Long
  Dim J     As Long
  Dim LSPos As Long

  If Not JustACommentOrBlank(strSearch) Then
    For I = 0 To UBound(ArrQMaintainSpacing)
      LSPos = ChrPos - Len(ArrQMaintainSpacing(I))
      If LSPos > 0 Then
        If Mid$(strSearch, LSPos, Len(ArrQMaintainSpacing(I))) = ArrQMaintainSpacing(I) Then
          For J = 0 To UBound(ArrQMaintainSpacing)
            If ChrPos + Len(ArrQMaintainSpacing(J)) + 1 <= Len(strSearch) Then
              If Mid$(strSearch, ChrPos + 1, Len(ArrQMaintainSpacing(J))) = ArrQMaintainSpacing(J) Then
                SpaceBetweenReservedWords = True
                Exit For 'unction
              End If
            End If
          Next J
        End If
      End If
      If SpaceBetweenReservedWords Then
        Exit For
      End If
    Next I
  End If

End Function

Private Function SpaceOffsetClip(VarStr As Variant) As String

  Dim CutPoint As Long

  'is not always needed but lets some Fixers to operate properly
  ' by temporarily removeing any leading blanks
  If Left$(VarStr, 1) = SngSpace Then
    CutPoint = 1
    Do While Mid$(VarStr, CutPoint, 1) = SngSpace
      CutPoint = CutPoint + 1
    Loop
    CutPoint = CutPoint - 1
    SpaceOffsetClip = String$(CutPoint, SngSpace)
    VarStr = Mid$(VarStr, CutPoint + 1)
  End If

End Function

Public Function SpacePad(varA As Variant) As String

  SpacePad = SngSpace & varA & SngSpace

End Function

Public Function strCodeOnly(ByVal varCode As Variant) As String

  'only use this when comment deleion is not an issue
  'ie testing codeline without changing them

  strCodeOnly = varCode
  ExtractCode strCodeOnly

End Function

Public Function strInBrackets(varA As Variant) As String

  strInBrackets = LBracket & varA & RBracket

End Function

Public Function strInDQuotes(varA As Variant, _
                             Optional ByVal bPad As Boolean = False) As String

  strInDQuotes = DQuote & varA & DQuote
  If bPad Then
    strInDQuotes = SpacePad(strInDQuotes)
  End If

End Function

Public Function strInSQuotes(varA As Variant, _
                             Optional ByVal bPad As Boolean = False) As String

  strInSQuotes = SQuote & varA & SQuote
  If bPad Then
    strInSQuotes = SpacePad(strInSQuotes)
  End If

End Function

Public Function StripDoubleSpace(ByVal strCode As String) As String

  Dim FromPos As Long

  'Replace any DoubleSpace with single
  'Unless it is in a comment or String Literal
  FromPos = InStr(strCode, DblSpace)
  If FromPos Then
    Do
      If InCode(strCode, FromPos) Then
        strCode = Left$(strCode, FromPos - 1) & Mid$(strCode, FromPos + 1)
        FromPos = InStr(FromPos, strCode, DblSpace)
       Else
        FromPos = InStr(FromPos + 1, strCode, DblSpace)
      End If
    Loop While FromPos > 0
  End If
  StripDoubleSpace = strCode

End Function

Private Sub SuggestControlName(strGuard As String)

  Dim HungarianPref As String
  Dim strCaption    As String
  Dim StrNumSuffix  As String
  Dim I             As Long

  With frm_CodeFixer
    If ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl) > -1 Then
      HungarianPref = StandardCtrPrefix(ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl))
    End If
    If Len(CntrlDesc(LngCurrentControl).CDCaption) Then
      strCaption = CleanCaption(CntrlDesc(LngCurrentControl).CDCaption)
    End If
AlreadyExists:
    If Len(strCaption) Then ' a menu seperator's caption will be reduced to nothing
      If bCtrlDescExists Then
        For I = LBound(CntrlDesc) To UBound(CntrlDesc)
          If CntrlDesc(I).CDName = HungarianPref & strCaption & StrNumSuffix Then
            If InStr(strCaption, CntrlDesc(LngCurrentControl).CDForm) = 0 Then
              strCaption = strCaption & CntrlDesc(LngCurrentControl).CDForm
             Else
              StrNumSuffix = Val(StrNumSuffix) + 1
            End If
            GoTo AlreadyExists
          End If
        Next I
      End If
      If strCaption = HungarianPref & CntrlDesc(LngCurrentControl).CDName Then
        'stops lblLabel1, cmdCommand1 etc
        strCaption = vbNullString
      End If
      If Len(strCaption) Then
        .lblOldName(3).Caption = HungarianPref & strCaption & StrNumSuffix
        SmartAdd .lstPrefixSuggest, HungarianPref & strCaption & StrNumSuffix, strGuard
      End If
    End If
  End With
  Err.Number = 0

End Sub

Private Sub SuggestHungarianName(bSelectedArrayNAme As Boolean, _
                                 strGuard As String)

  Dim HungarianPref    As String
  Dim strSimpleForm    As String
  Dim strSimpleControl As String
  Dim strOldTxt        As String

  If ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl) > -1 Then
    HungarianPref = StandardCtrPrefix(ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl))
  End If
  If Len(HungarianPref) = 0 Then
    If InQSortArray(ArrQActiveControlClass, CntrlDesc(LngCurrentControl).CDClass) Then
      HungarianPref = "ctl"
    End If
  End If
  With frm_CodeFixer
    If .lstPrefixSuggest.ListCount Then
      'save previous selected suggestion
      strOldTxt = .lstPrefixSuggest.List(.lstPrefixSuggest.ListIndex)
    End If
    'clear inteface
    .txtCtrlNewName = vbNullString
    .lblOldName(1) = vbNullString
    .lblOldName(3) = vbNullString
    .lstPrefixSuggest.Clear
  End With
  With frm_CodeFixer
    SendMessage .lstPrefixSuggest.hWnd, WM_SETREDRAW, False, 0
    bSelectedArrayNAme = False
    If Len(HungarianPref) Then
      strSimpleForm = Ucase1st(NoHungarianPrefix(CntrlDesc(LngCurrentControl).CDForm))
      strSimpleControl = Ucase1st(NoHungarianPrefix(CntrlDesc(LngCurrentControl).CDName))
      'if you think of  a new format this is where you create it
      SmartAdd .lstPrefixSuggest, HungarianPref, strGuard
      If CntrlDesc(LngCurrentControl).CDBadType = BNDefault Then
        'VB default names just add the number to the prefixes
        SmartAdd .lstPrefixSuggest, HungarianPref & strSimpleForm & DefNameNumber(CntrlDesc(LngCurrentControl).CDName), strGuard
      End If
      SmartAdd .lstPrefixSuggest, HungarianPref & strSimpleForm, strGuard
      SmartAdd .lstPrefixSuggest, HungarianPref & Mid$(strSimpleForm, 4), strGuard
      SmartAdd .lstPrefixSuggest, HungarianPref & strSimpleControl, strGuard
      SmartAdd .lstPrefixSuggest, HungarianPref & strSimpleForm & strSimpleControl, strGuard
      SmartAdd .lstPrefixSuggest, HungarianPref & Mid$(strSimpleForm, 4) & strSimpleControl, strGuard
      '   End With 'frm_CodeFixer
    End If
    If Len(strOldTxt) Then
      If PosInList(strOldTxt, .lstPrefixSuggest) > -1 Then
        .lstPrefixSuggest.ListIndex = PosInList(strOldTxt, .lstPrefixSuggest)
        bSelectedArrayNAme = True
      End If
    End If
    SendMessage .lstPrefixSuggest.hWnd, WM_SETREDRAW, False, 0
  End With 'fCodeFix.lstPrefixSuggest

End Sub

Public Sub TabPostActivate()

  On Error Resume Next
  'Select Case tbsMain.SelectedItem.Key
  Select Case FrameActive
   Case TPFile  '"backup"
    FormCaptionDisplay "File Tool"
   Case TPModules  '"module"
    Generate_ModuleArray
    FormCaptionDisplay "Module Tool"
   Case TPControls  '"controls"
    FormCaptionDisplay "Control Tool"
    If Not Xcheck(XNoCntrlWarning) Then
      mObjDoc.Safe_MsgBox "PLEASE READ HELP FILE FOR HOW TO USE THIS TOOL" & vbNewLine & _
                    vbNewLine & _
                    "DO NOT USE ON ORIGINAL SOURCE CODE." & vbNewLine & _
                    "Control names that conflict with Variable, Procedure or Parameter names are not marked until Start button has been used." & vbNewLine & _
                    "This Tab is a very new tool. It may fail unexpectedly." & vbNewLine & _
                    "POSSIBLE PROBLEMS: " & vbNewLine & _
                    "1. Fails to update all references." & vbNewLine & _
                    "2. Variables/Properties that overlap spelling may be wrongly changed." & vbNewLine & _
                    "3. Code Fixer didn't recognise how Control is embedded in the code(let me know)." & vbNewLine & _
                    "4. Same name on 2+ forms may not update correctly if code on FormA calls control on FormB from a With structure." & vbNewLine & _
                    "5. Same name as VB Reserved word; Code Fixer may update reserved word (Menus called 'Exit' or 'End' can do this)." & vbNewLine & _
                    "FIX: Run the code using 'Ctrl+F5' every few changes to check it still works. Use Code Fixer comments to repair damage." & vbNewLine & _
                    "NOTE: You can add single controls to a control array but you cannot remove array members or blend 2 control arrays." & strTurnOffMsg, vbCritical

    End If
    bShowctrlPRoject = VBInstance.VBProjects.Count > 1
    bShowctrlComponent = ControlBearingCount > 1
    If Not UpdateCtrlnameLists Then
      HideInitiliser
    End If
  End Select
  On Error GoTo 0

End Sub

Public Sub TabPreActivate()

  With frm_CodeFixer
    .mnuDelete.Visible = False
    .mnuFindShow.Visible = False
    .mnuPopUpDelete.Visible = False
    .mnuPopControls.Visible = False
  End With
  Select Case FrameActive 'tbsMain.SelectedItem.Key
   Case TPModules '"module"
    frm_CodeFixer.tbsModule.SelectedItem = frm_CodeFixer.tbsModule.Tabs("module")
   Case TPFile '"reload"
    SetUpFileTool
    Do_tbsFileToolClick
  End Select

End Sub


':)Code Fixer V3.0.9 (25/03/2005 4:14:46 AM) 46 + 1914 = 1960 Lines Thanks Ulli for inspiration and lots of code.

