Attribute VB_Name = "mod_FindSupport"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
'This is a very slightly modified version of routines I have used in earlier versions
'Slight changes to workaround UserDocument limits
Public bDontDisplayCodePane      As Boolean
' Let FInd control in code fil list without showing codepane
Public lngFindCounter            As Long
Public bCancel                   As Boolean
Public Enum EnumMsg
  Search
  Complete
  inComplete
  Missing
  found
  replaced
  replacing
  deleteing
  Finished
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private Search, Complete, inComplete, Missing, found, replaced, replacing, deleteing, Finished
#End If
Public Enum CaseConvert
  myeUpperCase = 1
  myeLowerCase = 2
  myeProperCase = 3
  myeSimpleSentence = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private myeUpperCase, myeLowerCase, myeProperCase, myeSimpleSentence
#End If
Public Enum Range
  AllCode
  ModCode
  ProcCode
  SelCode
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private AllCode, ModCode, ProcCode, SelCode
#End If
'Search Settings
Private Const LongLimit          As Single = 2147483645
'Prevent FlexGrid from over-flowing;
'very unlikely that anything could be found that often
'just a safety valve inherited from the old listbox/IntegerLimit
Public mObjDoc                   As FindDoc
Public VBInstance                As VBE
Public bLaunchOnStart            As Boolean
Public bToolBarButton            As Boolean
Public bSaveHistory              As Boolean
Public bBlankWarning             As Boolean
Public bFilterWarning            As Boolean
Public bReplace2Search           As Boolean
Public bRemFilters               As Boolean
Public bLoadingSettings          As Boolean
Public bAutoSelectText           As Boolean
Public bLocTop                   As Boolean
Public bWholeWordonly            As Boolean
Public bCaseSensitive            As Boolean
Public bNoComments               As Boolean
Public bCommentsOnly             As Boolean
Public bNoStrings                As Boolean
Public bStringsOnly              As Boolean
Public BPatternSearch            As Boolean
Public bGridlines                As Boolean
Public bShowProject              As Boolean
Public bShowComponent            As Boolean
Public bShowRoutine              As Boolean
Public bShowCompLineNo           As Boolean
Public bShowProcLineNo           As Boolean
Public bFindSelectWholeLine      As Boolean
'
Public iRange                    As Long
Private ReplaceCount             As Long
Public HistDeep                  As Long
Public ColourTextFore            As Long
Public ColourTextBack            As Long
Public ColourFindSelectBack      As Long
Public ColourFindSelectFore      As Long
Public ColourHeadFore            As Long
Public ColourHeadDefault         As Long
Public ColourHeadWork            As Long
Public ColourHeadPattern         As Long
Public ColourHeadNoFind          As Long
Public ColourHeadReplace         As Long
Public GridSizer(4)              As String
Public mWindow                   As Window

Public Sub AddToSearchBox(ByVal StrCom As String)

  'this adds any comments to the Search Combo
  'and the general comment marker if necessary

  mObjDoc.ComboSetText SearchB, StrCom

End Sub

Public Function AppendStr(strHead As String, _
                          strAdd As String, _
                          Optional sep As String = SngSpace) As String

  If Len(strHead) Then
    AppendStr = strHead & sep & strAdd
   Else
    AppendStr = strAdd
  End If

End Function

Public Sub ApplySelectedTextRestriction(cde As String, _
                                        ByVal FStartR As Long, _
                                        ByVal FStartC As Long, _
                                        ByVal FEndR As Long, _
                                        ByVal FEndC As Long, _
                                        ByVal SStartR As Long, _
                                        ByVal SStartC As Long, _
                                        ByVal SEndR As Long, _
                                        ByVal SEndC As Long)

  Dim InSRange                  As Boolean

  'Code is needed in DoSearch and DoReplace so a separate Procedure
  If iRange = SelCode Then
    If Len(cde) Then
      InSRange = False
      If BetweenLng(SStartR, FStartR, SEndR) Then
        If BetweenLng(SStartR, FEndR, SEndR) Then
          InSRange = True
          'ver 2.2.1           'Refinement of test; only checks column if in first or last line of selected text
          ' any other line must be in the range
          If SStartR = FStartR Then
            InSRange = SStartC >= FStartC
           ElseIf SEndR = FEndR Then '
            InSRange = FEndC <= SEndC
          End If
          If FStartR = FEndR Then
            'special case single line selected
            InSRange = SStartC >= FStartC And FEndC <= SEndC
          End If
        End If
      End If
      If Not InSRange Then
        cde = vbNullString
      End If
    End If
  End If

End Sub

Public Sub ApplyStringCommentFilters(cde As String, _
                                     ByVal strTarget As String)

  Dim Codepos As Long

  Codepos = InStr(1, cde, strTarget, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
  If bNoComments Then
    If InComment(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bCommentsOnly Then
    If Not InComment(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bStringsOnly Then
    If Not InQuotes(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bNoStrings Then
    If InQuotes(cde, Codepos) Then
      cde = vbNullString
    End If
  End If

End Sub

Private Sub AutoPatternOff(strF As String)

  'v 2.1.5 expanded to modify patterns used

  If BPatternSearch Then
    If Not InstrArrayHard(strF, arrPatternDetector) Then
      BPatternSearch = False
      mObjDoc.ClearForPattern
    End If
  End If
  If BPatternSearch Then
    If InStr(strF, "\ASC(") Then
      strF = ConvertAsciiSearch(strF)
      AddToSearchBox strF
    End If
  End If

End Sub

Public Sub AutoSelectInitialize(ByVal PrevRange As Long, _
                                ByVal AutoRevert As Boolean)

  Dim strTest As String

  'Code is needed in DoSearch and DoReplace so a separate Procedure
  strTest = GetSelectedText
  If bAutoSelectText Then
    If InStr(strTest, vbNewLine) > 0 Then
      'v2.1.3 fix for selecting one line by clicking at left of line which incorperates a newline chr
      If InStr(RStrip(LStrip(strTest, vbNewLine), vbNewLine), vbNewLine) > 0 Then
        PrevRange = iRange
        iRange = SelCode
        AutoRevert = True
        mObjDoc.ToggleButtonFaces
      End If
    End If
  End If

End Sub

Public Function BetweenLng(ByVal MinV As Long, _
                           ByVal lngVal As Long, _
                           ByVal MaxV As Long, _
                           Optional ByVal InClusive As Boolean = True) As Boolean

  If InClusive Then
    If lngVal >= MinV Then
      If lngVal <= MaxV Then
        BetweenLng = True
      End If
    End If
   Else
    If lngVal > MinV Then
      If lngVal < MaxV Then
        BetweenLng = True
      End If
    End If
  End If

End Function

Private Function BlankWarningCancel(StrFnd As String, _
                                    ByVal StrRep As String) As Boolean

  If bBlankWarning Then
    If LenB(StrRep) = 0 Then
      BlankWarningCancel = vbCancel = MsgBox("Replace" & strInSQuotes(StrFnd, True) & "with blank?", vbExclamation + vbOKCancel, "Blank Warning " & AppDetails)
    End If
  End If

End Function

Public Function Bool2Lng(bValue As Boolean) As Long

  'used to simplify settings

  Bool2Lng = IIf(bValue, 1, 0)

End Function

Private Function CancelReplace(strF As String, _
                               strR As String) As Boolean

  If LenB(strF) = 0 Then
    CancelReplace = True
   ElseIf FilterWarningCancel Then
    CancelReplace = True
   ElseIf BlankWarningCancel(strF, strR) Then
    CancelReplace = True
  End If

End Function

Public Function ControlBearingCount() As Long

  Dim Proj As VBProject
  Dim Comp As VBComponent

  'the counts are used to control column visiblity
  On Error Resume Next
  '  ProjCount = VBInstance.VBProjects.Count
  'CompCount includes ProjCount just in case a group includes projects with only one component
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        ControlBearingCount = ControlBearingCount + IIf(IsComponent_ControlHolder(Comp), 1, 0)
      End If
    Next Comp
  Next Proj
  On Error GoTo 0

End Function

Private Function ConvertAsciiSearch(ByVal strStrConv As String) As String

  Dim StartAsc As Long
  Dim EndAsc   As Long
  Dim AscVal   As Long

  'this routine adds the ability to search for Characters referred to by ascii value
  'ONLY works if pattern search is ON
  'other wise you are searching for the literal string
  StartAsc = 1
  Do While InStr(StartAsc, strStrConv, "\ASC(")
    StartAsc = InStr(StartAsc, strStrConv, "\ASC(")
    EndAsc = InStr(StartAsc, strStrConv, RBracket)
    If EndAsc > StartAsc Then
      AscVal = Val(Mid$(strStrConv, StartAsc + 5, EndAsc - 1))
      If Between(0, AscVal, 255) Then
        strStrConv = Left$(strStrConv, StartAsc - 1) & Chr$(AscVal) & Mid$(strStrConv, EndAsc + 1)
       Else '
        StartAsc = EndAsc
        'this will jump you past an ivalid asc value and check for other "\ASC(" triggers
      End If
     Else
      Exit Do
      'this jumps out of Loop if the end bracket is missing
    End If
  Loop
  ConvertAsciiSearch = strStrConv

End Function

Public Sub ConvertSelectedText(Optional Conversion As CaseConvert)

  Dim StartLine As Long
  Dim startCol  As Long
  Dim EndLine   As Long
  Dim endCol    As Long
  Dim codeText  As String
  Dim cpa       As VBIDE.CodePane
  Dim cmo       As VBIDE.CodeModule
  Dim I         As Long

  'ConvertSelectedTex - Convert text selected in code window
  'Date: 7/31/1999
  'Versions: VB5 VB6 Level: Advanced
  'Author: The VB2TheMax Team
  ' Convert to uppercase, lowercase, or propercase the text that is
  ' currently selected in the active code window
  On Error Resume Next
  ' get a reference to the active code window and the underlying module
  ' exit if no one is available
  Set cpa = GetActiveCodePane
  Set cmo = cpa.CodeModule
  If Err.Number = 0 Then
    ' get the current selection coordinates
    cpa.GetSelection StartLine, startCol, EndLine, endCol
    ' exit if no text is highlighted
    If StartLine <> EndLine Then
      If startCol <> endCol Then
        ' get the code text
        If StartLine = EndLine Then
          ' only one line is partially or fully highlighted
          codeText = cmo.Lines(StartLine, 1)
          If Conversion = myeSimpleSentence Then
            SimpleSentenceCase codeText
           Else
            Mid$(codeText, startCol, endCol - startCol) = StrConv(Mid$(codeText, startCol, endCol - startCol), Conversion)
            If Conversion = vbProperCase Then
              ProperProperCase codeText
            End If
          End If
          cmo.ReplaceLine StartLine, codeText
         Else
          ' the selection spans multiple lines of code
          ' first, convert the highlighted text on the first line
          codeText = cmo.Lines(StartLine, 1)
          If Conversion = myeSimpleSentence Then
            SimpleSentenceCase codeText
           Else
            Mid$(codeText, startCol, Len(codeText) + 1 - startCol) = StrConv(Mid$(codeText, startCol, Len(codeText) + 1 - startCol), Conversion)
            If Conversion = vbProperCase Then
              ProperProperCase codeText
            End If
          End If
          cmo.ReplaceLine StartLine, codeText
          ' then convert the lines in the middle, that are fully highlighted
          For I = StartLine + 1 To EndLine - 1
            codeText = cmo.Lines(I, 1)
            If Conversion = myeSimpleSentence Then
              SimpleSentenceCase codeText
             Else
              codeText = StrConv(codeText, Conversion)
              If Conversion = vbProperCase Then
                ProperProperCase codeText
              End If
            End If
            cmo.ReplaceLine I, codeText
          Next I
          ' finally, convert the highlighted portion of the last line
          codeText = cmo.Lines(EndLine, 1)
          If Conversion = myeSimpleSentence Then
            SimpleSentenceCase codeText
           Else
            Mid$(codeText, 1, endCol - 1) = StrConv(Mid$(codeText, 1, endCol - 1), Conversion)
            If Conversion = vbProperCase Then
              ProperProperCase codeText
            End If
          End If
          cmo.ReplaceLine EndLine, codeText
        End If
        ' after replacing code we must restore the old selection
        ' this seems to be a side-effect of the ReplaceLine method
        cpa.SetSelection StartLine, startCol, EndLine, endCol
      End If
    End If
  End If
  On Error GoTo 0

End Sub

Public Sub DefaultGridSizes()

  GridSizer(0) = "Project"
  GridSizer(1) = "Component"
  GridSizer(2) = "Line"
  GridSizer(3) = "Procedure"
  GridSizer(4) = "Line"
  mObjDoc.GridReSize

End Sub

Public Sub DoFind(Flsv As ListView)

  Dim GotIt       As Boolean
  Dim StartLine   As Long
  Dim StartText   As Long
  Dim EndText     As Long
  Dim StrProjName As String
  Dim strCompName As String
  Dim strTarget   As String
  Dim strRoutine  As String
  Dim strFound    As String
  Dim LItem       As ListItem
  Dim Pane        As CodePane
  Dim CompMod     As CodeModule
  Dim Proj        As VBProject
  Dim Comp        As VBComponent
  Dim SCol        As Long

  On Error Resume Next
  Set LItem = Flsv.SelectedItem
  With LItem
    StrProjName = .Text
    strCompName = .SubItems(1)
    StartLine = CLng(.SubItems(2))
    strRoutine = .SubItems(3)
    'ProcLineNo = CLng(.SubItems(4))
    strFound = .SubItems(5)
  End With 'LItem
  strTarget = mObjDoc.ComboGetText(SearchB)
  'this is the fast Find used if no editing has been done
  GotIt = GetFoundCodeLine(StrProjName, strCompName, strFound, Comp, StartLine, SCol, True)
  '  If Not GotIt Then
  '  GotIt = GetFoundCodeLine(StrProjName, strCompName, strFound, Comp, StartLine, SCol, False)
  '  End If
  'this will do Find if the line number has been changed by editing
  If Not GotIt Then
    StartLine = 1
    For Each Proj In VBInstance.VBProjects
      If StrProjName = Proj.Name Then
        For Each Comp In Proj.VBComponents
          If LenB(Comp.Name) Then
            If Comp.Name = strCompName Then
              Set CompMod = Comp.CodeModule
              StartLine = 1
              Do While CompMod.Find(strFound, StartLine, SCol, -1, -1, True, True)
                '        L_CodeLine = .CodeModule.Lines(StartLine, 1)
                '              Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFound, "", StartLine, , SCol)
                'this do loop takes care of the possibility of identical lines being present in different routines
                If strRoutine = GetProcName(CompMod, StartLine) Then
                  GotIt = True
                  'reset the Line data so fast Find will be used next time
                  LItem.SubItems(2) = StartLine
                  LItem.SubItems(4) = GetProcLineNumber(CompMod, StartLine)
                  Exit Do
                End If
                StartLine = StartLine + 1
              Loop
            End If
            If GotIt Then
              Exit For
            End If
          End If
        Next Comp
        If GotIt Then
          Exit For
        End If
      End If
    Next Proj
  End If
  If GotIt Then
    Set CompMod = Comp.CodeModule
    If bFindSelectWholeLine Then
      'select the whole line
      StartText = InStr(1, CompMod.Lines(StartLine, 1), strFound, vbTextCompare)
      EndText = StartText + Len(strFound)
     Else
      'select the search word
      If Not BPatternSearch Then
        'StartText = SCol '
        StartText = InStr(1, CompMod.Lines(StartLine, 1), strTarget, vbTextCompare)
        'vbBinaryCompare) '
        EndText = StartText + Len(strTarget)
       Else
        StartText = InStr(1, CompMod.Lines(StartLine, 1), strFound, vbTextCompare)
        EndText = StartText + Len(strFound)
      End If
    End If
    If GotIt Then
      ReportAction Flsv, found
      Set Pane = CompMod.CodePane
      With Pane
        ShowPane Pane
        'v2.0.5 fixed the VB toolbars not coming back into focus when searching
        'v2.0.5 the next line reactivates the vb menu buttons
        VBInstance.MainWindow.SetFocus
        mObjDoc.DoResize
        .Window.SetFocus
        .TopLine = Abs(Int(.CountOfVisibleLines / 2) - StartLine) + 1
        .SetSelection StartLine, StartText, StartLine, EndText
        DoEvents
      End With
      Set Pane = Nothing
      ' Exit Sub
    End If
   Else
    'Line is missing (probably edited out) so re do whole search
    '(may be a pain if really large search was done),
    '    SendMessage VBInstance.MainWindow.hWnd, WM_SETREDRAW, True, 0
    DoSearch Flsv
  End If
  On Error GoTo 0

End Sub

Public Sub DoReplace(Flsv As ListView)

  
  Dim AutoRangeRevert As Boolean
  Dim StartLine       As Long
  Dim EndLine         As Long
  Dim startCol        As Long
  Dim endCol          As Long
  Dim SelstartLine    As Long
  Dim selendline      As Long
  Dim SelStartCol     As Long
  Dim SelEndCol       As Long
  Dim PrevCurCodePane As Long
  Dim code            As String
  Dim strFind         As String
  Dim StrReplace      As String
  Dim curProc         As String
  Dim curModule       As String
  Dim ProcName        As String
  Dim CompMod         As CodeModule
  Dim Comp            As VBComponent
  Dim Proj            As VBProject

  strFind = mObjDoc.ComboGetText(SearchB)
  StrReplace = mObjDoc.ComboGetText(ReplaceB)
  If Not CancelReplace(strFind, StrReplace) Then
    On Error Resume Next
    With mObjDoc
      .ShowWorking True, "Replacing..."
      .ComboBoxSave SearchB, , HistDeep
      .ComboBoxSave ReplaceB, , HistDeep
    End With 'mobjDoc
    ReplaceSpecialChar StrReplace
    DoEvents
    bCancel = False
    'GetCounts
    curProc = GetCurrentProcedureName
    curModule = GetActiveModuleName
    AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
    If iRange = SelCode Then
      GetActiveCodePane.GetSelection SelstartLine, SelStartCol, selendline, SelEndCol
    End If
    ReplaceCount = 0
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If LenB(Comp.Name) Then
          ReportAction Flsv, replacing
          If iRange > AllCode Then
            If Comp.Name <> curModule Then
              GoTo SkipComp
            End If
          End If
          Set CompMod = Comp.CodeModule
          StartLine = 1
          If CompMod.Find(strFind, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
            Do
              startCol = 1
              EndLine = -1
              endCol = -1
              If CompMod.Find(strFind, StartLine, startCol, EndLine, endCol, bWholeWordonly, bCaseSensitive, False) Then
                ProcName = CompMod.ProcOfLine(StartLine, vbext_pk_Proc)
                If LenB(ProcName) = 0 Then
                  ProcName = "(Declarations)"
                End If
                If iRange = ProcCode Then
                  If ProcName <> curProc Then
                    GoTo SkipProc
                  End If
                End If
                code = CompMod.Lines(StartLine, 1)
                ApplyStringCommentFilters code, strFind
                ApplySelectedTextRestriction code, StartLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, selendline, SelEndCol
                If Len(code) Then
                  If bWholeWordonly Then
                    WholeWordReplacer code, strFind, StrReplace
                   Else
                    code = Replace$(code, strFind, StrReplace, , , IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
                  End If
                  CompMod.ReplaceLine StartLine, code
                  ReplaceCount = ReplaceCount + 1
                  mObjDoc.ShowWorking True, "Replacing...", strInBrackets(ReplaceCount) & " Items"
                End If
SkipProc:
              End If
              code = vbNullString
              StartLine = StartLine + 1
              If mObjDoc.CancelSearch Then
                Exit Do
              End If
              If StartLine > CompMod.CountOfLines Then
                Exit Do
              End If
            Loop While CompMod.Find(strFind, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False)
            'StartLine > 0 And StartLine <= CompMod.CountOfLines
          End If
        End If
SkipComp:
        Set Comp = Nothing
        If mObjDoc.CancelSearch Then
          Exit For
        End If
      Next Comp
      If mObjDoc.CancelSearch Then
        Exit For
      End If
    Next Proj
    'this turns off auto Selected text only
    If AutoRangeRevert Then
      iRange = PrevCurCodePane
      mObjDoc.ToggleButtonFaces
    End If
    Set Proj = Nothing
    Set CompMod = Nothing
    ReportAction Flsv, replaced
    mObjDoc.ShowWorking False
    If bReplace2Search Then
      If Len(StrReplace) Then
        mObjDoc.ComboBoxSave ReplaceB, strFind, HistDeep
        mObjDoc.ComboBoxSave SearchB, StrReplace, HistDeep
        DoFind Flsv
      End If
    End If
    On Error GoTo 0
  End If

End Sub

Public Sub DoSearch(Flsv As ListView)

  
  Dim AutoRangeRevert As Boolean
  Dim SecondRun       As Boolean
  Dim bComplete       As Boolean
  Dim PrevCurCodePane As Long
  Dim SelstartLine    As Long
  Dim SeEndLine       As Long
  Dim SelStartCol     As Long
  Dim SelEndCol       As Long
  Dim ProcLineNo      As Long
  Dim OldPosition     As Long
  Dim StartLine       As Long
  Dim startCol        As Long
  Dim EndLine         As Long
  Dim endCol          As Long
  Dim strStrComTest   As String
  Dim curModule       As String
  Dim code            As String
  Dim ProcName        As String
  Dim strFind         As String
  Dim curProc         As String
  Dim LItem           As ListItem
  Dim CompMod         As CodeModule
  Dim Comp            As VBComponent
  Dim Proj            As VBProject
  Dim RepType         As EnumMsg
  Dim MyhourGlass     As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  'ver1.1.02 major rewrite to allow simple pattern searching
  On Error Resume Next
  lngFindCounter = 0
  strFind = mObjDoc.ComboGetText(SearchB)
  If Not EmptySpaceGuard(strFind) Then
    'if strFind doesn't have any triggers and Pattern Search is on then turn it off
    AutoPatternOff strFind
    With Flsv
      If Not .SelectedItem Is Nothing Then
        'v2.3.5 added safety
        OldPosition = .SelectedItem.Index
        .ListItems.Clear
      End If
    End With 'Flsv
    DefaultGridSizes
    mObjDoc.ShowWorking True, "Searching..."
ReTry:
    If LenB(strFind) > 0 Then ' just in case shoul always be true at this point
      bComplete = False
      mObjDoc.ComboBoxSave SearchB, strFind, HistDeep
      mObjDoc.CancelButton True
      DoEvents
      bCancel = False
      curModule = GetActiveModuleName
      curProc = GetCurrentProcedureName
      'this does the auto switching if multiple code lines are selected
      AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
      ' this sets search limits if there is selected code.
      If iRange = SelCode Then
        GetSelectedText
        GetActiveCodePane.GetSelection SelstartLine, SelStartCol, SeEndLine, SelEndCol
      End If
      For Each Proj In VBInstance.VBProjects
        If Len(Proj.Name) > Len(GridSizer(0)) Then
          GridSizer(0) = Proj.Name
        End If
        For Each Comp In Proj.VBComponents
          If Flsv.ListItems.Count < LongLimit Then
            If LenB(Comp.Name) Then
              If iRange > AllCode Then
                If Comp.Name <> curModule Then
                  GoTo SkipComp
                End If
              End If
              If Len(Comp.Name) > Len(GridSizer(1)) Then
                GridSizer(1) = Comp.Name
              End If
              With Comp
                Set CompMod = .CodeModule
                If LenB(.Name) = 0 Then
                  bCancel = True
                  bComplete = True
                End If
              End With
              If mObjDoc.CancelSearch Then
                Exit For
              End If
              'Safety turns off filters if comment/double quote is actually in the search phrase
              If bNoComments Then
                If InStr(strFind, SQuote) > 0 Then
                  bNoComments = True
                  mObjDoc.SetFilterButtons
                End If
              End If
              If bNoStrings Then
                If InStr(strFind, DQuote) > 0 Then
                  bNoStrings = False
                  mObjDoc.SetFilterButtons
                End If
              End If
              StartLine = 1 'initialize search range
              startCol = 1
              EndLine = -1
              endCol = -1
              Do While CompMod.Find(strFind, StartLine, startCol, EndLine, endCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch)
                DoEvents
                code = CompMod.Lines(StartLine, 1)
                If BPatternSearch Then
                  ' the string/comment filters cannot work on a PatternSearch StrFind
                  'but you can get the actual string found and test that
                  CompMod.CodePane.SetSelection StartLine, startCol, EndLine, endCol
                  strStrComTest = Mid$(code, startCol, endCol - startCol)
                 Else
                  strStrComTest = strFind
                End If
                'apply nostring no comment filters
                ApplyStringCommentFilters code, strStrComTest
                'v3.0.3 ignore underscores
                If bWholeWordonly Then
                  If LenB(code) Then
                    If InStrWholeWordRX(code, strFind, , IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)) = 0 Then
                      code = vbNullString
                    End If
                  End If
                End If
                ApplySelectedTextRestriction code, StartLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, SeEndLine, SelEndCol
                If mObjDoc.CancelSearch Then
                  Exit For
                End If
                'ver 2.0.3 changed so that resize code only fires if a find is made
                If LenB(code) Then
                  ProcName = GetProcName(CompMod, StartLine)
                  ProcLineNo = GetProcLineNumber(CompMod, StartLine)
                  If iRange = ProcCode Then
                    If ProcName <> curProc Then
                      GoTo SkipProc
                    End If
                  End If
                  ProcLineNo = GetProcLineNumber(CompMod, StartLine)
                  If Len(CStr(StartLine)) > Len(GridSizer(2)) Then
                    GridSizer(2) = CStr(StartLine)
                  End If
                  If Len(ProcName) > Len(GridSizer(3)) Then
                    GridSizer(3) = ProcName
                  End If
                  If Len(CStr(ProcLineNo)) > Len(GridSizer(4)) Then
                    GridSizer(4) = CStr(ProcLineNo)
                  End If
                  'old position of find and test limits code
                  Set LItem = Flsv.ListItems.Add(, , Proj.Name)
                  With LItem
                    .ListSubItems.Add , , Comp.Name
                    .ListSubItems.Add , , StartLine
                    .ListSubItems.Add , , ProcName
                    .ListSubItems.Add , , ProcLineNo
                    .ListSubItems.Add , , Trim$(code)
                    lngFindCounter = lngFindCounter + 1
                  End With 'LItem
                  code = vbNullString
                  ReportAction Flsv, Search
                End If
SkipProc:
                StartLine = StartLine + 1
                If mObjDoc.CancelSearch Then
                  Exit Do
                End If
                If StartLine > CompMod.CountOfLines Then
                  Exit Do
                End If
                startCol = 1
                EndLine = -1
                endCol = -1
              Loop
            End If
          End If
SkipComp:
          Set Comp = Nothing
          If mObjDoc.CancelSearch Then
            Exit For
          End If
        Next Comp
        If mObjDoc.CancelSearch Then
          Exit For
        End If
      Next Proj
      Set Proj = Nothing
      Set CompMod = Nothing
      mObjDoc.GridReSize
      mObjDoc.CancelButton False
    End If
    If Flsv.ListItems.Count = 0 Then
      If Not BPatternSearch Then
        If InstrArrayHard(strFind, arrPatternDetector) Then
          BPatternSearch = Not BPatternSearch
          mObjDoc.ClearForPattern
          SecondRun = True
          GoTo ReTry
        End If
      End If
      If SecondRun Then
        'autoPattern Search is on
        If Flsv.ListItems.Count = 0 Then
          'turn it off it still no finds
          BPatternSearch = Not BPatternSearch
          mObjDoc.ClearForPattern
        End If
      End If
    End If
    If Flsv.ListItems.Count Then
      RepType = Complete
     Else
      RepType = Missing
    End If
    If mObjDoc.CancelSearch Then
      RepType = inComplete
    End If
    ReportAction Flsv, IIf(bComplete, found, RepType)
    'this turns off auto Selected text only
    If AutoRangeRevert Then
      iRange = PrevCurCodePane
      mObjDoc.ToggleButtonFaces
    End If
    mObjDoc.ShowWorking False
    ReportAction Flsv, Finished
    'this turns auto switch to pattern search off if it was used
    If Flsv.ListItems.Count = LongLimit Then
      ' as this is 2147483647 rows it is unlikely that this will ever hit but just in case :)
      MsgBox "Search halted because number of finds reached limit of Find ListView", vbCritical
    End If
    Flsv.Refresh
    SetFocus_Safe Flsv
    With Flsv
      If .ListItems.Count Then
        If OldPosition > .ListItems.Count Then
          ''ver 2.0.2 search window returns to close to previous location
          ''if you repeat a search
          OldPosition = .ListItems.Count
        End If
        'OldPosition = 1
        'v2.3.5 changed to '> 0' from '> -1'
        Set .SelectedItem = .ListItems(IIf(OldPosition > 0, OldPosition, 1))
        .SelectedItem.EnsureVisible
        If Not bDontDisplayCodePane Then
          DoFind Flsv
        End If
      End If
    End With 'Flsv
  End If
  On Error GoTo 0

End Sub

Private Function EmptySpaceGuard(ByVal strFind As String) As Boolean

  If LenB(strFind) = 0 Then
    EmptySpaceGuard = True
    mObjDoc.ComboSetFocus SearchB
   ElseIf strFind = SngSpace Then
    MsgBox "Search for single spaces is cancelled, it overloads the system", vbInformation
    EmptySpaceGuard = True
    mObjDoc.ComboSetFocus SearchB
  End If

End Function

Private Function FilterWarningCancel() As Boolean

  Dim Msg As String

  If bFilterWarning Then
    Msg = ONOFF(bWholeWordonly) & " Whole Word"
    Msg = Msg & vbNewLine & ONOFF(bCaseSensitive) & " Case Sensitive"
    Msg = Msg & vbNewLine & ONOFF(bNoComments) & " No Comments"
    Msg = Msg & vbNewLine & ONOFF(bCommentsOnly) & " Comments Only"
    Msg = Msg & vbNewLine & ONOFF(bNoStrings) & " No Strings"
    Msg = Msg & vbNewLine & ONOFF(bStringsOnly) & " Strings Only"
    FilterWarningCancel = vbCancel = mObjDoc.Safe_MsgBox(Msg & vbNewLine & _
                                                     vbNewLine & _
                                                     "Proceed with Replace anyway?", vbExclamation + vbOKCancel, "Filter Warning " & AppDetails)
  End If

End Function

Public Sub FindInCode(ByVal strFind As String, _
                      strCmp As String, _
                      strPrj As String)

  Dim GotIt     As Boolean
  Dim StartLine As Long
  Dim StartText As Long
  Dim EndText   As Long
  Dim CompSize  As Long
  Dim CompMod   As CodeModule
  Dim Pane      As CodePane

  StartLine = 1
  Set CompMod = GetComponent(strPrj, strCmp).CodeModule
  CompSize = CompMod.CountOfLines
  Do While CompMod.Find(strFind, StartLine, 1, -1, -1, True, True)
    '        L_CodeLine = .CodeModule.Lines(StartLine, 1)
    'Do While GetWholeCaseMatchCodeLine(strPrj, strCmp, strFind, "", StartLine)
    'this do loop takes care of the possibility of identical lines being present in different routines
    If strFind = GetProcName(CompMod, StartLine) Then
      StartText = InStr(1, CompMod.Lines(StartLine, 1), strFind, vbTextCompare)
      EndText = StartText + Len(strFind)
      GotIt = True
      Exit Do
    End If
    StartLine = StartLine + 1
    If StartLine > CompSize Then
      Exit Do 'Sub
    End If
  Loop
  If GotIt Then
    Set Pane = CompMod.CodePane
    With Pane
      '   ShowPane Pane
      'v2.0.5 fixed the VB toolbars not coming back into focus when searching
      'v2.0.5 the next line reactivates the vb menu buttons
      '  VBInstance.MainWindow.SetFocus
      '   mObjDoc.DoResize
      '   .Window.SetFocus
      .TopLine = Abs(Int(.CountOfVisibleLines / 2) - StartLine) + 1
      .SetSelection StartLine, StartText, StartLine, EndText
      DoEvents
    End With
    Set Pane = Nothing
  End If
  Set CompMod = Nothing

End Sub

Public Sub ForceDisplayFind()

  If Not mWindow.Visible Then
    mWindow.Visible = True
    ' mObjDoc.Show
  End If

End Sub

Public Function GetActiveCodeModule() As CodeModule

  On Error Resume Next
  Set GetActiveCodeModule = GetActiveCodePane.CodeModule
  On Error GoTo 0

End Function

Public Function GetActiveCodePane() As CodePane

  On Error Resume Next
  Set GetActiveCodePane = VBInstance.ActiveCodePane
  On Error GoTo 0

End Function

Public Function GetActiveModuleName() As String

  On Error Resume Next
  GetActiveModuleName = GetActiveCodePane.CodeModule.Name
  On Error GoTo 0

End Function

Public Function GetActiveProject() As VBProject

  On Error Resume Next
  Set GetActiveProject = VBInstance.ActiveVBProject
  On Error GoTo 0

End Function

Private Function GetCurrentProcedureName() As String

  Dim StartLine As Long
  Dim lJunk     As Long

  With GetActiveCodePane
    .GetSelection StartLine, lJunk, lJunk, lJunk
    lJunk = 0
    GetLineData .CodeModule, StartLine, GetCurrentProcedureName, lJunk, lJunk, lJunk, lJunk, lJunk
  End With

End Function

Public Function GetFoundCodeLine(VarProjName As Variant, _
                                 VarModName As Variant, _
                                 varFind As Variant, _
                                 vComp As VBComponent, _
                                 Optional lngStartLine As Long, _
                                 Optional lngColStart As Long, _
                                 Optional bWholeWord As Boolean = False) As Boolean

  'Check that a Found line is still in the code
  
  Dim lngSCol As Long

  Set vComp = VBInstance.VBProjects(VarProjName).VBComponents(VarModName)
  lngSCol = 1
  If Not vComp Is Nothing Then
    GetFoundCodeLine = vComp.CodeModule.Find(varFind, lngStartLine, lngSCol, -1, -1, bWholeWord, bWholeWord, False)
    lngColStart = lngSCol
  End If

End Function

Public Sub GetLineData(cmpMod As CodeModule, _
                       ByVal cdeline As Long, _
                       ProcName As String, _
                       ProcLineNo As Long, _
                       PRocStartLine As Long, _
                       ProcEndLine As Long, _
                       ProcHeadLine As Long, _
                       Proc1stInsertLine As Long)

  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

  For I = 1 To 4
    K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
    CleanElem = Null
    On Error Resume Next
    'IF you crash here first check that Error Trapping is not ON
    CleanElem = cmpMod.ProcOfLine(cdeline, K)
    On Error GoTo 0
    If Not IsNull(CleanElem) Then
      ProcName = CleanElem
      If Len(ProcName) = 0 Then
        ProcName = "(Declarations)"
        ProcLineNo = cdeline
        PRocStartLine = 1
        ProcEndLine = cmpMod.CountOfDeclarationLines
       Else
        ProcLineNo = cmpMod.ProcBodyLine(ProcName, K)
        PRocStartLine = cmpMod.PRocStartLine(ProcName, K)
        ProcHeadLine = PRocStartLine
        Proc1stInsertLine = ProcHeadLine
        Proc1stInsertLine = Proc1stInsertLine + 1
        ProcEndLine = PRocStartLine + cmpMod.ProcCountLines(ProcName, K)
      End If
      Exit For
    End If
  Next I

End Sub

Public Function getProcedureName() As String

  Dim StartLine As Long
  Dim cpa       As VBIDE.CodePane
  Dim lngdummy  As Long

  ' This function can only be used inside an add-in.
  On Error Resume Next
  ' get a reference to the active code window and the underlying module
  ' exit if no one is available
  Set cpa = GetActiveCodePane
  ' get the current selection coordinates
  cpa.GetSelection StartLine, lngdummy, lngdummy, lngdummy
  GetProcStartLine cpa.CodeModule, StartLine, getProcedureName ', lngPRocKind
  On Error GoTo 0

End Function

Public Function GetProcLineNumber(cmpMod As CodeModule, _
                                  CodeLineNo As Long) As String

  Dim LProcName As String
  Dim I         As Long
  Dim CleanElem As Variant

  LProcName = GetProcName(cmpMod, CodeLineNo)
  If LProcName = "(Declarations)" Then
    GetProcLineNumber = CodeLineNo
   Else
    'The + 1 is because ProcBodyLine returns a 0 based count but most people like 1 based counts
    'Oddly CodeLineNo which is generated by VB's Find is 1 based
    For I = 1 To 4
      CleanElem = Null
      On Error Resume Next
      'IF you crash here first check that Error Trapping is not ON
      CleanElem = CodeLineNo - cmpMod.ProcBodyLine(LProcName, Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)) + 1
      On Error GoTo 0
      If Not IsNull(CleanElem) Then
        GetProcLineNumber = CleanElem
        Exit For
      End If
    Next I
  End If

End Function

Public Function GetSelectedText() As String

  Dim StartLine As Long
  Dim cmo       As VBIDE.CodeModule
  Dim codeText  As String
  Dim cpa       As VBIDE.CodePane
  Dim endCol    As Long
  Dim EndLine   As Long
  Dim startCol  As Long

  'Date: 4/27/1999
  'Versions: VB5 VB6 Level: Intermediate
  'Author: The VB2TheMax Team
  ' Return the string of code the is selected in the code window
  ' that is currently active.
  ' This function can only be used inside an add-in.
  On Error Resume Next
  ' get a reference to the active code window and the underlying module
  ' exit if no one is available
  Set cpa = GetActiveCodePane
  Set cmo = cpa.CodeModule
  If Err.Number = 0 Then
    ' get the current selection coordinates
    cpa.GetSelection StartLine, startCol, EndLine, endCol
    ' exit if no text is highlighted
    If StartLine + EndLine + startCol + endCol > 0 Then
      ' get the code text
      If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        codeText = Mid$(cmo.Lines(StartLine, 1), startCol, endCol - startCol)
       Else
        ' the selection spans multiple lines of code
        ' first, get the selection of the first line
        codeText = Mid$(cmo.Lines(StartLine, 1), startCol) & vbNewLine
        ' then get the lines in the middle, that are fully highlighted
        If StartLine + 1 < EndLine Then
          codeText = codeText & cmo.Lines(StartLine + 1, EndLine - StartLine - 1)
        End If
        ' finally, get the highlighted portion of the last line
        codeText = codeText & Left$(cmo.Lines(EndLine, 1), endCol - 1)
      End If
      GetSelectedText = codeText
    End If
  End If
  On Error GoTo 0

End Function

Public Function InDeclaration() As Boolean

  Dim StartLine As Long
  Dim cpa       As VBIDE.CodePane
  Dim lngdummy  As Long

  ' This function can only be used inside an add-in.
  On Error Resume Next
  ' get a reference to the active code window and the underlying module
  ' exit if no one is available
  Set cpa = GetActiveCodePane
  ' get the current selection coordinates
  cpa.GetSelection StartLine, lngdummy, lngdummy, lngdummy
  InDeclaration = StartLine <= cpa.CodeModule.CountOfDeclarationLines
  On Error GoTo 0

End Function

Private Function InQuotes(ByVal code As String, _
                          ByVal Codepos As Long) As Boolean

  Dim CL As Long
  Dim CR As Long

  'ver 2.0.1 improved test
  If InStr(code, DQuote) Then        ' exists
    If InLiteral(code, Codepos) Then '
      CL = CountSubString(Left$(code, Codepos), DQuote)
      If CL > 0 Then
        CR = CountSubString(Mid$(code, Codepos), DQuote)
        If CR > 0 Then
          If IsOdd(CL) Then
            InQuotes = True
          End If
        End If
      End If
    End If
  End If

End Function

Public Sub InsertInCode(ByVal strFind As String, _
                        ByVal strProc As String, _
                        strCmp As String, _
                        strPrj As String, _
                        ByVal StrInsert As String)

  Dim StartLine As Long
  Dim CompSize  As Long
  Dim CompMod   As CodeModule

  StartLine = 1
  Set CompMod = GetComponent(strPrj, strCmp).CodeModule
  CompSize = CompMod.CountOfLines
  Do While CompMod.Find(strFind, StartLine, 1, -1, -1, True, True)
    '  Do While GetWholeCaseMatchCodeLine(strPrj, strCmp, strFind, "", StartLine)
    'this do loop takes care of the possibility of identical lines being present in different routines
    If strProc = GetProcName(CompMod, StartLine) Then
      CompMod.InsertLines StartLine + 1, StrInsert
      Exit Do
    End If
    StartLine = StartLine + 1
    If StartLine > CompSize Then
      Exit Do 'Sub
    End If
  Loop

End Sub

Public Function InSRange(Lnum As Long, _
                         cmpMod As CodeModule, _
                         curProc As String, _
                         selendline As Long) As Boolean

  Select Case iRange
   Case AllCode, ModCode
    InSRange = (Lnum <= cmpMod.CountOfLines)
   Case ProcCode
    InSRange = (curProc = GetProcName(cmpMod, Lnum))
   Case SelCode
    InSRange = (Lnum <= selendline)
  End Select

End Function

Public Function LastWord(ByVal varChop As Variant) As String

  Dim TmpA As Variant

  If LenB(varChop) Then
    TmpA = Split(varChop)
    LastWord = TmpA(UBound(TmpA))
  End If

End Function

Public Sub LoadFormPosition(frm As Form)

  'Requires AppDetails to supply top of Registry branch
  'You could also hard code it if you want

  With frm
    .Left = GetSetting(AppDetails, "Settings", .Name & "Left", .Left)
    .Top = GetSetting(AppDetails, "Settings", .Name & "Top", .Top)
    If .BorderStyle = vbSizableToolWindow Or .BorderStyle = vbSizable Then
      'don't bother to load if form is not resizable
      .Width = GetSetting(AppDetails, "Settings", .Name & "Width", .Width)
      .Top = GetSetting(AppDetails, "Settings", .Name & "Height", .Height)
    End If
  End With 'Me

End Sub

Private Function ONOFF(bVal As Boolean) As String

  ONOFF = IIf(bVal, "ON ", "OFF")

End Function

Public Function PosInCombo(ByVal strA As String, _
                           ByVal c As ComboBox, _
                           Optional CaseSensitive As Boolean = True) As Long

  Const CB_FINDSTRINGEXACT As Long = &H158
  Const CB_FINDSTRING      As Long = &H14C

  'find if strA is in Combolist
  'returns -1 if not found
  PosInCombo = SendMessage(c.hWnd, IIf(CaseSensitive, CB_FINDSTRINGEXACT, CB_FINDSTRING), 0, ByVal strA)

End Function

Private Sub ProperProperCase(strCode As String, _
                             Optional strTrigger As String = vbNullString)

  Dim hasquotes As Long

  If LenB(strTrigger) = 0 Then
    strTrigger = DQuote
  End If
  hasquotes = InStr(strCode, strTrigger)
  Do While hasquotes
    strCode = Left$(strCode, hasquotes) & UCase$(Mid$(strCode, hasquotes + 1, 1)) & Mid$(strCode, hasquotes + 2)
    hasquotes = InStr(hasquotes + 1, strCode, strTrigger)
  Loop

End Sub

Private Sub ReplaceSpecialChar(StrReplace As String)

  If InStr(StrReplace, "^N^") Then
    StrReplace = Replace$(StrReplace, "^N^", vbNewLine)
  End If
  If InStr(StrReplace, "^T^") Then
    StrReplace = Replace$(StrReplace, "^T^", vbTab)
  End If

End Sub

Public Sub ReportAction(Flsv As ListView, _
                        ByVal Act As EnumMsg, _
                        Optional ByVal AppendStr As String)

  Dim Msg                As String
  Dim StrItems           As String
  Dim StrFilterWarning   As String
  Dim strSearchEndStatus As String
  Dim StrPatternWarning  As String
  Dim strStatusMsg       As String
  Dim strStatusActMsg    As String

  StrItems = strInBrackets(Flsv.ListItems.Count) & " Item" & IIf(Flsv.ListItems.Count, "s", vbNullString)
  StrFilterWarning = IIf(mObjDoc.AnyFilterOn, " <Filter>", vbNullString)
  StrPatternWarning = IIf(BPatternSearch, " <Pattern>", vbNullString)
  Select Case Act
   Case Search
    strSearchEndStatus = " Searching " & IIf(Len(AppendStr), " in " & AppendStr, "...")
    strStatusMsg = "Searching..."
    strStatusActMsg = StrItems & " Found" & StrFilterWarning & StrPatternWarning
   Case Complete
    strSearchEndStatus = " Search Complete."
   Case inComplete
    strSearchEndStatus = " Search Cancelled."
  End Select
  Select Case Act
   Case replacing
    strStatusMsg = "Replacing..."
    strStatusActMsg = vbNullString
    Msg = "Replacing.." & String$(Int(Rnd * 5 + 1), ".")
   Case replaced
    Msg = strInBrackets(ReplaceCount) & " Item" & IIf(ReplaceCount <> 1, "s", vbNullString) & " replaced"
   Case Missing
    Msg = StrItems & " Found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
   Case deleteing, Finished
    If Flsv.ListItems.Count Then
      Msg = StrItems & " Found" & StrFilterWarning & StrPatternWarning
     Else
      Msg = "No Matches"
    End If
   Case Else
    Msg = StrItems & " Found" & StrItems & StrFilterWarning & strSearchEndStatus & StrPatternWarning
  End Select
  Select Case Act
   Case Search, replacing
    mObjDoc.ShowWorking True, strStatusMsg, strStatusActMsg
    Flsv.ColumnHeaders(6).Text = "Code: " & strStatusActMsg 'Msg
   Case Complete, inComplete, replaced, Missing, deleteing, Finished
    Flsv.ColumnHeaders(6).Text = "Code: " & Msg
  End Select
  DoEvents

End Sub

Public Sub SaveFormPosition(frm As Form)

  'Requires AppDetails to supply top of Registry branch
  'You could also hard code it if you want

  With frm
    SaveSetting AppDetails, "Settings", .Name & "Left", .Left
    SaveSetting AppDetails, "Settings", .Name & "Top", .Top
    If .BorderStyle = vbSizableToolWindow Or .BorderStyle = vbSizable Then
      'don't bother to save if form is not resizable
      SaveSetting AppDetails, "Settings", .Name & "Width", .Width
      SaveSetting AppDetails, "Settings", .Name & "Height", .Height
    End If
  End With 'frm

End Sub

Public Sub SelectedText(cmb As ComboBox, _
                        Cmd As CommandButton)

  Dim HiLitSelection As String

  HiLitSelection = GetSelectedText
  If LenB(HiLitSelection) Then
    If InStr(HiLitSelection, vbNewLine) Then
      If HiLitSelection <> vbNewLine Then
        HiLitSelection = Left$(HiLitSelection, InStr(HiLitSelection, vbNewLine) - 1)
      End If
    End If
    If LenB(HiLitSelection) Then
      SetFocus_Safe cmb
      cmb.Text = HiLitSelection
      Cmd = True
    End If
  End If

End Sub

Public Sub SetFocus_Safe(ctl As Control)

  '*PURPOSE: protect SetFocus from any of the many conditions which can stuff it

  On Error Resume Next
  With ctl
    If .Visible Then
      If .Enabled Then
        .SetFocus
      End If
    End If
  End With
  On Error GoTo 0

End Sub

Private Sub ShowPane(CPane As CodePane)

  On Error Resume Next
  With CPane
    If VBInstance.MainWindow.VBE.DisplayModel = vbext_dm_SDI Then
      'v 2.0.6 SDI needs different handling to MDI
      'Morgan Haueisen let me know there was a problem and tracked it to this
      .Window.Show
      .Window.Visible = True
     Else
      '  SendMessage VBInstance.MainWindow.hWnd, WM_SETREDRAW, False, 0
      'ver 2.0.3
      'Only first instance of a find would highlight except by
      'setting window to Visible=False but that produced an ugly
      'flicker as window went and returned
      'Thanks to Brian Barnett <>< (babarnett@mindspring.com)
      'and his Advanced Find/Replace Add-In for the clue to how
      'to do this properly. He used LockWindowUpdate API
      ' but that caused an ugly artifact while locking
      .Window.Visible = False
      '.Window.Visible = True 'Falses.Show
      ' SendMessage VBInstance.MainWindow.hWnd, WM_SETREDRAW, True, 0
    End If
  End With
  On Error GoTo 0

End Sub

Private Sub SimpleSentenceCase(strCode As String)

  Dim PunctSpace As Variant
  Dim Punct      As Variant
  Dim I          As Long

  PunctSpace = Array(RSpacePad("!"), RSpacePad("?"), RSpacePad("."), RSpacePad(","), RSpacePad(":"), RSpacePad(";"))
  Punct = Array("[", "{", LBracket, "<", SQuote, DQuote)
  ProperProperCase strCode, DQuote
  For I = 0 To UBound(Punct)
    ProperProperCase strCode, CStr(Punct(I))
  Next I
  For I = 0 To UBound(PunctSpace)
    ProperProperCase strCode, CStr(PunctSpace(I))
  Next I

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:26:10 AM) 88 + 1386 = 1474 Lines Thanks Ulli for inspiration and lots of code.

