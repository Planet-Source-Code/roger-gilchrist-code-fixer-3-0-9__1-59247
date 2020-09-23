Attribute VB_Name = "mod_WithStructureFix"
Option Explicit

Private Function AcceptableWithTarget(ByVal strCode As String, _
                                      ByVal strTest As String, _
                                      ByVal strConnector As String, _
                                      ByVal hits As Long) As Boolean

  Dim TPos            As Long

  TPos = InStr(strCode, strTest & strConnector)
  If TPos Then
    If Not HasWithUsage(strCode) Then
      If InCode(strCode, TPos) Then
        AcceptableWithTarget = True
        If hits = 1 Then
          If Left$(strCode, 6) = "ElseIf" Then
            AcceptableWithTarget = False
          End If
        End If
      End If
    End If
    '    If AcceptableWithTarget Then
    '      'test that you have not crossed bouudry between code referencing indexed members
    '      LBPos = InStr(StrTest, LBracket)
    '      If LBPos Then
    'AcceptableWithTarget = InStr(strCode, Left$(StrTest, LBPos)) > 0 And InStr(strCode, Mid$(StrTest, InStr(StrTest, LBracket) + 1)) > 0
    '      End If
    '    End If
  End If

End Function

Private Function BlockNoWithDrawCommands(ByVal varTest As Variant, _
                                         ByVal strTest As String) As Boolean

  'Takes care of those occasions where there should be a reference
  'because some commands cannot use With structures
  'ver1.1.79 adjusted line to test for space  (prevents 'codemodule.Lines' being misidentified)

  BlockNoWithDrawCommands = InstrAtPositionArray(varTest, ipAny, False, strTest & ".Line ", strTest & ".Circle ", strTest & ".PSet ", strTest & ".Print ", strTest & ".Scale ")
  If Not BlockNoWithDrawCommands Then
    'v 2.2.1 in case print is used to create a blank line
    BlockNoWithDrawCommands = InstrAtPosition(varTest, ".Print", IpRight, False)
  End If

End Function

Private Function GetModuleType(ByVal strTest As String) As Long

  Dim I As Long

  If bModDescExists Then
    If Len(strTest) Then ' not a res file etc
      For I = LBound(ModDesc) To UBound(ModDesc)
        If strTest = ModDesc(I).MDName Then
          GetModuleType = ModDesc(I).MDTypeNum
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function GetStructureDepthLine(arrR As Variant, _
                                      ByVal LineNo As Long) As Long

  Dim I       As Long
  Dim lngDeep As Long

  For I = 0 To LineNo
    lngDeep = StructureDeep2(lngDeep, arrR(I))
  Next I
  GetStructureDepthLine = lngDeep

End Function

Private Sub GetStructureEnd(arrR As Variant, _
                            ByVal lngCounter As Long, _
                            strTest As String, _
                            EPos As Long)

  Dim I          As Long
  Dim lngEmbeded As Long

  EPos = -1 ' if not found then return -1
  For I = lngCounter To UBound(arrR)
    Select Case strTest
     Case "Do"
      If LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = "Loop" Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "For"
      If LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = "Next" Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "Select Case"
      If InstrAtPosition(arrR(I), strTest, IpLeft) Then
        lngEmbeded = lngEmbeded + 1
       ElseIf InstrAtPosition(arrR(I), "End Select", IpLeft) Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "If"
      'v 2.1.6 Thanks to Tom Law who showed me the bug here
      'the StandardIfThenLine test surronded the whole code so it wasn't finding the end of the structure
      'I had cut and pasted from the other structure detectors but 'If' needs to cope with the special case
      'just incase some one has disabled the expand structures fix so has 'If X Then Y' code lines which don't have an 'End If'  If StandardIfThenLine(arrR, I) Then
      If LeftWord(arrR(I)) = strTest Then
        If StandardIfThenLine(arrR, I) Then
          ' moved inside so it is only conducted on 'If' code lines
          lngEmbeded = lngEmbeded + 1
        End If
      End If
      If InstrAtPosition(arrR(I), "End If", IpLeft) Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "While"
      If LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = "Wend" Then
        lngEmbeded = lngEmbeded - 1
      End If
    End Select
    If lngEmbeded = 0 Then
      EPos = I
      Exit For
    End If
  Next I

End Sub

Private Sub GetStructureTop(arrR As Variant, _
                            ByVal lngCounter As Long, _
                            strTest As String, _
                            SPos As Long)

  Dim I          As Long
  Dim lngEmbeded As Long

  SPos = -1 ' if not found then return -1
  For I = lngCounter To LBound(arrR) Step -1
    Select Case strTest
     Case "Loop"
      If LeftWord(arrR(I)) = "Do" Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "Next"
      If LeftWord(arrR(I)) = "For" Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "End Select"
      If InstrAtPosition(arrR(I), "Select Case", IpLeft) Then
        lngEmbeded = lngEmbeded + 1
       ElseIf InstrAtPosition(arrR(I), strTest, IpLeft) Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "End If"
      'v 2.1.6 Thanks to Tom Law who showed me the bug here
      If LeftWord(arrR(I)) = strTest Then
        If StandardIfThenLine(arrR, I) Then
          ' moved inside so it is only conducted on 'If' code lines
          lngEmbeded = lngEmbeded + 1
        End If
      End If
      If InstrAtPosition(arrR(I), "End If", IpLeft) Then
        lngEmbeded = lngEmbeded - 1
      End If
     Case "Wend"
      If LeftWord(arrR(I)) = "While" Then
        lngEmbeded = lngEmbeded + 1
       ElseIf LeftWord(arrR(I)) = strTest Then
        lngEmbeded = lngEmbeded - 1
      End If
    End Select
    If lngEmbeded = 0 Then
      SPos = I
      Exit For
    End If
  Next I

End Sub

Private Sub GetWithEmbedded(varCode As Variant, _
                            embed As Long)

  If InstrAtPosition(varCode, "With", IpLeft, True) Then
    embed = embed + 1
  End If
  If InstrAtPosition(varCode, "End With", IpLeft, True) Then
    embed = embed - 1
  End If

End Sub

Private Function HasTestBit(varCode As Variant, _
                            strTestBit As String) As Boolean

  HasTestBit = InStr(varCode, strTestBit & ".")
  If Not HasTestBit Then
    HasTestBit = InStr(varCode, strTestBit & "!")
  End If

End Function

Public Function HasWithUsage(varTest As Variant) As Boolean

  If InStr(varTest, ".") Then
    If Left$(Trim$(varTest), 1) = "." Or (InStrCode(varTest, " .") > 0) Or (InStrCode(varTest, "(.") > 0) Then
      HasWithUsage = True
    End If
  End If
  If InStr(varTest, "!") Then
    If Left$(Trim$(varTest), 1) = "!" Or (InStrCode(varTest, " !") > 0) Or (InStrCode(varTest, "(!") > 0) Then
      HasWithUsage = True
    End If
  End If
  'v2.4.2 Thanks Chad Gould. Used in case something that
  'should be in the existing With already is outside of it.
  If Not HasWithUsage Then
    HasWithUsage = Left$(varTest, 5) = "With "
  End If

End Function

Private Function IsEnumOrType(strTest As String) As Boolean

  'v2.4.1 Thanks Paul Caton
  'used to stop with being applied to full <EnumName>.<EnumMember> referenceing of Enum members
  ' which is not safe in With Structures
  '(added Type just in case someone tries that, i'd never seen this used till Paul told me about it)

  IsEnumOrType = InQSortArray(arrQEnumTypePresence, strTest)

End Function

Private Function IsNumericConnector(ByVal strTest As String, _
                                    ByVal TestPos As Long) As Boolean

  'ignore decimal points when testing for with structures
  'v 2.1.9 added bang detector for <number>! code (just in case the removers haven't done it

  If TestPos > 1 Then
    If TestPos < Len(strTest) Then
      If Mid$(strTest, TestPos, 1) = "." Then
        If IsNumeric(Mid$(strTest, TestPos - 1, 1)) Then
          If IsNumeric(Mid$(strTest, TestPos + 1, 1)) Then
            IsNumericConnector = True
          End If
        End If
       ElseIf Mid$(strTest, TestPos, 1) = "!" Then
        If IsNumeric(Mid$(strTest, TestPos - 1, 1)) Then
          IsNumericConnector = True
        End If
      End If
    End If
  End If

End Function

Private Function IsReference(ByVal strTmp As String) As Boolean

  Dim I    As Long
  Dim Proj As VBProject

  'v2.4.1 Thanks Paul Caton
  'used to stop With being applied to full <RefLibname>.<Procedure/Property> referencing of dll members
  ' which is not safe in With Structures
  For Each Proj In VBInstance.VBProjects
    For I = 1 To Proj.References.Count
      On Error GoTo MissingRef
      If Proj.References.Item(I).Name = strTmp Then
        IsReference = True
        GoTo NormalExit
      End If
    Next I
  Next Proj
NormalExit:

Exit Function

MissingRef:
  If Err.Number = -2147319779 Then
    'v2.5.4 added trap '-2147319779 Automation error Library not registered
    'the error is caused by running CF against Dll code without compiling/registering it first
    Err.Number = 0
    Resume Next
  End If

End Function

Private Function LineUnacceptableForWithTextBit(ByVal strCodeLine As String) As Boolean

  If HasWithUsage(strCodeLine) Then
    'With structure already exists so don't start a new one
    LineUnacceptableForWithTextBit = True
   Else
    If isProcHead(strCodeLine) Then
      '     v 2.0.7  procedures need to be avoided (this relates to not moving comments above proc head)
      LineUnacceptableForWithTextBit = True
     Else
      If InstrAtPositionArray(strCodeLine, ipAny, False, ".Line", ".Circle", ".PSet", ".Print", ".Scale", " As New ", "ElseIf", "Else", "ADODB.", "Case ") Then
        '    Line,Circle and Pset and Print methods do not work in With...End With structures
        '    Database code of Set Rs As New DB.RecordSet also fails
        '    'Else' & 'ElseIf' needs to be blocked so that the line is not included as initial line of a With Structure
        '     ADODB. does not supprot the With structure
        LineUnacceptableForWithTextBit = True
      End If
    End If
  End If

End Function

Private Function NotAcceptableForWith(varTest As Variant) As Boolean

  'test for conditions which t may lead to memory leaks in With structures

  If HasWithUsage(varTest) Then
    NotAcceptableForWith = True
   ElseIf InstrAtPositionArray(varTest, IpLeft, True, "Exit Sub", "Exit Function", "Exit Property", "Exit Do", "Exit For", "GoTo", "Case") Then
    'these are legal but will may lead to memory leaks
    NotAcceptableForWith = True
  End If

End Function

Public Sub ProcWithStructurePurify(ArrProc As Variant, _
                                   UpDated As Boolean, _
                                   MissingWith As Boolean, _
                                   Optional Embedded As Long = 0)

  Dim L_CodeLine     As String
  Dim RLine          As Long
  Dim CurLine        As Long
  Dim EndLine        As Long
  Dim TestBit        As String
  Dim OldLine        As String
  Dim DoneatLeastOne As Boolean

  For RLine = LBound(ArrProc) To UBound(ArrProc)
    L_CodeLine = ArrProc(RLine)
    If Not JustACommentOrBlank(L_CodeLine) Then
      If InstrAtPosition(L_CodeLine, "With", IpLeft) Then
        CurLine = RLine + 1
        Embedded = 0
        If WithEmbedded(ArrProc, CurLine, Embedded) Then
          MissingWith = True
          Exit For 'Sub
        End If
        If CurLine > 0 Then
          EndLine = CurLine
          'v 2.1.5 improved testbit fetching to include parametered objects
          'v 2.2.1 added trim just incase an end of line space exists
          TestBit = Trim$(Replace$(strCodeOnly(L_CodeLine), "With ", vbNullString))
          CurLine = RLine + 1
          Do
            OldLine = ArrProc(CurLine)
            'stop if there is an embeded With
            If Left$(OldLine, 5) = "With " Then
              Embedded = Embedded + 1
            End If
            If Embedded = 0 Then
              SuggestWithDoInternalRemoval ArrProc(CurLine), TestBit, UpDated, True
              If OldLine <> ArrProc(CurLine) Then
                Select Case FixData(WithPurity).FixLevel
                 Case CommentOnly
                  'undo change we only want the comment
                  ArrProc(CurLine) = OldLine
                 Case Else
                  'count the change
                  AddNfix WithPurity
                End Select
                DoneatLeastOne = True
              End If
            End If
            'restart after embedded with
            If Left$(OldLine, 8) = "End With" Then
              Embedded = Embedded - 1
            End If
            CurLine = CurLine + 1
            If MultiLeft(ArrProc(CurLine), True, "End Sub", "End Function", "End Property", "End Event") Or CurLine > UBound(ArrProc) Then
              MissingWith = True
              Exit For 'Sub 'GoTo IncompleteWith
            End If
          Loop Until (MultiLeft(ArrProc(CurLine), True, "End With") And Embedded = 0) Or CurLine = EndLine
          If DoneatLeastOne Then
            Select Case FixData(WithPurity).FixLevel
             Case CommentOnly
              ArrProc(RLine) = Marker(ArrProc(RLine), SUGGESTION_MSG & "'With' structure contains unnecessary refereneces to the object " & TestBit, MAfter, UpDated)
             Case Else
              ArrProc(RLine) = Marker(ArrProc(RLine), RGSignature & "Purified With", MAfter, UpDated)
            End Select
            DoneatLeastOne = False
          End If
        End If
      End If
    End If
  Next RLine

End Sub

Public Sub ProcWithStructureRemove(ArrProc As Variant, _
                                   UpDated As Boolean, _
                                   MissingWith As Boolean)

  Dim L_CodeLine    As String
  Dim RLine         As Long
  Dim CurLine       As Long
  Dim EndLine       As Long
  Dim Embedded      As Long
  Dim TestBit       As String
  Dim OldLine       As String
  Dim FoundEmbedded As Boolean   ' allows fix to run again it it finds an embedded with

DealWithEmbedded:
  FoundEmbedded = False
  For RLine = LBound(ArrProc) To UBound(ArrProc)
    L_CodeLine = strCodeOnly(ArrProc(RLine))
    If Not JustACommentOrBlank(L_CodeLine) Then
      If InstrAtPosition(L_CodeLine, "With", IpLeft) Then
        CurLine = RLine + 1
        Embedded = 0
        If WithEmbedded(ArrProc, CurLine, Embedded) Then
          MissingWith = True
         Else
          Exit For 'Sub
        End If
        If CurLine > 0 Then
          EndLine = CurLine
          'v 2.1.5 improved testbit fetching to include parametered objects
          TestBit = Replace$(L_CodeLine, "With ", vbNullString)
          CurLine = RLine + 1
          Do
            OldLine = strCodeOnly(ArrProc(CurLine))
            If Left$(OldLine, 5) = "With " Then
              Embedded = Embedded + 1
              FoundEmbedded = True
            End If
            If Embedded = 0 Then
              SuggestWithDoInternalRestore ArrProc(CurLine), TestBit, ".", UpDated
              SuggestWithDoInternalRestore ArrProc(CurLine), TestBit, "!", UpDated
            End If
            'restart after embedded with
            If Left$(OldLine, 8) = "End With" Then
              Embedded = Embedded - 1
            End If
            CurLine = CurLine + 1
            If MultiLeft(ArrProc(CurLine), True, "End Sub", "End Function", "End Property", "End Event") Or CurLine > UBound(ArrProc) Then
              MissingWith = True
              FoundEmbedded = False
              Exit For
            End If
          Loop Until (SmartLeft(ArrProc(CurLine), "End With") And Embedded = 0) Or CurLine = EndLine
          'v 2.2.1 added a 2nd RgSignature to the comments so that the comments can be edited to restore the With structure
          '        just by removing the the 1st Comment marker and using the Purify tool
          ArrProc(RLine) = Replace$(ArrProc(RLine), "With " & TestBit, vbNullString) & RGSignature & "With " & TestBit & RGSignature & " Removed"
          UpDated = True
          'v 2.2.1 fixed "End With " replaced with "End With"
          ArrProc(CurLine) = Replace$(ArrProc(CurLine), "End With", RGSignature & "End With " & RGSignature & TestBit & " Removed")
        End If
      End If
    End If
  Next RLine
  If FoundEmbedded Then
    GoTo DealWithEmbedded
  End If

End Sub

Private Function StandardIfThenLine(ByVal arrT As Variant, _
                                    ByVal LPos As Long) As Boolean

  Dim strTest As String

  strTest = strCodeOnly(arrT(LPos))
  If HasLineCont(strTest) Then
    Do While HasLineCont(strTest)
      LPos = LPos + 1
      strTest = strTest & strCodeOnly(arrT(LPos))
      If LPos > UBound(arrT) Then
        Exit Do
      End If
    Loop
  End If
  If Right$(strTest, 5) = " Then" Then
    StandardIfThenLine = True
  End If

End Function

Private Function StructuralIntegrity(Sline As Long, _
                                     ELine As Long, _
                                     arrR As Variant) As Boolean

  Dim lngTestBot     As Long
  Dim lngCounter     As Long
  Dim lngStructEnd   As Long
  Dim strTest        As String
  Dim lngStructStart As Long

  'v2.1.3 new version of this test
  ' much more consistent
  '
ReTry:
  'lngTestTop = Sline
  lngTestBot = ELine
  lngCounter = Sline
  Do
    If InstrAtPositionArray(arrR(lngCounter), IpLeft, True, "Do", "For", "Select Case", "If", "While") Then
      strTest = LeftWord(Trim$(arrR(lngCounter)))
      If strTest = "Select" Or strTest = "Case" Then
        strTest = "Select Case"
      End If
      GetStructureEnd arrR, lngCounter, strTest, lngStructEnd
      If lngStructEnd = -1 Then
        Exit Do
      End If
      If lngCounter < lngStructEnd Then
        ' whole structure in range test on
        lngCounter = lngStructEnd + 1 ' exclude end from next test by jumping one line
      End If
      If lngTestBot < lngStructEnd Then
        ' end of whole structure is outside range
        lngTestBot = lngStructEnd
        ELine = lngStructEnd
      End If
     ElseIf InstrAtPositionArray(arrR(lngCounter), IpLeft, True, "Loop", "Next", "End Select", "End If", "Wend") Then
      ' start is inside a structure that starts before it
      strTest = LeftWord(arrR(lngCounter))
      If strTest = "End" Then
        strTest = strTest & " " & WordInString(arrR(lngCounter), 2)
      End If
      GetStructureTop arrR, lngCounter, strTest, lngStructStart
      If lngStructStart = -1 Then
        Exit Do
      End If
      If lngStructStart < Sline Then
        Sline = lngStructStart
        GoTo ReTry
      End If
     ElseIf InstrAtPosition(arrR(lngCounter), "Case", IpLeft) Then
      GetStructureTop arrR, lngCounter, "End Select", lngStructStart
      GetStructureEnd arrR, lngCounter, "Select Case", lngStructEnd
      If lngStructEnd = -1 Then
        Exit Do
      End If
      If lngCounter < lngStructEnd Then
        ' whole structure in range test on
        lngCounter = lngStructEnd + 1 ' exclude end from next test by jumping on line
      End If
      If lngTestBot < lngStructEnd Then
        ' end of whole structure is outside range
        lngTestBot = lngStructEnd
        ELine = lngStructEnd
      End If
      If lngStructStart = -1 Then
        Exit Do
      End If
      If lngStructStart < Sline Then
        Sline = lngStructStart
        GoTo ReTry
      End If
     ElseIf InstrAtPosition(arrR(lngCounter), "ElseIf", IpLeft) Then
      GetStructureTop arrR, lngCounter, "End If", lngStructStart
      GetStructureEnd arrR, lngCounter, "If", lngStructEnd
      If lngStructEnd = -1 Then
        Exit Do
      End If
      If lngCounter < lngStructEnd Then
        ' whole structure in range test on
        lngCounter = lngStructEnd + 1 ' exclude end from next test by jumping on line
      End If
      If lngTestBot < lngStructEnd Then
        ' end of whole structure is outside range
        lngTestBot = lngStructEnd
        ELine = lngStructEnd
      End If
      If lngStructStart = -1 Then
        Exit Do
      End If
      If lngStructStart < Sline Then
        Sline = lngStructStart
        GoTo ReTry
      End If
    End If
    lngCounter = lngCounter + 1
  Loop Until lngCounter > lngTestBot
  If lngStructStart <> -1 Then
    If lngStructEnd <> -1 Then
      StructuralIntegrity = True
    End If
  End If

End Function

Private Function StructureDeep(ByVal CurDepth As Long, _
                               varSearch As Variant) As Long

  Select Case LeftWord(varSearch)
   Case "Select", "For", "Do", "While", "With", "If"
    StructureDeep = CurDepth + 1
   Case "Next", "Loop", "Wend"
    StructureDeep = CurDepth - 1
   Case "End"
    If MultiLeft(varSearch, True, "End If", "End With", "End Select") Then
      StructureDeep = CurDepth - 1
     Else
      StructureDeep = CurDepth 'do nothing
    End If
   Case "Exit"
    StructureDeep = 0
   Case Else
    StructureDeep = CurDepth
  End Select

End Function

Private Function StructureDeep2(ByVal CurDepth As Long, _
                                varSearch As Variant) As Long

  Select Case LeftWord(varSearch)
   Case "Select", "For", "Do", "While", "With", "If"
    StructureDeep2 = CurDepth + 1
   Case "Next", "Loop", "Wend"
    StructureDeep2 = CurDepth - 1
   Case "End"
    If MultiLeft(varSearch, True, "End If", "End With", "End Select") Then
      StructureDeep2 = CurDepth - 1
     Else
      StructureDeep2 = CurDepth 'do nothing
    End If
    '   Case "Exit"
    '    StructureDeep = 0
   Case Else
    StructureDeep2 = CurDepth
  End Select

End Function

Private Sub SuggestWith_Apply(arrR As Variant, _
                              Myi As Long, _
                              TestBit As String, _
                              TopOfWith As Long, _
                              UpDated As Boolean, _
                              CountOfWithBits As Long, _
                              ByVal CompName As String)

  
  Dim IfGuard             As Long
  Dim AddMsgAfterBefore   As Boolean
  Dim ExitIfBackToInit    As Boolean
  Dim BottomOfWith        As Long
  Dim EndOfRoutine        As Long
  Dim J                   As Long
  Dim InitStructuralDepth As Long
  Dim StructuralDepth     As Long
  Dim StrApproval         As String

  If TopOfWith Then
    'v2.1.1 Thanks Tom Law I'd forgotten all about Bang(!) connector and With
    EndOfRoutine = GetEndOfRoutine(arrR)
    'cope with blanks or comments after last routine in module
    'This is a safety mechanism for early developement phase
    'Keeps this routine from going beyond end of routine it is working on
    BottomOfWith = Myi - 1
    'v 2.0.5 thanks tom law a very rare combination of structures caused
    'an over flow here which still activated the fix and applied it falsely
    'expand up(line cont test applies to previous line)
    Do While HasLineCont(arrR(TopOfWith - 1)) Or (HasTestBit(arrR(TopOfWith - 1), TestBit) And Not HasWithUsage(arrR(TopOfWith - 1)))
      'v2.6.5 prevents back tracking including Set lines
      If SmartLeft(arrR(TopOfWith - 1), "Set " & TestBit) Then
        Exit Do
      End If
      TopOfWith = TopOfWith - 1
      If TopOfWith = 0 Then
        Exit Do
      End If
    Loop
    'expand down (line cont test applies to present line)
    Do While HasLineCont(arrR(BottomOfWith)) Or (HasTestBit(arrR(BottomOfWith + 1), TestBit) And Not HasWithUsage(arrR(BottomOfWith + 1)))
      BottomOfWith = BottomOfWith + 1
      If BottomOfWith = UBound(arrR) Then
        Exit Do
      End If
    Loop
    For J = LBound(arrR) To TopOfWith - 1
      InitStructuralDepth = StructureDeep(InitStructuralDepth, arrR(J))
    Next J
    StructuralDepth = InitStructuralDepth
    For J = TopOfWith To EndOfRoutine
      StructuralDepth = StructureDeep(StructuralDepth, arrR(J))
      If StructuralDepth >= InitStructuralDepth Then
        ExitIfBackToInit = True
      End If
      If BottomOfWith < J Then
        If InstrAtPositionArray(arrR(J), IpLeft, True, "Exit", "Case") Then
          'FIXME "Case" is a kludge try to remove it
          'NOTE While legal using 'Exit' inside With structures leads to a
          'subtle memory leak as the With reference is not always cleared properly
          'So unless the 'Exit' is outside any other structures quit rather than risk it.
          GoTo Skip
        End If
      End If
      If BottomOfWith < J Then
        BottomOfWith = J
      End If
      If J >= BottomOfWith Then
        If StructuralDepth = InitStructuralDepth Then
          AddMsgAfterBefore = True
          Exit For
        End If
      End If
      If StructuralDepth = InitStructuralDepth Then
        AddMsgAfterBefore = True
        Exit For
      End If
      If ExitIfBackToInit Then
        If StructuralDepth < InitStructuralDepth Then
          Exit For
        End If
      End If
    Next J
    If BottomOfWith >= EndOfRoutine Then
      'safety should never hit
      Do
        BottomOfWith = EndOfRoutine - 1
      Loop While LenB(arrR(BottomOfWith)) = 0
      AddMsgAfterBefore = True
    End If
    '
    If IsGotoLabel(arrR(BottomOfWith), CompName) Then
      Do
        BottomOfWith = BottomOfWith - 1
      Loop While IsGotoLabel(arrR(BottomOfWith), CompName)
    End If
    If BottomOfWith - TopOfWith > 1 Then
      For J = TopOfWith To BottomOfWith
        If NotAcceptableForWith(arrR(J)) Then
          GoTo Skip
        End If
      Next J
      'v2.5.5 added test to prevent With being inserted incorrectly in Large If structures
      'IfGuard will be 0 if the 'If's and 'End If's balance
      For J = TopOfWith To BottomOfWith
        If InStr(arrR(J), "If ") Then
          IfGuard = IfGuard + 1
         ElseIf InStr(arrR(J), "End If") Then
          IfGuard = IfGuard - 1
        End If
      Next J
      If IfGuard Then
        GoTo Skip
      End If
      If StructuralIntegrity(TopOfWith, BottomOfWith, arrR) Then
        If SmartLeft(arrR(TopOfWith - 1), "'APPROVED(Y) " & TestBit) Then
          'This means that user has approved this change on previous run using Mark only fix mode
          arrR(TopOfWith - 1) = RGSignature & "Auto-inserted With End...With Structure"
          For J = TopOfWith To BottomOfWith
            SuggestWithDoInternalRemoval arrR(J), TestBit, UpDated
          Next J
          arrR(TopOfWith) = "With " & TestBit & vbNewLine & arrR(TopOfWith)
          If Left$(arrR(BottomOfWith), 12) = "'APPROVED(Y)" Then
            'this kludge allows successive Approvals to work
            BottomOfWith = BottomOfWith - 1
          End If
          arrR(BottomOfWith) = Marker(arrR(BottomOfWith), "End With '" & TestBit, IIf(AddMsgAfterBefore, MAfter, MBefore))
         Else
          StrApproval = IIf(FixData(DetectWithStructure).FixLevel > CommentOnly, "'APPROVED(Y) ", "'APPROVED(Y ) ")
          If FixData(DetectWithStructure).FixLevel > CommentOnly Then
            arrR(TopOfWith) = StrApproval & TestBit & vbNewLine & arrR(TopOfWith)
            bWithSuggested = True
           Else
            'bWithSuggested = True
            arrR(TopOfWith) = SUGGESTION_MSG & "Possible Start: With " & TestBit & vbNewLine & _
             StrApproval & TestBit & " [Remove the space after the 'Y' in brackets and next run of Code Fixer will create the With Structure for you." & vbNewLine & _
             arrR(TopOfWith)
            arrR(BottomOfWith) = Marker(arrR(BottomOfWith), SUGGESTION_MSG & "Possible End: End With '" & TestBit, IIf(AddMsgAfterBefore, MAfter, MBefore), UpDated)
          End If
        End If
        UpDated = True
        AddNfix DetectWithStructure
Skip:
        Myi = BottomOfWith '- 1
        CountOfWithBits = 0
        TestBit = vbNullString
        TopOfWith = 0
      End If
    End If
  End If

End Sub

Public Sub SuggestWithDoInternalRemoval(varCode As Variant, _
                                        ByVal strTest As String, _
                                        Hit As Boolean, _
                                        Optional bPurify As Boolean = False)

  Dim Possible As Long
  Dim I        As Long
  Dim J        As Long
  Dim arrCon   As Variant

  'v2.1.1 Thanks Tom Law I'd forgotten all about Bang(!) connector and With
  'This routine removes the With... End With object from lines between structure limits
  ' might be a target line
  arrCon = Array(".", "!")
  If InStr(varCode, strTest & arrCon(0)) Or InStr(varCode, strTest & arrCon(1)) Then
    '\There is already a With Structure around this code
    If Not HasWithUsage(varCode) Or bPurify Then
      '        If InStr(Varcode, b Then
      '            '/so skip it this time. This allows nested With structures to be created
      '            If Left$(Varcode, 1) <> "." Then
      If Not BlockNoWithDrawCommands(strCodeOnly(varCode), strTest) Then
        'skip commands which cannot be used in the With structure
        Hit = True
        'v2.1.6 Thanks Tom Law this bug also showed up if the target was in the same format more than once in a line
        'v2.5.5 Integrated the dot and bang in one cycle because it is possible to use them both in a single structure
        '       and previous version would fail at the HasWithUsage test (2nd) above
        For J = 0 To 1
          Possible = CountSubString(varCode, strTest & arrCon(J))
          For I = 1 To Possible
            varCode = Safe_Replace(varCode, SngSpace & strTest & arrCon(J), SngSpace & arrCon(J))
            varCode = Safe_Replace(varCode, LBracket & strTest & arrCon(J), LBracket & arrCon(J))
            varCode = Safe_Replace(varCode, "-" & strTest & arrCon(J), "-" & arrCon(J))
            varCode = Safe_Replace(varCode, "- " & strTest & arrCon(J), "- " & arrCon(J))
          Next I
          If SmartLeft(varCode, strTest & arrCon(J)) Then
            ' because the previous tests use buffers to avoid hitting similarly named
            ' but different elements they will miss the very common case of target being first element
            'so this section gets such lines but only does a single change
            varCode = Safe_Replace(varCode, strTest & arrCon(J), arrCon(J), , 1)
          End If
        Next J
      End If
    End If
  End If

End Sub

Public Sub SuggestWithDoInternalRestore(varCode As Variant, _
                                        ByVal strTest As String, _
                                        ByVal strConnector As String, _
                                        Hit As Boolean)

  Dim Possible As Long
  Dim I        As Long

  'v2.1.1 Thanks Tom Law I'd forgotten all about Bang(!) connector and With
  'This routine removes the With... End With object from lines between structure limits
  ' might be a target line
  '\There is already a With Structure around this code
  If HasWithUsage(varCode) Then
    '        If InStr(Varcode, b Then
    '            '/so skip it this time. This allows nested With structures to be created
    '            If Left$(Varcode, 1) <> "." Then
    'skip commands which cannot be used in the With structure
    Hit = True
    Possible = CountSubString(varCode, strConnector)
    'v2.1.6 Thanks Tom Law this bug also showed up if the target was in the same format more than once in a line
    For I = 1 To Possible
      varCode = Safe_Replace(varCode, SngSpace & strConnector, SngSpace & strTest & strConnector)
      varCode = Safe_Replace(varCode, LBracket & strConnector, LBracket & strTest & strConnector)
      varCode = Safe_Replace(varCode, "-" & strConnector, "-" & strTest & strConnector)
      varCode = Safe_Replace(varCode, "- " & strConnector, "- " & strTest & strConnector)
    Next I
    If Left$(Trim$(varCode), 1) = "." Then
      ' because the previous tests use buffers to avoid hitting similarly named
      ' but different elements they will miss the very common case of target being first element
      'so this section gets such lines but only does a single change
      varCode = Safe_Replace(varCode, strConnector, strTest & strConnector, , 1)
    End If
  End If

End Sub

Public Sub SuggestWithStructure(cMod As CodeModule)

  Dim ModuleNumber As Long
  Dim I            As Long
  Dim arrMembers   As Variant
  Dim UpDated      As Boolean
  Dim MaxFactor    As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Identify potential locations for applying With...End With stuctures
  'also carries out auto code if yser approved a structure previous run
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, DetectWithStructure) Then
    arrMembers = GetMembersArray(cMod)
    MaxFactor = UBound(arrMembers)
    If MaxFactor > 0 Then
      'ver 2.0.9 Thanks Tom Law the With in Declarations bug was here
      MaxFactor = UBound(arrMembers)
      For I = 1 To MaxFactor
        MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
        SuggestWithStructureProcedure arrMembers(I), cMod.Parent.Name, UpDated
      Next I
      ReWriteMembers cMod, arrMembers, UpDated
    End If
  End If

End Sub

Private Function SuggestWithStructureGetTestBit(ByVal strCodeLine As String, _
                                                ByVal COfBits As Long, _
                                                ByVal ClinNo As Long, _
                                                ByVal strConnector As String, _
                                                TopWith As Long) As String

  Dim TPos   As Long
  Dim strTmp As String
  Dim EqPos  As Long

  'Support routine for SuggestWithStructure
  'extracts a possible target word for that routine
  'NOTE StrCodeLine is code only
  If Not LineUnacceptableForWithTextBit(strCodeLine) Then
    TPos = InStr(strCodeLine, strConnector)
    If TPos Then
      If Not IsNumericConnector(strCodeLine, TPos) Then
        'v2.8.3 Thanks Joakim Schramm stops detection of Variables with Typesuffix !(Single)
        'v2.9.8 ignore calls to procedures using the proc typesuffix!
        If IsPunct(Mid$(strCodeLine, TPos + 1, 1)) = False Then
          If Mid$(strCodeLine, TPos + 1, 1) <> " " Then
            strTmp = Left$(strCodeLine, TPos - 1)
            If LenB(strTmp) Then
              EqPos = InStr(strTmp, EqualInCode)
              If EqPos Then
                strTmp = Mid$(strTmp, EqPos + 3)
              End If
              If GetLeftBracketPos(strTmp) > 0 Then
                If GetRightBracketPos(strTmp) > 0 Then
                  If CountSubString(strTmp, LBracket) > CountSubString(strTmp, RBracket) Then
                    Do
                      strTmp = Mid$(strTmp, GetLeftBracketPos(strTmp) + 1)
                    Loop While CountSubStringImbalance(strTmp, LBracket, RBracket)
                   ElseIf CountSubString(strTmp, LBracket) < CountSubString(strTmp, RBracket) Then
                    Do
                      'ver 1.1.93 this locked up if last char was the ")"
                      'strTmp = Left$(strTmp, InStrRev(strTmp, ")") + 1)
                      strTmp = Left$(strTmp, InStrRev(strTmp, RBracket) - 1)
                      'ver 1.1.93 also changed = to >= in next line (in fact any line that hits here will almost certainly fail
                    Loop Until CountSubString(strTmp, LBracket) >= CountSubString(strTmp, RBracket)
                  End If
                End If
              End If
              Do While GetSpacePos(ConcealParameterSpaces(strTmp))
                'remove more than one word unless it includes an index
                strTmp = Mid$(strTmp, InStrRev(ConcealParameterSpaces(strTmp), SngSpace) + 1)
              Loop
              If IsNumeric(strTmp) Then
                strTmp = vbNullString
              End If
              If Left$(strTmp, 1) = "-" Then
                strTmp = Mid$(strTmp, 2)
              End If
              If COfBits = 0 Then
                If LenB(strTmp) Then
                  TopWith = ClinNo
                End If
              End If
              'If CountSubString(strTmp, LBracket) <> CountSubString(strTmp, RBracket) Then
              If CountSubStringImbalance(strTmp, LBracket, RBracket) Then
                If ((CountSubString(strTmp, LBracket) = 1 And CountSubString(strTmp, RBracket) = 0)) Then
                  strTmp = Mid$(strTmp, GetLeftBracketPos(strTmp) + 1)
                 Else
                  strTmp = vbNullString
                End If
              End If
              'you can't use With structure with VBA calls
              If UCase$(strTmp) = "VBA" Then
                strTmp = vbNullString
              End If
              'ver 1.0.93 Fixed Full path references to bas module methods do not work in With structure
              If InQSortArray(QSortModBasArray, strTmp) Then
                If GetModuleType(strTmp) = vbext_ct_StdModule Then
                  strTmp = vbNullString
                End If
              End If
              If IsEnumOrType(strTmp) Then
                strTmp = vbNullString
              End If
              If IsReference(strTmp) Then
                strTmp = vbNullString
              End If
              If Left$(strTmp, 1) = "." Then
                strTmp = vbNullString
              End If
              SuggestWithStructureGetTestBit = strTmp
            End If
          End If
        End If
      End If
    End If
  End If

End Function

Public Sub SuggestWithStructureProcedure(VarProc As Variant, _
                                         strModName As String, _
                                         UpDated As Boolean)

  Dim ArrRoutine      As Variant
  Dim J               As Long
  Dim L_CodeLine      As String
  Dim TestBit         As String
  Dim TestBitA        As String
  Dim TestBitB        As String
  Dim TopOfWith       As Long
  Dim CountOfWithBits As Long
  Dim I               As Long
  Dim strOutsider     As String
  Dim arrTest         As Variant

  'v2.4.3 included ReDim in the list
  arrTest = Array("With", "Load", "Set", "Unload", "Next", "RaiseEvent", "End", "ReDim")
  ArrRoutine = Split(VarProc, vbNewLine)
  TestBit = vbNullString
  TopOfWith = GetProcCodeLineOfRoutine(ArrRoutine) + 1
  For J = GetProcCodeLineOfRoutine(ArrRoutine) To UBound(ArrRoutine)
    L_CodeLine = strCodeOnly(ArrRoutine(J))
    If Not JustACommentOrBlank(L_CodeLine) Then
      If Not IsDimLine(L_CodeLine) Then
        If Not isProcHead(L_CodeLine) Then
          If LenB(TestBit) Then
            If AcceptableWithTarget(L_CodeLine, TestBit, ".", CountOfWithBits) Or AcceptableWithTarget(L_CodeLine, TestBit, "!", CountOfWithBits) Then
              CountOfWithBits = CountOfWithBits + 1
             ElseIf CountOfWithBits > 1 Then
              SuggestWith_Apply ArrRoutine, J, TestBit, TopOfWith, UpDated, CountOfWithBits, strModName
             Else
              'v2.4.2 thanks to Chad Gould whose 'ListViewGraphical' code helped me find a bug
              ' in CF caused by code that suggested the need for this comment.
              If SmartLeft(L_CodeLine, "With " & TestBit) Then
                If InStr(ArrRoutine(J - 1), SUGGESTION_MSG & " The previous ") = 0 Then
                  For I = TopOfWith To TopOfWith + CountOfWithBits + 1
                    strOutsider = LeftWord(ArrRoutine(I))
                    If IsInArray(strOutsider, arrTest) Then
                      GoTo noPrevMsg
                    End If
                    If IsProcedure(strOutsider) Then
                      GoTo noPrevMsg
                    End If
                  Next I
                  ArrRoutine(J - 1) = ArrRoutine(J - 1) & vbNewLine & _
                   SUGGESTION_MSG & " The previous " & IIf(CountOfWithBits = 0, "line", CountOfWithBits + 1 & " lines") & " may be able to be placed within the following 'With' structure."
                  UpDated = True
noPrevMsg:
                End If
              End If
              CountOfWithBits = 0
              TestBit = vbNullString
              TopOfWith = 0 'GetProcCodeLineOfRoutine(ArrRoutine) + 1
            End If
           Else
            TestBitA = SuggestWithStructureGetTestBit(L_CodeLine, CountOfWithBits, J, ".", TopOfWith)
            TestBitB = SuggestWithStructureGetTestBit(L_CodeLine, CountOfWithBits, J, "!", TopOfWith)
            If LenB(TestBitA) Then
              TestBit = TestBitA
             ElseIf LenB(TestBitB) Then
              TestBit = TestBitB
             Else
              TestBit = vbNullString
            End If
          End If
        End If
      End If
    End If
  Next J
  VarProc = Join(ArrRoutine, vbNewLine)

End Sub

Private Function WithEmbedded(arrR As Variant, _
                              curlin As Long, _
                              ByVal embed As Long) As Boolean

  Dim strTest As String

  'v2.2.0 strtest added for procedure rightMenu support
  strTest = Trim$(arrR(curlin))
  Do
    If MultiLeft(strTest, True, "End Sub", "End Function", "End Property", "End Event") Or curlin > UBound(arrR) Then
      WithEmbedded = True
      Exit Do 'Function
    End If
    'NEW 1193 may be too strong Designed to deal with error in 'PGN Reader'
    If Left$(strTest, 8) = "End With" Then
      Exit Do
    End If
    GetWithEmbedded strTest, embed
    curlin = curlin + 1
    If curlin > UBound(arrR) Then
      WithEmbedded = True
      Exit Do 'Function
    End If
    strTest = Trim$(arrR(curlin))
  Loop Until Left$(strTest, 8) = "End With" And embed = 0

End Function

Public Sub WithStructurePurityFast(cMod As CodeModule)

  Dim PSLine      As Long
  Dim PEndLine    As Long
  Dim ArrProc     As Variant
  Dim Sline       As Long
  Dim WSPmsg      As String
  Dim UpDated     As Boolean
  Dim MissingWith As Boolean
  Dim Embedded    As Long

  With cMod
    If dofix(ModDescMember(.Parent.Name), WithPurity) Then
      Do While .Find("With", Sline, 1, -1, -1, True, True)
        If InStr(.Lines(Sline, 1), "With") = 1 Then
          ArrProc = ReadProcedureCodeArray2(cMod, Sline, PSLine, PEndLine)
          MemberMessage GetProcName(cMod, Sline), Sline, .CountOfLines
          ProcWithStructurePurify ArrProc, UpDated, MissingWith, Embedded
          If MissingWith Then
            Exit Do
          End If
InsertErrorMEssage:
          If UpDated Then
            ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine, False
          End If
        End If
        Sline = Sline + IIf(PEndLine - PSLine > 1, PEndLine - PSLine, 1)
        If Sline > .CountOfLines Then
          Exit Do
        End If
      Loop
    End If
  End With

Exit Sub

  MissingWith = False
  'This is a safety mechanism should not hit unless corresponding End With is missing
  'It inserts a comment, displays a msgbox then continues processing
  WSPmsg = WARNING_MSG & ":ERROR: The With Structure Purity system has failed because the expected 'End With' was missing." & vbNewLine & _
   "Code Fixer has placed an 'End With' here but it may not be correct." & vbNewLine & _
   "A Structural Error should occur if this will not work. Inspect the code or run with [Ctrl]+[F5]"
  If Embedded = -1 Then
    ArrProc(GetEndOfRoutine(ArrProc)) = Marker(ArrProc(GetEndOfRoutine(ArrProc)), WARNING_MSG & "Purify With had a problem with this procedure; possible empty With structure.", MBefore, UpDated)
   Else
    ArrProc(GetEndOfRoutine(ArrProc)) = Marker("End With" & vbNewLine & _
     ArrProc(GetEndOfRoutine(ArrProc)), WSPmsg, MBefore, UpDated)
    mObjDoc.Safe_MsgBox WSPmsg & vbNewLine & "Click OK to continue", vbCritical
  End If
  UpDated = True
  GoTo InsertErrorMEssage

End Sub

'

':)Code Fixer V3.0.9 (25/03/2005 4:27:03 AM) 1 + 1117 = 1118 Lines Thanks Ulli for inspiration and lots of code.

