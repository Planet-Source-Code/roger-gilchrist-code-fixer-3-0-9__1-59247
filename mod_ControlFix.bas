Attribute VB_Name = "mod_ControlFix"
Option Explicit

Public Function ApplyDefProp(VarOriginal As Variant, _
                             ByVal DefProp As String, _
                             UpDated As Boolean, _
                             CommentUpdated As Boolean) As String

  If LenB(DefProp) Then
    ApplyDefProp = VarOriginal & "." & DefProp & SngSpace
    UpDated = True
   Else
    ApplyDefProp = VarOriginal
    CommentUpdated = True
  End If

End Function

Public Function BadNameType(varCtrlName As Variant) As Long

  Dim I As Long

  I = CntrlDescMember(varCtrlName)
  If I > -1 Then
    BadNameType = CntrlDesc(I).CDBadType
  End If

End Function

Public Function ControlNameBase(ByVal strCtrl As String) As String

  ControlNameBase = strCtrl
  If InStr(ControlNameBase, LBracket) Then
    ControlNameBase = Left$(ControlNameBase, InStr(ControlNameBase, LBracket) - 1)
  End If

End Function

Public Function ControlNeedsFormRef(ByVal CtrlName As String, _
                                    ByVal CompName As String) As Boolean

  Dim I   As Long
  Dim J   As Long
  Dim Hit As Boolean

  If bCtrlDescExists Then
    CtrlName = ControlNameBase(CtrlName)
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDForm = CompName Then
        For J = I To UBound(CntrlDesc)
          If CntrlDesc(J).CDForm <> CompName Then
            ControlNeedsFormRef = True
            Hit = True
            Exit For
           Else
            If CntrlDesc(J).CDName = CtrlName Then
              ControlNeedsFormRef = False
              Hit = True
              Exit For
            End If
          End If
        Next J
      End If
      If Hit Then
        Exit For
      End If
    Next I
  End If

End Function

Public Function DefFormType(ByVal CmpType As Long) As String

  Select Case CmpType
   Case vbext_ct_DocObject
    DefFormType = "UserDocument"
   Case vbext_ct_UserControl
    DefFormType = "UserControl"
   Case vbext_ct_MSForm, vbext_ct_VBMDIForm, vbext_ct_VBForm
    DefFormType = "Me"
  End Select

End Function

Public Function ExpandForDetection2(ByVal VarStr As Variant) As Variant

  Dim arrLine    As Variant
  Dim I          As Long
  Dim L_CodeLine As String
  Dim J          As Long
  Dim strComment As String

  L_CodeLine = VarStr
  ExtractCode L_CodeLine, strComment
  arrLine = Split(L_CodeLine)
  If InStr(L_CodeLine, DQuote) Then
    CompactLiteral arrLine
  End If
  For I = LBound(arrLine) To UBound(arrLine)
    If InStr(arrLine(I), DQuote) = 0 Then
      arrLine(I) = ReplaceArray(arrLine(I), LBracket, SpacePad(LBracket), RBracket, SpacePad(RBracket), "-", SpacePad("-"), ",", SpacePad(","), ":=", SpacePad(":="), "= -", SpacePad("= -"))
      arrLine(I) = Replace$(arrLine(I), DblSpace, SngSpace)
      'ver 2.0.3 David Carter found a bug with
      'rs(0)!Name gets changed to "rs ( 0 ) !Name"
      'this line puts the bang back against the right bracket
      arrLine(I) = Replace$(arrLine(I), ") !", ")!")
      'v2.1.4 Thanks Tom Law do'h this was it, should have fixed when I did above
      arrLine(I) = Replace$(arrLine(I), ") .", ").")
    End If
  Next I
  'special case X = -FunctionName(y)
  'ver 1.0.94 thanks to Martins Skujenieks 's 'FX.DLL 1.03 SDK' which made this fix necessary
  For I = LBound(arrLine) To UBound(arrLine)
    If LenB(arrLine(I)) Then
      If Left$(arrLine(I), 1) <> SQuote Then
        If InStr(arrLine(I), LBracket) Then
          J = 1
          Do While CountCodeSubString(arrLine(I), LBracket) <> CountCodeSubString(arrLine(I), RBracket)
            If I + J > UBound(arrLine) Then
              Exit Do
            End If
            arrLine(I) = arrLine(I) & SngSpace & arrLine(I + J)
            arrLine(I + J) = vbNullString
            J = J + 1
            If I + J > UBound(arrLine) Then
              Exit Do
            End If
          Loop
        End If
      End If
    End If
  Next I
  ExpandForDetection2 = CleanArray(arrLine)
  If Len(strComment) Then
    ExpandForDetection2 = AppendArray(ExpandForDetection2, strComment)
  End If

End Function

Public Function FindDimLine(ByVal varFind As Variant, _
                            ByVal arrTmp As Variant) As Long

  Dim I As Long

  FindDimLine = -1
  For I = LBound(arrTmp) To UBound(arrTmp)
    If IsDimLine(arrTmp(I)) Then
      If InstrAtPosition(arrTmp(I), varFind, ip2nd, True) Then
        FindDimLine = I
        Exit For
      End If
    End If
  Next I

End Function

Public Sub FreeStandingProperty(Comp As VBComponent, _
                                tmpB As Variant, _
                                ByVal J As Long, _
                                L_CodeLine As String, _
                                CFMsg3 As String, _
                                ExistingDimArray As Variant, _
                                TmpBSize As Long)

  Dim oldCodeLine As String

  If Not InQSortArray(ExistingDimArray, tmpB(J)) Then
    If Not isVBCommandNotProperty(tmpB, J) Then
      If Not IsDeclareName(tmpB(J)) Then 'IsDeclaration(tmpB(J)) Then
        oldCodeLine = L_CodeLine
        If LenB(DefFormType(Comp.Type)) Then
          If DefFormType(Comp.Type) = "UserControl" And IsProcedure(CStr(tmpB(J)), "Public", "Property") Then
            L_CodeLine = Safe_Replace(oldCodeLine, SngSpace & tmpB(J), SngSpace & " Me." & tmpB(J))
          End If
          If SmartLeft(L_CodeLine, tmpB(J)) Then
            If DefFormType(Comp.Type) = "UserControl" And IsProcedure(CStr(tmpB(J)), "Public", "Property") Then
              L_CodeLine = Safe_Replace(L_CodeLine, SngSpace & tmpB(J), SngSpace & " Me." & tmpB(J), , 1)
             Else
              L_CodeLine = Safe_Replace(L_CodeLine, tmpB(J), DefFormType(Comp.Type) & "." & tmpB(J), , 1)
            End If
          End If
        End If
        If oldCodeLine <> L_CodeLine Then
          tmpB = Split(ExpandForDetection(Trim$(L_CodeLine)))
          TmpBSize = UBound(tmpB)
          If LenB(DefFormType(Comp.Type)) Then
            CFMsg3 = CFMsg3 & WARNING_MSG & "It is clearer to use the " & DefFormType(Comp.Type) & " reference than depend on VB's auto behaviour."
           Else
            CFMsg3 = CFMsg3 & WARNING_MSG & "The referencing '.' inserted above may be incorrect."
          End If
        End If
      End If
    End If
  End If

End Sub

Public Sub InsertDefProp(Codeline As String, _
                         CtrlName As Variant, _
                         ByVal DefProp As String, _
                         UpDated As Boolean, _
                         CommentUpdated As Boolean, _
                         Optional InsertFullQualified As Boolean = False)

  Dim DoX         As String
  Dim strTmp      As String
  Dim FrmRef      As String
  Dim lngTmpIndex As Long

  DoX = ApplyDefProp(CtrlName, DefProp, UpDated, CommentUpdated)
  strTmp = Codeline
  Codeline = Safe_Replace(strTmp, CtrlName, DoX, , , True, True, True)
  If InsertFullQualified Then
    If Len(DoX) Then
      lngTmpIndex = CntrlDescMember(ControlNameBase(CtrlName))
      If lngTmpIndex > -1 Then
        If CntrlDesc(lngTmpIndex).CDBadType = 0 Then
          FrmRef = CntrlDesc(lngTmpIndex).CDForm
        End If
        If Len(Trim$(FrmRef)) Then
          If BadNameType(CtrlName) <> BNMultiForm Then
            Codeline = Safe_Replace(Codeline, Trim$(DoX), FrmRef & "." & DoX, , , True, True, True)
            InsertFullQualified = True
          End If
        End If
      End If
    End If
  End If
  If strTmp = Codeline Then
    UpDated = False
    CommentUpdated = False
  End If

End Sub

Public Function isCalculatedParameter(ByVal arrTmp As Variant, _
                                      ByVal TPos As Long) As Boolean

  Dim I       As Long
  Dim arrTest As Variant

  arrTest = Array("*", "/", "\", "+", "-")
  For I = LBound(arrTmp) To UBound(arrTmp)
    If IsProcedure(CStr(arrTmp(I)), "", "Sub") Or IsProcedure(CStr(arrTmp(I)), "", "Function") Then
      If TPos > I Then
        If IsInArray(arrTmp(TPos - 1), arrTest) Then
          isCalculatedParameter = True
          Exit For
        End If
      End If
      If TPos < UBound(arrTmp) Then
        If IsInArray(arrTmp(TPos + 1), arrTest) Then
          isCalculatedParameter = True
          Exit For
        End If
      End If
    End If
  Next I

End Function

Public Function IsControlMethod(ByVal strTest As String, _
                                Optional ByVal strCaseFix As String) As Boolean

  Dim I As Long

  For I = LBound(ArrQActiveControlClass) To UBound(ArrQActiveControlClass)
    If TLibMethodFinder(strTest, CStr(ArrQActiveControlClass(I)), strCaseFix) Then
      IsControlMethod = True
      Exit For
    End If
  Next I

End Function

Public Function IsParameter2(ByVal arrTmp As Variant, _
                             ByVal TPos As Long, _
                             ByVal CompName As String, _
                             Optional ByVal oldStyle As Boolean = True) As Boolean

  
  Dim Tval    As Variant
  Dim I       As Long
  Dim arrTest As Variant

  arrTest = Array("*", "/", "\", "+", "-")
  'v2.3.8 modification for faster test
  Tval = arrTmp(0)
  If ArrayHasContents(arrTmp) Then
    If Tval = "Call" Or Tval = "Set" Then
      IsParameter2 = True
      GoTo SingleExit
    End If
    For I = 0 To TPos
      If IsProcedure(CStr(arrTmp(I))) Then
        IsParameter2 = True
        GoTo SingleExit
      End If
      'v2.9.1 being passed to a procedure adressed by Form.Proc in a With structure. Thanks Richard Brisley
      If InStr("!.", Left$(arrTmp(I), 1)) Then
        If IsProcedure(Mid$(arrTmp(I), 2)) Then
          IsParameter2 = True
          GoTo SingleExit
        End If
      End If
      If InStr(arrTmp(I), ".") Then
        If IsProcedure(Mid$(arrTmp(I), InStr(arrTmp(I), ".") + 1), "") Then
          IsParameter2 = True
          GoTo SingleExit
        End If
      End If
    Next I
    If TPos > 0 Then
      Tval = arrTmp(TPos - 1)
      'v2.3.7 >>
      TPos = TPos - 1
     Else
      GoTo SingleExit
    End If
    If IsPunct(Tval) Then
      If TPos > 1 Then
        If TPos < UBound(arrTmp) Then
          If IsInArray(arrTmp(TPos - 1), arrTest) Or IsInArray(arrTmp(TPos + 1), arrTest) Then
            'it's being used in a calculation so almost certainly not the control being used
            IsParameter2 = False
            GoTo SingleExit
          End If
        End If
        If TPos = UBound(arrTmp) Then
          If IsInArray(arrTmp(TPos - 1), arrTest) Then
            'it's being used in a calculation so almost certainly not the control being used
            IsParameter2 = False
            GoTo SingleExit
          End If
        End If
      End If
      If Not ArrayMember(arrTmp(TPos - 1), "=", "&") Then
        If oldStyle Then
          IsParameter2 = True 'IT IS ASSIGNING THE CONTROL TO SOME THING THAT CAN ACCEPT IT
          GoTo SingleExit
         Else
          GoTo PsoPAr
        End If
      End If
      If Tval = LBracket Then
        If TPos > 1 Then
          Tval = arrTmp(TPos - 2)
          IsParameter2 = False
        End If
       Else
        If oldStyle Then
          GoTo SingleExit
        End If
      End If
      ' Else
    End If
PsoPAr:
    If isCalculatedParameter(arrTmp, TPos) Then
      ' if a control is being used in a calculation then it must be using the default value
      IsParameter2 = False
    End If
    'just in case is qualified call
    If InStr(Tval, ".") Then
      Tval = Mid$(Tval, InStr(Tval, ".") + 1)
    End If
    'v2.8.4 new tests to ignore impossible parameters
    '    If IsNumeric(Tval) Then
    '    IsParameter2 = True ' this is a dummy result that allow it to ignore these
    '    ElseIf IsPunct(CStr(Tval)) Then
    '    'IsParameter2 = True ' this is a dummy result that allow it to ignore these
    '    Else
    If IsControlMethod(CStr(Tval)) Then
      IsParameter2 = True
     ElseIf IsControlEvent(CStr(Tval)) Then
      IsParameter2 = True
     ElseIf IsControlEvent(CStr(Tval)) Then
      IsParameter2 = True
     ElseIf IsControlMethod(CStr(Tval)) Then
      IsParameter2 = True
     ElseIf Not TestParam("Function", Tval, arrTmp, TPos + 1) Then
      IsParameter2 = True
     ElseIf Not TestParam("Sub", Tval, arrTmp, TPos + 1) Then
      IsParameter2 = True
     ElseIf Not TestParam("Sub", Tval, arrTmp, TPos + 1) Then
      IsParameter2 = True
     ElseIf Not TestParam("Property", Tval, arrTmp, TPos + 1) Then
      IsParameter2 = True
     ElseIf IsProcedure(CStr(Tval), "Public", "Function") Or IsProcedure(CStr(Tval), "Private", "Function", CompName) Then
      'This works but there is a problem with control arrays
      IsParameter2 = TestParam("Function", Tval, arrTmp, TPos + 1)
     ElseIf IsProcedure(CStr(Tval), "Public", "Sub") Or IsProcedure(CStr(Tval), "Private", "Sub", CompName) Then
      IsParameter2 = TestParam("Sub", Tval, arrTmp, TPos + 1)
     ElseIf IsProcedure(CStr(Tval), "Public", "Property") Or IsProcedure(CStr(Tval), "Private", "Property", CompName) Then
      IsParameter2 = TestParam("Property", Tval, arrTmp, TPos + 1)
     ElseIf IsProcedure(CStr(arrTmp(TPos))) Then
      IsParameter2 = True
     Else
      IsParameter2 = False
    End If
  End If
SingleExit:

End Function

Public Function isVBCommandNotProperty(ByVal VarArray As Variant, _
                                       ByVal MemNo As Long) As Boolean

  'Tests if Left or Right is a Property  < Me.Left > or Command  < Right(x, 1) >
  'Has to be tested as the Code which adds the $ to Left/Right commands hits after the routine that fixes Default Properties

  If ArrayMember(VarArray(MemNo), "Left", "Right") Then
    If MemNo + 1 <= UBound(VarArray) Then
      isVBCommandNotProperty = VarArray(MemNo + 1) = LBracket
    End If
  End If

End Function

Public Function isVBReservedWord(ByVal strTest As String) As Boolean

  isVBReservedWord = InQSortArray(ArrQVBReservedWords, strTest)

End Function

Public Function notBadNameOrFixableBad(ByVal varCtrlName As Variant) As Boolean

  Dim lngTmpIndex As Long

  'ver 1.1.42 tests for control names that are unsafe (for Code Fixer)
  If Not IsDeclareName(varCtrlName) Then 'IsDeclaration(varCtrlName) Then
    'v2.3.7 stop hitting on Module level declarations
    If InStr(varCtrlName, LBracket) Then
      varCtrlName = Left$(varCtrlName, InStr(varCtrlName, LBracket) - 1)
    End If
    lngTmpIndex = CntrlDescMember(varCtrlName)
    If lngTmpIndex > -1 Then
      If CntrlDesc(lngTmpIndex).CDBadType = 0 Then
        'not bad so go ahead
        notBadNameOrFixableBad = True
       Else
        If ArrayMember(CntrlDesc(lngTmpIndex).CDBadType, BNClass, BNMultiForm, BNDefault, BNSingle) Then
          'not so bad as to stop the tool working
          notBadNameOrFixableBad = True
        End If
        notBadNameOrFixableBad = True
      End If
    End If
  End If

End Function

Public Function ParamSearch(ByVal strFind As String) As String

  Dim CompMod      As CodeModule
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim StartLine    As Long
  Dim CurCompCount As Long

  If LenB(strFind) Then
    On Error Resume Next
    If LenB(strFind) > 0 Then
      CurCompCount = 0
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If SafeCompToProcess(Comp, CurCompCount) Then
            Set CompMod = Comp.CodeModule
            StartLine = 1 'initialize search range
            If CompMod.Find(strFind, StartLine, 1, -1, -1, False, True, False) Then
              ParamSearch = CompMod.Lines(StartLine, 1)
              GoTo GotIt
            End If
          End If
        Next Comp
      Next Proj
GotIt:
      Set Comp = Nothing
      Set CompMod = Nothing
      Set Proj = Nothing
    End If
    On Error GoTo 0
  End If

End Function

Public Function SafeReattachComment(VarLine As Variant, _
                                    ByVal strComment As String) As String

  If InStr(VarLine, strComment) = 0 Then
    SafeReattachComment = strComment
  End If

End Function

Public Function TestParam(ByVal procClass As String, _
                          Tval As Variant, _
                          arrLine As Variant, _
                          ByVal TPos As Long) As Boolean

  Dim lngTmpIndex As Long
  Dim I           As Long
  Dim CheckParam  As Long
  Dim Codeline    As String
  Dim PLine       As String
  Dim arrP        As Variant
  Dim Pcode       As String
  Dim arrC        As Variant

  TestParam = True
  If procClass = "Property" Then
    'Property might be any of these but the searches ar in order of probability for this
    Codeline = ParamSearch(procClass & " Let " & Tval)
    If LenB(Codeline) = 0 Then
      Codeline = ParamSearch(procClass & " Set " & Tval)
    End If
    If LenB(Codeline) = 0 Then
      Codeline = ParamSearch(procClass & " Get " & Tval)
    End If
    'v2.3.8
   Else
    Codeline = ParamSearch(procClass & SngSpace & Tval)
  End If
  If Len(Codeline) Then
    If InStr(Codeline, LBracket) Then
      If InStr(Codeline, RBracket) Then
        PLine = Mid$(Codeline, InStr(Codeline, LBracket) + 1)
        PLine = Left$(PLine, InStrRev(PLine, RBracket) - 1)
        arrP = Split(PLine, CommaSpace)
        Pcode = Join(arrLine)
        Pcode = Trim$(Mid$(Pcode, InStr(Pcode, Tval) + Len(Tval)))
        Pcode = Trim$(RStrip(LStrip(Pcode, LBracket), RBracket))
        arrC = Split(Pcode, CommaSpace)
        For I = LBound(arrC) To UBound(arrC)
          If InStrWholeWordRX(arrC(I), arrLine(TPos)) Then
            CheckParam = I
            Exit For
          End If
        Next I
        'This detects if parameters are of the correct type or should be changed
        'Disabled because it is buggy in some circumstances (two different references being passed ie (text1, Text1(2).Text)
        'where text1 is the whole array being passed as a variant
        If UBound(arrP) > -1 Then
          If CheckParam < UBound(arrP) Then
            If Has_AS(arrP(CheckParam)) Then
              lngTmpIndex = CntrlDescMember(arrLine(TPos))
              If lngTmpIndex > -1 Then
                'v2.3.8
                If GetType(arrP(CheckParam)) = CntrlDesc(lngTmpIndex).CDClass Then
                  ' ReferenceLibraryControlDefaultPropertyType(CntrlDesc(lngTmpIndex).CDClass) Then
                  TestParam = False
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:25:29 AM) 1 + 536 = 537 Lines Thanks Ulli for inspiration and lots of code.

