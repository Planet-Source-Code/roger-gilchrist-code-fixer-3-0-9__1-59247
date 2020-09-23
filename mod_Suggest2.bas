Attribute VB_Name = "mod_Suggest2"
Option Explicit
Public Enum TypeRange
  TRBool
  TRByte
  TRCur
  TRDate
  TRDbl
  TRDec
  TRInt
  TRLng
  TRSng
  TRStr
  TRVar
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private TRBool, TRByte, TRCur, TRDate, TRDbl, TRDec, TRInt, TRLng, TRSng, TRStr, TRVar
#End If
Private Const EmbedMsg      As String = ":EMBED:HERE:"

Public Function FirstCodeLineInProc(arrR As Variant) As Long

  ''v2.9.6 keep comments at top of procedure

  FirstCodeLineInProc = GetProcCodeLineOfRoutine(arrR, True)
  Do
    FirstCodeLineInProc = FirstCodeLineInProc + 1
  Loop While JustACommentOrBlank(arrR(FirstCodeLineInProc))
  FirstCodeLineInProc = FirstCodeLineInProc - 1

End Function

Public Function GetDeclarationType(ByVal strTest As String, _
                                   Optional ByVal StrScope As String, _
                                   Optional ByVal strClass As String, _
                                   Optional ByVal strOnForm As String) As String

  Dim I            As Long
  Dim strScopePlus As String

  If bDeclExists Then
    If StrScope = "Public" Then
      strScopePlus = "Friend"
     ElseIf StrScope = "Private" Then
      strScopePlus = "Static"
    End If
    For I = LBound(DeclarDesc) To UBound(DeclarDesc)
      If LenB(strOnForm) = 0 Or strOnForm = DeclarDesc(I).DDComp Then
        If LenB(strClass) = 0 Or strClass = DeclarDesc(I).DDType Then
          If LenB(StrScope) = 0 Or StrScope = DeclarDesc(I).DDScope Or strScopePlus = DeclarDesc(I).DDScope Then
            If strTest = DeclarDesc(I).DDName Then
              GetDeclarationType = DeclarDesc(I).DDType
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function GetEndOfRoutine(ByVal ArrayTest As Variant) As Long

  'Just in case there are comments or rubbish added to code by processing
  'this finds real end of routine( assuming the code is damaged)

  GetEndOfRoutine = UBound(ArrayTest)
  If Not MultiLeft(ArrayTest(GetEndOfRoutine), True, "End Sub", "End Property", "End Function") Then
    Do
      GetEndOfRoutine = GetEndOfRoutine - 1
      If GetEndOfRoutine < 0 Then
        'the 'routine' is a declarations section with no routines to find
        GetEndOfRoutine = UBound(ArrayTest)
        Exit Do
      End If
    Loop While Not MultiLeft(ArrayTest(GetEndOfRoutine), True, "End Sub", "End Property", "End Function")
  End If

End Function

Public Function GetProcCodeLineOfRoutine(arrR As Variant, _
                                         Optional ByVal bCorretForlineContinuation As Boolean = False) As Long

  Dim I As Long

  For I = LBound(arrR) To UBound(arrR)
    If Not JustACommentOrBlank(arrR(I)) Then
      If isProcHead(arrR(I)) Then
        If bCorretForlineContinuation Then
          GetProcCodeLineOfRoutine = I
          If HasLineCont(arrR(I)) Then
            GetProcCodeLineOfRoutine = GetSafeInsertLineArray(arrR, I)
          End If
          Exit For
         Else
          GetProcCodeLineOfRoutine = I
          Exit For
        End If
      End If
    End If
  Next I

End Function

Private Function GetUserFunctionType(varTest As Variant) As String

  Dim I As Long

  If bProcDescExists Then
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      If PRocDesc(I).PrDName = varTest Then
        If PRocDesc(I).PrDClass = "Function" Then
          GetUserFunctionType = PRocDesc(I).PrDType
          Exit For
        End If
      End If
    Next I
  End If

End Function

Public Function InternalEvidence(varTest As Variant) As String

  If InStr(varTest, "Join") And InStr(varTest, "Split") Then
    If InStr(varTest, "Join") < InStr(varTest, "Split") Then
      InternalEvidence = " As String"
     Else
      InternalEvidence = " As Variant"
    End If
   ElseIf IsUserFunction(varTest) Then
    If Len(GetUserFunctionType(varTest)) Then
      InternalEvidence = " As " & GetUserFunctionType(varTest)
    End If
   ElseIf isRefLibKnownVBConstant(varTest) Then
    If Len(GetRefLibKnownVBConstantType(varTest)) Then
      InternalEvidence = " As " & GetRefLibKnownVBConstantType(varTest)
    End If
   ElseIf InstrArrayHard(varTest, arrInternalTest1) Then
    InternalEvidence = " As String"
   ElseIf MultiRight(varTest, True, ".Filename", ".Name", ".SubItems", ".Path", "&") Then
    InternalEvidence = " As String"
   ElseIf MultiLeft(varTest, True, "&H", "InStr", "InStrRev", "Len") Then
    InternalEvidence = " As Long"
   ElseIf MultiLeft(varTest, True, ".ScaleHeight", ".ScaleWidth", ".Height", ".Width", ".ForeColor", ".BackColor", "+") Then
    InternalEvidence = " As Long"
   ElseIf InstrArrayHard(varTest, arrInternalTest2) Then
    InternalEvidence = " As VbMsgBoxResult"
   ElseIf InstrArrayHard(varTest, arrInternalTest3) Then
    InternalEvidence = " As Double"
   ElseIf InstrArrayHard(varTest, arrPatternDetector2) Then
    InternalEvidence = " As Single"
   ElseIf InstrArrayWholeWord(varTest, "True", "False") Then
    InternalEvidence = " As Boolean"
   ElseIf InstrArrayWholeWord(varTest, "Array", "Split") Then
    InternalEvidence = " As Variant"
   ElseIf InstrArrayWholeWord(varTest, "Int", "Fix", "FreeFile") Then
    InternalEvidence = " As Integer"
   ElseIf Right$(varTest, 10) = ".ListIndex" Then
    InternalEvidence = " As Integer"
   ElseIf InstrArrayWholeWord(varTest, "Join", "CreateObject") Then
    InternalEvidence = " As Variant"
   ElseIf IsNumeric(varTest) Then
    InternalEvidence = " As " & LowestTypeFit(varTest)
    'Byte is extremely unlikely choice (If some one is using it they tend to set it) so reset
    If InternalEvidence = " As Byte" Then
      InternalEvidence = " As Integer"
    End If
   ElseIf isRefLibVBCommands(varTest) Then
    If Len(GetRefLibVBCommandsReturnType(varTest)) Then
      InternalEvidence = " As " & GetRefLibVBCommandsReturnType(varTest)
    End If
   ElseIf Right$(varTest, 1) = "$" Then
    If isRefLibVBCommands(Left$(varTest, Len(varTest) - 1)) Then
      If Len(GetRefLibVBCommandsReturnType(Left$(varTest, Len(varTest) - 1))) Then
        InternalEvidence = " As " & GetRefLibVBCommandsReturnType(Left$(varTest, Len(varTest) - 1))
      End If
    End If
   ElseIf IsDeclareName(varTest) Then 'IsDeclaration(varTest) Then
    'v2.3.4 extra test
    If LenB(GetDeclarationType(varTest)) Then
      InternalEvidence = " As " & GetDeclarationType(varTest)
    End If
    'InternalEvidence=
   Else
    InternalEvidence = vbNullString
  End If

End Function

Public Function InternalEvidenceType(varTest As Variant, _
                                     ByVal CompName As String) As String

  Dim strTest As String
  Dim strTmp  As String

  strTest = varTest
  InternalEvidenceType = InternalEvidence(strTest)
  If Not LenB(InternalEvidenceType) Then
    If IsDeclaration(strTest, "Private") Then
      strTmp = GetDeclarationType(strTest, "Private", , CompName)
      If LenB(strTmp) Then
        InternalEvidenceType = " As " & strTmp
      End If
     ElseIf IsDeclaration(strTest, "Public") Then
      strTmp = GetDeclarationType(strTest)
      If LenB(strTmp) Then
        InternalEvidenceType = " As " & strTmp
      End If
     Else
      If LenB(LowestTypeFit(varTest)) Then
        InternalEvidenceType = " As " & LowestTypeFit(varTest)
      End If
    End If
  End If

End Function

Private Function InTypeRange(ByVal TValue As Variant, _
                             ByVal TRange As TypeRange) As Boolean

  'Copyright 2003 Roger Gilchrist
  'email: rojagilkrist@hotmail.com
  'Returns True if tValue is within the TypeRange asked

  On Error GoTo OutOfRange
  Select Case TRange
   Case TRBool
    InTypeRange = (TValue = CBool(TValue))
   Case TRByte
    InTypeRange = (TValue = CByte(TValue))
   Case TRCur
    InTypeRange = (TValue = CCur(TValue))
   Case TRDec
    InTypeRange = (TValue = CDec(TValue))
   Case TRInt
    InTypeRange = (TValue = CInt(TValue))
   Case TRLng
    InTypeRange = (TValue = CLng(TValue))
   Case TRSng
    InTypeRange = (TValue = CSng(TValue))
   Case TRDbl
    InTypeRange = (TValue = CDbl(TValue))
   Case TRDate
    InTypeRange = IsDate(TValue)
   Case TRStr
    InTypeRange = (TValue = CStr(TValue))
   Case TRVar
    InTypeRange = (TValue = CVar(TValue))
  End Select
  On Error GoTo 0

Exit Function

OutOfRange:
  InTypeRange = False

End Function

Private Function IsFunctionCall(varTest As Variant, _
                                ByVal CompName As String) As Boolean

  If IsProcedure(CStr(varTest), "Public", "Function") Then
    IsFunctionCall = True
   ElseIf IsProcedure(CStr(varTest), "Private", "Function", CompName) Then
    IsFunctionCall = True
  End If

End Function

Public Function IsGotoLabel(ByVal varTest As Variant, _
                            ByVal CompName As String, _
                            Optional strFix As String, _
                            Optional arrGoTo As Variant) As Boolean

  Dim strPoorNameType As String

  'Update ver 1.0.87
  'misrecognized if the original line was of format
  '*If Y = 3 then VarTest: x = 3 Else sub_call2: x = 4
  'where VarTest is also a sub call
  strFix = vbNullString
  'v2.4.4 improved support for GoTo labels on same line as other code
  If Len(varTest) Then
    If InStr(varTest, SngSpace) = 0 Then
      If Right$(Trim$(varTest), 1) = Colon Then
        'v2.8.3 improved detector for single word reserved words
        If Not isVBReservedWord(Replace$(varTest, ":", "")) Then
          If InQSortArray(arrGoTo, IIf(InStr(varTest, ":"), Left$(varTest, Len(varTest) - 1), varTest)) Then
            IsGotoLabel = True ' this is a called goto target so ok
           Else
            If GetSpacePos(Trim$(varTest)) = 0 Then
              If ExtractCode(varTest) Then
                If InStr(varTest, ".") = 0 Then
                  'v2.8.3 improved test
                  IsGotoLabel = isVBReservedWord(Replace$(varTest, ":", vbNullString))
                  If MultiLeft(varTest, True, "Do:", "Else:", "Loop:", "Next:") Then
                    strFix = Replace$(varTest, ":", vbNullString)
                    IsGotoLabel = False
                   Else
                    If PoorNameGoto(varTest, CompName, strPoorNameType) Then
                      strFix = varTest & vbNewLine & _
                       WARNING_MSG & "This Goto Target label is named with a " & strPoorNameType & "." & vbNewLine & _
                       RGSignature & "Legal but makes code hard to read."
                      IsGotoLabel = Right$(Trim$(varTest), 1) = Colon
                    End If
                    ' End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  End If

End Function

Private Function IsUserFunction(varTest As Variant) As Boolean

  Dim I As Long

  If bProcDescExists Then
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      If PRocDesc(I).PrDName = varTest Then
        If PRocDesc(I).PrDClass = "Function" Then
          IsUserFunction = True
          Exit For
        End If
      End If
    Next I
  End If

End Function

Public Function LowestTypeFit(VValue As Variant) As String

  'Copyright 2003 Roger Gilchrist
  'email: rojagilkrist@hotmail.com
  'Returns the smallest (in term of memory usage) Type which can hold Value

  If InTypeRange(VValue, TRByte) Then
    LowestTypeFit = "Byte"
   ElseIf InTypeRange(VValue, TRInt) Then
    LowestTypeFit = "Integer"
   ElseIf InTypeRange(VValue, TRLng) Then
    LowestTypeFit = "Long"
   ElseIf InTypeRange(VValue, TRSng) Then
    LowestTypeFit = "Single"
   ElseIf InTypeRange(VValue, TRDbl) Then
    LowestTypeFit = "Double"
   ElseIf InTypeRange(VValue, TRCur) Then
    LowestTypeFit = "Currancy"
   ElseIf InTypeRange(VValue, TRDec) Then
    LowestTypeFit = "Variant"
   ElseIf InTypeRange(VValue, TRBool) Then
    LowestTypeFit = "Boolean"
   Else
    If IsDate(VValue) Then
      LowestTypeFit = "Date"
    End If
  End If
  'ver 1.0.90
  'following edited out as they can cause nasty errors
  'ElseIf value = CStr(value) Then
  'LowestTypeFit = "String"
  'Else
  'LowestTypeFit = "Variant"
  '   Else
  'do nothing

End Function

Public Function Marker(varCode As Variant, _
                       Msg As String, _
                       MPos As MarkerPos, _
                       Optional bUpdate As Boolean) As String

  'ver 2.0.5 now automatically sets the Updated triggers
  'THis is a simple way of attaching comments to immediate code line
  'see also SmartMarker which is designed for message to be separated from trigger line
  '*If there are newline characters in the msg then they are replaced with Newline & RGSignature

  Msg = CleanMsg(Msg)
  bUpdate = Len(Msg)
  Select Case MPos
   Case MBefore
    Marker = Msg & vbNewLine & varCode
   Case MAfter
    Marker = varCode & vbNewLine & Msg
   Case MEoL
    Marker = varCode & Msg
   Case MSoL
    Marker = Msg & varCode
   Case MEmbed
    Marker = Replace$(Msg, EmbedMsg, varCode)
  End Select

End Function

Public Function PoorNameGoto(ByVal varTest As Variant, _
                             ByVal CompName As String, _
                             strType As String) As Boolean

  'GotoLabel is also Sub/Function or vb Command or Reserved Word

  If Right$(varTest, 1) = ":" Then
    varTest = Left$(varTest, Len(varTest) - 1)
  End If
  If IsProcedure(varTest, , "Sub") Then
    PoorNameGoto = True
    strType = " Procedure name"
   ElseIf IsFunctionCall(varTest, CompName) Then
    PoorNameGoto = True
    strType = " Procedure name"
   ElseIf isRefLibVBCommands(varTest) Then
    PoorNameGoto = True
    strType = "VB Command name"
   ElseIf InQSortArray(ArrQVBReservedWords, varTest) Then
    strType = "VB Reserved word"
    PoorNameGoto = True
  End If

End Function

Public Function RightWord(ByVal varChop As Variant) As String

  Dim arrTmp As Variant

  If LenB(varChop) Then
    arrTmp = Split(varChop)
    RightWord = arrTmp(UBound(arrTmp))
  End If

End Function

Public Sub SafeInsertArray(Arr As Variant, _
                           Lcontcount As Long, _
                           ByVal Msg As String)

  Lcontcount = GetSafeInsertLineArray(Arr, Lcontcount)
  Arr(Lcontcount) = Arr(Lcontcount) & Msg

End Sub

Public Sub SafeInsertArrayMarker(Arr As Variant, _
                                 Lcontcount As Long, _
                                 ByVal Msg As String)

  Lcontcount = GetSafeInsertLineArray(Arr, Lcontcount)
  Arr(Lcontcount) = Marker(Arr(Lcontcount), Msg, MAfter)

End Sub

Public Function SearchRoutineArray2(ByVal ArrProc As Variant, _
                                    ByVal varFind As Variant) As Boolean

  Dim I       As Long

  For I = LBound(ArrProc) To UBound(ArrProc)
    If InstrAtPosition(ArrProc(I), varFind, ipAny, True) Then
      SearchRoutineArray2 = True
      Exit For
    End If
  Next I

End Function

Public Sub TopCommentsIntoRoutine(cMod As CodeModule)

  Dim MyStr             As String
  Dim UpDated           As Boolean
  Dim arrMembers        As Variant
  Dim ArrRoutine        As Variant
  Dim I                 As Long
  Dim J                 As Long
  Dim MaxFactor         As Long
  Dim OptionExplicitPos As Long
  Dim Jump1             As Boolean
  Dim ModuleNumber      As Long

  'Move any comments outside routine into routine
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  ModuleNumber = ModDescMember(cMod.Parent.Name)
  If dofix(ModuleNumber, MoveCommentInside) Then
    OptionExplicitPos = -1
    arrMembers = GetMembersArray(cMod)
    MaxFactor = UBound(arrMembers)
    If MaxFactor > -1 Then
      For I = 1 To MaxFactor
        MemberMessage GetProcNameStr(arrMembers(I)), I, MaxFactor
        ArrRoutine = Split(arrMembers(I), vbNewLine)
        If UBound(ArrRoutine) > -1 Then
          For J = LBound(ArrRoutine) To UBound(ArrRoutine)
            If Left$(ArrRoutine(J), 15) = "Option Explicit" Then
              OptionExplicitPos = J
            End If
            If JustACommentOrBlank(ArrRoutine(J)) Then
              If ArrRoutine(J) <> RGSignature Then
                MyStr = MyStr & vbNewLine & ArrRoutine(J)
              End If
              ArrRoutine(J) = vbNullString
             Else
              J = J + 1
              If LenB(MyStr) Then
                If I = 0 Then
                  If ArrRoutine(0) <> "Option Explicit" Then
                    If cMod.CountOfDeclarationLines = 1 Then
                      'only Option Explicit in Declaration so move it to end of Declaration comments
                      If OptionExplicitPos > -1 Then
                        ArrRoutine(OptionExplicitPos) = vbNullString
                        ArrRoutine(J) = MyStr & vbNewLine & "Option Explicit" & vbNewLine & ArrRoutine(J)
                       Else
                        If HasLineCont(ArrRoutine(J)) Then
                          J = GetSafeInsertLineArray(ArrRoutine, J)
                          Jump1 = True
                          'J = J + 1
                          ArrRoutine(J) = MyStr & vbNewLine & ArrRoutine(J)
                        End If
                      End If
                     Else
                      J = GetSafeInsertLineArray(ArrRoutine, J)
                      If Jump1 Then
                        J = J + 1
                      End If
                      ArrRoutine(J) = MyStr & vbNewLine & ArrRoutine(J)
                    End If
                  End If
                 ElseIf LenB(MyStr) Then
                  If MyStr <> vbNewLine Then
                    SafeInsertArray ArrRoutine, J, MyStr
                  End If
                End If
                UpDated = True
                MyStr = vbNullString
                arrMembers(I) = Join(CleanArray(ArrRoutine), vbNewLine)
              End If
              Exit For
            End If
          Next J
          arrMembers(I) = Join(CleanArray(ArrRoutine), vbNewLine)
          UpDated = True
        End If
      Next I
      ReWriteMembers cMod, arrMembers, UpDated
    End If
  End If

End Sub

Public Sub UpdateMember(arrMemb As Variant, _
                        ArrProc As Variant, _
                        UpdateME As Boolean, _
                        UpdateMod As Boolean)

  If UpdateME Then
    arrMemb = Join(ArrProc, vbNewLine)
    UpdateME = False
    UpdateMod = True
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:21:29 AM) 18 + 524 = 542 Lines Thanks Ulli for inspiration and lots of code.

