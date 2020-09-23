Attribute VB_Name = "mod_Declarations2"
Option Explicit
Public Enum CtrlDescNo
  e_CDName
  e_CDIndex
  e_CDClass
  e_CDform
  e_CDCaption
  e_CDUsage
  e_CDProj
  e_CDXPFrameBug
  e_CDBadType
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private e_CDName, e_CDIndex, e_CDClass, e_CDform, e_CDCaption, e_CDUsage, e_CDProj, e_CDXPFrameBug, e_CDBadType
#End If
Private Const LVM_FIRST                 As Long = &H1000
Private Const LVM_GETCOUNTPERPAGE       As Double = LVM_FIRST + 40
Public Const LVM_SETCOLUMNWIDTH         As Long = LVM_FIRST + 30

Public Sub Assign_Type_To_Constants(ByVal ModuleNumber As Long, _
                                    dArray As Variant)

  Dim OrigLine     As String
  Dim L_CodeLine   As String
  Dim SpaceOffSet  As String
  Dim CommentStore As String
  Dim TmpDecArr    As Variant
  Dim UpDated      As Boolean
  Dim I            As Long
  Dim MaxFactor    As Long

  'Prevents Constants using unnecessary Variant memory space
  If dofix(ModuleNumber, UpdateConstExplicit) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > -1 Then
      For I = 0 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If InstrAtPosition(L_CodeLine, "Const", ipLeftOr2nd) Then
            If ExtractCode(L_CodeLine, CommentStore, SpaceOffSet) Then
              If Not Has_AS(L_CodeLine) Then
                L_CodeLine = ConcealParameterCommas(L_CodeLine, True)
                L_CodeLine = SpaceOffSet & L_CodeLine & CommentStore
                OrigLine = L_CodeLine
                L_CodeLine = ConstantExpander(L_CodeLine, IIf(ModuleNumber = -1, True, False))
                If OrigLine <> L_CodeLine Then
                  TmpDecArr(I) = L_CodeLine
                  UpDated = True
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

Public Function ConstantExpander(ByVal varName As String, _
                                 ByVal SupressComment As Boolean) As String

  Dim TypeName     As String
  Dim MyStr        As String
  Dim CommentStore As String
  Dim LngMode      As Long
  Dim ConstName    As String
  Dim ConstValue   As String
  Dim ShowComment  As Boolean
  Dim LDoFix       As Boolean
  Dim CFComment    As String
  Dim MsgHead      As String
  Dim strTmp       As String
  Dim strComOnly   As String

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Constants often contain clues to their type so they can be analysed to get a Type
  'Rewrite un-Typed Const Definitions in following circumstances
  ' Hex numbers to use As Long
  ' Value is enclosed in brackets to use As Long
  ' string value to use As String
  ' True or False values to Boolean
  '  If there are math operations in value , set to Double
  '*If the Const Name is all caps with undescore _ (and optionally numerals) it is probably an API parameter and they are usually Long
  ' numerals to smallest Type that will hold it(Byte is set to Integer because poeple rarely use Byte with out setting it.
  MsgHead = UPDATED_MSG & "UnTyped Const with"
  LngMode = FixData(UpdateConstExplicit).FixLevel
  LDoFix = IIf(LngMode >= FixAndComment, True, False)
  ShowComment = IIf(LngMode = FixAndComment, True, False)
  strComOnly = IIf(LngMode = CommentOnly, "could be", SngSpace)
  On Error GoTo BadError
  MyStr = varName
  MyStr = Trim$(MyStr)
  If ExtractCode(MyStr, CommentStore) Then
    If CountCodeSubString(MyStr, "=") = 1 Then
      ConstValue = Trim$(Mid$(MyStr, InStr(MyStr, "=") + 1))
      ConstName = WordBefore(MyStr, "=")
      strTmp = TypeSuffix2String(ConstName)
      If LenB(strTmp) Then
        If FixData(UpdateDecTypeSuffix).FixLevel > Off Or LDoFix Then
          MyStr = Safe_Replace(MyStr, ConstName, Left$(ConstName, Len(ConstName) - 1))
          MyStr = Safe_Replace(MyStr, EqualInCode, " As " & strTmp & EqualInCode)
        End If
        CFComment = IIf(ShowComment, UPDATED_MSG & "Const with Type suffix " & Right$(ConstName, 1) & strComOnly & "updated to use As " & strTmp, vbNullString)
       ElseIf InStr(ConstValue, DQuote) Then
        'String character " in Value
        ConstantExpanderApply MyStr, "String", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " string value" & strComOnly & "changed to As String", vbNullString)
       ElseIf TypeSuffixExists(ConstValue) Then
        If LDoFix Then
          ConstantExpanderApply MyStr, TypeUpdate(Right$(ConstValue, 1), , False), LDoFix
          MyStr = Safe_Replace(MyStr, " " & ConstValue, " " & Left$(ConstValue, Len(ConstValue) - 1))
        End If
        CFComment = IIf(ShowComment, MsgHead & " value ending in " & strInSQuotes(Right$(ConstValue, 1)) & strComOnly & "changed to" & TypeUpdate(Right$(ConstValue, 1)) & " ", vbNullString)
       ElseIf InStr(ConstValue, "&H") Then
        '&H in value
        ConstantExpanderApply MyStr, "Long", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " Hex (&H) value" & strComOnly & "changed to As Long", vbNullString)
       ElseIf InstrAtPositionArray(ConstValue, ipAny, True, "True", "False", "<", ">") Then
        'Boolean values in value
        ConstantExpanderApply MyStr, "Boolean", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " value using 'True', 'False', '<' or '>'" & strComOnly & "changed to As Boolean", vbNullString)
       ElseIf InstrAtPositionArray(ConstValue, ipAny, True, "*", "/", "+", "-", "\", "Or", "And", "Xor") Then
        'Math performed in value
        ConstantExpanderApply MyStr, "Double", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " value using Math or Logic operation" & strComOnly & "changed to As Double", vbNullString)
       ElseIf IsNumeric(ConstValue) Then
        'Const value is numeric so set it to the smallest Type able to hold the value
        TypeName = LowestTypeFit(CVar(ConstValue))
        If LDoFix Then
          If TypeName = "Byte" Then
            TypeName = "Integer"
          End If
          ConstantExpanderApply MyStr, TypeName, LDoFix
        End If
        CFComment = IIf(ShowComment, MsgHead & " numeric value" & strComOnly & "changed to As " & TypeName, vbNullString)
       ElseIf ConstNameIsAPIStyle(ConstName) Or (ConstName = UCase$(ConstName)) Then
        'Const name has form AAA_BBBB123 probably an API and they are usually Long
        'OR the name is all capitals
        ConstantExpanderApply MyStr, "Long", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " API sytle Name" & strComOnly & "changed set to As Long", vbNullString)
       ElseIf MultiRight(LCase$(ConstName), False, "color", "colour") Then
        ConstantExpanderApply MyStr, "Long", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " Name containing 'color' or 'colour'" & strComOnly & "changed to As Long", vbNullString)
       ElseIf CountSubString(ConstName, "_") Then
        ConstantExpanderApply MyStr, "Long", LDoFix
        CFComment = IIf(ShowComment, MsgHead & " Name containing underscores '_'" & strComOnly & "changed to As Long", vbNullString)
       ElseIf isRefLibKnownVBConstant(ConstValue) Then
        If Len(GetRefLibKnownVBConstantType(ConstValue)) Then
          ConstantExpanderApply MyStr, GetRefLibKnownVBConstantType(ConstValue), LDoFix
          CFComment = IIf(ShowComment, MsgHead & " known Constant value" & strComOnly & "changed to As Long", vbNullString)
        End If
      End If
    End If
    If SupressComment Then
      CFComment = vbNullString
    End If
    ConstantExpander = Marker(MyStr & CommentStore, CFComment, IIf(HasLineCont(MyStr), MBefore, MAfter))
    If ConstantExpander <> varName Then
      AddNfix UpdateConstExplicit
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  ConstantExpander = varName

End Function

Private Sub ConstantExpanderApply(strCode As String, _
                                  ByVal StrReplace As String, _
                                  ByVal DoIt As Boolean)

  If DoIt Then
    strCode = Safe_Replace(strCode, EqualInCode, " As " & StrReplace & EqualInCode)
  End If

End Sub

Private Function ConstNameIsAPIStyle(ByVal strTest As String) As Boolean

  Dim I As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Test if a constant name has the API style CAPITALS and underscores (optionally includes numbers)
  If InStr(strTest, "_") Then
    strTest = Safe_Replace(strTest, "_", vbNullString)
    For I = 0 To 9
      strTest = Safe_Replace(strTest, I, vbNullString)
    Next I
    If LenB(strTest) Then
      If strTest = UCase$(strTest) Then
        ConstNameIsAPIStyle = True
      End If
    End If
  End If

End Function

Public Function CountCodeSubString(varSearch As Variant, _
                                   varFind As Variant) As Long

  Dim I As Long

  I = InStr(varSearch, varFind)
  Do While I
    If InCode(varSearch, I) Then
      CountCodeSubString = CountCodeSubString + 1
    End If
    I = InStr(I + 1, varSearch, varFind)
  Loop

End Function

Private Sub Declaration_PrivateTypeLister(ByVal ModuleNumber As Long, _
                                          dArray As Variant)

  Dim arrLine    As Variant
  Dim L_CodeLine As String
  Dim strTemp    As String
  Dim TmpDecArr  As Variant
  Dim I          As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Generate an array of private Types
  TmpDecArr = CleanArray(dArray)
  For I = LBound(dArray) To UBound(dArray) - 1
    L_CodeLine = TmpDecArr(I)
    If Not JustACommentOrBlank(L_CodeLine) Then
      arrLine = Split(L_CodeLine)
      Select Case arrLine(0)
       Case "Private"
        Select Case arrLine(1)
         Case "Type"
          strTemp = AccumulatorString(strTemp, arrLine(2))
        End Select
      End Select
    End If
  Next I
  If LenB(strTemp) Then
    PrivateTypeArray(ModuleNumber) = QuickSortArray(Split(strTemp, ","))
   Else
    PrivateTypeArray(ModuleNumber) = Split("")
  End If

End Sub

Public Sub DeclareDimGlobalUpdate(ByVal ModuleNumber As Long, _
                                  Comp As VBComponent, _
                                  DecArray As Variant)

  Dim strFound  As String
  Dim StartLine As Long
  Dim Hit       As Boolean

  'v2.7.2 use find to update dim/global
  On Error Resume Next
  If dofix(ModuleNumber, DimGlobal2PublicPrivate) Then
    StartLine = 1
    With Comp
      Do While .CodeModule.Find("Dim", StartLine, 1, -1, -1, True, True)
        If StartLine <= .CodeModule.CountOfDeclarationLines Then
          strFound = .CodeModule.Lines(StartLine, 1)
          If LeftWord(strFound) = "Dim" Then
            InsertNewCodeComment .CodeModule, StartLine, 1, Replace$(strFound, "Dim ", "Private ", , 1), UPDATED_MSG & "Module Level 'Dim' to 'Private'"
            '.CodeModule.ReplaceLine StartLine, Replace$(strFound, "Dim ", "Private ", , 1)
            'SafeInsertModule .CodeModule, StartLine, UPDATED_MSG & "Module Level 'Dim' to 'Private'"
            Hit = True
          End If
          StartLine = StartLine + 1
          If StartLine > .CodeModule.CountOfDeclarationLines Then
            Exit Do
          End If
         Else
          Exit Do
        End If
      Loop
      StartLine = 1
      'v2.7.6 fixed now works
      Do While .CodeModule.Find("Global", StartLine, 1, -1, -1, True, True)
        'v3.0.4 left out the = sign
        If StartLine <= .CodeModule.CountOfDeclarationLines Then
          strFound = .CodeModule.Lines(StartLine, 1)
          If LeftWord(strFound) = "Global" Then
            InsertNewCodeComment .CodeModule, StartLine, 1, Replace$(strFound, "Global ", "Public ", , 1), UPDATED_MSG & "Module Level 'Global' to 'Public'"
            '    .CodeModule.ReplaceLine StartLine, Replace$(strFound, "Global ", "Public ", , 1)
            '   SafeInsertModule .CodeModule, StartLine, UPDATED_MSG & "Module Level 'Global' to 'Public'"
            Hit = True
          End If
          StartLine = StartLine + 1
          If StartLine > .CodeModule.CountOfDeclarationLines Then
            Exit Do
          End If
         Else
          Exit Do
        End If
      Loop
    End With 'Comp
    If Hit Then
      ' only re-get if it has been rewritten
      DecArray = GetDeclarationArray(Comp.CodeModule)
    End If
  End If
  On Error GoTo 0

End Sub

Private Function DOPartialFind(ByVal strPFind As String) As Boolean

  Dim Proj As VBProject
  Dim Comp As VBComponent

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If Comp.CodeModule.Find(strPFind, 1, 1, -1, -1, False, False, True) Then
          DOPartialFind = True
          Exit For
        End If
      End If
    Next Comp
    If DOPartialFind Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Public Sub Fix_Form_Public_Declarations(ByVal ModuleNumber As Long, _
                                        dArray As Variant, _
                                        Comp As VBComponent)

  
  Dim L_CodeLine      As String
  Dim arrLine         As Variant
  Dim TmpDecArr       As Variant
  Dim I               As Long
  Dim UpDated         As Boolean
  Dim MaxFactor       As Long
  Dim TestWord        As String
  Dim bSuppressMsg    As Boolean
  Dim MUpdated        As Boolean
  Dim SuggestProperty As Boolean
  Dim arrTest1        As Variant
  Dim arrTest2        As Variant

  arrTest1 = Array("Enum", "Type")
  arrTest2 = Array("Event", "WithEvents")
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Automatically convert declarations which are too widely scoped to Private
  If dofix(ModuleNumber, Public2Private) Then
    If LenB(GetActiveProject.FileName) Then
      If Comp.Type <> vbext_ct_ClassModule Then
        If Comp.Type <> vbext_ct_StdModule Then
          If Comp.Type <> vbext_ct_UserControl Then
            'classes may contain Public variables which are used as properties
            'to be referenced by new instances of the class so assume the author got it right
            Declaration_PrivateTypeLister ModuleNumber, dArray
            TmpDecArr = CleanArray(dArray)
            MaxFactor = UBound(TmpDecArr)
            If MaxFactor > -1 Then
              For I = 0 To MaxFactor
                bSuppressMsg = False
                SuggestProperty = False
                MemberMessage "", I, MaxFactor
                If Not JustACommentOrBlank(TmpDecArr(I)) Then
                  L_CodeLine = StripDoubleSpace(Trim$(TmpDecArr(I)))
                  If InstrAtPosition(L_CodeLine, "Public", IpLeft) Then
                    arrLine = Split(L_CodeLine)
                    TestWord = GetName(arrLine, True)
                    If LenB(arrLine(0)) Then
                      'check that a guessed Public is not scoped too large
                      If Not IsInArray(arrLine(1), arrTest1) Then
                        If Not IsInArray(arrLine(1), arrTest2) Then
                          ' all Events are public
                          If Not FindCodeUsage(TestWord, vbNullString, Comp.Name, False, False, True) Then
                            '*if not used outside this component then should be Private
                            arrLine(0) = "Private"
                            UpDated = True
                          End If
                          If arrLine(0) = "Public" Then
                            If IsComponent_ControlHolder(Comp) Or IsComponent_ClassMod(Comp.Type) Then
                              If FindCodeUsage(Comp.Name & "." & TestWord, L_CodeLine, Comp.Name, False, False, True) Or FindCodeUsage("." & TestWord, L_CodeLine, Comp.Name, False, False, True) Then
                                UpDated = True
                                SuggestProperty = True
                               Else
                                If FindCodeUsage(TestWord, L_CodeLine, Comp.Name, False, False, True) And InStr(L_CodeLine, "As Collection") Then
                                  UpDated = True
                                  SuggestProperty = True
                                 Else
                                  If Not ProtectConnectVBIDE(L_CodeLine) Then
                                    arrLine(0) = "Private"
                                    UpDated = True
                                  End If
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                      If IsInArray(arrLine(1), arrTest1) Then
                        ' all Events are public
                        If Not FindCodeUsage(TestWord, vbNullString, Comp.Name, False, False, True) Then
                          '*if not used outside this component then should be Private unless it is being used to type cast a procedure or parameter
                          If FindCodeUsage(" As " & TestWord, vbNullString, Comp.Name, False, False, False) Then
                            If arrLine(0) <> "Public" Then
                              arrLine(0) = "Public"
                              UpDated = True
                            End If
                          End If
                        End If
                      End If
                      If UpDated Then
                        If Not bSuppressMsg Then
                          Select Case FixData(Public2Private).FixLevel
                           Case CommentOnly
                            If SuggestProperty Then
                              If FixData(PublicVar2Property).FixLevel = Off Then
                                TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                                 SUGGESTION_MSG & "Public variables on Forms should be converted to properties"
                              End If
                             Else
                              TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                               SUGGESTION_MSG & "Scope should be changed to " & arrLine(0)
                            End If
                           Case FixAndComment
                            If SuggestProperty Then
                              If FixData(PublicVar2Property).FixLevel <> FixAndComment Then
                                TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                                 SUGGESTION_MSG & "Public variables on Forms should be converted to properties"
                              End If
                             Else
                              TmpDecArr(I) = Join(arrLine) & vbNewLine & WARNING_MSG & "Scope Changed to " & arrLine(0)
                            End If
                           Case JustFix
                            If SuggestProperty Then
                              If FixData(PublicVar2Property).FixLevel <> FixAndComment Then
                                TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                                 SUGGESTION_MSG & "Public variables on Forms should be converted to properties"
                              End If
                             Else
                              TmpDecArr(I) = Join(arrLine)
                            End If
                          End Select
                        End If
                        AddNfix Public2Private
                        UpDated = False
                        MUpdated = True
                      End If
                    End If
                  End If
                End If
              Next I
              dArray = CleanArray(TmpDecArr, MUpdated)
            End If
          End If
        End If
      End If
    End If
  End If

End Sub

Public Function GenerateDefTypeTargetArray(arrT As Variant) As Variant

  Dim strTmp As String
  Dim J      As Long
  Dim K      As Long

  'v2.8.3 improved to cope with A-C style DefXXX
  ' Thanks Joakim Schramm for leading me to this
  ' the old code was not working properly
  For J = 1 To UBound(arrT)
    arrT(J) = Safe_Replace(arrT(J), ",", vbNullString)
    If Len(arrT(J)) > 1 Then
      arrT(J) = vbNullString
    End If
    If InStr(arrT(J), "-") Then
      For K = Asc(Left$(arrT(J - 1), 1)) To Asc(Right$(arrT(J + 1), 1))
        strTmp = AccumulatorString(strTmp, UCase$(Chr$(K)))
      Next K
      arrT(J) = strTmp
      strTmp = vbNullString
    End If
  Next J
  GenerateDefTypeTargetArray = StripDuplicateArray(Split(Join(arrT, ","), ","))

End Function

Public Function GetName(varA As Variant, _
                        Optional ByVal PublicTest As Boolean = False) As String

  Dim DoTest As Boolean

  If PublicTest Then
    DoTest = Left$(varA(0), 6) = "Public"
   Else
    DoTest = MultiLeft(varA(0), True, "Dim", "Global", "Public", "Private")
  End If
  If DoTest Then
    'Declared with older styles
    Select Case varA(1)
     Case "Type", "Const", "Enum", "WithEvents", "Event"
      GetName = varA(2)
     Case "Declare"
      GetName = varA(3)
     Case Else
      GetName = varA(1)
    End Select
   ElseIf MultiLeft(varA(0), True, "Type", "Const", "Enum", "WithEvents", "Event", "Declare") Then
    'No explicit declare
    Select Case varA(0)
     Case "Declare"
      GetName = varA(2)
     Case Else
      If UBound(varA) > 0 Then
        GetName = varA(1)
       Else
        GetName = varA(0)
      End If
    End Select
  End If
  If InStr(GetName, LBracket) Then
    GetName = strGetLeftOf(GetName, LBracket)
  End If

End Function

Public Function InStructure(Struct As EnumStruct, _
                            ArrRoutine As Variant, _
                            ByVal TLine As Long, _
                            StartPos As Long, _
                            EndPos As Long) As Boolean

  Dim StructStart As Variant
  Dim StructEnd   As Variant
  Dim I           As Long
  Dim TPos        As InstrLocations
  Dim Hit         As Long
  Dim SoR         As Long
  Dim EoR         As Long

  'v2.9.0 extended to Do/Loop
  StartPos = -1
  EndPos = -1
  SoR = GetProcCodeLineOfRoutine(ArrRoutine)
  EoR = GetEndOfRoutine(ArrRoutine)
  StructStart = Array("If", "Select Case", "With", "Type", "Enum", "Do", "For")
  StructEnd = Array("End If", "End Select", "End With", "End Type", "End Enum", "Loop", "Next")
  TPos = IIf(Array(1, 1, 1, 2, 2, 1, 1)(Struct) = 1, IpLeft, ipLeftOr2nd)
  For I = TLine + 1 To SoR Step -1
    If InstrAtPosition(ArrRoutine(I), StructStart(Struct), TPos, True) Then
      Hit = Hit + 1
      StartPos = I
     ElseIf InstrAtPosition(ArrRoutine(I), StructEnd(Struct), TPos, True) Then
      Hit = Hit - 1
    End If
  Next I
  If Hit > 0 Then
    InStructure = True
    For I = TLine To EoR
      If InstrAtPosition(ArrRoutine(I), StructStart(Struct), TPos, True) Then
        Hit = Hit + 1
       ElseIf InstrAtPosition(ArrRoutine(I), StructEnd(Struct), TPos, True) Then
        Hit = Hit - 1
        StartPos = I
      End If
    Next I
  End If
  If Struct = IfStruct Then
    If TypeOfIf(ArrRoutine, TLine, EndPos) <> Simple Then
      For I = TLine To SoR Step -1
        If InstrAtPositionArray(ArrRoutine(I), IpLeft, True, "ElseIf", "Else") Then
          StartPos = I + 1
          Exit For
        End If
      Next I
      For I = TLine To EoR
        If InstrAtPositionArray(ArrRoutine(I), IpLeft, True, "ElseIf", "Else") Then
          EndPos = I - 1
          Exit For
        End If
      Next I
    End If
  End If
  If Struct = SelectStruct Then
    For I = TLine To SoR Step -1
      If InstrAtPosition(ArrRoutine(I), "Case", IpLeft, True) Then
        StartPos = I + 1
        Exit For
      End If
    Next I
    For I = TLine To EoR
      If InstrAtPosition(ArrRoutine(I), "Case", IpLeft, True) Then
        EndPos = I - 1
        Exit For
      End If
    Next I
  End If

End Function

Public Function IsAPI(ByVal strTest As String) As Boolean

  Dim I            As Long

  IsAPI = IsDeclaration(strTest, , "Declare")
  If Not IsAPI Then
    If bProcDescExists Then
      For I = LBound(PRocDesc) To UBound(PRocDesc)
        If strTest = PRocDesc(I).PrDName Then
          IsAPI = PRocDesc(I).PrDAPI
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function IsAPIinLine(strCode As String, _
                            Optional StrAPI As String) As Boolean

  Dim I As Long

  If bDeclExists Then
    For I = LBound(DeclarDesc) To UBound(DeclarDesc)
      If DeclarDesc(I).DDKind = "Declare" Then
        If InStrCode(strCode, DeclarDesc(I).DDName) Then
          IsAPIinLine = True
          StrAPI = DeclarDesc(I).DDName
          Exit For 'unction
        End If
      End If
    Next I
  End If

End Function

Public Function IsComponent_ActiveXDesigner(ByVal c As VBComponent) As Boolean

  IsComponent_ActiveXDesigner = c.Type = vbext_ct_ActiveXDesigner

End Function

Public Function IsComponent_ClassMod(ByVal CompType As Long) As Boolean

  IsComponent_ClassMod = CompType = vbext_ct_ClassModule

End Function

Public Function IsComponent_ControlHolder(ByVal c As VBComponent) As Boolean

  IsComponent_ControlHolder = c.Type = vbext_ct_VBForm Or c.Type = vbext_ct_MSForm Or c.Type = vbext_ct_VBMDIForm Or c.Type = vbext_ct_UserControl Or c.Type = vbext_ct_PropPage Or c.Type = vbext_ct_DocObject

End Function

Public Function IsComponent_Reloadable(ByVal c As VBComponent) As Boolean

  Select Case c.Type
   Case vbext_ct_VBForm, vbext_ct_MSForm, vbext_ct_VBMDIForm, vbext_ct_DocObject, vbext_ct_ClassModule, vbext_ct_StdModule
    IsComponent_Reloadable = True
  End Select

End Function

Public Function IsComponent_User_Class(ByVal c As VBComponent) As Boolean

  IsComponent_User_Class = c.Type = vbext_ct_ClassModule

End Function

Public Function IsComponent_UserControl(ByVal CompType As Long) As Boolean

  IsComponent_UserControl = CompType = vbext_ct_UserControl

End Function

Public Function IsDeclaration(ByVal strTest As String, _
                              Optional ByVal StrScope As String, _
                              Optional ByVal strKind As String, _
                              Optional ByVal strOnForm As String) As Boolean

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
        If LenB(strKind) = 0 Or strKind = DeclarDesc(I).DDKind Then
          If LenB(StrScope) = 0 Or StrScope = DeclarDesc(I).DDScope Or strScopePlus = DeclarDesc(I).DDScope Then
            If strTest = DeclarDesc(I).DDName Then
              IsDeclaration = True
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function IsDeclaredVariable(ByVal strTest As String) As Boolean

  If bDeclExists Then
    If IsDeclaration(strTest, "Public") Then
      IsDeclaredVariable = True
     Else
      IsDeclaredVariable = IsDeclaration(strTest, , "Variable")
    End If
  End If

End Function

Public Function IsDuplicateProcName(ByVal strTest As String, _
                                    Optional ByVal StrScope As String, _
                                    Optional ByVal strProcType As String, _
                                    Optional ByVal strOnForm As String) As Boolean

  Dim I            As Long
  Dim strScopePlus As String

  If bProcDescExists Then
    If StrScope = "Public" Then
      strScopePlus = "Friend"
     ElseIf StrScope = "Private" Then
      strScopePlus = "Static"
    End If
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      If PRocDesc(I).PrDDuplicate Then ' this is a rare condition so test it first
        If LenB(strOnForm) = 0 Or strOnForm = PRocDesc(I).PrdComp Then
          If LenB(strProcType) = 0 Or strProcType = PRocDesc(I).PrDType Then
            If LenB(StrScope) = 0 Or StrScope = PRocDesc(I).PrDScope Or strScopePlus = PRocDesc(I).PrDScope Then
              If strTest = PRocDesc(I).PrDName Then
                If Not PRocDesc(I).PrDControl Then
                  'v2.3.6 added test ignores control procedures
                  IsDuplicateProcName = True
                  Exit For
                End If
              End If
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function isEvent(ByVal strTest As String, _
                        Optional ByVal StrScope As String, _
                        Optional ByVal strOnForm As String) As Boolean

  Dim I            As Long
  Dim strScopePlus As String

  If StrScope = "Public" Then
    strScopePlus = "Friend"
   ElseIf StrScope = "Private" Then
    strScopePlus = "Static"
  End If
  If bEventDescExists Then
    For I = LBound(EventDesc) To UBound(EventDesc)
      If LenB(strOnForm) = 0 Or strOnForm = EventDesc(I).EForm Then
        If LenB(StrScope) = 0 Or StrScope = EventDesc(I).EScope Or strScopePlus = EventDesc(I).EScope Then
          'V3.0.3 Thanks Bazz this allows it to recognize procedures dependant on Events/WithEvents
          If SmartLeft(strTest, EventDesc(I).EName) Then
            isEvent = True
            Exit For
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function IsProcedure(ByVal strTest As String, _
                            Optional ByVal StrScope As String, _
                            Optional ByVal strProcType As String, _
                            Optional ByVal strOnForm As String) As Boolean

  Dim I            As Long
  Dim J            As Long
  Dim strScopePlus As String

  If bProcDescExists Then
    If Not IsNumeric(strTest) Then
      If Not IsPunct(strTest) Then
        If StrScope = "Public" Then
          strScopePlus = "Friend"
         ElseIf StrScope = "Private" Then
          strScopePlus = "Static"
        End If
        For I = LBound(PRocDesc) To UBound(PRocDesc)
          If LenB(strOnForm) = 0 Or strOnForm = PRocDesc(I).PrdComp Then
            If LenB(strProcType) = 0 Or strProcType = PRocDesc(I).PrDClass Then
              If LenB(StrScope) = 0 Or StrScope = PRocDesc(I).PrDScope Or strScopePlus = PRocDesc(I).PrDScope Then
                If strTest = PRocDesc(I).PrDName Then
                  IsProcedure = True
                  Exit For
                End If
                'v2.8.4 Thanks Richard Brisley  This is the reported bug fix
                ' the control default property test didn't understand about Form.Procedure style calls
                If strTest = PRocDesc(I).PrdComp & "." & PRocDesc(I).PrDName Then
                  IsProcedure = True
                  Exit For
                End If
                'v2.9.1 improved test for cls referenced expands on v2.8.4 fix above
                If Not IsEmpty(PRocDesc(I).prClassAlias) Then
                  For J = 0 To UBound(PRocDesc(I).prClassAlias)
                    If strTest = PRocDesc(I).prClassAlias(J) & "." & PRocDesc(I).PrDName Then
                      IsProcedure = True
                      Exit For
                    End If
                  Next J
                End If
              End If
            End If
          End If
        Next I
      End If
    End If
  End If

End Function

Public Function IsProcedureName(ByVal strTest As String) As Boolean

  IsProcedureName = IsProcedure(strTest)
  If Not IsProcedureName Then
    IsProcedureName = IsDeclaration(strTest, , "Declare")
  End If

End Function

Public Function IsVBFunction(ByVal strTest As String) As Boolean

  IsVBFunction = InQSortArray(ArrQVBCommands, strTest)

End Function

Public Sub KillContaining(ByVal strTarget As String)

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim CurCompCount As Long
  Dim strKill      As String
  Dim X            As Long
  Dim CompMod      As CodeModule

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount) Then
        Set CompMod = Comp.CodeModule
        X = 1
        Do While CompMod.Find(strTarget, X, 1, -1, -1, False, False, True)
          strKill = CompMod.Lines(X, 1)
          If Len(strTarget) Then
            If strTarget = Trim$(strKill) Then
              CompMod.DeleteLines X, 1
              X = X - 1
             ElseIf SmartRight(strKill, strTarget) Then
              strKill = Left$(strKill, InStr(strKill, strTarget) - 1)
              CompMod.ReplaceLine X, strKill
             ElseIf SmartLeft(strKill, strTarget) Then
              'does away with plug line
              If SmartLeft(strTarget, Smiley) Then
                CompMod.DeleteLines X, 1
                X = X - 1
              End If
            End If
           Else
            If Left$(Trim$(strKill), Len(RGSignature)) = RGSignature Then
              CompMod.DeleteLines X, 1
              X = X - 1
             Else
              strKill = Left$(strKill, InStr(strKill, RGSignature) - 1)
              CompMod.ReplaceLine X, strKill
              X = X - 1
            End If
          End If
          X = X + 1
          If X > CompMod.CountOfLines Then
            Exit Do
          End If
        Loop
      End If
    Next Comp
  Next Proj
  On Error GoTo 0

End Sub

Public Function ListViewVisibleItems(lvw As ListView) As Long

  'Author: The VB2TheMax Team
  ' Return the number of items that can fit vertically in a ListView control
  ' when in List or Report mode (only fully visible items are counted)

  ListViewVisibleItems = SendMessage(lvw.hWnd, LVM_GETCOUNTPERPAGE, 0, ByVal 0&)

End Function

Private Function LSVWidthFromData(Lsv As ListView, _
                                  lngTypeMember As CtrlDescNo, _
                                  ByVal StrHeader As String, _
                                  Optional ByVal bForceShow As Boolean = False) As Long

  Dim strAny     As String
  Dim strTest    As String
  Dim bMulti     As Boolean
  Dim lngLongest As Long
  Dim I          As Long

  lngLongest = Len(StrHeader)
  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDUsage <> 2 Then
        If Not Lsv.FindItem(CntrlDesc(I).CDName, lvwSubItem) Is Nothing Then
          strTest = strTypeMember(I, lngTypeMember)
          If Not bMulti Then  ' only test until you get a True(satu cukup)
            If LenB(strAny) = 0 Then
              strAny = strTest
             ElseIf strAny <> strTest Then
              bMulti = True
            End If
          End If
          If lngLongest < Len(strTest) Then
            lngLongest = Len(strTest)
          End If
        End If
      End If
    Next I
  End If
  If Not bMulti And Not bForceShow Then
    LSVWidthFromData = 0
   Else
    LSVWidthFromData = Lsv.Parent.TextWidth(String$(lngLongest * 1.3, "X"))
  End If

End Function

Public Function ProtectConnectVBIDE(varCode As Variant) As Boolean

  'v 2.2.1 thanks rblanch
  ' this is the other fix you lead me to.
  ' case shift is probably not necessary but adds safety

  If InStr(LCase$(varCode), " as vbide.vbe") Or InStr(LCase$(varCode), " as connect") Then
    ProtectConnectVBIDE = True
  End If

End Function

Public Sub SetColumn(Lsv As ListView, _
                     ByVal strKey As String, _
                     lngTypeMember As CtrlDescNo, _
                     Optional bForceShow As Boolean = False)

  With Lsv.ColumnHeaders(strKey)
    If Not bForceShow Then
      .Width = 0
     Else
      .Width = LSVWidthFromData(Lsv, lngTypeMember, .Text, bForceShow)
    End If
  End With

End Sub

Public Sub SetDefArrays(arrT As Variant, _
                        ByVal strT As String, _
                        ByVal ModNum As Long)

  Select Case arrT(0)
   Case "DefLng"
    ArrQDefLng(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefSng"
    ArrQDefSng(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefInt"
    ArrQDefInt(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefBool"
    ArrQDefBool(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefByte"
    ArrQDefByte(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefCur"
    ArrQDefCur(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefDbl"
    ArrQDefDbl(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefDate"
    ArrQDefDate(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefStr"
    ArrQDefStr(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefObj"
    ArrQDefObj(ModNum) = QuickSortArray(Split(strT, ","))
   Case "DefVar"
    ArrQDefVar(ModNum) = QuickSortArray(Split(strT, ","))
  End Select

End Sub

Public Function SingletonCtrlArray(ByVal lngIndex As Long) As Boolean

  Dim countCtrl As Long
  Dim I         As Long

  If bCtrlDescExists Then
    If CntrlDesc(lngIndex).CDIndex <> "-1" Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If CntrlDesc(I).CDProj = CntrlDesc(lngIndex).CDProj Then
          If CntrlDesc(I).CDForm = CntrlDesc(lngIndex).CDForm Then
            If CntrlDesc(I).CDName = CntrlDesc(lngIndex).CDName Then
              countCtrl = countCtrl + 1
              If countCtrl > 1 Then
                Exit For 'unction
              End If
            End If
          End If
        End If
      Next I
      If countCtrl = 1 Then
        If Not DOPartialFind("Load " & CntrlDesc(lngIndex).CDName & "*") Then
          SingletonCtrlArray = True
        End If
      End If
    End If
  End If

End Function

Public Function StripPunctuation(ByVal VarInput As Variant, _
                                 Optional ByVal StrReplace As String = SngSpace, _
                                 Optional ByVal AllowInnerUnderScore As Boolean = False) As String

  Dim I As Long

  For I = 1 To Len(VarInput)
    If IsPunct(Mid$(VarInput, I, 1)) Then
      If AllowInnerUnderScore Then
        If Mid$(VarInput, I, 1) = "_" Then
          If Not I < Len(VarInput) Then
            Mid$(VarInput, I, 1) = StrReplace
          End If
         Else
          Mid$(VarInput, I, 1) = StrReplace
        End If
       Else
        Mid$(VarInput, I, 1) = StrReplace
      End If
    End If
  Next I
  Do While InStr(VarInput, StrReplace & StrReplace)
    VarInput = Replace$(VarInput, StrReplace & StrReplace, StrReplace)
  Loop
  VarInput = LStrip(VarInput, StrReplace)
  VarInput = RStrip(VarInput, StrReplace)
  If StrReplace = SngSpace Then
    Do While InStr(VarInput, SngSpace)
      VarInput = Replace$(VarInput, SngSpace, vbNullString)
    Loop
  End If
  StripPunctuation = VarInput

End Function

Private Function strTypeMember(ByVal lngTpyeNo As Long, _
                               lngTypeMember As CtrlDescNo) As String

  With CntrlDesc(lngTpyeNo)
    Select Case lngTypeMember
     Case e_CDName
      strTypeMember = .CDName
     Case e_CDform
      strTypeMember = .CDForm
     Case e_CDProj
      strTypeMember = .CDProj
     Case e_CDXPFrameBug
      strTypeMember = .CDXPFrameBug
     Case e_CDBadType
      strTypeMember = .CDBadType
     Case e_CDIndex
      strTypeMember = .CDIndex
     Case e_CDClass
      strTypeMember = .CDClass
     Case e_CDCaption
      strTypeMember = .CDCaption
     Case e_CDUsage
      strTypeMember = .CDUsage
    End Select
  End With

End Function

Private Function TypSuffixUpdateTarget(strCode As String, _
                                       StrTypeName As String) As Boolean

  Dim I        As Long

  'short circuites a multiple 'Or' line
  'v 2.2.2 reordered tests to make sure of catching API declares
  If Not JustACommentOrBlank(strCode) Then
    If InTypeDef(strCode, StrTypeName) Then
      TypSuffixUpdateTarget = True
     Else
      If InstrAtPositionArray(strCode, IpLeft, True, "Dim", "Global", "Public", "Private") Then
        If InStrCode(strCode, "Declare") Then
          'special tests for API declares
          If TypeSuffixExists(WordAfter(strCode, WordAfter(strCode, "Declare"))) Then
            'procedure has type suffix
            TypSuffixUpdateTarget = True
           Else
            'v2.2.2 detects paramaters with type suffix
            If InStr(strCode, "(") Then
              'v2.3.7 stop a crash if there are no brackets
              ' and the Type includes a dot separator
              ' to allow explicit referencing of Module.Type format
              For I = InStr(strCode, "(") To InStr(strCode, ")")
                If InQSortArray(TypeSuffixArray, Mid$(strCode, I, 1)) Then
                  TypSuffixUpdateTarget = True
                  Exit For 'unction
                End If
              Next I
            End If
          End If
         ElseIf Not Has_AS(strCode) Then
          'other declarations without 'As Type
          TypSuffixUpdateTarget = True
         ElseIf TypeSuffixExists(strCode) Then
          'or with type suffix
          TypSuffixUpdateTarget = True
        End If
      End If
    End If
  End If

End Function

Public Sub UpDate_Dim_Global_Declarations(ByVal ModuleNumber As Long, _
                                          dArray As Variant, _
                                          Comp As VBComponent)

  Dim L_CodeLine   As String
  Dim arrLine      As Variant
  Dim TmpDecArr    As Variant
  Dim I            As Long
  Dim UpDated      As Boolean
  Dim MaxFactor    As Long
  Dim TestWord     As String
  Dim ForcePrivate As Boolean
  Dim lngdummy     As Long
  Dim SuppressMsg  As Boolean
  Dim StrTypeName  As String ' v2.2.8 added tet for Type member named Type
  Dim arrTest1     As Variant
  Dim arrTest2     As Variant

  arrTest1 = Array("Const", "Declare", "Enum", "Type")
  arrTest2 = Array("Type", "Const")
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Automatically convert older sytle Dim and Global Declarations to Public
  'generate the PrivateTypeArray to help with deciding whether to go Public or Private
  If dofix(ModuleNumber, DimGlobal2PublicPrivate) Then
    Declaration_PrivateTypeLister ModuleNumber, dArray
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > -1 Then
      'v2.9.7 added ActiveXDesigner test
      ForcePrivate = IsComponent_ControlHolder(Comp) Or IsComponent_User_Class(Comp) Or IsComponent_ActiveXDesigner(Comp)
      For I = 0 To MaxFactor
        SuppressMsg = False
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          L_CodeLine = StripDoubleSpace(Trim$(TmpDecArr(I)))
          arrLine = Split(L_CodeLine)
          If IsInArray(arrLine(0), arrTest1) Then
            If InTypeDef(TmpDecArr(I), StrTypeName) Then
              'v2.3.6 added Const as name of a type member
              If IsInArray(arrLine(0), arrTest2) Then
                GoTo TypeMemberType
              End If
            End If
            TestWord = GetName(arrLine)
            Select Case arrLine(0)
              'Deal with unscoped stuff
             Case "Declare", "Const"
              '*if in class or form these are usually Private
              arrLine(0) = IIf(ForcePrivate, "Private ", "Public ") & arrLine(0)
             Case "Enum", "Event"
              'Event is always public (other wise it couldn't do anything)
              'Later tests will reduce Public Enum to Private if necessary for others
              arrLine(0) = "Public " & arrLine(0)
             Case "Type"
              If Not InStructure(TypeStruct, TmpDecArr, I, lngdummy, lngdummy) Then
                arrLine(0) = "Public " & arrLine(0)
               Else
                SuppressMsg = True
              End If
            End Select
            If LenB(arrLine(0)) Then
              'check that a guessed Public is not scoped too large
              ' all Event are public
              If InStr(Join(arrLine, SngSpace), CommaSpace) = 0 Then
                ' in case there is a multi-Dim with mixed members leave till later
                If Not MultiRight(arrLine(0), True, "Event", "Enum", "Type", "WithEvent") Then
                  If Left$(arrLine(0), 7) = "Public " Then
                    If Not FindCodeUsage(TestWord, vbNullString, Comp.Name, False, False, True) Then
                      arrLine(0) = Safe_Replace(arrLine(0), "Public", "Private")
                    End If
                  End If
                End If
              End If
              If Not SuppressMsg Then
                Select Case FixData(DimGlobal2PublicPrivate).FixLevel
                 Case CommentOnly
                  TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                   SUGGESTION_MSG & "Scope should be changed to " & arrLine(0)
                 Case FixAndComment
                  TmpDecArr(I) = Join(arrLine) & vbNewLine & WARNING_MSG & "Scope Changed to " & LeftWord(arrLine(0))
                 Case JustFix
                  TmpDecArr(I) = Join(arrLine)
                End Select
              End If
              AddNfix DimGlobal2PublicPrivate
              UpDated = True
            End If
TypeMemberType:
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

Public Sub UpDate_Excess_Public_To_Private_Declarations(ByVal ModuleNumber As Long, _
                                                        dArray As Variant, _
                                                        Comp As VBComponent)

  Dim L_CodeLine  As String
  Dim arrLine     As Variant
  Dim TmpDecArr   As Variant
  Dim I           As Long
  Dim UpDated     As Boolean
  Dim MaxFactor   As Long
  Dim TestWord    As String
  Dim SuppressMsg As Boolean
  Dim MUpdated    As Boolean
  Dim arrTest1    As Variant
  Dim arrTest2    As Variant

  arrTest1 = Array("Enum", "Type")
  arrTest2 = Array("Event", "WithEvents")
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Automatically convert declarations which are too widely scoped to PRivate
  If dofix(ModuleNumber, Public2Private) Then
    If LenB(GetActiveProject.FileName) Then
      If Comp.Type <> vbext_ct_ClassModule Then
        'classes may contain Public variables which are used as properties
        'to be referenced by new instances of the class so assume the author got it right
        Declaration_PrivateTypeLister ModuleNumber, dArray
        TmpDecArr = CleanArray(dArray)
        MaxFactor = UBound(TmpDecArr)
        If MaxFactor > -1 Then
          For I = 0 To MaxFactor
            SuppressMsg = False
            MemberMessage "", I, MaxFactor
            If Not JustACommentOrBlank(TmpDecArr(I)) Then
              L_CodeLine = StripDoubleSpace(Trim$(TmpDecArr(I)))
              If InstrAtPosition(L_CodeLine, "Public", IpLeft) Then
                arrLine = Split(L_CodeLine)
                TestWord = GetName(arrLine, True)
                If LenB(arrLine(0)) Then
                  'check that a guessed Public is not scoped too large
                  If Not IsInArray(arrLine(1), arrTest1) Then
                    If Not IsInArray(arrLine(1), arrTest2) Then
                      ' all Events are public
                      If Not FindCodeUsage(TestWord, vbNullString, Comp.Name, False, False, True) Then
                        '*if not used outside this component then should be Private
                        arrLine(0) = "Private"
                        UpDated = True
                      End If
                    End If
                  End If
                  If Not IsInArray(arrLine(1), arrTest1) Then
                    ' all Events are public
                    If Not FindCodeUsage(TestWord, vbNullString, Comp.Name, False, False, True) Then
                      '*if not used outside this component then should be Private unless it is being used to type cast a procedure or parameter
                      If FindCodeUsage(" As " & TestWord, vbNullString, Comp.Name, False, False, False) Then
                        If arrLine(0) <> "Public" Then
                          arrLine(0) = "Public"
                          UpDated = True
                        End If
                      End If
                    End If
                  End If
                  If UpDated Then
                    If Not SuppressMsg Then
                      Select Case FixData(Public2Private).FixLevel
                       Case CommentOnly
                        TmpDecArr(I) = TmpDecArr(I) & vbNewLine & _
                         SUGGESTION_MSG & "Scope should be changed to " & arrLine(0)
                       Case FixAndComment
                        TmpDecArr(I) = Join(arrLine) & vbNewLine & WARNING_MSG & "Scope Changed to " & arrLine(0)
                       Case JustFix
                        TmpDecArr(I) = Join(arrLine)
                      End Select
                    End If
                    AddNfix Public2Private
                    UpDated = False
                    MUpdated = True
                  End If
                End If
              End If
            End If
          Next I
          dArray = CleanArray(TmpDecArr, MUpdated)
        End If
      End If
    End If
  End If

End Sub

Public Sub UpDate_TypeSuffix_Declarations(ByVal ModuleNumber As Long, _
                                          dArray As Variant)

  Dim L_CodeLine  As String
  Dim UpDated     As Boolean
  Dim TmpDecArr   As Variant
  Dim I           As Long
  Dim MaxFactor   As Long
  Dim MyStr       As String
  Dim StrTypeName As String

  'This routine updates Type suffix Declares to 'As Type' format
  If dofix(ModuleNumber, UpdateDecTypeSuffix) Then
    TmpDecArr = CleanArray(dArray)
    MaxFactor = UBound(TmpDecArr)
    If MaxFactor > 0 Then
      For I = 1 To MaxFactor
        MemberMessage "", I, MaxFactor
        L_CodeLine = TmpDecArr(I)
        If Not JustACommentOrBlank(L_CodeLine) Then
          If TypSuffixUpdateTarget(L_CodeLine, StrTypeName) Then
            MyStr = TypeSuffixExtender(L_CodeLine)
            If MyStr <> L_CodeLine Then
              TmpDecArr(I) = MyStr
              MyStr = vbNullString
              AddNfix UpdateDecTypeSuffix
              UpDated = True
            End If
          End If
        End If
      Next I
      dArray = CleanArray(TmpDecArr, UpDated)
    End If
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:13:40 AM) 18 + 1316 = 1334 Lines Thanks Ulli for inspiration and lots of code.

