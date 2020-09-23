Attribute VB_Name = "mod_Arrays"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public ArrAllScopes                        As Variant
Public ArrFuncPropSub                      As Variant
Public arrOldChr                           As Variant
Public arrNewConst                         As Variant
Public arrQDeclarPresence                  As Variant
Public arrQEnumTypePresence                As Variant
Public Type DeclarDescriptor
  DDName                                   As String
  DDProj                                   As String
  DDComp                                   As String
  DDKind                                   As String
  DDClsRef                                 As Boolean
  DDType                                   As String
  DDScope                                  As String
  DDIndexing                               As Boolean
  DDUsage                                  As Variant
  DDComment                                As String
  DDArguments                              As String
End Type
Public DeclarDesc()                        As DeclarDescriptor
Public bDeclExists                         As Boolean
Public Type ProcedureDescriptor
  PrDName                                  As String
  PrDProj                                  As String
  PrdComp                                  As String
  prClassAlias                             As Variant
  prUSerCtlAlias                           As Variant
  PrdCompType                              As Long
  PrDSize                                  As Long
  PrDClass                                 As String
  PrDType                                  As String
  PrDControl                               As Boolean    'User Interface
  PrDLegalControl                          As Boolean    ' Control_Event is legal
  PrDLocation                              As String
  PrDInternal                              As Boolean    '_Initialize, _Terminate Form_ etc
  PrDGeneric                               As Boolean    'Form_ UserControl_ MdiForm_
  PrDArguments                             As String
  PrDArgumentTypes                         As String
  PrDScope                                 As String
  PrDCallsTo                               As Variant
  PrDCallsFrom                             As String
  PrDDims                                  As String
  PrDDuplicate                             As Boolean
  PrDAPI                                   As Boolean
  PrDUDTParam                              As Boolean
  PrDComment                               As String
End Type
Public PRocDesc()                          As ProcedureDescriptor
Public bProcDescExists                     As Boolean
Private ArrEndFuncPropSub                  As Variant
Public ArrExitFuncPropSub                  As Variant
Public UserCtrlEventArray                  As Variant
Private AlreadyLoaded                      As Boolean
Public QSortModNameArray                   As Variant
Public QSortModClassArray                  As Variant
Public QSortModBasArray                    As Variant
Public StandardTypes                       As Variant
Public StandardPreFix                      As Variant
Public StandardControl                     As Variant
Public StandardCtrPrefix                   As Variant
Public ArrayDummySplitPoint                As String
Public arrInternalTest1                    As Variant
Public arrInternalTest2                    As Variant
Public arrInternalTest3                    As Variant
Public arrPatternDetector                  As Variant
Public arrPatternDetector2                 As Variant
Public ArrQMaintainSpacing                 As Variant
Public ArrQVBReservedWords                 As Variant
Public ArrQIntTypes                        As Variant
Public ArrQVBCommands                      As Variant
'
Public PrivateTypeArray()                  As Variant
Public ArrQNoNullControl                   As Variant
Public ArrQDefLng()                        As Variant
Public ArrQDefSng()                        As Variant
Public ArrQDefInt()                        As Variant
Public ArrQDefBool()                       As Variant
Public ArrQDefByte()                       As Variant
Public ArrQDefCur()                        As Variant
Public ArrQDefDbl()                        As Variant
Public ArrQDefDate()                       As Variant
Public ArrQDefStr()                        As Variant
Public ArrQDefObj()                        As Variant
Public ArrQDefVar()                        As Variant
Public ArrQVBStructureWords                As Variant
Private arrOperators                       As Variant
Public TypeSuffixArray                     As Variant
Public AsTypeArray                         As Variant
Public ArrQStrVarFunc                      As Variant
Public CodeFixProtectedArray               As Variant
Public ObsoleteCodeArray                   As Variant
Public DangerousCodeArray                  As Variant
Public DangerousCodeLevelArray             As Variant
Public DangerousScriptArray                As Variant
Public DangerousScriptLevelArray           As Variant
Public DangerousAPIArray                   As Variant
Public DangerousStringArray                As Variant
Public DangerousAPILevelArray              As Variant
Public DangerousReferenceArray             As Variant
Private DangerousReferenceLevelArray       As Variant
Public ArrQKnowVBWord                      As Variant
Public ComponentCount                      As Long

Public Function AccumulatorString(ByVal StrAccum As String, _
                                  VarAdd As Variant, _
                                  Optional Delimiter As String = ",", _
                                  Optional ByVal NoRepeats As Boolean = True) As String

  'Allows you to build up a delimited string with no duplicate members or excess delimiters
  'Call:
  '           SomeString= AccumulatorString (SomeString, SomeOtherData)
  '
  'VarAdd allows you to add array members or strings
  'Optional Delimiter allows you to do further formatting if needed
  'Optional NoRepeats default True exclude duplicates set to false if you want them
  'NOTE if you want to add blanks make sure that VarAdd is at least a single space (" ")

  If LenB(VarAdd) Then
    If LenB(StrAccum) Then
      'Not already collected
      'v2.4.9 improced speed of test by using seperate function
      If Not InDelimitedString(StrAccum, VarAdd, Delimiter, NoRepeats) Then
        AccumulatorString = StrAccum & Delimiter & VarAdd
       Else
        AccumulatorString = StrAccum
      End If
     Else
      AccumulatorString = VarAdd
    End If
   Else
    AccumulatorString = StrAccum
  End If

End Function

Public Function AppendArray(ByVal VarArray As Variant, _
                            ByVal strAdd As String) As Variant

  Dim strT As String

  If Not IsEmpty(VarArray) Then
    strT = Join(VarArray, ",")
  End If
  strT = AccumulatorString(strT, strAdd, , False)
  AppendArray = Split(strT, ",")

End Function

Public Sub CompactLiteral(Arr As Variant)

  Dim I         As Long
  Dim J         As Long
  Dim QMCOUNT   As Long
  Dim KillCount As Long
  Dim strKill   As String

  Do
    strKill = RandomString(48, 122, 3, 6)
  Loop While InStr(Join(Arr), strKill)
  Do
    If Len(Arr(I)) Then
      If Left$(Arr(I), 1) = DQuote Then
        If IsOdd(CountSubString(Arr(I), DQuote)) Then
          J = I + 1
NotFinished:
          Do
            If J > UBound(Arr) Then
              Exit Do
            End If
            Arr(I) = Arr(I) & SngSpace & Arr(J)
            Arr(J) = strKill
            KillCount = KillCount + 1
            J = J + 1
            If J > UBound(Arr) Then
              Exit Do
            End If
            QMCOUNT = CountSubString(Arr(I), DQuote)
          Loop While IsOdd(QMCOUNT) 'QMCOUNT < 1
          If J < UBound(Arr) Then
            If IsOdd(QMCOUNT) Then
              Arr(I) = Arr(I) & SngSpace & Arr(J)
              Arr(J) = strKill
              KillCount = KillCount + 1
              I = J ' - 1
              'v2.1.2 modified to stop expanding spaces falsely
              ' Else
              'GoTo NotFinished
            End If
          End If
        End If
      End If
    End If
    I = I + 1
  Loop While I <= UBound(Arr)
  J = 0
  ReDim tarr(UBound(Arr) - KillCount) As Variant
  For I = LBound(Arr) To UBound(Arr)
    If Arr(I) <> strKill Then
      tarr(J) = Arr(I)
      J = J + 1
    End If
  Next I
  Arr = tarr

End Sub

Private Function EnumArray(ByVal strName As String, _
                           ByVal strPrj As String, _
                           ByVal StrFrm As String) As Variant

  
  Dim cMod  As CodeModule
  Dim Sline As Long
  Dim strA  As String

  Set cMod = VBInstance.VBProjects.Item(strPrj).VBComponents.Item(StrFrm).CodeModule
  cMod.Find "Enum " & strName, Sline, 1, -1, -1
  Do
    Sline = Sline + 1
    strA = AccumulatorString(strA, LeftWord(cMod.Lines(Sline, 1)), , False)
    strA = AccumulatorString(strA, cMod.Lines(Sline, 1), , False)
  Loop Until InStr(cMod.Lines(Sline + 1, 1), "End Enum")
  EnumArray = Split(strA, ",")

End Function

Private Function EnumUseageArray(strName As String, _
                                 strPrj As String, _
                                 StrFrm As String) As Variant

  
  Dim arrTest  As Variant
  Dim arrAccum As Variant
  Dim I        As Long

  arrTest = EnumArray(strName, strPrj, StrFrm)
  For I = LBound(arrTest) To UBound(arrTest) Step 2
    arrAccum = MergeArray(arrAccum, EnumUseageArrayTest(CStr(arrTest(I)), CStr(arrTest(I + 1))))
  Next I
  EnumUseageArray = NoBlanksArray(arrAccum)

End Function

Public Function ExpandForDetection(ByVal VarStr As Variant) As String

  If Not JustACommentOrBlank(VarStr) Then
    VarStr = ReplaceArray(VarStr, LBracket, SpacePad(LBracket), RBracket, SpacePad(RBracket), "-", SpacePad("-"), ",", SpacePad(","), ":=", SpacePad(":="), "= -", SpacePad("= -"))
    'special case X = -FunctionName(y)
    'ver 1.0.94 thanks to Martins Skujenieks's 'FX.DLL 1.03 SDK' which made this fix necessary
    Do While InStrCode(VarStr, DblSpace)
      VarStr = Replace$(VarStr, DblSpace, SngSpace)
    Loop
    ExpandForDetection = VarStr
  End If

End Function

Public Function ExpandForDetection3(ByVal VarStr As Variant) As String

  If Not JustACommentOrBlank(VarStr) Then
    VarStr = Safe_ReplaceArray(VarStr, LBracket, SpacePad(LBracket), RBracket, SpacePad(RBracket), "-", SpacePad("-"), ",", SpacePad(","), ":=", SpacePad(":="), "= -", SpacePad("= -"))
    'special case X = -FunctionName(y)
    'ver 1.0.94 thanks to Martins Skujenieks's 'FX.DLL 1.03 SDK' which made this fix necessary
    Do While InStrCode(VarStr, DblSpace)
      VarStr = Replace$(VarStr, DblSpace, SngSpace)
    Loop
  End If
  ExpandForDetection3 = VarStr

End Function

Public Sub FillArray(Arr As Variant, _
                     ByVal strNewData As String, _
                     Optional ByVal DoMerge As Boolean = False, _
                     Optional ByVal Delim As String = ",", _
                     Optional ByVal bQSort As Boolean = False)

  'ver 1.0.94
  'centralized common structure for array builders

  If DoMerge Then
    If LenB(strNewData) Then
      Arr = MergeArray(Arr, Split(strNewData, Delim), Delim)
     Else
      On Error GoTo oops
      If UBound(Arr) = -1 Then
        Arr = Split("")
      End If
    End If
   Else
    If LenB(strNewData) Then
      Arr = Split(strNewData, Delim)
      If bQSort Then
        Arr = QuickSortArray(Arr)
      End If
     Else
      If Not ArrayHasContents(Arr) Then
        Arr = Split("")
      End If
    End If
  End If
  On Error GoTo 0

Exit Sub

oops:
  Arr = Split("")

End Sub

Private Sub Generate_DeclarationArray()

  
  Dim strTmp       As String
  Dim HasIndex     As Boolean
  Dim I            As Long
  Dim MaxFactor    As Long
  Dim CurCompCount As Long
  Dim DDC          As Long
  Dim L_CodeLine   As String
  Dim TypeDef      As String
  Dim arrLine      As Variant
  Dim arrDecl      As Variant
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim arrTest      As Variant
  Dim strOrig      As String

  'b3.0.8 skips data already gathered by GenerateVariableData
  On Error GoTo BugTrap
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  arrTest = Array("Private", "Public", "Dim", "Global", "Static", "Friend", "Type", "Enum")
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        ModuleMessage Comp, CurCompCount
        If Comp.CodeModule.CountOfDeclarationLines Then
          arrDecl = GetDeclarationArray(Comp.CodeModule)
          MaxFactor = UBound(arrDecl)
          If MaxFactor > -1 Then
            'make the array enormous to avoid multiple Redim's
            DDC = UBound(DeclarDesc) + 1
            ReDim Preserve DeclarDesc(UBound(DeclarDesc) + MaxFactor) As DeclarDescriptor
            For I = 0 To MaxFactor
              L_CodeLine = arrDecl(I)
              HasIndex = False
              If Not JustACommentOrBlank(L_CodeLine) Then
                If ExtractCode(L_CodeLine) Then
                  'If Not InEnumCapProtection(Comp.CodeModule, arrDecl, I) Then
                  If IsInArray(WordInString(Trim$(L_CodeLine), 1), arrTest) Then
                    strOrig = L_CodeLine
                    L_CodeLine = ExpandForDetection(L_CodeLine)
                    arrLine = Split(L_CodeLine)
                    If Not IsDeclaration(GetName(arrLine), arrLine(0), , Comp.Name) Then
                      If Left$(Trim$(L_CodeLine), 5) = "Type " Then
                        If Has_AS(L_CodeLine) Then
                          GoTo IgnoreTypeAsMember
                        End If
                      End If
                      If UBound(arrLine) >= 1 Then
                        TypeDef = GetType(L_CodeLine)
                        If GetLeftBracketPos(L_CodeLine) Then
                          If arrLine(1) <> "Enum" Then
                            If UBound(arrLine) > 1 Then
                              If GetLeftBracketPos(arrLine(2)) Then
                                HasIndex = True
                              End If
                            End If
                          End If
                        End If
                        SetDeclareData DDC, Comp.Name, Proj.Name, strOrig, L_CodeLine, arrLine, TypeDef, HasIndex
                        MemberMessage DeclarDesc(DDC).DDName, I, MaxFactor
IgnoreTypeAsMember:
                      End If
                    End If
                    ' End If
                  End If
                End If
              End If
            Next I
          End If
        End If
      End If
      If Comp.Type = vbext_ct_PropPage Then
        'Changed is a special variable for property pages which is not explicitly declared
        'adding it in here stops it being detected as missing in dim section
        'v2.7.7 Thanks Evan Toder this was the bug with the PropertyPage Changed Property
        With DeclarDesc(DDC) 'UBound(DeclarDesc))
          .DDName = "Changed"
          .DDType = "Boolean"
          .DDIndexing = -1
          .DDComp = Comp.Name
          .DDProj = Proj.Name
          .DDScope = "Private"
          .DDKind = "Property"
        End With 'DeclarDesc(UBound(DeclarDesc))
        DDC = DDC + 1
      End If
      If DDC = UBound(DeclarDesc) Then
        'just a safety in case the enormous array is not enough
        ReDim Preserve DeclarDesc(UBound(DeclarDesc) + 30) As DeclarDescriptor
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  If DDC Then
    'trim off the extra blank members
    ReDim Preserve DeclarDesc(DDC - 1) As DeclarDescriptor
    bDeclExists = True
  End If
  For I = LBound(DeclarDesc) To UBound(DeclarDesc)
    strTmp = AccumulatorString(strTmp, DeclarDesc(I).DDName)
  Next I
  arrQDeclarPresence = QuickSortArray(Split(strTmp, ","))

Exit Sub

BugTrap:
  BugTrapComment "Generate_DeclarationArray"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Sub Generate_Publics()

  If Not bAborting Then
    ReDim DeclarDesc(0) As DeclarDescriptor
    WorkingMessage "Build Variable arrays", 1, 4
    GenerateVariableData ' this is fast but doesn't get Enums/Types
    WorkingMessage "Build Declarations Array", 2, 4
    Generate_DeclarationArray
    'ver 1.1.29 unscoped set as public initially mainly so that the default property function will recognize on first run.
    ReDim PRocDesc(0) As ProcedureDescriptor
    WorkingMessage "Get Procedure Data", 3, 4
    GenerateProcedureData
    WorkingMessage "Duplicate Procedure Data", 4, 4
    TestForDuplicateProcedures
  End If

End Sub

Private Sub GenerateProcedureData()

  Dim Comp          As VBComponent
  Dim memb          As Member
  Dim CurCompCount  As Long
  Dim Proj          As VBProject
  Dim CCMCounter    As Long
  Dim TotalProcs    As Long
  Dim CurProcNumber As Long

  'Copyright 2004 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'ver 1.1.79 simplified multiplee procedures into single sweep
  'ver 1.1.88 total replacement to use the Members access to subs
  'TODO get total number of procs then redim once
  TotalProcs = GetMemberCount
  ReDim PRocDesc(TotalProcs) As ProcedureDescriptor
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        ModuleMessage Comp, CurCompCount
        CCMCounter = 0
        For Each memb In Comp.CodeModule.Members
          If memb.Type = vbext_mt_Method Or memb.Type = vbext_mt_Property Then
            CCMCounter = CCMCounter + 1
            With Comp
              MemberMessage memb.Name, CCMCounter, .CodeModule.Members.Count
              If SetProcedureData(CurProcNumber, GetWholeLineCodeModule(.CodeModule, GetCodeDefLine(Comp, memb)), memb, Comp, Proj) Then
                CurProcNumber = CurProcNumber + 1
              End If
            End With 'Comp
          End If
        Next memb
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  bProcDescExists = CurProcNumber > 0
  If CurProcNumber Then
    ReDim Preserve PRocDesc(CurProcNumber - 1) As ProcedureDescriptor
  End If

End Sub

Private Sub GenerateVariableData()

  Dim strCode      As String
  Dim Comp         As VBComponent
  Dim CompMod      As CodeModule
  Dim memb         As Member
  Dim Proj         As VBProject
  Dim arrLine      As Variant
  Dim HasIndex     As Boolean
  Dim strOrig      As String
  Dim CCMCounter   As Long
  Dim TotalMemb    As Long
  Dim CurMemb      As Long
  Dim CurCompCount As Long
  Dim arrTest1     As Variant
  Dim arrTest2     As Variant

  arrTest1 = Array(vbext_mt_Const, vbext_mt_Event, vbext_mt_Variable)
  arrTest2 = Array("Private", "Public", "Dim", "Global", "Static", "Friend", "Const", "Declare")
  'Copyright 2004 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'ver 1.1.79 simplified multiple procedures into single sweep
  'ver 1.1.88 total replacement to use the Members access to subs
  TotalMemb = GetMemberCount
  ReDim DeclarDesc(TotalMemb) As DeclarDescriptor
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        ModuleMessage Comp, CurCompCount
        CCMCounter = 0
        Set CompMod = Comp.CodeModule
        For Each memb In CompMod.Members
          If IsInArray(memb.Type, arrTest1) Then
            CCMCounter = CCMCounter + 1
            With CompMod
              MemberMessage memb.Name, CCMCounter, .Members.Count
              strCode = GetWholeLineCodeModule(CompMod, memb.CodeLocation)
            End With 'Comp
            strOrig = strCode
            If ExtractCode(strCode) Then
              If IsInArray(LeftWord(strCode), arrTest2) Then
                strCode = ExpandForDetection(strCode)
                arrLine = Split(strCode)
                If UBound(arrLine) >= 1 Then
                  If GetLeftBracketPos(strCode) Then
                    Select Case arrLine(1)
                     Case "Const", "Event", "WithEvents"
                      HasIndex = False
                     Case Else
                      If UBound(arrLine) > 1 Then
                        If GetLeftBracketPos(arrLine(2)) Then
                          HasIndex = True
                        End If
                      End If
                    End Select
                  End If
                  With DeclarDesc(CurMemb)
                    .DDComp = Comp.Name
                    .DDProj = Proj.Name
                    .DDScope = ScopeName(memb.Scope, memb.Static)
                    .DDType = ProcType(strCode)
                    .DDKind = GetVariableClassfromMemb(memb)
                    .DDName = memb.Name
                    .DDIndexing = HasIndex
                    If arrLine(1) <> "Declare" Then
                      If InStr(strCode, " New ") Then
                        .DDClsRef = True
                      End If
                    End If
                    .DDUsage = VariableUseageArray(.DDName, strOrig)
                    If UBound(.DDUsage) = -1 Then
                      If .DDKind = "Enum" Then
                        'Check Use OF Enum Members
                        .DDUsage(0) = "Enum Name not used. UNDER DEVELOPMENT check for Enum Members"
                      End If
                    End If
                  End With
                  CurMemb = CurMemb + 1
IgnoreTypeAsMember:
                End If
              End If
            End If
          End If
        Next memb
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  If CurMemb Then
    ReDim Preserve DeclarDesc(CurMemb - 1) As DeclarDescriptor
    bDeclExists = True
  End If

End Sub

Private Function GetClassAlias(ByVal strCompName As String) As Variant

  Dim I      As Long
  Dim strTmp As String

  For I = 0 To UBound(DeclarDesc)
    If DeclarDesc(I).DDType = strCompName Then
      strTmp = AccumulatorString(strTmp, DeclarDesc(I).DDName)
    End If
  Next I
  GetClassAlias = Split(strTmp, ",")

End Function

Private Function GetCodeDefLine(Comp As VBComponent, _
                                mem As Member) As Long

  Dim I     As Long
  Dim K     As Long
  Dim telem As Variant

  With Comp.CodeModule
    For I = 1 To 4
      K = Choose(I, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set, vbext_pk_Proc)
      On Error Resume Next
      telem = Null
      telem = Array(Comp.Name, .ProcBodyLine(mem.Name, K), .ProcCountLines(mem.Name, K), .PRocStartLine(mem.Name, K), K)
      On Error GoTo 0
      If Not IsNull(telem) Then
        If telem(1) < mem.CodeLocation Then
          GetCodeDefLine = mem.CodeLocation
          Exit For
         Else
          GetCodeDefLine = telem(1)
          Exit For
        End If
      End If
    Next I
    If IsNull(telem) Then
      GetCodeDefLine = mem.CodeLocation
    End If
  End With

End Function

Private Function GetMemberCount() As Long

  Dim Proj As VBProject
  Dim Comp As VBComponent

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        GetMemberCount = GetMemberCount + Comp.CodeModule.Members.Count
      End If
    Next Comp
  Next Proj

End Function

Private Function GetParamNameString(ByVal strP As String) As String

  Dim arrTmp As Variant
  Dim I      As Long

  strP = GetParamString(strP)
  If LenB(strP) Then
    If strP <> "NONE" Then
      arrTmp = Split(strP, ",")
      For I = LBound(arrTmp) To UBound(arrTmp)
        arrTmp(I) = Replace$(arrTmp(I), "Optional ", vbNullString)
        arrTmp(I) = Replace$(arrTmp(I), "ByVal ", vbNullString)
        arrTmp(I) = Replace$(arrTmp(I), "ByRef ", vbNullString)
        GetParamNameString = GetParamNameString & WordInString(Trim$(arrTmp(I)), 1) & ","
      Next I
      Do While InStr(GetParamNameString, ",,")
        GetParamNameString = Replace$(GetParamNameString, ",,", ",")
      Loop
      GetParamNameString = Left$(GetParamNameString, Len(GetParamNameString) - 1)
    End If
  End If

End Function

Public Function GetParamString(ByVal strC As String) As String

  Dim LBPos As Long

  LBPos = InStr(strC, LBracket)
  If LBPos Then
    GetParamString = Mid$(strC, LBPos)
    GetParamString = Left$(GetParamString, InStrRev(GetParamString, RBracket))
    Do While CountSubString(GetParamString, LBracket) <> CountSubString(GetParamString, RBracket)
      GetParamString = Left$(GetParamString, Len(GetParamString) - 1)
    Loop
    'GetParamString = RStrip(LStrip(GetParamString, "("), ")")
    GetParamString = RStrip(LStrip(GetParamString, LBracket), RBracket)
    If LenB(GetParamString) = 0 Then
      GetParamString = "NONE"
    End If
  End If

End Function

Private Function GetParamTypeString(ByVal strP As String) As String

  Dim arrTmp As Variant
  Dim I      As Long

  If LenB(strP) Then
    If strP <> "NONE" Then
      strP = GetParamString(strP)
      arrTmp = Split(strP, ",")
      For I = LBound(arrTmp) To UBound(arrTmp)
        If Has_AS(arrTmp(I)) Then
          GetParamTypeString = GetParamTypeString & GetType(arrTmp(I)) & ","
         Else
          If Len(arrTmp(I)) Then
            GetParamTypeString = GetParamTypeString & "UnDefinedVariant,"
          End If
        End If
      Next I
      Do While InStr(GetParamTypeString, ",,")
        GetParamTypeString = Replace$(GetParamTypeString, ",,", ",")
      Loop
      GetParamTypeString = Left$(GetParamTypeString, Len(GetParamTypeString) - 1)
    End If
  End If

End Function

Private Function GetProcKind(ByVal strCode As String) As String

  If ContainsWholeWord(strCode, "Sub") Then
    GetProcKind = "Sub"
   ElseIf ContainsWholeWord(strCode, "Function") Then
    GetProcKind = "Function"
   ElseIf ContainsWholeWord(strCode, "Property") Then
    GetProcKind = "Property"
  End If

End Function

Public Function GetType(ByVal varCode As Variant) As String

  'v2.1.5 centralised this action

  If Get_As_Pos(varCode) > 0 Then
    GetType = WordAfter(varCode, " As ")
    If GetType = "New" Then
      GetType = WordAfter(varCode, " New ")
    End If
    Do While IsPunct(Right$(GetType, 1))
      GetType = Left$(GetType, Len(GetType) - 1)
    Loop
   Else
    GetType = vbNullString
  End If

End Function

Private Function GetVariableClassfromMemb(ByVal memb As Member) As String

  Select Case memb.Type
   Case vbext_mt_Const
    GetVariableClassfromMemb = "Const"
   Case vbext_mt_Event
    GetVariableClassfromMemb = "Event"
   Case vbext_mt_Method
    GetVariableClassfromMemb = "Method"
   Case vbext_mt_Property
    GetVariableClassfromMemb = "Property"
   Case vbext_mt_Variable
    GetVariableClassfromMemb = "Variable"
   Case 5 ' NO SUPPORTED Const vbext_mt_Enum
    GetVariableClassfromMemb = "Enum"
   Case 7 ' NO SUPPORTED Const vbext_mt_EventSink
    GetVariableClassfromMemb = "EventSink"
  End Select

End Function

Public Function GetWholeLineCodeModule(cMod As CodeModule, _
                                       ByVal Lnum As Long, _
                                       Optional LUpDate As Long) As String

  GetWholeLineCodeModule = cMod.Lines(Lnum, 1)
  Do While HasLineCont(GetWholeLineCodeModule)
    Lnum = Lnum + 1
    GetWholeLineCodeModule = Left$(GetWholeLineCodeModule, Len(GetWholeLineCodeModule) - 1) & Trim$(cMod.Lines(Lnum, 1))
  Loop
  LUpDate = Lnum

End Function

Private Function InDelimitedString(ByVal strMain As String, _
                                   ByVal varFind As Variant, _
                                   ByVal StrDelim As String, _
                                   ByVal NoRepeats As Boolean) As Boolean

  'support for AccumulatorString
  'v2.4.9 improced speed of test by using seperate function

  If NoRepeats Then
    If InStr(strMain, varFind) Then
      If strMain = varFind Then 'single member
        InDelimitedString = True
       ElseIf InStr(strMain, StrDelim & varFind & StrDelim) Then ' middle member
        InDelimitedString = True
       ElseIf InStr(strMain, varFind & StrDelim) = 1 Then 'firstmember
        InDelimitedString = True
       ElseIf InStr(strMain, StrDelim & varFind) = Len(strMain) - Len(StrDelim & varFind) + 1 Then
        'last member
        InDelimitedString = True
      End If
    End If
   Else
    'don't care, just do it
    InDelimitedString = False
  End If

End Function

Public Function InQSortArray(ByVal SortedArray As Variant, _
                             ByVal FindMe As String) As Boolean

  Dim Low    As Long
  Dim Middle As Long
  Dim High   As Long
  Dim Trap   As Boolean
  Dim TestMe As Variant

  If Not IsEmpty(SortedArray) Then
    If Not IsMissing(SortedArray) Then
      If UBound(SortedArray) > -1 Then
        'Binary search module very fast but requires array to be sorted
        Low = LBound(SortedArray)
        High = UBound(SortedArray)
        If High >= Low Then
          ' invert for Descending sorted Arrays
          If SortedArray(Low) > SortedArray(High) Then
            SwapAnyThing Low, High
          End If
          High = High + 1
          Do Until High - Low = 0
            Middle = (Low + High) \ 2
            ' see note below*
            If Trap Then
              Middle = Low
              High = Low
            End If
            TestMe = SortedArray(Middle) ' assign once to test twice
            If TestMe >= FindMe Then
              ' Only tests half the time
              If TestMe = FindMe Then
                InQSortArray = True
                Exit Do 'Function
              End If
              High = Middle
             Else
              Low = Middle
            End If
            Trap = (Low = High - 1)
          Loop
         ElseIf High = Low Then
          'single member test
          InQSortArray = SortedArray(Low) = FindMe
        End If
      End If
    End If
  End If

End Function

Public Sub LoadArrays()

  'only needed once

  If Not AlreadyLoaded Then
    arrOldChr = Array("vbCrLf", "Chr$(13) & Chr$(10)", "Chr$(9)", "Chr$(13)", "Chr$(10)", "Chr$(0)", "Chr$(8)")
    arrNewConst = Array("vbNewline", "vbNewline", "vbTab", "vbCr", "vbLf", "vbNullChar", "vbBack")
    arrOperators = Array("+", "-", "*", "/", "\", "^")
    ArrQVBStructureWords = QuickSortArray(Array("For", "Next", "Do", "While", "Until", "Loop", "Exit", "If", "Else", "ElseIf", "End", "Sub", "Function", "Property", "Case", "Select", "Option", "Explicit", "Base", "Wend", "Open", "Close"))
    UserCtrlEventArray = Array("_AccessKeyPress", "_AmbientChanged", "_AsyncReadComplete", "_AsyncReadProgress", "_Click", "_DblClick", _
                               "_DragDrop", "_DragOver", "_EnterFocus", "_ExitFocus", "_GetDataMember", "_GotFocus", "_Hide", "_HitTest", _
                               "_Initialize", "_InitProperties", "_KeyDown", "_KeyPress", "_KeyUp", "_LostFocus", "_MouseDown", "_MouseMove", _
                               "_MouseUp", "_OLECompleteDrag", "_OLEDragDrop", "_OLEDragOver", "_OLEGiveFeedback", "_OLESetData", "_OLEStartDrag", _
                               "_Paint", "_ReadProperties", "_Resize", "_Show", "_Terminate", "_WriteProperties")
    'Initialize public arrays
    'These allow the ReDoLineContinuation system to work
    ArrAllScopes = Array("Public", "Private", "Friend", "Static")
    ArrFuncPropSub = Array("Function", "Sub", "Property")
    ArrExitFuncPropSub = Array("Exit Function", "Exit Sub", "Exit Property")
    ArrEndFuncPropSub = Array("End Function", "End Sub", "End Property")
    'These allow translating old style Type suffixes into As Type
    'v2.3.8 alpha sorted of InQSortArray on TypeSuffixArray
    TypeSuffixArray = Array("!", "#", "$", "%", "&", "@")
    AsTypeArray = Array("Single", "Double", "String", "Integer", "Long", "Currency")
    'AsTypeArray is NOT qsortable it has to align with TypeSuffixArray
    StandardTypes = Array("Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long", "Object", "OLE_COLOR", "Single", _
                          "String", "Variant")
    StandardPreFix = Array("bln", "byt", "cur", "dte", "dbl", "int", "lng", "obj", "ole", "sng", "str", "var")
    'v2.6.6 added aDocd/ado to arrays
    StandardControl = Array("PictureBox", "Label", "TextBox", "Frame", "Form", "CommandButton", "CheckBox", "OptionButton", "ComboBox", _
                            "ListBox", "Timer", "Image", "HScrollBar", "VScrollBar", "DriveListBox", "DirListBox", "FileListBox", _
                            "Shape", "Line", "Data", "OLE", "TabStrip", "Toolbar", "StatusBar", "ProgressBar", "TreeView", "ListView", _
                            "ImageList", "Slider", "ImageCombo", "RichTextBox", "MSFlexGrid", "SSTab", "CommonDialog", "Animation", _
                            "UpDown", "MonthView", "DTPicker", "FlatScrollBar", "Inet", "MSChart", "Winsock", "MAPISession", "MAPIMessages", _
                            "MMControl", "PictureClip", "SysInfo", "MSComm", "MaskEdBox", "DataGrid", "DataList", "DataCombo", "CoolBar", _
                            "MSHFlexGrid", "Menu", "WebBrowser", "Adodc")
    StandardCtrPrefix = Array("pic", "lbl", "txt", "fra", "frm", "cmd", "chk", "opt", "cmb", "lst", "tmr", "img", "hsc", "vsc", "drv", _
                              "dir", "fil", "shp", "lin", "dta", "ole", "tbs", "tlb", "stb", "prg", "trv", "lsv", "iml", "sld", "imc", _
                              "rtb", "grd", "sstab", "cdl", "ani", "udn", "mnv", "dtp", "fsc", "net", "cht", "wns", "mps", "mpm", _
                              "mmc", "ptc", "snf", "msc", "meb", "dgrd", "dlst", "dcmb", "cbr", "mshgrd", "mnu", "wbb", "ado")
    'KnownVBWord contains VB Properties which can appear in code without
    'Me/FormName/UserControl identifier 'Property = X'
    '                       rather than 'Me.Property = X'
    'and is used to protect them from the DimMissing detector
    ArrQKnowVBWord = Array("AccessKeys", "Appearance", "AutoRedraw", "BackColor", "BorderStyle", "Caption", "ClipControls", "ControlBox", _
                           "DrawMode", "DrawStyle", "DrawWidth", "Enabled", "FillColor", "FillStyle", "Font", "FontTransparent", _
                           "ForeColor", "HasDC", "Height", "HelpContextID", "Icon", "KeyPreview", "Left", "LinkMode", "LinkTopic", _
                           "MaskColor", "MaskPicture", "MaxButton", "MouseIcon", "MDIChild", "MinButton", "MouseIcon", "MousePointer", _
                           "Moveable", "NegociateMenus", "OLDDropMode", "Palette", "PaletteMode", "Picture", "RightToLeft", "ScaleHeight", _
                           "ScaleLeft", "ScaleMode", "ScaleTop", "ScaleWidth", "ShowInTaskbar", "StartUpPosition", "Tag", "Top", _
                           "Visible", "WhatsThisButton", "WhatsThisHelp", "Width", "WindowState")
    ArrQKnowVBWord = QuickSortArray(ArrQKnowVBWord)
    'These are Ulli's orignal list of VB functions whose efficency is greatly enhanced by using them as string functions rather than Variants
    'by adding a $ to calls program speed is increased and overhead reduced
    'I have added Replace to the list
    'thanks to Rudz for suggesting Dir
    'ver 11.49 now gets the names from VB itself
    'v2.8.8 thanks to Anele Mbanga whose code let me track down a bug in this
    ' the VB lst misses  Relpace command so added it
    ArrQStrVarFunc = QuickSortArray(AppendArray(RefLibGenerateVBCommandStrVarArray, "Replace"))
    'these are teporary code that Code Fixer inserts and deletes while running
    CodeFixProtectedArray = Array(StrReverse("naelooB sA snoitaralceD_fo_dne_ta_PCE_tcetorP_oT_ekaF etavirP"), StrReverse("naelooB sA snoitaralceDnIstnemmoCceDpeeKoTymmuD etavirP"), _
                                  StrReverse("naelooB sA DETELED_EB_LLIW_YMMUD_NOITARALCED_FO_DNE etavirP"))
    'these are older style code which are marked for updating
    ObsoleteCodeArray = Array("GoSub", "Return", "Switch", "Choose", "IsMissing")
    'NB While Wend have their own fixer
    'These are VB commands with at least some level of risk in that they could be used to damage your files and disks
    DangerousCodeArray = Array("Kill", "MkDir", "RmDir", "FileCopy", "SetAttr", "Save", "ChDrive", "ChDir", "OutPut")
    DangerousCodeLevelArray = Array(4, 2, 4, 2, 3, 3, 1, 1, 1)
    DangerousScriptArray = Array(".shell", ".regwrite", ".regdelete", ".regread")
    DangerousScriptLevelArray = Array(4, 5, 5, 3)
    'These are API calls with at least some level of risk in that they could be used to damage your files and disks
    'NB this list is not exhaustive it is based on an hour or so reading thru an API book looking for dangerous sounding APIs
    DangerousAPIArray = Array("WinExec", "TerminateProcess", "ExitProcess", "RemoveDirectory", "DeleteFile", "SetVolumeLabel", _
                              "ExitWindows", "ExitWindowsEx", "SetComputerName", "DeletePrinter", "DeletePort", "DeletePrintDrive", _
                              "DeleteMoniter", "DeletePrintProcessor", "DeletePrintProvider", "RegDeleteKey", "RegDeleteKeyEx", "RegDeleteValue", _
                              "RegReplaceKey", "RegSetValue", "RegSetValueEx", "RegCreateKey", "RegCreateKeyEx", "MoveFile", "MoveFileEx", _
                              "WriteFile", "LoadKeyBoardLayout", "SetLocaleInfo", "SetLocalTime", "SetSysTime", "SetSystemTimeAdjustInformation", _
                              "SetSysColors", "SetEnvironmentVariable", "SystemParametersInfo", "SystemParametersInfoByval", "UnloadKeyboardLayout", _
                              "RtlAdjustPrivilege", "NtShutdownSystem", "RegisterServiceProcess")
    'These danger levels determine which of the messages to display
    '(Except for highest 5 (Never, ever) the levels are subjective
    DangerousAPILevelArray = Array(3, 4, 4, 5, 4, 5, 5, 5, 5, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 3, 2, 2, 3, 3, 2, 5, 5, 4, 4, 4, 4, 4, _
                                   4, 4, 5, 5, 5, 5)
    'Irritating code
    ' ShowCursor - Check it is turned back on in code some where
    'These are references with
    DangerousReferenceArray = Array("Shell32.Shell")
    DangerousReferenceLevelArray = Array(3)
    'These are strings with  at least some level of risk in that they could be used to damage your files and disks
    'v2.6.4 added a few more
    'v2.9.0 added regsvr32
    DangerousStringArray = Array("regsvr32", "Software\Microsoft\Windows\", "\CurrentVersion\Policies\System", "\Policies\System", "\System", _
                                 "\CurrentVersion\Run")
    'because code fix reacts to any line containing these comments
    'the string reverse protects Code Fixer from reacting to itself
    '    InCodeDontTouchOn = StrReverse("NO_REXIF_EDOC_DNEPSUS'")
    '   InCodeDontTouchOff = StrReverse("FFO_REXIF_EDOC_DNEPSUS'")
    'this array is not used at present but will add soon
    'array lifeted directly from PSC upload
    'Author: 80 SpitFire 08
    'Title: Convert Your VB code to color coded HTML (Updated finnally)
    'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=45658&lngWId=1
    ArrQVBReservedWords = QuickSortArray(Array("Alias", "And", "As", "Base", "Binary", "Boolean", "Byte", _
"ByVal", "Call", "Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "Close", _
"Compare", "Const", "CSng", "CStr", "Currency", "CVar", "CVErr", "Date", "Decimal", "Declare", "DefBool", _
"DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng", "DefObj", "DefSng", "DefStr", _
"DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "Enum", "Eqv", "Erase", "Error", "Exit", _
"Explicit", "False", "For", "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "In", "Input", _
"Input", "Integer", "Is", "LBound", "Let", "Lib", "Like", "Line", "Lock", "Long", "Loop", "LSet", "Name", _
"New", "Next", "Not", "Object", "Open", "Option", "On", "Or", "Output", "Preserve", "Print", "PrintForm", _
"Private", "Property", "Public", "Put", "Random", "Read", "ReDim", "Resume", "Return", "RSet", "Seek", _
"Select", "Set", "Single", "Spc", "Static", "String", "Stop", "Sub", "Tab", "Then", "True", "UBound", _
"Val", "Variant", "While", "Wend", "With"))

    'v2.6.5 added safety a list of types know to require integer members
    ' test is case insensitive but if you use VB's API Viewer or AllAPI's APIViewer
    ' the strings will usually be all caps (a few exceptions) Thanks Aaron Spivey
    ArrQIntTypes = QuickSortArray(Array("ACCEL", "ACL", "ACTION_HEADER", "ADAPTER_STATUS", "AUXCAPS", "BITMAP", _
"BITMAPCODEHEADER", "BITMAPFILEHEADER", "BITMAPINFOHEADER", "BITMAPV4HEADER", "CHAR_INFO", "COLORADJUSTMENT", _
"COMMCONFIG", "COMMPROP", "CONSOLE_sCREEN_BUFFER_INFO", "COORD", "CREATE_PROCESS_DEBUG_INFO", "DCB", _
"DDEACK", "DDEADVISE", "DDEDATA", "DDELN", "DDEPOKE", "DDEUP", "DEVMODE", "DEVMANES", "DLGITEMTEMPLATE", _
"DLGTEMPLATE", "EMREXTSELECTCLIPRGN", "EMRFILLRGN", "EMRFRAMERGN", "EMRGDICOMMENT", "EMRINVERTRGN", _
"EMRPAINTRGN", "EMRPOLYDRAW", "EMRPOLYDRAW16", "ENHMETAHEADER", "EVENTLOGRECORD", "FIND_NAME_BUFFER", _
"FIND_NAME_HEADER", "FINDREPLACE", "FIXED", "GLYPHMETRICS", "GUID", "ICONDIR", "ICONDIRENTRY", "JOYCAPS", _
"JOYINFO", "KERNINGPAIR", "KEY_EVENT_RECODR", "LANA_ENUM", "LDT_ENTRY", "LOAD_DLL_DEBUG_INFO", "LOGPALETTE", _
"MCI_STATUS_PARMS", "MCI_WAVE_SET_PARMS", "MEMICONDIRENTRY", "MENUITEMTEMPLATE", "MENUITEMTEMPLATEHEADER", _
"METAHEADER", "METARECORD", "MIDIINCAPS", "MIDIOUTCAPS", "MIXERCAPS", "NAME_BUFFER", "NCB", "NMLVKEYDOWN", _
"OFSTRUCT", "OPENFILENAME", "OUTPUT_DEBUG_STRING_INFO", "PCMWAVEFORMAT", "PELARRAY", "PIXELFORMATDESCRIPTOR", _
"POINTS", "PRINTDLG", "RASTERIZER_STATUS", "SAFEARRAY2D", "SECURITY_QUALITY_OF_SERVICE", "SESSION_BUFFER", _
"SESSION_HEADER", "SHFILEOPSTRUCT", "SMALL_RECT", "SOCKADDR", "STARTUPINFO", "SYSTEMTIME", "TARGET", _
"TIME_ZONE_INFORMATION", "TSYSTEM_PROCESSOR_INFORMATION", "TTPOLYCURVE", "WAVEFORMAT", "WAVEINCAPS", _
"WAVEOUTCAPS", "WSADATA"))

    ArrQVBCommands = QuickSortArray(RefLibGenerateVBCommandArray)
    ArrQNoNullControl = QuickSortArray(Array("Inet"))
    ArrQMaintainSpacing = MergeArray(ArrQVBReservedWords, ArrQKnowVBWord, ",")
    ArrQMaintainSpacing = QuickSortArray(MergeArray(ArrQMaintainSpacing, ArrQStrVarFunc, ","))
    ArrayDummySplitPoint = StrReverse("YMMUDYARRAREXIFEDOC")
    LoadUserSettings
    'v3.0.4 added to speed up InternalEvidence tests
    arrInternalTest1 = Array(DQuote, "$", " & ", "Join", "InputBox", "vbNullString")
    arrInternalTest2 = Array("MsgBox", "vbOK", "vbYes", "vbNo", "vbCancel", "vbAbort", "vbRetry", "vbIgnore")
    arrInternalTest3 = Array("Sqr(", "Sqr", "Timer", "Round", "Shell")
    arrPatternDetector = Array("*", "!", "[", "]", "\")
    arrPatternDetector2 = Array("/", "*")
    'This is a trick to keep the function from being marked unused
    'until I find a use for it :)
    AlreadyLoaded = InQSortArray(ArrQVBReservedWords, "Alias")
    AlreadyLoaded = True
  End If

End Sub

Private Function MergeArray(A As Variant, _
                            B As Variant, _
                            Optional Delim As String = SngSpace) As Variant

  ' TESTED BY: RG 30-Aug-98 14:22:09

  If IsEmpty(A) Then
    If IsEmpty(B) Then
      MergeArray = Split("")
      GoTo SafeExit
    End If
  End If
  If IsArray(A) Then
    If Not IsArray(B) Then
      If LenB(B) Then
        MergeArray = Split(Join(A, Delim) & Delim & B)
       Else
        MergeArray = A
      End If
      GoTo SafeExit
    End If
  End If
  If IsArray(B) Then
    If Not IsArray(A) Then
      If LenB(A) Then
        MergeArray = Split(Join(B, Delim) & Delim & A)
       Else
        MergeArray = B
      End If
      GoTo SafeExit
    End If
  End If
  MergeArray = Split(Join(A, Delim) & Delim & Join(B, Delim), Delim)
SafeExit:

End Function

Public Function NoBlanksArray(ByVal Arr As Variant) As Variant

  Dim lngNewIndex As Long
  Dim lngCount    As Long

  'set temp array to maximum possible size
  ReDim arrTmp(UBound(Arr)) As String
  Do
    'add 1st member (it can't be duplicate;))
    If Len(Arr(lngCount)) Then
      arrTmp(lngNewIndex) = Arr(lngCount)
      lngNewIndex = lngNewIndex + 1
    End If
    'increment the temp array counter
    lngCount = lngCount + 1
  Loop Until lngCount > UBound(Arr)
  'delete the unused members of the temp array
  'including the empty one generated by last pass over 'lngNewIndex = lngNewIndex + 1'
  'v2.7.5 crash if there is only a single blank member
  If lngNewIndex - 1 > -1 Then
    ReDim Preserve arrTmp(lngNewIndex - 1)
    NoBlanksArray = arrTmp
  End If

End Function

Private Function NonStandardArg(ByVal strArg As String) As Boolean

  Dim arrTmp As Variant
  Dim I      As Long
  Dim J      As Long

  If Len(strArg) Then
    arrTmp = Split(strArg, ",")
    For I = LBound(arrTmp) To UBound(arrTmp)
      If Not InQSortArray(StandardTypes, arrTmp(I)) Then
        If bDeclExists Then
          For J = LBound(DeclarDesc) To UBound(DeclarDesc)
            If DeclarDesc(J).DDKind = "Type" Then
              If DeclarDesc(J).DDName = arrTmp(I) Then
                NonStandardArg = True
                Exit For
              End If
            End If
          Next J
         Else
          NonStandardArg = True
        End If
      End If
      If NonStandardArg Then
        Exit For
      End If
    Next I
  End If

End Function

Private Function ProcType(ByVal strC As String) As String

  If WordInString(strC, -2) = "As" Then
    ProcType = WordInString(strC, -1)
    If Right$(ProcType, 1) = RBracket Then
      If Right$(ProcType, 2) <> "()" Then
        ProcType = vbNullString
      End If
    End If
   Else
    If Has_AS(strC) Then
      If InStr(strC, "Const ") Then
        ProcType = GetType(strC)
      End If
    End If
  End If

End Function

Public Function QSortArrayPos(ByVal SortedArray As Variant, _
                              ByVal FindMe As String) As Long

  Dim Low    As Long
  Dim Middle As Long
  Dim High   As Long
  Dim Trap   As Boolean
  Dim TestMe As Variant

  QSortArrayPos = -1 ' default missing
  'Binary search module very fast but requires array to be sorted
  Low = LBound(SortedArray)
  High = UBound(SortedArray)
  If High >= Low Then
    ' invert for Descending sorted Arrays
    If SortedArray(Low) > SortedArray(High) Then
      SwapAnyThing Low, High
    End If
    High = High + 1
    Do Until High - Low = 0
      Middle = (Low + High) \ 2
      ' see note below*
      If Trap Then
        Middle = Low
        High = Low
      End If
      TestMe = SortedArray(Middle) ' assign once to test twice
      If TestMe >= FindMe Then
        ' Only tests half the time
        If TestMe = FindMe Then
          QSortArrayPos = Middle
          Exit Do 'Function
        End If
        High = Middle
       Else
        Low = Middle
      End If
      Trap = (Low = High - 1)
    Loop
  End If

End Function

Private Sub QuickSortAD(AnArray As Variant, _
                        Lo As Long, _
                        Hi As Long, _
                        Optional Ascending As Boolean = True)

  Dim NewHi      As Long
  Dim CurElement As Variant
  Dim NewLo      As Long

  NewLo = Lo
  NewHi = Hi
  CurElement = AnArray((Lo + Hi) / 2)
  Do While (NewLo <= NewHi)
    If Ascending Then
      Do While AnArray(NewLo) < CurElement And NewLo < Hi 'Ascending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement < AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
     Else
      Do While AnArray(NewLo) > CurElement And NewLo < Hi 'Descending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement > AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
    End If
    If NewLo <= NewHi Then
      SwapAnyThing AnArray(NewLo), AnArray(NewHi)
      NewLo = NewLo + 1
      NewHi = NewHi - 1
    End If
  Loop
  If Lo < NewHi Then
    QuickSortAD AnArray, Lo, NewHi, Ascending
  End If
  If NewLo < Hi Then
    QuickSortAD AnArray, NewLo, Hi, Ascending
  End If

End Sub

Public Function QuickSortAppend(ByVal Arr As Variant, _
                                varAppend As Variant, _
                                Optional ByVal bAscending As Boolean = True) As Variant

  If IsEmpty(Arr) Then
    QuickSortAppend = Split(varAppend)
   Else
    If InQSortArray(Arr, varAppend) Then
      QuickSortAppend = Arr
     Else
      QuickSortAppend = QuickSortArray(AppendArray(Arr, varAppend), bAscending)
    End If
  End If

End Function

Public Function QuickSortArray(ByVal A As Variant, _
                               Optional Ascending As Boolean = True) As Variant

  On Error GoTo Not_AnArray
  If IsEmpty(A) Then
    QuickSortArray = Split("")
   Else
    QuickSortAD A, LBound(A), UBound(A), Ascending
    QuickSortArray = A
  End If

Exit Function

Not_AnArray:
  QuickSortArray = Split("")

End Function

Public Function QuickSortRemove(ByVal Arr As Variant, _
                                varRemove As Variant, _
                                Optional ByVal bAscending As Boolean = True) As Variant

  If InQSortArray(Arr, varRemove) Then
    Arr(QSortArrayPos(Arr, varRemove)) = vbNullString
    QuickSortRemove = QuickSortArray(NoBlanksArray(Arr), bAscending)
   Else
    QuickSortRemove = Arr
  End If

End Function

Private Function ScopeName(ByVal MScopeNum As Long, _
                           ByVal bStatic As Boolean) As String

  Select Case MScopeNum
   Case vbext_Private '1
    ScopeName = "Private"
    If bStatic Then
      ScopeName = ScopeName & " Static"
    End If
   Case vbext_Public '2
    ScopeName = "Public"
    If bStatic Then
      ScopeName = ScopeName & " Static"
    End If
   Case vbext_Friend '3
    ScopeName = vbNullString
    If bStatic Then
      ScopeName = ScopeName & " Static"
    End If
  End Select

End Function

Private Sub SetDeclareData(DDC As Long, _
                           ByVal strCompName As String, _
                           ByVal StrProjName As String, _
                           ByVal strOrig As String, _
                           ByVal strExpand As String, _
                           arrLine As Variant, _
                           ByVal TypeDef As String, _
                           ByVal HasIndex As Boolean)

  On Error GoTo BugTrap
  '  If Not IsDeclaration(GetName(arrLine), arrLine(0), , strCompName) Then
  With DeclarDesc(DDC)
    .DDComp = strCompName
    .DDProj = StrProjName
    .DDScope = arrLine(0)
    .DDType = TypeDef
    Select Case arrLine(1)
     Case "Enum", "Type"
      .DDType = arrLine(1)
      .DDKind = arrLine(1)
      .DDName = arrLine(2)
      .DDIndexing = False
      arrQEnumTypePresence = QuickSortArray(AppendArray(arrQEnumTypePresence, .DDName))
     Case "Declare"
      If GetLeftBracketPos(arrLine(3)) Then
        arrLine(3) = Left$(arrLine(3), GetLeftBracketPos(arrLine(3)) - 1)
      End If
      .DDKind = "Declare"
      .DDName = arrLine(3)
      .DDIndexing = False
      .DDArguments = Mid$(strOrig, InStr(strOrig, LBracket))
      .DDArguments = Left$(.DDArguments, InStr(.DDArguments, RBracket))
     Case "Const"
      .DDKind = arrLine(1)
      .DDName = arrLine(2)
      .DDIndexing = False
     Case Else
      If GetLeftBracketPos(arrLine(1)) Then
        arrLine(1) = Left$(arrLine(1), GetLeftBracketPos(arrLine(1)) - 1)
      End If
      .DDIndexing = HasIndex
      .DDKind = "Variable"
      If InStr(strExpand, " New ") Then
        .DDClsRef = True
      End If
      .DDName = arrLine(1)
    End Select
    If .DDKind = "Enum" Then
      .DDUsage = VariableUseageArray(.DDName, strOrig)
      If UBound(.DDUsage) = -1 Then
        .DDUsage = EnumUseageArray(.DDName, .DDProj, .DDComp)
      End If
     Else
      .DDUsage = VariableUseageArray(.DDName, strOrig)
    End If
  End With
  DDC = DDC + 1
  ' End If

Exit Sub

BugTrap:
  BugTrapComment "SetDeclareData"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Function SetProcedureData(ByVal PRocNo As Long, _
                                  strCode As String, _
                                  memb As Member, _
                                  Comp As VBComponent, _
                                  Proj As VBProject) As Boolean

  Dim strProc As String

  'v3.0.4 called 2X so use a variable
  If ExtractCode(strCode) Then
    strProc = GetProcKind(strCode)
    If LenB(strProc) Then
      With PRocDesc(PRocNo)
        .PrDClass = strProc
        .PrdComp = Comp.Name
        .PrdCompType = Comp.Type
        If Comp.Type = vbext_ct_ClassModule Then
          .prClassAlias = GetClassAlias(Comp.Name)
        End If
        .PrDProj = Proj.Name
        .PrDScope = ScopeName(memb.Scope, memb.Static)
        .PrDName = memb.Name
        .PrDCallsTo = ProcedureUseageArray(.PrDName)
        .PrDControl = RoutineNameIsVBGenerated(.PrDName, Comp)
        .PrDInternal = RoutineNameIsInternalToModule(.PrDName, Comp)
        .PrDGeneric = RoutineNameIsGenericToModule(.PrDName, Comp)
        If .PrDControl Then
          If Comp.Type = vbext_ct_UserControl Then
            .PrDLegalControl = IsUserControlEvent(.PrDName, Comp)
            If Not .PrDLegalControl Then  'Control on UserControl
              .PrDLegalControl = LegalControlProcedure(.PrDName, .PrdComp)
            End If
           Else
            .PrDLegalControl = LegalControlProcedure(.PrDName, .PrdComp)
          End If
        End If
        .PrDType = ProcType(strCode)
        .PrDAPI = ContainsWholeWord(strCode, "Declare") > 0
        .PrDArguments = GetParamNameString(strCode)
        .PrDArgumentTypes = GetParamTypeString(strCode)
        .PrDUDTParam = NonStandardArg(.PrDArgumentTypes)
      End With
      SetProcedureData = True
    End If
  End If

End Function

Public Function StripDuplicateArray(arrMem As Variant) As Variant

  Dim I           As Long
  Dim J           As Long
  Dim lngNewIndex As Long

  '
  'for some reason some but not all USerControls generate duplicate methods.
  'this strips out the copies and removes blank members of the array
  'comment from Ulli's Code Formatter
  For I = LBound(arrMem) To UBound(arrMem)
    For J = LBound(arrMem) To UBound(arrMem)
      If I <> J Then
        If LenB(arrMem(I)) Then
          If arrMem(I) = arrMem(J) Then
            arrMem(J) = vbNullString
          End If
        End If
      End If
      'v2.4.4 reconfigured to short circuit
      If LenB(arrMem(I)) Then
        If J = LBound(arrMem) Then
          lngNewIndex = lngNewIndex + 1
        End If
      End If
    Next J
  Next I
  ReDim TmpA(lngNewIndex - 1) As Variant
  lngNewIndex = 0
  For I = LBound(arrMem) To UBound(arrMem)
    If LenB(arrMem(I)) Then
      TmpA(lngNewIndex) = arrMem(I)
      lngNewIndex = lngNewIndex + 1
    End If
  Next I
  StripDuplicateArray = TmpA
  Erase arrMem

End Function

Public Sub SwapAnyThing(Var1 As Variant, _
                         Var2 As Variant)

  Dim Var3 As Variant

  Var3 = Var1
  Var1 = Var2
  Var2 = Var3

End Sub

Private Sub TestForDuplicateProcedures()

  Dim I As Long
  Dim J As Long

  'ver 1.1.31 test for same name public procedures in different modules, VB picks them up if in same module
  If bProcDescExists Then
    For I = LBound(PRocDesc) To UBound(PRocDesc)
      MemberMessage PRocDesc(I).PrDName, 1, UBound(PRocDesc)
      For J = LBound(PRocDesc) To UBound(PRocDesc)
        If I <> J Then
          With PRocDesc(I)
            If .PrDName = PRocDesc(J).PrDName Then
              'v 2.1.6 classes and usercontrols can legitimately have same name stuff
              If .PrdCompType <> vbext_ct_UserControl Then
                If .PrdCompType <> vbext_ct_ClassModule Then
                  If PRocDesc(J).PrdCompType <> vbext_ct_UserControl Then
                    If PRocDesc(J).PrdCompType <> vbext_ct_ClassModule Then
                      If .PrDScope = "Public" Then
                        Select Case PRocDesc(J).PrDClass
                         Case "Sub", "Function"
                          'Properties are often legit so ignore them
                          If Not InQSortArray(QSortModClassArray, PRocDesc(J).PrdComp) Then
                            PRocDesc(J).PrDDuplicate = True
                            .PrDDuplicate = True
                          End If
                        End Select
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End With 'PRocDesc(I)
        End If
      Next J
    Next I
  End If

End Sub



':)Code Fixer V3.0.9 (25/03/2005 4:12:17 AM) 105 + 1329 = 1434 Lines Thanks Ulli for inspiration and lots of code.

