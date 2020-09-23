Attribute VB_Name = "mod_ReadWrite"
Option Explicit
Public Enum ReWriteType
  RWCode
  RWDeclaration
  RWMembers
  RWModule
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private RWCode, RWDeclaration, RWMembers, RWModule
#End If
Public SortElems()     As Variant

Public Function BlankLineStripper(VarInput As Variant) As String

  Dim strTemp         As String

  strTemp = Join(VarInput, vbNewLine)
  Do While InStr(strTemp, vbNewLine & vbNewLine)
    strTemp = Replace$(strTemp, vbNewLine & vbNewLine, vbNewLine)
  Loop
  BlankLineStripper = strTemp

End Function

Public Function FullMemberExtraction(cMod As CodeModule) As Variant

  Dim J          As Long
  Dim I          As Long
  Dim Tmpstring1 As String
  Dim K          As Long
  Dim TempElem   As Variant
  Dim strTmp     As String

  'mObjDoc.SuspendErrorTrap
  'ver 1.0.93 simplified coding
  With cMod
    If Xcheck(XVisScan) Then
      .CodePane.TopLine = 1
    End If
    ReDim SortElems(0)
    'collect module descriptions -> (Name, StartingLine, Length)
    For J = 1 To .Members.Count
      With .Members(J)
        Tmpstring1 = .Name
        I = (.Type = vbext_mt_Property Or .Type = vbext_mt_Method)
      End With
      If I Then
        For I = 1 To 4
          K = Choose(I, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set, vbext_pk_Proc)
          'determines seq of equal named modules
          TempElem = Null
          On Error Resume Next
          'crash here if debug is all
          TempElem = Array(Tmpstring1, .PRocStartLine(Tmpstring1, K), .ProcCountLines(Tmpstring1, K))
          On Error GoTo 0
          If Not IsNull(TempElem) Then
            ReDim Preserve SortElems(UBound(SortElems) + 1)
            SortElems(UBound(SortElems)) = TempElem
          End If
        Next I
      End If
    Next J
    'build sorted component
    Tmpstring1 = vbNullString
    For I = 1 To UBound(SortElems)
      Select Case I
       Case 1
        'Sub or Function
        If SortElems(I)(1) > .CountOfDeclarationLines Then
          Tmpstring1 = Tmpstring1 & .Lines(SortElems(I)(1), SortElems(I)(2)) & ArrayDummySplitPoint
        End If
       Case Else
        'there's a quirk in VB: it returns Events as methods and if an
        'Event has the same name as a Sub/Function then this results in
        'duplicates, so here duplicates are filtered out
        If SortElems(I)(1) <> SortElems(I - 1)(1) Then
          'Sub or Function
          If SortElems(I)(1) > .CountOfDeclarationLines Then
            strTmp = .Lines(SortElems(I)(1), SortElems(I)(2))
            If InStr(Tmpstring1, strTmp) = 0 Then
              Tmpstring1 = Tmpstring1 & strTmp & ArrayDummySplitPoint
            End If
          End If
        End If
      End Select
    Next I
  End With
  FullMemberExtraction = Split(Tmpstring1, ArrayDummySplitPoint)
  'mObjDoc.ResumeErrorTrap

End Function

Public Function GetDeclarationArray(cMod As CodeModule) As Variant

  Dim Tmpstring1 As String

  'Gets the Declaration section as an array
  With cMod
    'fix me only do this after checking if #If occurs in the module at all
    'v2.3.8 fix so that it uses short method if optional compile not needed
    'v2.3.9 if the there is an opt compile at end get it and rewrite so long way doesn't hit again
    ' and adds the necessary guard if needed
    If NeedOptionalCompileFix(.Parent.Name) Then
      'End of declaration Optional Compile (#If X Then/#End If)
      'to be fully read as part of declaration section
      Tmpstring1 = .Lines(1, .CountOfDeclarationLines) 'get most of the Declarations
      .DeleteLines 1, .CountOfDeclarationLines
      Do While CountSubStringImbalance(Tmpstring1, vbNewLine & "#If", vbNewLine & "#End If")
        If Not isProcHead(.Lines(1, 1)) Then
          ' collect lines until balanced or is proc head
          ' which means that it is a case of Whole Code in Opt Compile
          Tmpstring1 = Tmpstring1 & vbNewLine & Trim$(.Lines(1, 1))
          ' delete '#End If that belongs in Declare but VB sees as part of 1st proc
          .DeleteLines 1
         Else
          Exit Do
        End If
      Loop
      'v 2.0.7 Thanks Paul Caton This trap allows a whole code enclosed in Optional compilation
      'to be processed properly while still protecting an end of declartion section
      ' Optional compilation to be fully in the declaration
      Tmpstring1 = Tmpstring1 & vbNewLine & CodeFixProtectedArray(endDec)
      'rewrite the Declaration sect
      .InsertLines 1, Tmpstring1
      Tmpstring1 = vbNullString
    End If
    If .Parent.Type <> vbext_ct_UserControl Then
      InsertOptionExplicit cMod
    End If
    'now if fixed or not needed
    If .CountOfDeclarationLines Then
      Tmpstring1 = .Lines(1, .CountOfDeclarationLines)
      GetDeclarationArray = Split(.Lines(1, .CountOfDeclarationLines), vbNewLine)
     Else
      GetDeclarationArray = Split("")
    End If
  End With

End Function

Public Sub InsertOptionExplicit(cMod As CodeModule)

  'v3.0.5 simplified

  If Not bNoOptExp Then 'v3.0.0 Block fix in Format only thanks Ian K
    With cMod
      If .CountOfDeclarationLines = 0 Then
        'v2.9.5 changed message layout, commment now above code
        .InsertLines 1, IIf(Xcheck(XVerbose), OptExplicitMsgVerbose, OptExplicitMsg) & vbNewLine & _
         "Option Explicit"
       ElseIf Not FindCodeUsage("Option Explicit", "", .Parent.Name, True, True, False) Then
        .InsertLines 1, IIf(Xcheck(XVerbose), OptExplicitMsgVerbose, OptExplicitMsg) & vbNewLine & _
         "Option Explicit"
      End If
    End With
  End If

End Sub

Private Function NeedOptionalCompileFix(strModName As String) As Boolean

  Dim A As Long
  Dim B As Long

  'v2.3.8 fast way to decide whether to apply the Fix (slow line by line read of file or get declaration simply way
  ' Thanks Paul Caton found this while working on the bracket problem
  If FindCodeUsage("#If", "", strModName, True, True, False) Then
    A = CountCodeUsage("#If", vbNullString, strModName, True, True, False)
    B = CountCodeUsage("#End If", vbNullString, strModName, True, True, False)
    NeedOptionalCompileFix = A <> B
  End If

End Function

Public Sub ReWriter(cMod As CodeModule, _
                    CodArr As Variant, _
                    ByVal Mode As ReWriteType, _
                    Optional ByVal DeleteBlanks As Boolean = True)

  Dim strTemp As String

  On Error GoTo BugHit
  'delete Blank lines introduced by calling routine
  If DeleteBlanks Then
    '*if DeleteBlanks = True then delete blank lines introduced by calling routine
    'Some routines introduce blank lines by deleting lines
    '(usually to move the codearray member to another position)
    'This routine can delete those blanks before rewriting the module
    strTemp = BlankLineStripper(CodArr)
   Else
    strTemp = Join(CodArr, vbNewLine)
  End If
  'this turns Code Fixer back on if it was disabled for any lines
  With cMod
    Select Case Mode
     Case RWCode
      .DeleteLines .CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1 - 1
      .AddFromString strTemp
     Case RWDeclaration
      .DeleteLines 1, .CountOfDeclarationLines
      'v3.0.5 newline at end should keep final declaration comment in declaration when Proc Sorting occurs
      .AddFromString strTemp & vbNewLine
     Case RWModule
      .DeleteLines 1, .CountOfLines
      .AddFromString strTemp
     Case RWMembers
      .DeleteLines 1, .CountOfLines
      .AddFromString strTemp
    End Select
  End With
  strTemp = vbNullString
  Safe_Sleep
  'DoEvents
  On Error GoTo 0

Exit Sub

BugHit:
  BugTrapComment "ReWriter"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If
  '   Case 40192 ' "Too many line continuations"

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:17:16 AM) 11 + 214 = 225 Lines Thanks Ulli for inspiration and lots of code.

