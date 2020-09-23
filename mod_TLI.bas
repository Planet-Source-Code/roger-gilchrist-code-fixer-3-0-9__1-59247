Attribute VB_Name = "mod_TLI"
Option Explicit
'WARNING THIS MODULE IS ONLY PARTIALLY DEVELOPED SO IF YOU BORROW IT BE CAREFUL
'*IF YOU MAKE ANY INTERESTING MODIFICATIONS LET ME KNOW
Private ReferenceMembers      As Variant
Private tliTypeLibInfo        As TypeLibInfo
Private tliTypeInfo           As TypeInfo

Public Function GenerateReferencesEnumArray() As Boolean

  'v3.0.6 updated to abort fix/format if there is a missing reference Thanks Ian K
  
  Dim Proj   As VBProject
  Dim strLib As String
  Dim I      As Long

  'get the references and stick them in an array to drive the outer 'For' structure of
  'the references tests
  GenerateReferencesEnumArray = True ' assume it will work
  For Each Proj In VBInstance.VBProjects
    For I = 1 To Proj.References.Count
      On Error Resume Next
      'the Reference collection may contain Common control references due to their presense in the toolbox
      'but generate the '<Application-defined or object-defined error>' if none of the controls
      'are used in project
      If Proj.References.Item(I).IsBroken Then
        mObjDoc.Safe_MsgBox "Code Fixer will stop processing." & vbNewLine & _
                    "There is at least one missing reference." & vbNewLine & _
                    "Please resolve the problem before trying again." & vbNewLine & _
                    "Open Project|References... menu and check for references marked 'MISSING:'" & vbNewLine & _
                    "You may have a different version of the reference available, uncheck the missing one and scroll down to find it." & vbNewLine & _
                    "WARNING Some Properties/Methods of your version may be incompatible with the coder's version." & vbNewLine & _
                    "Check documentation for possible replacments." & vbNewLine & _
                    "E.G. Word 10 '.Documents.Add DocumentType:=wdNewBlankDocument' becomes Word 8 '.Documents.Add NewTemplate:=True'" & vbNewLine & _
                    "If you do not have a different version (or it is highly incompatible) contact the orignal coder for help.", vbCritical + vbOKOnly

        GenerateReferencesEnumArray = False
        GoTo SafeExit
       Else
        strLib = AccumulatorString(strLib, Proj.References.Item(I).FullPath, "|")
      End If
    Next I
  Next Proj
  '
  strLib = AccumulatorString(strLib, "vba6.dll", "|")
  'this holds such stuff as vbColor values but is not listed explicitly in reference tool so add it
  'NOTE lack of full path is not a problem; If the reference exists the code can access from filename alone
  'I could have striped the pathes off all of them but it is a bit messy because of the \3 element
  'that distinguishes VBA="msvbvm60.dll"  from VBRUN ="msvbvm60.dll\3"
  'I have to assume that other references may also use this or similar
  'and allow for it by just getting the full path
  FillArray ReferenceMembers, strLib, False, "|"
SafeExit:
  On Error GoTo 0

End Function

Public Function GetRefLibKnownVBConstantType(varTest As Variant) As String

  Dim INTERFACE As tli.Constants
  Dim I         As Long
  Dim J         As Long

  'Get value of a known VB Constant
  Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      If varTest = INTERFACE.Item(J).Members.Item(I).Name Then
        GetRefLibKnownVBConstantType = RefLibType2CodeType(INTERFACE.Item(J).Members.Item(I).ReturnType.VarType)
        Exit For
      End If
    Next I
    If LenB(GetRefLibKnownVBConstantType) Then
      Exit For
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Public Function GetRefLibVBCommandsReturnType(varTest As Variant) As String

  Dim INTERFACE As Declarations
  Dim strMName  As String
  Dim I         As Long
  Dim J         As Long

  'detect that word is a VBCommand
  Set INTERFACE = tli.TypeLibInfoFromFile("vba6.dll").Declarations
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      strMName = INTERFACE.Item(J).Members(I).Name
      If InStr(strMName, "_") Then ' cope with _B_str_ & _B_var_
        strMName = strGetRightOf(strMName, "_")
      End If
      If varTest = strMName Then
        GetRefLibVBCommandsReturnType = RefLibType2CodeType(INTERFACE.Item(J).Members.Item(I).ReturnType.VarType)
        Exit For
      End If
    Next I
    If LenB(GetRefLibVBCommandsReturnType) Then
      Exit For
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Public Sub InitializeTLib()

  Set tliTypeLibInfo = New TypeLibInfo
  tliTypeLibInfo.AppObjString = "<Global>"

End Sub

Public Function isReadWriteProperty(strCtrlClass As String, _
                                    ByVal strPropertyName As String, _
                                    Optional ByVal bForm As Boolean = False) As Boolean

  Dim TestCount As Long
  Dim INTERFACE As tli.Interfaces
  Dim I         As Long
  Dim J         As Long
  Dim strLib    As String
  Dim arrTest   As Variant

  strLib = RefLibContainsControldata(strCtrlClass)
  If LenB(strLib) Then
    Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
    arrTest = Array(tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT)
    For I = 1 To INTERFACE.Count
      If Right$(INTERFACE.Item(I).Name, Len(strCtrlClass)) = strCtrlClass Then
        For J = 1 To INTERFACE.Item(I).Members.Count
          If IsInArray(INTERFACE.Item(I).Members.Item(J).InvokeKind, arrTest) Then
            If INTERFACE.Item(I).Members.Item(J).Name = strPropertyName Then
              TestCount = TestCount + 1
              If TestCount > IIf(bForm, 2, 1) Then
                isReadWriteProperty = True
                GoTo CleanExit ': Exit Function
              End If
            End If
          End If
        Next J
      End If
    Next I
  End If
CleanExit:
  Set INTERFACE = Nothing

End Function

Public Function isRefLibKnownVBConstant(varTest As Variant) As Boolean

  Dim INTERFACE As tli.Constants
  Dim I         As Long
  Dim J         As Long

  'test vattest is knonw VB Constant
  Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      If varTest = INTERFACE.Item(J).Members.Item(I).Name Then
        isRefLibKnownVBConstant = True
        Exit For 'unction
      End If
    Next I
    If isRefLibKnownVBConstant Then
      Exit For
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Public Function IsRefLibKnownVBConstantType(varTest As Variant) As Boolean

  Dim INTERFACE As tli.Constants
  Dim J         As Long

  'Get value of a known VB Constant
  Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
  For J = 1 To INTERFACE.Count
    If varTest = INTERFACE.Item(J).Name Then
      IsRefLibKnownVBConstantType = True
      Exit For 'unction
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Public Function isRefLibVBCommands(varTest As Variant, _
                                   Optional ByVal CaseAware As Boolean = True) As Variant

  Dim INTERFACE As Declarations
  Dim strMName  As String
  Dim I         As Long
  Dim J         As Long

  'detect that word is a VBCommand
  Set INTERFACE = tli.TypeLibInfoFromFile("vba6.dll").Declarations
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      strMName = INTERFACE.Item(J).Members(I).Name
      If InStr(strMName, "_") Then ' cope with _B_str_ & _B_var_
        strMName = strGetRightOf(strMName, "_")
      End If
      If CaseAware Then
        If varTest = strMName Then
          isRefLibVBCommands = True
          Exit For 'unction
        End If
       Else
        If LCase$(varTest) = LCase$(strMName) Then
          isRefLibVBCommands = True
          Exit For 'unction
        End If
      End If
    Next I
    If isRefLibVBCommands Then
      Exit For
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Private Function isRefLibVBCommandsStrVar(varTest As Variant) As Boolean

  Dim INTERFACE As Declarations
  Dim strMName  As String
  Dim I         As Long
  Dim J         As Long
  Dim lFound    As Long

  'detect that string and variant forms of a Command exist
  'support for RefLibGenerateVBCommandStrVarArray
  Set INTERFACE = tli.TypeLibInfoFromFile("vba6.dll").Declarations
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      strMName = INTERFACE.Item(J).Members(I).Name
      If InStr(strMName, "_") Then ' cope with _B_str_ & _B_var_
        If strMName = "_B_str_" & varTest Then
          lFound = lFound + 1
        End If
        If strMName = "_B_var_" & varTest Then
          lFound = lFound + 1
        End If
        If lFound = 2 Then
          isRefLibVBCommandsStrVar = True
          Exit For 'unction
        End If
      End If
    Next I
    If isRefLibVBCommandsStrVar Then
      Exit For
    End If
  Next J
  Set INTERFACE = Nothing

End Function

Private Function ProduceDefaultValue(DefVal As Variant, _
                                     ByVal TI As TypeInfo) As String

  Dim lTrackVal As Long
  Dim MI        As MemberInfo
  Dim TKind     As TypeKinds

  If TI Is Nothing Then
    Select Case VarType(DefVal)
     Case vbString
      If Len(DefVal) Then
        ProduceDefaultValue = """" & DefVal & """"
      End If
     Case vbBoolean 'Always show for Boolean
      ProduceDefaultValue = DefVal
     Case vbDate
      If DefVal Then
        ProduceDefaultValue = "#" & DefVal & "#"
      End If
     Case Else 'Numeric Values
      If DefVal <> 0 Then
        ProduceDefaultValue = DefVal
       Else
        ProduceDefaultValue = VarType(DefVal)
      End If
    End Select
   Else
    'See if we have an enum and track the matching member
    'If the type is an object, then there will never be a
    'default value other than Nothing
    TKind = TI.TypeKind
    Do While TKind = TKIND_ALIAS
      TKind = TKIND_MAX
      On Error Resume Next
      Set TI = TI.ResolvedType
      If Err.Number = 0 Then
        TKind = TI.TypeKind
      End If
      On Error GoTo 0
    Loop
    If TI.TypeKind = TKIND_ENUM Then
      lTrackVal = DefVal
      For Each MI In TI.Members
        If MI.Value = lTrackVal Then
          ProduceDefaultValue = MI.Name
          Exit For
        End If
      Next MI
    End If
  End If

End Function

Public Function ReferenceLibraryConstant(ByVal strTest As String) As Boolean

  Dim CNST As tli.Constants
  Dim I    As Long
  Dim K    As Long

  'v2.8.6 Thanks Ian K for spotting the need for this
  'Finds all constants in reference libraries
  '"msvbvm60.dll" = VBA
  ', "msvbvm60.dll\3") Then 'VBRUN
  ', "vb6.olb") Then 'Visual Basic objects and procedures
  '', "Comdlg32.ocx")
  For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
    Set CNST = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Constants
    For I = 1 To CNST.Count
      If CNST.Item(I).Name = strTest Then
        ReferenceLibraryConstant = True
        Exit For
      End If
    Next I
    If ReferenceLibraryConstant Then
      Exit For
    End If
  Next K

End Function

Public Function ReferenceLibraryControlDefaultProperty(ByVal strCtrlClass As String) As String

  Dim INTERFACE     As tli.Interfaces
  Dim IFACE_MEMB    As tli.MemberInfo
  Dim I             As Long
  Dim J             As Long
  Dim M             As Long
  Dim strLib        As String
  Dim LVTableOffset As String
  Dim arrTest       As Variant

  'the '_Default' method of a control has the same VTableOffset as the method itself
  'SO find a '_Default' method and get its offset
  'Search the offsets for matching offset (and ignore the 2nd'_Default" PROPERTYPUT mmember)
  'Return Default Property of control
  strLib = RefLibContainsControldata(strCtrlClass)
  If LenB(strLib) Then
    Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
    arrTest = Array(tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT, tli.INVOKE_PROPERTYPUTREF)
    For I = 1 To INTERFACE.Count
      If Right$(INTERFACE.Item(I).Name, Len(strCtrlClass)) = strCtrlClass Then
        For J = 1 To INTERFACE.Item(I).Members.Count
          Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(J)
          If IsInArray(IFACE_MEMB.InvokeKind, arrTest) Then
            If IFACE_MEMB.Name = "_Default" Then
              LVTableOffset = IFACE_MEMB.VTableOffset
              For M = J To INTERFACE.Item(I).Members.Count
                Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(M)
                If LVTableOffset = IFACE_MEMB.VTableOffset Then
                  If IFACE_MEMB.Name <> "_Default" Then
                    ReferenceLibraryControlDefaultProperty = IFACE_MEMB.Name
                    GoTo CleanExit 'Exit Function
                  End If
                End If
              Next M
            End If
          End If
        Next J
      End If
    Next I
  End If
CleanExit:
  Set INTERFACE = Nothing

End Function

Public Function ReferenceLibraryControlProperty(ByVal strTest As String, _
                                                ByVal strCtrlClass As String, _
                                                Optional strType As String) As Boolean

  Dim INTERFACE As tli.Interfaces
  Dim I         As Long
  Dim J         As Long
  Dim strLib    As String
  Dim arrTest   As Variant

  strType = vbNullString
  strLib = RefLibContainsControldata(strCtrlClass)
  If LenB(strLib) Then
    Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
    arrTest = Array(tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT, tli.INVOKE_PROPERTYPUTREF)
    For I = 1 To INTERFACE.Count
      If Right$(INTERFACE.Item(I).Name, Len(strCtrlClass)) = strCtrlClass Then
        For J = 1 To INTERFACE.Item(I).Members.Count
          If IsInArray(INTERFACE.Item(I).Members.Item(J).InvokeKind, arrTest) Then
            If INTERFACE.Item(I).Members.Item(J).Name = strTest Then
              strType = RefLibType2CodeType(INTERFACE.Item(I).Members.Item(J).ReturnType.VarType)
              ReferenceLibraryControlProperty = True
              Exit For 'unction
             ElseIf LCase$(INTERFACE.Item(I).Members.Item(J).Name) = LCase$(strTest) Then
              'this will reset the string to correct case
              strTest = INTERFACE.Item(I).Members.Item(J).Name
              strType = RefLibType2CodeType(INTERFACE.Item(I).Members.Item(J).ReturnType)
              ReferenceLibraryControlProperty = True
              Exit For 'unction
            End If
          End If
        Next J
      End If
      If ReferenceLibraryControlProperty Then
        Exit For
      End If
    Next I
  End If
  Set INTERFACE = Nothing

End Function

Private Function RefLibContainsControldata(ByVal strCtrlClass As String) As String

  Dim INTERFACE As CoClasses
  Dim I         As Long
  Dim J         As Long

  'get library containing the control data
  For I = LBound(ReferenceMembers) To UBound(ReferenceMembers)
    On Error GoTo oops
    Set INTERFACE = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(I))).CoClasses
    For J = 1 To INTERFACE.Count
      If INTERFACE.Item(J).Name = strCtrlClass Then
        RefLibContainsControldata = ReferenceMembers(I)
        Exit For 'unction
      End If
    Next J
    If LenB(RefLibContainsControldata) Then
      Exit For
    End If
  Next I
oops:
  Set INTERFACE = Nothing

End Function

Public Function RefLibGenerateVBCommandArray() As Variant

  Dim INTERFACE As Declarations
  Dim strTmp    As String
  Dim strMName  As String
  Dim I         As Long
  Dim J         As Long

  'create array of ArrQVBCommands
  Set INTERFACE = tli.TypeLibInfoFromFile("vba6.dll").Declarations
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      strMName = INTERFACE.Item(J).Members(I).Name
      If InStr(strMName, "_") Then
        strMName = strGetRightOf(strMName, "_")
      End If
      strTmp = AccumulatorString(strTmp, strMName)
    Next I
  Next J
  FillArray RefLibGenerateVBCommandArray, strTmp
CleanExit:
  Set INTERFACE = Nothing
  strTmp = vbNullString

End Function

Public Function RefLibGenerateVBCommandStrVarArray() As Variant

  Dim INTERFACE As Declarations
  Dim strTmp    As String
  Dim strMName  As String
  Dim I         As Long
  Dim J         As Long

  'creates an array containing all the commands that can take string & variant forms
  'ie Chr/LEft/Mid etc
  Set INTERFACE = tli.TypeLibInfoFromFile("vba6.dll").Declarations
  For J = 1 To INTERFACE.Count
    For I = 1 To INTERFACE.Item(J).Members.Count
      strMName = INTERFACE.Item(J).Members(I).Name
      If InStr(strMName, "_") Then
        'get rightmost part of the string which is in format X_XXX_CommandName
        strMName = StrReverse(Left$(StrReverse(strMName), InStr(StrReverse(strMName), "_") - 1))
        If isRefLibVBCommandsStrVar(strMName) Then
          strTmp = AccumulatorString(strTmp, strMName)
        End If
      End If
    Next I
  Next J
  FillArray RefLibGenerateVBCommandStrVarArray, strTmp
CleanExit:
  Set INTERFACE = Nothing
  strTmp = vbNullString

End Function

Private Function RefLibType2CodeType(ByVal strRLType As Long) As String

  Select Case strRLType
   Case VT_BSTR, VT_LPWSTR, VT_LPSTR
    RefLibType2CodeType = "String"
   Case VT_I1
    RefLibType2CodeType = "Byte"
   Case VT_I2
    RefLibType2CodeType = "Integer"
   Case VT_I4, VT_INT
    RefLibType2CodeType = "Long"
   Case VT_R4
    RefLibType2CodeType = "Single"
   Case VT_R8, VT_I8
    RefLibType2CodeType = "Double"
   Case VT_BOOL
    RefLibType2CodeType = "Boolean"
   Case VT_ARRAY, VT_VARIANT, VT_CARRAY
    RefLibType2CodeType = "Variant"
   Case VT_DATE
    RefLibType2CodeType = "Date"
   Case VT_CY, VT_DECIMAL
    RefLibType2CodeType = "Currency"
   Case VT_UI4
    RefLibType2CodeType = "OLE_COLOR"
   Case Else
    RefLibType2CodeType = "Variant' Default Value; May not be correct but will work"
  End Select
  'char [(n)] (1 = n = 255) DBTYPE_STR
  'varchar [(n)] (1 = n = 255) DBTYPE_STR
  'binary [(n)] (1 = n = 255) DBTYPE_BYTES
  'varbinary [(n)] (1 = n = 255) DBTYPE_BYTES
  'numeric [(p[,s])] DBTYPE_NUMERIC
  'decimal [(p[,s])] DBTYPE_NUMERIC
  'Tinyint DBTYPE_UI1
  'SmallInt DBTYPE_I2
  'Int DBTYPE_I4
  'Real DBTYPE_R4
  'float [(n)] DBTYPE_R8
  'Smalldatetime DBTYPE_DATE, DBTYPE_DBTIMESTAMP
  'DateTime DBTYPE_DATE, DBTYPE_DBTIMESTAMP
  'Timestamp DBTYPE_BYTES (DBCOLUMNFLAGS_ISROWVER is set)
  'text (= 2**31 bytes) DBTYPE_STR
  'image (= 2**31 bytes) DBTYPE_BYTES
  'Smallmoney DBTYPE_CY
  'Money DBTYPE_CY
  'user-defined-type DBTYPE_UDT
  '
  'DBTYPE_I2 Integer data stored in 2 bytes (16 bits).
  'DBTYPE_I4 Integer data stored in 4 bytes (32 bits).
  'DBTYPE_R4 Single precision IEEE floating-point data stored in 4 bytes (32 bits).
  'DBTYPE_R8 Double precision floating-point data stored in 8 bytes (64 bits).
  ''Case VT_I1
  ''Case VT_I2
  ''Case VT_I4
  ''Case VT_I8
  ''Case VT_INT
  ''Case VT_R4
  ''Case VT_R8
  ''Case VT_BLOB
  ''Case VT_BLOB_OBJECT
  ''Case VT_BYREF
  ''Case VT_CARRAY
  ''Case VT_CF
  ''Case VT_CLSID
  ''Case VT_CY
  ''Case VT_DISPATCH
  ''Case VT_EMPTY
  ''Case VT_ERROR
  ''Case VT_FILETIME
  ''Case VT_HRESULT
  ''Case VT_INT
  ''Case VT_LPSTR
  ''Case VT_LPWSTR
  ''Case VT_NULL
  ''Case VT_PTR
  ''Case VT_I1
  ''Case VT_I2
  ''Case VT_I4
  ''Case VT_I8
  ''Case VT_INT
  ''Case VT_R4
  ''Case VT_R8
  ''Case VT_RECORD
  ''Case VT_RESERVED
  ''Case VT_SAFEARRAY
  ''Case VT_STORAGE
  ''Case VT_STORED_OBJECT
  ''Case VT_STREAM
  ''Case VT_STREAMED_OBJECT
  ''Case VT_UI1
  ''Case VT_UI2
  ''Case VT_UI4
  ''Case VT_UI8
  ''Case VT_UINT
  ''Case VT_UNKNOWN
  ''Case VT_USERDEFINED
  ''Case VT_VECTOR
  ''Case VT_VOID

End Function

Public Function RefLibVBColorConstFromValue(lngColour As Long) As Variant

  Dim INTERFACE As tli.Constants
  Dim I         As Long
  Dim J         As Long

  'If a colour value matches any of the vb color Constants
  'returns the Constant otherwise returns the value
  RefLibVBColorConstFromValue = lngColour ' set to long value by default
  Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
  For J = 1 To INTERFACE.Count
    If INTERFACE.Item(J).Name = "ColorConstants" Or INTERFACE.Item(J).Name = "SystemColorConstants" Then
      For I = 1 To INTERFACE.Item(J).Members.Count
        If INTERFACE.Item(J).Members.Item(I).Value = lngColour Then
          RefLibVBColorConstFromValue = INTERFACE.Item(J).Members.Item(I).Name
          GoTo CleanExit '        Exit Function
        End If
      Next I
    End If
  Next J
CleanExit:
  Set INTERFACE = Nothing

End Function

Public Function TLibEventFinder(ByVal strTest As String, _
                                ByVal strCtrlClass As String, _
                                Optional strResetCase As String) As Boolean

  Dim I   As Long
  Dim tli As Interfaces

  On Error GoTo oops
  With tliTypeLibInfo
    .ContainingFile = RefLibContainsControldata(strCtrlClass)
    Set tliTypeInfo = .GetTypeInfo(strCtrlClass)
    Set tli = .CoClasses.NamedItem(strCtrlClass).Interfaces
  End With 'tliTypeLibInfo
  For I = 1 To tli.Item(2).Members.Count
    If tli.Item(2).Members(I).Name = strTest Then
      TLibEventFinder = True
      strResetCase = strTest
      Exit For
    End If
    If LCase$(tli.Item(2).Members(I).Name) = LCase$(strTest) Then
      TLibEventFinder = True
      strResetCase = tli.Item(2).Members(I).Name
      Exit For
    End If
  Next I
oops:

End Function

Public Function TLibMethodFinder(ByVal strTest As String, _
                                 ByVal strCtrlClass As String, _
                                 Optional strResetCase As String) As Boolean

  Dim I   As Long
  Dim tli As Interfaces

  On Error GoTo oops
  If LenB(RefLibContainsControldata(strCtrlClass)) Then
    With tliTypeLibInfo
      .ContainingFile = RefLibContainsControldata(strCtrlClass)
      Set tliTypeInfo = .GetTypeInfo(strCtrlClass)
      Set tli = .CoClasses.NamedItem(strCtrlClass).Interfaces
    End With 'tliTypeLibInfo
    For I = 1 To tli.Item(1).Members.Count
      If tli.Item(1).Members(I).Name = strTest Then
        TLibMethodFinder = True
        strResetCase = strTest
        Exit For
      End If
      If LCase$(tli.Item(1).Members(I).Name) = LCase$(strTest) Then
        TLibMethodFinder = True
        strResetCase = tli.Item(1).Members(I).Name
        Exit For
      End If
    Next I
  End If
oops:

End Function

Public Function TypeOfProperty(strCtrlClass As String, _
                               ByVal strPropertyName As String) As Variant

  Dim INTERFACE As tli.Interfaces
  Dim I         As Long
  Dim J         As Long
  Dim strLib    As String
  Dim TI        As TypeInfo
  Dim arrTest   As Variant

  strLib = RefLibContainsControldata(strCtrlClass)
  If LenB(strLib) Then
    Set TI = TypeLibInfoFromFile(strLib).GetTypeInfo(strCtrlClass)
    Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
    arrTest = Array(tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT)
    For I = 1 To INTERFACE.Count
      If Right$(INTERFACE.Item(I).Name, Len(strCtrlClass)) = strCtrlClass Then
        For J = 1 To INTERFACE.Item(I).Members.Count
          If IsInArray(INTERFACE.Item(I).Members.Item(J).InvokeKind, arrTest) Then
            If INTERFACE.Item(I).Members.Item(J).Name = strPropertyName Then
              TypeOfProperty = RefLibType2CodeType(ProduceDefaultValue(INTERFACE.Item(I).Members.Item(J).ReturnType, Nothing))
              Exit For 'unction
            End If
          End If
        Next J
      End If
      If LenB(TypeOfProperty) Then
        Exit For
      End If
    Next I
  End If
  Set INTERFACE = Nothing

End Function

Public Function UsingScripting() As Boolean

  Dim I As Long

  'v 2.2.2 activated
  For I = LBound(ReferenceMembers) To UBound(ReferenceMembers)
    If InStr(LCase$(ReferenceMembers(I)), "scrrun.dll") Then
      UsingScripting = True
      Exit For
    End If
  Next I

End Function

''Public Function TLibPropertyFinder(ByVal strTest As String, ByVal strCtrlClass As String) As Boolean
''
''
''Dim tli As Interfaces
''Dim I   As Long
''If Len(RefLibContainsControldata(strCtrlClass)) Then
''With tliTypeLibInfo
''.ContainingFile = RefLibContainsControldata(strCtrlClass)
''Set tliTypeInfo = .GetTypeInfo(strCtrlClass)
''Set tli = .CoClasses.NamedItem(strCtrlClass).Interfaces
''End With 'tliTypeLibInfo
''For I = 1 To tli.Item(2).Members.Count
''If tli.Item(1).Members(I).Name = strTest Then
''TLibPropertyFinder = True
''Exit For
''End If
''Next I
''End If
''End Function
''Public Function BuildSearchData(ByVal TypeInfoNumber As Integer, ByVal SearchTypes As TliSearchTypes, Optional ByVal LibNum As Integer, Optional ByVal Hidden As Boolean = False) As Long
''If SearchTypes And &H80 Then
''BuildSearchData = (TypeInfoNumber And &H1FFF) Or ((SearchTypes And &H7F) * &H1000000) Or &H80000000
''Else
''BuildSearchData = (TypeInfoNumber And &H1FFF) Or (SearchTypes * &H1000000)
''End If
''If LibNum Then
''BuildSearchData = BuildSearchData Or ((LibNum And &HFF) * &H10000) Or ((LibNum And &H700) * &H20)
''End If
''If Hidden Then
''BuildSearchData = BuildSearchData Or &H1000
''End If
''End Function
''
''Public Function GetHidden(ByVal SearchData As Long) As Boolean
''If SearchData And &H1000 Then
''GetHidden = True
''End If
''End Function
''
''Public Function GetLibNum(ByVal SearchData As Long) As Integer
''SearchData = SearchData And &H7FFFFFFF
''GetLibNum = ((SearchData \ &H2000 And &H7) * &H100) Or (SearchData \ &H10000 And &HFF)
''End Function
''
''Public Function GetTypeInfoNumber(ByVal SearchData As Long) As Integer
''GetTypeInfoNumber = SearchData And &HFFF
''End Function
''
'Public Function PrototypeMember(TLInf As TypeLibInfo, ByVal SearchData As Long, ByVal InvokeKinds As InvokeKinds, Optional ByVal MemberId As Long = -1, Optional ByVal MemberName As String) As String
'Dim pi              As ParameterInfo
'Dim fFirstParameter As Boolean
'Dim fIsConstant     As Boolean
'Dim fByVal          As Boolean
'Dim retVal          As String
'Dim ConstVal        As Variant
'Dim strTypeName     As String
'Dim VarTypeCur      As Long
'Dim fDefault        As Boolean
'Dim fOptional       As Boolean
'Dim fParamArray     As Boolean
'Dim TIType          As TypeInfo
'Dim TIResolved      As TypeInfo
'Dim TKind           As TypeKinds
'With TLInf
'fIsConstant = GetSearchType(SearchData) And tliStConstants
'With .GetMemberInfo(SearchData, InvokeKinds, MemberId, MemberName)
'If fIsConstant Then
'retVal = "Const "
'ElseIf InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then
'Select Case .ReturnType.VarType
'Case VT_VOID, VT_HRESULT
'retVal = "Sub "
'Case Else
'retVal = "Function "
'End Select
'Else
'retVal = "Property "
'End If
'retVal = retVal & .Name
'With .Parameters
'If .Count Then
'retVal = retVal & "("
'fFirstParameter = True
'fParamArray = .OptionalCount = -1
'For Each pi In .Me
'If Not fFirstParameter Then
'retVal = retVal & ", "
'End If
'fFirstParameter = False
'fDefault = pi.Default
'fOptional = fDefault Or pi.Optional
'If fOptional Then
'If fParamArray Then
''This will be the only optional parameter
'retVal = retVal & "[ParamArray "
'Else
'retVal = retVal & "["
'End If
'End If
'With pi.VarTypeInfo
'Set TIType = Nothing
'Set TIResolved = Nothing
'TKind = TKIND_MAX
'VarTypeCur = .VarType
'If (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
''If Not .TypeInfoNumber Then 'This may error, don't use here
'On Error Resume Next
'Set TIType = .TypeInfo
'If Not TIType Is Nothing Then
'Set TIResolved = TIType
'TKind = TIResolved.TypeKind
'Do While TKind = TKIND_ALIAS
'TKind = TKIND_MAX
'Set TIResolved = TIResolved.ResolvedType
'If Err Then
'Err.Clear
'Else
'TKind = TIResolved.TypeKind
'End If
'Loop
'End If
'Select Case TKind
'Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
'fByVal = .PointerLevel = 1
'Case TKIND_RECORD
''Records not passed ByVal in VB
'fByVal = False
'Case Else
'fByVal = .PointerLevel = 0
'End Select
'If fByVal Then
'retVal = retVal & "ByVal "
'End If
'retVal = retVal & pi.Name
'If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
'retVal = retVal & "()"
'End If
''Error
'If TIType Is Nothing Then
'retVal = retVal & " As ?"
'Else
'If .IsExternalType Then
'retVal = retVal & " As " & .TypeLibInfoExternal.Name & "." & TIType.Name
'Else
'retVal = retVal & " As " & TIType.Name
'End If
'End If
'On Error GoTo 0
'Else
'If .PointerLevel = 0 Then
'retVal = retVal & "ByVal "
'End If
'retVal = retVal & pi.Name
'If VarTypeCur <> vbVariant Then
'strTypeName = TypeName(.TypedVariant)
'If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
'retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
'Else
'retVal = retVal & " As " & strTypeName
'End If
'End If
'End If
'If fOptional Then
'If fDefault Then
'retVal = retVal & ProduceDefaultValue(pi.DefaultValue, TIResolved)
'End If
'retVal = retVal & "]"
'End If
'End With
'Next pi
'retVal = retVal & ")"
'End If
'End With
'If fIsConstant Then
'ConstVal = .Value
'retVal = retVal & " = " & ConstVal
'Select Case VarType(ConstVal)
'Case vbInteger, vbLong
'If ConstVal < 0 Or ConstVal > 15 Then
'retVal = retVal & " (&H" & Hex$(ConstVal) & ")"
'End If
'End Select
'Else
'With .ReturnType
'VarTypeCur = .VarType
'If VarTypeCur = 0 Or (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
''If Not .TypeInfoNumber Then 'This may error, don't use here
'On Error Resume Next
'If Not .TypeInfo Is Nothing Then
''Information not available
'If Err Then
'retVal = retVal & " As ?"
'Else
'If .IsExternalType Then
'retVal = retVal & " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
'Else
'retVal = retVal & " As " & .TypeInfo.Name
'End If
'End If
'End If
'If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
'retVal = retVal & " () "
'End If
'Else
'Select Case VarTypeCur
'Case VT_VARIANT, VT_VOID, VT_HRESULT
'Case Else
'strTypeName = TypeName(.TypedVariant)
'If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
'retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
'Else
'retVal = retVal & " As " & strTypeName
'End If
'End Select
'End If
'End With
'End If
'PrototypeMember = retVal & vbNewLine & " " & "Member of " & TLInf.Name & "." & TLInf.GetTypeInfo(SearchData And &HFFFF).Name & vbNewLine & " " & .HelpString
'End With
'End With
'End Function
'
''Public Function ReferenceLibraryClass(ByVal strTest As String) As Boolean
'''Finds all control classes in reference libraries
''"msvbvm60.dll" = VBA
'', "msvbvm60.dll\3") Then 'VBRUN
'', "vb6.olb") Then 'Visual Basic objects and procedures
'', "Comdlg32.ocx")
''Dim CLAS As tli.Interfaces
''Dim I    As Long
''Dim J    As Long
''Dim K    As Long
''For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
''Set CLAS = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Interfaces
''For I = 1 To CLAS.Count
''For J = 1 To CLAS.Item(I).Members.Count
'''copes with commands that have a Variant and String version ie Chr and Chr$
''If CLAS.Item(I).Members.Item(J).name = strTest Then
''ReferenceLibraryClass = True
''Exit Function
''End If
''Next J
''Next I
''Next K
''End Function
''
''Public Function ReferenceLibraryCommand(ByVal strTest As String) As Boolean
'''Finds all constants in reference libraries
'''"msvbvm60.dll" = VBA
''', "msvbvm60.dll\3") Then 'VBRUN
''', "vb6.olb") Then 'Visual Basic objects and procedures
'', "Comdlg32.ocx")
''Dim CMND As tli.Declarations
''Dim I    As Long
''Dim J    As Long
''Dim K    As Long
''For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
''Set CMND = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Declarations
''For I = 1 To CMND.Count
''For J = 1 To CMND.Item(I).Members.Count
'''copes with commands that have a Variant and String version ie Chr and Chr$
''If ArrayMember(CMND.Item(I).Members.Item(J).name, strTest, "_B_str_" & strTest, "_B_var_" & strTest) Then
''ReferenceLibraryCommand = True
''Exit Function
''End If
''Next J
''Next I
''Next K
''End Function
''
''Public Function ReferenceLibraryControlDefaultPropertyType(ByVal StrCtrlClass As String) As String
'''the '_Default' method of a control has the same VTableOffset as the method itself
'''SO find a '_Default' method and get its offset
'''Search the offsets for matching offset (and ignore the 2nd'_Default" PROPERTYPUT mmember)
'''Return Default Property of control
''Dim INTERFACE     As tli.Interfaces
''Dim IFACE_MEMB    As tli.MemberInfo
''Dim I             As Long
''Dim J             As Long
''Dim M             As Long
''Dim strLib        As String
''Dim LVTableOffset As String
''strLib = RefLibContainsControldata(StrCtrlClass)
''If LenB(strLib) Then
''Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
''For I = 1 To INTERFACE.Count
''If Right$(INTERFACE.Item(I).name, Len(StrCtrlClass)) = StrCtrlClass Then
''For J = 1 To INTERFACE.Item(I).Members.Count
''Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(J)
''If ArrayMember(IFACE_MEMB.InvokeKind, tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT, tli.INVOKE_PROPERTYPUTREF) Then
''If IFACE_MEMB.name = "_Default" Then
''LVTableOffset = IFACE_MEMB.VTableOffset
''For M = J To INTERFACE.Item(I).Members.Count
''Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(M)
''If LVTableOffset = IFACE_MEMB.VTableOffset Then
''If IFACE_MEMB.name <> "_Default" Then
''ReferenceLibraryControlDefaultPropertyType = RLibType2VBType(IFACE_MEMB.ReturnType.VarType)
''Exit Function
''End If
''End If
''Next M
''End If
''End If
''Next J
''End If
''Next I
''End If
''End Function
''
''
''Public Function ReferenceLibraryVarStrCommand(ByVal strTest As String) As Boolean
'''Finds all constants in reference libraries
'''"msvbvm60.dll" = VBA
''', "msvbvm60.dll\3") Then 'VBRUN
''', "vb6.olb") Then 'Visual Basic objects and procedures
'', "Comdlg32.ocx")
''Dim CMND   As tli.Declarations
''Dim I      As Long
''Dim J      As Long
''Dim K      As Long
''Dim HasVar As Boolean
''Dim HasStr As Boolean
''For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
''Set CMND = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Declarations
''For I = 1 To CMND.Count
''For J = 1 To CMND.Item(I).Members.Count
'''copes with commands that have a Variant and String version ie Chr and Chr$
''If CMND.Item(I).Members.Item(J).name = "_B_str_" & strTest Then
''HasStr = True
''End If
''If CMND.Item(I).Members.Item(J).name = "_B_var_" & strTest Then
''HasVar = True
''End If
''If HasStr And HasVar Then
''ReferenceLibraryVarStrCommand = True
''Exit Function
''End If
''Next J
''Next I
''Next K
''End Function
''
''Public Sub RefLibAttributes(ByVal strLib As String, ByVal StrCtrlClass As String)
''Dim TI           As TypeInfo
''Dim Attributes() As String
''Dim I            As Long
''Set TI = TypeLibInfoFromFile(strLib).GetTypeInfo(StrCtrlClass)
''For I = 1 To TI.AttributeStrings(Attributes)
'''  Debug.Print Attributes(I)
''Next I
''End Sub
''
'Public Function RefLibConstantType(ByVal strTest As String) As String
''returns the type of vbConstants ie if strTest = vbMonday then returns VbDayOfWeek
'Dim CNST As tli.Constants
'Dim I    As Long
'Dim J    As Long
'Dim K    As Long
'For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
'Set CNST = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Constants
'For I = 1 To CNST.Count
'For J = 1 To CNST.Item(I).Members.Count
'If CNST.Item(I).Members.Item(J).Name = strTest Then
'RefLibConstantType = CNST.Item(I)
'Exit Function
'End If
'Next J
'Next I
'Next K
'End Function
''
''Public Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
''If SearchData And &H80000000 Then
'''VB SearchData routines
''GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F) Or &H80
''Else
''GetSearchType = SearchData \ &H1000000 And &HFF
''End If
''End Function
''
''Public Function ReferenceLibraryControlEvent(ByVal strTest As String, ByVal StrCtrlClass As String, Optional strResetCase As String) As Boolean
''Dim INTERFACE As tli.Interfaces
''Dim I         As Long
''Dim J         As Long
''Dim strLib    As String
''strLib = RefLibContainsControldata(StrCtrlClass)
''If LenB(strLib) Then
''Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
''Set iface = TLI.TypeLibInfoFromFile(strLib).GetMemberInfo
''For I = 1 To INTERFACE.Count
''If Right$(INTERFACE.Item(I).name, Len(StrCtrlClass)) = StrCtrlClass Then
''skip over the hidden methods that cannot be legally addressed from code
''For J = 1 To INTERFACE.Item(I).Members.Count
''If INTERFACE.Item(I).Members.Item(J).InvokeKind = tli.INVOKE_FUNC Then
'' skip hidden Events such as '_Default'
''If Left$(INTERFACE.Item(I).Members.Item(J).name, 1) <> "_" Then
''If INTERFACE.Item(I).Members.Item(J).name = strTest Then
''ReferenceLibraryControlEvent = True
''Exit Function
''ElseIf LCase$(INTERFACE.Item(I).Members.Item(J).name) = LCase$(strTest) Then
''this will reset the string to correct case
''strResetCase = INTERFACE.Item(I).Members.Item(J).name
''ReferenceLibraryControlEvent = True
''Exit Function
''End If
''End If
''End If
''Next J
''End If
''Next I
''End If
''End Function
''
''Public Function RLibType2VBType(ByVal RLibType As Variant) As String
''Select Case RLibType
''Function ReferenceLibraryControlProperty(strtest As String, strLibrary As String)
'''Finds all constants in reference libraries
''Dim CNST As tli.Constants
''Dim I As Long
''Dim J As Long
''Dim K As Long
''For K = LBound(ReferenceMembers) To UBound(ReferenceMembers)
''  Set CNST = tli.TypeLibInfoFromFile(CStr(ReferenceMembers(K))).Constants
''  For I = 1 To CNST.Count
''  For J = 1 To CNST.Item(I).Members.Count
''  If CNST.Item(I).Members.Item(J).Name = strtest Then
''  ReferenceLibraryControlProperty = True
''  Exit Function
''  End If
''  Next J
'' Next
''Next
''
'''End Function
''
''Public Function RefLibGenerateVBColorConstArray() As Variant
'''create array of VBColor Constants
''  Dim INTERFACE As tli.Constants
''  Dim strTmp As String
''  Dim strMName As String
''  Dim I         As Long
''  Dim J         As Long
''    Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
''    For J = 1 To INTERFACE.Count
''    If INTERFACE.Item(J).Name = "ColorConstants" Or INTERFACE.Item(J).Name = "SystemColorConstants" Then
''    For I = 1 To INTERFACE.Item(J).Members.Count
''      strTmp = AccumulatorString(strTmp, INTERFACE.Item(J).Members.Item(I).Name)
''      Next I
''      End If
''    Next J
''FillArray RefLibGenerateVBColorConstArray, strTmp
'''End Function
''
''Public Function GetRefLibKnownVBConstantValue(varTest As Variant) As Long
'''Get value of a known VB Constant
''Dim INTERFACE As tli.Constants
'''Dim strTmp As String
'''Dim strMName As String
''Dim I         As Long
''Set INTERFACE = tli.TypeLibInfoFromFile("msvbvm60.dll\3").Constants
''For J = 1 To INTERFACE.Count
''For I = 1 To INTERFACE.Item(J).Members.Count
''If varTest = INTERFACE.Item(J).Members.Item(I).name Then
''GetRefLibKnownVBConstantValue = INTERFACE.Item(J).Members.Item(I).Value
''Exit Function
''End If
''Next I
''Next J
''End Function
''
''Public Function ReferenceLibraryControlDefaultPropertyType(ByVal strCtrlClass As String) As String
''Dim INTERFACE     As tli.Interfaces
''Dim IFACE_MEMB    As tli.MemberInfo
''Dim I             As Long
''Dim J             As Long
''Dim M             As Long
''Dim strLib        As String
''Dim LVTableOffset As String
'''Returns a string describing the Type of the Default Property of a control class
''strLib = RefLibContainsControldata(strCtrlClass)
''If LenB(strLib) Then
''Set INTERFACE = tli.TypeLibInfoFromFile(strLib).Interfaces
''For I = 1 To INTERFACE.Count
''If Right$(INTERFACE.Item(I).Name, Len(strCtrlClass)) = strCtrlClass Then
''For J = 1 To INTERFACE.Item(I).Members.Count
''Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(J)
''If ArrayMember(IFACE_MEMB.InvokeKind, tli.INVOKE_PROPERTYGET, tli.INVOKE_PROPERTYPUT, tli.INVOKE_PROPERTYPUTREF) Then
''If IFACE_MEMB.Name = "_Default" Then
''LVTableOffset = IFACE_MEMB.VTableOffset
''For M = J To INTERFACE.Item(I).Members.Count
''Set IFACE_MEMB = INTERFACE.Item(I).Members.Item(M)
''If LVTableOffset = IFACE_MEMB.VTableOffset Then
''If IFACE_MEMB.Name <> "_Default" Then
''ReferenceLibraryControlDefaultPropertyType = RefLibType2CodeType(INTERFACE.Item(I).Members.Item(M).ReturnType.VarType)
''Exit Function
''End If
''End If
''Next M
''End If
''End If
''Next J
''End If
''Next I
''End If
''End Function


':)Code Fixer V3.0.9 (25/03/2005 4:11:58 AM) 6 + 1230 = 1236 Lines Thanks Ulli for inspiration and lots of code.

