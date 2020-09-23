Attribute VB_Name = "mod_HardCodeCtrls"
Option Explicit

Private Function CreateHardCodeFormValue(Comp As VBComponent, _
                                         K As Long) As String

  CreateHardCodeFormValue = Comp.Properties(K).Value
  If Not IsNumeric(CreateHardCodeFormValue) Then
    If Not ArrayMember(CreateHardCodeFormValue, "True", "False") Then
      CreateHardCodeFormValue = strInDQuotes(CreateHardCodeFormValue)
    End If
   Else
    If InStr(Comp.Properties(K).Name, "Color") Then
      CreateHardCodeFormValue = RefLibVBColorConstFromValue(CLng(CreateHardCodeFormValue))
    End If
  End If

End Function

Private Function CreateHardCodeValue(vbc As VBControl, _
                                     K As Long, _
                                     ByVal strType As String) As String

  On Error GoTo WriteOnlyProp
  CreateHardCodeValue = vbc.Properties(vbc.Properties(K).Name)
  If strType = "String" Then
    CreateHardCodeValue = strInDQuotes(CreateHardCodeValue)
  End If
  If strType = "Long" Then
    If InStr(vbc.Properties(K).Name, "Color") Then
      CreateHardCodeValue = RefLibVBColorConstFromValue(CLng(CreateHardCodeValue))
    End If
  End If

Exit Function

WriteOnlyProp:
  If Err.Number = 394 Then
    CreateHardCodeValue = "Write Only Property"
  End If

End Function

Public Sub DoHardCodeAll(Optional ByVal bStrOnly As Boolean = False)

  Dim CCtrl       As Long
  Dim vbc         As VBControl
  Dim vbf         As VBForm
  Dim Comp        As VBComponent
  Dim Proj        As VBProject
  Dim strAllCData As String
  Dim MyhourGlass As cls_HourGlass
  Dim StrFormDesc As String
  Dim PrjCount    As Long

  Set MyhourGlass = New cls_HourGlass
  mObjDoc.ShowWorking True, "Hard-coding All Forms and Controls", , False
  If GenerateReferencesEnumArray Then
    PrjCount = GetProjectCount
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If LenB(Comp.Name) Then
          If IsComponent_ControlHolder(Comp) Then
            ActivateDesigner Comp, vbf, False
            If vbf Is Nothing Then
              ActivateDesigner Comp, vbf, True
            End If
            StrFormDesc = vbNullString
            GenerateFormData Comp, StrFormDesc, bStrOnly
            For Each vbc In vbf.VBControls
              CCtrl = CCtrl + 1
              mObjDoc.ShowWorking True, IIf(PrjCount > 0, "Hard-coding all Form and Controls", "Hard-coding current Form and Controls"), StrPercent(CCtrl, vbf.VBControls.Count), False
              GenerateControlData vbc, Comp.Name, strAllCData, bStrOnly
            Next vbc
            WriteHardCode Comp, Nothing, strAllCData, StrFormDesc, bStrOnly
          End If
        End If
      Next Comp
    Next Proj
    On Error GoTo 0
  End If
  mObjDoc.ShowWorking False

End Sub

Public Sub DoHardCodeForm(Optional ByVal FrmCount As Long = 0, _
                          Optional ByVal bStrOnly As Boolean = False)

  Dim Comp        As VBComponent
  Dim vbc         As VBControl
  Dim vbf         As VBForm
  Dim strAllCData As String
  Dim MyhourGlass As cls_HourGlass
  Dim StrFormDesc As String
  Dim CCtrl       As Long

  Set MyhourGlass = New cls_HourGlass
  If GenerateReferencesEnumArray Then
    ActivateDesigner VBInstance.SelectedVBComponent, vbf, False
    If Not vbf Is Nothing Then
      Set Comp = vbf.Parent
      GenerateFormData Comp, StrFormDesc, bStrOnly
      For Each vbc In vbf.VBControls
        CCtrl = CCtrl + 1
        mObjDoc.ShowWorking True, "Hard-coding" & IIf(bStrOnly, " Strings ", " ") & IIf(FrmCount > 0, "all Forms", "current Form") & " and Controls", StrPercent(CCtrl, vbf.VBControls.Count), False
        GenerateControlData vbc, Comp.Name, strAllCData, bStrOnly
      Next vbc
      WriteHardCode Comp, Nothing, strAllCData, StrFormDesc, bStrOnly
    End If
    If FrmCount = 0 Then
      mObjDoc.ShowWorking False
    End If
  End If
  On Error GoTo 0

End Sub

Public Sub DoHardOneControl(Optional ByVal bStrOnly As Boolean = False)

  Dim vbc         As VBControl
  Dim vbf         As VBForm
  Dim Comp        As VBComponent
  Dim strAllCData As String

  If GenerateReferencesEnumArray Then
    Set Comp = GetActiveCodeModule.Parent
    ActivateDesigner VBInstance.SelectedVBComponent, vbf, False
    If vbf.SelectedVBControls.Count = 1 Then
      mObjDoc.ShowWorking True, "Hard-coding one control", , False
      Set vbc = vbf.SelectedVBControls.Item(0)
      GenerateControlData vbc, VBInstance.SelectedVBComponent.Name, strAllCData, bStrOnly
      WriteHardCode VBInstance.SelectedVBComponent, vbc, strAllCData, "", bStrOnly
      mObjDoc.ShowWorking False
    End If
  End If

End Sub

Public Sub DoRighClickControlsFinder()

  Dim vbf                                            As VBForm

  ActivateDesigner VBInstance.SelectedVBComponent, vbf, False
  If vbf.SelectedVBControls.Count = 1 Then
    bDontDisplayCodePane = True
    mObjDoc.ForceFind vbf.SelectedVBControls.Item(0).Properties("name")
  End If
  bDontDisplayCodePane = False

End Sub

Public Sub DoRighClickFormFinder()

  Dim vbf                                            As VBForm

  ActivateDesigner VBInstance.SelectedVBComponent, vbf, False
  bDontDisplayCodePane = True
  mObjDoc.ForceFind vbf.Parent.Name
  bDontDisplayCodePane = False

End Sub

Private Function FontData(Obj As Variant) As String

  Dim StrIndt As String

  StrIndt = Space$(IndentSize)
  FontData = StrIndt & ".Font.Name = " & strInDQuotes(Obj.Item(1).Value) & vbNewLine
  FontData = FontData & StrIndt & ".Font.Size = " & Obj.Item(2).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.Bold = " & Obj.Item(3).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.Italic = " & Obj.Item(4).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.UnderLine = " & Obj.Item(5).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.Strikethrough = " & Obj.Item(6).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.Weight = " & Obj.Item(7).Value & vbNewLine
  FontData = FontData & StrIndt & ".Font.Charset = " & Obj.Item(8).Value

End Function

Private Sub GenerateControlData(vbc As VBControl, _
                                ByVal CompName As String, _
                                OutPutString As String, _
                                Optional ByVal bStrOnly As Boolean = False)

  
  Dim StrIndt     As String
  Dim strCData    As String
  Dim Tindex      As String
  Dim TFullName   As String
  Dim K           As Long
  Dim strType     As String
  Dim strPropName As String
  Dim strValue    As String
  Dim strTmp      As String
  Dim arrTest1    As Variant
  Dim arrTest2    As Variant

  arrTest1 = Array("True", "False")
  arrTest2 = Array("Columns", "MultiSelect", "LinkMode", "WhatsThisHelp", "StartUpPosition", "Name", "Moveable")
  StrIndt = Space$(IndentSize)
  With vbc
    Tindex = .Properties("index").Value
    TFullName = CompName & "." & .Properties("name").Value & IIf(Tindex <> -1, strInBrackets(Tindex), vbNullString)
    For K = 1 To .Properties.Count
      strType = TypeOfProperty(.ClassName, .Properties(K).Name)
      strPropName = .Properties(K).Name
      strValue = CreateHardCodeValue(vbc, K, strType)
      Select Case strPropName
       Case "Font"
        strTmp = strTmp & FontData(.Properties("Font").Value) & vbNewLine
       Case "List"
        strTmp = strTmp & StrIndt & "'." & strPropName & " = (List)'Code Fixer cannot access the List property; if the List exists you should transfer it by hand to this routine using AddItem"
       Case "Index"
        strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue & " ' -1 means no index. Read only at Run-time. MUST be set from Properties window"
       Case "Picture", "Icon", "DragIcon", "MouseIcon", "DisabledPicture", "DownPicture", "Palette"
        If IsNumeric(strValue) Or IsInArray(strValue, arrTest1) Then
          If .Properties(strPropName).Value.Item(1) <> 0 Then
            strValue = "LoadPicture( " & DQuote & "[ImageFilePath]" & DQuote & " )'Image loaded from Properties window. To hard code fill in [ImageFilePath]"
            strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue
          End If
          If strPropName = "Palette" Then
            strCData = strCData & vbNewLine & _
             StrIndt & "'." & strPropName & " = [ImageFilePath]" & DQuote & " ) ' to hard code"
           Else
            strCData = strCData & vbNewLine & _
             StrIndt & "'." & strPropName & " = (None) 'Use  LoadPicture( " & DQuote & "[ImageFilePath]" & DQuote & " ) ' to hard code"
          End If
        End If
       Case Else
        If InStr(strPropName, "Picture") Then
          If IsNumeric(strValue) Or IsInArray(strValue, arrTest1) Then
            If .Properties(strPropName).Value.Item(1) <> 0 Then
              strTmp = strTmp & StrIndt & "'." & strPropName & " = " & _
                                                              "LoadPicture( " & DQuote & "[ImageFilePath]" & DQuote & " )'Image/icon loaded from Properties window. To hard code fill in [ImageFilePath]"
            End If
          End If
         Else
          Select Case strType
           Case "Write Only Property"
            strTmp = strTmp & "'" & strPropName & "= " & strValue
           Case "String", "Boolean", "Long", "Single", "Integer"
            If isReadWriteProperty(.ClassName, strPropName) And Not IsInArray(strPropName, arrTest2) Then
              On Error Resume Next
              If LenB(.Properties(strPropName)) Then
                If Err.Number = 0 Then
                  strTmp = strTmp & StrIndt & "." & strPropName & " = " & strValue
                End If
                Err.Clear
              End If
             Else
              If strPropName = "Columns" Then
                If strValue = 0 Then
                  strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue & "' If Value = 0 then Read only at Run-time. MUST be set from Properties window"
                 Else
                  strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue & "' Value cannot be set to 0 at Run-time. MUST be set from Properties window"
                End If
               Else
                strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue & "' Read only at Run-time. MUST be set from Properties window"
              End If
            End If
          End Select
        End If
      End Select
      If Len(strTmp) Then
        If bStrOnly Then
          If strType = "String" Then
            If Left$(Trim$(strTmp), 1) <> "'" Then
              strCData = strCData & strTmp & vbNewLine
            End If
          End If
         Else
          strCData = strCData & strTmp & vbNewLine
        End If
        strTmp = vbNullString
      End If
    Next K
  End With 'vbc
  If Len(strCData) Then
    Do While InStr(strCData, vbNewLine & vbNewLine)
      strCData = Replace$(strCData, vbNewLine & vbNewLine, vbNewLine)
    Loop
    If Right$(strCData, 2) = vbNewLine Then
      strCData = Left$(strCData, Len(strCData) - 2)
    End If
    If bStrOnly Then
      If InStr(strCData, vbNewLine) Then
        OutPutString = OutPutString & vbNewLine & _
         "With " & TFullName & vbNewLine & _
         Join(QuickSortArray(Split(strCData, vbNewLine)), vbNewLine) & vbNewLine & _
         "End with" & vbNewLine
       Else
        OutPutString = OutPutString & vbNewLine & TFullName & Trim$(strCData)
      End If
     Else
      OutPutString = OutPutString & vbNewLine & _
       "'" & TFullName & " Properties Hard-Coded" & vbNewLine & _
       "With " & TFullName & Join(QuickSortArray(Split(strCData, vbNewLine)), vbNewLine) & vbNewLine & _
       "End With" & vbNewLine
    End If
  End If
  strCData = vbNullString
  On Error GoTo 0

End Sub

Private Sub GenerateFormData(Comp As VBComponent, _
                             OutPutString As String, _
                             Optional ByVal bStrOnly As Boolean = False)

  Dim StrIndt     As String
  Dim K           As Long
  Dim strPropName As String
  Dim strValue    As String
  Dim strTmp      As String
  Dim arrTest     As Variant

  arrTest = Array("LinkMode", "WhatsThisHelp", "StartUpPosition", "Name", "Moveable")
  StrIndt = Space$(IndentSize)
  For K = 1 To Comp.Properties.Count
    'bCommentOut = False
    On Error Resume Next
    strValue = CreateHardCodeFormValue(Comp, K)
    strPropName = Comp.Properties(K).Name
    Select Case strPropName
     Case "Font"
      strTmp = strTmp & vbNewLine & FontData(Comp.Properties("Font").Value)
     Case "Picture", "Icon", "MouseIcon"
      If Comp.Properties(strPropName).Value.Item(1) <> 0 Then
        strTmp = strTmp & StrIndt & "'." & strPropName & " = " & _
                                                        "LoadPicture( " & DQuote & "[ImageFilePath]" & DQuote & " )'Image/icon loaded from Properties window. To hard code fill in [ImageFilePath]"
       Else
        strTmp = strTmp & StrIndt & "'." & strPropName & " = (None) 'Use  LoadPicture( " & DQuote & "ImagePath" & DQuote & " ) ' to hard code"
      End If
     Case Else
      If isReadWriteProperty(ModuleType(Comp.CodeModule), strPropName) And Not IsInArray(strPropName, arrTest) Then
        strTmp = strTmp & StrIndt & "." & strPropName & " = " & strValue
       Else
        strTmp = strTmp & StrIndt & "'." & strPropName & " = " & strValue & "' Read only at Run-time. MUST be set from Proper  s window"
      End If
    End Select
    If bStrOnly Then
      Select Case strPropName
       Case "Caption", "Text", "ToolTipText"
        OutPutString = OutPutString & strTmp & vbNewLine
      End Select
     Else
      OutPutString = OutPutString & strTmp & vbNewLine
    End If
    strTmp = vbNullString
  Next K
  Do While InStr(OutPutString, vbNewLine & vbNewLine)
    OutPutString = Replace$(OutPutString, vbNewLine & vbNewLine, vbNewLine)
  Loop
  If Right$(OutPutString, 2) = vbNewLine Then
    OutPutString = Left$(OutPutString, Len(OutPutString) - 2)
  End If
  If bStrOnly Then
    If InStr(OutPutString, vbNewLine) Then
      OutPutString = "With " & Comp.Name & vbNewLine & _
                     Join(QuickSortArray(Split(OutPutString, vbNewLine)), vbNewLine) & vbNewLine & _
                     "End with" & vbNewLine
     Else
      OutPutString = Comp.Name & Trim$(OutPutString)
    End If
   Else
    OutPutString = "With " & Comp.Name & vbNewLine & _
                   Join(QuickSortArray(Split(OutPutString, vbNewLine)), vbNewLine) & vbNewLine & _
                   "End with" & vbNewLine
  End If
  On Error GoTo 0

End Sub

Public Function GetProjectCount() As Long

  On Error Resume Next
  GetProjectCount = VBInstance.VBProjects.Count
  On Error GoTo 0

End Function

Public Sub HardCodeObject(ByVal intIndex As Long, _
                          Optional bStrOnly As Boolean = False)

  Select Case intIndex
   Case 0
    DoHardOneControl bStrOnly
   Case 1
    DoHardCodeForm , bStrOnly
   Case 2
    DoHardCodeAll bStrOnly
  End Select

End Sub

Private Function StrPercent(ByVal CurV As Long, _
                            ByVal MaxV As Long, _
                            Optional strFormat As String = "###") As String

  Dim c As Single

  'calculate Percentage given two numbers default display as units
  If CurV > MaxV Then
    CurV = MaxV
  End If
  If MaxV = 0 Then
    StrPercent = Format$(0, strFormat) & "%"
   Else
    c = CurV / MaxV * 100
    If c > 100 Then
      c = 100
    End If
    StrPercent = Format$(c, strFormat) & "%"
  End If

End Function

Private Sub WriteHardCode(Comp As VBComponent, _
                          vbc As VBControl, _
                          strAllCData As String, _
                          ByVal StrFormDesc As String, _
                          Optional ByVal bStrOnly As Boolean = False)

  Dim StartLine    As Long
  Dim lngdummy     As Long
  Dim ProcName     As String
  Dim StrHeader    As String
  Dim StrSubHeader As String
  Dim strCallProc  As String
  Dim Sline        As Long
  Dim SCol         As Long
  Dim ELine        As Long
  Dim Ecol         As Long

  If Not vbc Is Nothing Then
    ProcName = "HardCode_" & IIf(bStrOnly, "String_", vbNullString) & vbc.Properties("name")
    StrSubHeader = "for the control " & vbc.Properties("name")
   Else
    ProcName = "HardCode_" & IIf(bStrOnly, "String_", vbNullString) & Comp.Name
    StrSubHeader = "for this Form and its Controls"
  End If
  Select Case Comp.Type
   Case vbext_ct_DocObject
    strCallProc = "UserDocument_InitProperties"
   Case vbext_ct_UserControl
    strCallProc = "UserControl_InitProperties"
   Case Else
    strCallProc = "Form_Load"
  End Select
  StrHeader = "' This Sub created by Code Fixer contains " & IIf(bStrOnly, "String", "all") & " Properties " & StrSubHeader & vbNewLine & _
              "' The Sub is called from the Form_Load or InitProperties Event. " & vbNewLine
  StrHeader = StrHeader & IIf(bStrOnly, "' You may use this to quickly translate the interface to other languages." & vbNewLine & _
   "' Create a Select Case Structure where each Case contains a copy of the code generated." & vbNewLine & _
   "' Create a Module level variable ('lngLanguage') to select the language to use." & vbNewLine & _
   "' Create a Menu that allows the user to select the language to use." & vbNewLine & _
   "' Be careful about sizing issues, different languages may require longer labels, etc.", "' Many of the Properties will be the default values and can be deleted. " & vbNewLine & _
   "' Properties which cannot be set in code are commented out. " & vbNewLine & _
   "' Properties which Code Fixer cannot access are also commented out.")
  If LenB(strAllCData) Then
    With Comp
      If .CodeModule.Find("Sub " & strCallProc, StartLine, lngdummy, lngdummy, lngdummy) Then
        If Not .CodeModule.Find("HardCode" & .Name, StartLine, lngdummy, lngdummy, lngdummy) Then
          StartLine = GetSafeInsertLine(.CodeModule, StartLine)
          .CodeModule.InsertLines StartLine, ProcName & "' Code Fixer created this"
        End If
       Else
        .CodeModule.AddFromString "Private Sub " & strCallProc & "()" & vbNewLine & _
         ProcName & vbNewLine & _
         "End Sub"
      End If
    End With 'Comp
    If Not Comp.CodeModule.Find("Sub " & ProcName, 1, lngdummy, lngdummy, lngdummy) Then
      Comp.CodeModule.AddFromString "Private Sub " & ProcName & "()" & vbNewLine & _
       StrHeader & vbNewLine & _
       StrFormDesc & vbNewLine & _
       strAllCData & vbNewLine & _
       "End Sub"
    End If
    'v2.3.6 new this compacts the HardCode_String procedures
    '(the full hard-coder does it better for the more complex requirement of that fix)
    If bStrOnly Then
      Comp.CodeModule.Find "Sub " & ProcName, Sline, SCol, ELine, Ecol
      Comp.CodeModule.CodePane.SetSelection Sline, SCol, ELine, Ecol
      RightMenuWithRemove
      RightMenuWithCreate
    End If
  End If
  strAllCData = vbNullString
  On Error GoTo 0

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:27:57 AM) 1 + 449 = 450 Lines Thanks Ulli for inspiration and lots of code.

