Attribute VB_Name = "mod_RenameControls"
'code behind Controls Tab
Option Explicit
Public Enum ListParts
  LPProj
  LPForm
  LPName
  LPIndexedName
  LPIndex
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private LPProj, LPForm, LPName, LPIndexedName, LPIndex
#End If
Public Enum BadNameClasses
  BNNone = 0
  BNClass = 1
  BNReserve = 2
  BNKnown = 3
  BNCommand = 4
  BNVariable = 5
  BNProc = 6
  BNDefault = 7
  BNSingle = 8
  BNStructural = 9
  BNSingletonArray = 10
  BNMultiForm = 20
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private BNNone, BNClass, BNReserve, BNKnown, BNCommand, BNVariable, BNProc, BNDefault, BNSingle, BNStructural, BNSingletonArray, BNMultiForm
#End If


Private Sub AllPRojectsDisplay(Lsv As ListView, _
                               ByVal LPos As Long)

  Dim LItem As ListItem
  Dim I     As Long

  HideColumnsReturnPosition Lsv, LPos
  If bProjDescExists Then
    SendMessage Lsv.hWnd, WM_SETREDRAW, False, 0
    For I = LBound(ProjDesc) To UBound(ProjDesc)
      Set LItem = Lsv.ListItems.Add(, , ProjDesc(I).PDName)
      With LItem
        .ListSubItems.Add , , ProjDesc(I).PDFilename
        .ListSubItems.Add , , I
        If ProjDesc(I).PDBadType > 0 Then
          .ListSubItems.Add , , "Poor Name" & BadNameMsg2(ProjDesc(I).PDBadType)
          .ListSubItems(4).ForeColor = vbRed
         Else
          .ListSubItems.Add , , ""
        End If
      End With 'LItem
NotBad:
    Next I
    DoListViewWidth Lsv, Array(3)
    If UBound(ModDesc) > -1 Then
      If LPos = 0 Then
        LPos = 1
      End If
    End If
    SendMessage Lsv.hWnd, WM_SETREDRAW, True, 0
    SetCurrentLSVLine Lsv, LPos
    SetProjectListitem
  End If

End Sub

Public Sub ControlRenameButton(ByVal bEnabled As Boolean)

  If Not (AutoReName And bEnabled) Then
    With frm_CodeFixer
      .cmdCtrlChange.Enabled = bEnabled
      '.cmdEditSuggest.Enabled = bEnabled
      ControlAutoEnabled bEnabled
      If bEnabled Then
        UpdateControlsCaption .lsvAllControls
       Else
        WarningLabel "Working...", vbRed
      End If
    End With
  End If

End Sub

Public Sub DoFrontPageListView(Optional Smsg As String)

  Dim I     As Long
  Dim LItem As ListItem

  With frm_FindSettings.lsvModNames
    .ListItems.Clear
    SendMessage .hWnd, WM_SETREDRAW, False, 0
    FrontTabEnable False
  End With
  SetVBP_VBGRead_Write
  bSomeFilesReadOnly = False
  For I = LBound(ModDesc) To UBound(ModDesc)
    With ModDesc(I)
      If Len(.MDName) Then
        If FileSize(.MDFullPath) > 64 Then
          Smsg = Smsg & .MDFilename & SngSpace & ModuleSize(.MDFullPath) & vbNewLine
        End If
        If ModuleAttributeDanger(I) Then
          If Xcheck(XReadWrite) Then
            MakeVBFilesReadable I
            .MDReadOnly = False
           Else '
            With ModDesc(I)
              If mObjDoc.Safe_MsgBox("The file" & strInSQuotes(.MDFilename, True) & "or one of its support files has attributes which will interfer with Editing." & vbNewLine & _
                       "Would you like to change this Attribute?(Recommended)" & vbNewLine & _
                       "(See Settings|General Tab to make this automatic)" & vbNewLine & _
                       "NOTE: Answering 'No' disables Reload and Restore And Code Fixer fixes for this file.", vbQuestion + vbYesNo) = vbYes Then
                MakeVBFilesReadable I
                .MDReadOnly = False
              End If
            End With 'ModDesc(I)
          End If
        End If
        If .MDReadOnly Then
          bSomeFilesReadOnly = True
        End If
        Set LItem = frm_FindSettings.lsvModNames.ListItems.Add(, , .MDName)
        LItem.SubItems(1) = .MDProj
        LItem.Tag = IIf(.MDReadOnly, "ReadOnly", vbNullString)
        LItem.Checked = (Not .MDReadOnly)
        'this stops user selecting a file that is read only
        Set LItem = frm_CodeFixer.lsvUnDone.ListItems.Add(, , .MDProj)
        LItem.SubItems(1) = .MDName
        ' LItem.Checked = (Not .MDReadOnly)
      End If
    End With
  Next I
  'DoListViewWidth frm_CodeFixer.lsvUnDone
  With frm_CodeFixer
    DoListViewWidth .lsvUnDone
    If VBInstance.VBProjects.Count > 1 Then
      DoListViewWidth frm_FindSettings.lsvModNames, Array(-1)
     Else
      DoListViewWidth frm_FindSettings.lsvModNames, Array(2)
      frm_FindSettings.lsvModNames.ColumnHeaders(1).Width = frm_FindSettings.lsvModNames.Width - IIf(frm_FindSettings.lsvModNames.ListItems.Count > ListViewVisibleItems(frm_FindSettings.lsvModNames), 310, 60)
    End If
    frm_FindSettings.lsvModNames.ColumnHeaders(2).Position = 1
    SendMessage frm_FindSettings.lsvModNames.hWnd, WM_SETREDRAW, True, 0
  End With
  FrontTabEnable True

End Sub

Public Sub Generate_ModuleArray(Optional ByVal bDoDisplay As Boolean = True)

  Dim CurCompCount As Long
  Dim MyhourGlass  As cls_HourGlass
  Dim I            As Long
  Dim Comp         As VBComponent
  Dim Proj         As VBProject

  Set MyhourGlass = New cls_HourGlass
  On Error Resume Next
  'forms with Missing References can cause a crash when this attempts to Activate it
  'so this needs Error Trap
  ' if the controldata already matches the controls don't redo search
  ReDim ModDesc(GetComponentCount) As ModuleDescriptor
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        CurCompCount = CurCompCount + 1
        If mObjDoc.GraphVisible Then
          ModuleMessage Comp, CurCompCount
        End If
        SetModuleData CurCompCount, Comp, Proj
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  bModDescExists = True
  If bDoDisplay Then
    If frm_CodeFixer.lsvAllModules.ListItems.Count Then
      AllModulesDisplay frm_CodeFixer.lsvAllModules, frm_CodeFixer.lsvAllModules.SelectedItem.Index
     Else
      AllModulesDisplay frm_CodeFixer.lsvAllModules, 1
    End If
  End If
  For I = 1 To UBound(ModDesc)
    ModDesc(I).MDDontTouch = UnTouchable(ModDesc(I).MDName)
  Next I
  On Error GoTo 0
  OptionalCompileTempDisable

End Sub

Public Sub Generate_ProjectArray(Optional ByVal bDoDisplay As Boolean = False)

  Dim Proj As VBProject
  Dim I    As Long

  On Error Resume Next
  'forms with Missing References can cause a crash when this attempts to Activate it
  'so this needs Error Trap
  ReDim ProjDesc(VBInstance.VBProjects.Count - 1) As ProjectDescriptor
  For Each Proj In VBInstance.VBProjects
    With ProjDesc(I)
      .PDName = Proj.Name
      .PDFullPath = Proj.FileName
      .PDFilename = FileNameOnly(Proj.FileName)
      .PDBadType = BadModuleName(Proj.Name, vbNullString)
      .PDReadOnly = IIf(IsFileReadOnly(.PDFullPath), 1, 0)
      .PMDHidden = IIf(IsFileHidden(.PDFullPath), 1, 0)
      .PDAttributes = GetAttributes(.PDFullPath)
    End With
    I = I + 1
  Next Proj
  bProjDescExists = True
  If bDoDisplay Then
    If frm_CodeFixer.lsvAllProjects.ListItems.Count Then
      AllPRojectsDisplay frm_CodeFixer.lsvAllProjects, frm_CodeFixer.lsvAllProjects.SelectedItem.Index
     Else
      AllPRojectsDisplay frm_CodeFixer.lsvAllProjects, 1
    End If
  End If
  On Error GoTo 0

End Sub

Private Function GetAttributes(ByVal FName As String) As Long

  If LenB(FName) Then
    GetAttributes = FSO.GetFile(FName).Attributes
  End If

End Function

Public Sub OptionalCompileReEnable()

  Dim I         As Long
  Dim StartLine As Long
  Dim Comp      As VBComponent
  Dim CompMod   As CodeModule

  'v2.0.7 Thanks Paul Caton
  'fix for whole code being in Opt Compile structure
  'This replaces the structure removed in OptionalCompileTempDisable
  If bModDescExists Then
    For I = 1 To UBound(ModDesc)
      If ModDesc(I).MDWholeOptCompile Then
        Set Comp = VBInstance.VBProjects(ModDesc(I).MDProj).VBComponents(ModDesc(I).MDName)
        Set CompMod = Comp.CodeModule
        ' the "" & "" stops Code Fixer detecting itself and damaging this code
        With CompMod
          If .Find("'OptCompGuard", StartLine, 1, -1, -1, True, True, False) Then
            If InComment(.Lines(StartLine, 1), InStr(.Lines(StartLine, 1), "'OptCompGuard")) Then
              .InsertLines .CountOfLines + 1, "#End If"
              .ReplaceLine StartLine, Replace$(.Lines(StartLine, 1), "'OptCompGuard ", "")
            End If
          End If
        End With 'CompMod
      End If
    Next I
  End If

End Sub

Private Sub OptionalCompileTempDisable()

  Dim I         As Long
  Dim StartLine As Long
  Dim Comp      As VBComponent
  Dim CompMod   As CodeModule

  'v2.0.7 Thanks Paul Caton
  'fix for whole code being in Opt Compile structure
  'This temporarily disables the structure
  If bModDescExists Then
    For I = 1 To UBound(ModDesc)
      If ModDesc(I).MDWholeOptCompile Then
        Set Comp = VBInstance.VBProjects(ModDesc(I).MDProj).VBComponents(ModDesc(I).MDName)
        Set CompMod = Comp.CodeModule
        If CompMod.Find("#If", StartLine, 1, -1, -1, True, True, False) Then
          ' the "" & "" stops Code Fixer detecting itself and damaging this code
          CompMod.ReplaceLine StartLine, "'OptCompGuard " & CompMod.Lines(StartLine, 1)
          'Check this is legit
        End If
      End If
    Next I
  End If

End Sub

Private Sub SetModuleData(ByVal CurCompCount As Long, _
                          Comp As VBComponent, _
                          Proj As VBProject)

  With ModDesc(CurCompCount)
    .MDName = Comp.Name
    .MDBadType = BadModuleName(.MDName, .MDType)
    .MDType = ModuleType(Comp.CodeModule)
    .MDTypeNum = Comp.Type
    .MDisControlHolder = IsComponent_ControlHolder(Comp)
    .MDProj = Proj.Name
    If Len(Comp.FileNames(1)) Then
      .MDFullPath = Comp.FileNames(1)
      .MDFilename = FileNameOnly(Comp.FileNames(1))
      .MDAttributes = GetAttributes(.MDFullPath)
      .MDReadOnly = IsFileReadOnly(.MDFullPath)
      .MDHidden = IsFileHidden(.MDFullPath)
      .MDWholeOptCompile = TestOptCompileModule(Comp.CodeModule)
     Else
      .MDFilename = strUnsavedModule
    End If
    If ArrayMember(Comp.Type, vbext_ct_MSForm, vbext_ct_VBMDIForm, vbext_ct_VBForm) Then
      On Error Resume Next
      .MDCaption = Comp.Properties("caption").Value
      On Error GoTo 0
    End If
  End With

End Sub

Public Sub SetProjectListitem()

  Dim I    As Long

  With frm_CodeFixer
    For I = 0 To 3
      .cmdEditProj(I).Enabled = False
    Next I
    If bProjDescExists Then
      I = GetHiddenDescriptorIndex(.lsvAllProjects, 2)
      .txtProjectEdit(0) = ProjDesc(I).PDName
      .txtProjectEdit(1) = ProjDesc(I).PDFilename
      .txtProjectEdit(2) = ProjDesc(I).PDFullPath
    End If
  End With
  suggestProjName

End Sub

Private Function TestOptCompileModule(Cmp As CodeModule) As Boolean

  Dim StartLine  As Long
  Dim EndOptComp As Long

  StartLine = 1
  '2.4.0 Thanks again Paul Caton, this was the bug with
  'a Declaration only module ending with opt compilation.
  'in this circumstance VB gets it right; with the final '#End If'
  'as part of the Decaratio section so none of the
  'Optional Complile Module protection needs to be used
  If Cmp.CountOfDeclarationLines = Cmp.CountOfLines Then
    TestOptCompileModule = False
   Else
    'v2.0.7 Thanks Paul Caton
    'fix for whole code being in Opt Compile structure
    'Tests if the code is in this state
    If Cmp.Find("#If", StartLine, 1, -1, -1, True, True, False) Then
      If StartLine < Cmp.CountOfDeclarationLines Then
        Do While Cmp.Find("#End If", EndOptComp, 1, -1, -1, True, True, False)
          If EndOptComp = Cmp.CountOfLines Then
            TestOptCompileModule = True
            Exit Do
          End If
          EndOptComp = EndOptComp + 1
          If EndOptComp > Cmp.CountOfLines Then
            Exit Do
          End If
        Loop
        If Not TestOptCompileModule Then
          TestOptCompileModule = True
          For StartLine = EndOptComp To Cmp.CountOfLines
            If Not JustACommentOrBlank(Cmp.Lines(StartLine, 1)) Then
              TestOptCompileModule = False
              Exit For 'unction
            End If
          Next StartLine
        End If
      End If
    End If
  End If

End Function

Public Function UpdateCtrlnameLists() As Boolean

  ControlRenameButton False
  UpdateCtrlnameLists = Generate_ControlArray
  ControlRenameButton True

End Function

Public Sub UpdateModuleList()

  Generate_ProjectArray True
  Generate_ModuleArray True
  Generate_ProjectArray
  DoFrontPageListView

End Sub

Public Sub UpdateProjectList()

  Generate_ProjectArray True
  Generate_ModuleArray True
  Generate_ModuleArray
  DoFrontPageListView

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:17:21 AM) 29 + 375 = 404 Lines Thanks Ulli for inspiration and lots of code.

