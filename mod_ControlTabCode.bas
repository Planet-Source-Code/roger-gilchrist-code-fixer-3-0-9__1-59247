Attribute VB_Name = "mod_ControlTabCode"
Option Explicit
Public Type ControlDescriptor
  CDName                      As String
  CDOldName                   As String
  CDOldIndex                  As String
  CDFullName                  As String
  CDIndex                     As Long
  CDClass                     As String
  CDDefProp                   As String
  CDImageListLink             As String
  CDImageListLinkedTo         As String
  CDForm                      As String
  CDCaption                   As String
  CDUsage                     As Long
  CDProj                      As String
  CDXPFrameBug                As Boolean
  CDBadType                   As Long
  CDIsContainer               As Boolean
  CDContains                  As String    'v2.1.3 added to speed up the XPFRameBug detection
  CDStyle                     As Long
End Type
Public CntrlDesc()            As ControlDescriptor
Public bCtrlDescExists        As Boolean
Public LngCurrentControl      As Long
Private PrevListPos           As Long
Public AutoReName             As Boolean

Public Sub AutoFixSingletonControlArrays()

  Dim LItem          As ListItem
  Dim SaveSort       As Boolean
  Dim lngWorkingDesc As Long
  Dim MyhourGlass    As cls_HourGlass
  Dim I              As Long

  Set MyhourGlass = New cls_HourGlass
  WarningLabel "Auto Renaming(Caption)...", vbRed
  With frm_CodeFixer
    SaveSort = .lsvAllControls.Sorted
    .lsvAllControls.Sorted = False
    For I = 1 To .lsvAllControls.ListItems.Count
      SetCurrentLSVLine .lsvAllControls, I
      Set LItem = .lsvAllControls.ListItems(.lsvAllControls.SelectedItem.Index)
      lngWorkingDesc = GetTag(.lsvAllControls)
      If CntrlDesc(lngWorkingDesc).CDBadType <> 0 Then
        DoBadSelect
        With CntrlDesc(lngWorkingDesc)
          If HasSingletonCArray(.CDBadType) Then
            ChangeControlNameSingleton .CDName, .CDName '& "_" & .CDIndex
            LItem.SubItems(2) = .CDName
            .CDBadType = .CDBadType - BNSingletonArray
            LItem.ListSubItems(4).ForeColor = vbGreen
            LItem.SubItems(4) = "FIXED"
          End If
        End With 'CntrlDesc(lngWorkingDesc)
      End If
    Next I
    .lsvAllControls.Sorted = SaveSort
  End With

End Sub

Public Function BadNameMsg(varCtrlName As Variant) As String

  Dim lngTmpIndex As Long

  lngTmpIndex = CntrlDescMember(varCtrlName)
  If lngTmpIndex > -1 Then
    BadNameMsg = BadNameMsg2(CntrlDesc(lngTmpIndex).CDBadType)
  End If

End Function

Public Function BadReasonComment(ByVal lngIndex As Long) As String

  If CntrlDesc(lngIndex).CDBadType > 0 Then
    BadReasonComment = " Reason " & BadNameMsg2(CntrlDesc(lngIndex).CDBadType)
  End If

End Function

Private Sub ChangeControlInWith(ByVal strCtrl As String, _
                                ByVal strNewName As String)

  Dim L_CodeLine As String
  Dim Proj       As VBProject
  Dim LModule    As CodeModule
  Dim StartLine  As Long
  Dim Targets    As Variant
  Dim LUpdated   As Boolean
  Dim K          As Long
  Dim Comp       As VBComponent

  If CntrlDesc(LngCurrentControl).CDUsage <> 0 Then
    Targets = Array("With " & CntrlDesc(LngCurrentControl).CDForm, "With Me", "With UserControl", "With ActiveForm", "With UserDocument")
    For Each Proj In VBInstance.VBProjects
      If Proj.Name = CntrlDesc(LngCurrentControl).CDProj Then
        For Each Comp In Proj.VBComponents
          If LenB(Comp.Name) Then
            If DOPartialFindModule(Comp.Name, strCtrl) Then
              Set LModule = Comp.CodeModule
              DisplayCodePane Comp
              For K = LBound(Targets) To UBound(Targets)
                StartLine = 1
                If LModule.Find(Targets(K), StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
                  '                  Do
                  Do While LModule.Find(strCtrl, StartLine, 1, -1, -1, True, True)
                    L_CodeLine = LModule.Lines(StartLine, 1)
                    'If LModule.Find(strCtrl, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
                    ' L_CodeLine = LModule.Lines(StartLine, 1)
                    If Not JustACommentOrBlank(L_CodeLine) Then
                      If Not ProhibitedControlUpdate(L_CodeLine, strCtrl) Then
                        If CntrlDesc(LngCurrentControl).CDForm = Comp.Name Then
                          WholeWordReplacer L_CodeLine, strCtrl, strNewName, LUpdated
                         Else
                          If PreviousWith(LModule, StartLine, Targets) Then
                            WholeWordReplacer L_CodeLine, "." & strCtrl, "." & strNewName, LUpdated
                          End If
                        End If
                        If LUpdated Then
                          If LCase$(strNewName) <> LCase$(strCtrl) Then
                            SafeInsertModule LModule, StartLine, WARNING_MSG & "Control" & strInSQuotes(strCtrl, True) & "renamed to" & strInSQuotes(strNewName, True)
                          End If
                          LModule.ReplaceLine StartLine, L_CodeLine
                          LUpdated = False
                        End If
                      End If
                      '  End If
                    End If
                    StartLine = StartLine + 1
                    If Left$(LModule.Lines(StartLine, 1), 8) = "End With" Then
                      Exit Do
                    End If
                    If StartLine > LModule.CountOfLines Then
                      Exit Do
                    End If
                  Loop
                  'While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strCtrl, L_CodeLine, StartLine)
                End If
              Next K
            End If
          End If
        Next Comp
      End If
    Next Proj
  End If

End Sub

Private Sub ChangeControlOnModules(ByVal strCtrl As String, _
                                   ByVal strNewName As String, _
                                   Optional ByVal bSingletonSwitch As Boolean = False)

  
  Dim strOldName         As String
  Dim StartLine          As Long
  Dim strTargetPropError As String
  Dim L_CodeLine         As String
  Dim Proj               As VBProject
  Dim Comp               As VBComponent
  Dim LModule            As CodeModule
  Dim LUpdated           As Boolean
  Dim CurCompCount       As Long
  Dim lngPrevFindLine    As Long
  Dim strIndex           As String

  ' break Loop if the only codeline is a comment that prog should ignore
  If CntrlDesc(LngCurrentControl).CDUsage <> 0 Then
    If bSingletonSwitch Then
      strOldName = strCtrl
     Else
      strOldName = CntrlDesc(LngCurrentControl).CDOldName
    End If
    For Each Proj In VBInstance.VBProjects
      If Proj.Name = CntrlDesc(LngCurrentControl).CDProj Then
        For Each Comp In Proj.VBComponents
          If SafeCompToProcess(Comp, CurCompCount, False) Then
            If DOPartialFindModule(Comp.Name, strCtrl) Then
              Set LModule = Comp.CodeModule
              StartLine = 1
              If LModule.Find(strCtrl, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
                lngPrevFindLine = StartLine
                ' Do
                Do While LModule.Find(strCtrl, StartLine, 1, -1, -1, True, True)
                  L_CodeLine = LModule.Lines(StartLine, 1)
                  If StartLine < lngPrevFindLine Then
                    Exit Do
                  End If
                  If LModule.Find(strCtrl, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
                    L_CodeLine = LModule.Lines(StartLine, 1)
                    If Not JustACommentOrBlank(L_CodeLine) Then
                      If Not ProhibitedControlUpdate(L_CodeLine, strCtrl) Then
                        'v3.0.7 improved search (new has same nae as old but do it anyway
                        If InStr(L_CodeLine, strCtrl) <> InStr(L_CodeLine, strNewName) Or bSingletonSwitch Then
                          'v2.2.7 fixed to stop double insertion of parameters
                          If CntrlDesc(LngCurrentControl).CDForm = Comp.Name Then
                            WholeWordReplacer L_CodeLine, strCtrl, strNewName, LUpdated
                            'v 2.1.5 correct code calls to newly indexed controls
                            strIndex = vbNullString
                            If InStr(L_CodeLine, Left$(strNewName, Len(strNewName) - 1) & ")_") Then
                              If Not isProcHead(L_CodeLine) Then
                                'don't update the temporarily bad proc heads that creating the initial array member requires
                                strIndex = Mid$(L_CodeLine, InStr(L_CodeLine, strNewName))
                                strIndex = Mid$(L_CodeLine, InStr(L_CodeLine, "("))
                                strIndex = Left$(strIndex, InStr(strIndex, ")"))
                                L_CodeLine = Replace$(L_CodeLine, strIndex, vbNullString) & strIndex
                              End If
                            End If
                           Else
                            WholeWordReplacer L_CodeLine, CntrlDesc(LngCurrentControl).CDForm & "." & strCtrl, CntrlDesc(LngCurrentControl).CDForm & "." & strNewName, LUpdated
                            If Not LUpdated Then
                              If InStr(SngSpace & L_CodeLine, "." & strCtrl) = 2 Then
                                WholeWordReplacer L_CodeLine, " ." & strCtrl, " ." & strNewName, LUpdated
                               ElseIf SmartLeft(L_CodeLine, "." & strCtrl) Then
                                WholeWordReplacer L_CodeLine, "." & strOldName, "." & strNewName, LUpdated
                              End If
                            End If
                          End If
                          'v2.8.2 Quick'n'Dirty fix for a problem with using ActiveForm.
                          If InStr(L_CodeLine, "ActiveForm." & strOldName) Then
                            WholeWordReplacer L_CodeLine, "ActiveForm." & strOldName, "ActiveForm." & strNewName, LUpdated
                          End If
                          'End If
                          If LUpdated Then
                            'ver 1.1.63 this copes with whether or not the update is acceptable (99% accurate)
                            strTargetPropError = WordIncluding(Trim$(ExpandForDetection(L_CodeLine)), InStr(Trim$(ExpandForDetection(L_CodeLine)), "." & strNewName))
                            'get potential target point
                            If Len(strTargetPropError) Then ' potential ezists
                              If Left$(strTargetPropError, 1) = "." Then
                                If inLegalWith(StartLine, LModule, CntrlDesc(LngCurrentControl).CDForm) Then
                                  GoTo doNoRepair
                                 Else
                                  GoTo DoRepair
                                End If
                              End If
DoRepair:
                              'ver 1.1.62 this stops controls with same name as Properties from updating the Property calls
                              Do While Not MultiLeft(strTargetPropError, True, CntrlDesc(LngCurrentControl).CDForm, "Me", "UserDocument", "UserControl")
                                L_CodeLine = Replace$(L_CodeLine, strTargetPropError, Replace$(strTargetPropError, strNewName, strCtrl), , 1)
                                strTargetPropError = WordIncluding(L_CodeLine, InStr(L_CodeLine, "." & strNewName))
                                'retest
                                If LenB(strTargetPropError) = 0 Then
                                  'exit if no more fixed potential problem so
                                  Exit Do
                                 Else
                                  GoTo doNoRepair2
                                End If
                              Loop
doNoRepair:
                            End If
                            If InStr(L_CodeLine, strNewName) Or Len(strIndex) Then
                              '2.0.1 logic fixed
                              LModule.ReplaceLine StartLine, L_CodeLine
                              If InStr(L_CodeLine, strNewName) Or Len(strIndex) Then
                                If LCase$(strNewName) <> LCase$(strOldName) Then
                                  SafeInsertModule LModule, StartLine, WARNING_MSG & "Control" & strInSQuotes(strCtrl, True) & "renamed to" & strInSQuotes(strNewName, True)
                                End If
                              End If
                            End If
                            LUpdated = False
                          End If
                        End If
                      End If
                    End If
                  End If
doNoRepair2:
                  StartLine = StartLine + 1
                  If StartLine > LModule.CountOfLines Then
                    Exit Do
                  End If
                  lngPrevFindLine = StartLine
                Loop
                ' While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strCtrl, L_CodeLine, StartLine)
              End If
            End If
          End If
        Next Comp
      End If
    Next Proj
  End If

End Sub

Public Function CntrlDescMember2(varName As Variant, _
                                 varForm As Variant, _
                                 Optional ByVal varProj As Variant, _
                                 Optional ByVal StartPoint As Long = 0) As Long

  Dim I As Long

  If IsMissing(varProj) Then
    varProj = vbNullString
  End If
  CntrlDescMember2 = -1 ' report control does not exist with this return
  If bCtrlDescExists Then
    For I = StartPoint To UBound(CntrlDesc)
      If CntrlDesc(I).CDProj = varProj Or LenB(varProj) = 0 Then
        If CntrlDesc(I).CDForm = varForm Then
          If CntrlDesc(I).CDName = varName Then
            CntrlDescMember2 = I
            Exit For
          End If
        End If
      End If
    Next I
  End If

End Function

Private Sub ControlArrayMessage()

  fixList PrevListPos
  mObjDoc.Safe_MsgBox "NOTE:" & vbNewLine & _
                    "You are creating a new Control Array." & vbNewLine & _
                    "Previous individual Events will be rename to start with ' DELETE_ME_INDEXED_VERSION_CREATED_'" & vbNewLine & _
                    "The Code and comments of this procedure will be copied in to the new Event procedure's Select Case structure." & vbNewLine & _
                    "NOTE: You may need to do some hand re-coding to make this Control Array fully functional." & vbNewLine & _
                    "Remove duplicated declarations and seperate out any shared code that would be best done once before or after the Select case structure.", vbExclamation

End Sub

Public Function DeleteControl() As Boolean

  Dim LItem          As ListItem
  Dim Comp           As VBComponent
  Dim Targetcontrol  As VBControl
  Dim vbf            As VBForm
  Dim lngTargetIndex As Long
  Dim MyhourGlass    As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  'This is based on the XP Frame solution
  'stripped down to simple select and cut the control
  Set LItem = frm_CodeFixer.lsvAllControls.SelectedItem
  lngTargetIndex = LItem.Tag
  If IsThisControlDeletable(lngTargetIndex) Then
    If vbYes = mObjDoc.Safe_MsgBox("Code Fixer has tested this control for any potential usage in code." & vbNewLine & _
                               "However there is a potential that it is addressed in code only by reference to the Controls Collection or in some way I have not thought of." & vbNewLine & _
                               "Are you sure you want to delete it?", vbYesNo + vbCritical) Then
      Set Comp = GetComponent(LItem.Text, LItem.SubItems(1))
      ActivateDesigner Comp, vbf
      Set Targetcontrol = GetControlItemFromName(CntrlDesc(lngTargetIndex).CDFullName, vbf.VBControls)
      Targetcontrol.InSelection = True
      If vbf.SelectedVBControls.Count Then
        'cut selected controls
        vbf.SelectedVBControls.Cut
        DeleteControl = True
      End If
    End If
   Else
    mObjDoc.Safe_MsgBox "This control did not pass all tests for being unused." & vbNewLine & _
                    "Code Fixer will not delete it!", vbInformation
  End If
  If frm_CodeFixer.txtCtrlNewName.Visible Then
    'just removes focus from button
    SetFocus_Safe frm_CodeFixer.txtCtrlNewName
  End If

End Function

Public Sub DeselectAll(AllCntrls As VBControls)

  Dim TmpC   As Variant

  'Clear selection
  'When you cut controls from form the selection jumps to next
  'TabOrder control and when you use InSelection previous
  'selections are not de-selected (if they are in the same ontainer) so this has to be used
  On Error Resume Next
  For Each TmpC In AllCntrls
    If TmpC.InSelection Then
      TmpC.InSelection = False
    End If
  Next TmpC
  On Error GoTo 0

End Sub

Public Sub DoBadSelect()

  Dim CurName As String

  With frm_CodeFixer
    GetCtrlDataLSV
    CurName = CntrlDesc(LngCurrentControl).CDName
    .lblOldName(1).Caption = CurName
    If .lstPrefixSuggest.ListIndex < 0 Then
      SetNewName CurName
      If CurName = .txtCtrlNewName Then
        SetNewName .lstPrefixSuggest.List(0)
      End If
     Else
      'v2.0.3 last line and just renamed arrays no longer trigger the possible array member signal
      If .lsvAllControls.SelectedItem.Index <> .lsvAllControls.ListItems.Count Then
        If Not SmartLeft(.txtCtrlNewName.Text, CurName) Then
          WarningLabel "Control could be made part of a control array with the name selected", vbYellow
        End If
      End If
    End If
    ' End If
  End With

End Sub

Private Function DoControlRename(strNewName As String, _
                                 Optional ByVal bMenuMode As Boolean = False) As Boolean

  
  Dim lngNewIndex            As Long
  Dim lngIndexTarget         As Long
  Dim vbf                    As VBForm
  Dim Targetcontrol          As VBControl
  Dim oldTargetControl       As VBControl
  Dim ArrayCntrls            As Variant
  Dim ctrla                  As Variant
  Dim Comp                   As VBComponent
  Dim lngTmpCtrlDescIndex    As Long
  Dim lngTmpCtrlOldDescIndex As Long
  Dim strHomeComp            As String
  Dim StrCap                 As String

  'ver 1193 changed to function to return own success value
  Set Comp = GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm)
  strHomeComp = Comp.Name
  ActivateDesigner Comp, vbf, False
  If vbf Is Nothing Then 'safety should never hit
    ActivateDesigner Comp, vbf, True
    If vbf Is Nothing Then
      Exit Function
    End If
  End If
  With vbf
    DoEvents
    'single controls
    ArrayCntrls = GetControlArrayFromName(CntrlDesc(LngCurrentControl).CDName, .VBControls)
    lngNewIndex = UBound(ArrayCntrls)
    lngIndexTarget = UBound(GetControlArrayFromName(strNewName, .VBControls))
    '-1 means no control with newname exists
    If UBound(ArrayCntrls) > -1 Then
      For Each ctrla In ArrayCntrls
        Set Targetcontrol = GetControlItemFromName(ctrla, .VBControls)
        If Targetcontrol Is Nothing Then
          'ver 2.0.1 bug fix stops crash
          GoTo SafeExit
          'Find max index for new array and reset index first
        End If
        lngTmpCtrlDescIndex = MatchCtrlToDescArrayIndex(Targetcontrol)
        If bMenuMode Then
          'menus need to change the index before the name
          'as changing the name first generates a stopping error
          'because the indexes are not consecative/in correct order
          If lngIndexTarget > -1 Then
            Targetcontrol.Properties("index").Value = lngIndexTarget + 1
          End If
        End If
        On Error GoTo NoMergeLegal 'Error Resumes safely in the With
        Targetcontrol.Properties("name").Value = strNewName
        'v3.0.0 adds the Caption to Case statements for generated control arrays
        StrCap = ExtractControlCaption(Targetcontrol, True)
        On Error GoTo 0
        If Targetcontrol.Properties("index").Value = 1 Then
          'creating a new array so fix code and rename initial member
          If UBound(ArrayCntrls) = 0 Then
            'setting newindex auto updates old control's name/index
            RenameControlsStandard strNewName, strNewName & "(0)" ', True
            Set oldTargetControl = GetControlItemFromName(strNewName & "(0)", .VBControls)
            'v3.0.0 adds the Caption to Case statements for generated  control  arrays
            FixNewlyIndexed strNewName, strHomeComp, ExtractControlCaption(oldTargetControl, True), True
            lngTmpCtrlOldDescIndex = GetCtrlDescArrayIndex(strNewName, strHomeComp)
            SetControlData lngTmpCtrlOldDescIndex, Targetcontrol, GetComponent(CntrlDesc(lngTmpCtrlOldDescIndex).CDProj, CntrlDesc(lngTmpCtrlOldDescIndex).CDForm), vbf
            ReWriteSelectedLineArrays lngTmpCtrlOldDescIndex, GetControlItemFromName(strNewName & "(0)", .VBControls)
            DoControlRename = True
            ControlArrayMessage
          End If
          ImageListWarning Targetcontrol, Comp
          Comp.CodeModule.InsertLines 1, WARNING_MSG & WARNING_MSG & "Control" & strInSQuotes(ctrla, True) & "renamed to" & strInSQuotes(ControlName(Targetcontrol), True) & BadReasonComment(lngTmpCtrlDescIndex)
          ReWriteSelectedLineArrays lngTmpCtrlDescIndex, Targetcontrol
          WarningLabel "Working... Updating " & CntrlDesc(LngCurrentControl).CDOldName & " Code", vbRed
          With Targetcontrol
            If .Properties("index").Value <> CStr(UBound(ArrayCntrls)) Then
              If UBound(ArrayCntrls) = 0 Then
                RenameControlsStandard CntrlDesc(LngCurrentControl).CDOldName, strNewName & strInBrackets(.Properties("index").Value)
                If DoControlRename Then
                  If .Properties("index").Value > -1 Then
                    If lngNewIndex < 1 Then
                      FixNewlyIndexed strNewName & strInBrackets(.Properties("index").Value), strHomeComp, StrCap
                    End If
                  End If
                End If
              End If
              'v 2.2.2 Thanks Mike Ulik these lines were outside the IF structure and messed up renaming arrays of controls
              DoControlRename = True
              SetControlData LngCurrentControl, Targetcontrol, GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm), vbf
            End If
          End With 'Targetcontrol
         Else
          ' control array already exists
          ' singleton control creation
          ImageListWarning Targetcontrol, Comp
          Comp.CodeModule.InsertLines 1, WARNING_MSG & "Control" & strInSQuotes(ctrla, True) & "renamed to" & strInSQuotes(ControlName(Targetcontrol), True) & BadReasonComment(lngTmpCtrlDescIndex)
          ReWriteSelectedLineArrays lngTmpCtrlDescIndex, Targetcontrol
          WarningLabel "Working... Updating " & CntrlDesc(LngCurrentControl).CDOldName & " Code", vbRed
          If Targetcontrol.Properties("index").Value = -1 Then
            'singleton code update
            RenameControlsStandard CntrlDesc(LngCurrentControl).CDOldName, strNewName
            DoControlRename = True
            SetControlData LngCurrentControl, Targetcontrol, GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm), vbf
           Else
            If UBound(ArrayCntrls) > 0 Then
              RenameControlsStandard CntrlDesc(LngCurrentControl).CDOldName, strNewName
              DoControlRename = True
              SetControlData LngCurrentControl, Targetcontrol, GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm), vbf
             ElseIf UBound(ArrayCntrls) = 0 Then
              If CntrlDesc(LngCurrentControl).CDIndex <> 0 Then
                RenameControlsStandard CntrlDesc(LngCurrentControl).CDOldName, strNewName & strInBrackets(Targetcontrol.Properties("index").Value)
               Else
                'v2.6.5 renaming singleton control array Thanks Mike Ulik
                RenameControlsStandard CntrlDesc(LngCurrentControl).CDOldName, strNewName
                '& strInBrackets(Targetcontrol.Properties("index").Value)
              End If
              If Targetcontrol.Properties("index").Value > -1 Then
                FixNewlyIndexed strNewName & strInBrackets(Targetcontrol.Properties("index").Value), strHomeComp, StrCap
                DoControlRename = True
                SetControlData LngCurrentControl, Targetcontrol, GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm), vbf
              End If
            End If
          End If
        End If
      Next ctrla
    End If
SafeExit:
  End With

Exit Function

NoMergeLegal:
  mObjDoc.Safe_MsgBox "Code Fixer cannot rename a control from one control array to another", vbExclamation
  Resume SafeExit

End Function

Private Function DOPartialFindModule(ByVal strMod As String, _
                                     ByVal strPFind As String) As Boolean

  Dim Proj      As VBProject
  Dim Comp      As VBComponent
  Dim Hit       As Boolean
  Dim StartLine As Long

  On Error Resume Next
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If Comp.Name = strMod Then
          Hit = True
          StartLine = 1
          If Comp.CodeModule.Find(strPFind, StartLine, 1, -1, -1, False, False, True) Then
            DOPartialFindModule = True
          End If
          Exit For
        End If
      End If
    Next Comp
    If DOPartialFindModule Or Hit Then
      Exit For
    End If
  Next Proj
  On Error GoTo 0

End Function

Public Sub EditControlName(ByVal strNewName As String)

  Dim bmenuWarn   As Boolean
  Dim MyhourGlass As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  FixControlCase strNewName
  If CntrlDesc(LngCurrentControl).CDOldName <> CntrlDesc(LngCurrentControl).CDName Then
    mObjDoc.Safe_MsgBox "You need to refresh the list before renaming a second time", vbInformation
   Else
    If strNewName <> CntrlDesc(LngCurrentControl).CDName Then
      If IllegalCharacters(strNewName) Then
        mObjDoc.Safe_MsgBox "Illegal characters in the control name " & strInSQuotes(strNewName) & vbNewLine & _
                    IllegalCharacterMsg(strNewName) & vbNewLine & _
                    "No action was taken.", vbInformation
        Exit Sub
      End If
      If Len(BadNameMsg(strNewName)) Then
        If mObjDoc.Safe_MsgBox("New control name" & strInSQuotes(strNewName, True) & "has the following problem: " & BadNameMsg(strNewName) & vbNewLine & _
                       "Proceed anyway?", vbInformation + vbYesNo) = vbYes Then
          GoTo DoItAnyway
         Else
          Exit Sub
        End If
      End If
      If Not IsMenu(CntrlDesc(LngCurrentControl).CDName) Then
        bmenuWarn = False
       Else
        If Not bmenuWarn Then
          bmenuWarn = True
          If mObjDoc.Safe_MsgBox("When creating Menu arrays you must be careful to rename them in top to bottom order and include any seperators." & vbNewLine & _
                       "1. Click on the Form and then open the Menu Editor to check menu structure." & vbNewLine & _
                       "2. Leave Menu Editor open to check targets NOTE: Menu Editor will not update Names while Code Fixer is working." & vbNewLine & _
                       "3. Close Menu Editor using the 'Cancel' button or it will undo the name changes(but NOT the code changes)" & vbNewLine & _
                       "Proceed?", vbInformation + vbYesNo) = vbYes Then
            GoTo DoItAnyway
           Else
            bmenuWarn = False
            Exit Sub
          End If
        End If
      End If
DoItAnyway:
      WarningLabel "Working... Renaming", vbRed
      If DoControlRename(strNewName, IsMenu(CntrlDesc(LngCurrentControl).CDName)) Then
        If IsImageList(CntrlDesc(LngCurrentControl).CDName) Then
          ForceUpDateImageListLink
        End If
       Else
        If Not AutoReName Then
          mObjDoc.Safe_MsgBox "The Control Rename Tool has failed to change the control name." & vbNewLine & _
                    "No action taken", vbInformation
        End If
      End If
      WarningLabel
      If Not AutoReName Then
        ReWriteSelectedLine
        With frm_CodeFixer
          If .lsvAllControls.SelectedItem.Index < .lsvAllControls.ListItems.Count Then
            PrevListPos = .lsvAllControls.SelectedItem.Index
            SetCurrentLSVLine .lsvAllControls, .lsvAllControls.SelectedItem.Index + 1
            AllControlsColumnWidths          'SetColumn .lsvAllControls, "control", e_CDName
            GetCtrlDataLSV
          End If
        End With 'fCodeFix
      End If
     Else
      mObjDoc.Safe_MsgBox "You didn't change the name; no action was taken.", vbInformation
    End If
  End If

End Sub

Private Function ExtractControlCaption(Ctrl As VBControl, _
                                       Optional ByVal bShort As Boolean = False) As String

  'v3.0.0 used by Control array builder

  If ControlHasProperty(Ctrl, "Caption") Then
    ExtractControlCaption = Ctrl.Properties("Caption")
    If bShort Then
      If Len(ExtractControlCaption) > 25 Then
        ExtractControlCaption = Left$(ExtractControlCaption, 20) & "..."
      End If
    End If
  End If

End Function

Private Function ExtractNumber(ByVal strExt As String) As String

  Dim I As Long

  strExt = Mid$(strExt, InStr(strExt, LBracket))
  For I = 1 To Len(strExt)
    If IsNumeric(Mid$(strExt, I, 1)) Then
      ExtractNumber = ExtractNumber & Mid$(strExt, I, 1)
    End If
  Next I

End Function

Private Sub fixList(ByVal lngOldListPos As Long)

  Dim LItem          As ListItem
  Dim lngIndex       As Long
  Dim SaveCurListPos As Long

  With frm_CodeFixer
    SaveCurListPos = .lsvAllControls.SelectedItem.Index
    SetCurrentLSVLine .lsvAllControls, lngOldListPos
    Set LItem = .lsvAllControls.SelectedItem
  End With 'fCodeFix
  lngIndex = CLng(LItem.Tag)
  With CntrlDesc(lngIndex)
    LItem.Text = .CDProj
    LItem.SubItems(1) = .CDForm
    .CDIndex = 0
    LItem.SubItems(2) = .CDFullName
    'v3.0.0 stopped bad display of wrong caption when creating 0th member of new array
    'LItem.SubItems(3) = .CDCaption
    .CDBadType = 0
    If GetBadNameType(lngIndex) <> BNSingletonArray Then
      .CDBadType = GetBadNameType(lngIndex)
    End If
    NoCodeCommentry LItem, lngIndex, True
  End With
  SetCurrentLSVLine frm_CodeFixer.lsvAllControls, SaveCurListPos

End Sub

Private Sub FixNewlyIndexed(ByVal strTarget As String, _
                            strHomeComp As String, _
                            StrCap As String, _
                            Optional ByVal InsertEmptyCaseStructure As Boolean = False)

  'v3.0.0 updated to add Caption after Case statement as a comment
  
  Dim G                As Long
  Dim lngPrevFindLine  As Long
  Dim GetCodeLine      As Long
  Dim StartLine        As Long
  Dim lngLegalEnd      As Long
  Dim LBPos            As Long
  Dim strIndexUpdate   As String
  Dim strDeleteWarning As String
  Dim strEvent         As String
  Dim strComment       As String
  Dim StrOldParam      As String
  Dim strCorrect       As String
  Dim strError         As String
  Dim L_CodeLine       As String
  Dim strLegalEnd      As String
  Dim LModule          As CodeModule
  Dim Proj             As VBProject
  Dim Comp             As VBComponent
  Dim LUpdated         As Boolean
  Dim arrTest          As Variant

  arrTest = Array("()", ContMark)
  If InsertEmptyCaseStructure Then
    strCorrect = strTarget
    strError = strTarget & "(0)_"
   Else
    strError = strTarget
    strCorrect = "DELETE_ME_INDEXED_VERSION_CREATED_" & CntrlDesc(LngCurrentControl).CDOldName
  End If
  strIndexUpdate = WARNING_MSG & "Procedure converted to use Indexed parameter." & vbNewLine & _
   RGSignature & "Check following code for lines which could be removed from the " & vbNewLine & _
   RGSignature & "Select Structure and applied before or after the structure" & vbNewLine & _
   RGSignature & "OR Repeated Dims which should be removed." & vbNewLine & _
   "Select Case Index" & vbNewLine & _
   "Case 0" & IIf(LenB(StrCap), "'" & StrCap, vbNullString) & vbNewLine
  strDeleteWarning = WARNING_MSG & "Delete this Procedure. It is no longer needed. (Code Fixer will comment it out on next run)" & vbNewLine & _
   WARNING_MSG & "Procedure converted to use Indexed parameter."
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        Set LModule = Comp.CodeModule
        StartLine = 1
        If LModule.Find(strError, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
          lngPrevFindLine = StartLine
          Do
            If StartLine < lngPrevFindLine Then
              Exit Do
            End If
            If LModule.Find(strError, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
              L_CodeLine = LModule.Lines(StartLine, 1)
              If ExtractCode(L_CodeLine, strComment) Then
                strEvent = Mid$(L_CodeLine, InStr(L_CodeLine, strError) + Len(strError))
                LBPos = InStr(strEvent, LBracket)
                If LBPos Then
                  strEvent = Left$(strEvent, LBPos - 1)
                End If
                StrOldParam = Mid$(L_CodeLine, (InStr(L_CodeLine, strEvent) + Len(strEvent)))
                'v 2.1.5 this makes sure that multiline parameters are collected properly
                G = StartLine + 1
                Do While HasLineCont(StrOldParam)
                  StrOldParam = StrOldParam & vbNewLine & LModule.Lines(G, 1)
                  G = G + 1
                Loop
                If CntrlDesc(LngCurrentControl).CDForm = Comp.Name Then
                  L_CodeLine = Replace$(L_CodeLine, strError, strCorrect & "_")
                  'v 2.1.5 stop double underscores created in code
                  L_CodeLine = Replace$(L_CodeLine, strCorrect & "__", strCorrect & "_")
                  If IsInArray(Right$(L_CodeLine, 2), arrTest) Or isProcHead(L_CodeLine) Then
                    LUpdated = True
                  End If
                  If InsertEmptyCaseStructure Then
                    If Right$(L_CodeLine, 2) = "()" Then
                      L_CodeLine = Replace$(L_CodeLine, "()", "(Index As Integer)")
                     ElseIf Right$(L_CodeLine, 2) = ContMark Then
                      L_CodeLine = Replace$(L_CodeLine, LBracket, "(Index As Integer,", , 1)
                     ElseIf isProcHead(L_CodeLine) Then
                      L_CodeLine = Replace$(L_CodeLine, LBracket, "(Index As Integer,", , 1)
                     Else
                      If InStr(L_CodeLine, strCorrect & "_" & strEvent) Then
                        If SmartLeft(L_CodeLine, "Call " & strCorrect) Then
                          L_CodeLine = L_CodeLine & strInBrackets(ExtractNumber(strError))
                         Else
                          L_CodeLine = L_CodeLine & SngSpace & ExtractNumber(strError)
                        End If
                       Else
                        L_CodeLine = Replace$(L_CodeLine, strCorrect, strCorrect & strInBrackets(ExtractNumber(strError)))
                      End If
                      LUpdated = True
                    End If
                    If LUpdated Then
                      If Not HasLineCont(L_CodeLine) Then
                        L_CodeLine = L_CodeLine & IIf(isProcHead(L_CodeLine), vbNewLine & strIndexUpdate, vbNullString)
                        LModule.ReplaceLine StartLine, L_CodeLine & strComment
                       Else
                        LModule.ReplaceLine StartLine, L_CodeLine
                        SafeInsertModule LModule, StartLine, strComment & IIf(isProcHead(L_CodeLine), vbNewLine & _
                         strIndexUpdate, "")
                      End If
                      LUpdated = False
                    End If
                    If isProcHead(L_CodeLine) Then
                      lngLegalEnd = StartLine
                      strLegalEnd = GetLegalEnd(L_CodeLine)
                      Do Until SmartLeft(LModule.Lines(lngLegalEnd, 1), strLegalEnd)
                        lngLegalEnd = lngLegalEnd + 1
                      Loop
                      LModule.ReplaceLine lngLegalEnd, RGSignature & "INSERT CODE HERE" & vbNewLine & _
                       "End Select" & vbNewLine & _
                       LModule.Lines(lngLegalEnd, 1)
                    End If
                   Else
                    If LUpdated Then
                      If Not HasLineCont(L_CodeLine) Then
                        L_CodeLine = L_CodeLine & vbNewLine & strDeleteWarning
                        LModule.ReplaceLine StartLine, L_CodeLine
                        GetCodeLine = StartLine + 1
                       Else
                        LModule.ReplaceLine StartLine, L_CodeLine
                        SafeInsertModule LModule, StartLine, strDeleteWarning
                        GetCodeLine = GetSafeInsertLine(LModule, StartLine)
                      End If
                      GetorDeleteProcedureCode StartLine, GetCodeLine, LModule, L_CodeLine, strError, strEvent, StrOldParam, strHomeComp, StrCap
                      LUpdated = False
                    End If
                  End If
                End If
              End If
            End If
            StartLine = StartLine + 1
            lngPrevFindLine = StartLine
          Loop While LModule.Find(strError, StartLine, 1, -1, -1, False, False, False)
        End If
      End If
    Next Comp
  Next Proj

End Sub

Private Sub FixNewRoutine(ByVal strProcType As String, _
                          ByVal StrFindNewProc As String, _
                          ByVal StrInsert As String, _
                          ByVal StrInsertAt As String, _
                          ByVal strOldParams As String, _
                          ByVal StrScope As String, _
                          ByVal strHomeComp As String)

  Dim LModule     As CodeModule
  Dim StartLine   As Long
  Dim Proj        As VBProject
  Dim Comp        As VBComponent
  Dim strFindThis As String
  Dim Runs        As Long
  Dim I           As Long

  If strProcType = "Property" Then ' this should never hit but someone might find a way to code it
    Runs = 3
   Else
    Runs = 1
  End If
  For I = 1 To Runs
    If Runs = 2 Then
      Select Case I
       Case 1
        strFindThis = strProcType & " Let " & StrFindNewProc
       Case 2
        strFindThis = strProcType & " Get " & StrFindNewProc
       Case 3
        strFindThis = strProcType & " Set " & StrFindNewProc
      End Select
     Else
      strFindThis = strProcType & SngSpace & StrFindNewProc
    End If
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If LenB(Comp.Name) Then
          Set LModule = Comp.CodeModule
NewCodeInsert:
          StartLine = 1
          If LModule.Find(strFindThis, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
            If LModule.Find(StrInsertAt, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
              LModule.ReplaceLine StartLine, StrInsert & vbNewLine & LModule.Lines(StartLine, 1)
              Exit Sub
            End If
           Else
            'the original control doesn't have this event so build new one
            If Comp.Name = strHomeComp Then
              If strProcType <> "Property" Then
                LModule.InsertLines LModule.CountOfLines + 1, StrScope & SngSpace & strFindThis & "(Index as integer" & IIf(Len(strOldParams) And strOldParams <> "()", CommaSpace & LStrip(strOldParams, LBracket), RBracket)
                SafeInsertModule LModule, LModule.CountOfLines + 1, WARNING_MSG & "This Event is not called by all controls in the control array." & vbNewLine & _
                 "Select Case Index" & vbNewLine & _
                 StrInsertAt & vbNewLine & _
                 "End Select" & vbNewLine & _
                 "End " & strProcType
                GoTo NewCodeInsert
              End If
            End If
          End If
        End If
      Next Comp
    Next Proj
  Next I

End Sub

Private Sub ForceUpDateImageListLink()

  Dim arrTmp           As Variant
  Dim ArrCtrl          As Variant
  Dim ArrCtrlID        As Variant
  Dim I                As Long
  Dim J                As Long
  Dim oldTargetControl As VBControl
  Dim Comp             As VBComponent
  Dim lngCtrolID       As Long
  Dim vbf              As VBForm

  'If a control has a 'ListImage' property it is automatically updated when you rename the listimage
  'but if it has 'Icons' or 'SmallIcons' properties but no 'ListImage' property assigned then
  'it does not auto update.
  'This routine assumes that it did not update automatically and does it
  arrTmp = Split(CntrlDesc(LngCurrentControl).CDImageListLinkedTo, ",")
  For I = LBound(arrTmp) To UBound(arrTmp)
    ArrCtrl = Split(arrTmp(I), ":")
    ArrCtrlID = Split(ArrCtrl(0), "|")
    lngCtrolID = CntrlDescMember2(ArrCtrlID(2), ArrCtrlID(1), ArrCtrlID(0))
    Set Comp = GetComponent(CntrlDesc(lngCtrolID).CDProj, CntrlDesc(lngCtrolID).CDForm)
    If IsComponent_ControlHolder(Comp) Then
      ActivateDesigner Comp, vbf, False
      If vbf Is Nothing Then
        ActivateDesigner Comp, vbf, True
      End If
      Set oldTargetControl = GetControlItemFromName(ArrCtrlID(2), vbf.VBControls)
      If Not oldTargetControl Is Nothing Then
        For J = 1 To UBound(ArrCtrl)
          oldTargetControl.Properties(ArrCtrl(J)).Value.Item("name").Value = CntrlDesc(LngCurrentControl).CDName
        Next J
      End If
    End If
  Next I

End Sub

Public Function GetBadNameType(ByVal lngIndex As Long, _
                               Optional ByVal IncludeSingleton As Boolean = True) As Long

  Dim I As Long

  With CntrlDesc(lngIndex)
    If isRefLibVBCommands(.CDName, False) Then
      GetBadNameType = BNCommand
     ElseIf InQSortArray(ArrQVBStructureWords, .CDName) Then
      GetBadNameType = BNStructural
     ElseIf InQSortArray(ArrQVBReservedWords, .CDName) Then
      GetBadNameType = BNReserve
     ElseIf IsControlProperty(.CDName) Then
      GetBadNameType = BNKnown
     ElseIf .CDName = .CDClass Then
      ' thanks Georg Veichtlbauer I moved this here so more serious porblems got detected first
      GetBadNameType = BNClass
     ElseIf IsDeclaredVariable(.CDName) Then
      If Not isEvent(.CDName) Then
        GetBadNameType = BNVariable
      End If
     ElseIf IsProcedureName(.CDName) Then
      GetBadNameType = BNProc
     ElseIf Len(.CDName) = 1 Then
      GetBadNameType = BNSingle
     ElseIf UsingDefVBName(.CDName, .CDClass) Then
      GetBadNameType = BNDefault
    End If
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDForm <> .CDForm Then ' only check on wrong forms
        If CntrlDesc(I).CDName = .CDName Then
          GetBadNameType = GetBadNameType + BNMultiForm
          Exit For
        End If
      End If
    Next I
    If IncludeSingleton Then
      'this optional value lets the edit control names to work without reporting the error
      'when renaming an array of controls(the first control will have this problem)
      If SingletonCtrlArray(lngIndex) Then
        GetBadNameType = GetBadNameType + BNSingletonArray
      End If
    End If
  End With

End Function

Public Function GetComponent(ByVal strPrj As String, _
                             ByVal StrFrm As String) As VBComponent

  Dim Proj As VBProject
  Dim Comp As VBComponent

  If Len(strPrj) Then
    Set GetComponent = VBInstance.VBProjects.Item(strPrj).VBComponents.Item(StrFrm)
   Else
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If StrFrm = Comp.Name Then
          Set GetComponent = Comp
          GoTo NormalExit
        End If
      Next Comp
    Next Proj
  End If
NormalExit:

End Function

Private Function GetCtrlDescArrayIndex(ByVal strName As String, _
                                       ByVal StrFrm As String) As Long

  Dim I        As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDForm = StrFrm Then
        If CntrlDesc(I).CDName = strName Then
          If CntrlDesc(I).CDIndex = -1 Then
            GetCtrlDescArrayIndex = I
            Exit For 'unction
          End If
        End If
      End If
    Next I
  End If

End Function

Public Function GetHiddenDescriptorIndex(Lsv As ListView, _
                                         ByVal LngColnumber As Long) As Long

  GetHiddenDescriptorIndex = CLng(Lsv.ListItems(Lsv.SelectedItem.Index).SubItems(LngColnumber))

End Function

Private Function GetLegalEnd(ByVal strCode As String) As String

  If ContainsWholeWord(strCode, "Sub") Then
    GetLegalEnd = "End Sub"
   ElseIf ContainsWholeWord(strCode, "Property") Then
    GetLegalEnd = "End Property"
   ElseIf ContainsWholeWord(strCode, "Function") Then
    GetLegalEnd = "End Function"
  End If

End Function

Private Sub GetorDeleteProcedureCode(ByVal StartLine As Long, _
                                     ByVal lngCodeLine As Long, _
                                     LModule As CodeModule, _
                                     L_CodeLine As String, _
                                     strError As String, _
                                     strEvent As String, _
                                     StrOldParam As String, _
                                     strHomeComp As String, _
                                     StrCap As String)

  Dim StrProcCode As String
  Dim lngLegalEnd As Long
  Dim strLegalEnd As String
  Dim StrScope    As String

  lngLegalEnd = lngCodeLine + 2
  strLegalEnd = GetLegalEnd(L_CodeLine)
  StrProcCode = "Case " & ExtractNumber(strError) & IIf(LenB(StrCap), "'" & StrCap, vbNullString)
  Do Until SmartLeft(LModule.Lines(lngLegalEnd, 1), strLegalEnd)
    StrProcCode = StrProcCode & vbNewLine & LModule.Lines(lngLegalEnd, 1)
    lngLegalEnd = lngLegalEnd + 1
  Loop
  If frm_CodeFixer.chkDelOldCode.Value = vbChecked Then
    LModule.DeleteLines StartLine, lngLegalEnd - StartLine + 1
  End If
  StrScope = LeftWord(L_CodeLine)
  If Not ArrayMember(StrScope, "Public", "Private", "Static", "Friend") Then
    StrScope = vbNullString
  End If
  FixNewRoutine WordInString(strLegalEnd, 2), Left$(strError, InStr(strError, LBracket) - 1) & strEvent, StrProcCode, RGSignature & "INSERT CODE HERE", StrOldParam, StrScope, strHomeComp

End Sub

Public Function GetProcEndLine(cMod As CodeModule, _
                               pStartLine As Long, _
                               Optional ProcName As String, _
                               Optional PKType As Long) As Long

  'v2.9.6 corrected so that it findes end of final procedure

  GetProcEndLine = GetProcStartLine(cMod, pStartLine, ProcName, PKType)
  Do
    GetProcEndLine = GetProcEndLine + 1
    If GetProcEndLine > cMod.CountOfLines Then
      GetProcEndLine = cMod.CountOfLines + 1
      Exit Do
    End If
  Loop While cMod.ProcOfLine(GetProcEndLine, PKType) = ProcName

End Function

Public Function GetProcStartLine(cMod As CodeModule, _
                                 ByVal Sline As Long, _
                                 Optional ProcName As String, _
                                 Optional PKType As Long) As Long

  Dim lngLineNo As Long

  'v2.8.5 restructure for greater safety
  On Error Resume Next
  ProcName = cMod.ProcOfLine(Sline, vbext_pk_Proc)
  If LenB(ProcName) Then
    lngLineNo = cMod.PRocStartLine(ProcName, vbext_pk_Proc)
    If lngLineNo Then
      GetProcStartLine = lngLineNo
      PKType = vbext_pk_Proc
      GoTo Done
    End If
  End If
  ProcName = cMod.ProcOfLine(Sline, vbext_pk_Let)
  If LenB(ProcName) Then
    lngLineNo = cMod.PRocStartLine(ProcName, vbext_pk_Let)
    If lngLineNo Then
      GetProcStartLine = lngLineNo
      PKType = vbext_pk_Let
      GoTo Done
    End If
  End If
  ProcName = cMod.ProcOfLine(Sline, vbext_pk_Get)
  If LenB(ProcName) Then
    lngLineNo = cMod.PRocStartLine(ProcName, vbext_pk_Get)
    If lngLineNo Then
      GetProcStartLine = lngLineNo
      PKType = vbext_pk_Get
      GoTo Done
    End If
  End If
  ProcName = cMod.ProcOfLine(Sline, vbext_pk_Set)
  If LenB(ProcName) Then
    lngLineNo = cMod.PRocStartLine(ProcName, vbext_pk_Set)
    If lngLineNo Then
      GetProcStartLine = lngLineNo
      PKType = vbext_pk_Set
    End If
  End If
  'dummy for detecting that item is in Declaration section
  ProcName = "(Declarations)"
  GetProcStartLine = 0
Done:
  On Error GoTo 0

End Function

Public Function GetTag(Lsv As ListView) As Long

  GetTag = CLng(Lsv.ListItems(Lsv.SelectedItem.Index).Tag) '.SubItems(LngColnumber)

End Function

Public Function HasNoHungarianPrefix(ByVal lngIndex As Long, _
                                     strCorrectPrefix As String) As Boolean

  If ArrayPos(CntrlDesc(lngIndex).CDClass, StandardControl) > -1 Then
    strCorrectPrefix = StandardCtrPrefix(ArrayPos(CntrlDesc(lngIndex).CDClass, StandardControl))
  End If
  If Len(strCorrectPrefix) Then
    HasNoHungarianPrefix = Not SmartLeft(LCase$(CntrlDesc(lngIndex).CDName), strCorrectPrefix, False)
   Else
    HasNoHungarianPrefix = True
  End If

End Function

Private Function IllegalCharacterMsg(ByVal strTest As String) As String

  Dim strTmp   As String
  Dim strIlChr As String
  Dim I        As Long

  If IsNumeric(Left$(strTest, 1)) Then
    IllegalCharacterMsg = "Numerals are not legal at start of names."
    strTmp = AccumulatorString(strTmp, Left$(strTest, 1), CommaSpace)
  End If
  For I = 1 To Len(strTest)
    If IsPunctExcept(Mid$(strTest, I, 1), "_1234567890") Then
      strIlChr = Mid$(strTest, I, 1)
      If strIlChr = DQuote Then
        strIlChr = """"
      End If
      strTmp = AccumulatorString(strTmp, strIlChr, CommaSpace)
      Exit For
    End If
  Next I
  If Len(strIlChr) Then
    IllegalCharacterMsg = IllegalCharacterMsg & " Illegal Characters: " & strTmp
  End If

End Function

Public Function IllegalCharacters(ByVal strTest As String) As Boolean

  Dim I As Long

  If IsNumeric(Left$(strTest, 1)) Then
    IllegalCharacters = True
  End If
  For I = 1 To Len(strTest)
    If IsPunctExcept(Mid$(strTest, I, 1), "_1234567890") Then
      IllegalCharacters = True
      Exit For
    End If
  Next I

End Function

Private Function inLegalWith(ByVal TestLine As Long, _
                             LModule As CodeModule, _
                             ByVal strForm As String) As Boolean

  Dim StartLine   As Long
  Dim EndWithLine As Long
  Dim Targets     As Variant
  Dim K           As Long

  Targets = Array("With " & strForm, "With Me", "With UserControl", "With UserDocument")
  For K = LBound(Targets) To UBound(Targets)
    StartLine = 1
    Do
      If LModule.Find(Targets(K), StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
        EndWithLine = StartLine
        If LModule.Find("End With", EndWithLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
          If StartLine < TestLine Then
            If TestLine < EndWithLine Then
              inLegalWith = True
              Exit For
            End If
          End If
        End If
      End If
      If StartLine = 0 Then
        Exit Do
      End If
      If StartLine = 1 Then
        Exit Do
      End If
      StartLine = StartLine + 1
      If StartLine > TestLine Then
        If MultiLeft(LModule.Lines(StartLine, 1), True, "End Sub", "End Function", "End Property") Then
          Exit Do
        End If
      End If
    Loop Until StartLine > LModule.CountOfLines
  Next K

End Function

Private Function IsImageList(ByVal varName As Variant) As Boolean

  Dim I        As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDName = varName Then
        IsImageList = CntrlDesc(I).CDClass = "ImageList"
        Exit For 'unction
      End If
    Next I
  End If

End Function

Public Function IsThisControlDeletable(ByVal Cno As Long) As Boolean

  If bCtrlDescExists Then
    With CntrlDesc(Cno)
      If .CDUsage = 0 Then
        If LenB(.CDImageListLinkedTo) = 0 Then
          If LenB(.CDImageListLink) = 0 Then
            If .CDClass <> "Menu" Then
              If Not IsGraphic(.CDClass) Then
                If Not IsFileTool(.CDClass) Then
                  If Not .CDIsContainer Then
                    IsThisControlDeletable = True
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End With 'CntrlDesc(Cno)
  End If

End Function

Public Function IsWholeWord(ByVal strCode As String, _
                            ByVal strF As String, _
                            ByVal TPos As Long, _
                            Optional ByVal bIgnoreUnderScore As Boolean = False) As Boolean

  If TPos = 1 Then
    If strCode = strF Then
      IsWholeWord = True
     Else
      IsWholeWord = IsPunct(Mid$(strCode, TPos + Len(strF), 1))
    End If
   ElseIf IsPunct(Mid$(strCode, TPos - 1, 1)) Then
    If SmartRight(strCode, strF) And TPos = Len(strCode) - Len(strF) + 1 Then
      IsWholeWord = True
     ElseIf IsPunct(Mid$(strCode, TPos + Len(strF), 1)) Then
      If bIgnoreUnderScore Then
        IsWholeWord = Mid$(strCode, TPos + Len(strF), 1) <> "_"
       Else
        IsWholeWord = True
      End If
    End If
  End If

End Function

Public Function NoHungarianPrefix(ByVal strObj As String) As String

  Dim HPref As String

  NoHungarianPrefix = strObj ' safety
  If ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl) > -1 Then
    HPref = StandardCtrPrefix(ArrayPos(CntrlDesc(LngCurrentControl).CDClass, StandardControl))
  End If
  If Len(HPref) Then
    If strObj <> CntrlDesc(LngCurrentControl).CDClass Then
      If SmartLeft(strObj, HPref, False) Then
        NoHungarianPrefix = Mid$(NoHungarianPrefix, Len(HPref) + 1)
      End If
    End If
  End If

End Function

Private Function PreviousWith(LMod As CodeModule, _
                              StartLine As Long, _
                              Targets As Variant) As Boolean

  Dim I         As Long
  Dim TopOfProc As Long
  Dim strTest   As String
  Dim TmpMemb   As Variant

  TopOfProc = GetProcStartLine(LMod, StartLine)
  For I = StartLine To TopOfProc Step -1
    If SmartLeft(Trim$(LMod.Lines(I, 1)), "With ") Then
      strTest = Trim$(LMod.Lines(I, 1))
      For Each TmpMemb In Targets
        If SmartLeft(strTest, TmpMemb) Then
          PreviousWith = True
          Exit For 'unction
        End If
      Next TmpMemb
      Exit For 'unction
    End If
  Next I

End Function

Private Function ProhibitedControlUpdate(ByVal strCode As String, _
                                         ByVal strCtrl As String) As Boolean

  'this is a kludge for certain very destructive rename control operations
  'thanks to Tom Law & Georg Veichtlbauer  who both reported this problem
  'basicly just stops renaming from occurring

  Select Case strCtrl
   Case "Exit"
    ProhibitedControlUpdate = IsInArray(WordInString(strCode, 2), ArrFuncPropSub)
   Case "Line"
    ProhibitedControlUpdate = WordInString(strCode, 2) = "Input"
  End Select
  'End If

End Function

Public Sub RenameControlsStandard(strOldName As String, _
                                  strNewName As String, _
                                  Optional bSingletonSwitch As Boolean = False)

  ChangeControlInWith strOldName, strNewName
  ChangeControlOnModules strOldName, strNewName, bSingletonSwitch

End Sub

Private Sub ReWriteSelectedLine(Optional ByVal lngNewIndex As Long = -2)

  Dim LItem    As ListItem
  Dim lngIndex As Long

  Set LItem = frm_CodeFixer.lsvAllControls.SelectedItem
  lngIndex = CLng(LItem.Tag)
  With CntrlDesc(lngIndex)
    LItem.Text = .CDProj
    LItem.SubItems(1) = .CDForm
    If lngNewIndex <> -2 Then
      .CDIndex = lngNewIndex
    End If
    .CDFullName = .CDName & IIf(.CDIndex > -1, strInBrackets(.CDIndex), vbNullString)
    LItem.SubItems(2) = .CDFullName
    LItem.SubItems(3) = .CDCaption
    .CDBadType = GetBadNameType(lngIndex, False) ' GetBadNameType(lngIndex)
    NoCodeCommentry LItem, lngIndex, True
  End With

End Sub

Private Function WordIncluding(ByVal strCode As String, _
                               ByVal TPos As Long) As String

  Dim I        As Long
  Dim EndPos   As Long
  Dim StartPos As Long

  If TPos > 0 Then
    For I = TPos To Len(strCode)
      If Mid$(strCode, I, 1) = SngSpace Then
        EndPos = I
        Exit For
      End If
    Next I
    If EndPos = 0 Then
      EndPos = Len(strCode)
    End If
    For I = TPos To 1 Step -1
      If Mid$(strCode, I, 1) = SngSpace Then
        StartPos = I
        Exit For
      End If
    Next I
    If StartPos = 0 Then
      StartPos = 1
    End If
    WordIncluding = Trim$(Mid$(strCode, StartPos, EndPos - StartPos))
   Else
    WordIncluding = vbNullString
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:13:19 AM) 26 + 1373 = 1399 Lines Thanks Ulli for inspiration and lots of code.

