Attribute VB_Name = "mod_ControlTabCode2"
Option Explicit
Private Type TopLeft
  Top                    As Long
  Left                   As Long
  TabIndex               As Long
  ZorderVal              As Long
End Type
Private ctl()            As TopLeft
Private Const vbGrey     As Long = 8355711    ' Control comments that need no fix

Public Sub ActivateDesigner(Cmp As VBComponent, _
                            vbfrm As VBForm, _
                            Optional ByVal BShowSelection As Boolean = True)

  Dim ErrorTrap As Boolean

  On Error GoTo ForceHard
ReTry:
  With Cmp
    If BShowSelection Then
      .CodeModule.CodePane.Show
      .Activate
    End If
    Set vbfrm = .Designer
  End With 'Comp

Exit Sub

ForceHard:
  'v2.4.9 increased error detection for the odd form that won't load
  If Not ErrorTrap Then
    BShowSelection = True
    ErrorTrap = True
    GoTo ReTry
   Else
    mObjDoc.Safe_MsgBox "An Error (" & Err.Number & ")" & vbNewLine & _
                    " '" & Err.Description & "'" & vbNewLine & _
                    " has occured while trying to load the component '" & Cmp.Name & "'.", vbCritical
  End If

End Sub

Public Sub AutoFixBadNames()

  Dim I              As Long
  Dim arrTarget      As Variant
  Dim LItem          As ListItem
  Dim SaveSort       As Boolean
  Dim lngWorkingDesc As Long
  Dim MyhourGlass    As cls_HourGlass
  Dim arrTest        As Variant

  arrTest = Array(BNStructural, BNClass, BNCommand, BNMultiForm)
  Set MyhourGlass = New cls_HourGlass
  WarningLabel "Auto Renaming(Caption)...", vbRed
  arrTarget = QuickSortArray(Split("CommandButton,Label,Frame,OptionButton,CheckBox,Menu", ","))
  With frm_CodeFixer
    SaveSort = .lsvAllControls.Sorted
    .lsvAllControls.Sorted = False
    For I = 1 To .lsvAllControls.ListItems.Count
      SetCurrentLSVLine .lsvAllControls, I
      Set LItem = .lsvAllControls.ListItems(.lsvAllControls.SelectedItem.Index)
      lngWorkingDesc = GetTag(.lsvAllControls)
      If CntrlDesc(lngWorkingDesc).CDBadType <> 0 Then
        DoBadSelect
        If CntrlDesc(lngWorkingDesc).CDIndex = "-1" Then
          ' skip arrays or else they get named after last member of array
          If InQSortArray(arrTarget, CntrlDesc(lngWorkingDesc).CDClass) Then
            If CntrlDesc(lngWorkingDesc).CDClass = "Menu" Then
              If IsInArray(CntrlDesc(lngWorkingDesc).CDBadType, arrTest) Then
                SetNewName "mnu" & CntrlDesc(lngWorkingDesc).CDName
                EditControlName "mnu" & CntrlDesc(lngWorkingDesc).CDName
                LItem.SubItems(2) = "mnu" & CntrlDesc(lngWorkingDesc).CDName
                CntrlDesc(lngWorkingDesc).CDName = "mnu" & CntrlDesc(lngWorkingDesc).CDName
                LItem.ListSubItems(4).ForeColor = vbGreen
                LItem.SubItems(4) = "FIXED"
              End If
             Else
              With .lblOldName(3)
                If Len(.Caption) Then
                  SetNewName .Caption
                  EditControlName .Caption
                  LItem.SubItems(2) = CntrlDesc(lngWorkingDesc).CDName
                  LItem.ListSubItems(4).ForeColor = vbGreen
                  LItem.SubItems(4) = "FIXED"
                End If
              End With
            End If
          End If
         Else
          If CntrlDesc(lngWorkingDesc).CDBadType = BNSingletonArray Then
            ChangeControlNameSingleton CntrlDesc(lngWorkingDesc).CDName, CntrlDesc(lngWorkingDesc).CDName & "_" & CntrlDesc(lngWorkingDesc).CDIndex
            LItem.SubItems(2) = CntrlDesc(lngWorkingDesc).CDName & "_" & CntrlDesc(lngWorkingDesc).CDIndex
            LItem.ListSubItems(4).ForeColor = vbGreen
            LItem.SubItems(4) = "FIXED"
          End If
        End If
      End If
    Next I
    .lsvAllControls.Sorted = SaveSort
  End With

End Sub

Public Sub AutoFixPrefix()

  Dim I                As Long
  Dim LBPos            As Long
  Dim strCorrectPrefix As String
  Dim strCName         As String
  Dim strNewName       As String
  Dim StrArrayName     As String
  Dim LItem            As ListItem
  Dim SaveSort         As Boolean
  Dim MyhourGlass      As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  With frm_CodeFixer
    WarningLabel "Auto Renaming(Prefix)...", vbRed
    I = .lsvAllControls.ListItems.Count
    SaveSort = .lsvAllControls.Sorted
    .lsvAllControls.Sorted = False
    For I = 1 To .lsvAllControls.ListItems.Count
      SetCurrentLSVLine .lsvAllControls, I
      DoAllSelect
      Set LItem = .lsvAllControls.ListItems(I)
      strCName = LItem.SubItems(2)
      LBPos = InStr(strCName, LBracket)
      If LBPos Then
        strCName = Left$(strCName, LBPos - 1)
        If StrArrayName <> strCName Then
          StrArrayName = strCName
         Else
          GoTo Skip2ndArrayMembers
        End If
       Else
        StrArrayName = vbNullString
      End If
      If HasNoHungarianPrefix(Val(LItem.Tag), strCorrectPrefix) Then
        strNewName = strCorrectPrefix & strCName
        FixControlCase strNewName
        SetNewName strNewName
        EditControlName strNewName
        LItem.SubItems(2) = strNewName
Skip2ndArrayMembers:
       Else
        ' fix casing
        strNewName = strCName
        FixControlCase strNewName
        If strCName <> strNewName Then
          SetNewName strNewName
          EditControlName strNewName
        End If
      End If
    Next I
    .lsvAllControls.Sorted = SaveSort
  End With

End Sub

Public Sub ChangeControlNameSingleton(ByVal strCtrl As String, _
                                      ByVal strNewName As String)

  Dim vbf                 As VBForm
  Dim Targetcontrol       As VBControl
  Dim ArrayCntrls         As Variant
  Dim ctrla               As Variant
  Dim Comp                As VBComponent
  Dim lngTmpCtrlDescIndex As Long
  Dim lngOldIndex         As Long

  Set Comp = GetComponent(CntrlDesc(LngCurrentControl).CDProj, CntrlDesc(LngCurrentControl).CDForm)
  If IsComponent_ControlHolder(Comp) Then
    ActivateDesigner Comp, vbf, False
    If vbf Is Nothing Then
      ActivateDesigner Comp, vbf, True
    End If
    On Error GoTo ChangeFail
    With vbf
      DoEvents
      'single controls
      ArrayCntrls = GetControlArrayFromName(strCtrl, .VBControls)
      If UBound(ArrayCntrls) = 0 Then
        For Each ctrla In ArrayCntrls
          Set Targetcontrol = GetControlItemFromName(ctrla, .VBControls)
          If Not Targetcontrol Is Nothing Then
            lngTmpCtrlDescIndex = MatchCtrlToDescArrayIndex(Targetcontrol)
            If UBound(ArrayCntrls) = 0 Then
              With Targetcontrol
                'lngOldIndex = .Properties("index").Value
                .Properties("name").Value = strNewName
                lngOldIndex = .Properties("index").Value
                .Properties("index").Value = -1
                FixDeIndexed CntrlDesc(LngCurrentControl).CDForm, strCtrl & "_", strNewName
                'v3.0.7 improved replacement engine using the RenameArray switch
                RenameControlsStandard strNewName & strInBrackets(lngOldIndex), strNewName, True
              End With 'TargetControl
            End If
            If LCase$(ControlName(Targetcontrol)) <> LCase$(strCtrl) Then
              ImageListWarning Targetcontrol, Comp
              Comp.CodeModule.InsertLines 1, WARNING_MSG & WARNING_MSG & "Control" & strInSQuotes(ctrla, True) & "renamed to" & strInSQuotes(ControlName(Targetcontrol), True) & BadReasonComment(lngTmpCtrlDescIndex)
              ReWriteSelectedLineArrays lngTmpCtrlDescIndex, Targetcontrol
            End If
          End If
        Next ctrla
      End If
    End With
  End If
ChangeFail:
  'this routine depends on this error detector to reset for following errors

End Sub

Private Sub CollectSelectedControls(TargetFrame As VBControl, _
                                    Optional ByVal SelType As Boolean = True)

  Dim I    As Long

  'Separated out for clarity purposes
  For I = 1 To TargetFrame.ContainedVBControls.Count
    TargetFrame.ContainedVBControls.Item(I).InSelection = SelType
  Next I

End Sub

Private Sub CreatePictureBox(c As VBComponent, _
                             TargetFrame As VBControl, _
                             PicTarget As VBControl, _
                             ofset As Long, _
                             FixCount As Long, _
                             voffset As Long)

  On Error Resume Next
  With TargetFrame
    .InSelection = True
    .Collection.Add "VB.PictureBox" ', TargetFrame, false
    Set PicTarget = .Collection.Item(.Collection.Count)
    ofset = IIf(.Properties("borderstyle").Value = 0, 0, 100)
    'v2.3.2 new font handler for inserting pictures
    If LenB(.Properties("caption").Value) Then
      voffset = .Properties("font").Value.Item("size").Value * 7  ' 2
    End If
  End With 'TargetFrame
  'set inital properties 'Name/BorderStyle/Width/Height
  ' separate variable needed for later resets of control
  ' after it is cut from form
  With PicTarget
    .Properties("name").Value = "picCFXPBugFix" & c.Name
    .Properties("scalemode").Value = vbTwips 'fCodeFix.picMenuIcon.ScaleMode
    ReDim Preserve CntrlDesc(UBound(CntrlDesc) + 1) As ControlDescriptor
    With CntrlDesc(UBound(CntrlDesc))
      .CDName = PicTarget.Properties("name").Value
      .CDClass = PicTarget.ClassName
      If FixCount Then
        PicTarget.Properties("index").Value = FixCount
        .CDIndex = FixCount
       Else
        .CDIndex = PicTarget.Properties("index").Value
      End If
      .CDFullName = .CDName & IIf(.CDIndex > -1, strInBrackets(.CDIndex), vbNullString)
      If .CDIndex = 1 Then
        'then update record of previous new picture (if in same set)
        If CntrlDesc(UBound(CntrlDesc) - 1).CDName = .CDName Then
          CntrlDesc(UBound(CntrlDesc) - 1).CDIndex = 0
          CntrlDesc(UBound(CntrlDesc) - 1).CDFullName = CntrlDesc(UBound(CntrlDesc) - 1).CDFullName & "(0)"
        End If
      End If
      .CDForm = c.Name
      .CDProj = GetActiveProject.Name
      FixCount = FixCount + 1
    End With
    'next line is just an optical effect used in debugging for this code
    'but produces a nice dramatic effect so I left it in
    .Properties("visible").Value = True
    .Properties("backcolor").Value = TargetFrame.Properties("backcolor")
    .Properties("borderstyle").Value = 0
    'controls on MDIForms cannot have Width property set so skip here and
    'fix after pasting to Frame
    If c.Type <> vbext_ct_VBMDIForm Then
      .Properties("width").Value = TargetFrame.Properties("width") - (ofset * 2)
      .Properties("height").Value = TargetFrame.Properties("height") - (offset * 2.5) - voffset * 1.5
    End If
  End With 'PicTarget
  On Error GoTo 0

End Sub

Private Sub ExtractControlData(vbc As VBControl, _
                               strPrj As String, _
                               strMod As String, _
                               strName As String, _
                               lngIndex As Long)

  With vbc
    strPrj = .VBE.ActiveVBProject.Name
    strMod = .VBE.SelectedVBComponent.Name
    strName = .Properties("name").Value
    lngIndex = .Properties("index").Value
  End With

End Sub

Public Sub FixControlCase(strNewName As String)

  Dim I As Long

  For I = LBound(StandardCtrPrefix) To UBound(StandardCtrPrefix)
    If SmartLeft(strNewName, StandardCtrPrefix(I), False) Then
      If Len(strNewName) > 3 Then
        Mid$(strNewName, Len(StandardCtrPrefix(I)) + 1, 1) = UCase$(Mid$(strNewName, Len(StandardCtrPrefix(I)) + 1, 1))
        strNewName = LCase$(Left$(strNewName, 3)) & Mid$(strNewName, 4)
       Else
        strNewName = StrConv(strNewName, vbProperCase)
      End If
      Exit For
    End If
  Next I

End Sub

Private Sub FixDeIndexed(ByVal strForm As String, _
                         ByVal strError As String, _
                         ByVal strCorrect As String)

  Dim lngPrevFindLine As Long
  Dim Proj            As VBProject
  Dim Comp            As VBComponent
  Dim L_CodeLine      As String
  Dim LModule         As CodeModule
  Dim StartLine       As Long
  Dim strRemIndex     As String

  ' break Loop if the only codeline is a comment that prog should ignore
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        Set LModule = Comp.CodeModule
        StartLine = 1
        If LModule.Find(strError, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) Then
          lngPrevFindLine = StartLine
          'Do
          Do While LModule.Find(strError, StartLine, 1, -1, -1, False, False, True)
            L_CodeLine = LModule.Lines(StartLine, 1)
            If StartLine < lngPrevFindLine Then
              Exit Do
            End If
            If Not JustACommentOrBlank(L_CodeLine) Then
              If strForm = Comp.Name Then
                'v3.0.7 improved replacement engine
                If isProcHead(L_CodeLine) Then
                  L_CodeLine = Replace$(L_CodeLine, "Index As Integer,", SngSpace)
                  L_CodeLine = Replace$(L_CodeLine, "Index As Integer", SngSpace)
                  LModule.ReplaceLine StartLine, L_CodeLine
                  SafeInsertModule LModule, StartLine, WARNING_MSG & "(Index As Integer) removed"
                 Else
                  If IsNumeric(WordAfter(L_CodeLine, (WordAfter(L_CodeLine, strError)))) Then
                    strRemIndex = WordAfter(L_CodeLine, (WordAfter(L_CodeLine, strError)))
                    L_CodeLine = Replace$(L_CodeLine, strRemIndex, SngSpace)
                    LModule.ReplaceLine StartLine, L_CodeLine
                    SafeInsertModule LModule, StartLine, WARNING_MSG & " Index Reference removed"
                   ElseIf IsNumeric(WordAfter(L_CodeLine, (WordAfter(L_CodeLine, strCorrect)))) Then
                    strRemIndex = WordAfter(L_CodeLine, (WordAfter(L_CodeLine, strCorrect)))
                    L_CodeLine = Replace$(L_CodeLine, strRemIndex, SngSpace)
                    LModule.ReplaceLine StartLine, L_CodeLine
                    SafeInsertModule LModule, StartLine, WARNING_MSG & " Index Reference removed"
                   Else
                    SafeInsertModule LModule, StartLine, WARNING_MSG & " You must remove Index Reference by hand"
                  End If
                End If
              End If
            End If
            'End If
            StartLine = StartLine + 1
            If StartLine > LModule.CountOfLines Then
              Exit Do
            End If
            lngPrevFindLine = StartLine
          Loop
          'While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strError, L_CodeLine, StartLine)
        End If
      End If
    Next Comp
  Next Proj

End Sub

Public Function GetControlArrayFromName(ByVal CtrlName As String, _
                                        vbc As VBControls) As Variant

  Dim LocalVbc As VBControl
  Dim strTmp   As String

  'Because you cannot be sure where in the Controls collection a specific
  'control is this routine searches the collection by finding the matching name.
  For Each LocalVbc In vbc
    If LocalVbc.Properties("name").Value = CtrlName Then
      strTmp = AccumulatorString(strTmp, ControlName(LocalVbc))
    End If
  Next LocalVbc
  GetControlArrayFromName = Split(strTmp, ",")

End Function

Public Function GetControlArraySizeFromName(ByVal CtrlName As String, _
                                            vbc As VBControls) As Long

  Dim LocalVbc As VBControl

  'Because you cannot be sure where in the Controls collection a specific
  'control is this routine searches the collection by finding the matching name.
  'Only needs to return 1 any more and the test fails anyway
  For Each LocalVbc In vbc
    If LocalVbc.Properties("name").Value = CtrlName Then
      GetControlArraySizeFromName = GetControlArraySizeFromName + 1
      If GetControlArraySizeFromName > 1 Then
        Exit For
      End If
    End If
  Next LocalVbc

End Function

Public Function GetControlItemFromName(ByVal CtrlName As String, _
                                       vbc As VBControls) As VBControl

  Dim LocalVbc    As VBControl

  'Because you cannot be sure where in the Controls collection a specific
  'control is this routine searches the collection by finding the matching name.
  For Each LocalVbc In vbc
    If ControlName(LocalVbc) = CtrlName Then
      Set GetControlItemFromName = LocalVbc
      Exit For 'unction
    End If
  Next LocalVbc

End Function

Private Sub GetControlPositioning(vbf As VBForm)

  Dim I    As Long
  Dim Ctrl As Variant

  With vbf
    ReDim ctl(.SelectedVBControls.Count) As TopLeft
    For Each Ctrl In .SelectedVBControls
      On Error Resume Next
      I = I + 1
      With ctl(I)
        .Left = Ctrl.Properties.Item("left")
        .Top = Ctrl.Properties.Item("top")
        .TabIndex = Ctrl.Properties.Item("tabindex")
      End With 'ctl(I)
      On Error GoTo 0
    Next Ctrl
  End With

End Sub

Private Function HasToolBar(vbf As VBForm) As Boolean

  Dim I As Long

  For I = 0 To vbf.SelectedVBControls.Count - 1
    If vbf.SelectedVBControls(I).ClassName = "Toolbar" Then
      HasToolBar = True
      Exit For
    End If
  Next I

End Function

Public Sub ImageListWarning(tCtl As VBControl, _
                            Cmp As VBComponent)

  If tCtl.ClassName = "ImageList" Then
    Cmp.CodeModule.InsertLines 1, WARNING_MSG & "renaming Imagelists may cause Error '35613'" & vbNewLine & _
     RGSignature & "'ImageList must be initialized before it can be used'" & vbNewLine & _
     RGSignature & "Occurs with ListView with no Column Header link and TreeView if any other Property is changed." & vbNewLine & _
     RGSignature & "Open Properties window of the Control(s) and reset Imagelist link(s)."
  End If

End Sub

Public Sub InsertXPPic2Frame2()

  
  Dim strPrevForm       As String    ' Reset test for restarting count on modules
  Dim MyhourGlass       As cls_HourGlass
  Dim PicTarget         As VBControl
  Dim Comp              As VBComponent
  Dim TargetFrame       As VBControl
  Dim vbf               As VBForm
  Dim strForm           As String
  Dim strProj           As String
  Dim strFrame          As String
  Dim LOffset           As Long
  Dim LItem             As ListItem
  Dim OriginalScale     As Long
  Dim OriginalSclLeft   As Single
  Dim OriginalSclTop    As Single
  Dim OriginalSclHeight As Single
  Dim OriginalSclWidth  As Single
  Dim dofix             As Boolean
  Dim I                 As Long
  Dim voffset           As Long      'v2.3.2 new font handler for inserting pictures
  Dim FixCounter        As Long

  ' this is used to force Index of pic box (a problem with containers of containers cofuses CF and it cant track index properly
  Set MyhourGlass = New cls_HourGlass
  'v2.9.6 new message
  mObjDoc.Safe_MsgBox "WARNING" & vbNewLine & _
                    "If your code resizes the frames in code you will need to add code to handle the PictureBoxes." & vbNewLine & _
                    "This does not apply if you use a class/OCX that reads the controls off the form", vbInformation, "WARNING"
  'This is the heart of my cure for the XP Frame problem
  'It is slow because it makes several sweeps through
  'the whole form to make sure it is dealing safely with the correct controls
  For I = 1 To frm_CodeFixer.lsvAllControls.ListItems.Count
    With frm_CodeFixer
      SetCurrentLSVLine .lsvAllControls, I
      'safety stuff
      Set LItem = .lsvAllControls.SelectedItem
      With LItem
        strProj = .Text
        strForm = .SubItems(1)
        strFrame = CntrlDesc(Val(.Tag)).CDFullName
        dofix = .SubItems(4) = "XP Style Frame Bug"
        If Not dofix Then
          'this copes with a poorly named controls which also have xpbug
          dofix = NeedsXPBugFix(strProj, strForm, strFrame)
        End If
      End With 'LItem
    End With 'fCodeFix
    If dofix Then
      If strForm <> strPrevForm Then
        FixCounter = 0
        strPrevForm = strForm
      End If
      Set Comp = GetComponent(strProj, strForm)
      ActivateDesigner Comp, vbf
      ReadWriteScale Comp, True, OriginalScale, OriginalSclWidth, OriginalSclHeight, OriginalSclTop, OriginalSclLeft
      On Error GoTo CutFail
      With vbf
        DeselectAll .VBControls
        DoEvents
        Set TargetFrame = GetControlItemFromName(strFrame, .VBControls)
        'collect controls to move
        CollectSelectedControls TargetFrame
        GetControlPositioning vbf
        DoEvents
        If .SelectedVBControls.Count Then
          'v2.7.3 warn that there is a toolbar
          If HasToolBar(vbf) Then
            mObjDoc.Safe_MsgBox "This frame contains at least one ToolBar control." & vbNewLine & _
                    "Inserting the XP Frame bug protection will trigger the Toolbar Wizard." & vbNewLine & _
                    "When it appears just Click its 'Cancel' button to continue.", vbOKOnly, "XP Frame bug Fix"
          End If
          'cut selected controls
          .SelectedVBControls.Cut
          'Create the new picturebox
          CreatePictureBox Comp, TargetFrame, PicTarget, LOffset, FixCounter, voffset
          'Paste controls to picturebox
          DoEvents
          PasteToOneControl vbf, PicTarget
          'restore data to controls
          SetControlPositioning vbf, LOffset, voffset
          SelectOneControl vbf, PicTarget
          ''Cut and paste Picturebox to Frame
          .SelectedVBControls.Cut
          PasteToOneControl vbf, GetControlItemFromName(strFrame, .VBControls) 'TargetFrame
          'Fit Picture to frame
          'This is similar to Function PictureToFrame but because we are dealing with
          'VBControls and other code we can't use that function
          Set PicTarget = GetControlItemFromName(CntrlDesc(UBound(CntrlDesc)).CDFullName, .VBControls)
          Set TargetFrame = GetControlItemFromName(strFrame, .VBControls)
          With PicTarget
            .Properties("backcolor").Value = TargetFrame.Properties("backcolor")
            .Properties("left").Value = LOffset  'PicTarget.Properties("left") + LOffset
            .Properties("Top").Value = LOffset * 1.75 + voffset * 1.75
            'controls on MDIForms cannot have Width property set
            'so it is set here instead
            If TargetFrame.Properties("width").Value - (LOffset * 2) > 0 Then
              .Properties("width").Value = TargetFrame.Properties("width").Value - (LOffset * 2)
             Else
              .Properties("width").Value = TargetFrame.Properties("width").Value
            End If
            If TargetFrame.Properties("height").Value - (LOffset * 2.5) - voffset * 1.5 > 0 Then
              .Properties("height").Value = TargetFrame.Properties("height").Value - (LOffset * 2.5) - voffset * 1.5
             ElseIf TargetFrame.Properties("height").Value - (LOffset * 2.5) - voffset * 1.5 > 0 Then
              .Properties("height").Value = TargetFrame.Properties("height").Value - (LOffset * 2.5)
             Else
              .Properties("height").Value = TargetFrame.Properties("height").Value
            End If
          End With 'PicTarget
          DeselectAll .VBControls
        End If
      End With
      ReadWriteScale Comp, False, OriginalScale, OriginalSclWidth, OriginalSclHeight, OriginalSclTop, OriginalSclLeft
      ''ver 1.0.95 update stops CF retesting all forms.
      LItem.ListSubItems(4).ForeColor = vbGreen
      LItem.SubItems(4) = "FIXED"
    End If
  Next I

Exit Sub

CutFail:
  BugTrapComment "InsertXPPic2Frame2"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Function IsDataDisplayTool(varName As Variant) As Boolean

  'v2.6.6 control tool comment to block Delete? suggestion

  IsDataDisplayTool = ArrayMember(varName, "DataGrid", "DataList", "DataCombo")

End Function

Public Function IsFileTool(varName As Variant) As Boolean

  'used to detect if a codeless tool is one of the file tools which
  ' may be set from IDE or used without code to navigate

  IsFileTool = ArrayMember(varName, "DriveListBox", "FileListBox", "DirListBox")

End Function

Public Function IsGraphic(varName As Variant) As Boolean

  IsGraphic = ArrayMember(varName, "Line", "Label", "Shape", "Frame", "PictureBox", "Image")

End Function

Private Function IsLinkedImageList(ByVal strTest As String, _
                                   ByVal Cno As Long, _
                                   strLinkedTo As String) As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDName = strTest Then
        If LenB(CntrlDesc(I).CDImageListLinkedTo) Then
          IsLinkedImageList = True
          strLinkedTo = CntrlDesc(Cno).CDImageListLinkedTo
          Exit For
        End If
      End If
    Next I
  End If

End Function

Public Function MatchCtrlToDescArrayIndex(vbc As VBControl, _
                                          Optional ByVal bForceNegOneIndex As Boolean = False) As Long

  Dim strPrj   As String
  Dim StrFrm   As String
  Dim strNam   As String
  Dim lngIndex As Long
  Dim I        As Long

  ExtractControlData vbc, strPrj, StrFrm, strNam, lngIndex
  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDProj = strPrj Then
        If CntrlDesc(I).CDForm = StrFrm Then
          If CntrlDesc(I).CDName = strNam Then
            If (CntrlDesc(I).CDIndex = lngIndex) Or (bForceNegOneIndex And (CntrlDesc(I).CDIndex = -1)) Then
              MatchCtrlToDescArrayIndex = I
              Exit For 'unction
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Private Function NeedsXPBugFix(ByVal strProj As String, _
                               ByVal strForm As String, _
                               ByVal StrFramename As String) As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDProj = strProj Then
        If CntrlDesc(I).CDForm = strForm Then
          If CntrlDesc(I).CDFullName = StrFramename Then
            If CntrlDesc(I).CDXPFrameBug Then
              NeedsXPBugFix = True
              Exit For
            End If
          End If
        End If
      End If
    Next I
  End If

End Function

Public Sub NoCodeCommentry(LItem As ListItem, _
                           ByVal I As Long, _
                           Optional Update As Boolean = False, _
                           Optional ByVal bIgnoreSingleton As Boolean = False)

  Dim strMsg      As String
  Dim strLinkedTo As String

  On Error GoTo BugTrap
  If CntrlDesc(I).CDXPFrameBug Then
    strMsg = "XP Style Frame Bug"
  End If
  With LItem
    If CntrlDesc(I).CDBadType > 0 Then
      If CntrlDesc(I).CDUsage = 0 Then
        strMsg = "(No Code)"
        If CntrlDesc(I).CDClass = "ImageList" Then
          If IsLinkedImageList(CntrlDesc(I).CDName, I, strLinkedTo) Then
            'display form and control only
            strMsg = "Linked(" & Replace$(Mid$(strLinkedTo, InStr(strLinkedTo, "|") + 1), "|", "-") & RBracket
          End If
        End If
      End If
      UpdateAddComment LItem, Update, strMsg & "Poor Name" & BadNameMsg(CntrlDesc(I).CDName) & IIf(Len(strMsg), ":" & strMsg, "")
      .ListSubItems(4).ForeColor = vbRed
     ElseIf Not bIgnoreSingleton And CntrlDesc(I).CDXPFrameBug <> 0 Then
      UpdateAddComment LItem, Update, strMsg
      .ListSubItems(4).ForeColor = vbBlue
     ElseIf CntrlDesc(I).CDUsage = 0 Then
      If CntrlDesc(I).CDClass = "Menu" Then
        If CntrlDesc(I).CDCaption = "-" Then
          'ver 2.1.0 Thanks to Dipankar Basu for suggestion
          'ver 2.1.3 oops fixed works properly now
          UpdateAddComment LItem, Update, "No Code(Menu Seperator)"
         Else
          UpdateAddComment LItem, Update, "No Code(Sub-Menu Header?)"
        End If
        .ListSubItems(4).ForeColor = vbGrey
       ElseIf IsGraphic(CntrlDesc(I).CDClass) Then
        If InStr(CntrlDesc(I).CDName, "picCFXPBugFix") Then
          UpdateAddComment LItem, Update, "XP Frame Bug Solution"
          .ListSubItems(4).ForeColor = vbGrey
         ElseIf CntrlDesc(I).CDIsContainer Then
          UpdateAddComment LItem, Update, "No Code(Container)"
          .ListSubItems(4).ForeColor = vbGrey
         Else
          UpdateAddComment LItem, Update, "No Code(Graphic)"
          .ListSubItems(4).ForeColor = vbGrey
        End If
       ElseIf IsFileTool(CntrlDesc(I).CDClass) Then
        UpdateAddComment LItem, Update, "No Code(File Control)"
        .ListSubItems(4).ForeColor = vbGrey
       ElseIf IsDataDisplayTool(CntrlDesc(I).CDClass) Then
        UpdateAddComment LItem, Update, "No Code(Data Display)"
        .ListSubItems(4).ForeColor = vbGrey
       Else
        If CntrlDesc(I).CDIsContainer Then
          UpdateAddComment LItem, Update, "No Code(Container)"
          .ListSubItems(4).ForeColor = vbGrey
         ElseIf CntrlDesc(I).CDClass = "ImageList" Then
          If IsLinkedImageList(CntrlDesc(I).CDName, I, strLinkedTo) Then
            UpdateAddComment LItem, Update, "Linked(" & Replace$(Mid$(strLinkedTo, InStr(strLinkedTo, "|") + 1), "|", "-") & RBracket
            .ListSubItems(4).ForeColor = vbGrey
           Else
            UpdateAddComment LItem, Update, "No Code(Delete?)"
            .ListSubItems(4).ForeColor = vbRed
          End If
         Else
          UpdateAddComment LItem, Update, "No Code(Delete?)"
          .ListSubItems(4).ForeColor = vbRed
        End If
      End If
     Else
      UpdateAddComment LItem, Update, ""
    End If
  End With

Exit Sub

BugTrap:
  BugTrapComment "NoCodeCommentry"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Private Sub PasteToOneControl(vbfrm As VBForm, _
                              ByVal CName As VBControl)

  With vbfrm
    DeselectAll .VBControls
    CName.InSelection = True
    .Paste
  End With

End Sub

Private Sub ReadWriteScale(Cmp As VBComponent, _
                           ByVal RW As Boolean, _
                           OrigScale As Long, _
                           OrigSclWidth As Single, _
                           OrigSclHeight As Single, _
                           OrigSclTop As Single, _
                           OrigSclLeft As Single)

  With Cmp
    If RW Then
      If .DesignerID <> "VB.MDIForm" Then
        OrigScale = .Properties("scalemode").Value
        OrigSclWidth = .Properties("scalewidth").Value
        OrigSclHeight = .Properties("scaleheight").Value
        OrigSclTop = .Properties("scaletop").Value
        OrigSclLeft = .Properties("scaleleft").Value
        .Properties("scalemode") = vbTwips
      End If
     Else
      If .DesignerID <> "VB.MDIForm" Then
        .Properties("scalemode").Value = OrigScale
        .Properties("scalewidth").Value = OrigSclWidth
        .Properties("scaleheight").Value = OrigSclHeight
        .Properties("scaletop").Value = OrigSclTop
        .Properties("scaleleft").Value = OrigSclLeft
      End If
    End If
  End With

End Sub

Public Sub ReWriteSelectedLineArrays(ByVal lngTmpIndex As Long, _
                                     Targetcontrol As VBControl)

  Dim I              As Long
  Dim LItem          As ListItem
  Dim SaveCurListPos As Long

  With frm_CodeFixer
    SaveCurListPos = .lsvAllControls.SelectedItem.Index
    For I = 1 To .lsvAllControls.ListItems.Count
      If CLng(.lsvAllControls.ListItems.Item(I).Tag) = lngTmpIndex Then
        SetCurrentLSVLine .lsvAllControls, I
        Exit For
      End If
    Next I
  End With 'frm_CodeFixer
  With CntrlDesc(lngTmpIndex)
    SetCurrentLSVLine frm_CodeFixer.lsvAllControls, I
    .CDName = Targetcontrol.Properties("name").Value
    .CDIndex = Targetcontrol.Properties("index").Value
    .CDFullName = .CDName & IIf(.CDIndex > -1, strInBrackets(.CDIndex), vbNullString)
    Set LItem = frm_CodeFixer.lsvAllControls.SelectedItem
    LItem.Text = .CDProj
    LItem.SubItems(2) = .CDFullName
    LItem.SubItems(2) = .CDFullName
    .CDBadType = GetBadNameType(lngTmpIndex, False)
    If .CDBadType = BNSingletonArray Then 'stops errant message while rewriting a whole array
      .CDBadType = BNNone
    End If
    NoCodeCommentry LItem, lngTmpIndex, True, True
  End With
  SetCurrentLSVLine frm_CodeFixer.lsvAllControls, SaveCurListPos

End Sub

Private Sub SelectOneControl(vbfrm As VBForm, _
                             ByVal CName As VBControl)

  With vbfrm
    DeselectAll .VBControls
    CName.InSelection = True
  End With

End Sub

Private Sub SetControlPositioning(vbf As VBForm, _
                                  ByVal ofset As Long, _
                                  ByVal voffset As Long)

  Dim I    As Long
  Dim Ctrl As Variant

  For Each Ctrl In vbf.SelectedVBControls
    I = I + 1
    On Error Resume Next
    With Ctrl
      .Properties.Item("left") = ctl(I).Left - ofset
      .Properties.Item("top") = ctl(I).Top - ofset * 2 - voffset
      .Properties.Item("tabindex") = ctl(I).TabIndex
      .ZOrder ctl(I).ZorderVal
    End With
    On Error GoTo 0
  Next Ctrl
  Erase ctl

End Sub

Public Sub SetCurrentLSVLine(Lsv As ListView, _
                             LPos As Long)

  If LPos Then
    If LPos > Lsv.ListItems.Count Then
      LPos = Lsv.ListItems.Count
    End If
   Else
    If Lsv.ListItems.Count Then
      LPos = 1
    End If
  End If
  If LPos Then
    With Lsv.ListItems(LPos)
      .Selected = True
      .EnsureVisible
    End With
  End If

End Sub

Public Sub SetNewName(ByVal strName As String)

  With frm_CodeFixer.txtCtrlNewName
    .Text = strName
    If Not AutoReName Then
      If strName <> CntrlDesc(LngCurrentControl).CDName Then
        .SelStart = Len(strName)
       Else
        .SelStart = 0
        .SelLength = Len(strName)
      End If
      SetFocus_Safe frm_CodeFixer.txtCtrlNewName
    End If
  End With

End Sub

Private Sub UpdateAddComment(LItm As ListItem, _
                             ByVal bUpdate As Boolean, _
                             ByVal strMsg As String)

  If bUpdate Then
    LItm.ListSubItems(4) = strMsg
   Else
    LItm.ListSubItems.Add , , strMsg
  End If

End Sub

Public Sub WarningLabel(Optional ByVal strMsg As String, _
                        Optional ByVal lBackColor As Long = vbRed)

  With frm_CodeFixer.lblOldName(10)
    If Len(strMsg) Then
      .BackColor = lBackColor
      .Caption = strMsg
     Else
      .Caption = vbNullString
      .BackColor = vbButtonFace
    End If
    .Refresh
  End With

End Sub

Public Sub WholeWordReplacer(cde As String, _
                             ByVal strF As String, _
                             ByVal StrRep As String, _
                             Optional bUpdated As Boolean)

  Dim Cmpr As VbCompareMethod
  Dim FPos As Long

  If Len(cde) Then
    If Len(strF) Then
      If InStr(cde, strF) Then
        Cmpr = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
        FPos = InStr(1, cde, strF, Cmpr)
        Do While InStr(FPos, cde, strF, Cmpr)
          If InCode(cde, FPos) Then
            If IsWholeWord(cde, strF, FPos) Then
              cde = Left$(cde, FPos - 1) & StrRep & Mid$(cde, FPos + Len(strF))
              bUpdated = True
            End If
          End If
          'FPos = InStr(FPos + Len(StrRep), cde, strF, Cmpr)
          'v2.6.5 dumb error missed matches near end of code; finally tracked down
          FPos = InStr(FPos + Len(strF), cde, strF, Cmpr)
          If FPos = 0 Then
            Exit Do
          End If
        Loop
      End If
    End If
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:12:58 AM) 9 + 951 = 960 Lines Thanks Ulli for inspiration and lots of code.

