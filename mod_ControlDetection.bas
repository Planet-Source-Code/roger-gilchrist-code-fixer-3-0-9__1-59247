Attribute VB_Name = "mod_ControlDetection"

'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Option Explicit
Public arrQCtrlPresence            As Variant
Public bControlsSet                As Boolean
Public bForceReload                As Boolean
'*****************
Public ImplementsArray             As Variant
Public Type EventDescriptor
  EName                            As String
  EProj                            As String
  EForm                            As String
  EScope                           As String
  EWhich                           As String
End Type
Public EventDesc()                 As EventDescriptor
Public bEventDescExists            As Boolean
Public bShowctrlPRoject            As Boolean
Public bShowctrlComponent          As Boolean
Public bShowctrlName               As Boolean
Public bShowctrlCaption            As Boolean
Public bShowctrlComment            As Boolean
Public Type ModuleDescriptor
  MDName                           As String
  MDType                           As String
  MDTypeNum                        As Long
  MDCaption                        As String
  MDUsage                          As Long
  MDProj                           As String
  MDBadType                        As Long
  MDAttributes                     As Long
  MDIsClass                        As Boolean
  MDFilename                       As String
  MDFullPath                       As String
  MDReadOnly                       As Boolean
  MDHidden                         As Boolean
  MDDontTouch                      As Boolean
  MDWholeOptCompile                As Boolean
  MDisControlHolder                As Boolean
End Type
Public ModDesc()                   As ModuleDescriptor
Public bModDescExists              As Boolean
Public Type ProjectDescriptor
  PDName                           As String
  PDClass                          As String
  PDAttributes                     As Long
  PDScope                          As String
  PDFilename                       As String
  PDFullPath                       As String
  PDReadOnly                       As Boolean
  PMDHidden                        As Boolean
  PDBadType                        As Long
End Type
Public ProjDesc()                  As ProjectDescriptor
Public bProjDescExists             As Boolean
Public Enum LVSCW_Styles
  LVSCW_AUTOSIZE = -1
  LVSCW_AUTOSIZE_USEHEADER = -2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private LVSCW_AUTOSIZE, LVSCW_AUTOSIZE_USEHEADER
#End If
'--- ListView Set Column Width Messages ---'
'Thanks to  Karl E. Peterson  http://www.mvps.org/vb
'who published these API ways of sizing the ListView
'Modification Added ability to Suppress specific columns
Public Const strUnsavedModule      As String = "[Module is not Saved]"

Public Sub AllControlsColumnWidths()

  Dim Lsv       As ListView
  Dim lngComWid As Long
  Dim I         As Long

  Set Lsv = frm_CodeFixer.lsvAllControls
  SetColumn Lsv, "proj", e_CDProj, bShowctrlPRoject
  SetColumn Lsv, "component", e_CDform, bShowctrlComponent
  SetColumn Lsv, "control", e_CDName, True
  SetColumn Lsv, "captions", e_CDCaption, bShowctrlCaption
  SetColumn Lsv, "use", e_CDUsage, bShowctrlComment
  With Lsv
    For I = 1 To 4 'set to 5 to see the hidden index
      lngComWid = lngComWid + .ColumnHeaders(I).Width + 15
    Next I
    If .ListItems.Count > ListViewVisibleItems(Lsv) Then
      lngComWid = lngComWid + 250
    End If
    'deal with everything being too wide for the comments coloumn to display
    Do
      .ColumnHeaders("proj").Width = .ColumnHeaders("proj").Width * 0.9
      .ColumnHeaders("component").Width = .ColumnHeaders("component").Width * 0.9
      .ColumnHeaders("captions").Width = .ColumnHeaders("captions").Width * 0.9
      lngComWid = 0
      For I = 1 To 4 'set to 5 to see the hidden index
        lngComWid = lngComWid + .ColumnHeaders(I).Width + 15
      Next I
      If .ListItems.Count > ListViewVisibleItems(Lsv) Then
        lngComWid = lngComWid + 250
      End If
    Loop Until .Width - lngComWid > .Width / 5
    .ColumnHeaders("use").Width = .Width - lngComWid
  End With

End Sub

Private Sub AllControlsDisplay(Lsv As ListView, _
                               ByVal LPos As Long)

  Dim LItem           As ListItem
  Dim I               As Long
  Dim lngSepNameCount As Long
  Dim lngArrayCount   As Long
  Dim strPrevName     As String
  Dim BadCount        As Long
  Dim XPCount         As Long

  On Error GoTo BugTrap
  Lsv.ListItems.Clear
  SendMessage Lsv.hWnd, WM_SETREDRAW, False, 0
  frm_CodeFixer.frapage(TPControls).Caption = "Loading List..."
  frm_CodeFixer.frapage(TPControls).Refresh
  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDXPFrameBug Then
        XPCount = XPCount + 1
      End If
      If CntrlDesc(I).CDBadType <> 0 Then
        BadCount = BadCount + 1
      End If
      If CntrlDesc(I).CDUsage <> 2 Then
        Set LItem = Lsv.ListItems.Add(, , CntrlDesc(I).CDProj)
        With LItem
          .ListSubItems.Add , , CntrlDesc(I).CDForm
          .ListSubItems.Add , , CntrlDesc(I).CDFullName
          If Len(CntrlDesc(I).CDCaption) Then
            .ListSubItems.Add , , CntrlDesc(I).CDCaption
           Else
            .ListSubItems.Add , , ""
          End If
          NoCodeCommentry LItem, I
          .Tag = I
          If CntrlDesc(I).CDName <> strPrevName Then
            strPrevName = CntrlDesc(I).CDName
            lngSepNameCount = lngSepNameCount + 1
            If CntrlDesc(I).CDIndex <> -1 Then
              lngArrayCount = lngArrayCount + 1
            End If
          End If
        End With 'LItem
      End If
    Next I
    AllControlsColumnWidths
    Lsv.SortKey = 0
    DoEvents
    Lsv.SortKey = 1
    DoEvents
    Lsv.SortKey = 2
    DoEvents
    Lsv.SortKey = 1
    SetCurrentLSVLine Lsv, LPos
    SendMessage Lsv.hWnd, WM_SETREDRAW, True, 0
    With frm_CodeFixer
      .mnupopControlLev3(2).Checked = True
      .frapage(TPControls).Caption = "Controls [" & Lsv.ListItems.Count & "]" & IIf(lngSepNameCount > 0 And lngSepNameCount <> Lsv.ListItems.Count, " [" & lngSepNameCount & " Names / " & lngArrayCount & " Control Arrays]", "[No Control Arrays]") & IIf(BadCount, "[" & BadCount & " Poorly Named]", vbNullString) & IIf(XPCount, "[" & XPCount & " XP-Frame Bug]", vbNullString)
      If LPos Then
        DoAllSelect
      End If
    End With
  End If

Exit Sub

BugTrap:
  BugTrapComment "AllControlsDisplay"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Sub

Public Sub AllModulesDisplay(Lsv As ListView, _
                             ByVal LPos As Long)

  Dim LItem As ListItem
  Dim I     As Long

  HideColumnsReturnPosition Lsv, LPos
  SendMessage Lsv.hWnd, WM_SETREDRAW, False, 0
  'lsv
  For I = LBound(ModDesc) To UBound(ModDesc)
    If Len(ModDesc(I).MDName) Then
      Set LItem = Lsv.ListItems.Add(, , ModDesc(I).MDProj)
      With LItem
        .ListSubItems.Add , , ModDesc(I).MDName
        .ListSubItems.Add , , ModDesc(I).MDFilename
        If ModDesc(I).MDFilename = strUnsavedModule Then
          .ListSubItems(2).ForeColor = vbRed
        End If
        .ListSubItems.Add , , I 'hiddenindex
        If ModDesc(I).MDBadType > 0 Then
          .ListSubItems.Add , , "Poor Name" & BadNameMsg2(ModDesc(I).MDBadType)
          .ListSubItems(4).ForeColor = vbRed
         Else
          .ListSubItems.Add , , ""
        End If
      End With 'LItem
    End If
NotBad:
  Next I
  DoListViewWidth Lsv, Array(1, 4)
  If UBound(ModDesc) > -1 Then
    If LPos = 0 Then
      LPos = 1
    End If
  End If
  SendMessage Lsv.hWnd, WM_SETREDRAW, True, 0
  SetCurrentLSVLine Lsv, LPos
  SetModuleListitem

End Sub

Public Function AnyBadNames() As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDBadType <> 0 Then
        AnyBadNames = True
        Exit For
      End If
    Next I
  End If

End Function

Private Function AnyNamesWithoutPrefix() As Boolean

  Dim I       As Long
  Dim strJunk As String

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDUsage < 2 Then
        If HasNoHungarianPrefix(I, strJunk) Then
          AnyNamesWithoutPrefix = True
          Exit For
        End If
      End If
    Next I
  End If

End Function

Private Function AnySingletonCArrays() As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If HasSingletonCArray(CntrlDesc(I).CDBadType) Then
        AnySingletonCArrays = True
        Exit For
      End If
    Next I
  End If

End Function

Private Function AnyXPBugs() As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDXPFrameBug Then
        AnyXPBugs = True
        Exit For
      End If
    Next I
  End If

End Function

Public Function BadInitialization(strPrj As String, _
                                  strCmp As String) As String

  Dim CompMod    As CodeModule
  Dim arrMembers As Variant
  Dim I          As Long
  Dim J          As Long
  Dim ArrProc    As Variant
  Dim arrErr(3)  As Boolean   ' only report each error once
  Dim strErrorIn As String
  Dim strActMsg  As String
  Dim strTest    As String

  strActMsg = WARNING_MSG & "Illegal call to $$ Object. Causes Error(398) 'Client Site not available'" & vbNewLine & _
   SUGGESTION_MSG & "Move to 'UserControl_ReadProperties'." & vbNewLine & _
   WARNING_MSG & "You will have to close VB and restart to fully re-initialise the controls."
  Set CompMod = GetComponent(strPrj, strCmp).CodeModule
  If CompMod.Members.Count Then
    arrMembers = GetMembersArray(CompMod)
    For I = 1 To UBound(arrMembers)
      If GetProcNameStr(arrMembers(I)) = "UserControl_Initialize" Then
        strErrorIn = GetProcNameStr(arrMembers(I))
        ArrProc = Split(arrMembers(I), vbNewLine)
        For J = LBound(ArrProc) To UBound(ArrProc)
          If Not JustACommentOrBlank(ArrProc(J)) Then
            strTest = strCodeOnly(ArrProc(J))
            If Not arrErr(0) Then
              If InStr(strTest, ".Ambient") Then
                BadInitialization = BadInitialization & "Illegal call to Ambient Object in" & strInSQuotes(strErrorIn, True) & "." & vbNewLine
                InsertInCode CStr(ArrProc(J)), strErrorIn, strCmp, strPrj, Replace$(strActMsg, "$$", "Ambient")
                arrErr(0) = True
              End If
            End If
            If Not arrErr(1) Then
              If strErrorIn <> "UserControl_InitProperties" Then
                If InStr(strTest, ".Extender") Then
                  BadInitialization = BadInitialization & "Illegal call to Extender Object in" & strInSQuotes(strErrorIn, True) & "." & vbNewLine
                  InsertInCode CStr(ArrProc(J)), strErrorIn, strCmp, strPrj, Replace$(strActMsg, "$$", "Extender")
                  arrErr(1) = True
                End If
              End If
            End If
            If Not arrErr(2) Then
              If InStr(strTest, ".ParentControls") Then
                BadInitialization = BadInitialization & "Illegal call to ParentControls Object in" & strInSQuotes(strErrorIn, True) & "." & vbNewLine
                InsertInCode CStr(ArrProc(J)), strErrorIn, strCmp, strPrj, Replace$(strActMsg, "$$", "ParentControls")
                arrErr(2) = True
              End If
            End If
          End If
        Next J
      End If
    Next I
  End If
  If LenB(BadInitialization) Then
    BadInitialization = BadInitialization & vbNewLine & _
     "These will produce Error (398) 'Client Site not available'"
    DisplayCodePane CompMod.Parent, True
    FindInCode strErrorIn, strCmp, strPrj
   Else
    BadInitialization = "There may be an illegal call in 'UserControl_InitProperties' or 'UserControl_Inititialize'." & vbNewLine & _
                        "These will produce Error (398) 'Client Site not available'"
    DisplayCodePane CompMod.Parent, True
  End If

End Function

Public Function BadModuleName(ByVal strName As String, _
                              ByVal strClass As String) As Long

  If isRefLibVBCommands(strName, False) Then
    BadModuleName = BNCommand
   ElseIf InQSortArray(ArrQVBStructureWords, strName) Then
    BadModuleName = BNStructural
   ElseIf InQSortArray(ArrQVBReservedWords, strName) Then
    BadModuleName = BNReserve
   ElseIf IsControlProperty(strName) Then
    BadModuleName = BNKnown
   ElseIf strName = strClass Then
    ' thanks Georg Veichtlbauer I moved this here so more serious porblems got detected first
    BadModuleName = BNClass
   ElseIf IsDeclaredVariable(strName) Then
    If Not isEvent(strName) Then
      BadModuleName = BNVariable
    End If
   ElseIf IsDeclareName(strName) Then
    If Not isEvent(strName) Then
      BadModuleName = BNVariable
    End If
   ElseIf IsProcedureName(strName) Then
    BadModuleName = BNProc
   ElseIf Len(strName) = 1 Then
    BadModuleName = BNSingle
   ElseIf UsingDefVBName(strName, strClass) Then
    BadModuleName = BNDefault
   Else
    BadModuleName = BNNone
  End If

End Function

Private Sub BadUserControlMsg(vbc As VBControl, _
                              Comp As VBComponent, _
                              strPrj As String, _
                              strCmp As String)

  Dim strErrMsg       As String
  Dim strBadInit      As String
  Const StrCFAborting As String = vbNewLine & _
   vbNewLine & _
   "Code Fixer will now abort processing." & vbNewLine & _
   "Please wait."

  'v 2.1.4
  'thanks to Fred.cpp  whose 'isExplorerBar 1.62' contained the problems that let me develop this
  If Len(vbc.ClassName) Then
    strBadInit = BadInitialization(strPrj, strCmp)
    If LenB(strBadInit) Then
      strErrMsg = "User Control" & strInSQuotes(strCmp, True) & "has following possible problem(s):" & vbNewLine & _
                  strBadInit & vbNewLine & _
                  "Move the lines into 'User_Control_ReadProperties'" & vbNewLine & _
                  "After repairing these problems, please close the project and reopen it to allow VB to fully reinitialise the control." & StrCFAborting
     Else
      If UserControlUsesResumeNext(Comp) Then
        strErrMsg = "User Control" & strInSQuotes(strCmp, True) & "contains at least one 'On Error Resume Next' which allows it to fall past a line containing an illegal statement." & vbNewLine & _
                    "There are 2 possible fixes:" & vbNewLine & _
                    "A. (Quick but dirty)" & vbNewLine & _
                    "Find all the error traps and make sure that there is a following 'On Error Goto 0' or 'Error.Clear' to reset the error trap." & vbNewLine & _
                    "B. (Hard but better)" & vbNewLine & _
                    "  i. Search for error traps & set Watchpoints (F9)" & vbNewLine
        strErrMsg = strErrMsg & " ii. Set Error Trapping to 'Break on All Errors'. (Tools/Options.../General)" & vbNewLine & _
         "iii. Run the code. (Ctrl+F5 for full compile mode)" & vbNewLine & _
         " iv. When you reach a Watchpoint step through the code. (F8)." & vbNewLine & _
         "  v. Try to correct the error causing statment. If you succeed you may be able to remove the Error trap, otherwise apply Fix A." & StrCFAborting
       Else
        strErrMsg = "A control of Class:" & strInSQuotes(strCmp, True) & "on" & strInSQuotes(Comp.Name, True) & "can not be initialised." & vbNewLine & _
                    "Possible causes" & vbNewLine & _
                    "1. A UserControl Property has an invalid value." & vbNewLine & _
                    "FIX A: Save the project. Close And Re-open VB. This will force the form to update its files." & vbNewLine & _
                    "FIX B: Open the Form, if VB generates a log file read it. Save the code. VB will warn you about overwriting a file, do it." & vbNewLine & _
                    "2. A UserControl is not initializing properly." & vbNewLine & _
                    "FIX: Open the UserControl code by itself in VB (*.ctl file or *.vbp file if it has one and run Code Fixer." & StrCFAborting
      End If
    End If
   Else
    strErrMsg = "A Control on" & strInSQuotes(Comp.Name, True) & "does not have a name." & vbNewLine & _
                "Possible causes" & vbNewLine & _
                "1. UserControl is not initializing properly." & vbNewLine & _
                "FIX: See 'UserControl Problems' in the Help File for details." & vbNewLine & _
                "2. There was a loading error (missing Dll, OCX or UserControl)." & vbNewLine & _
                "FIX: Save the project. Answer Yes when warned that you are overwriting Forms." & StrCFAborting
  End If
  If LenB(strErrMsg) Then
    mObjDoc.Safe_MsgBox strErrMsg, vbCritical
  End If

End Sub

Public Function CleanCaption(ByVal StrCap As String) As String

  Dim arrCaption As Variant
  Dim I          As Long

  arrCaption = Split(StrCap)
  For I = LBound(arrCaption) To UBound(arrCaption)
    arrCaption(I) = Replace$(arrCaption(I), "&&", vbNullString)
    arrCaption(I) = Replace$(arrCaption(I), "&", vbNullString)
    arrCaption(I) = StripPunctuation(arrCaption(I), "_")
    arrCaption(I) = Ucase1st(arrCaption(I)) 'Make Each Character Word Proper Case
  Next I
  For I = LBound(arrCaption) To UBound(arrCaption)
    If Len(arrCaption(I)) Then
      CleanCaption = CleanCaption & arrCaption(I)
      If Len(CleanCaption) > 25 Then
        If I > 1 Then
          CleanCaption = Left$(CleanCaption, Len(CleanCaption) - Len(arrCaption(I)))
         Else
          CleanCaption = Left$(CleanCaption, 25)
        End If
        Exit For
      End If
    End If
  Next I

End Function

Public Function Control_Engine() As Boolean

  If Not bAborting Then
    WorkingMessage "Build Implements Array", 1, 3
    Generate_ImplementsArray
    ReDim EventDesc(0) As EventDescriptor
    WorkingMessage "Build Events Array", 2, 3
    Generate_EventArray
    'Generate_WithEventArray
    WorkingMessage "Build Controls Array", 3, 3
    Control_Engine = Generate_ControlArray
  End If

End Function

Public Sub ControlAutoEnabled(ByVal bOnOff As Boolean)

  With frm_CodeFixer
    .cmdXPStyle.Enabled = bOnOff And WeAreRunningUnderWinXP
    .cmdXPStyle.Visible = WeAreRunningUnderWinXP
    .cmdXPStyle.Caption = XPStyleCaption
    .cmdAutoLabel(0).Enabled = bOnOff And AnyBadNames
    .cmdAutoLabel(1).Enabled = bOnOff And AnyNamesWithoutPrefix
    .cmdAutoLabel(2).Enabled = bOnOff
    .chkDelOldCode.Enabled = bOnOff
    .cmdFindInCode.Enabled = bOnOff
    .cmdAutoLabel(3).Enabled = bOnOff And AnyXPBugs
    .cmdAutoLabel(4).Enabled = bOnOff And AnySingletonCArrays
    .lblDeletableExist.BackColor = IIf(IsAnyControlDeletable, vbRed, vbButtonFace)
    .chkDelOldCode.Enabled = bOnOff
    DoEvents
  End With

End Sub

Public Function ControlHasProperty(vbc As VBControl, _
                                   ByVal StrProp As String) As Boolean

  'Return True if the property exists
  'but remember it may not be used

  On Error GoTo notfound
  If Len(vbc.Properties(StrProp).Name) Then
    ControlHasProperty = True
  End If
notfound:
  On Error GoTo 0

End Function

''
Public Function ControlHasPropertyValid(vbc As VBControl, _
                                        ByVal StrProp As String) As Boolean

  Dim varTest As Variant 'v2.9.7 Has to be 2 timer

  'Return True if the property exists
  'but remember it may not be used
  On Error GoTo notfound
  With vbc
    If Len(.Properties(StrProp).Name) Then
      varTest = .Properties(StrProp).Value
      ControlHasPropertyValid = True
    End If
  End With 'vbc
notfound:
  On Error GoTo 0

End Function

''
Public Function ControlName(VBCtrl As VBControl) As String

  'extracts name and index( if any) for easy control identification

  With VBCtrl
    ControlName = .Properties("name").Value
    If .Properties("index").Value > -1 Then
      ControlName = ControlName & strInBrackets(.Properties("index").Value)
    End If
  End With 'VBctrl

End Function

Public Sub DoListViewWidth(lv As ListView, _
                           Optional SupressArray As Variant)

  Dim ColumnIndex As Long

  'based on LVSetAllColWidths
  '  Copyright ©1997, Karl E. Peterson
  '  http://www.mvps.org/vb
  '--- loop through all of the columns in the listview and size each
  'Modification1 Added ability to Suppress specific columns
  'Modification2 HardCoded style
  'lv.Visible = False
  If IsMissing(SupressArray) Then
    SupressArray = Array(-1)
  End If
  With lv
    For ColumnIndex = 1 To .ColumnHeaders.Count
      If IsInArray(ColumnIndex, SupressArray) Then
        .ColumnHeaders(ColumnIndex).Width = 0
       Else
        LVSetColWidth lv, ColumnIndex, LVSCW_AUTOSIZE_USEHEADER
      End If
    Next ColumnIndex
  End With
  ' lv.Visible = True

End Sub

Private Function DOPartialUsageX(ByVal strPFind As String, _
                                 ByVal blnWhole As Boolean) As Boolean

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim X          As Long
  Dim L_CodeLine As String
  Dim varFind    As Variant

  On Error Resume Next
  If Len(strPFind) Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If LenB(Comp.Name) Then
          X = 1
          Do While Comp.CodeModule.Find(strPFind, X, 1, -1, -1, blnWhole, blnWhole, Not blnWhole)
            L_CodeLine = Comp.CodeModule.Lines(X, 1)
            If InCode(L_CodeLine, InStr(L_CodeLine, varFind)) Then
              DOPartialUsageX = True
              'v2.3.8
              GoTo SafeExit 'Exit Function 'safe because no with
            End If
            X = X + 1
            If X > Comp.CodeModule.CountOfLines Then
              Exit Do
            End If
          Loop
        End If
      Next Comp
    Next Proj
  End If
SafeExit:
  Set Comp = Nothing
  Set Proj = Nothing
  On Error GoTo 0

End Function

Private Sub FakeControls(strNme As String, _
                         strClas As String)

  'these fakes stop the deleted control code detector from hitting
  'VB default names and objects

  FakeControlsSingle "Form", "Form", strNme, strClas
  FakeControlsSingle "Class", SngSpace, strNme, strClas
  FakeControlsSingle "UserControl", SngSpace, strNme, strClas
  FakeControlsSingle "MDIForm", "MDIForm", strNme, strClas
  FakeControlsSingle "AddinInstance", SngSpace, strNme, strClas
  FakeControlsSingle "UserDocument", SngSpace, strNme, strClas
  FakeControlsSingle "PropertyPage", "PropertyPage", strNme, strClas
  FakeControlsSingle "Printer", "Printer", strNme, strClas
  FakeControlsSingle "Scripting", "FileSystem", strNme, strClas
  FakeControlsFromArray ImplementsArray, strNme, strClas
  FakeControlsFromDescEvents strNme, strClas

End Sub

Private Sub FakeControlsFromArray(arrC As Variant, _
                                  strNme As String, _
                                  strClas As String)

  Dim I As Long

  'these fakes stop the deleted control code detector from hitting
  'VB default names and objects
  'this routine extracts possible fakes from arrays
  If Not IsEmpty(arrC) Then
    If Not IsArrayEmpty(arrC) Then
      For I = LBound(arrC) To UBound(arrC)
        strNme = AccumulatorString(strNme, arrC(I), ",", False)
        strClas = AccumulatorString(strClas, SngSpace, ",", False)
      Next I
    End If
  End If

End Sub

Private Sub FakeControlsFromDescEvents(strNme As String, _
                                       strClas As String)

  Dim I As Long

  If bEventDescExists Then
    For I = LBound(EventDesc) To UBound(EventDesc)
      strNme = AccumulatorString(strNme, EventDesc(I).EName, ",", False)
      strClas = AccumulatorString(strClas, SngSpace, ",", False)
    Next I
  End If

End Sub

Private Sub FakeControlsSingle(ByVal strCName As String, _
                               ByVal strCClass As String, _
                               strNme As String, _
                               strClas As String)

  'these fakes stop the deleted control code detector from hitting
  'VB default names and objects
  'this routine adds fakes to the strings used to generate the controls array

  strNme = AccumulatorString(strNme, strCName, ",", False)
  strClas = AccumulatorString(strClas, strCClass, ",", False)

End Sub

Public Function FileBaseName(ByVal filespec As String) As String

  If LenB(filespec) Then
    FileBaseName = FSO.GetBaseName(filespec)
  End If

End Function

Public Function FileExtention(ByVal filespec As String) As String

  If LenB(filespec) Then
    FileExtention = FSO.GetExtensionName(filespec)
  End If

End Function

Public Function FileNameOnly(ByVal filespec As String) As String

  If LenB(filespec) Then
    FileNameOnly = FSO.GetFileName(filespec)
  End If

End Function

Public Sub FillAllControlsList(ByVal ListPos As Long, _
                               ByVal bRefreshBad As Boolean, _
                               ByVal UpdateRate As Long, _
                               Optional ByVal bRefreshXPBug As Boolean = False)

  Dim I    As Long
  Dim vbf  As VBForm
  Dim vbc  As VBControl
  Dim Comp As VBComponent

  frm_CodeFixer.lsvAllControls.ListItems.Clear
  If bRefreshXPBug Then
    If bCtrlDescExists Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If CntrlDesc(I).CDXPFrameBug Then
          Set Comp = GetComponent(CntrlDesc(I).CDProj, CntrlDesc(I).CDForm)
          ActivateDesigner Comp, vbf, False
          If vbf Is Nothing Then
            ActivateDesigner Comp, vbf, True
          End If
          For Each vbc In vbf.VBControls
            If vbc.Properties("name").Value = CntrlDesc(I).CDName Then
              If vbc.Properties("index").Value = CntrlDesc(I).CDIndex Then
                If Not NeedsXPFrameFix(vbc) Then
                  CntrlDesc(I).CDXPFrameBug = False
                  Exit For
                End If
              End If
            End If
          Next vbc
        End If
      Next I
    End If
  End If
  If bRefreshBad Then
    If bCtrlDescExists Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If CInt(I / UBound(CntrlDesc) * 100) Mod UpdateRate = 0 Then
          DoEvents
          frm_CodeFixer.frapage(TPControls).Caption = "Checking Poorly named Controls..." & Format$(I / UBound(CntrlDesc) * 100, "##0") & "%"
          frm_CodeFixer.frapage(TPControls).Refresh
        End If
        With CntrlDesc(I)
          If .CDUsage <> 2 Then
            .CDBadType = GetBadNameType(I)
          End If
        End With
      Next I
    End If
  End If
  If FrameActive = TPControls Then
    AllControlsDisplay frm_CodeFixer.lsvAllControls, ListPos
  End If

End Sub

Public Function Generate_ControlArray() As Boolean

  
  Dim OnReportTab           As Boolean
  Dim ListPos               As Long
  Dim CurControlOnForm      As Long
  Dim CDC                   As Long
  Dim TotalControls         As Long
  Dim RealControlCount      As Long
  Dim Oldsize               As Long
  Dim CurCompCount          As Long
  Dim strName               As String
  Dim strClass              As String
  Dim ExistingControlsArray As Variant
  Dim ControlClassArray     As Variant
  Dim vbc                   As VBControl
  Dim vbf                   As VBForm
  Dim Comp                  As VBComponent
  Dim Proj                  As VBProject
  Dim MyhourGlass           As cls_HourGlass
  Dim I                     As Long
  Dim UpdateRate            As Long
  Dim bHasImageLists        As Boolean

  Set MyhourGlass = New cls_HourGlass
  'ver1.1.31 added ControlOwnerArray for detection of incomplete references to controls
  'Name and Index are separate because Code Fixer just uses name to identify controls
  'Index is only used to allow Code Fixer to know whether to search past brackets for properties
  On Error GoTo BugHit
  'forms with Missing References can cause a crash when this attempts to Activate it
  'so this needs Error Trap
  ' if the controldata already matches the controls don't redo search
  With frm_CodeFixer
    OnReportTab = mObjDoc.GraphVisible '.tbsMain.SelectedItem.Key = "report"
    If Not OnReportTab Then
      HideColumnsReturnPosition .lsvAllControls, ListPos
      .frapage(TPControls).Caption = "Refreshing Controls..."
      WarningLabel "Up-dating, please wait...", vbRed
    End If
  End With
  GenerateReferencesEnumArray
  TotalControls = TotalControlCount
  Select Case TotalControls
   Case Is > 200
    UpdateRate = 1
   Case Is > 100
    UpdateRate = 2
   Case Is > 50
    UpdateRate = 5
   Case Else
    UpdateRate = 10
  End Select
  If TotalControls Then ' comment out this line and its counterpart to test error detector
    ReDim CntrlDesc(TotalControls - 1) As ControlDescriptor
    ' -1 because 0-Based
    CDC = -1
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If SafeCompToProcess(Comp, CurCompCount) Then
          If IsComponent_ControlHolder(Comp) Then
            If mObjDoc.GraphVisible Then
              ModuleMessage Comp, CurCompCount
            End If
            ActivateDesigner Comp, vbf, False
            If vbf Is Nothing Then
              ActivateDesigner Comp, vbf, True
            End If
            CurControlOnForm = 0
            bHasImageLists = False
            For Each vbc In vbf.VBControls
              If vbc.ClassName = "ImageList" Then
                bHasImageLists = True
                Exit For
              End If
            Next vbc
            For Each vbc In vbf.VBControls
              CDC = CDC + 1
              RealControlCount = RealControlCount + 1
              If OnReportTab Then
                CurControlOnForm = CurControlOnForm + 1
               Else
                If CInt(CDC / TotalControls * 100) Mod UpdateRate = 0 Then
                  frm_CodeFixer.frapage(TPControls).Caption = "Refreshing Controls..." & Format$(CDC / TotalControls * 100, "##0") & "%"
                  'frm_CodeFixer.frapage(TPControls).Refresh
                End If
              End If
              'DoEvents
              SetControlData CDC, vbc, Comp, vbf, False, bHasImageLists
              If bAborting Then
                Exit For
              End If
              MemberMessage CntrlDesc(CDC).CDFullName, CurControlOnForm, vbf.VBControls.Count
            Next vbc
          End If
        End If
        If bAborting Then
          Exit For
        End If
      Next Comp
      If bAborting Then
        Exit For
      End If
    Next Proj
    '    If bAborting Then
    '      Exit Function
    '    End If
    'generate the fakes for Code Fixer
    If Not OnReportTab Then
      frm_CodeFixer.frapage(TPControls).Caption = "Tidying up..."
      frm_CodeFixer.frapage(TPControls).Refresh
    End If
  End If
  For I = 0 To CDC
    arrQCtrlPresence = QuickSortAppend(arrQCtrlPresence, CntrlDesc(I).CDName)
  Next I
  FakeControls strName, strClass
  If LenB(strName) Then
    If TotalControls Then
      Oldsize = UBound(CntrlDesc)
    End If
    ExistingControlsArray = Split(strName, ",")
    ReDim Preserve CntrlDesc(Oldsize + UBound(ExistingControlsArray) + 1) As ControlDescriptor
    ControlClassArray = Split(strClass, ",")
    For CDC = LBound(ExistingControlsArray) To UBound(ExistingControlsArray)
      With CntrlDesc(Oldsize + 1 + CDC)
        .CDName = ExistingControlsArray(CDC)
        .CDClass = ControlClassArray(CDC)
        .CDDefProp = vbNullString
        .CDIndex = -1
        .CDForm = vbNullString
        .CDProj = vbNullString
        .CDCaption = vbNullString
        .CDXPFrameBug = False
        .CDIsContainer = False
        .CDUsage = 2
      End With
    Next CDC
    strClass = vbNullString
    For CDC = LBound(CntrlDesc) To UBound(CntrlDesc)
      ''v2.8.4 Thanks Richard Brisley this was the hidden secondary bug
      ' some of the fake controls (don't ask;)) that CF uses have a CDClass name of " "
      ' and originally the next code line only tested Len not Trim (also change it to LEnB for (micro) more speed)
      ' caused handled error later in the code
      If LenB(Trim$(CntrlDesc(CDC).CDClass)) Then
        strClass = AccumulatorString(strClass, CntrlDesc(CDC).CDClass, ",")
      End If
    Next CDC
    ArrQActiveControlClass = QuickSortArray(StripDuplicateArray(Split(strClass, ",")))
  End If
  bCtrlDescExists = UBound(CntrlDesc) > -1
  SetCDImageListLinkedTo
  For CDC = LBound(CntrlDesc) To UBound(CntrlDesc)
    If CntrlDesc(CDC).CDUsage <> 2 Then
      arrQCtrlPresence = QuickSortAppend(arrQCtrlPresence, CntrlDesc(CDC).CDName)
    End If
  Next CDC
  If Not OnReportTab Then
    If RealControlCount Then
      FillAllControlsList ListPos, True, UpdateRate
     Else
      mObjDoc.Safe_MsgBox "This Project has no Controls.", vbInformation
      frm_CodeFixer.frapage(TPControls).Caption = "No Controls in Project"
      frm_CodeFixer.frapage(TPControls).Refresh
      WarningLabel "No Controls in Project"
    End If
  End If
  If Xcheck(XStayOnTop) Then
    SetTopMost frm_CodeFixer, True
  End If
  Generate_ControlArray = True
  On Error GoTo 0

Exit Function

BugHit:
  BugTrapComment "Generate_ControlArray"
  If RunningInIDE Then
    Resume
   Else
    Resume Next
  End If

End Function

Private Sub Generate_EventArray()

  Dim CurCompCount As Long
  Dim X            As Long
  Dim TmpIndex     As Long
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim L_CodeLine   As String
  Dim arrLine      As Variant
  Dim MaxFactor    As Long
  Dim CompMod      As CodeModule
  Dim lngMaxComp   As Long
  Dim arrFind(1)   As String
  Dim I            As Long
  Dim GuardLine    As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'ver 1.1.79
  'events are recognized as controls
  arrFind(0) = "Event"
  arrFind(1) = "WithEvents"
  For Each Proj In VBInstance.VBProjects
    lngMaxComp = GetComponentCount
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        ModuleMessage Comp, CurCompCount
        Set CompMod = Comp.CodeModule
        MaxFactor = CompMod.CountOfDeclarationLines
        MemberMessage "", CurCompCount, lngMaxComp
        If MaxFactor Then
          For I = 0 To 1
            X = 1
            GuardLine = 0
            Do While CompMod.Find(arrFind(I), X, 1, CompMod.CountOfDeclarationLines + 1, -1, True, True, False)
              'v2.3.2 makes sure that the search exits
              If GuardLine > 0 Then
                If GuardLine > X Then
                  Exit Do
                End If
              End If
              If X <= CompMod.CountOfDeclarationLines Then
                L_CodeLine = Trim$(CompMod.Lines(X, 1))
                If ExtractCode(L_CodeLine) Then
                  arrLine = Split(ExpandForDetection(L_CodeLine))
                  If InStrWholeWordRX(L_CodeLine, "Event") Then
                    TmpIndex = PosInStringWholeWord(L_CodeLine, "Event") - 1
                   Else
                    TmpIndex = PosInStringWholeWord(L_CodeLine, "WithEvents") - 1
                  End If
                  If TmpIndex > -1 Then
                    Select Case arrLine(TmpIndex)
                     Case "Event", "WithEvents"
                      If InCode(L_CodeLine, InStr(L_CodeLine, arrLine(TmpIndex))) Then
                        With EventDesc(UBound(EventDesc))
                          .EName = arrLine(TmpIndex + 1)
                          .EWhich = arrLine(TmpIndex)
                          .EForm = Comp.Name
                          .EProj = Proj.Name
                          If TmpIndex > 0 Then
                            .EScope = arrLine(0)
                           Else
                            .EScope = "Public"
                          End If
                        End With 'EventDesc(UBound(EventDesc))
                        ReDim Preserve EventDesc(UBound(EventDesc) + 1) As EventDescriptor
                        bEventDescExists = True
                      End If
                    End Select
                  End If
                End If
                X = X + 1
                GuardLine = X
                'v 2.2.4 fixed it was missing the event if it was last line of declaration
                If X > MaxFactor Then
                  Exit Do
                End If
                'v 2.3.2 (Very unlikely) if all the Detected events are commented out and there are no Procedures
                If X >= Comp.CodeModule.CountOfLines Then
                  Exit Do
                End If
              End If
            Loop
          Next I
        End If
      End If
SkipComp:
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  'v3.0.7
  If UBound(EventDesc) > 0 Then
    ReDim Preserve EventDesc(UBound(EventDesc) - 1) As EventDescriptor
  End If

End Sub

Private Sub Generate_ImplementsArray()

  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim L_CodeLine   As String
  Dim strTemp      As String
  Dim MaxFactor    As Long
  Dim CurCompCount As Long
  Dim StartLine    As Long
  Dim CompMod      As CodeModule
  Dim GuardLine    As Long

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'ver 1.0.95 moved to here to be called before Generate_ControlArray so that
  'ver 1.1.79' simplifed to use Find
  'Implements targets are recognized as controls
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        ModuleMessage Comp, CurCompCount
        StartLine = 0
        GuardLine = 0
        Set CompMod = Comp.CodeModule
        MaxFactor = CompMod.CountOfDeclarationLines
        Do While CompMod.Find("Implements", StartLine, 1, MaxFactor + 1, -1, True, True, False)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          MemberMessage "", StartLine, MaxFactor
          L_CodeLine = Trim$(CompMod.Lines(StartLine, 1))
          If InCode(L_CodeLine, InStr(L_CodeLine, "Implements ")) Then
            strTemp = AccumulatorString(strTemp, Split(L_CodeLine)(1))
            'v 2.0.7 Thanks Paul Caton
            'This deals with implements of the form RefObject.EventName
            'CF treats them as legit controls
            If InStr(Split(L_CodeLine)(1), ".") Then
              strTemp = AccumulatorString(strTemp, Mid$(Split(L_CodeLine)(1), InStr(Split(L_CodeLine)(1), ".") + 1))
            End If
          End If
          StartLine = StartLine + 1
          GuardLine = StartLine
          If StartLine > CompMod.CountOfDeclarationLines Then
            Exit Do
          End If
        Loop
      End If
    Next Comp
    If bAborting Then
      Exit For 'Sub
    End If
  Next Proj
  If Len(strTemp) Then
    FillArray ImplementsArray, strTemp, , , True
   Else
    ImplementsArray = Split("")
  End If

End Sub

Private Function GetControlCaption(ByVal strName As String) As String

  Dim I As Long

  If bModDescExists Then
    For I = LBound(ModDesc) To UBound(ModDesc)
      If ModDesc(I).MDName = strName Then
        If Len(ModDesc(I).MDCaption) Then
          GetControlCaption = ModDesc(I).MDCaption
        End If
        Exit For
      End If
    Next I
  End If

End Function

Private Sub GetImageListLinks(vbc As VBControl, _
                              ByVal CDC As Long)

  On Error Resume Next
  With CntrlDesc(CDC)
    If ControlHasProperty(vbc, "imagelist") Then
      If LenB(vbc.Properties("imagelist").Value.Item("name")) Then
        .CDImageListLink = "ImageList:" & .CDProj & "|" & .CDForm & "|" & vbc.Properties("imagelist").Value.Item("name").Value
      End If
    End If
    If ControlHasProperty(vbc, "ColumnHeaderIcons") Then
      If LenB(vbc.Properties("ColumnHeaderIcons").Value.Item("name").Value) Then
        .CDImageListLink = AccumulatorString(.CDImageListLink, "ColumnHeaderIcons:" & .CDProj & "|" & .CDForm & "|" & vbc.Properties("ColumnHeaderIcons").Value.Item("name").Value, , False)
      End If
    End If
    If ControlHasProperty(vbc, "icons") Then
      If LenB(vbc.Properties("icons").Value.Item("name").Value) Then
        .CDImageListLink = AccumulatorString(.CDImageListLink, "Icons:" & .CDProj & "|" & .CDForm & "|" & vbc.Properties("icons").Value.Item("name").Value, , False)
      End If
    End If
    If ControlHasProperty(vbc, "smallicons") Then
      If LenB(vbc.Properties("smallicons").Value.Item("name").Value) Then
        .CDImageListLink = AccumulatorString(.CDImageListLink, "SmallIcons:" & .CDProj & "|" & .CDForm & "|" & vbc.Properties("smallicons").Value.Item("name").Value, , False)
      End If
    End If
  End With
  On Error GoTo 0

End Sub

Public Function HasSingletonCArray(ByVal lngBadNameType As Long) As Boolean

MultiFormMsg:
  Select Case lngBadNameType
    '   Case BNNone, BNClass, BNReserve, BNKnown, BNCommand, BNVariable, BNProc
    '   Case BNMultiForm, BNDefault, BNSingle, BNStructural
   Case BNSingletonArray
    If lngBadNameType = BNSingletonArray Then
      HasSingletonCArray = True
      'Exit Function
    End If
   Case Is > BNSingletonArray
    If lngBadNameType > BNMultiForm Then
      lngBadNameType = lngBadNameType - BNMultiForm
      GoTo MultiFormMsg
     Else
      lngBadNameType = lngBadNameType - BNSingletonArray
      HasSingletonCArray = True
      ' Exit Function
    End If
  End Select

End Function

Public Sub HideColumnsReturnPosition(Lsv As ListView, _
                                     Optional ListPos As Long)

  Dim I As Long

  ListPos = GetCurrentLSVLine(Lsv)
  With Lsv
    .ListItems.Clear
    For I = 1 To .ColumnHeaders.Count
      .ColumnHeaders(I).Width = 0
    Next I
  End With

End Sub

Private Function IsAnyControlDeletable() As Boolean

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDUsage = 0 Then
        If LenB(CntrlDesc(I).CDImageListLinkedTo) = 0 Then
          If LenB(CntrlDesc(I).CDImageListLink) = 0 Then
            If CntrlDesc(I).CDClass <> "Menu" Then
              If Not IsGraphic(CntrlDesc(I).CDClass) Then
                If Not IsFileTool(CntrlDesc(I).CDClass) Then
                  If Not CntrlDesc(I).CDIsContainer Then
                    IsAnyControlDeletable = True
                    Exit For
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next I
   Else
    IsAnyControlDeletable = False
  End If

End Function

Public Function IsFileHidden(ByVal FName As String) As Boolean

  If LenB(FName) Then
    IsFileHidden = FSO.GetFile(FName).Attributes And Hidden
  End If

End Function

Private Sub LVSetColWidth(lv As ListView, _
                          ByVal ColumnIndex As Long, _
                          ByVal lvsStyle As LVSCW_Styles)

  '  Copyright ©1997, Karl E. Peterson
  '  http://www.mvps.org/vb
  '------------------------------------------------------------------------------
  '--- If you include the header in the sizing then the last column will
  '--- automatically size to fill the remaining listview width.
  '------------------------------------------------------------------------------

  With lv
    ' verify that the listview is in report view and that the column exists
    If .View = lvwReport Then
      If ColumnIndex >= 1 Then
        If ColumnIndex <= .ColumnHeaders.Count Then
          SendMessage .hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal lvsStyle
        End If
      End If
    End If
  End With

End Sub

Public Function ModuleType(codeMod As CodeModule) As String

  'Based on code found at
  'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42065&lngWId=1
  'Submitted on: 1/1/2003 4:26:20 PM
  'By: Mark Nemtsas

  Select Case codeMod.Parent.Type
   Case vbext_ct_ClassModule
    ModuleType = "Class"
   Case vbext_ct_MSForm
    ModuleType = "Form"
   Case vbext_ct_StdModule
    ModuleType = "Module"
   Case vbext_ct_UserControl
    ModuleType = "UserControl"
   Case vbext_ct_VBForm
    ModuleType = "Form"
   Case vbext_ct_VBMDIForm
    ModuleType = "VBMDIForm"
   Case vbext_ct_ActiveXDesigner
    ModuleType = "ActiveXDesigner"
    'ModuleType = "DataReport"
    'ModuleType = "AddInDesigner
    'ModuleType = "DataEnvironment"
    'ModuleType = "DHTMLPage"
   Case vbext_ct_ResFile
    ModuleType = "Resource File"
   Case vbext_ct_RelatedDocument
    ModuleType = "Related Document"
   Case vbext_ct_PropPage
    ModuleType = "PropertyPage"
   Case vbext_ct_DocObject
    ModuleType = "DocObject"
    'Case Else
  End Select

End Function

Public Function NeedsXPFrameFix(CCtrl As VBControl) As Boolean

  With CCtrl
    If ArrayMember(.ClassName, "CommandButton", "OptionButton", "Frame", "Label") Then
      If ArrayMember(.ClassName, "CommandButton", "OptionButton") Then
        NeedsXPFrameFix = .Properties("style") = 0
        'Graphic option/command-buttons (= 1) are immune to the XP Frame bug
       ElseIf ArrayMember(.ClassName, "Frame", "Label") Then
        'captionless frames dont need it either
        NeedsXPFrameFix = Len(.Properties("caption")) > 0
      End If
    End If
  End With 'CCtrl

End Function

Public Function PosInList(ByVal strA As String, _
                          Lst As ListBox, _
                          Optional CaseSensitive As Boolean = True) As Long

  Const LB_FINDSTRINGEXACT As Long = &H1A2
  Const LB_FINDSTRING      As Long = &H18F

  PosInList = SendMessage(Lst.hWnd, IIf(CaseSensitive, LB_FINDSTRINGEXACT, LB_FINDSTRING), 0, ByVal strA)

End Function

Private Sub SetCDImageListLinkedTo()

  Dim I      As Long
  Dim J      As Long
  Dim K      As Long
  Dim strTmp As String
  Dim arrTmp As Variant

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If LenB(CntrlDesc(I).CDImageListLink) Then
        arrTmp = Split(CntrlDesc(I).CDImageListLink, ",")
        For K = LBound(arrTmp) To UBound(arrTmp)
          For J = LBound(CntrlDesc) To UBound(CntrlDesc)
            With CntrlDesc(J)
              If InStr(arrTmp(K), "|" & .CDName) Then
                If InStr(arrTmp(K), "|" & .CDForm & "|") Then
                  If InStr(arrTmp(K), ":" & .CDProj & "|") Then
                    strTmp = CntrlDesc(I).CDImageListLink
                    strTmp = Replace$(strTmp, .CDName, vbNullString)
                    strTmp = Replace$(strTmp, .CDForm, vbNullString)
                    strTmp = Replace$(strTmp, .CDProj, vbNullString)
                    strTmp = Replace$(strTmp, ",", vbNullString)
                    strTmp = Left$(strTmp, Len(strTmp) - 3)
                    strTmp = .CDProj & "|" & CntrlDesc(I).CDForm & "|" & CntrlDesc(I).CDName & ":" & strTmp
                    .CDImageListLinkedTo = AccumulatorString(.CDImageListLinkedTo, strTmp)
                  End If
                End If
              End If
            End With 'CntrlDesc(J)
          Next J
        Next K
      End If
    Next I
  End If

End Sub

Public Sub SetControlData(CDC As Long, _
                          vbc As VBControl, _
                          Comp As VBComponent, _
                          vbf As VBForm, _
                          Optional ByVal bSingleRename As Boolean = True, _
                          Optional ByVal bHasImageLists As Boolean = False)
Dim I As Long
Dim J As Long
  With CntrlDesc(CDC)
    'v 2.2.4 correction
    .CDProj = Comp.Collection.Parent.Name
    .CDForm = Comp.Name
    .CDIndex = vbc.Properties("index").Value
    .CDClass = vbc.ClassName
    'v2.2.1 Thanks rblanch
    'Added explicit Value property allows CF to get through a complex bug in source code.
    'If a control has a damaged link to an ImageList (errors in the relevant frx file)
    'then on calling the next control CF crashed because the no name test is designed
    'for a different problem in UserControls. Explicit Value call reestablishes CF link to
    'Code's controls Collection
    .CDName = vbc.Properties("name") '.Value
    If LenB(.CDName) = 0 Then
      'v3.0.9  retry ( for some reason some UserControls disrupt the detection of names, but get it the 2nd try)
      .CDName = vbc.Properties("name")
      'v 2.0.5 new error trap for badly built UsedControls
      'v 2.1.1 added test
      'v 2.1.4 seperate sub added
      'v 2.2.1 Thanks rblanch 2. added extra test to generate different message boxes for above problem
      If LenB(.CDName) = 0 Then
        If IsComponent_UserControl(Comp.Type) Then
          If LenB(.CDName) = 0 Then
            BadUserControlMsg vbc, Comp, .CDProj, .CDForm
            Err.Clear
            bForceReload = True
            bAborting = True
            GoTo SafeExit
          End If
        Else
          'v3.0.9 added caption to help ID problem (inactive until tested further)
'          If ControlHasProperty(vbc, "Caption") Then
'            If ControlHasPropertyValid(vbc, "Caption") Then
'              .CDCaption = ShortCaption(vbc.Properties("Caption").Value)
'            End If
'          End If
          'mObjDoc.Safe_MsgBox "Code Fixer has encountered an error which suggests that your project has a problem with an FRX file or with a '" & .CDClass & "' " & IIf(LenB(.CDCaption), .CDCaption, "") & " on " & .CDForm & vbNewLine &
          mObjDoc.Safe_MsgBox "Code Fixer has encountered an error which suggests that your project has a problem with an FRX file or with a '" & .CDClass & "' on " & .CDForm & vbNewLine & _
                    "Check for a log file in the source folder(s). It is probably an error of the type:" & vbNewLine & _
                    "'Property X in <formName|ControlName> had an invalid file reference.'" & vbNewLine & _
                    "This may mean that the form|control has a damaged or missing binary data file." & vbNewLine & _
                    "Fix 1 is to Update the UserControls on the Form." & vbNewLine & _
                    "Fix 2 is to Save the Form, forcing VB to recreate the binary data file(s)." & vbNewLine & _
                    "Fix 3 If the controls are Locked on the form unlock them (PadLock icon on menubar or Form's Right-Click Context menu)." & vbNewLine & _
                    vbNewLine & _
                    "NOTE You may have to close VB and restart for Code Fixer to operate." & vbNewLine & _
                    vbNewLine & _
                    "Code Fixer will now abort, please wait.", vbCritical
          Err.Clear
          bForceReload = True
          bAborting = True
          GoTo SafeExit
        End If
      End If
    End If
    'v2.3.8 keeps arrQCtrlPresence in sync when you rename controls
    If bSingleRename Then
      'v 2.4.4 don't do this if generating whole list
      If Len(.CDOldName) Then
        If .CDOldName <> .CDName Then
          If .CDIndex = -1 Then
            arrQCtrlPresence = QuickSortRemove(arrQCtrlPresence, .CDOldName)
          Else
            'v3.0.7 move inside the rename trap because it's only needed here (speed increase)
            If GetControlArraySizeFromName(.CDName, vbf.VBControls) = 1 Then
              arrQCtrlPresence = QuickSortRemove(arrQCtrlPresence, .CDOldName)
            End If
          End If
        End If
      End If
      arrQCtrlPresence = QuickSortAppend(arrQCtrlPresence, .CDName)
    End If
    If vbc.ContainedVBControls.Count > 0 Then
      .CDIsContainer = True
      If vbc.ContainedVBControls.Count Then
        For J = 1 To vbc.ContainedVBControls.Count
          .CDContains = AccumulatorString(.CDContains, vbc.ContainedVBControls(J).Properties("name").Value & IIf(vbc.ContainedVBControls(J).Properties("index") > -1, strInBrackets(vbc.ContainedVBControls(J).Properties("index").Value), vbNullString))
        Next J
      End If
    End If
    If ControlHasProperty(vbc, "style") Then
      ' SStab Style value is not relevant
      ' ImageCombo has a Property Style but it is not available for programming
      ' an interesting bug that suggests that it was not finished properly
      ' this produced a CF crash until I blocked it out
      ' thanks to Jimmy C. Broadhead, Jr. whose 'mySQL Query Analyzer' showed me this bug
      'v2.3.0 simplified and copes with anything with a 'Style' Property
      'Thanks to Dean Camera and his Deeplook 4.7 which triggered this error
      'and lead me to a more generalized fix
      On Error Resume Next
      .CDStyle = vbc.Properties("style").Value
      If Err Then
        .CDStyle = 0
      End If
      On Error GoTo 0
    End If
    .CDDefProp = ReferenceLibraryControlDefaultProperty(.CDClass)
    .CDOldName = .CDName
    .CDOldIndex = .CDIndex
    .CDFullName = .CDName
    If .CDIndex > -1 Then
      .CDFullName = .CDName & strInBrackets(.CDIndex)
    End If
    If .CDIndex = -1 Then
      .CDUsage = IIf(DOPartialUsageX(.CDName, True), 1, 0)
    Else 'Usage = 0 Then
      .CDUsage = IIf(DOPartialUsageX(.CDName, False), 1, 0)
    End If
    'strOneCaption = vbNullString
    If vbc.ClassName = "Menu" Then
      .CDCaption = ShortCaption(vbc.Properties("caption").Value)
    ElseIf ControlHasProperty(vbc, "Caption") Then
      If ControlHasPropertyValid(vbc, "Caption") Then
        .CDCaption = ShortCaption(vbc.Properties("Caption").Value)
      End If
    End If
    '.CDCaption = strOneCaption
    If .CDClass = "Frame" Then
      For I = 1 To vbc.ContainedVBControls.Count
        If NeedsXPFrameFix(vbc.ContainedVBControls.Item(I)) Then
          .CDXPFrameBug = True
          Exit For
        End If
      Next I
    End If
    If bSingleRename Or bHasImageLists Then
      GetImageListLinks vbc, CDC
    End If
SafeExit:

  End With 'CntrlDesc(CDC)
End Sub

Public Sub SetModuleListitem()

  Dim Comp As VBComponent
  Dim I    As Long

  With frm_CodeFixer
    Set Comp = GetComponent(.lsvAllModules.SelectedItem.Text, .lsvAllModules.SelectedItem.SubItems(1))
    Comp.CodeModule.CodePane.Show
    Comp.Activate
    For I = 0 To 3
      .cmdEditMod(I).Enabled = False
    Next I
    I = GetHiddenDescriptorIndex(.lsvAllModules, 3)
    .txtModuleEdit(0) = ModDesc(I).MDName
    .txtModuleEdit(1) = ModDesc(I).MDFilename
    .txtModuleEdit(2) = ModDesc(I).MDFullPath
  End With
  SuggestModName

End Sub

Private Function ShortCaption(strC As String) As String

  If Len(strC) Then
    ShortCaption = Trim$(Replace$(strC, ",", SngSpace))
    If Len(ShortCaption) > 20 Then
      ShortCaption = Left$(ShortCaption, 20)
    End If
  End If

End Function

Public Sub SmartAdd(L As ListBox, _
                    ByVal strAdd As String, _
                    strGuard As String)

  'v2.4.9 Lcase to stop Frame1 > fraMe1

  If InStr(LCase$(strGuard), LCase$("*" & strAdd & "*")) = 0 Then
    L.AddItem strAdd
    strGuard = strGuard & "*" & strAdd & "*"
  End If

End Sub

Private Sub SuggestModName()

  Dim strCurName     As String
  Dim strCurFileName As String
  Dim strExt         As String
  Dim StrProjName    As String
  Dim StrProjName2   As String
  Dim StrGotOne      As String

  strCurName = frm_CodeFixer.lsvAllModules.SelectedItem.ListSubItems(1)
  strCurFileName = frm_CodeFixer.lsvAllModules.SelectedItem.ListSubItems(2)
  strExt = LCase$(FileExtention(strCurFileName))
  StrProjName = frm_CodeFixer.lsvAllModules.SelectedItem.Text
  If LCase$(Left$(StrProjName, 3)) = "prj" Then
    StrProjName2 = Mid$(frm_CodeFixer.lsvAllModules.SelectedItem.Text, 4)
  End If
  With frm_CodeFixer.lstSuggestModName
    .Clear
    SendMessage .hWnd, WM_SETREDRAW, False, 0
    StrGotOne = "*" & strCurName & "*"
    SmartAdd frm_CodeFixer.lstSuggestModName, Ucase1st(strCurName), StrGotOne
    If LCase$(Left$(strCurName, 3)) <> strExt Then
      SmartAdd frm_CodeFixer.lstSuggestModName, FileExtention(strCurFileName) & Ucase1st(strCurName), StrGotOne
      SmartAdd frm_CodeFixer.lstSuggestModName, FileExtention(strCurFileName) & "_" & Ucase1st(strCurName), StrGotOne
    End If
    If LCase$(Left$(strCurName, 1)) <> Left$(strExt, 1) Then
      SmartAdd frm_CodeFixer.lstSuggestModName, Left$(strExt, 1) & Ucase1st(strCurName), StrGotOne
      SmartAdd frm_CodeFixer.lstSuggestModName, Left$(strExt, 1) & "_" & Ucase1st(strCurName), StrGotOne
    End If
    If Len(GetControlCaption(strCurName)) Then
      With frm_CodeFixer
        SmartAdd .lstSuggestModName, strExt & CleanCaption(GetControlCaption(strCurName)), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & CleanCaption(GetControlCaption(strCurName)), StrGotOne
        SmartAdd .lstSuggestModName, strExt & "_" & CleanCaption(GetControlCaption(strCurName)), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & "_" & CleanCaption(GetControlCaption(strCurName)), StrGotOne
      End With
    End If
    If LCase$(strExt) = "bas" Then
      With frm_CodeFixer
        SmartAdd .lstSuggestModName, "mod" & CleanCaption(strCurName), StrGotOne
        SmartAdd .lstSuggestModName, "m" & CleanCaption(strCurName), StrGotOne
        SmartAdd .lstSuggestModName, "mod_" & CleanCaption(strCurName), StrGotOne
        SmartAdd .lstSuggestModName, "m_" & CleanCaption(strCurName), StrGotOne
      End With
    End If
    If Len(StrProjName) Then
      With frm_CodeFixer
        SmartAdd .lstSuggestModName, strExt & Ucase1st(StrProjName), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & Ucase1st(StrProjName), StrGotOne
        SmartAdd .lstSuggestModName, strExt & "_" & Ucase1st(StrProjName), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & "_" & Ucase1st(StrProjName), StrGotOne
      End With
      If LCase$(strExt) = "bas" Then
        With frm_CodeFixer
          SmartAdd .lstSuggestModName, "mod" & Ucase1st(StrProjName), StrGotOne
          SmartAdd .lstSuggestModName, "m" & Ucase1st(StrProjName), StrGotOne
          SmartAdd .lstSuggestModName, "mod_" & Ucase1st(StrProjName), StrGotOne
          SmartAdd .lstSuggestModName, "m_" & Ucase1st(StrProjName), StrGotOne
        End With
      End If
    End If
    If Len(StrProjName2) Then
      With frm_CodeFixer
        SmartAdd .lstSuggestModName, strExt & Ucase1st(StrProjName2), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & Ucase1st(StrProjName2), StrGotOne
        SmartAdd .lstSuggestModName, strExt & "_" & Ucase1st(StrProjName2), StrGotOne
        SmartAdd .lstSuggestModName, Left$(strExt, 1) & "_" & Ucase1st(StrProjName2), StrGotOne
      End With
    End If
    SendMessage .hWnd, WM_SETREDRAW, True, 0
  End With

End Sub

Public Sub suggestProjName()

  Dim strCurName     As String
  Dim strCurFileName As String
  Dim strExt         As String
  Dim StrGotOne      As String

  strCurName = frm_CodeFixer.lsvAllProjects.SelectedItem.Text
  strCurFileName = FileNameOnly(VBInstance.VBProjects.Item(frm_CodeFixer.lsvAllProjects.SelectedItem.Text).FileName)
  strExt = LCase$(FileExtention(strCurFileName))
  With frm_CodeFixer.lstSuggestProjName
    .Clear
    SendMessage .hWnd, WM_SETREDRAW, False, 0
    StrGotOne = "*" & strCurName & "*"
    SmartAdd frm_CodeFixer.lstSuggestProjName, Ucase1st(strCurName), StrGotOne
    If LCase$(Left$(strCurName, 3)) <> strExt Then
      SmartAdd frm_CodeFixer.lstSuggestProjName, FileExtention(strCurFileName) & Ucase1st(strCurName), StrGotOne
      SmartAdd frm_CodeFixer.lstSuggestProjName, FileExtention(strCurFileName) & "_" & Ucase1st(strCurName), StrGotOne
    End If
    If LCase$(Left$(strCurName, 1)) <> Left$(strExt, 1) Then
      SmartAdd frm_CodeFixer.lstSuggestProjName, Left$(strExt, 1) & Ucase1st(strCurName), StrGotOne
      SmartAdd frm_CodeFixer.lstSuggestProjName, Left$(strExt, 1) & "_" & Ucase1st(strCurName), StrGotOne
    End If
    If LCase$(Left$(strCurName, 3)) <> "prj" Then
      With frm_CodeFixer
        SmartAdd .lstSuggestProjName, "prj" & Ucase1st(strCurName), StrGotOne
        SmartAdd .lstSuggestProjName, "p" & Ucase1st(strCurName), StrGotOne
        SmartAdd .lstSuggestProjName, "prj_" & Ucase1st(strCurName), StrGotOne
        SmartAdd .lstSuggestProjName, "p_" & Ucase1st(strCurName), StrGotOne
      End With
    End If
    SendMessage .hWnd, WM_SETREDRAW, True, 0
  End With

End Sub

Private Function TotalControlCount() As Long

  Dim Comp As VBComponent
  Dim Proj As VBProject

  'v2.4.3 seperated to stop any errors flowing through
  For Each Proj In VBInstance.VBProjects ' count the controls
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        If IsComponent_ControlHolder(Comp) Then
          ' only control bearers have a Designer so skip others
          On Error GoTo EventNotMatch
          TotalControlCount = TotalControlCount + Comp.Designer.VBControls.Count
          'FIXME a crash here means that a UserControl contains an Event that doesn't match its description.
        End If
      End If
    Next Comp
  Next Proj
  On Error GoTo 0

Exit Function

EventNotMatch:
  mObjDoc.Safe_MsgBox "An instance of a UserControl on " & Comp.Name & " contains an error," & vbNewLine & _
                    "possibly an Event that doesn't match its description" & vbNewLine & _
                    "or a missing reference.", vbCritical
  Err.Clear
  Resume Next

End Function

Public Function UnTouchable(ByVal strName As String) As Boolean

  Dim I As Long

  'this routine test whether or not a control is checked
  'Thanks to Neil who suggested that you should be able to turn off
  'Code Fixer for modules
  With frm_FindSettings.lsvModNames
    For I = 1 To .ListItems.Count
      If .ListItems(I).Text = strName Then
        UnTouchable = .ListItems(I).Checked = False
        Exit For
      End If
    Next I
  End With

End Function

Public Sub UpdateControlsCaption(Lsv As ListView)

  Dim I               As Long
  Dim lngSepNameCount As Long
  Dim lngArrayCount   As Long
  Dim strPrevName     As String ' allows control array counter to increase
  Dim BadCount        As Long
  Dim XPCount         As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDXPFrameBug Then
        XPCount = XPCount + 1
      End If
      If CntrlDesc(I).CDBadType <> 0 Then
        BadCount = BadCount + 1
      End If
      With CntrlDesc(I)
        If .CDUsage <> 2 Then
          If .CDName <> strPrevName Then
            strPrevName = .CDName
            lngSepNameCount = lngSepNameCount + 1
            If .CDIndex <> -1 Then
              lngArrayCount = lngArrayCount + 1
            End If
          End If
        End If
      End With 'CntrlDesc(I)
NotBad:
    Next I
    With frm_CodeFixer
      .frapage(TPControls).Caption = "Controls [" & Lsv.ListItems.Count & "]" & IIf(lngSepNameCount > 0 And lngSepNameCount <> Lsv.ListItems.Count, " [" & lngSepNameCount & " Names / " & lngArrayCount & " Control Arrays]", "[No Control Arrays]") & IIf(BadCount, "[" & BadCount & " Poorly Named]", vbNullString) & IIf(XPCount, "[" & XPCount & " XP-Frame Bug]", vbNullString)
    End With
  End If

End Sub

Public Sub UpDateOldNameIndex()

  Dim I As Long

  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      With CntrlDesc(I)
        .CDOldName = .CDName
        .CDOldIndex = .CDIndex
      End With
    Next I
  End If

End Sub
Public Function UserControlFontRefFix() As Boolean
Dim Comp         As VBComponent
Dim Proj         As VBProject
Dim vbf          As VBForm
Dim Sline        As Long
Dim CurCompCount As Long
Dim ArrProc      As Variant
Dim TopRPline    As Long
Dim I            As Long
Dim UpDated      As Boolean
Dim bUpDateMsg   As Boolean
Dim PSLine       As Long
Dim PEndLine     As Long
Dim UpDateCount As Long
Dim UpDatedName As String
  'v3.0.9 Thanks to STabXP 1.1.0 txtCodeId=59595 (author unclear)
  'If a user control sets it's own font in code then there is a possiblility that other
  'properties will fail if the font has not been set yet
  'This fix moves the ReadProperties line that sets the Font to the top of the ReadProperties cycle
  'This is rarely a real problem in code but is one cause of Code Fixer not being able to get the
  'Name property for a UserControl because the UC doesn't initialize properly when CF calls for the Name
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        If dofix(CurCompCount, UCFontFix) Then
          If Comp.CodeModule.Parent.Type = vbext_ct_UserControl Then
            ActivateDesigner Comp, vbf, False
            Sline = 1
            TopRPline = 0
            With Comp.CodeModule
              If .Find("Private Sub UserControl_ReadProperties(PropBag As PropertyBag)", Sline, 1, -1, -1, True) Then
              'get the UserControl_ReadProperties procedure
                ArrProc = ReadProcedureCodeArray2(Comp.CodeModule, Sline, PSLine, PEndLine)
              'search for first non-font read property line
                For I = LBound(ArrProc) To UBound(ArrProc)
                  If InCode(ArrProc(I), InStr(ArrProc(I), "ReadProperty")) Then
                    If InCode(ArrProc(I), InStr(ArrProc(I), "Font")) = 0 Then
                     'ignore if already is Font setter
                      TopRPline = I
                      Exit For
                    End If
                  End If
                Next
                If TopRPline Then
               'search for Font setting lines that ar not at top of code
               'and move them to the top
                  For I = TopRPline + 1 To UBound(ArrProc)
                    If InCode(ArrProc(I), InStr(ArrProc(I), "Font")) Then
                      SwapAnyThing ArrProc(I), ArrProc(TopRPline)
                      UpDated = True
                      ArrProc(TopRPline) = ArrProc(TopRPline) & WARNING_MSG & "Font property setting moved to top of routine"
                      TopRPline = TopRPline + 1
                    End If
                  Next
                End If
                If UpDated Then
                  ReplaceProcedureCode Comp.CodeModule, ArrProc, PSLine, PEndLine, False
                  'comp.Designer.controls
                  UpDated = False
                  bUpDateMsg = True
                  UpDateCount = UpDateCount + 1
                  UpDatedName = UpDatedName & ", " & Comp.Name
                End If
              End If
            End With 'Comp
          End If
        End If
      End If
      Set vbf = Nothing
    Next Comp
  Next Proj
  If bUpDateMsg Then
  UpDatedName = strInDQuotes(Left(UpDatedName, Len(UpDatedName) - 2))
  
    mObjDoc.Safe_MsgBox "Code Fixer has rearranged the order of ReadProperties in " & IIf(UpDateCount = 1, "a UserControl:", UpDateCount & " UserControls:") & UpDatedName & vbNewLine & _
                    "Properties which set the Font of the control now occur first." & vbNewLine & _
                    "This should decrease the number of times error trapping needs to deal with the Font not being set." & vbNewLine & _
                    "You should check that this has not disrupted the order of other PRoprties being initalised." & vbNewLine & _
                    "You will need to update the controls before Code Fixer can process the code." & vbNewLine & _
                    "You may need to restart VB.", vbCritical
    UserControlFontRefFix = True
  End If
End Function
Public Sub UserControlActiveTimerFix()

  'v3.0.4 improved UserControl set up
  'This fix works by inserting code into the UserControl_ReadProperties routine(created if necessary)
  'This procedure is used because it always fires.
  '(If InvisibleAtRuntime is True then Resize and other drawing routines don't hit)
  '(Initialize doesn't work because UserControl.Ambient is not yet initialised)
  '
  'Only inserts the code if the Timer is Enabled.
  '
  
  Dim Comp         As VBComponent
  Dim Proj         As VBProject
  Dim vbc          As VBControl
  Dim vbf          As VBForm
  Dim strTimerFix  As String
  Dim Sline        As Long
  Dim CurCompCount As Long

  strTimerFix = ".Enabled = UserControl.Ambient.UserMode = True" & vbNewLine & _
                EARLYWARNING_MSG & "Active UserControl Timers make coding difficult this will turn Timer OFF in IDE and ON when running" & vbNewLine & _
                EARLYWARNING_MSG & "NOTE the Timer has been set Enabled = False"
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp, CurCompCount, False) Then
        If dofix(CurCompCount, UCTimerDisable) Then
          If Comp.CodeModule.Parent.Type = vbext_ct_UserControl Then
            ActivateDesigner Comp, vbf, False
            If vbf Is Nothing Then
              ActivateDesigner Comp, vbf, True
            End If
            For Each vbc In vbf.VBControls
              If vbc.ClassName = "Timer" Then
                If vbc.Properties("Enabled") Then
                  Sline = 1
                  With Comp.CodeModule
                    If .Find("Private Sub UserControl_ReadProperties(PropBag As PropertyBag)", Sline, 1, -1, -1, True) Then
                      'Sub already exists so just insert the code
                      If Not JustACommentOrBlank(.Lines(Sline, 1)) Then
                        'but first check that it is actually code and not a comment with the trigger words in it
                        .InsertLines Sline + 1, vbc.Properties("Name") & IIf(vbc.Properties("Index") > -1, "(" & vbc.Properties("Name") & ")", "") & strTimerFix
                       Else
                        .AddFromString "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine & _
                         vbc.Properties("Name") & IIf(vbc.Properties("Index") > -1, "(" & vbc.Properties("Name") & ")", "") & strTimerFix & vbNewLine & _
                         "End Sub"
                      End If
                     Else
                      'Sub doesn't exist so create it
                      .AddFromString "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine & _
                       vbc.Properties("Name") & IIf(vbc.Properties("Index") > -1, "(" & vbc.Properties("Name") & ")", "") & strTimerFix & vbNewLine & _
                       "End Sub"
                    End If
                  End With 'Comp
                  vbc.Properties("Enabled") = False
                End If
              End If
            Next vbc
          End If
        End If
      End If
      Set vbf = Nothing
    Next Comp
  Next Proj

End Sub

Private Function UserControlUsesResumeNext(Comp As VBComponent) As Boolean

  Dim lngdummy As Long

  UserControlUsesResumeNext = Comp.CodeModule.Find("On Error Resume", 1, lngdummy, lngdummy, lngdummy)

End Function

Public Function UsingDefVBName(ByVal strName As String, _
                               strCtrlClas As Variant) As Boolean

  Dim strTmp As String

  'once all the fixes are done do this for safety reasons
  If strCtrlClas <> "Menu" Then ' numbered menus don't count as VB doesn't create their names
    strTmp = strName
    Do While IsNumeric(Right$(strTmp, 1))
      strTmp = Left$(strTmp, Len(strTmp) - 1)
    Loop
    If SmartLeft(strCtrlClas, strTmp) And strTmp <> strName Then
      '*if name is left or whole of class name then is is the VB defname
      If strCtrlClas = strTmp Then 'Exactly same as Class (Label# etc)
        UsingDefVBName = True
       ElseIf Mid$(strCtrlClas, Len(strTmp) + 1, 1) = UCase$(Mid$(strCtrlClas, Len(strTmp) + 1, 1)) Then
        'VB def names seem to be based on part of classname up to the second embedded Capital so
        'PictureBox is detected as Picture# but Pic# is not.
        UsingDefVBName = True
      End If
    End If
  End If

End Function

''Sub TestNeedsUCActiveTimerFix()
''Dim Comp        As VBComponent
''Dim Proj        As VBProject
''Dim vbc         As VBControl
''Dim vbf         As VBForm
'' Dim Hit As Boolean
''  For Each Proj In VBInstance.VBProjects
''    For Each Comp In Proj.VBComponents
''   If LenB(Comp.Name) Then
''      If Comp.CodeModule.Parent.Type = vbext_ct_UserControl Then
''        ActivateDesigner Comp, vbf, False
''        If vbf Is Nothing Then
''          ActivateDesigner Comp, vbf, True
''        End If
''        For Each vbc In vbf.VBControls
''          If vbc.ClassName = "Timer" Then
''            If vbc.Properties("Enabled") = True Then
''            Hit = True
''            DisplayCodePane Comp, True
''            Exit For
''            End If
''          End If
''        Next
''      End If
''      If Hit Then
''      Exit For
''      End If
''       Next
''      If Hit Then
''      Exit For
''      End If
''      Next
''
''End Sub
''

':)Code Fixer V3.0.9 (25/03/2005 4:12:40 AM) 70 + 1741 = 1811 Lines Thanks Ulli for inspiration and lots of code.

