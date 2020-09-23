Attribute VB_Name = "mod_Settings"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
Public Const offset              As Long = 120
Private bWorking                 As Boolean
Public Enum FixMode
  Off = 0
  CommentOnly = 1
  FixAndComment = 2
  JustFix = 3
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private Off, CommentOnly, FixAndComment, JustFix
#End If
'OK this will work because I lifted the values from registry
'v2.2.6 Note added so I update properly. Thanks Mike Ulik
''Don't forget to update SettingsEngine and RefreshUserSettings <ArrEnd> when adding/deleting fixes
'                                               Declaration     |Restructure                |Param  |Dim       |Suggest      |Unused       |Format
Public Const DefNoSettings       As String = "033300201101200000|0000000000000000000000000000|0000000|00000000000|0000000000000|0000000000000|0300000"
Public Const DefMinSettings      As String = "133310211111211213|3333332111111111111111011111|1111111|11112111133|1111111111111|1111111111110|1333333"
Public Const DefAvgSettings      As String = "133310222222222213|3333332122222222222222222222|2221222|22222222233|1111111111111|1122222222220|2333333"
Public Const DefMaxSettings      As String = "133310333333333313|3333332223333333222222222222|2221222|33333323333|1111111111111|1122222222222|3333333"
Public UserSettings              As String
Public Type NFix
  FixLevel                       As FixMode
  FixModCount                    As Long
  FixTotalCount                  As Long
End Type
Public Enum NfixDesc
  'Declaration Section Fixers
  LargeSourceWarning_CXX
  RemoveIndent
  RemoveCodeFixComments
  RemoveLineContinuation
  XPStyleTest
  MoveCommentInside
  InsertOptExp
  MoveAPIDown
  DimGlobal2PublicPrivate
  Public2Private
  UpdateConstExplicit
  ExpandDecSingleLineSingleType
  ExpandDecSingleLine
  UpdateDecTypeSuffix
  UpdateDefType2AsType
  CreateEnumCapProtect
  BadNameVariableWarning_CXX
  LayOutAtTypeEOLComment
  'ReStructure Section Fixers
  UCTimerDisable
  UCFontFix
  RemoveLineNum
  UpDateStrConcat
  UpdateChrToConstant
  UpdateStrFunction
  UpdateErrError
  UpdateInteger2Long
  UnNeededExit
  NIfThenExpand
  SeperateCompounds
  UpdateWend
  NPleonasmFix
  NCloseResumeLoad
  NCloseResume
  WithPurity
  DetectShortCircuit
  DetectZeroLengthStringTests
  DetectZeroLengthStringAssign
  DetectWithStructure
  RoutineCaseFix
  PublicVar2Property
  UnNeededIfBracket
  UnNeededAmpersand
  UnNeededCall
  UneededCaseIs
  CollapseIfThenBool
  CollapseCaseBool
  SecondryExitFix
  'Parameter Fixers
  ParameterExpandTypeSuffix
  ParameterNoType
  ParameterBadName_CXX
  ParameterDead_CXX
  ParameterByVal
  FunctionTypeCast
  ConvertFunction2Sub
  'Local Section Fixers
  MoveDim2Top
  ReWriteDimMultiSingleType
  ReWriteDimExpandMultiDim
  UpdateDimType
  DetectReWriteDimUntyped
  DetectDimUnused
  DetectDimMissing
  DetectDimDuplicate
  NoDimInitialise
  LayOutDim
  CodeTypeSuffixStrip
  'Suggestion Section Fixers
  DetectObsoleteCodeStructure_CXX
  DetectDangerousCode_CXX
  DetectDangerousAPICode_CXX
  DetectDangerousString_CXX
  DetectDangerousReference_CXX
  DetectStatic2Private_CXX
  DetectHardPath_CXX
  DetectBadCtrlName_CXX
  DuplicatePublicProc_CXX
  DetectIllegalWithExit_CXX
  DoElseIfEval_CXX
  GoToComments_CXX
  VeryBigProc_CXX
  'Unused Section Fixers
  EmptyRoutine_CXX
  REmptyStruct
  UnusedDecConst
  UnusedDecAPI
  UnusedDecVariable
  DeletedControlCode
  UnusedFunction
  UnusedSub
  UnusedProperty
  UnusedEvents
  UnusedWithEvents
  ActiveDebugStop
  CommentOutClassmembers
  'Format Section Fixers
  NForNextVar
  FormatCodeFix
  FormatArrayLineContinuation
  FormatStringLineContinuation
  FormatDeclareLineContinuation
  FormatRoutineLineContinuation
  NSortModules
  DummyFixEnd
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private LargeSourceWarning_CXX, RemoveIndent, RemoveCodeFixComments, RemoveLineContinuation
Private XPStyleTest, MoveCommentInside, InsertOptExp, MoveAPIDown, DimGlobal2PublicPrivate, Public2Private
Private UpdateConstExplicit, ExpandDecSingleLineSingleType, ExpandDecSingleLine, UpdateDecTypeSuffix
Private UpdateDefType2AsType, CreateEnumCapProtect, BadNameVariableWarning_CXX, LayOutAtTypeEOLComment
Private UCTimerDisable, UCFontFix, RemoveLineNum, UpDateStrConcat, UpdateChrToConstant, UpdateStrFunction
Private UpdateErrError, UpdateInteger2Long, UnNeededExit, NIfThenExpand, SeperateCompounds, UpdateWend
Private NPleonasmFix, NCloseResumeLoad, NCloseResume, WithPurity, DetectShortCircuit, DetectZeroLengthStringTests
Private DetectZeroLengthStringAssign, DetectWithStructure, RoutineCaseFix, PublicVar2Property, UnNeededIfBracket
Private UnNeededAmpersand, UnNeededCall, UneededCaseIs, CollapseIfThenBool, CollapseCaseBool, SecondryExitFix, ParameterExpandTypeSuffix, ParameterNoType, ParameterBadName_CXX, ParameterDead_CXX
Private ParameterByVal, FunctionTypeCast, ConvertFunction2Sub, MoveDim2Top, ReWriteDimMultiSingleType
Private ReWriteDimExpandMultiDim, UpdateDimType, DetectReWriteDimUntyped
Private DetectDimUnused, DetectDimMissing, DetectDimDuplicate, NoDimInitialise, LayOutDim, CodeTypeSuffixStrip
Private DetectObsoleteCodeStructure_CXX, DetectDangerousCode_CXX, DetectDangerousAPICode_CXX
Private DetectDangerousString_CXX, DetectDangerousReference_CXX, DetectStatic2Private_CXX
Private DetectHardPath_CXX, DetectBadCtrlName_CXX, DuplicatePublicProc_CXX, DetectIllegalWithExit_CXX
Private DoElseIfEval_CXX, GoToComments_CXX, VeryBigProc_CXX, EmptyRoutine_CXX, REmptyStruct, UnusedDecConst
Private UnusedDecAPI, UnusedDecVariable, DeletedControlCode, UnusedFunction, UnusedSub, UnusedProperty
Private UnusedEvents, UnusedWithEvents, ActiveDebugStop, CommentOutClassmembers, NForNextVar
Private FormatCodeFix, FormatArrayLineContinuation, FormatStringLineContinuation, FormatDeclareLineContinuation
Private FormatRoutineLineContinuation, NSortModules, DummyFixEnd
#End If
Public FixData()                 As NFix

Private Sub GetFixSetting2(ByVal LsvNo As Long, _
                           ByVal ListRow As Long, _
                           ByVal FixNo As NfixDesc)

  'This routine just makes code easier to read in SettingCollector by allowing the changing

  FixData(FixNo).FixLevel = Val(Mid$(frm_FindSettings.ctlSettingsList(LsvNo).DataString, ListRow, 1))

End Sub

Private Sub InitalisePropertyList(ByVal UCNo As Long)

  'v2.5.3 general format for PropertyList

  With frm_FindSettings.ctlSettingsList(UCNo)
    .Clear
    .FormatString = "Fix|Mode"
    .DefaultList = "&Off|&Mark|Fi&x && Mark|&Fix"
    .DefaultButtons = True
    .DescriptionShow = True
  End With

End Sub

Public Sub PictureToFrame(P As PictureBox, _
                          F As Frame, _
                          Optional L As Label)

  Dim LOffset As Long

  'Places a PictureBox neatly in a Frame with borders or completely covers a Frame without borders
  'Quick and Dirty :) you may need something more complex in the real world
  'Creator: Roger Gilchrist
  With P
    LOffset = IIf(F.BorderStyle = 0, 0, 100)
    ' if you don't like the offset method adjust the final values below
    '.Move offset, offset * 1.75, F.Width - offset * 2, F.Height - offset * 2.5
    .Move 0, 0, F.Width - LOffset * 2, F.Height - LOffset * 2.5
    .BackColor = F.BackColor
  End With
  If Not L Is Nothing Then
    L.Move 0, 0
  End If

End Sub

Private Sub RefreshUserSettings()

  Dim I       As Long
  Dim arrTest As Variant

  arrTest = Array(LayOutAtTypeEOLComment, SecondryExitFix, ConvertFunction2Sub, CodeTypeSuffixStrip, VeryBigProc_CXX, _
                  CommentOutClassmembers, FormatRoutineLineContinuation)
  'ver 2.0.2 Thanks to Ing. Miguel Angel Guzman Robles
  'v2.2.6 Oops; This was the rest of the bug with new settings. Thanks Mike Ulik
  'v2.5.5 Thanks Mike Ulik this was the apply deteriorating settings error too!
  SettingCollector
  UserSettings = vbNullString
  For I = 0 To DummyFixEnd - 1
    UserSettings = UserSettings & FixData(I).FixLevel
    If IsInArray(I, arrTest) Then
      UserSettings = UserSettings & "|"
    End If
  Next I

End Sub

Public Sub RefreshUserSettingsFromString()

  Dim TSettings As String
  Dim I         As Long

  'ver 2.0.2 Thanks to Ing. Miguel Angel Guzman Robles
  'prog wasn't saving user setting properly
  'UserSettings = ""
  TSettings = Replace$(UserSettings, "|", vbNullString)
  For I = 0 To DummyFixEnd - 1
    FixData(I).FixLevel = Mid$(TSettings, I + 1, 1)
  Next I

End Sub

Public Sub SaveUserSettings()

  Dim TurnOn As Long

  Xcheck.SaveCheck
  RefreshUserSettings
  SaveSetting AppDetails, "Options", "UserSet", UserSettings
  SaveSetting AppDetails, "Options", "BigProc", lngBigProcLines
  SaveSetting AppDetails, "Options", "LngLine", lngLineLength
  SaveSetting AppDetails, "Options", "ElseIndnt", lngElseIndent
  SaveSetting AppDetails, "Options", "CaseIndnt", lngCaseIndent
  Select Case UserSettings
    'Case DefNoSettings
    'TurnOn = 0
   Case DefMinSettings
    TurnOn = 1
   Case DefAvgSettings
    TurnOn = 2
   Case DefMaxSettings
    TurnOn = 3
   Case Else
    TurnOn = 4
  End Select
  With frm_FindSettings
    .optLevelSetting(TurnOn).Value = True
  End With
  StandardSettings TurnOn

End Sub

Public Sub SettingCollector()


  With frm_CodeFixer
    'Declarations Section Fixers
    '              List,  ListMember ,Fixtype
    GetFixSetting2 0, 1, LargeSourceWarning_CXX
    GetFixSetting2 0, 2, RemoveIndent
    GetFixSetting2 0, 3, RemoveCodeFixComments
    GetFixSetting2 0, 4, RemoveLineContinuation
    GetFixSetting2 0, 5, XPStyleTest
    GetFixSetting2 0, 6, MoveCommentInside
    GetFixSetting2 0, 7, InsertOptExp
    GetFixSetting2 0, 8, MoveAPIDown
    GetFixSetting2 0, 9, DimGlobal2PublicPrivate
    GetFixSetting2 0, 10, Public2Private
    GetFixSetting2 0, 11, UpdateConstExplicit
    GetFixSetting2 0, 12, ExpandDecSingleLineSingleType
    GetFixSetting2 0, 13, ExpandDecSingleLine
    GetFixSetting2 0, 14, UpdateDecTypeSuffix
    GetFixSetting2 0, 15, UpdateDefType2AsType
    GetFixSetting2 0, 16, CreateEnumCapProtect
    GetFixSetting2 0, 17, BadNameVariableWarning_CXX
    GetFixSetting2 0, 18, LayOutAtTypeEOLComment
    'ReStructure Section Fixers
    GetFixSetting2 1, 1, UCTimerDisable
    GetFixSetting2 1, 2, UCFontFix
    GetFixSetting2 1, 3, RemoveLineNum
    GetFixSetting2 1, 4, UpDateStrConcat
    GetFixSetting2 1, 5, UpdateChrToConstant
    GetFixSetting2 1, 6, UpdateStrFunction
    GetFixSetting2 1, 7, UpdateErrError
    GetFixSetting2 1, 8, UpdateInteger2Long
    GetFixSetting2 1, 9, UnNeededExit
    GetFixSetting2 1, 10, NIfThenExpand
    GetFixSetting2 1, 11, SeperateCompounds
    GetFixSetting2 1, 12, UpdateWend
    GetFixSetting2 1, 13, NPleonasmFix
    GetFixSetting2 1, 14, NCloseResumeLoad
    GetFixSetting2 1, 15, NCloseResume
    GetFixSetting2 1, 16, WithPurity
    GetFixSetting2 1, 17, DetectShortCircuit
    GetFixSetting2 1, 18, DetectZeroLengthStringTests
    GetFixSetting2 1, 19, DetectZeroLengthStringAssign
    GetFixSetting2 1, 20, DetectWithStructure
    GetFixSetting2 1, 21, RoutineCaseFix
    GetFixSetting2 1, 22, PublicVar2Property
    GetFixSetting2 1, 23, UnNeededIfBracket
    GetFixSetting2 1, 24, UnNeededAmpersand
    GetFixSetting2 1, 25, UnNeededCall
    GetFixSetting2 1, 26, UneededCaseIs
    GetFixSetting2 1, 27, CollapseIfThenBool
    GetFixSetting2 1, 28, CollapseCaseBool
    GetFixSetting2 1, 29, SecondryExitFix
    'Parameter Section Fixers
    GetFixSetting2 2, 1, ParameterExpandTypeSuffix
    GetFixSetting2 2, 2, ParameterNoType
    GetFixSetting2 2, 3, ParameterBadName_CXX
    GetFixSetting2 2, 4, ParameterDead_CXX
    GetFixSetting2 2, 5, ParameterByVal
    GetFixSetting2 2, 6, FunctionTypeCast
    GetFixSetting2 2, 7, ConvertFunction2Sub
    'Local Section Fixers
    GetFixSetting2 3, 1, MoveDim2Top
    GetFixSetting2 3, 2, ReWriteDimMultiSingleType
    GetFixSetting2 3, 3, ReWriteDimExpandMultiDim
    'GetFixSetting2 3, 4, UpdateCtrlFrmRef
    'GetFixSetting2 3, 5, UpdateCtrlDefProp
    GetFixSetting2 3, 4, UpdateDimType
    GetFixSetting2 3, 5, DetectReWriteDimUntyped
    GetFixSetting2 3, 6, DetectDimUnused
    GetFixSetting2 3, 7, DetectDimMissing
    GetFixSetting2 3, 8, DetectDimDuplicate
    GetFixSetting2 3, 9, NoDimInitialise
    GetFixSetting2 3, 10, LayOutDim
    GetFixSetting2 3, 11, CodeTypeSuffixStrip
    'Suggestion Section Fixers
    GetFixSetting2 4, 1, DetectObsoleteCodeStructure_CXX
    GetFixSetting2 4, 2, DetectDangerousCode_CXX
    GetFixSetting2 4, 3, DetectDangerousAPICode_CXX
    GetFixSetting2 4, 4, DetectDangerousString_CXX
    GetFixSetting2 4, 5, DetectDangerousReference_CXX
    GetFixSetting2 4, 6, DetectStatic2Private_CXX
    GetFixSetting2 4, 7, DetectHardPath_CXX
    GetFixSetting2 4, 8, DetectBadCtrlName_CXX
    GetFixSetting2 4, 9, DuplicatePublicProc_CXX
    GetFixSetting2 4, 10, DetectIllegalWithExit_CXX
    GetFixSetting2 4, 11, DoElseIfEval_CXX
    GetFixSetting2 4, 12, GoToComments_CXX
    GetFixSetting2 4, 13, VeryBigProc_CXX
    'Unused Section Fixers
    GetFixSetting2 5, 1, EmptyRoutine_CXX
    GetFixSetting2 5, 2, REmptyStruct
    GetFixSetting2 5, 3, UnusedDecConst
    GetFixSetting2 5, 4, UnusedDecAPI
    GetFixSetting2 5, 5, UnusedDecVariable
    GetFixSetting2 5, 6, DeletedControlCode
    GetFixSetting2 5, 7, UnusedFunction
    GetFixSetting2 5, 8, UnusedSub
    GetFixSetting2 5, 9, UnusedProperty
    GetFixSetting2 5, 10, UnusedEvents
    GetFixSetting2 5, 11, UnusedWithEvents
    GetFixSetting2 5, 12, ActiveDebugStop
    GetFixSetting2 5, 13, CommentOutClassmembers
    'Format Section Fixers
    GetFixSetting2 6, 1, NForNextVar
    GetFixSetting2 6, 2, FormatCodeFix
    GetFixSetting2 6, 3, FormatArrayLineContinuation
    GetFixSetting2 6, 4, FormatStringLineContinuation
    GetFixSetting2 6, 5, FormatDeclareLineContinuation
    GetFixSetting2 6, 6, FormatRoutineLineContinuation
    GetFixSetting2 6, 7, NSortModules
    'Next I
  End With

End Sub

Private Sub SettingFrameCleanUp()

  Dim I               As Long

  With frm_FindSettings
    'Clean up the frames for presentation
    .picCFXPBugFixSettings(13).BorderStyle = 0
    .tbsSettings.Move 0, 0
    For I = 0 To .fraSubSetting.Count - 1
      .fraSubSetting(I).Caption = vbNullString
      .fraSubSetting(I).BorderStyle = 0
      'v2.4.1 Thanks Paul Caton this was hiding the active page of settings
      '.fraSubSetting(I).Visible = False
      .fraSubSetting(I).Visible = .tbsSettings.SelectedItem.Index - 1 = I
    Next I
  End With

End Sub

Private Sub SettingFrameFormat2(ByVal LsvNo As Long)

  'Mar 04
  'set tab to show whole listview

  With frm_FindSettings.ctlSettingsList(LsvNo)
    .Top = offset / 2
    .Left = offset
    frm_FindSettings.fraSubSetting(LsvNo + 1).Height = .Height - offset
    frm_FindSettings.fraSubSetting(LsvNo + 1).Width = .Width + offset
  End With

End Sub

Private Sub SettingFrameSize(ByVal FramSep As Long)

  Dim I               As Long
  Dim HighestSubFrame As Long
  Dim WidestSubFrame  As Long

  For I = 0 To 6
    SettingFrameFormat2 I
  Next I
  With frm_FindSettings
    For I = 0 To 6
      'find highest and widest Sub Setting Frame
      If .fraSubSetting(I).Height > HighestSubFrame Then
        HighestSubFrame = .fraSubSetting(I).Height
      End If
      If .fraSubSetting(I).Width > WidestSubFrame Then
        WidestSubFrame = .fraSubSetting(I).Width
      End If
    Next I
    'Size tab to hold frame
    .tbsSettings.Height = HighestSubFrame + FramSep * 3.5
    .tbsSettings.Width = WidestSubFrame + FramSep * 1.5
    'Size back frame to hold TabStrip
    .fraprop(1).BorderStyle = 0
    .fraprop(1).Height = .tbsSettings.Height
    .fraprop(1).Width = .tbsSettings.Width + FramSep
    'Fit backing PictureBox to Frame
    PictureToFrame .picCFXPBugFixSettings(13), .fraprop(1)
  End With

End Sub

Public Sub SettingsEngine(LsvNo As Long, _
                          strData As String)

  Dim ArrStart As Variant

  Dim arrEnd   As Variant
  'v2.2.6 Update following when adding/removing fixes Note added so I update properly. Thanks Mike Ulik
  ArrStart = Array(LargeSourceWarning_CXX, NSortModules, ParameterExpandTypeSuffix, MoveDim2Top, DetectObsoleteCodeStructure_CXX, _
                   EmptyRoutine_CXX, NForNextVar)
  arrEnd = Array(LayOutAtTypeEOLComment, SecondryExitFix, ConvertFunction2Sub, CodeTypeSuffixStrip, VeryBigProc_CXX, _
                 CommentOutClassmembers, FormatRoutineLineContinuation)
  UserSet2 LsvNo, CStr(Split(strData, "|")(LsvNo)), CLng(ArrStart(LsvNo)), CLng(arrEnd(LsvNo))

End Sub

Public Sub SetUpSettingFrame()

  
  Dim I            As Long
  Dim safeSettings As String

  'Rewritten to use a second tabcontrol
  'Thanks to Terry L <t4feb@comcast.net> who pointed out that this screen was too large on lo-res screens
  'Rewritten to use a second TabStrip
  'Clean up the frames for presentation
  'v2.5.3 rewite to use PRopertList UC
  safeSettings = UserSettings
  'protect real values from the effects of setting up the control
  SettingFrameCleanUp
  ' limitstr: | 1= Open 0=Blocked X=Default
  With frm_FindSettings.ctlSettingsList(0)
    InitalisePropertyList 0
    'v2.5.5 Thanks Mike Ulik this was the typo error. "100X"
    .AddProperty "1X00", "Large Source File Warning", "Add a comment to top of code if file ove 56K"
    .AddProperty "000X", "Remove all indenting", "Necessary for processing."
    .AddProperty "000X", "Remove Code Fixer Comments", "Delete old Code Fixer comments"
    .AddProperty "000X", "Remove Line Continuation", "Simplify Code Scanning"
    .AddProperty "0X00", "XP Style Frame Problem", "Add a comment to top of code if XP Frame bug detected"
    .AddProperty "X001", "Move Comments outside routine inside", "Layout style issue"
    .AddProperty "01X0", "Insert Option Explicit", "Makes coding safer"
    .AddProperty "100X", "Move API calls to bottom of Declarations", "Layout style issue"
    .AddProperty "01X1", "Dim, Global Declarations to Public/Private", "Replace Obsolete Module level Scope key words"
    .AddProperty "01X1", "Reduce Unneeded Public Dec to Private", "Reduce Public to Private to save memory"
    .AddProperty "11X1", "Rewrite Constant Declares to Explicit Types", "Un-Typed Constants are Varaint & wasteful of memory"
    .AddProperty "01X1", "Correct Multiple Declare Single Line Single Type", "'Dim X, Y As Long' is a Long (Y) and a Variant (X). Fix assumes you really wanted 2 Longs."
    .AddProperty "100X", "Expand Multiple Declare Single Line to Separate lines", "Increases readability of code"
    .AddProperty "100X", "Updates Type suffix Declares to 'As Type' format", "'As Type' is easier to read than !#%&"
    .AddProperty "100X", "DefType to 'As Type' ", "Replaces DefType with easier to read As Type"
    .AddProperty "100X", "Enum Capitalisation Protection", "A trick to preserve case when typing Enums in code"
    .AddProperty "10X0", "Poorly named Variable Warning", "Variables with same name as controls/properties or some VB comands are legal but make code hard to read"
    .AddProperty "100X", "Format Declarations(line up 'As Type' & EoL comments)", "Increases readability"
    .DrawPropertyBox
  End With
  With frm_FindSettings.ctlSettingsList(1)
    InitalisePropertyList 1
    .AddProperty "100X", "Disable UserControl Timers", "Stop Enabled Timers on UserControls (inserts code to activate them when code runs)"
    .AddProperty "100X", "UserControl Font fix", "If a UserControl stores/sets a Font Property then it should do so before any other Ppoperty is set."
    .AddProperty "100X", "Line Number Removal [GoTo targets protected]"
    .AddProperty "100X", "Update String Concatenation ' + ' => ' & '", "Removes a potential bug"
    .AddProperty "100X", "Update some Chr$() to VB Constants"
    .AddProperty "100X", "Modify VB String functions (Ulli)", "Some VB functions have String and Variant versions, this converts Variants to the Smaller, faster String version."
    .AddProperty "10X0", "Update old style Err to new Err Object", ""
    .AddProperty "X010", "Update Integer to Long", "In 32-bit systems Long is more efficent than Integer. However some API Types contain required Integer members."
    .AddProperty "11X0", "Unneeded Exit Routine", "Explicitly exiting procedures at natural exit points waste time."
    .AddProperty "11X1", "Single Line 'If Then' Expand", "Clearer coding style, reveals subtle logic flaws."
    .AddProperty "11X1", "Expand Compound lines", "Compound ( colon sepatated) code is hard to read and can conceal logic flaws."
    .AddProperty "11X1", "Update While..Wend to Do While..Loop", "Wend is a back-compatibility command"
    .AddProperty "11X1", "Repair Pleonasm ('If <var> = True Then')", "Removes unnecessary double testing of Boolean truth."
    .AddProperty "11X1", "Protect Load/Unload with 'On Error Resume Next'", "Error Traps a potential crash point in code."
    .AddProperty "11X1", "Close 'On Error Resume' w/ 'On Error Goto 0'", "Prevent flow on error trapping by restricting Error traps to a per procedure level. May not be correct but usually is."
    .AddProperty "11X1", "Check With End With Purity", "Remove unneeded doubleing up of references in With structures"
    .AddProperty "11X0", "Short Circuit of 'If...And...Then'", "Save time by testing conditions one at a time rather than having to perform all test before branching "
    .AddProperty "11X0", "Detect Zero Length String", "String comparison 'If X = " & DQuote & DQuote & " Then' is less efficent than 'If LenB(X) Then'"
    .AddProperty "11X0", "Assign Zero Length String", "Replace X = " & DQuote & DQuote & " with X = vbNullString"
    .AddProperty "11X0", "Suggest/Make 'With ...End With'", "Convert code to use more efficent With Structures."
    .AddProperty "11X0", "Routine Case Fix", "Adjust case of hand-coded control events to match VB standard. 'form_load' > 'Form_Load'"
    .AddProperty "11X0", "Convert Public Class/Form Variable to Property. Does not apply to Implements Classes."
    .AddProperty "11X0", "Unneeded 'If (Brackets) Then'", "Brackets used to force Boolean testing are often unneeded"
    .AddProperty "10X0", "Unneeded Literal Strings and Ampersand", " joining Double quotes strings with Anpersands (&) is unnecessary unless being used to format very long lines. This fix removes them in short code lines."
    .AddProperty "11X0", "Unneeded 'Call'", "'Call' is a back-compatability command that is not required"
    .AddProperty "11X0", "Unneeded 'Case Is ='", "'Case Is' is only required if using > ,=>,,<=,< or <>"
    .AddProperty "11X0", "Collapse If Then Boolean settings", "Using If..Then to simulate Boolean testing is slow"
    .AddProperty "11X0", "Collapse Case Select T/F to IIf", "Using Select Case to simulate Boolean testing is slow"
    .AddProperty "11X0", "Exit Procedure re-code/suggetions", "Secondry Exits from procedures can make code difficult to read. This suggests ways to code around them and can fix some simpler ones."
    .DrawPropertyBox
  End With
  With frm_FindSettings.ctlSettingsList(2)
    InitalisePropertyList 2
    .AddProperty "11X0", "Update Type-suffix", "Type suffixes are back-compatible and can obscure coding"
    .AddProperty "11X0", "UnTyped Parameters", "Variants are wasteful of memory unless absolutely necessary"
    .AddProperty "11X0", "Poorly named Parameters", "Names which match VB terms/commands are usually legal but cn make code obscure"
    .AddProperty "1X00", "Unused Parameters", "Passing unused parameters wastes time"
    .AddProperty "11X0", "Insert ByVal", "Using ByVal on values which will not e changed simplifies VB's data handling"
    .AddProperty "11X0", "Update No-Return Functions to Subs", "VB has to porvide support code for the return even if you never use it."
    .AddProperty "11X0", "Untyped Functions", "Will return a Variant which will almost certainly be immediately coerced to some Type, wastes time/memory"
    .DrawPropertyBox
  End With
  With frm_FindSettings.ctlSettingsList(3)
    InitalisePropertyList 3
    .AddProperty "11X1", "Move Dims To Top", "Improved format, avoids bug of trying to use a variable before it is dimmed."
    .AddProperty "11X1", "Update Multi Dim with Single As Type", "'Dim X, Y As Long' is a Long(Y) and a Variant(X). Fix assumes you really wanted 2 Longs."
    .AddProperty "11X1", "Expand Multi Dims to one per line", "Easier reading"
    .AddProperty "11X1", "Update Type suffixed Dims", "Type suffixes are back-compatible and can obscure coding"
    '.AddProperty "11X1", "Fully reference Form.Controls in bas"
    ' .AddProperty "11X0", "Set Control Default Property"
    .AddProperty "10X1", "Type untyped Dims", "Variant Dims waste memory"
    .AddProperty "11X0", "Detect Unused Dims", "Waste of time setting up unused dims"
    .AddProperty "11X1", "Identify Missing Dims", "relying on VB to create variable on the fly is risky. Only needed if 'Option Explicit' had to be inserted. "
    .AddProperty "11X1", "Dim Duplicates Private/Public", "Dims with the same name as module level variables can introduce subtle bugs."
    .AddProperty "10X0", "Unneeded Dim initialization", "It is unnecessary to initialise Dim variables to their default values."
    .AddProperty "100X", "Format Dims(line up 'As Type' & EoL comments) ", "Layout "
    .AddProperty "100X", "Remove Type Suffixes from code", "Using Type suffixes in code is unnecessary and can block the Integer to Long fix."
    .DrawPropertyBox
  End With
  '
  With frm_FindSettings.ctlSettingsList(4)
    InitalisePropertyList 4
    .AddProperty "1X00", "Obsolete/Archaic Code", "Mark back-compatible code "
    .AddProperty "1X00", "Risky commands", "Warn of dangerous commands"
    .AddProperty "1X00", "Risky API", "Warn of dangerous API"
    .AddProperty "1X00", "Risky String", "Warn of dangerous Strings which may contain instructions for scripting/API."
    .AddProperty "1X00", "Risky Reference", "Warn of dangerous external dlls"
    .AddProperty "1X00", "Static [Suggest Private]", "Static is about 3 times as memory intensive as module level variables"
    .AddProperty "1X00", "Hard-coded Path [Suggest App.Path]", "Not all systems have the same paths, App.Path is the only path that MUST exist on a machine."
    .AddProperty "1X00", "Poorly named Controls"
    .AddProperty "1X00", "Duplicate Public procedures warning", "Procedures with the same name are either complete duplicates or make code hard to read."
    .AddProperty "1X00", "Detect Dangerous With structure Exits", "Using Exit or GoTo inside a With structure is a potentail source of memory leaks."
    .AddProperty "1X00", "ElseIf Evaluation"
    .AddProperty "1X00", "Evaluate GoTo"
    .AddProperty "1X00", "Large Procedures", "Add warning; large procedures are more likely to contain poor coding as i is harder to keep the whole procedure in mind."
    .DrawPropertyBox
  End With
  '
  With frm_FindSettings.ctlSettingsList(5)
    InitalisePropertyList 5
    .AddProperty "1X00", "Empty Routine (<Stub> code)", "UserControls and some sub-classing code need empty procedures and this fix provides a default comment to mark these, all others can be removed."
    .AddProperty "1X00", "Detect Empty Structures", "Empty With/For/Case/If/ElseIf structures are usually a waste of space (unless being used to avoid a default 'Else' or 'Case Else')"
    .AddProperty "1X10", "Constants", ""
    .AddProperty "1X10", "API delcarations", ""
    .AddProperty "1X10", "Public/Private Variables", ""
    .AddProperty "1X10", "Deleted Control Code", ""
    .AddProperty "1X10", "Functions", ""
    .AddProperty "1X10", "Subs", ""
    .AddProperty "1X10", "Properties", ""
    .AddProperty "1X10", "Events", ""
    .AddProperty "1X10", "WithEvents", ""
    .AddProperty "1X10", "Active Debug/Stop", "While Debug is not an issue (compiler will ignore it) extensive code with the sole purpose of producing Debug data will waste time. Stop (unless carefully isolated) should neve occur in compiled code."
    .AddProperty "X010", "Comment Out Unused Class Members", "If you are using a library class you may not wish to comment out class members. It is advisable to create a copy of such classes and develop program specific version of the class."
    .DrawPropertyBox
  End With
  '
  With frm_FindSettings.ctlSettingsList(6)
    InitalisePropertyList 6
    .AddProperty "11X1", "Insert 'For..Next' terminal Variable", ""
    .AddProperty "000X", "Format Code", ""
    .AddProperty "100X", "Array Line Continuation", ""
    .AddProperty "100X", "String Line Continuation", ""
    .AddProperty "100X", "Declare Line Continuation", ""
    .AddProperty "100X", "Routine Parameter Line Continuation", ""
    .AddProperty "100X", "Sort Modules", "Alphabetic sort of Procedure names"
    .DrawPropertyBox
  End With
  SettingFrameSize 100 '120
  UserSettings = safeSettings
  For I = 0 To 6
    SettingsEngine I, UserSettings
  Next I

End Sub

Public Sub StandardSettings(ByVal intIndex As Long)

  Dim strTmp As String
  Dim I      As Long

  If Not bWorking Then
    bWorking = True
    frm_FindSettings.optLevelSetting(intIndex).Value = True
    Select Case intIndex
     Case LevelOff
      strTmp = DefNoSettings
     Case LevelComment
      strTmp = DefMinSettings
     Case LevelCommentFix
      strTmp = DefAvgSettings
     Case LevelFix
      strTmp = DefMaxSettings
     Case LevelUser
      strTmp = UserSettings
    End Select
    For I = 0 To 6
      SettingsEngine I, strTmp
    Next I
    bWorking = False
  End If

End Sub

Public Sub UserSet2(ByVal LsvNo As Long, _
                    ByVal StrSetting As String, _
                    ByVal Lnum As Long, _
                    ByVal Unum As Long)

  Dim I    As Long

  'Write UserSetting String into the settings listviews
  For I = Lnum To Unum
    FixData(I).FixLevel = Val(Mid$(StrSetting, I - Lnum + 1, 1))
  Next I
  If Len(frm_FindSettings.ctlSettingsList(LsvNo).DataString) Then
    frm_FindSettings.ctlSettingsList(LsvNo).DataString = StrSetting
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:20:55 AM) 158 + 503 = 661 Lines Thanks Ulli for inspiration and lots of code.

