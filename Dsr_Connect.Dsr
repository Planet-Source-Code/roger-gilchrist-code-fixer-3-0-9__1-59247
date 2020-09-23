VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Dsr_Connect 
   ClientHeight    =   13785
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13890
   _ExtentX        =   24500
   _ExtentY        =   24315
   _Version        =   393216
   Description     =   "Fix and Format code"
   DisplayName     =   "Code Fixer 3"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Dsr_Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'The sections of this code to do with docking is
'based on code from Jason A. Fisher fisher@scapromo.com
'used in his CodeBrowser program
'I used this to avoid learning UserDocuments from the bottom up.
'Events handlers
'Private WithEvents mobjCBEvts      As CommandBarEvents
'(I don't use these but have left them in for later experimental purposes
'Private WithEvents mobjPrjEvts      As VBProjectsEvents
'Private WithEvents mobjCmpEvts      As VBComponentsEvents
'Private WithEvents mobjFCEvts       As FileControlEvents
'Private mobjMCBCtl                 As CommandBarControl
'Module-level extensibility objects
'see notes below for what this is about
Public mWindow                                      As Window
'Dockable add-in needs a GUID
'You MUST generate your own values for the guid string constant otherwise you may clash with other tools
'Using a tool called Guidgen.exe, which is located in the \tools\idgen directory of Visual Basic.
'not installed automatically so time to dig out the disk.
'this supports the USerDocument way of addressing the COmboBoxes
'I put it here to over come the problem of Optional Compilation (#End If etc)
'not being detected properly if it is the last statement in a Declaration section
Public Enum CBoxID
  SearchB
  ReplaceB
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private SearchB, ReplaceB
#End If
Private WithEvents MenuEvents                       As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1
Private MenuItem                                    As CommandBarControl
Private Const mstrGuid                              As String = "{981D8E33-6A0A-4609-8DE6-8F70E1A5CEA1}"
Private mcbMenuCommandBar                           As Office.CommandBarControl
'command bar object
Public WithEvents MenuHandler                       As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
'command bar event handler
'
Public WithEvents MenuHandlerMyRightCode            As CommandBarEvents
Attribute MenuHandlerMyRightCode.VB_VarHelpID = -1
Private WithEvents MenuHandlerMyRightControl        As CommandBarEvents
Attribute MenuHandlerMyRightControl.VB_VarHelpID = -1
Private WithEvents MenuHandlerMyRightForm           As CommandBarEvents
Attribute MenuHandlerMyRightForm.VB_VarHelpID = -1
' Gain access to the events by referencing the object.
' 'components event handler
Public WithEvents CtrlHandler                       As VBControlsEvents
Attribute CtrlHandler.VB_VarHelpID = -1
'controls event handler
Public WithEvents PrjHandler                        As VBProjectsEvents
Attribute PrjHandler.VB_VarHelpID = -1
'    'projects event handler
Public WithEvents CmpHandler                        As VBComponentsEvents
Attribute CmpHandler.VB_VarHelpID = -1

Private Sub AddinInstance_Initialize()

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub AddinInstance_OnAddinsUpdate(custom() As Variant)

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)

  Dim CommandBarMenu As CommandBar
  Dim I              As Long

  bForceReload = False
  bControlsSet = True
  BadUserControlTestConducted = False
  'This event fires when the user loads the add-in from the Add-In Manager
  On Error GoTo ErrorHandler
  'Retain handle to the current instance of Visual Basic for later use
  Set VBInstance = Application
  If ConnectMode = ext_cm_External Then
    FormatAll
   Else
    If AddInMenuAvailable Then
      On Error Resume Next
      Set CommandBarMenu = VBInstance.CommandBars("Add-Ins")
      On Error GoTo 0
      If CommandBarMenu Is Nothing Then
        MsgBox "Code Fixer was loaded but could not be connected to the 'Add-Ins' menu.", vbCritical
       Else
        DoEvents
        With CommandBarMenu
          Set MenuItem = .Controls.Add(msoControlButton) '
          I = .Controls.Count - 1
          If .Controls(I).BeginGroup Then
            If Not .Controls(I - 1).BeginGroup Then
              'menu separator required
              MenuItem.BeginGroup = True
            End If
          End If
        End With
        'set menu caption
        '[Alt]+[A]+[R]'
        'The "&" in next line sets the [R] in the hotkey you may need to change this for your language/font
        MenuItem.Caption = "&" & AppDetails & "..."
        On Error Resume Next
        bInitializing = True
        PasteAFace MenuItem
        InitializeTLib
        Xcheck.Init frm_FindSettings.chkUser, AppDetails, "Options", "AllChecks", , 2495
        ReDim FixData(DummyFixEnd - 1) As NFix
        bInitializing = False
        'set event handler
        With VBInstance
          Set MenuEvents = .Events.CommandBarEvents(MenuItem)
          Set Me.CtrlHandler = .Events.VBControlsEvents(Nothing, Nothing)
          Set Me.PrjHandler = .Events.VBProjectsEvents
          Set Me.CmpHandler = .Events.VBComponentsEvents(Nothing)
          'done connecting
        End With 'VBInstance
      End If
    End If
  End If
  'Hook project and components events
  '''  With VBInstance
  '''    Set mobjPrjEvts = .Events.VBProjectsEvents
  '''    Set mobjCmpEvts = .Events.VBComponentsEvents(Nothing)
  '''    Set mobjFCEvts = .Events.FileControlEvents(Nothing)
  '''  End With
  'Convert the ActiveX document into a dockable tool window in the VB IDE
  'Uses the CreateToolWindow function
  'VB help doesn't explain this line very clearly so here's my explanation.
  'Set to an object of type Window (this can be private to the Add-in designer, it's only purpose is to
  '              allow you to make the add-in visible either when the menu button is clicked or during VB startup (see below)
  '1st parameter comes form this routine's parameters (don't touch)
  '2nd parameter is project name (up at top of project window on right (usually) plus a '.' connector plus the name of your UserDocument (bottom of project window)
  '3rp parameter the name you want to appear on tool (I use a routine to keep the Ver number up to date)
  '4th parameter a Guid number (you must generate a new one for each program DO NOT just cut and paste the one in the help file
  '              Use a tool called Guidgen.exe, which is located in the \tools\idgen directory of Visual Basic CD.
  '5th parameter The name of your Userdocument (or a variable holding a reference to it which is declared  As <Type = UserDocument's name>
  '              I used 'Public mobjDoc           As docfind' which I placed in a bas Module so that it was public to other parts of code
  Set mWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "prj_CodeFix3.FindDoc", AppDetails, mstrGuid, mObjDoc)
  'Read Preferences from Registry (if they exist)
  mObjDoc.SettingsLoad
  MenusCreate
  SetUpSettingFrame
  LoadArrays
  'InitArrays
  DefaultGridSizes
  'This is what allows Find to act like part of VB and appear during VB's startup process
  'started from the addin toolbar
  If ConnectMode = vbext_cm_AfterStartup Then
    AddToCommandBar
  End If
  If bLaunchOnStart Then
    ToggleTool
  End If
ExitSub:

Exit Sub

ErrorHandler:
  MsgBox "An error has occured!" & vbNewLine & Err.Number & ": " & Err.Description, vbCritical
  Resume ExitSub

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)

  Dim frm As Form

  'This event fires when the user explicitly unloads the add-in
  On Error Resume Next
  'save properties
  'Persist Preferences to Registry
  mObjDoc.SettingsSave
  Xcheck.SaveCheck
  SaveSetting AppDetails, "Options", "UserSet", UserSettings '
  'SaveUserSettings
  'ver 2.0.3 a miilion thanks to necro static who finally found
  'this bug which made CF attack the VB menus
  'still not sure why not everyone got the problem
  '(or at least noticed it)
  'v2.8.5 not a bug for most but mucks up my toolbars during test runs.
  If Not RunningInIDE Then
    mcbMenuCommandBar.Delete
    'right-click menu removal
  End If
  MenuDestroyAll
  MenuItem.Delete
  Set MenuItem = Nothing
  'mobjMCBCtl.Delete
  'destroy the support forms
  'Added to v1.1.01 clean up properly
  bAddinTerminate = True
  For Each frm In Forms
    Unload frm
    Set frm = Nothing
  Next frm
  'Destroy the UserDocument
  'Unload mobjDoc
  Set mObjDoc = Nothing
  'Destroy the add-in window
  ' Unload mWindow
  Set mWindow = Nothing
  On Error GoTo 0

End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

  AddToCommandBar
  mObjDoc.ApplyChanges

End Sub

Private Sub AddinInstance_Terminate()

  bAddinTerminate = True

End Sub

Private Function AddInMenuAvailable() As Boolean

  AddInMenuAvailable = Not VBInstance.CommandBars("Add-Ins") Is Nothing
  If Not AddInMenuAvailable Then
    mObjDoc.Safe_MsgBox "'Add-Ins' Menu is unavailable.", vbCritical
  End If

End Function

Private Sub AddToCommandBar()

  'v2.0.6 add a CF button to Standard toolbar

  On Error GoTo AddToCommandBarErr
  If bToolBarButton Then
    'make sure the standard toolbar is visible
    VBInstance.CommandBars(2).Visible = True
    'add it to the command bar
    'the following line will add the TabOrder manager to the
    'Standard toolbar to the right of the ToolBox button
    Set mcbMenuCommandBar = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    'set the caption
    mcbMenuCommandBar.Caption = AppDetails
    'copy the icon to the clipboard
    Clipboard.SetData frm_CodeFixer.picMenuIcon.Image
    'set the icon for the button
    mcbMenuCommandBar.PasteFace
    'sink the event
    Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    'restore the last state
    Clipboard.Clear
  End If

Exit Sub

AddToCommandBarErr:
  MsgBox Err.Description

End Sub

Private Sub DoPopup(MyMnu As Menu)

  On Error Resume Next
  frm_RCMenus.PopupMenu MyMnu
  Unload frm_RCMenus 'needed otherwise the form blocks access
  On Error GoTo 0

End Sub

Private Sub MenuDestroy(Obj As Object)

  Obj.Delete
  Set Obj = Nothing

End Sub

Private Sub MenuDestroyAll()

  'Destroy menus in reverse order to creation
  ' MenuDestroy MyRightFormMenu

  MenuDestroy MyRightCodeMenu
  MenuDestroy MyRightControlsMenu

End Sub

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, _
                             handled As Boolean, _
                             CancelDefault As Boolean)

  ToggleTool

End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, _
                              handled As Boolean, _
                              CancelDefault As Boolean)

  'V3.0.3 Thanks Bazz these 2 were what killed the toolbar and context menu

  ToggleTool

End Sub

Private Sub MenuHandlerMyRightCode_Click(ByVal CommandBarControl As Object, _
                                         handled As Boolean, _
                                         CancelDefault As Boolean)

  'V3.0.3 Thanks Bazz these 2 were what killed the toolbar and context menu

  DoPopup frm_RCMenus.mnuRCCode

End Sub

Private Sub MenuHandlerMyRightControl_Click(ByVal CommandBarControl As Object, _
                                            handled As Boolean, _
                                            CancelDefault As Boolean)

  DoPopup frm_RCMenus.mnuRCControl

End Sub

Private Sub MenuItemPaint(CmdBar As CommandBar, _
                          MItem As CommandBarButton)

  Set MItem = CmdBar.Controls.Add(msoControlButton, , , , True)
  On Error Resume Next
  With MItem
    '   On Error Resume Next
    .BeginGroup = True
    .Caption = "Code Fixer"
  End With
  PasteAFace MItem
  On Error GoTo 0

End Sub

Private Sub MenusCreate()

  Dim MyEvents            As Events

  'these are the VB codePane and Designer right-click menus
  Set MyEvents = VBInstance.Events ' Event sink allows CF to respond to clicks on the controls
  '
  MenuItemPaint VBInstance.CommandBars("Code Window"), MyRightCodeMenu ' index 15
  Set MenuHandlerMyRightCode = MyEvents.CommandBarEvents(MyRightCodeMenu) 'index 21
  '
  MenuItemPaint VBInstance.CommandBars("Controls"), MyRightControlsMenu
  Set MenuHandlerMyRightControl = MyEvents.CommandBarEvents(MyRightControlsMenu)
  '
  '    MenuItemPaint VBInstance.CommandBars("Forms"), MyRightFormMenu 'index 20
  '    Set MenuHandlerMyRightForm = MyEvents.CommandBarEvents(MyRightFormMenu)

End Sub

Private Sub PasteAFace(ByVal Obj As Variant)

  Dim Tmpstring1          As String
  Dim PreserveClipGraphic As Variant

  'v2.2.0 apply the yellow hammer to commandbar buttons/menus etc
  With Clipboard
    'set menu picture
    Tmpstring1 = .GetText
    PreserveClipGraphic = .GetData
    .SetData frm_CodeFixer.picMenuIcon.Image
    Obj.PasteFace
    .Clear
    If IsObject(PreserveClipGraphic) Then
      .SetData PreserveClipGraphic
    End If
    If LenB(Tmpstring1) Then
      .SetText Tmpstring1
    End If
  End With

End Sub

Private Sub ToggleTool()

  'show/hide tool form menu or toolbar

  If mWindow.Visible Then
    mWindow.Visible = False
   Else
    mWindow.Visible = True
    'mObjDoc.Show
    mObjDoc.Startup
  End If

End Sub

''
''Private Sub MenuHandlerMyRightForm_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''  DoPopup frm_RCMenus.mnuRCForm
''End Sub
''
''Private Sub CtrlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
''If bCtrlDescExists Then
''bControlsSet = True
''End If
''End Sub
''
''Private Sub CtrlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
''If bCtrlDescExists Then
''bControlsSet = True
''End If
''End Sub
''
''Private Sub CtrlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
''If bCtrlDescExists Then
''bControlsSet = True
''End If
''End Sub
''

':)Code Fixer V3.0.9 (25/03/2005 4:11:33 AM) 47 + 365 = 412 Lines Thanks Ulli for inspiration and lots of code.

