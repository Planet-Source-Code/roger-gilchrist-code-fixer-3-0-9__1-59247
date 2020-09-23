Attribute VB_Name = "mod_XPStyle"
Option Explicit
'© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
'with extensive modifications
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
'these routines were extracted from Ulli's Code Formatter
'I just like modularization
'Win XP Look
Private Const XPLookAPIProto      As String = "Private Declare Sub InitCommonControls Lib ""comctl32"" ()"
'standard manifest file
Private Const XPLookXML           As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbNewLine & _
                                              "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbNewLine & _
                                              "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""$"" type=""win32"" />" & vbNewLine & _
                                              "<description>XP-Look - Created by Code Fixer</description>" & vbNewLine & _
                                              "<dependency>" & vbNewLine & _
                                              "<dependentAssembly>" & vbNewLine & _
                                              "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*"" />" & vbNewLine & _
                                              "</dependentAssembly>" & vbNewLine & _
                                              "</dependency>" & vbNewLine & _
                                              "</assembly>"





Private Declare Function WinVersion Lib "kernel32" Alias "GetVersion" () As Long

Private Function ComponentContainingProc(ByVal strProc As String) As VBComponent

  Dim Proj    As VBProject
  Dim Comp    As VBComponent
  Dim CompMod As CodeModule

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If LenB(Comp.Name) Then
        Set CompMod = Comp.CodeModule
        If CompMod.Find(strProc, 1, 1, -1, -1) Then
          Set ComponentContainingProc = Comp
          GoTo NormalExit
        End If
      End If
    Next Comp
  Next Proj
NormalExit:

End Function

Public Sub DoXPStyle()

  Dim SComp           As VBComponent
  Const strXPInsert   As String = RGSignature & "XP Style support inserted by Code Fixer"
  Const strXPDelete   As String = RGSignature & "XP Style removed by Code Fixer"
  Dim SCompMod        As CodeModule
  Dim StartLine       As Long
  Dim StrStartUpProc  As String
  Const XPLookAPICall As String = "InitCommonControls"

  StrStartUpProc = StartUpProcedure
  Set SComp = GetStartUpComponent
  Select Case HasXPStyle
   Case False
    If Not SComp Is Nothing Then
      Set SCompMod = SComp.CodeModule
      SCompMod.InsertLines SCompMod.CountOfDeclarationLines + 1, XPLookAPIProto & vbNewLine & strXPInsert
      If SCompMod.Find(StrStartUpProc, StartLine, 1, -1, -1, True) Then
        SafeInsertModule SCompMod, StartLine + 1, vbTab & XPLookAPICall & vbNewLine & _
         strXPInsert & vbNewLine
       Else
        SCompMod.InsertLines SCompMod.CountOfDeclarationLines + 1, vbNewLine & _
         "Private " & StrStartUpProc & vbNewLine & _
         strXPInsert & vbNewLine & _
         vbTab & XPLookAPICall & vbNewLine & _
         strXPInsert & vbNewLine & _
         vbNewLine & _
         "End Sub"
      End If
      'just a tidy up for neatness
      StartLine = 1
      Do While SCompMod.Find(strXPDelete, StartLine, 1, -1, -1, True)
        SCompMod.DeleteLines StartLine, 1
      Loop
      ManifestWrite
      DisplayCodePane SComp
     Else
      mObjDoc.Safe_MsgBox "Code contains no StartUp Object to insert XP Style support code." & vbNewLine & _
                    "Open Project|Properties menu and set a StartUp Object." & vbNewLine & _
                    "If you select 'Sub Main' you must create the Sub for Code Fixer to find it.", vbInformation
      'frm_CodeFixer.cmdXPStyle.Enabled = False
    End If
   Case True
    If Not SComp Is Nothing Then
      Set SCompMod = SComp.CodeModule
      '  SCompMod.InsertLines SCompMod.CountOfDeclarationLines + 1, XPLookAPIProto & vbNewLine & strXPInsert
      If SCompMod.Find(XPLookAPIProto, StartLine, 1, -1, -1, True) Then
        SCompMod.ReplaceLine StartLine, strXPDelete
      End If
      ' tidy up for neatness
      StartLine = 1
      Do While SCompMod.Find(XPLookAPICall, StartLine, 1, -1, -1, True)
        SCompMod.ReplaceLine StartLine, strXPDelete
      Loop
      StartLine = 1
      Do While SCompMod.Find(strXPInsert, StartLine, 1, -1, -1, True)
        SCompMod.DeleteLines StartLine, 1
      Loop
      ManifestDelete
      DisplayCodePane SComp
     Else
      mObjDoc.Safe_MsgBox "Code contains no XP Style support code", vbInformation
      frm_CodeFixer.cmdXPStyle.Enabled = False
    End If
  End Select

End Sub

Public Function GetStartUpComponent() As VBComponent

  'finds Sub Main or if missing the startup form and
  'find or create the Sub [MDI]Form_Initialize
  'for instertion point of InitCommonControls call.

  If Len(StartUpProcedure) Then
    Set GetStartUpComponent = ComponentContainingProc(StartUpProcedure)
    If GetStartUpComponent Is Nothing Then
      Set GetStartUpComponent = GetActiveProject.VBComponents.StartUpObject
      'no error - startup is a form
    End If
  End If

End Function

Public Function HasXPStyle() As Boolean

  Dim SComp As VBComponent

  Set SComp = GetStartUpComponent
  If Not SComp Is Nothing Then
    HasXPStyle = SComp.CodeModule.Find(XPLookAPIProto, 1, 1, -1, -1)
  End If

End Function

Public Sub ManifestDelete()

  On Error Resume Next
  Kill GetActiveProject.BuildFileName & ".Manifest"
  On Error GoTo 0

End Sub

Public Sub ManifestWrite()

  Dim EXEName    As String
  Const UserName As String = "generic"   'personalize here
  Dim FFile      As Long

  EXEName = GetActiveProject.BuildFileName
  'valid directory; create manifest file
  If LenB(EXEName) Then
    If LenB(Dir(Left$(EXEName, InStrRev(EXEName, "\")), vbDirectory)) Then
      FFile = FreeFile
      Open EXEName & ".Manifest" For Output As FFile
      Print #FFile, Replace$(XPLookXML, "$", UserName & "." & Replace$(Mid$(EXEName, InStrRev(EXEName, "\") + 1), ".exe", vbNullString, , , vbTextCompare))
      Close #FFile
    End If
  End If

End Sub

Public Function StartUpComponent() As String

  On Error Resume Next
  With GetActiveProject
    If .Type = vbext_pt_StandardExe Or .Type = vbext_pt_ActiveXExe Then
      StartUpComponent = .VBComponents.StartUpObject.Name
    End If
  End With
  On Error GoTo 0

End Function

Public Function StartUpProcedure() As String

  Dim Comp As VBComponent

  If Len(StartUpComponent) Then
    With GetActiveProject
      StartUpProcedure = "Sub Main"
      On Error Resume Next
      With .VBComponents.StartUpObject
        Set Comp = GetActiveProject.VBComponents.StartUpObject
        If Err.Number = 0 Then 'no error - startup is a form
          If .Type = vbext_ct_VBForm Then
            StartUpProcedure = "Sub Form_Initialize" 'so now we know
           Else
            StartUpProcedure = "Sub MDIForm_Initialize"
          End If
        End If
      End With
      On Error GoTo 0
    End With
  End If

End Function

Public Function WeAreRunningUnderWinXP() As Boolean

  WeAreRunningUnderWinXP = ((WinVersion And &HFF) >= 5) 'RequiredWinVersion)

End Function

Public Sub XPManifestFrameWarning()

  Dim Msg                As String
  Dim GotOne             As Boolean
  Dim bBadDrawnCtrl      As Boolean
  Dim I                  As Long
  Dim J                  As Long
  Dim K                  As Long
  Dim ArrSuspectControls As Variant
  Dim DecArray           As Variant
  Dim Comp               As VBComponent
  Dim Proj               As VBProject
  Dim CurCompCount       As Long

  'v2.1.3 updated to use the CntrlDesc data instead of checking controls directly (much faster)
  'This routine tests all forms for the XP & Frame problem
  'The following Controls don't draw properly on Frames
  'OptionButton:- Black ForeColor and BackColor (Cannot be set to anything else).
  'CommandButtons:- get a black border and moving mouse over the button causes all controls on the same frame to flicker.
  'Frames:- (within Frames) Caption is very large, truncated (approx 8 letters long depending on TextWidth value for characters) and Bold.
  'check data arrays to see if anything was tagged
  If Not bAborting Then
    If XPManifestFrameWarningNeeded Then
      For I = LBound(CntrlDesc) To UBound(CntrlDesc)
        If CntrlDesc(I).CDXPFrameBug Then
          Exit For
        End If
      Next I
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If dofix(CurCompCount, XPStyleTest) Then
            If IsComponent_ControlHolder(Comp) Then
              ModuleMessage Comp, CurCompCount
              DecArray = GetDeclarationArray(Comp.CodeModule)
              DisplayCodePane Comp, True
              If Xcheck(XVisScan) Then
                Comp.Activate
              End If
              For I = LBound(CntrlDesc) To UBound(CntrlDesc)
                With CntrlDesc(I)
                  If .CDXPFrameBug Then
                    If .CDForm = Comp.Name Then
                      If Len(.CDContains) Then
                        ArrSuspectControls = Split(.CDContains, ",")
                        For J = LBound(ArrSuspectControls) To UBound(ArrSuspectControls)
                          For K = LBound(CntrlDesc) To UBound(CntrlDesc)
                            If CntrlDesc(K).CDFullName = ArrSuspectControls(J) Then
                              If CntrlDesc(K).CDForm = .CDForm Then
                                Select Case CntrlDesc(K).CDClass
                                 Case "CommandButton", "OptionButton"
                                  If CntrlDesc(K).CDStyle = 0 Then
                                    'Graphic option/command-buttons (= 1) are immune to the XP Frame bug
                                    bBadDrawnCtrl = True
                                    'BadDrawnCtrl = AccumulatorString(BadDrawnCtrl, CntrlDesc(K).CDFullName & SngSpace & IIf(Len(CntrlDesc(K).CDCaption), " Captioned " & DQuote & CntrlDesc(K).CDCaption & DQuote, "[UnCaptioned]"), vbNewLine & "*")
                                  End If
                                 Case "Frame"
                                  If Len(CntrlDesc(K).CDCaption) > 0 Then
                                    'captionless frames dont need it either
                                    bBadDrawnCtrl = True
                                    'BadDrawnCtrl = AccumulatorString(BadDrawnCtrl, CntrlDesc(K).CDFullName & SngSpace & IIf(Len(CntrlDesc(K).CDCaption), " Captioned " & DQuote & CntrlDesc(K).CDCaption & DQuote, "[UnCaptioned]"), vbNewLine & "*")
                                  End If
                                 Case "Label"
                                  'a label may be captioned by code
                                  bBadDrawnCtrl = True
                                End Select
                              End If
                            End If
                          Next K
                        Next J
                        If bBadDrawnCtrl Then
                          'multiples = LBound(ArrSuspectControls) > 0
                          GotOne = True
                          bBadDrawnCtrl = False
                        End If
                      End If
                    End If
                  End If
                End With 'CntrlDesc(I)
              Next I
              If GotOne Then 'LenB(Msg) Then
                Msg = RGSignature & "POTENTIAL XP FRAME BUG DETECTED" & vbNewLine & _
                 " CommandButton, OptionBox, large Labels and captioned Frame (within another Frame)" & vbNewLine & _
                 " Do not draw properly if used in XP-Style (a Manifest for the program or VB itself will produce this bug)." & vbNewLine & _
                 "SOLUTION: Use the XP-Frame Bug Button on Controls tool."
                'DecArray may not be an Array.
                If UBound(DecArray) > -1 Then
                  DecArray(0) = Marker(DecArray(0), Msg, MBefore)
                 Else
                  'This copes with that
                  'Also inserts option explicit as this
                  'is tested before option explicit inserter
                  If Not bNoOptExp Then 'v3.0.0 Block fix in Format only thanks Ian K
                    'v3.0.4 prevent crash caused by UC not yet having missing Dims fixed
                    If Comp.Type <> vbext_ct_UserControl Then
DecArray = Array(Marker(IIf(Xcheck(XVerbose), OptExplicitMsgVerbose, OptExplicitMsg) & vbNewLine & "Option Explicit", Msg, MAfter))
                      AddNfix InsertOptExp
                    End If
                  End If
                End If
                ReWriter Comp.CodeModule, DecArray, RWDeclaration
                GotOne = False
                Msg = vbNullString
              End If
            End If
SkipComp:
          End If
        Next Comp
        If bAborting Then
          Exit For
        End If
      Next Proj
    End If
  End If

End Sub

Private Function XPManifestFrameWarningNeeded() As Boolean

  Dim I                 As Long

  'check data arrays to see if anything was tagged
  If bCtrlDescExists Then
    For I = LBound(CntrlDesc) To UBound(CntrlDesc)
      If CntrlDesc(I).CDXPFrameBug Then
        XPManifestFrameWarningNeeded = True
        Exit For
      End If
    Next I
  End If

End Function

Public Function XPStyleCaption() As String

  XPStyleCaption = IIf(HasXPStyle, "Remove XP Style", "Apply XP Style")

End Function


':)Code Fixer V3.0.9 (25/03/2005 4:24:45 AM) 17 + 317 = 334 Lines Thanks Ulli for inspiration and lots of code.

