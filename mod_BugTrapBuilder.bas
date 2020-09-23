Attribute VB_Name = "mod_BugTrapBuilder"
'<<<<<<<<<<<<<<<<<<<NEW >>>>>>>>>>>>>>>>>>>>>>>>>>
'This module is the start of adding Error Trapping to code.
' TO USE
'1 Place cursor in the Procedure you wish to Bug Trap
'2 Select the menu option.
'3  What happens next depends on the module you are in, whether it has non-BugTrap error detection
'   and whether BugTrap is already present.
' a. if in a Class or UserControl the routines 'BugTrapComment' and 'RunningInIDE' are added to the
'    code module and Scoped to Private . This maintains portablility.
' b. If in any other module then if this is the first time you have added BugTrap a new module
'    'mod_BugTrap(mod_BugTrap.bas)' is added to the program.
'    This supports BugTrap for all non Class/UserControl modules
' 4  CF checks that there is not already some form of Error Trapping in the routine
'    (Looks for 'On Error Goto XXX' and the presence of an 'Exit <ProcedureType>
'     somewhere in code above the 'XXX:' label)
' 5  CF Then checks for existing BugTrap code. If it already exists CF will remove it.
' 6  If there is no BugTrap code or other Error detection CF inserts BugTrap
Option Explicit
Private Const TrapTop     As String = "On Error GoTo BugTrap"

Private Sub AddBugTrapSupportProcedures()

  Dim cMod        As VBIDE.CodeModule

  'Test for existance of BugTrap support procedures and inserts as needed
  ' if you are in a form or bas module CF adds another module containing the
  ' support routines
  ' if you are in a UserControl or Class CF adds the procedures to the module
  ' and sets the bug report name to match the modules name in order to maintain containment.
  Set cMod = GetActiveCodeModule
  With cMod
    If .Parent.Type = vbext_ct_ClassModule Or .Parent.Type = vbext_ct_UserControl Then
      If Not HasBugTrap(.Parent.Name, True) Then
        mObjDoc.Safe_MsgBox "Code Fixer is adding the necessary support procedures for BugTrap error detection to" & strInSQuotes(.Parent.Name, True) & "." & vbNewLine & _
                    "Adding further BugTraps in this module will not require this step.", vbInformation
        .InsertLines .CountOfLines + 1, StrBugTrapSub(True, .Parent.Name)
      End If
      If Not HasRunningInIDE(.Parent.Name, True) Then
        .InsertLines .CountOfLines + 1, StrCreateRunningInIDE(True)
      End If
     Else
      CreateBugTrapModule
    End If
  End With 'Cmod

End Sub

Private Function AlreadyTrapped(ByVal arrR As Variant) As Boolean

  Dim I As Long

  'Test for presense of BugTrap
  For I = LBound(arrR) To UBound(arrR)
    If SmartLeft(arrR(I), TrapTop) Then
      AlreadyTrapped = True
      Exit For 'unction
    End If
  Next I

End Function

Public Sub BugTrapAddRemoveToggle()

  Dim cMod         As CodeModule
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant

  'If a procedure has BugTrap delete it else add it
  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind)
  If PName = "(Declarations)" Then
    mObjDoc.Safe_MsgBox "You cannot add BugTrap error detection to Declaration sections." & vbNewLine & _
                    "Please place cursor in a procedure and try again", vbInformation
   Else
    AddBugTrapSupportProcedures
    If Not TestOtherDebug(ArrProc) Then
      If AlreadyTrapped(ArrProc) Then
        ArrProc = CleanArray(ArrProc)
        ' strip out blanks so that the proc conforms to BugTrap standard layout
        If BugTrapComplete(ArrProc, ProcKind, PName) Then
          BugTrapRemove ArrProc, ProcKind, PName
          ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
         Else
          mObjDoc.Safe_MsgBox strInSQuotes(PName) & " contains elements of BugTrap" & vbNewLine & _
                    "but they do not conform to the generic format." & vbNewLine & _
                    "You will have to remove it by hand.", vbInformation
        End If
       Else
        ArrProc(TopOfRoutine) = ArrProc(TopOfRoutine) & vbNewLine & TrapTop
        ArrProc(UBound(ArrProc)) = strBugTrap(ProcKind, PName) & ArrProc(UBound(ArrProc))
        ReplaceProcedureCode cMod, ArrProc, PSLine, PEndLine
      End If
    End If
  End If

End Sub

Private Function BugTrapComplete(arrR As Variant, _
                                 ByVal ProcKind As String, _
                                 ByVal PName As String) As Boolean

  Dim I As Long

  'test that bugtrap code is in standard format
  For I = LBound(arrR) To UBound(arrR)
    If arrR(I) = TrapTop Then
      BugTrapComplete = True
     ElseIf arrR(I) = "BugTrap:" Then
      If BugTrapComplete Then
        If Trim$(arrR(I - 1)) = "Exit " & ProcKind Then
          If Trim$(arrR(I + 1)) = "BugTrapComment " & DQuote & PName & DQuote Then
            If Trim$(arrR(I + 2)) = "If RunningInIDE Then" Then
              If Trim$(arrR(I + 3)) = "Resume" Then
                If Trim$(arrR(I + 4)) = "Else" Then
                  If Trim$(arrR(I + 5)) = "Resume Next" Then
                    If Trim$(arrR(I + 6)) = "End If" Then
                      BugTrapComplete = True
                      Exit For 'unction
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next I

End Function

Private Sub BugTrapRemove(arrR As Variant, _
                          ByVal strPKind As String, _
                          ByVal strPName As String)

  Dim I          As Long
  Dim J          As Long
  Dim TrapTopPos As Long

  'Delete bugtrap code that is in standard format
  For I = LBound(arrR) To UBound(arrR)
    If arrR(I) = TrapTop Then
      TrapTopPos = I
     ElseIf arrR(I) = "BugTrap:" Then
      If Trim$(arrR(I - 1)) = "Exit " & strPKind Then
        If Trim$(arrR(I + 1)) = "BugTrapComment " & DQuote & strPName & DQuote Then
          If Trim$(arrR(I + 2)) = "If RunningInIDE Then" Then
            If Trim$(arrR(I + 3)) = "Resume" Then
              If Trim$(arrR(I + 4)) = "Else" Then
                If Trim$(arrR(I + 5)) = "Resume Next" Then
                  If Trim$(arrR(I + 6)) = "End If" Then
                    For J = I - 1 To I + 6
                      arrR(J) = vbNullString
                    Next J
                    arrR(TrapTopPos) = vbNullString
                    Exit For
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next I
  arrR = CleanArray(arrR)

End Sub

Public Sub CreateBugTrapModule()

  Dim FFile    As Long
  Dim strCode  As String
  Dim FullPath As String
  Dim Proj     As VBProject

  'based on Sub Project_AddFile in 'Building Add-Ins for Visual Basic 4.0' MSDN help files
  Set Proj = GetActiveProject
  If Not HasBugTrap("", , True) Then
    mObjDoc.Safe_MsgBox "Code Fixer is adding a module 'mod_BugTrap(mod_BugTrap.bas)' to your code." & vbNewLine & _
                    "This module provides support for BugTrap error detection." & vbNewLine & _
                    "Adding Further BugTraps to Forms and .bas Modules will not require this step.", vbInformation
    Set Proj = GetActiveProject
    FullPath = GetProjFolder & "\mod_BugTrap.bas"
    strCode = "'BugTrap module Created by " & AppDetails & vbNewLine & _
              "Option Explicit" & StrBugTrapSub & StrCreateRunningInIDE
    FFile = FreeFile
    Open FullPath For Output As #FFile
    Print #FFile, strCode
    Close FFile
    Proj.VBComponents.AddFile FullPath  '.FileNames(1) '(FullPath$).Name = "mod_BugTrap"
    Proj.VBComponents(Proj.VBComponents.Count).Name = "mod_BugTrap"
  End If

End Sub

Private Function DetectNonBugTrapErrorHandling(arrR As Variant) As Boolean

  Dim I        As Long
  Dim J        As Long
  Dim K        As Long
  Dim strTest  As String
  Dim ProcKind As String
  Dim arrTest  As Variant

  arrTest = Array("0", "BugTrap")
  'test for non-BugTrap error handling (BugTrap might stuff this up so don't do it)
  ProcKind = GetProcClassStr(arrR(GetProcCodeLineOfRoutine(arrR)))
  For I = LBound(arrR) To UBound(arrR)
    'v2.9.0
    If IsOnErrorCode(arrR(I), 1) Then
      strTest = WordAfter(arrR(I), "GoTo")
      If Not IsInArray(strTest, arrTest) Then
        For J = I To UBound(arrR)
          If SmartLeft(Trim$(arrR(J)), strTest & ":") Then
            For K = J To I Step -1
              If SmartLeft(Trim$(arrR(K)), "Exit " & ProcKind) Then
                DetectNonBugTrapErrorHandling = True
                Exit For
              End If
            Next K
          End If
          If DetectNonBugTrapErrorHandling Then
            Exit For
          End If
        Next J
      End If
    End If
    If DetectNonBugTrapErrorHandling Then
      Exit For
    End If
  Next I

End Function

Private Function DetectResumeErrorHandling(arrR As Variant) As Boolean

  Dim I         As Long

  'test for Resume error handling (BugTrap might stuff this up so don't do it)
  For I = LBound(arrR) To UBound(arrR)
    'v2.9.0
    If IsOnErrorCode(arrR(I), 2) Then
      DetectResumeErrorHandling = True
      Exit For 'unction
    End If
  Next I

End Function

Private Function HasBugTrap(CompName As String, _
                            Optional bCurModuleOnly As Boolean = False, _
                            Optional bIngnoreClsUserD As Boolean = False) As Boolean

  'Test BugTrap main procedure exits

  HasBugTrap = FindCodeUsage("Sub BugTrapComment", vbNullString, CompName, False, bCurModuleOnly, False, , bIngnoreClsUserD)

End Function

Public Function HasBugTrap2(Optional ByVal bSilent As Boolean = False) As Boolean

  Dim cMod         As CodeModule
  Dim PSLine       As Long
  Dim PEndLine     As Long
  Dim TopOfRoutine As Long
  Dim Ptype        As Long
  Dim PName        As String
  Dim ProcKind     As String
  Dim ArrProc      As Variant

  Set cMod = GetActiveCodeModule
  ArrProc = ReadProcedureCodeArray(PSLine, PEndLine, PName, Ptype, TopOfRoutine, ProcKind)
  If PName <> "(Declarations)" Then
    '    AddBugTrapSupportProcedures
    If Not TestOtherDebug(ArrProc, bSilent) Then
      If AlreadyTrapped(ArrProc) Then
        HasBugTrap2 = True
      End If
    End If
  End If

End Function

Private Function HasRunningInIDE(CompName As String, _
                                 Optional bCurModuleOnly As Boolean = False, _
                                 Optional bIngnoreClsUserD As Boolean = False) As Boolean

  'Test BugTrap support procedure exits

  HasRunningInIDE = FindCodeUsage("Function RunningInIDE", vbNullString, CompName, False, bCurModuleOnly, False, , bIngnoreClsUserD)

End Function

Private Function strBugTrap(strExitType As String, _
                            strProcName As String) As String

  'Template code for BugTrap inserted at end of procedures
  'v 2.1.5 added "On Error GoTo' to turn the trap off

  strBugTrap = "On Error GoTo 0" & vbNewLine & _
               "Exit " & strExitType & vbNewLine & _
               "BugTrap:" & vbNewLine & _
               "BugTrapComment " & DQuote & strProcName & DQuote & vbNewLine & _
               "If RunningInIDE Then" & vbNewLine & _
               "  Resume" & vbNewLine & _
               "Else" & vbNewLine & _
               "  Resume Next" & vbNewLine & _
               "End if" & vbNewLine

End Function

Private Function StrBugTrapSub(Optional bPrivate As Boolean = False, _
                               Optional ByVal strCompName As String = "BugTrapReport") As String

  Dim StrProjName   As String
  Dim StrOutPutFile As String

  'Template of BugTrap Procedure
  If Not bPrivate Then
    StrProjName = FileBaseName(GetActiveProject.FileName)
    If LenB(StrProjName) Then
      StrOutPutFile = StrProjName & "_BugReport.txt"
     Else
      StrOutPutFile = strCompName & ".txt" ' default name if projname doesn't exist yet
    End If
   Else
    StrOutPutFile = strCompName & "_BugReport.txt"
  End If
  'StrOutputPath = "App.Path & " & DQuote & "\" & StrOutPutFile & DQuote
  StrBugTrapSub = vbNewLine & _
   IIf(bPrivate, "Private", "Public") & " Sub BugTrapComment(ByVal strloc As String) " & vbNewLine & _
   "  Dim strMsg    As String" & vbNewLine & _
   "  Dim strAdvice As String" & vbNewLine & _
   "  Dim FFile     As Integer" & vbNewLine & _
   "  If Not RunningInIDE Then" & vbNewLine & _
   "    strAdvice = " & DQuote & "When reporting the bug please include the file" & strInSQuotes(StrOutPutFile, True) & DQuote & " & vbNewLine & _" & vbNewLine & _
   "                " & DQuote & "You will find the file in the project's folder." & DQuote & vbNewLine & _
   "  End If" & vbNewLine & _
   "  strMsg = strloc & vbNewLine & _" & vbNewLine & _
   DQuote & "Error(" & DQuote & " & Err.Number & " & DQuote & ") " & DQuote & " & Err.Description" & vbNewLine & _
   "  If Err.Number <> 0 Then " & vbNewLine & _
   "    MsgBox strMsg & strAdvice, vbCritical, App.EXEName" & vbNewLine & _
   "    FFile = FreeFile" & vbNewLine & _
   " Open App.Path & " & DQuote & "\" & StrOutPutFile & DQuote & " For Append As #FFile" & vbNewLine & _
   "    Print #FFile, strMsg" & vbNewLine & _
   "    Print #FFile, " & DQuote & "----------------------------------------------------------" & DQuote & vbNewLine & _
   "    Close #FFile " & vbNewLine & _
   "  End If" & vbNewLine & _
   "  If RunningInIDE Then" & vbNewLine & _
   "  'This Stop is not accessable in compiled programs" & vbNewLine & _
   "    Stop" & vbNewLine & _
   "  End If" & vbNewLine & _
   "End Sub"


End Function

Private Function StrCreateRunningInIDE(Optional bPrivate As Boolean = False) As String

  'Template of a support function for BugTrap

  StrCreateRunningInIDE = vbNewLine & _
   IIf(bPrivate, "Private", "Public") & " Function RunningInIDE() As Boolean" & vbNewLine & _
   "  On Error GoTo RunningInIDEErr" & vbNewLine & _
   "  Debug.Print 1 / 0  'Divide by zero (fails within IDE)" & vbNewLine & _
   "Exit Function  'Exit if no error" & vbNewLine & _
   "RunningInIDEErr:" & vbNewLine & _
   "  RunningInIDE = True 'We get error if Debug.Print was evaluated" & vbNewLine & _
   "End Function"

End Function

Private Function TestOtherDebug(ArrProc As Variant, _
                                Optional ByVal bSilent As Boolean = False) As Boolean

  Dim strMsg As String

  'v2.5.9 added bSilent which stops this aborting the menu when it pops up
  'check for existing error handling and issue warning if necessary
  If DetectNonBugTrapErrorHandling(ArrProc) Then
    strMsg = "Procedure already contains non-BugTrap error handling."
    TestOtherDebug = True
   ElseIf DetectResumeErrorHandling(ArrProc) Then
    strMsg = "Procedure contains 'On Error Resume' error handling."
    TestOtherDebug = True
  End If
  If Not bSilent Then
    If TestOtherDebug Then
      mObjDoc.Safe_MsgBox strMsg & vbNewLine & _
                    "It is unwise to double up on error handling." & vbNewLine & _
                    "You must remove other error handling before inserting BugTrap.", vbInformation
    End If
  End If

End Function


':)Code Fixer V3.0.9 (25/03/2005 4:25:47 AM) 19 + 329 = 348 Lines Thanks Ulli for inspiration and lots of code.

