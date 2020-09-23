Attribute VB_Name = "mod_BugTrap"
'This is a modified version of mod_BugTrap which contains a few adaptations for CF specific purposes
Option Explicit

Public Sub BugTrapComment(ByVal strloc As String)

  Dim strMsg    As String
  Dim strAdvice As String
  Dim FFile     As Long

  'next 4 are not created by mod_BugHit
  If Err.Number = 462 Then
    MsgBox "VB terminated while Code Fixer was Formatting and/or Fixing", vbCritical, AppDetails
    bAborting = True
  End If
  If Not RunningInIDE Then
    strAdvice = "When reporting the bug please include the file 'CF2BUGReports.txt' " & vbNewLine & _
                "You will find the file in the project's folder."
  End If
  strMsg = strloc & vbNewLine & _
   "Error(" & Err.Number & ") " & Err.Description & vbNewLine & _
   mObjDoc.ErrorPos
  'ERRORPOS is not a part of the Template version of BugTrap---^^^^^^^^
  If Err.Number <> 0 Then
    MsgBox strMsg & strAdvice, vbCritical, AppDetails
    FFile = FreeFile
    Open App.Path & "\CF2BUGReports.txt" For Append As #FFile
    Print #FFile, GetActiveProject.VBComponents.Item(1).FileNames(1)
    Print #FFile, strMsg
    Print #FFile, "----------------------------------------------------------"
    Close #FFile
  End If
  If RunningInIDE Then    'This Stop is not accessable in compiled programs
    Stop ' this is legal as you can only get here in IDE
  End If

End Sub

Public Function RunningInIDE() As Boolean

  'WARNING This is NOT for bugTrap only DO not Delete if removing BugTrap.

  On Error GoTo RunningInIDEErr
  Debug.Print 1 / 0  'Divide by zero (fails within IDE) DO NOT DELETE THIS LINE

Exit Function

RunningInIDEErr:
  RunningInIDE = True 'We get error if Debug.Print was evaluated

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:26:51 AM) 2 + 50 = 52 Lines Thanks Ulli for inspiration and lots of code.

