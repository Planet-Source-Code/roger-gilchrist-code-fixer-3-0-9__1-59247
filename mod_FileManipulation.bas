Attribute VB_Name = "mod_FileManipulation"
Option Explicit
Public FSO                     As New FileSystemObject
Public bSomeFilesReadOnly      As Boolean

Public Sub BackUpDeleteEngine(Optional ByVal strBackUpBaseName As String = "BackUp", _
                              Optional DeleteNamedOnly As String = vbNullString, _
                              Optional DeleteAll As Boolean = False)

  Dim AllSubFolders   As Folder
  Dim AFolder         As Folder
  Dim TerminalFolder  As String
  Dim DoDelete        As Boolean
  Dim strSourceFolder As String

  'Copyright 2003 Roger GIlchrist
  'PURPOSE:
  'Modifed Oct 10. Stripped down version specific to Code Fixer
  'Delete Folder(s) below strSourceFolder
  'strSourceFolder = now calculated within procedure, holds the path to the project folder
  'Optional Strings
  'strBackUpBaseName = Default="BackUp"
  'DeleteNamedOnly not ""  then only the named sub-folder is deleted
  '(NOTE there is no check that it exists and code will crash if it doesn't so only call on known folders (from a listbox in CF)
  'Deleteall =  deletes all sub-folders whose names start with the DeleteNamedOnly
  strSourceFolder = GetProjFolder
  If Len(strSourceFolder) Then
    If FSO.FolderExists(strSourceFolder) Then
      Set AllSubFolders = FSO.GetFolder(strSourceFolder)
      For Each AFolder In AllSubFolders.SubFolders
        If InStr(AFolder, strBackUpBaseName) Then
          With FSO
            If Not .GetFolder(AFolder).IsRootFolder Then
              If Not .GetSpecialFolder(WindowsFolder) = strSourceFolder Then
                If Not .GetSpecialFolder(SystemFolder) = strSourceFolder Then
                  TerminalFolder = Mid$(AFolder, InStrRev(AFolder, "\"))
                  If Not InStr("|\NEVERBACKMEUP|\My Documents|\Program Files|\Windows|\System|\System32", TerminalFolder) Then
                    If InStr(TerminalFolder, "_" & Safe_Replace(Date, "/", "-") & "_") = 0 Or LenB(DeleteNamedOnly) Or DeleteAll Then
                      If Len(DeleteNamedOnly) Then
                        DoDelete = True
                      End If
                      If DeleteAll Then
                        DoDelete = True
                      End If
                      If Not DoDelete Then
                        If frm_CodeFixer.chkOdlBackWarn = vbChecked Then
                          'Xcheck(XOldBackUpWarning) Then
                          Select Case mObjDoc.Safe_MsgBox("Backup folders not made today have been found." & vbNewLine & _
                                "Delete Them?", vbCritical + vbYesNo)
                           Case vbYes
                            DoDelete = True
                           Case vbNo
                            DoDelete = False
                          End Select
                         Else
                          DoDelete = False
                        End If
                      End If
                      If DoDelete Then
                        If Len(DeleteNamedOnly) Then
                          If SmartLeft(.GetFileName(AFolder), DeleteNamedOnly) Then
                            .DeleteFolder AFolder
                          End If
                         Else
                          .DeleteFolder AFolder
                        End If
                      End If
                    End If
                   Else
                    mObjDoc.Safe_MsgBox "Cannot delete " & TerminalFolder, vbCritical
                  End If
                 Else
                  mObjDoc.Safe_MsgBox "Cannot delete System Folder", vbCritical
                End If
               Else
                mObjDoc.Safe_MsgBox "Cannot delete Windows Folder", vbCritical
              End If
             Else
              mObjDoc.Safe_MsgBox "Cannot delete Root Folder", vbCritical
            End If
          End With
          If Not DoDelete Then
            Exit For
          End If
        End If
      Next AFolder
     Else
      mObjDoc.Safe_MsgBox "Folder: " & strSourceFolder & vbNewLine & _
                    "does not exist, so no sub-folders to delete.", vbInformation
    End If
  End If

End Sub

Public Sub BackUpDeleteSelected(ByVal BSelected As Boolean, _
                                Optional ByVal DeleteAll As Boolean = False)

  Dim strDir As String
  Dim I      As Long

  With frm_CodeFixer
    For I = .lstRestore.ListCount - 1 To 0 Step -1
      If BSelected Then
        If .lstRestore.Selected(I) Or DeleteAll Then
          strDir = .lstRestore.List(I)
        End If
       Else
        If Not .lstRestore.Selected(I) Then
          strDir = .lstRestore.List(I)
        End If
      End If
      If LenB(strDir) Then
        BackUpDeleteEngine "CodeFixBackUp", strDir
        strDir = vbNullString
      End If
    Next I
  End With
  loadRestore "CodeFixBackUp"

End Sub

Private Sub BackUpEngine(Optional ByVal strBackUpBaseName As String = "BackUp", _
                         Optional strBackUpDir As String)

  Dim strBackUpPath   As String
  Dim TerminalFolder  As String
  Dim strSourceFolder As String

  'Copyright 2003 Roger GIlchrist
  'PURPOSE:
  'Create a Back Up Folder below strSourceFolder
  'Modifed Oct 10. Stripped down version specific to Code Fixer
  'strBackUpDir returns the name generated in this procedure
  'for use in further processing
  'REQUIREMENTS
  ' in Declaration section
  '
  'in Project|References select and check
  '           MicroSoft Scripting Runtime
  'NOTE: This code does not back up back-ups with the same BaseName
  '      (otherwise you quickly end up with many, many copies of each back-up)
  strSourceFolder = GetProjFolder
  If Len(strSourceFolder) Then
    SectionMessage "Backing Up Source Folder", 1 / 1
    With FSO
      ' strSourceFolder = ParentFolderName(strSourceFolder)
      TerminalFolder = Mid$(strSourceFolder, InStrRev(strSourceFolder, "\"))
      If .FolderExists(strSourceFolder) Then
        If Not .GetFolder(strSourceFolder).IsRootFolder Then
          If Not .GetSpecialFolder(WindowsFolder) = strSourceFolder Then
            If Not .GetSpecialFolder(SystemFolder) = strSourceFolder Then
              If Not MultiRight(TerminalFolder, False, "\NEVERBACKMEUP", "\My Documents", "\Program Files", "\Windows", "\System", "\System32") Then
                Do
                  strBackUpPath = strSourceFolder & "\" & strBackUpBaseName & "_" & Safe_Replace(Date, "/", "-") & "_" & Safe_Replace(Time, ":", "-")
                  DoEvents
                  'this loop stops you making 2 backups in the same second (naming convention won't allow it)
                Loop While LenB(Dir(strBackUpPath, vbDirectory))
                'Back up all files in strSourceFolder
                strBackUpDir = strBackUpPath
                .CreateFolder strBackUpPath
                ExportProject strBackUpPath
               Else
                mObjDoc.Safe_MsgBox "Routine BackUpEngine will not make a back up folder from " & TerminalFolder, vbCritical
              End If
             Else
              mObjDoc.Safe_MsgBox "Routine BackUpEngine will not make a back up folder from the System Folder", vbCritical
            End If
           Else
            mObjDoc.Safe_MsgBox "Routine BackUpEngine will not make a back up folder from the Windows Folder", vbCritical
          End If
         Else
          mObjDoc.Safe_MsgBox "Routine BackUpEngine will not make a back up folder from the Root Folder", vbCritical
        End If
       Else
        mObjDoc.Safe_MsgBox "Folder: " & strSourceFolder & vbNewLine & _
                    "does not exist, so cannot be backed up.", vbCritical
      End If
    End With
  End If

End Sub

Public Sub BackUpMakeOne(Optional strBackUpName As String = "CodeFixBackUp", _
                         Optional strReturned As String)

  'v2.16 thanks Johnnie Parrish not your bug but also related
  ' Byval ahd been wrongly applied to strReturn blocking the return value

  BackUpEngine strBackUpName, strReturned
  loadRestore strBackUpName
  frm_CodeFixer.lstRestore.ListIndex = frm_CodeFixer.lstRestore.ListCount - 1

End Sub

Private Sub ExportProject(ByVal strBackUpPath As String)

  Dim Proj            As VBProject
  Dim Comp            As VBComponent
  Dim I               As Long
  Dim J               As Long
  Dim GRPName         As String
  Dim VBPName         As String
  Dim strSourceFolder As String

  'MOve the project files to a backup file
  strSourceFolder = GetProjFolder
  If Len(GetActiveProject.FileName) Then
    'only copy thhese if project is loaded
    'GRPName = Dir ( strSourceFolder&"\"& "*.vbg")
    GRPName = Dir(strSourceFolder & "\*.vbg")
    If Len(GRPName) Then
      Do
        FileCopy strSourceFolder & "\" & GRPName, strBackUpPath & "\" & GRPName
        GRPName = Dir()
      Loop While LenB(GRPName)
    End If
    If Len(strSourceFolder) Then
      For I = 1 To GetActiveProject.Collection.Count
        VBPName = GetActiveProject.Collection(I).FileName
        FileCopy VBPName, strBackUpPath & "\" & FSO.GetFileName(VBPName)
      Next I
    End If
  End If
  If LenB(strSourceFolder) Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        With Comp
          For J = 1 To .FileCount
            If LenB(.FileNames(J)) Then
              FileCopy .FileNames(J), strBackUpPath & "\" & FSO.GetFileName(.FileNames(J))
             Else
              mObjDoc.Safe_MsgBox .Name & " will not be backed up because it has not been saved yet.", vbCritical
            End If
          Next J
        End With 'Comp
      Next Comp
    Next Proj
  End If

End Sub

Private Function FileExists(ByVal FName As String) As Boolean

  If LenB(FName) Then
    FileExists = FSO.FileExists(FName)
  End If

End Function

Public Function FileKill(ByVal FName As String) As Boolean

  If FileExists(FName) Then
    FSO.DeleteFile FName
  End If

End Function

Public Function FilePathOnly(ByVal filespec As String) As String

  If LenB(filespec) Then
    FilePathOnly = FSO.GetParentFolderName(filespec)
  End If

End Function

Public Function GetProjFolder() As String

  'get project home folder

  On Error Resume Next
  GetProjFolder = ParentFolderName(GetActiveProject.FileName)
  If LenB(GetProjFolder) = 0 Then
    'not a pr0ject but either new unsaved project or single module
    GetProjFolder = ParentFolderName(GetActiveProject.VBComponents.Item(1).FileNames(1))
  End If
  On Error GoTo 0

End Function

Public Function isFileDirty2(ByVal Item As MSComctlLib.ListItem) As Boolean

  If Not IsFileReadOnlylsvItem(Item) Then
    isFileDirty2 = VBInstance.VBProjects(Item.Text).VBComponents(Item.SubItems(1)).IsDirty
  End If

End Function

Public Function IsFileReadOnly(ByVal FName As String) As Boolean

  If LenB(FName) Then
    IsFileReadOnly = FSO.GetFile(FName).Attributes And ReadOnly
  End If

End Function

Public Function IsFileReadOnlylsvItem(ByVal Item As MSComctlLib.ListItem) As Boolean

  'Only used to get value for lsvUnDone

  IsFileReadOnlylsvItem = IsFileReadOnly(VBInstance.VBProjects(Item.Text).VBComponents(Item.SubItems(1)).FileNames(1))

End Function

Public Function IsProjectInSingleFolder() As Boolean

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim strSourceDir As String

  IsProjectInSingleFolder = True
  strSourceDir = GetProjFolder
  If LenB(strSourceDir) Then
    For Each Proj In VBInstance.VBProjects
      For Each Comp In Proj.VBComponents
        If Not SmartLeft(Comp.FileNames(1), strSourceDir) Then
          IsProjectInSingleFolder = False
        End If
      Next Comp
    Next Proj
  End If

End Function

Private Function IsProjSaved() As Boolean

  IsProjSaved = LenB(Dir(GetActiveProject.FileName))

End Function

Public Function KillFindLine(LItem As ListItem) As Boolean

  Dim strKill     As String
  Dim strMatching As String
  Dim StartLine   As Long
  Dim strProj     As String
  Dim strCompName As String
  Dim CompMod     As CodeModule

  strProj = LItem.Text
  strCompName = LItem.SubItems(1)
  StartLine = 1
  strMatching = LItem.SubItems(5)
  If InStr(strMatching, RGSignature) Then
    strMatching = Mid$(strMatching, InStr(strMatching, RGSignature))
  End If
  'v2.1.0 fixed
  Set CompMod = GetComponent(strProj, strCompName).CodeModule
  Do While GetFoundCodeLine(strProj, strCompName, strMatching, CompMod.Parent, StartLine)
    strKill = CompMod.Lines(StartLine, 1)
    If strMatching = Trim$(strKill) Then
      CompMod.DeleteLines StartLine, 1
      KillFindLine = True
      StartLine = StartLine - 1
     ElseIf SmartRight(strKill, strMatching) Then
      strKill = Left$(strKill, InStr(strKill, strMatching) - 1)
      If Len(strKill) = 0 Then
        CompMod.DeleteLines StartLine
        KillFindLine = True
       Else
        CompMod.ReplaceLine StartLine, strKill
        KillFindLine = True
      End If
      StartLine = StartLine - 1
     Else
      If SmartLeft(strKill, strMatching) Then
        Exit Do
      End If
    End If
    StartLine = StartLine + 1
    If StartLine > CompMod.CountOfLines Then
      Exit Do
    End If
  Loop
  On Error GoTo 0

End Function

Private Sub loadRestore(Optional ByVal strBackUpBaseName As String = "BackUp")

  Dim strTmp          As String
  Dim strSourceFolder As String

  strSourceFolder = GetProjFolder
  If LenB(strSourceFolder) Then
    frm_CodeFixer.lstRestore.Clear
    strSourceFolder = GetProjFolder
    If LenB(strSourceFolder) Then
      strTmp = Dir(strSourceFolder & "\" & strBackUpBaseName & "*", vbDirectory)
      With frm_CodeFixer
        SendMessage .lstRestore.hWnd, WM_SETREDRAW, False, 0
        If Len(strTmp) Then
          Do
            .lstRestore.AddItem strTmp
            strTmp = Dir(, vbDirectory)
          Loop While LenB(strTmp)
        End If
        If .lstRestore.ListCount > 0 Then
          .lstRestore.ListIndex = .lstRestore.ListCount - 1
         Else
          .LstRestoreFiles.Clear
          .lblFiles.Caption = .LstRestoreFiles.ListCount & " File(s)"
        End If
        .lblBackupcount.Caption = .lstRestore.ListCount & " Backup(s)"
        .cmdBackup(1).Enabled = .lstRestore.ListCount
        .cmdBackup(2).Enabled = .lstRestore.ListCount
        .cmdBackup(3).Enabled = .lstRestore.ListCount
        SendMessage .lstRestore.hWnd, WM_SETREDRAW, True, 0
      End With
    End If
  End If

End Sub

Public Function ParentFolderName(ByVal strFileName As String) As String

  'cut off the filename in a fullyqualified path

  ParentFolderName = FSO.GetParentFolderName(strFileName)

End Function

Private Sub ProcessVBPFiles(ByVal StrDirName As String)

  Dim strFile As String
  Dim strTmp  As String
  Dim arrTmp  As Variant
  Dim I       As Long
  Dim TStream As TextStream

  'read through the VBP file(s) and remove filepaths so that VB will   look in strtup folder for all files.
  If Len(Dir(StrDirName, vbDirectory)) Then
    strFile = Dir(StrDirName & "\*.vbp")
    If Len(strFile) Then
      Do
        Set TStream = FSO.OpenTextFile(StrDirName & "\" & strFile, ForReading)
        strTmp = TStream.ReadAll
        arrTmp = Split(strTmp, vbNewLine)
        For I = LBound(arrTmp) To UBound(arrTmp)
          strTmp = arrTmp(I)
          strTmp = Replace$(strTmp, "=", SngSpace)
          Select Case WordInString(strTmp, 1)
           Case "File", "Module", "Designer", "Class"
            If InStr(arrTmp(I), "\") Then
              arrTmp(I) = Left$(arrTmp(I), InStr(arrTmp(I), SngSpace)) & Mid$(arrTmp(I), InStrRev(arrTmp(I), "\") + 1)
            End If
          End Select
          'Module=newerrors; ..\ModTest\newerrors.bas    'folder is not proj folder or sub-folder
          'Module=newerrors; \ModTest\newerrors.bas      'sub-folder below proj folder
        Next I
        Set TStream = FSO.OpenTextFile(StrDirName & "\" & strFile, ForWriting)
        TStream.Write (Join(arrTmp, vbNewLine))
        Set TStream = Nothing
        strFile = Dir()
      Loop While Len(strFile)
    End If
  End If
  '

End Sub

Public Sub RestoreBackup(ByVal inIndex As Long, _
                         ByVal strBackUpLocation As String, _
                         Optional ByVal forceSingle As Boolean = False)

  Dim Proj          As VBProject
  Dim Comp          As VBComponent
  Dim strBackUpPath As String
  Dim strSource     As String
  Dim strFile       As String
  Dim I             As Long
  Dim strTarget     As String
  Dim arrNewFiles   As Variant
  Dim strNewFiles   As String

  strSource = GetProjFolder
  If LenB(strSource) Then
    If SmartLeft(strBackUpLocation, strSource) = 0 Then
      strBackUpPath = strSource & "\" & strBackUpLocation
     Else
      strBackUpPath = strBackUpLocation
    End If
    If Right$(strBackUpPath, 1) <> "\" Then
      strBackUpPath = strBackUpPath & "\"
    End If
    With frm_CodeFixer
      For I = 0 To .LstRestoreFiles.ListCount - 1
        strFile = .LstRestoreFiles.List(I)
        If .LstRestoreFiles.Selected(I) Or inIndex = 1 Or forceSingle Then
          If forceSingle Then
            strTarget = strSource & "\" & strFile
           Else
            'put file back where ever it came from
            strTarget = SourceHomeFolder(strFile) & strFile
            If strTarget = strFile Then
              strTarget = strSource & "\" & strFile
            End If
          End If
          'v2.16 thanks Johnnie Parrish this was the problem trying to copy files already in the
          'single folder over themselves is illegal
          If strBackUpPath & strFile <> strTarget Then
            If Len(Dir(strTarget)) Then
              Kill strTarget
            End If
            FileCopy strBackUpPath & strFile, strTarget
            strNewFiles = AccumulatorString(strNewFiles, strTarget)
          End If
        End If
      Next I
      arrNewFiles = Split(strNewFiles, ",")
      For Each Proj In VBInstance.VBProjects
        For Each Comp In Proj.VBComponents
          If forceSingle Then
            For I = LBound(arrNewFiles) To UBound(arrNewFiles)
              If FSO.GetFileName(Comp.FileNames(1)) = FSO.GetFileName(arrNewFiles(I)) Then
                If Comp.FileNames(1) <> arrNewFiles(I) Then
                  'v2.16 thanks Johnnie Parrish not your bug but also related
                  'only files that appear in the project window need to reload
                  If IsComponent_Reloadable(Comp) Then
                    Comp.SaveAs arrNewFiles(I)
                    Comp.Reload
                  End If
                  Exit For
                End If
              End If
            Next I
          End If
        Next Comp
      Next Proj
    End With
  End If

End Sub

Public Sub restoreClick()

  Dim strTmp          As String
  Dim strSourceFolder As String

  strSourceFolder = GetProjFolder
  If Len(strSourceFolder) Then
    'get the name of an exiting backup folder folder
    ' and update the file display listbox
    With frm_CodeFixer
      .LstRestoreFiles.Clear
      strTmp = Dir(strSourceFolder & "\" & .lstRestore.Text & "\*")
      If Len(strTmp) Then
        SendMessage .LstRestoreFiles.hWnd, WM_SETREDRAW, False, 0
        Do
          .LstRestoreFiles.AddItem strTmp
          strTmp = Dir()
        Loop While Len(strTmp)
      End If
      .lblFiles.Caption = .LstRestoreFiles.ListCount & " File(s)"
      .cmdRestore(0).Enabled = .LstRestoreFiles.ListCount > 0 And (Not bSomeFilesReadOnly) And IsProjSaved And True
      .cmdRestore(1).Enabled = .LstRestoreFiles.ListCount > 0 And (Not bSomeFilesReadOnly) And IsProjSaved And True
      SendMessage .LstRestoreFiles.hWnd, WM_SETREDRAW, False, 0
    End With
  End If

End Sub

Private Sub SetDirty()

  Dim LItem As ListItem
  Dim I     As Long

  'v2.9.8  error trap for removed modules
  On Error Resume Next
  With frm_CodeFixer.lsvUnDone
    For I = 1 To .ListItems.Count
      SetCurrentLSVLine frm_CodeFixer.lsvUnDone, I
      'safety stuff
      Set LItem = .SelectedItem
      With LItem
        .Checked = VBInstance.VBProjects(.Text).VBComponents(.SubItems(1)).IsDirty
      End With
    Next I
  End With
  On Error GoTo 0

End Sub

Public Sub SetUpFileTool()

  With frm_CodeFixer
    'should really be in fcodefix but placed here for easy of documentation and to keep form file size down
    .lblReload.Caption = "RELOAD" & vbNewLine & _
                         "Names with a check were edited since last saved. Reload the file from its source folder using this tool. If you have already saved the file then use the Backup & Restore Tab"
    .lblNoSavedProj.Visible = Not IsProjSaved
    .lblNoSavedProj.Caption = "You cannot use these tools until you have saved the project/file."
    .cmdReload(0).Enabled = IsProjSaved And (Not bSomeFilesReadOnly) And True
    .cmdReload(1).Enabled = IsProjSaved And (Not bSomeFilesReadOnly) And True
    .cmdSingleFolder(0).Enabled = IsProjSaved
    '(even if there are read_only files you can still make a 'Release' folder
    'v2.16 thanks Johnnie Parrish  not your bug but also related
    'button was not behaving properly because the tests were being performed seperately and turning on falsely
    .cmdSingleFolder(1).Enabled = (Not IsProjectInSingleFolder) And IsProjSaved And (Not bSomeFilesReadOnly)
    .cmdRestore(0).Enabled = IsProjSaved And (Not bSomeFilesReadOnly) And True
    .cmdRestore(1).Enabled = IsProjSaved And (Not bSomeFilesReadOnly) And True
    .cmdBackup(0).Enabled = IsProjSaved
    .cmdBackup(1).Enabled = IsProjSaved
    .cmdBackup(2).Enabled = IsProjSaved
    .cmdBackup(3).Enabled = IsProjSaved
    If IsProjSaved Then
      .lblNoSavedProj.Visible = bSomeFilesReadOnly
      .lblNoSavedProj.Caption = "Reload, Restore && Convert Source are disabled while some files Attributes = Read-Only. Backup Maintanance and Make Release are still available"
    End If
    '
    .lblsinglefolder.Caption = "This tool allows you to bring all the files in your project into a single folder. " & _
                               "OPTIONS" & vbNewLine & _
                               "1. 'Create Release'" & vbNewLine & _
                               "    Copies all current VB project files to a single sub-folder of the program's source folder named 'Release_DATE_TIME' where DATE and TIME are the creation date/time. Your project retains any multiple folder associations it may have while the Release is a single folder version. Why? Project is easier to zip, upload and distibute." & vbNewLine
    .lblsinglefolder.Caption = .lblsinglefolder.Caption & "2. 'Convert Source'" & vbNewLine & _
     "    Copies any VB files not in the source folder to it and rewrites the VBP file(s) to use these. Why? If you have utility modules that you use regularly, copying them to the project folder allows you to modify them for the current project while preserving the original file.(If disabled your project is already in a single folder.)"
  End With
  loadRestore "CodeFixBackUp"
  SetDirty

End Sub

Public Sub SingleFolderWrap(ByVal intIndex As Long)

  Dim strFileName As String

  'making the current project a single folder version
  'involves creating a Release folder copying its contents to the source folder
  'then immediately destroying the Release folder
  'Just click 'make release' if you want one
  'make a backup file in a folder called ProjectName + Release + Date
  'v3.0.9
  BackUpMakeOne IIf(Len(Dir(GetProjFolder & "\*.vbg")), FileBaseName(Dir(GetProjFolder & "\*.vbg")), FileBaseName(GetActiveProject.FileName)) & " Release", strFileName
  'edit the VBP file to used files in release
  ProcessVBPFiles strFileName
  If intIndex = 1 Then
    RestoreBackup 1, strFileName, True
    'v2.16 thanks Johnnie Parrish not your bug but also related
    If Len(strFileName) Then
      FSO.DeleteFolder strFileName
    End If
   Else
    mObjDoc.Safe_MsgBox "Release Created", vbInformation
  End If

End Sub

Private Function SourceHomeFolder(ByVal strFileName As String) As String

  Dim Proj As VBProject
  Dim Comp As VBComponent
  Dim J    As Long
  Dim Hit  As Boolean

  'Search project to find which folder to copy the backup to.
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      For J = 1 To Comp.FileCount
        If SmartRight(Comp.FileNames(J), strFileName) Then
          SourceHomeFolder = ParentFolderName(Comp.FileNames(J))
          Hit = True
          Exit For
        End If
      Next J
      If Hit Then
        Exit For
      End If
    Next Comp
    If Hit Then
      Exit For
    End If
  Next Proj

End Function

Public Function strGetLeftOf(ByVal varIn As Variant, _
                             ByVal strOf As String) As String

  strGetLeftOf = Left$(varIn, InStr(varIn, strOf) - 1)

End Function

Public Function strGetRightOf(ByVal varIn As Variant, _
                              ByVal strOf As String) As String

  strGetRightOf = Mid$(varIn, InStr(varIn, strOf) + Len(strOf))

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:25:00 AM) 3 + 675 = 678 Lines Thanks Ulli for inspiration and lots of code.

