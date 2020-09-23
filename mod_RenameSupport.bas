Attribute VB_Name = "mod_RenameSupport"
Option Explicit

Public Sub FrontTabEnable(ByVal bEnabled As Boolean)

  Dim I As Long

  With frm_FindSettings
    For I = XVerbose To XUsageComments
      .chkUser(I).Enabled = bEnabled
    Next I
  End With

End Sub

Public Function ModuleAttributeDanger(ByVal lngIndex As Long) As Boolean

  Dim I    As Long
  Dim Comp As VBComponent

  Set Comp = VBInstance.VBProjects(ModDesc(lngIndex).MDProj).VBComponents(ModDesc(lngIndex).MDName)
  For I = 1 To Comp.FileCount
    If IsFileReadOnly(Comp.FileNames(I)) Then
      ModuleAttributeDanger = True
      Exit For
    End If
    If IsFileHidden(Comp.FileNames(I)) Then
      ModuleAttributeDanger = True
      Exit For
    End If
  Next I

End Function

Public Sub SetVBP_VBGRead_Write()

  Dim strSourceFolder As String
  Dim GRPName         As String
  Dim I               As Long

  'VBG and VBP files should never be read-only
  'but if you get them off a CD, DVD the attribute may have been set
  'so this automatically converts them to Read-Write
  strSourceFolder = GetProjFolder
  If Len(GetActiveProject.FileName) Then
    'only copy thhese if project is loaded
    GRPName = Dir(strSourceFolder & "\*.vbg")
    If Len(GRPName) Then
      Do
        MakeReadable strSourceFolder & "\" & GRPName
        GRPName = Dir()
      Loop While LenB(GRPName)
    End If
    If Len(strSourceFolder) Then
      For I = 1 To GetActiveProject.Collection.Count
        MakeReadable GetActiveProject.Collection(I).FileName
      Next I
    End If
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:17:29 AM) 1 + 62 = 63 Lines Thanks Ulli for inspiration and lots of code.
