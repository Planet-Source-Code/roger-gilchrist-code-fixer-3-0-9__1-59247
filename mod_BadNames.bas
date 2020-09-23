Attribute VB_Name = "mod_BadNames"
Option Explicit

Public Function BadNameMsg2(ByVal lngBadNameType As Long) As String

  BadNameMsg2 = vbNullString
MultiFormMsg:
  Select Case lngBadNameType
    'Case BNNone '
   Case BNClass 'classname
    BadNameMsg2 = ":Control's class name." & BadNameMsg2
   Case BNReserve 'ReserveWord
    BadNameMsg2 = ":VB reserved word." & BadNameMsg2
   Case BNKnown, BNCommand 'VBCommand
    BadNameMsg2 = ":VB command" & BadNameMsg2
   Case BNVariable 'Variable
    BadNameMsg2 = ":Code variable" & BadNameMsg2
   Case BNProc 'Procname
    BadNameMsg2 = ":Code Procedure/Parameter" & BadNameMsg2
   Case BNMultiForm 'multiforms
    BadNameMsg2 = ":Control on another form"
   Case BNDefault 'DefaultVB
    BadNameMsg2 = ":VB default" & BadNameMsg2
   Case BNSingle 'singleletter
    BadNameMsg2 = ":Single letter" & BadNameMsg2
   Case BNStructural
    BadNameMsg2 = ":VB Strucutral word" & BadNameMsg2
   Case BNSingletonArray
    BadNameMsg2 = ": Singleton Control Array" & BadNameMsg2
   Case Is > BNSingletonArray
    Select Case lngBadNameType
     Case BNMultiForm
      BadNameMsg2 = ":Control on another form"
     Case Is > BNMultiForm
      BadNameMsg2 = " AND Control on another form"
      lngBadNameType = lngBadNameType - BNMultiForm
      GoTo MultiFormMsg
     Case Else
      BadNameMsg2 = " AND Singleton Array"
      lngBadNameType = lngBadNameType - BNSingletonArray
      GoTo MultiFormMsg
    End Select
  End Select

End Function

Private Function IllegalFileCharacters(ByVal strTest As String) As Boolean

  Dim I As Long

  For I = 1 To Len(strTest)
    If InStr("\/:*?<>|" & DQuote, Mid$(strTest, I, 1)) Then
      IllegalFileCharacters = True
      Exit For
    End If
  Next I

End Function

Public Function LegalName(Strnew As String, _
                          strPath As String, _
                          ByVal Mode As Long, _
                          Cmp As VBComponent, _
                          Prj As VBProject) As Boolean

  Dim strNameType As String

  Strnew = Trim$(Strnew)
  Select Case Mode
   Case 0 'Modulename
    If Len(Strnew) < 45 Then
      If Not IllegalCharacters(Strnew) Then
        If InStr(Strnew, SngSpace) = 0 Then
          LegalName = True
        End If
      End If
    End If
    strNameType = "module name."
   Case 1 'filename comp
    If Not IllegalFileCharacters(Strnew) Then
      If InStr(Strnew, ".") = 0 Then
        Strnew = Strnew & "." & FileExtention(Cmp.FileNames(1))
      End If
      If FileExtention(Cmp.FileNames(1)) = FileExtention(Strnew) Then
        strPath = Replace$(Cmp.FileNames(1), FileNameOnly(Cmp.FileNames(1)), Strnew)
        If Len(strPath) < 255 Then
          LegalName = True
        End If
      End If
    End If
    strNameType = "filename."
   Case 2 'filename Proj
    If Not IllegalFileCharacters(Strnew) Then
      If InStr(Strnew, ".") = 0 Then
        Strnew = Strnew & "." & FileExtention(Prj.FileName)
      End If
      If FileExtention(Prj.FileName) = FileExtention(Strnew) Then
        strPath = Replace$(Prj.FileName, FileNameOnly(Prj.FileName), Strnew)
        If Len(strPath) < 255 Then
          LegalName = True
        End If
      End If
    End If
    strNameType = "filename."
  End Select
  If Not LegalName Then
    mObjDoc.Safe_MsgBox strInSQuotes(Strnew) & " is not a valid " & strNameType, vbCritical
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:12:25 AM) 1 + 107 = 108 Lines Thanks Ulli for inspiration and lots of code.

