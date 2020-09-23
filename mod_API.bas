Attribute VB_Name = "mod_API"
Option Explicit
'© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
'modified mostly by deleting stuff Ulli used but I don't
'Listbox
'Misc
Public Enum ErrLevel
  BreakUnhandled
  BreakinClass
  BreakonAll
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private BreakUnhandled, BreakinClass, BreakonAll
#End If
Public Const WM_SETREDRAW            As Long = &HB
Private Const HKEY_CURRENT_USER      As Long = &H80000001
Private Const ERROR_NONE             As Long = 0
Private Const VBSettings             As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Private Const REG_SZ                 As Long = 1
Private Const REG_DWORD              As Long = 4
Private Const KEY_ALL_ACCESS         As Long = &H3F
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                       ByVal wMsg As Long, _
                                                                       ByVal wParam As Long, _
                                                                       LParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                            ByVal lpValueName As String, _
                                                                                            ByVal lpReserved As Long, _
                                                                                            lpType As Long, _
                                                                                            ByVal lpData As String, _
                                                                                            lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          ByVal lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                        ByVal lpValueName As String, _
                                                                                        ByVal Reserved As Long, _
                                                                                        ByVal dwType As Long, _
                                                                                        ByVal lpValue As String, _
                                                                                        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal Reserved As Long, _
                                                                                      ByVal dwType As Long, _
                                                                                      lpValue As Long, _
                                                                                      ByVal cbData As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function GetErrorTrapLevel() As ErrLevel

  GetErrorTrapLevel = QueryValue(VBSettings, "BreakOnAllErrors") * 2 + QueryValue(VBSettings, "BreakOnServerErrors")

End Function

Public Function GetFullTabWidth() As Long

  'v 2.1.6 simplifed

  GetFullTabWidth = QueryValue(VBSettings, "TabWidth")
  GetErrorTrapLevel

End Function

Private Function QueryValue(sKeyName As String, _
                            sValueName As String) As Variant

  Dim VValue As Variant   'setting of queried value
  Dim hKey   As Long

  'handle of opened key
  RegOpenKeyEx HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey
  QueryValueEx hKey, sValueName, VValue
  QueryValue = VValue
  RegCloseKey (hKey)

End Function

Private Function QueryValueEx(ByVal lhKey As Long, _
                              ByVal szValueName As String, _
                              VValue As Variant) As Long

  Dim cch    As Long
  Dim lrc    As Long
  Dim LType  As Long
  Dim lValue As Long
  Dim sValue As String

  On Error GoTo QueryValueExError
  ' Determine the size and type of data to be read
  lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, LType, 0&, cch)
  If lrc <> ERROR_NONE Then
    Error 5
  End If
  Select Case LType
    ' For strings
   Case REG_SZ
    sValue = String$(cch, 0)
    lrc = RegQueryValueExString(lhKey, szValueName, 0&, LType, sValue, cch)
    If lrc = ERROR_NONE Then
      VValue = Left$(sValue, cch - 1)
     Else
      VValue = Empty
    End If
    ' For DWORDS
   Case REG_DWORD
    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, LType, lValue, cch)
    If lrc = ERROR_NONE Then
      VValue = lValue
    End If
   Case Else
    'all other data types not supported
    lrc = -1
  End Select
QueryValueExExit:
  QueryValueEx = lrc

Exit Function

QueryValueExError:
  Resume QueryValueExExit

End Function

Public Sub Safe_Sleep()

  'DoEvents allows other programs to run but forces CPU usage to stay at near 100%
  'For some portables this causes heat problems Sleep allows some cycles terminate prematurely
  'allowing empty cycles (or nearly empty) to allow cooling time
  'Thanks to Arron Spivey for suggesting this

  If Xcheck(XLowCPU) Then
    If Rnd > 0.9 Then
      'the rnd means that this hits less often
      'otherwise the program runs waaaay too slow
      Sleep 1
    End If
   Else
    DoEvents
  End If

End Sub

Public Sub SetErrorTrapLevel(ET As ErrLevel)

  Select Case ET
   Case BreakonAll
    SetKeyValue VBSettings, "BreakOnAllErrors", 1, REG_DWORD
    SetKeyValue VBSettings, "BreakOnServerErrors", 0, REG_DWORD
   Case BreakinClass
    SetKeyValue VBSettings, "BreakOnAllErrors", 0, REG_DWORD
    SetKeyValue VBSettings, "BreakOnServerErrors", 1, REG_DWORD
   Case BreakUnhandled
    SetKeyValue VBSettings, "BreakOnAllErrors", 0, REG_DWORD
    SetKeyValue VBSettings, "BreakOnServerErrors", 0, REG_DWORD
  End Select

End Sub

Private Sub SetKeyValue(sKeyName As String, _
                        sValueName As String, _
                        vValueSetting As Variant, _
                        lValueType As Long)

  Dim hKey    As Long         'handle of open key

  'open the specified key
  RegOpenKeyEx HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey
  SetValueEx hKey, sValueName, lValueType, vValueSetting
  RegCloseKey (hKey)

End Sub

Public Sub SetTopMost(frm1 As Form, _
                      ByVal isTopMost As Boolean)

  Const SWP_NOACTIVATE As Long = &H10
  Const SWP_NOSIZE     As Long = &H1
  Const SWP_NOMOVE     As Long = &H2
  Const HWND_TOPMOST   As Long = (-1)
  Const HWND_NOTOPMOST As Long = (-2)

  ' Sets a window to topmost or non-topmost
  ' Does not activate, resize or move the window
  'Uses: Declare Sub SetWindowPos Lib "User" (ByVal hWnd%, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Integer)
  ' Parameters:
  '  frm1        form to be made topmost/non-topmost
  '  isTopMost   true if the form is to be made topmost, false to reset the form to normal
  SetWindowPos frm1.hWnd, IIf(isTopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
  DoEvents

End Sub

Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           LType As Long, _
                           VValue As Variant) As Long

  Dim sValue As String

  Select Case LType
   Case REG_SZ
    sValue = VValue & vbNullChar
    SetValueEx = RegSetValueExString(hKey, sValueName, 0&, LType, sValue, Len(sValue))
   Case REG_DWORD
    SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, LType, CLng(VValue), 4)
  End Select

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:11:52 AM) 30 + 154 = 184 Lines Thanks Ulli for inspiration and lots of code.

