VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ByteCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'(c) 2003 roger gilchrist
'<rojagilkrist@hotmail.con>
'feel free to use this code but leave copyrights in source
'
'this enum simplifies getting various output strings representing the CheckBox settings
Public Enum CheckValues
  XVerbose
  XPrevCom
  XStructCom
  XBackup
  XStayOnTop
  XVisScan
  XLowCPU
  XReadWrite
  XSpaceSep
  XBlankPreserve
  XIgnoreCom
  XNoIndentCom
  XNoLargeFileMsg
  XNoPrjWarning
  XNoCntrlWarning
  ' XnoSubClassWarning
  XUsageComments
  'XDeleteOldCarrayCode
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private XVerbose, XPrevCom, XStructCom, XBackup, XStayOnTop, XVisScan, XLowCPU, XReadWrite
Private XSpaceSep, XBlankPreserve, XIgnoreCom, XNoIndentCom, XOldBackUpWarning, XNoPrjWarning, XNoCntrlWarning, XUsageComments, XnoSubClassWarning, XDeleteOldCarrayCode
#End If
Public Enum StrMode
  IndexNo
  CaptionStr
  BinaryStr
  BinaryInverseStr
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private IndexNo, CaptionStr, BinaryStr, BinaryInverseStr
#End If
Public Enum Inversion
  Standard
  Inverted
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Standard, Inverted
#End If
Private arrCBox        As Variant
Private m_app          As String
Private m_section      As String
Private m_key          As String
Private m_MaxValue     As Long
Private m_Default      As Long

Private Function ByteValue(ByVal ByteNumber As Long) As Long

  'you could do this in-line but this makes the operation clearer
  'and lowers chance of a mistype

  ByteValue = 2 ^ ByteNumber

End Function

Public Sub Init(arrCB As Variant, _
                Optional ByVal AppName As String, _
                Optional ByVal Section As String = "ChkSection", _
                Optional ByVal Key As String, _
                Optional ByVal AutoLoad As Boolean = True, _
                Optional ByVal DefaultValue As Long = 0)

  Dim I As Long

  'Optional Section and Key strings mean that you can use this class without thinking about it
  'The key value is either user specified or automatically generated form the CheckBox
  '
  'AutoLoad: set to False if you are loading multiple groups of interactive CheckBoxes
  'as you do not want to trigger their clcik events until all the
  'CheckBoxes are initialized
  Set arrCBox = arrCB
  If LegalControls Then
    m_section = Section
    If Len(AppName) Then
      m_app = AppName
     Else
      m_app = App.Title
    End If
    If Len(Key) Then
      m_key = Key
     Else ' if key is not specified then generate one from the controls's own name
      m_key = arrCBox(I).Name
    End If
    For I = 0 To arrCBox.Count - 1 ' initialize values
      m_MaxValue = m_MaxValue + ByteValue(I)
    Next I
    m_Default = DefaultValue
    If AutoLoad Then
      LoadCheck
    End If
  End If

End Sub

Public Property Get Inversion() As Long

  Dim I      As Long

  Inversion = m_MaxValue
  For I = 0 To arrCBox.Count - 1
    Inversion = Inversion - arrCBox(I) * ByteValue(I)
  Next I

End Property

Private Function LegalControls() As Boolean

  Dim I      As Long
  Dim arName As String

  'this is procedure is designed to help programmers
  'it prevents you from using incorrect controls with the class
  'you can delete it (and the calling code) from compiled releases
  'but leave it in if you are releasing code as next user may also need it
  'This class only works with contiguous arrays of CheckBoxes so check that none have been skipped
  On Error GoTo InCompleteArray
  For I = 0 To arrCBox.Count - 1
    arName = arrCBox(I).Name ' if the control doesn't exist this will cause an error
    ' NOTE if the missing element is '0' then no name will appear in the MsgBox
  Next I
  On Error GoTo 0
  'becuase you need to use Variant to pass the CheckBox array into the class
  'the following test is a safety mechanism to make sure you have passed in CheckBoxes
  If Not TypeOf arrCBox(0) Is CheckBox Then
    MsgBox "clsByteCheckBox needs a contiguous array of checkBoxes to work", vbCritical
   Else
    LegalControls = True
  End If

Exit Function

InCompleteArray:
  MsgBox "The array of Checkboxes " & IIf(Len(arName), strInSQuotes(arName), "") & " must be a contiguous array numbered from '0' for clsByteCheckBox to work.", vbCritical

End Function

Public Sub LoadCheck()

  'default(first run) value is all OFF

  TotalValue = CLng(GetSetting(m_app, m_section, m_key, m_Default))

End Sub

Public Sub SaveCheck()

  SaveSetting m_app, m_section, m_key, TotalValue

End Sub

Public Property Get TotalValue() As Long

  Dim I      As Long

  'read the value that the class stores to set the checkbox array
  For I = 0 To arrCBox.Count - 1
    TotalValue = TotalValue + arrCBox(I) * ByteValue(I)
  Next I

End Property

Public Property Let TotalValue(ByVal Prop As Long)

  Dim I As Long

  'set the value the class will store for the checkbox array
  For I = 0 To arrCBox.Count - 1
    If Prop And ByteValue(I) Then
      arrCBox(I).Value = 1
     Else
      arrCBox(I).Value = 0
    End If
  Next I

End Property

Public Property Get Value(IchkNumber As CheckValues) As Boolean
Attribute Value.VB_UserMemId = 0

  'Get individual member of an array of CheckBoxes
  'Equivalent of 'If Check(index).Value = vbValue Then'
  'Allows you to set up an Enum using descriptive names and then call using
  '- If Xcheck.Value(CBBackUp) Then'
  'NOTE this is the default property of the class
  'this means that you can use the format
  '<InstanceName>(index)
  'to set and get the value
  'Example
  '- If Xcheck(CBBackUp) Then

  Value = TotalValue And ByteValue(IchkNumber)

End Property

Public Property Let Value(IchkNumber As CheckValues, _
                          ByVal Nval As Boolean)

  'set individual member of an array of CheckBoxes
  'Equivalent of 'Check(index).Value = vbValue Then'
  'Allows you to set up an Enum using descriptive names and then call using
  'Value(QBByte1)= True' not you can also use 1|0 for more checkbox-like coding
  'v2.4.4 reconfigured for speed

  If Nval Then
    If (TotalValue And ByteValue(IchkNumber)) Then
      TotalValue = TotalValue - ByteValue(IchkNumber)
     Else
      TotalValue = TotalValue + ByteValue(IchkNumber)
    End If
  End If

End Property

':)Code Fixer V3.0.9 (25/03/2005 4:11:50 AM) 52 + 163 = 215 Lines Thanks Ulli for inspiration and lots of code.

