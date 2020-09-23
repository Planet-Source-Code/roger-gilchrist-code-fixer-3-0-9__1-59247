VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl PropertyList 
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   ScaleHeight     =   4515
   ScaleMode       =   0  'User
   ScaleWidth      =   5280
   Begin VB.TextBox txtDescription 
      BackColor       =   &H8000000F&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "PropertyList.ctx":0000
      Top             =   2280
      Width           =   4215
   End
   Begin VB.PictureBox picButtonContainer 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdSetter 
         Appearance      =   0  'Flat
         Caption         =   "Dummy"
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.ListBox lstInput 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picButton 
      Height          =   375
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdPropertyList 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin MSComctlLib.ImageList imlPropBox 
      Left            =   2760
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyList.ctx":000D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyList.ctx":032F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyList.ctx":0651
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyList.ctx":0A17
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PropertyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'by roger gilchrist
'modified from 'VB6 PropertyList' by Octoni at PSC
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55807&lngWId=1
'2nd version
'
'Partial key board access you can scroll select and open listboxes with keys but
' you can only select one up or down in the list from the keyboard and have to leave the
' row and reenter to reopen the list. If you can help please let me know
'
Option Explicit
Private bListScroll                        As Boolean
Private bDescriptionShow                   As Boolean
Private bIinitialized                      As Boolean
'Default Property Values:
Private lngLongestName                     As Long
Private lngLongestOption                   As Long
Private StrDefault                         As String
Private strLongestOption                   As String
Private PropertyCount                      As Variant
Private mDefaultButtons                    As Boolean
Private mForeColor                         As Long
Private mListColor                         As Long
'
Private WithEvents cGrid                   As MSFlexGrid
Attribute cGrid.VB_VarHelpID = -1
Private WithEvents cText                   As VB.TextBox
Attribute cText.VB_VarHelpID = -1
Private WithEvents cList                   As VB.ListBox
Attribute cList.VB_VarHelpID = -1
Private WithEvents cButton                 As VB.PictureBox
Attribute cButton.VB_VarHelpID = -1
'
Private Const m_def_BackColor              As Long = vbWhite
Private Const m_def_ListColor              As Long = vbWhite
Private Const m_def_ForeColor              As Long = vbBlack
Private Const m_def_BackColorFixed         As Long = vbButtonFace
Private Const m_def_BackColorSel           As Long = vbWhite
Private Const m_def_Locked                 As Long = 0
'Property Variables:
Private m_BackColor                        As OLE_COLOR
Private m_BackColorFixed                   As OLE_COLOR
Private m_BackColorSel                     As OLE_COLOR
Private m_Font                             As Font
Private m_Locked                           As Boolean
Private mCurrentItem                       As Long
Private Enum ColNames 'identify the columns in the FlexiGrid tool
  '<FIXME:-) :WARNING: This UserControl member is not used in the current project,
  ePropName           'displays property name
  eValue              'displays current property setting
  eType               'hidden store for value input mode (only eList impelimented
  eOptions            'hidden store stores string of possible inputs for elist
  eDefValue           'hidden store default value (may not be needed)
  eTag                'hidden store anything not used yet perhaps a description of the property
  eOptLimits
  eButtonType
  eNoOptions
  'hidden store button type can be Elipsis ... or ListArrow (always arrow in this control
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eOptLimits, eButtonType, eNoOptions
#End If
Public Enum eInputMode
  eText
  eList
  eButton
  eColor
  efile
  eFont
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eText, eList, eButton, eColor, efile, eFont
#End If
'Event Declarations:
Public Event ButtonClick(strProperty As String)
Public Event ComboClick(strProperty As String)
Public Event Change()
'Public Event Resize()
Private mDefaultlist                       As String
Private Const GWL_STYLE                    As Long = (-16)
Private Const WS_HSCROLL                   As Long = &H100000
Private Const WS_VSCROLL                   As Long = &H200000
'Used by VisibleScrollBars function
Public Enum Enum_VisibleScrollBars
  vs_none = 0
  vs_vertical = 1
  vs_horizontal = 2
  vs_both = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private vs_none, vs_vertical, vs_horizontal, vs_both
#End If
Private Const SM_CXVSCROLL                 As Long = 2
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Sub AddProperty(ByVal LimitStr As String, _
                       ByVal strProperty As String, _
                       Optional ByVal strDesc As String, _
                       Optional ByVal strListValues As String, _
                       Optional ByVal PropertyType As eInputMode = eList)

  Dim I        As Long
  Dim InsertAt As Long

  With grdPropertyList
    For I = 1 To .Rows
      If I = .Rows Then
        .Rows = .Rows + 1
      End If
      If LenB(Trim$(.TextMatrix(I, ePropName))) = 0 Then
        If .Rows = I Then
          InsertAt = I - 1
          Exit For
         Else
          InsertAt = I
          Exit For
        End If
      End If
    Next I
    .TextMatrix(InsertAt, ePropName) = strProperty
    If lngLongestName < Len(strProperty) Then
      lngLongestName = Len(strProperty)
    End If
    .TextMatrix(InsertAt, eType) = PropertyType
    .TextMatrix(InsertAt, eTag) = strDesc
    'list only modification
    .TextMatrix(InsertAt, eButtonType) = PropertyType
    If PropertyType = eList Then
      If LenB(strListValues) = 0 Then
        If Len(mDefaultlist) Then
          strListValues = mDefaultlist
         Else
          strListValues = "False|True"
          LimitStr = "X1"
        End If
      End If
      .TextMatrix(InsertAt, eOptions) = OptionClean(strListValues)
      .TextMatrix(InsertAt, eOptLimits) = LimitStr
      If InStr(LimitStr, "X") Then
        .TextMatrix(InsertAt, eDefValue) = InStr(LimitStr, "X") - 1
        StrDefault = StrDefault & InStr(LimitStr, "X") - 1
       Else
        .TextMatrix(InsertAt, eDefValue) = "0"
        StrDefault = StrDefault & "0"
      End If
      .TextMatrix(InsertAt, eNoOptions) = IIf(CountSubString(LimitStr, "0") <> 3, "1", "0")
      If .TextMatrix(InsertAt, eNoOptions) = "0" Then
        .TextMatrix(InsertAt, eTag) = "(NO OPTIONS) " & .TextMatrix(InsertAt, eTag)
      End If
      LongestOption strListValues
    End If
    PropertyCount = InsertAt '.Rows
  End With '

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

  BackColor = m_BackColor
  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=10,0,0,0

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

  m_BackColor = New_BackColor
  grdPropertyList.BackColor = m_BackColor
  PropertyChanged "BackColor"
  PropertyChanged "BackColorFixed"

End Property

Private Sub cButton_Click()

  Select Case cGrid.TextMatrix(mCurrentItem, eType)
   Case eList
    With cList
      If .Visible Then
        .Visible = False
       Else
        ShowList
      End If
    End With 'cList
   Case eText
    RaiseEvent ButtonClick(cGrid.TextMatrix(mCurrentItem, ePropName))
   Case eButton
    RaiseEvent ButtonClick(cGrid.TextMatrix(mCurrentItem, ePropName))
  End Select

End Sub

Public Sub Clear()

  grdPropertyList.FixedRows = 0
  grdPropertyList.Rows = 1

End Sub

Private Sub cList_Click()

  If cList.Enabled Then
    If Not bListScroll Then
      cText.Text = cList.Text
      cGrid.TextMatrix(mCurrentItem, eValue) = " " & Trim$(cText.Text)
      RaiseEvent ComboClick(CInt(mCurrentItem))
      RaiseEvent Change
      ControlsHide
    End If
  End If

End Sub

Private Sub cList_KeyPress(KeyAscii As Integer)

  If cList.Enabled Then
    Select Case KeyAscii
     Case vbKeyReturn
      bListScroll = False
      cList_Click
     Case vbKeyDown
      If cList.ListIndex < cList.ListCount - 1 Then
        KeyAscii = 0
        bListScroll = True
        cList.ListIndex = cList.ListIndex + 1
      End If
     Case vbKeyUp
      If cList.ListIndex > 0 Then
        cList.ListIndex = cList.ListIndex - 1
      End If
    End Select
  End If

End Sub

Private Sub cList_MouseDown(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

  bListScroll = False

End Sub

Private Sub cmdSetter_Click(Index As Integer)

  If Index = 0 Then
    DataString = StrDefault
   Else
    DataString = String$(PropertyCount, CStr(Index - 1))
  End If

End Sub

Private Sub ControlAppearance(ctl As Control)

  On Error Resume Next ' not all controls have all properties
  With ctl
    Set .Font = cGrid.Font
    .Width = cGrid.CellWidth - cGrid.GridLineWidth * 2 '- ScrollOffset
    .Appearance = vbFlat
    .BorderStyle = 0
    .BackColor = mListColor
    .ForeColor = mForeColor
    .Left = cGrid.ColPos(eValue)
    .ZOrder 0
    .Visible = False
  End With
  On Error GoTo 0

End Sub

Private Sub ControlsHide()

  cList.Visible = False
  cButton.Visible = False
  cText.Visible = False
  If cGrid.Visible Then
    cGrid.SetFocus
  End If

End Sub

Private Sub ControlsInitialise()

  ControlAppearance cText
  cText.Height = cGrid.CellHeight - cGrid.GridLineWidth * 2
  ControlAppearance cButton
  With cButton
    .AutoSize = True
    .Picture = imlPropBox.ListImages(1).Picture
    .Left = cGrid.ColPos(eValue + 1) - .Width
  End With 'cButton
  ControlAppearance cList
  'v2.6.3 apparently not necessary
  '  With cList
  '    If .Width < getTextWidth(strLongestOption) Then
  '      .Left = .Left - (getTextWidth(strLongestOption) - .Width)
  '      .Width = getTextWidth(strLongestOption)
  '    End If
  '  End With

End Sub

Private Function CountSubString(strTest As String, _
                                strFind As String) As Long

  CountSubString = UBound(Split(strTest, strFind))

End Function

Public Property Get DataString() As String

  Dim I    As Long
  Dim strD As String

  If bIinitialized Then
    If grdPropertyList.Rows > 2 Then
      For I = 1 To cGrid.Rows - 1
        strD = strD & DataValue(I)
      Next I
    End If
    DataString = strD
  End If

End Property

Public Property Let DataString(ByVal DataStr As String)

  Dim I      As Long
  Dim ValPos As Long
  Dim arrTmp As Variant

  ControlsHide
  If Len(DataStr) < PropertyCount Then
    DataStr = DataStr & String$(PropertyCount - 1 - Len(DataStr), "0")
   ElseIf Len(DataStr) > PropertyCount Then
    DataStr = Left$(DataStr, PropertyCount)
  End If
  For I = 1 To Len(DataStr)
    arrTmp = Split(cGrid.TextMatrix(I, eOptions), "|")
    ValPos = Val(Mid$(DataStr, I, 1)) + 1
    If Mid$(cGrid.TextMatrix(I, eOptLimits), ValPos, 1) <> "0" Then
      If Val(Mid$(DataStr, I, 1)) <= UBound(arrTmp) Then
        cGrid.TextMatrix(I, eValue) = " " & arrTmp(ValPos - 1)
       Else
        cGrid.TextMatrix(I, eValue) = " " & arrTmp(UBound(arrTmp))
      End If
    End If
  Next I
  RaiseEvent Change

End Property

Private Function DataValue(ByVal lngPos As Long) As Long

  Dim arrTmp As Variant
  Dim I      As Long

  arrTmp = Split(cGrid.TextMatrix(lngPos, eOptions), "|")
  For I = LBound(arrTmp) To UBound(arrTmp)
    If arrTmp(I) = Trim$(cGrid.TextMatrix(lngPos, eValue)) Then
      DataValue = I
      Exit For
    End If
  Next I

End Function

Public Property Get DefaultButtons() As Boolean

  DefaultButtons = mDefaultButtons

End Property

Public Property Let DefaultButtons(ByVal bShow As Boolean)

  mDefaultButtons = bShow
  picButtonContainer.Visible = mDefaultButtons
  PropertyChanged "defBut"

End Property

Public Property Get DefaultList() As String

  DefaultList = mDefaultlist

End Property

Public Property Let DefaultList(ByVal strDef As String)

  Dim I      As Long
  Dim arrTmp As Variant

  On Error Resume Next
  If Len(strDef) Then
    mDefaultlist = Replace$(strDef, "&&", "////")
    mDefaultlist = Replace$(mDefaultlist, "&", vbNullString)
    mDefaultlist = Replace$(mDefaultlist, "////", "&")
    arrTmp = Split(strDef, "|")
    LongestOption mDefaultlist
    With cmdSetter(0)
      .Caption = "&Default"
      .Left = 0
      .Width = CLng(UserControl.ScaleWidth \ (UBound(arrTmp) + 2)) - UBound(arrTmp) - 3
    End With 'cmdSetter(0)
    For I = 1 To UBound(arrTmp) + 1
      If I > cmdSetter.Count - 1 Then
        Load cmdSetter(I)
      End If
      With cmdSetter(I)
        .Left = cmdSetter(I - 1).Left + cmdSetter(I - 1).Width - UBound(arrTmp) - 3
        .Caption = arrTmp(I - 1)
        .Visible = True
      End With 'cmdSetter(I)
    Next I
   Else
    DefaultButtons = False
  End If
  PropertyChanged "DefList"
  On Error GoTo 0

End Property

Public Property Get DescriptionShow() As Boolean

  DescriptionShow = bDescriptionShow

End Property

Public Property Let DescriptionShow(ByVal bShow As Boolean)

  bDescriptionShow = bShow
  txtDescription.Visible = bDescriptionShow
  PropertyChanged "ShowDesc"
  UserControl_Resize

End Property

Private Sub Display()

  'hide list if visible

  ControlsHide
  With cGrid
    If .RowIsVisible(.Row) Then
      If .Row > 0 Then
        mCurrentItem = .Row
        txtDescription.Text = "Description: " & .TextMatrix(mCurrentItem, eTag)
        If .Col = eValue Then
          '.Col = ePropName
          .HighLight = flexHighlightAlways
          mCurrentItem = .Row
          If .TextMatrix(mCurrentItem, eNoOptions) = "1" Then
            ShowText
            ShowButton
          End If
        End If
      End If
    End If
  End With

End Sub

Public Sub DrawPropertyBox()

  bIinitialized = True
  UserControl_Resize
  Set cGrid = grdPropertyList
  With cGrid
    .BackColor = m_BackColor
    .ForeColor = mForeColor
  End With 'cGrid
  Set cText = txtInput
  Set cButton = picButton
  Set cList = lstInput
  SetValue
  RaiseEvent Change

End Sub

Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for a FlexGrid."

  FixedRows = grdPropertyList.FixedRows
  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,FixedRows

End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)

  grdPropertyList.FixedRows() = New_FixedRows
  PropertyChanged "FixedRows"

End Property

Public Property Set Font(ByVal New_Font As Font)
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

  Set m_Font = New_Font
  PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR

  ForeColor = mForeColor

End Property

Public Property Let ForeColor(ByVal oleCol As OLE_COLOR)

  mForeColor = oleCol
  PropertyChanged "ForeColor"

End Property

Public Property Get FormatString() As String

  FormatString = grdPropertyList.FormatString
  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,FormatString

End Property

Public Property Let FormatString(ByVal New_FormatString As String)

  With grdPropertyList
    If .FixedRows = 0 Then
      .Rows = .Rows + 1
      .FixedRows = 1
      PropertyChanged "Rows"
    End If
  End With 'MSFlexGrid1
  If grdPropertyList.FixedRows = 1 Then
    If LenB(New_FormatString) = 0 Then
      With grdPropertyList
        .Clear
        .FixedRows = 0
        .Rows = .Rows - 1
      End With 'MSFlexGrid1
      PropertyChanged "Rows"
    End If
  End If
  grdPropertyList.FormatString() = New_FormatString
  PropertyChanged "FormatString"
  UserControl_Resize

End Property

Private Function getTextWidth(pString As String) As Long

  getTextWidth = UserControl.TextWidth(pString)

End Function

Private Sub grdPropertyList_Click()

  Display

End Sub

Private Sub grdPropertyList_EnterCell()

  Display

End Sub

Private Sub grdPropertyList_KeyDown(KeyCode As Integer, _
                                    Shift As Integer)

  If cButton.Visible Then
    cButton_Click
  End If

End Sub

Private Sub grdPropertyList_SelChange()

  cText.Visible = False
  cButton.Visible = False
  cList.Visible = False
  txtDescription.Text = "Description:"

End Sub

Public Property Get ListColor() As OLE_COLOR

  ListColor = mListColor

End Property

Public Property Let ListColor(ByVal oleCol As OLE_COLOR)

  mListColor = oleCol
  PropertyChanged "ListColor"

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."

  Locked = m_Locked
  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=0,0,0,0

End Property

Public Property Let Locked(ByVal New_Locked As Boolean)

  m_Locked = New_Locked
  PropertyChanged "Locked"

End Property

Private Sub LongestOption(ByVal strOpts As String)

  Dim I      As Long
  Dim arrTmp As Variant

  arrTmp = Split(strOpts, "|")
  For I = LBound(arrTmp) To UBound(arrTmp)
    If lngLongestOption < Len(arrTmp(I)) Then
      lngLongestOption = Len(arrTmp(I))
      strLongestOption = UCase$(arrTmp(I))
    End If
  Next I

End Sub

Private Function OptionClean(strOpt As String) As String

  OptionClean = strOpt
  Do While Left$(OptionClean, 1) = "|"
    OptionClean = Mid$(OptionClean, 2)
  Loop
  Do While Right$(OptionClean, 1) = "|"
    OptionClean = Left$(OptionClean, Len(OptionClean) - 1)
  Loop

End Function

Public Property Get Rows() As Long

  Rows = grdPropertyList.Rows - grdPropertyList.FixedRows
  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Rows

End Property

Public Property Let Rows(ByVal New_Rows As Long)

  If grdPropertyList.Rows = 0 Then
    If New_Rows <= 1 Then
      If grdPropertyList.FixedRows Then
        New_Rows = 2
      End If
    End If
  End If
  grdPropertyList.Rows() = New_Rows + grdPropertyList.FixedRows
  PropertyChanged "Rows"
  UserControl_Resize

End Property

Private Function ScrollOffset() As Long

  If VisibleScrollBars(grdPropertyList) Then
    ScrollOffset = GetSystemMetrics(SM_CXVSCROLL)
    ScrollOffset = ScaleY(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
  End If

End Function

Private Sub SetValue()

  Dim I             As Long

  With cGrid
    Set .Font = UserControl.Ambient.Font
    For I = .FixedRows To PropertyCount
      .TextMatrix(I, eValue) = CStr(Split(.TextMatrix(I, eOptions), "|")(Val(.TextMatrix(I, eDefValue))))
    Next I
    .ScrollTrack = True
    .ColAlignment(eValue) = flexAlignLeftCenter
    .Col = eValue
  End With '
  ControlsInitialise

End Sub

Private Sub ShowButton()

  If mCurrentItem > 0 Then
    With cButton
      .Picture = imlPropBox.ListImages(IIf(cGrid.TextMatrix(mCurrentItem, eType) = eList, 1, 2)).Picture
      .Top = cText.Top
      .Visible = True
    End With
  End If

End Sub

Private Sub ShowList()

  Dim I          As Long
  Dim lngLongest As Long
  Dim arrTmp     As Variant

  If mCurrentItem > 0 Then
    With cList
      .Top = cText.Top + cText.Height
      arrTmp = Split(cGrid.TextMatrix(mCurrentItem, eOptions), "|")
      .Clear
      .Height = 0
      For I = LBound(arrTmp) To UBound(arrTmp)
        If LenB(arrTmp(I)) Then
          If Mid$(cGrid.TextMatrix(mCurrentItem, eOptLimits), I + 1, 1) <> "0" Then
            .AddItem arrTmp(I)
            If lngLongest < Len(arrTmp(I)) Then
              lngLongest = Len(arrTmp(I))
            End If
            .Height = .Height + cGrid.CellHeight
          End If
        End If
      Next I
      .Height = .Height - cGrid.CellHeight / 2
      If .Top + .Height > UserControl.Height Then
        'if too high the show above target text
        .Top = .Top - .Height - cGrid.CellHeight
      End If
      .Width = cText.Width
      'v2.6.3 simplified to avoid
      .Left = cText.Left
      If .ListCount > 1 Then
        .Visible = True
        .Enabled = False
        .Text = cText.Text
        .Enabled = True
        .SetFocus
       Else
        ControlsHide
      End If
    End With
  End If

End Sub

Private Sub ShowText()

  If mCurrentItem > 0 Then
    With cText
      .Text = Trim$(cGrid.TextMatrix(mCurrentItem, eValue))
      .Top = cGrid.Top + cGrid.CellTop
      .Visible = True
    End With
  End If

End Sub

Private Sub UserControl_InitProperties()

  m_BackColor = m_def_BackColor
  'Initialize Properties for User Control
  ListColor = m_def_BackColor
  ForeColor = m_def_ForeColor
  m_BackColorFixed = m_def_BackColorFixed
  m_BackColorSel = m_def_BackColorSel
  Set m_Font = Ambient.Font
  m_Locked = m_def_Locked
  bDescriptionShow = True
  Rows = 5
  UserControl_Resize
  FormatString = "Property|Value"

End Sub

Private Sub UserControl_LostFocus()

  ControlsHide

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  With PropBag
    mForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
    mListColor = .ReadProperty("ListColor", m_def_ListColor)
    DescriptionShow = .ReadProperty("ShowDesc", True)
    m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
    m_BackColorFixed = .ReadProperty("BackColorFixed", m_def_BackColorFixed)
    m_BackColorSel = .ReadProperty("BackColorSel", m_def_BackColorSel)
    grdPropertyList.Cols = .ReadProperty("Cols", 1)
    If .ReadProperty("FixedRows", 0) Then
      If .ReadProperty("Rows", 1) = 1 Then
        grdPropertyList.Rows = 2
       Else
        grdPropertyList.Rows = .ReadProperty("Rows", 1)
      End If
     Else
      grdPropertyList.Rows = .ReadProperty("Rows", 1)
    End If
    grdPropertyList.FixedRows = .ReadProperty("FixedRows", 0)
    Set m_Font = .ReadProperty("Font", Ambient.Font)
    grdPropertyList.FormatString = .ReadProperty("FormatString", vbNullString)
    m_Locked = .ReadProperty("Locked", m_def_Locked)
    DefaultList = .ReadProperty("DefList", vbNullString)
    DefaultButtons = .ReadProperty("defBut", False)
  End With
  UserControl_Resize

End Sub

Private Sub UserControl_Resize()

  Dim I                 As Long
  Dim GHFactor          As Long
  Dim VisibleCols       As Long
  Dim VisibleColsWidth  As Long
  Const XPOffice2003Bug As Boolean = False

  'v2.6.3 This is for Ariel. change Const XPOffice2003Bug to True and the listview should work
  'Once we work out what the real cause of the bug is and build a detector this will be automatic
  'the name of the constant is temporary and may not be the real cause
  'which I suspect may be an internationalisation problem
  With grdPropertyList
    For I = 0 To .Cols - 1
      If Len(.TextArray(I)) = 0 Then
        'hide columns without headers
        'v2.6.3 XPOffice2003Bug
        .ColWidth(I) = IIf(XPOffice2003Bug, 0, -1)
       Else
        'count those with headers
        VisibleCols = VisibleCols + 1
      End If
    Next I
  End With
  picButtonContainer.Visible = mDefaultButtons
  '''-------------------------
  If VisibleCols Then
    If XPOffice2003Bug Then
      If VisibleCols Then
        grdPropertyList.ColWidth(eValue) = 900
        grdPropertyList.ColWidth(ePropName) = 5100
      End If
     Else
      With grdPropertyList
        VisibleColsWidth = (UserControl.ScaleWidth - ScrollOffset + VisibleCols * .GridLineWidth) / VisibleCols
        For I = 0 To .Cols - 1
          If Len(.TextArray(I)) Then
            .ColWidth(I) = VisibleColsWidth
          End If
        Next I
        If Len(strLongestOption) Then
          .ColWidth(eValue) = getTextWidth(strLongestOption) + .GridLineWidth * 14
          .ColWidth(ePropName) = .Width - .ColWidth(eValue)
        End If
      End With 'grdPropertyList
    End If
  End If
  '''-------------------------
  GHFactor = (grdPropertyList.Rows + 1) * ((grdPropertyList.CellHeight + grdPropertyList.GridLineWidth * 14))
  If GHFactor < 0 Then
    GHFactor = (grdPropertyList.CellHeight + grdPropertyList.GridLineWidth * 4)
  End If
  With UserControl
    .Height = GHFactor + IIf(bDescriptionShow, txtDescription.Height, 0) + IIf(mDefaultButtons, picButtonContainer.Height, 0)
    picButtonContainer.Move .ScaleLeft, .ScaleTop, .ScaleWidth
    grdPropertyList.Move .ScaleLeft, IIf(mDefaultButtons, picButtonContainer.Height, .ScaleTop), .ScaleWidth, GHFactor
  End With 'UserControl
  GHFactor = grdPropertyList.RowPos(grdPropertyList.Rows - 1) + grdPropertyList.CellHeight
  If bDescriptionShow Then
    txtDescription.Move UserControl.ScaleLeft, GHFactor - (grdPropertyList.Rows + 1) * grdPropertyList.GridLineWidth + IIf(mDefaultButtons, picButtonContainer.Height, 0), UserControl.ScaleWidth
  End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  With PropBag
    .WriteProperty "ForeColor", mForeColor, m_def_ForeColor
    .WriteProperty "ListColor", mListColor, m_def_ListColor
    .WriteProperty "ShowDesc", bDescriptionShow, True
    .WriteProperty "BackColor", m_BackColor, m_def_BackColor
    .WriteProperty "BackColorFixed", m_BackColorFixed, m_def_BackColorFixed
    .WriteProperty "BackColorSel", m_BackColorSel, m_def_BackColor
    .WriteProperty "Cols", grdPropertyList.Cols, 1
    .WriteProperty "Rows", grdPropertyList.Rows, 1
    .WriteProperty "FixedRows", grdPropertyList.FixedRows, 0
    .WriteProperty "Font", m_Font, Ambient.Font
    .WriteProperty "FormatString", grdPropertyList.FormatString, ""
    .WriteProperty "Locked", m_Locked, m_def_Locked
    .WriteProperty "DefList", DefaultList, ""
    .WriteProperty "defBut", DefaultButtons, False
  End With

End Sub

Private Function VisibleScrollBars(ControlName As Control) As Enum_VisibleScrollBars

  Dim MyStyle As Long

  '
  ' Returns an enumerated type constant depicting the
  ' type(s) of scrollbars visible on the passed control.
  '
  MyStyle = GetWindowLong(ControlName.hWnd, GWL_STYLE)
  'Use a bitwise comparison
  If (MyStyle And (WS_VSCROLL Or WS_HSCROLL)) = (WS_VSCROLL Or WS_HSCROLL) Then
    'Both are visible
    VisibleScrollBars = vs_both
   ElseIf (MyStyle And WS_VSCROLL) = WS_VSCROLL Then
    'Only Vertical is visible
    VisibleScrollBars = vs_vertical
   ElseIf (MyStyle And WS_HSCROLL) = WS_HSCROLL Then
    'Only Horizontal is visible
    VisibleScrollBars = vs_horizontal
   Else
    'No scrollbars are visible
    VisibleScrollBars = vs_none
  End If

End Function

':)Code Fixer V3.0.9 (25/03/2005 4:29:05 AM) 90 + 821 = 911 Lines Thanks Ulli for inspiration and lots of code.

