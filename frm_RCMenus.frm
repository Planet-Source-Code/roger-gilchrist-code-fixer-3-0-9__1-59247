VERSION 5.00
Begin VB.Form frm_RCMenus 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMenuWarning 
      Caption         =   "RCForm is not implemented as it doesn't work properly"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu mnuRCCode 
      Caption         =   "RCCode"
      Begin VB.Menu mnuRCSelectedText 
         Caption         =   "Selected Text"
         Begin VB.Menu mnuRCFind 
            Caption         =   "Find"
            Begin VB.Menu mnuRCFindOpt 
               Caption         =   "Project"
               Index           =   0
            End
            Begin VB.Menu mnuRCFindOpt 
               Caption         =   "Module"
               Index           =   1
            End
            Begin VB.Menu mnuRCFindOpt 
               Caption         =   "Procedure"
               Index           =   2
            End
         End
         Begin VB.Menu muRCSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRCCase 
            Caption         =   "Case"
            Begin VB.Menu mnuRCCaseOpt 
               Caption         =   "lower"
               Index           =   0
            End
            Begin VB.Menu mnuRCCaseOpt 
               Caption         =   "UPPER"
               Index           =   1
            End
            Begin VB.Menu mnuRCCaseOpt 
               Caption         =   "Proper"
               Index           =   2
            End
            Begin VB.Menu mnuRCCaseOpt 
               Caption         =   "Sentence"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRCPocedure 
         Caption         =   "Procedure"
         Begin VB.Menu mnuRCPocedureOpt 
            Caption         =   "Format"
            Index           =   0
            Begin VB.Menu mnuRCFormatOpt 
               Caption         =   "Apply"
               Index           =   0
            End
            Begin VB.Menu mnuRCFormatOpt 
               Caption         =   "Remove"
               Index           =   1
            End
            Begin VB.Menu mnuRCFormatOpt 
               Caption         =   "UnDo Last"
               Index           =   2
            End
         End
         Begin VB.Menu mnuRCPocedureOpt 
            Caption         =   "With"
            Index           =   1
            Begin VB.Menu mnuRCWithOpt 
               Caption         =   "Purify"
               Index           =   0
            End
            Begin VB.Menu mnuRCWithOpt 
               Caption         =   "Apply"
               Index           =   1
            End
            Begin VB.Menu mnuRCWithOpt 
               Caption         =   "Remove"
               Index           =   2
            End
         End
         Begin VB.Menu mnuRCPocedureOpt 
            Caption         =   "Bugtrap"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRCModule 
         Caption         =   "Module"
         Begin VB.Menu mnuRCModuleOpt 
            Caption         =   "Sort"
         End
      End
   End
   Begin VB.Menu mnuRCControl 
      Caption         =   "RCControl"
      Begin VB.Menu mnuRCFindInCode 
         Caption         =   "Find In Code"
      End
      Begin VB.Menu mnuRCHardCode 
         Caption         =   "Hard-Code"
         Begin VB.Menu mnuRCHardCodeOpt 
            Caption         =   "Control"
            Index           =   0
         End
         Begin VB.Menu mnuRCHardCodeOpt 
            Caption         =   "Form"
            Index           =   1
         End
         Begin VB.Menu mnuRCHardCodeOpt 
            Caption         =   "Project"
            Index           =   2
         End
      End
      Begin VB.Menu MnuHardString 
         Caption         =   "Hard Code Strings"
         Begin VB.Menu MnuHardStringOpt 
            Caption         =   "Form"
            Index           =   0
         End
         Begin VB.Menu MnuHardStringOpt 
            Caption         =   "Project"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSepCtrlForm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCFindInCodeFORM 
         Caption         =   "Find Form In Code"
      End
   End
End
Attribute VB_Name = "frm_RCMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  mnuRCCode.Visible = False
  mnuRCControl.Visible = False
  '  mnuRCForm.Visible = False
  On Error GoTo 0

End Sub

Private Sub MnuHardStringOpt_Click(Index As Integer)

  HardCodeObject Index + 1, True

End Sub

Private Sub mnuRCCaseOpt_Click(Index As Integer)

  DoRightCase Index

End Sub

Private Sub mnuRCCode_Click()

  'set menu appearance

  With mnuRCPocedure
    .Caption = IIf(InDeclaration, "Declarations", "Procedure")
    If .Caption = "Procedure" Then
      .Caption = "Procedure [" & getProcedureName & "]"
    End If
  End With 'mnuRCPocedure
  mnuRCSelectedText.Enabled = LenB(GetSelectedText)
  mnuRCPocedureOpt(1).Enabled = Not InDeclaration
  With mnuRCPocedureOpt(2)
    .Enabled = Not InDeclaration
    If .Enabled Then
      .Caption = IIf(HasBugTrap2(True), "Remove BugTrap", "Insert BugTrap")
     Else
      .Caption = "BugTrap"
    End If
  End With 'mnuRCPocedureOpt(2)
  mnuRCModule.Caption = "Module [" & GetActiveModuleName & "]"

End Sub

Private Sub mnuRCFindInCode_Click()

  DoRighClickControlsFinder

End Sub

Private Sub mnuRCFindInCodeFORM_Click()

  DoRighClickFormFinder

End Sub

Private Sub mnuRCFindOpt_Click(Index As Integer)

  ForceDisplayFind
  DoRightMenuFind Index

End Sub

Private Sub mnuRCFormatOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    RightClickIndent
   Case 1
    ProcNoIndent
   Case 2
    UnDoProcFormat
  End Select

End Sub

Private Sub mnuRCHardCodeOpt_Click(Index As Integer)

  HardCodeObject Index

End Sub

Private Sub mnuRCModuleOpt_Click()

  Dim CompMod         As CodeModule

  ReDim Attributes(1) As Variant
  Set CompMod = GetActiveCodePane.CodeModule
  SaveMemberAttributes 1, CompMod.Members
  DoSorting CompMod
  RestoreMemberAttributes 1, CompMod.Members

End Sub

Private Sub mnuRCPocedureOpt_Click(Index As Integer)

  If Index = 2 Then
    BugTrapAddRemoveToggle
  End If

End Sub

Private Sub mnuRCWithOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    RightMenuWithPurify
   Case 1
    RightMenuWithCreate
   Case 2
    RightMenuWithRemove
  End Select

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:28:15 AM) 1 + 119 = 120 Lines Thanks Ulli for inspiration and lots of code.
