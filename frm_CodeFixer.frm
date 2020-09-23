VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CodeFixer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Roja's Code Fixer"
   ClientHeight    =   7695
   ClientLeft      =   4095
   ClientTop       =   5325
   ClientWidth     =   11580
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frm_CodeFixer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11580
   Begin VB.PictureBox picMenuIcon 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   7560
      Picture         =   "frm_CodeFixer.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   76
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame frapage 
      Caption         =   "Reload"
      Height          =   4935
      Index           =   3
      Left            =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox picCFXPBugFix 
         BorderStyle     =   0  'None
         Height          =   4575
         Index           =   2
         Left            =   0
         ScaleHeight     =   4575
         ScaleWidth      =   6015
         TabIndex        =   1
         Top             =   0
         Width           =   6015
         Begin VB.Frame fraUndoTab 
            Caption         =   "Reload"
            Height          =   4095
            Index           =   0
            Left            =   4800
            TabIndex        =   5
            Top             =   1560
            Visible         =   0   'False
            Width           =   5775
            Begin VB.PictureBox picCFXPBugFix 
               BorderStyle     =   0  'None
               Height          =   3845
               Index           =   11
               Left            =   100
               ScaleHeight     =   3840
               ScaleWidth      =   5580
               TabIndex        =   10
               Top             =   175
               Width           =   5575
               Begin MSComctlLib.ListView lsvUnDone 
                  Height          =   3255
                  Left            =   0
                  TabIndex        =   77
                  Top             =   0
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   5741
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Key             =   "project"
                     Text            =   "Project"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Key             =   "module"
                     Text            =   "Module"
                     Object.Width           =   2540
                  EndProperty
               End
               Begin VB.CommandButton cmdReload 
                  Caption         =   "Selected"
                  Height          =   375
                  Index           =   0
                  Left            =   15
                  TabIndex        =   12
                  Top             =   3400
                  Width           =   855
               End
               Begin VB.CommandButton cmdReload 
                  Caption         =   "All"
                  Height          =   375
                  Index           =   1
                  Left            =   1580
                  TabIndex        =   11
                  Top             =   3400
                  Width           =   855
               End
               Begin VB.Label lblNoSavedProj 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Restricted Access"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   3615
                  Left            =   2760
                  TabIndex        =   27
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2775
               End
               Begin VB.Label lblReload 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Label4"
                  Height          =   3735
                  Left            =   2775
                  TabIndex        =   13
                  Top             =   45
                  Width           =   2760
               End
            End
         End
         Begin VB.Frame fraUndoTab 
            Caption         =   "singlefolder"
            Height          =   4215
            Index           =   2
            Left            =   1560
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   5895
            Begin VB.PictureBox picCFXPBugFix 
               BorderStyle     =   0  'None
               Height          =   3845
               Index           =   10
               Left            =   100
               ScaleHeight     =   3840
               ScaleWidth      =   5580
               TabIndex        =   6
               Top             =   175
               Width           =   5575
               Begin VB.CommandButton cmdSingleFolder 
                  Caption         =   "Convert Source"
                  Height          =   375
                  Index           =   1
                  Left            =   3840
                  TabIndex        =   8
                  Top             =   3360
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSingleFolder 
                  Caption         =   "Create Release"
                  Height          =   375
                  Index           =   0
                  Left            =   120
                  TabIndex        =   7
                  Top             =   3360
                  Width           =   1575
               End
               Begin VB.Label lblsinglefolder 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Label4"
                  Height          =   3255
                  Left            =   0
                  TabIndex        =   9
                  Top             =   45
                  Width           =   5535
               End
            End
         End
         Begin VB.Frame fraUndoTab 
            Caption         =   "BackUp"
            Height          =   4095
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   5775
            Begin VB.PictureBox picCFXPBugFix 
               BorderStyle     =   0  'None
               Height          =   3845
               Index           =   12
               Left            =   100
               ScaleHeight     =   3840
               ScaleWidth      =   5580
               TabIndex        =   14
               Top             =   175
               Width           =   5575
               Begin VB.CheckBox chkOdlBackWarn 
                  Caption         =   "Old Backup Warning"
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   28
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.Frame fraDeletebackup 
                  Caption         =   "Delete Backups"
                  Height          =   1095
                  Left            =   3980
                  TabIndex        =   20
                  Top             =   640
                  Width           =   1455
                  Begin VB.PictureBox picCFXPBugFix 
                     BorderStyle     =   0  'None
                     Height          =   845
                     Index           =   13
                     Left            =   100
                     ScaleHeight     =   840
                     ScaleWidth      =   1260
                     TabIndex        =   23
                     Top             =   175
                     Width           =   1255
                     Begin VB.CommandButton cmdBackup 
                        Caption         =   "Not Selected"
                        Height          =   255
                        Index           =   3
                        Left            =   20
                        TabIndex        =   26
                        Top             =   295
                        Width           =   1215
                     End
                     Begin VB.CommandButton cmdBackup 
                        Caption         =   "Selected"
                        Height          =   255
                        Index           =   2
                        Left            =   20
                        TabIndex        =   25
                        Top             =   40
                        Width           =   1215
                     End
                     Begin VB.CommandButton cmdBackup 
                        Caption         =   "All"
                        Height          =   90
                        Index           =   1
                        Left            =   54900
                        TabIndex        =   24
                        Top             =   9195
                        Width           =   2.45745e5
                     End
                  End
               End
               Begin VB.CommandButton cmdBackup 
                  Caption         =   "Make Backup"
                  Height          =   255
                  Index           =   0
                  Left            =   3980
                  TabIndex        =   19
                  Top             =   280
                  Width           =   1455
               End
               Begin VB.CommandButton cmdRestore 
                  Caption         =   "Restore All"
                  Height          =   375
                  Index           =   1
                  Left            =   3260
                  TabIndex        =   18
                  Top             =   3400
                  Width           =   1575
               End
               Begin VB.ListBox LstRestoreFiles 
                  Height          =   1185
                  Left            =   480
                  Style           =   1  'Checkbox
                  TabIndex        =   17
                  Top             =   2080
                  Width           =   4335
               End
               Begin VB.CommandButton cmdRestore 
                  Caption         =   "Restore Selected"
                  Height          =   375
                  Index           =   0
                  Left            =   500
                  TabIndex        =   16
                  Top             =   3400
                  Width           =   1575
               End
               Begin VB.ListBox lstRestore 
                  Height          =   1425
                  Left            =   0
                  Sorted          =   -1  'True
                  TabIndex        =   15
                  Top             =   240
                  Width           =   3000
               End
               Begin VB.Label lblFiles 
                  Caption         =   "files(s)"
                  Height          =   255
                  Left            =   500
                  TabIndex        =   22
                  Top             =   1840
                  Width           =   3255
               End
               Begin VB.Label lblBackupcount 
                  Caption         =   "Backup(s)"
                  Height          =   255
                  Left            =   20
                  TabIndex        =   21
                  Top             =   40
                  Width           =   3255
               End
            End
         End
         Begin MSComctlLib.TabStrip tbsFileTool 
            Height          =   4575
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   8070
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Reload"
                  Key             =   "reload"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "BackUp && Restore"
                  Key             =   "backupnrestore"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Single Folder"
                  Key             =   "singlefolder"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame frapage 
      Caption         =   "Modules"
      Height          =   4695
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   6375
      Begin VB.PictureBox picCFXPBugFixfCodeFix 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   4440
         Index           =   2
         Left            =   0
         ScaleHeight     =   4440
         ScaleWidth      =   6300
         TabIndex        =   46
         Top             =   0
         Width           =   6295
         Begin VB.Frame fraModule 
            Caption         =   "Module(s)"
            Height          =   3855
            Index           =   1
            Left            =   2520
            TabIndex        =   59
            Top             =   960
            Visible         =   0   'False
            Width           =   6135
            Begin VB.PictureBox picCFXPBugFixfCodeFix 
               BorderStyle     =   0  'None
               Height          =   3605
               Index           =   0
               Left            =   100
               ScaleHeight     =   3600
               ScaleWidth      =   5940
               TabIndex        =   60
               Top             =   175
               Width           =   5935
               Begin VB.CommandButton cmdEditMod 
                  Caption         =   "6"
                  BeginProperty Font 
                     Name            =   "Marlett"
                     Size            =   14.25
                     Charset         =   2
                     Weight          =   500
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   3
                  Left            =   3240
                  TabIndex        =   68
                  ToolTipText     =   "Set Filename to Module Name"
                  Top             =   2520
                  Width           =   270
               End
               Begin VB.CommandButton cmdEditMod 
                  Caption         =   "Change Filename"
                  Height          =   255
                  Index           =   1
                  Left            =   20
                  TabIndex        =   67
                  Top             =   2110
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEditMod 
                  Caption         =   "Change Name"
                  Height          =   255
                  Index           =   0
                  Left            =   20
                  TabIndex        =   66
                  Top             =   3280
                  Width           =   1455
               End
               Begin VB.TextBox txtModuleEdit 
                  Height          =   285
                  Index           =   0
                  Left            =   1460
                  TabIndex        =   65
                  Top             =   3250
                  Width           =   2055
               End
               Begin VB.TextBox txtModuleEdit 
                  Height          =   285
                  Index           =   1
                  Left            =   1460
                  TabIndex        =   64
                  Top             =   2110
                  Width           =   2055
               End
               Begin VB.CommandButton cmdEditMod 
                  Caption         =   "5"
                  BeginProperty Font 
                     Name            =   "Marlett"
                     Size            =   14.25
                     Charset         =   2
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   1455
                  TabIndex        =   63
                  ToolTipText     =   "Set Module name to Filename"
                  Top             =   3000
                  Width           =   285
               End
               Begin VB.ListBox lstSuggestModName 
                  Height          =   1425
                  Left            =   3740
                  Sorted          =   -1  'True
                  TabIndex        =   62
                  Top             =   2110
                  Width           =   2055
               End
               Begin VB.TextBox txtModuleEdit 
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Index           =   2
                  Left            =   20
                  TabIndex        =   61
                  Top             =   1720
                  Width           =   5775
               End
               Begin MSComctlLib.ListView lsvAllModules 
                  Height          =   1575
                  Left            =   20
                  TabIndex        =   69
                  Top             =   40
                  Width           =   5775
                  _ExtentX        =   10186
                  _ExtentY        =   2778
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   5
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Key             =   "project"
                     Text            =   "Projects"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Key             =   "module"
                     Text            =   "Modules"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Key             =   "filename"
                     Text            =   "Filename"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Key             =   "hiddenindex"
                     Text            =   "hidden"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   4
                     Key             =   "comment"
                     Text            =   "Comments"
                     Object.Width           =   2540
                  EndProperty
               End
               Begin VB.Label lblReadOnlyMod 
                  Caption         =   "File is Read-Only"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   2640
                  Visible         =   0   'False
                  Width           =   2775
               End
            End
         End
         Begin VB.Frame fraModule 
            Caption         =   "Projects"
            Height          =   3855
            Index           =   2
            Left            =   0
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   6135
            Begin VB.PictureBox picCFXPBugFixfCodeFix 
               BorderStyle     =   0  'None
               Height          =   3605
               Index           =   1
               Left            =   100
               ScaleHeight     =   3600
               ScaleWidth      =   5940
               TabIndex        =   49
               Top             =   175
               Width           =   5935
               Begin VB.CommandButton cmdEditProj 
                  Caption         =   "6"
                  BeginProperty Font 
                     Name            =   "Marlett"
                     Size            =   14.25
                     Charset         =   2
                     Weight          =   500
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   3
                  Left            =   3375
                  TabIndex        =   57
                  ToolTipText     =   "Set Filename to Project name"
                  Top             =   2445
                  Width           =   270
               End
               Begin VB.TextBox txtProjectEdit 
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Index           =   2
                  Left            =   20
                  TabIndex        =   56
                  Top             =   1720
                  Width           =   5775
               End
               Begin VB.ListBox lstSuggestProjName 
                  Height          =   1425
                  Left            =   3740
                  Sorted          =   -1  'True
                  TabIndex        =   55
                  Top             =   2080
                  Width           =   2055
               End
               Begin VB.CommandButton cmdEditProj 
                  Caption         =   "5"
                  BeginProperty Font 
                     Name            =   "Marlett"
                     Size            =   14.25
                     Charset         =   2
                     Weight          =   500
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   1455
                  TabIndex        =   54
                  ToolTipText     =   "Set Project name to Filename"
                  Top             =   2940
                  Width           =   270
               End
               Begin VB.TextBox txtProjectEdit 
                  Height          =   285
                  Index           =   1
                  Left            =   1460
                  TabIndex        =   53
                  Top             =   2080
                  Width           =   2175
               End
               Begin VB.TextBox txtProjectEdit 
                  Height          =   285
                  Index           =   0
                  Left            =   1460
                  TabIndex        =   52
                  Top             =   3235
                  Width           =   2175
               End
               Begin VB.CommandButton cmdEditProj 
                  Caption         =   "Change Filename"
                  Height          =   255
                  Index           =   1
                  Left            =   20
                  TabIndex        =   51
                  Top             =   2080
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEditProj 
                  Caption         =   "Change Name"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   50
                  Top             =   3265
                  Width           =   1455
               End
               Begin MSComctlLib.ListView lsvAllProjects 
                  Height          =   1575
                  Left            =   20
                  TabIndex        =   58
                  Top             =   40
                  Width           =   5775
                  _ExtentX        =   10186
                  _ExtentY        =   2778
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   4
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Key             =   "project"
                     Text            =   "Project(s)"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Key             =   "filename"
                     Text            =   "Filename(s)"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Key             =   "hiddenindex"
                     Text            =   "hidden"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Key             =   "comments"
                     Text            =   "Comments"
                     Object.Width           =   2540
                  EndProperty
               End
               Begin VB.Label lblReadOnlyPrj 
                  Caption         =   "File is Read-Only"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   72
                  Top             =   2580
                  Visible         =   0   'False
                  Width           =   2775
               End
            End
         End
         Begin MSComctlLib.TabStrip tbsModule 
            Height          =   4335
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   7646
            MultiRow        =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Module(s)"
                  Key             =   "module"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Project(s)"
                  Key             =   "project"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame frapage 
      Caption         =   "Controls"
      Height          =   6015
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   2760
      Width           =   7335
      Begin VB.PictureBox picCFXPBugFix 
         BorderStyle     =   0  'None
         Height          =   5520
         Index           =   14
         Left            =   120
         ScaleHeight     =   5520
         ScaleWidth      =   7020
         TabIndex        =   30
         Top             =   240
         Width           =   7015
         Begin VB.CommandButton cmdFindInCode 
            Caption         =   "Find In Code"
            Height          =   255
            Left            =   4920
            TabIndex        =   79
            Top             =   3795
            Width           =   2055
         End
         Begin VB.CommandButton cmdXPStyle 
            Caption         =   "Remove XP Style"
            Height          =   255
            Left            =   5520
            TabIndex        =   78
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "Delete No Code Control"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   74
            ToolTipText     =   "RECOMMENDED do not use if poorly named controls exist (fix them first)"
            Top             =   4560
            Width           =   2175
         End
         Begin VB.CheckBox chkDelOldCode 
            Caption         =   "Delete old code when making arrays"
            Height          =   195
            Left            =   0
            TabIndex        =   73
            Top             =   4920
            Width           =   2895
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "Auto Fix Sinlgeton Array"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   70
            ToolTipText     =   "Does not touch seed controls for code loaded control arrays"
            Top             =   3390
            Width           =   2175
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "XP Frame Bug Fix"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   44
            Top             =   3645
            Width           =   2175
         End
         Begin MSComctlLib.ListView lsvAllControls 
            Height          =   2775
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "proj"
               Text            =   "Project"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "component"
               Text            =   "Component"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "control"
               Text            =   "Control"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "captions"
               Text            =   "Caption"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "use"
               Text            =   "Comment(s)"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "Refresh List"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "Auto Prefix Controls"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "RECOMMENDED do not use if poorly named controls exist (fix them first)"
            Top             =   3900
            Width           =   2175
         End
         Begin VB.ListBox lstPrefixSuggest 
            Height          =   1815
            Left            =   2760
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   3120
            Width           =   2055
         End
         Begin VB.CommandButton cmdAutoLabel 
            Caption         =   "Auto Fix Poorly Named"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Controls with Captions and Menus with Reserved Word/VBCommand names can be Auto Fixed"
            Top             =   3135
            Width           =   2175
         End
         Begin VB.CommandButton cmdCtrlChange 
            Caption         =   "Apply New Name"
            Height          =   255
            Left            =   4940
            TabIndex        =   32
            Top             =   4725
            Width           =   2055
         End
         Begin VB.TextBox txtCtrlNewName 
            Height          =   375
            Left            =   4940
            TabIndex        =   31
            Top             =   4365
            Width           =   2055
         End
         Begin VB.Label lblDeletableExist 
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   0
            TabIndex        =   75
            ToolTipText     =   "Red means deletable control(s) exist (button enabled if current selected control is deletable)"
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Label lblOldName 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   40
            Top             =   5160
            Width           =   7095
         End
         Begin VB.Label lblOldName 
            Caption         =   "Prefix Suggestions"
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   39
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label lblOldName 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   38
            Top             =   3960
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblOldName 
            Caption         =   "New Name"
            Height          =   255
            Index           =   2
            Left            =   4940
            TabIndex        =   37
            Top             =   4125
            Width           =   2055
         End
         Begin VB.Label lblOldName 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   4935
            TabIndex        =   36
            Top             =   3540
            Width           =   2055
         End
         Begin VB.Label lblOldName 
            Caption         =   "Current Name"
            Height          =   255
            Index           =   0
            Left            =   4935
            TabIndex        =   35
            Top             =   3285
            Width           =   2055
         End
      End
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Begin VB.Menu mnuDeleteOpt 
         Caption         =   "Matching Current Selection"
         Index           =   0
      End
      Begin VB.Menu mnuDeleteOpt 
         Caption         =   "Listed Code Fixer Comments"
         Index           =   1
      End
      Begin VB.Menu mnuDeleteOpt 
         Caption         =   "All Code Fixer Comments"
         Index           =   2
      End
   End
   Begin VB.Menu mnuFindShow 
      Caption         =   "View"
      Begin VB.Menu mnuViewOpt 
         Caption         =   "Columns..."
         Index           =   0
         Begin VB.Menu mnuColOpt 
            Caption         =   "Project"
            Index           =   0
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Component"
            Index           =   1
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Component line"
            Index           =   2
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Procedure"
            Index           =   3
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Procedure line"
            Index           =   4
         End
      End
      Begin VB.Menu mnuViewOpt 
         Caption         =   "Sort..."
         Index           =   1
         Begin VB.Menu mnuViewSortOpt 
            Caption         =   "Unsorted"
            Index           =   0
         End
         Begin VB.Menu mnuViewSortOpt 
            Caption         =   "Alpha (Code)"
            Index           =   1
         End
         Begin VB.Menu mnuViewSortOpt 
            Caption         =   "Module Order(Form)"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuPopUpDelete 
      Caption         =   "popDelete"
      Begin VB.Menu mnuCodeFixMarkers 
         Caption         =   "Code Fixer Messages"
         Index           =   0
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Match Selection"
            Index           =   1
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Like Selection"
            Index           =   2
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Missing Dims"
            Index           =   4
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Unused"
            Index           =   5
            Begin VB.Menu mnuUnused 
               Caption         =   "All"
               Index           =   0
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Variable"
               Index           =   1
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Dim"
               Index           =   2
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Declare"
               Index           =   3
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Type"
               Index           =   4
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Enum"
               Index           =   5
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Sub"
               Index           =   6
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Function"
               Index           =   7
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Property"
               Index           =   8
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Parameter"
               Index           =   9
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Dead Code"
               Index           =   10
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Empty Structure"
               Index           =   11
            End
            Begin VB.Menu mnuUnused 
               Caption         =   "Unnecessary"
               Index           =   12
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Untyped"
            Index           =   6
            Begin VB.Menu mnuUntyped 
               Caption         =   "All"
               Index           =   0
            End
            Begin VB.Menu mnuUntyped 
               Caption         =   "Dim"
               Index           =   1
            End
            Begin VB.Menu mnuUntyped 
               Caption         =   "Variable"
               Index           =   2
            End
            Begin VB.Menu mnuUntyped 
               Caption         =   "Constant"
               Index           =   3
            End
            Begin VB.Menu mnuUntyped 
               Caption         =   "Parameter"
               Index           =   4
            End
            Begin VB.Menu mnuUntyped 
               Caption         =   "Function"
               Index           =   5
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "With Structures"
            Index           =   7
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Auto-Inserted "
               Index           =   0
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Start of With"
               Index           =   1
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "End of With"
               Index           =   2
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Possible Approval points"
               Index           =   3
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Approved Structures Done"
               Index           =   4
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Excess With Object"
               Index           =   5
            End
            Begin VB.Menu mnuWithOpt 
               Caption         =   "Purified With Object"
               Index           =   6
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Procedures/Parameters"
            Index           =   8
            Begin VB.Menu mnuParam 
               Caption         =   "Unused"
               Index           =   0
            End
            Begin VB.Menu mnuParam 
               Caption         =   "Untyped"
               Index           =   1
            End
            Begin VB.Menu mnuParam 
               Caption         =   "ByVal Suggest"
               Index           =   2
            End
            Begin VB.Menu mnuParam 
               Caption         =   "ByVal Inserted"
               Index           =   3
            End
            Begin VB.Menu mnuParam 
               Caption         =   "Function to Sub"
               Index           =   4
            End
            Begin VB.Menu mnuParam 
               Caption         =   "Function Auto-Typed"
               Index           =   5
            End
            Begin VB.Menu mnuParam 
               Caption         =   "Large Code procedure"
               Index           =   6
            End
            Begin VB.Menu mnuParam 
               Caption         =   "Large Control Procedure"
               Index           =   7
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Goto"
            Index           =   9
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Orphan Target Label"
               Index           =   0
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "No Target Label"
               Index           =   1
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Unneeded Goto 0"
               Index           =   2
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Goto => Exit Proc"
               Index           =   3
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Goto Label => Exit  Proc"
               Index           =   4
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Illegal Goto Into"
               Index           =   5
            End
            Begin VB.Menu mnuGotoOpt 
               Caption         =   "Illegal Goto out of"
               Index           =   6
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "If  structures"
            Index           =   10
            Begin VB.Menu mnuElseOpt 
               Caption         =   "All"
               Index           =   0
            End
            Begin VB.Menu mnuElseOpt 
               Caption         =   "Case Conversion"
               Index           =   1
            End
            Begin VB.Menu mnuElseOpt 
               Caption         =   "Missing Else"
               Index           =   2
            End
            Begin VB.Menu mnuElseOpt 
               Caption         =   "If..And..End If Short Circuit"
               Index           =   3
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Structural Changes"
            Index           =   11
            Begin VB.Menu mnuStructural 
               Caption         =   "Structure Error"
               Index           =   0
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Scope Change"
               Index           =   1
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Scope Too Large"
               Index           =   2
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Integer to Long "
               Index           =   3
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Duplicated Name"
               Index           =   4
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Close 'On Error Resume'"
               Index           =   5
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Active Debug"
               Index           =   6
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Structure Expanded"
               Index           =   7
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Pleonasm Removed"
               Index           =   8
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "For-Variable Inserted"
               Index           =   9
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Zero String update"
               Index           =   10
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Unsafe Load/Unload"
               Index           =   11
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Hard Coded Path"
               Index           =   12
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "DefType"
               Index           =   13
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Goto Line Number"
               Index           =   14
            End
            Begin VB.Menu mnuStructural 
               Caption         =   "Obsolete Code"
               Index           =   15
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Controls"
            Index           =   12
            Begin VB.Menu mnuControlOpt 
               Caption         =   "Renamed"
               Index           =   0
            End
            Begin VB.Menu mnuControlOpt 
               Caption         =   "Removed control code"
               Index           =   1
            End
            Begin VB.Menu mnuControlOpt 
               Caption         =   "XP Style && Frame Bug"
               Index           =   2
            End
            Begin VB.Menu mnuControlOpt 
               Caption         =   "Default Property"
               Index           =   3
            End
            Begin VB.Menu mnuControlOpt 
               Caption         =   "Default Me/UserControl"
               Index           =   4
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "General message types"
            Index           =   13
            Begin VB.Menu mnuGeneralOpt 
               Caption         =   "UPDATED"
               Index           =   0
            End
            Begin VB.Menu mnuGeneralOpt 
               Caption         =   "SUGGESTION"
               Index           =   1
            End
            Begin VB.Menu mnuGeneralOpt 
               Caption         =   "WARNING"
               Index           =   2
            End
            Begin VB.Menu mnuGeneralOpt 
               Caption         =   "RISK"
               Index           =   3
            End
            Begin VB.Menu mnuGeneralOpt 
               Caption         =   "PREVIOUS CODE"
               Index           =   4
            End
         End
         Begin VB.Menu mnuCodeFixOpt 
            Caption         =   "Usage"
            Index           =   14
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "All"
               Index           =   0
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Header"
               Index           =   1
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Move suggested"
               Index           =   2
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Form to Code Entry point"
               Index           =   3
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "NOT USED"
               Index           =   4
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Form/Module Interface"
               Index           =   5
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Form/Module Variable"
               Index           =   6
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Module Level > Procedure level"
               Index           =   7
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Module/Module Interface"
               Index           =   8
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "RECOMMENDED"
               Index           =   9
            End
            Begin VB.Menu mnuCodeFixUsageOpt 
               Caption         =   "Dim Usage"
               Index           =   10
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Dim Usage (all)"
                  Index           =   0
               End
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Dim Usage (1)"
                  Index           =   1
               End
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Dim Usage (2)"
                  Index           =   2
               End
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Dim Usage (3)"
                  Index           =   3
               End
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Dim Usage (3) on (2)"
                  Index           =   4
               End
               Begin VB.Menu mnuDimusageOpt 
                  Caption         =   "Delete non-problematic"
                  Index           =   5
               End
            End
         End
      End
      Begin VB.Menu mnuCodeFixMarkers 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopUp_0 
         Caption         =   "Code Fixer Delete..."
         Begin VB.Menu mnuPopUpDeleteOpt 
            Caption         =   "Delete Match Selection"
            Index           =   0
         End
         Begin VB.Menu mnuPopUpDeleteOpt 
            Caption         =   "Delete Like Selection"
            Index           =   1
         End
         Begin VB.Menu mnuPopUpDeleteOpt 
            Caption         =   "Delete Listed CF Comments"
            Index           =   2
         End
         Begin VB.Menu mnuPopUpDeleteOpt 
            Caption         =   "Delete All CF Comments"
            Index           =   3
         End
         Begin VB.Menu mnuPopUpDeleteOpt 
            Caption         =   "Remove Code Fixer Tag line"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuPopControls 
      Caption         =   "popControls"
      Begin VB.Menu mnupopControlLev1 
         Caption         =   "Columns..."
         Index           =   0
         Begin VB.Menu mnupopControlLev2 
            Caption         =   "Project"
            Index           =   0
         End
         Begin VB.Menu mnupopControlLev2 
            Caption         =   "Component"
            Index           =   1
         End
         Begin VB.Menu mnupopControlLev2 
            Caption         =   "Caption"
            Index           =   2
         End
      End
      Begin VB.Menu mnupopControlLev1 
         Caption         =   "Sort..."
         Index           =   1
         Begin VB.Menu mnupopControlLev3 
            Caption         =   "Project"
            Index           =   0
         End
         Begin VB.Menu mnupopControlLev3 
            Caption         =   "Component"
            Index           =   1
         End
         Begin VB.Menu mnupopControlLev3 
            Caption         =   "Control"
            Index           =   2
         End
         Begin VB.Menu mnupopControlLev3 
            Caption         =   "Caption"
            Index           =   3
         End
         Begin VB.Menu mnupopControlLev3 
            Caption         =   "Comment"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frm_CodeFixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright 2003 Roger Gilchrist
'email: rojagilkrist@hotmail.com
' some of this code is based on Ulli's Code Formatter
'© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
Option Explicit
Private PrevTab     As Long

Public Sub ActivatePage()

  'v2.3.7 fixed Control tool not launching properly
  ' TabPreActivate

  frm_CodeFixer.WindowState = vbNormal
  'v2.7.3 this causes a problem if Control tool is active when Settings is called
  If FrameActive <> TPNone Then
    frapage(FrameActive).Visible = True
  End If
  frm_CodeFixer.Refresh
  PlaceTool
  'TabPostActivate

End Sub

Private Function AnyRenamed() As Boolean

  Dim I As Long

  For I = LBound(CntrlDesc) To UBound(CntrlDesc)
    If CntrlDesc(I).CDUsage <> 2 Then
      If CntrlDesc(I).CDOldName <> CntrlDesc(I).CDName Then
        AnyRenamed = True
        Exit For
      End If
    End If
  Next I

End Function

Private Sub cmdAutoLabel_Click(Index As Integer)

  ControlAutoEnabled False
  AutoReName = True
  ControlRenameButton False
  Select Case Index
   Case 0
    AutoFixBadNames
   Case 1
    If AnyBadNames Then
      If mObjDoc.Safe_MsgBox("Auto Prefix will apply a prefix to all members of the Poorly Named Controls list." & vbNewLine & _
                       "Controls with VB Default names will no longer be detected but have ugly names ('Test1' => 'txtText1')" & vbNewLine & _
                       "Proceed anyway?", vbInformation + vbYesNo) = vbYes Then
        AutoFixPrefix
       Else
        WarningLabel
        GoTo UpdatedAlready
      End If
      GoTo ExitAuto
    End If
    AutoFixPrefix
   Case 2
    UpdateCtrlnameLists
    GoTo UpdatedAlready
   Case 3
    If AnyRenamed Then
      If vbOK = mObjDoc.Safe_MsgBox("You have renamed some controls." & vbNewLine & _
                              "XP Frame Bug fix requires that you click 'Refresh List' first." & vbNewLine & _
                              "Select 'OK' to do so automatically or 'Cancel' to abort the fix", vbInformation + vbOKCancel) Then
        UpdateCtrlnameLists
        InsertXPPic2Frame2
       Else
        GoTo UpdatedAlready
      End If
     Else
      InsertXPPic2Frame2
    End If
   Case 4
    AutoFixSingletonControlArrays
   Case 5 'delete
    If Not DeleteControl Then
      GoTo UpdatedAlready
    End If
  End Select
ExitAuto:
  ControlAutoEnabled True
  UpdateCtrlnameLists
UpdatedAlready:
  AutoReName = False
  UpDateOldNameIndex
  ControlRenameButton True

End Sub

Private Sub cmdBackup_Click(Index As Integer)

  Select Case Index
   Case 0
    cmdBackup(0).Enabled = False
    BackUpMakeOne
    cmdBackup(0).Enabled = True
   Case 1
    BackUpDeleteSelected True, True
   Case 2
    BackUpDeleteSelected True
   Case 3
    BackUpDeleteSelected False
  End Select

End Sub

Private Sub cmdCtrlChange_Click()

  DO_cmdCtrlChangeClick

End Sub

Private Sub cmdEditMod_Click(Index As Integer)

  DO_cmdEditModClick Index

End Sub

Private Sub cmdEditProj_Click(Index As Integer)

  DO_cmdEditProjClick Index

End Sub

Private Sub cmdFindInCode_Click()

  mObjDoc.ForceFind lblOldName(1).Caption

End Sub

Private Sub cmdReload_Click(Index As Integer)

  Dim I     As Long
  Dim Comp  As VBComponent
  Dim LItem As ListItem

  With lsvUnDone
    For I = 1 To .ListItems.Count
      SetCurrentLSVLine lsvUnDone, I
      'safety stuff
      Set LItem = .SelectedItem
      If LItem.Checked Or Index = 1 Then
        Set Comp = GetComponent(LItem.Text, LItem.SubItems(1))
        If Comp.IsDirty Then
          If Not IsFileReadOnly(Comp.FileNames(1)) Then
            If IsComponent_Reloadable(Comp) Then
              Comp.Reload
            End If
          End If
          LItem.Checked = False
        End If
      End If
    Next I
  End With

End Sub

Private Sub cmdRestore_Click(Index As Integer)

  RestoreBackup Index, lstRestore.Text

End Sub

Private Sub cmdSingleFolder_Click(Index As Integer)

  SingleFolderWrap Index
  cmdSingleFolder(1).Enabled = Not IsProjectInSingleFolder
  Unload Me

End Sub

Private Sub cmdXPStyle_Click()

  DoXPStyle
  cmdXPStyle.Caption = XPStyleCaption

End Sub

Private Sub DO_cmdCtrlChangeClick()

  Dim MyhourGlass As cls_HourGlass

  Set MyhourGlass = New cls_HourGlass
  ControlRenameButton False
  If Len(txtCtrlNewName.Text) Then
    If Not AutoReName Then
      WarningLabel "Working...", vbRed
    End If
    EditControlName txtCtrlNewName.Text
    DoBadSelect
    If Not AutoReName Then
      ControlAutoEnabled False
    End If
  End If
  ControlRenameButton True

End Sub

Private Sub DO_cmdEditModClick(ByVal intIndex As Long)

  Dim MyhourGlass As cls_HourGlass
  Dim LItem       As ListItem
  Dim Strnew      As String
  Dim Comp        As VBComponent
  Dim Proj        As VBProject
  Dim strOldName  As String
  Dim strNewPath  As String
  Dim I           As Long

  Set MyhourGlass = New cls_HourGlass
  For I = 0 To 3
    cmdEditMod(I).Enabled = False
  Next I
  Set LItem = lsvAllModules.SelectedItem
  Set Proj = VBInstance.VBProjects.Item(LItem.Text)
  Set Comp = Proj.VBComponents.Item(LItem.SubItems(1))
  Select Case intIndex
   Case 0 'module name
    Strnew = txtModuleEdit(0).Text
    If Strnew <> LItem.SubItems(1) Then
      If LegalName(Strnew, strNewPath, 0, Comp, Proj) Then
        ReplaceName LItem.SubItems(1), Strnew
        LItem.SubItems(1) = Strnew
        'v2.4.6 Thanks Mike Ulik Found the need for this error trap while working on the other bug
        On Error Resume Next
        Comp.Name = Strnew
        If Err.Number = 32813 Then
          mObjDoc.Safe_MsgBox "A Module or Project with the name" & strInSQuotes(Strnew, True) & "already exists.", vbCritical
         Else
          'v2.4.6 Thanks Mike Ulik. This was the bug you reported
          'VB has a bug that doesn't recognize case changes in the module name
          'this trick changes it to something extremely unlikely( its reverse)
          'but with a letter added in case the newname ends in a numeral or other illegal 1st character
          'then back to the new casing which makes the name stick
          With Comp
            If LCase$(.Name) = LCase$(Strnew) Then
              .Name = "a" & StrReverse(Strnew)
              .Name = Strnew
            End If
          End With 'Comp
          Comp.SaveAs Comp.FileNames(1)
          Comp.Reload
        End If
        On Error GoTo 0
        '   End If
      End If
    End If
   Case 1 'module file
    Strnew = txtModuleEdit(1).Text
    If Strnew <> LItem.SubItems(2) Then
      strOldName = Comp.FileNames(1)
      If LegalName(Strnew, strNewPath, 1, Comp, Proj) Then
        If LItem.SubItems(2) = strUnsavedModule Then
          Comp.SaveAs FilePathOnly(LItem.Text) & "\" & Strnew
          FileKill strOldName ' needed??
         Else
          With Comp
            ReDim killus(.FileCount) As String ' get all potential attached files(FRX etc)
            For I = 1 To .FileCount
              killus(I) = .FileNames(I)
            Next I
          End With 'Comp
          'NOTE this will also generate the other support files automatically
          Comp.SaveAs Replace$(Comp.FileNames(1), LItem.SubItems(2), Strnew)
          If IsComponent_Reloadable(Comp) Then
            Comp.Reload
          End If
          For I = LBound(killus) To UBound(killus) 'Kill the old files
            FileKill killus(I)
          Next I
        End If
      End If
    End If
   Case 2 'filename from module name
    'Call this routine recursively to do the filename
    DO_cmdEditModClick 0
    'reset the filename and call this routine recursively
    txtModuleEdit(1).Text = txtModuleEdit(0).Text & Mid$(txtModuleEdit(1).Text, InStr(txtModuleEdit(1).Text, "."))
    DO_cmdEditModClick 1
   Case 3 'sync module file to mod name
    'Call this routine recursively to do the filename
    DO_cmdEditModClick 1
    'reset the modulename and call this routine recursively
    txtModuleEdit(0).Text = Left$(txtModuleEdit(1).Text, InStr(txtModuleEdit(1).Text, ".") - 1)
    DO_cmdEditModClick 0
  End Select
  'keep the vbp file up to date
  Proj.SaveAs Proj.FileName
  On Error Resume Next
  Comp.Reload
  On Error GoTo 0
  UpdateModuleList
  txtModuleEdit_Change 0

End Sub

Private Sub DO_cmdEditProjClick(ByVal intIndex As Long)

  Dim I           As Long
  Dim LItem       As ListItem
  Dim Strnew      As String
  Dim Comp        As VBComponent
  Dim Proj        As VBProject
  Dim strNewPath  As String
  Dim MyhourGlass As cls_HourGlass
  Dim strOldPath  As String

  Set MyhourGlass = New cls_HourGlass
  For I = 0 To 3
    cmdEditProj(I).Enabled = False
  Next I
  Set LItem = lsvAllProjects.SelectedItem
  Set Proj = VBInstance.VBProjects.Item(LItem.Text)
  Select Case intIndex
   Case 0 'proj name
    Strnew = txtProjectEdit(0).Text
    If Strnew <> LItem.Text Then
      If LegalName(Strnew, strNewPath, 0, Comp, Proj) Then
        On Error Resume Next
        'v2.4.6 Thanks Mike Ulik Found te need for this error trap while working on the other bug
        Proj.Name = Strnew
        If Err.Number = 32813 Then
          mObjDoc.Safe_MsgBox "A Module or Project with the name" & strInSQuotes(Strnew, True) & "already exists.", vbCritical
        End If
        On Error GoTo 0
        ReplaceName LItem.Text, Strnew
        UpdateProjectList
      End If
    End If
   Case 1 'proj file
    Strnew = txtProjectEdit(1).Text
    If Strnew <> LItem.SubItems(1) Then
      strOldPath = Proj.FileName
      If LegalName(Strnew, strNewPath, 2, Comp, Proj) Then
        Proj.SaveAs strNewPath
        FileKill strOldPath
        UpdateProjectList
      End If
    End If
   Case 2 'sync file to project name
    If LegalName(txtProjectEdit(0).Text, strNewPath, 2, Comp, Proj) Then
      DO_cmdEditProjClick 0
      txtProjectEdit(1).Text = txtProjectEdit(0).Text
      DO_cmdEditProjClick 1
    End If
   Case 3 'sync  project to file name
    If LegalName(Left$(txtProjectEdit(1).Text, InStr(txtProjectEdit(1).Text, ".") - 1), strNewPath, 0, Comp, Proj) Then
      DO_cmdEditProjClick 1
      txtProjectEdit(0).Text = Left$(txtProjectEdit(1).Text, InStr(txtProjectEdit(1).Text, ".") - 1)
      DO_cmdEditProjClick 0
    End If
  End Select
  txtProjectEdit_Change 0

End Sub

Private Sub Form_Activate()

  If Not (bInitializing Or bAddinTerminate) Then
    cmdXPStyle.Caption = XPStyleCaption
    cmdXPStyle.ToolTipText = "Insert/Delete support code for XP Style"
    LoadFormPosition Me
    mnuPopUpDelete.Visible = False
    ActivatePage
    SetTopMost frm_CodeFixer, True
  End If

End Sub

Private Sub Form_Deactivate()

  SaveFormPosition Me
  ' Xcheck.SaveCheck

End Sub

Private Sub Form_Load()

  Dim I As Long

  If Not (bInitializing Or bAddinTerminate) Then
    Me.Caption = AppDetails
    With frm_FindSettings.lsvModNames
      For I = 1 To .ListItems.Count
        .ListItems(I).Checked = True
        'ModDesc(I).MDDontTouch = True
      Next I
    End With
    SetTopMost frm_CodeFixer, True
  End If

End Sub

Private Sub Form_Resize()

  KeepOnScreen
  'Keep tool on screen

End Sub

Private Sub Form_Unload(Cancel As Integer)

  SaveFormPosition Me
  If bFixing Then
    If Not bAborting Then
      If mObjDoc.GraphVisible Then
        D0_Abort
        Cancel = 1
      End If
     Else
      If bAbortComplete Then
        Cancel = 1
        Exit Sub
      End If
    End If
  End If
  If FrameActive = TPControls Then
    If Not cmdAutoLabel(2).Enabled Then
      mObjDoc.Safe_MsgBox "Please wait until the Control list is refreshed before closing this tool.", vbExclamation
      Cancel = 1
     Else
      mObjDoc.ForceFind RGSignature
      If lngFindCounter > 0 Then
        mObjDoc.Safe_MsgBox "Please run code (Use Ctrl+F5) to check that Controls Tool performed correctly." & vbNewLine & _
                    "Use Code Fixer comments to repair any failures.", vbExclamation
      End If
    End If
  End If

End Sub

Private Sub lstPrefixSuggest_Click()

  If LenB(lstPrefixSuggest.Text) Then
    SetNewName lstPrefixSuggest.Text
   Else
    WarningLabel 'clear message box
  End If

End Sub

Private Sub lstPrefixSuggest_DblClick()

  DO_cmdCtrlChangeClick

End Sub

Private Sub lstRestore_Click()

  restoreClick

End Sub

Private Sub lstSuggestModName_Click()

  txtModuleEdit(0).Text = lstSuggestModName.Text
  txtModuleEdit(0).SelStart = Len(txtModuleEdit(0).Text)

End Sub

Private Sub lstSuggestModName_DblClick()

  cmdEditMod(0).Value = True

End Sub

Private Sub lstSuggestProjName_Click()

  txtProjectEdit(0).Text = lstSuggestProjName.Text
  txtProjectEdit(0).SelStart = Len(txtProjectEdit(0).Text)

End Sub

Private Sub lsvAllControls_Click()

  DoAllSelect

End Sub

Private Sub lsvAllControls_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  PopupMenu mnuPopControls

End Sub

Private Sub lsvAllModules_Click()

  SetModuleListitem

End Sub

Private Sub lsvAllProjects_Click()

  SetProjectListitem

End Sub

Private Sub lsvUnDone_ItemCheck(ByVal Item As MSComctlLib.ListItem)

  With Item
    'ignore clean files
    .Checked = isFileDirty2(Item)
    ' ignore read-only files
    .Checked = IsFileReadOnlylsvItem(Item)
  End With

End Sub

Private Sub mnuCodeFixOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = RGSignature
   Case 1
    strSearch = mObjDoc.CurrentSelection
   Case 2
    mObjDoc.LikeSearch
    'Case 3 seperator
   Case 4
    strSearch = RGSignature & "Missing Dims Auto-inserted"
    'Case anything else they are sub-menus
  End Select
  If Index <> 2 Then ' likeSearch does its own thing
    mObjDoc.ForceFind strSearch
  End If

End Sub

Private Sub mnuCodeFixUsageOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = RGSignature & "|"
   Case 1
    strSearch = RGSignature & "|Usage "
   Case 2
    strSearch = RGSignature & "|Could be moved to following"
   Case 3
    strSearch = RGSignature & "|Code Entry point from "
   Case 4
    strSearch = RGSignature & "|NOT USED IN THIS MODULE"
   Case 5
    strSearch = RGSignature & "|Form/Module Interface"
   Case 6
    strSearch = RGSignature & "|Form Only Variable"
   Case 7
    strSearch = RGSignature & "|Could be Dim in following procedure"
   Case 8
    strSearch = RGSignature & "|Module/Module Interface"
   Case 9
    strSearch = RGSignature & "|RECOMMENDED:"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuColOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    bShowProject = Not bShowProject
   Case 1
    bShowComponent = Not bShowComponent
   Case 2
    bShowCompLineNo = Not bShowCompLineNo
   Case 3
    bShowRoutine = Not bShowRoutine
   Case 4
    bShowProcLineNo = Not bShowProcLineNo
  End Select
  mObjDoc.GridReSize

End Sub

Private Sub mnuControlOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = " renamed to "
   Case 1
    strSearch = WARNING_MSG & "Unused Control Code"
   Case 2
    strSearch = RGSignature & "POTENTIAL XP FRAME BUG DETECTED"
   Case 3
    strSearch = WARNING_MSG & "Default Property of Control"
   Case 4
    strSearch = WARNING_MSG & "It is clearer to use"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuDeleteOpt_Click(Index As Integer)

  mObjDoc.DoDelete Index

End Sub

Private Sub mnuDimusageOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = RGSignature & "|Dim Usage:("
   Case 1
    strSearch = RGSignature & "|Dim Usage:(1)"
   Case 2
    strSearch = RGSignature & "|Dim Usage:(2)"
   Case 3
    strSearch = RGSignature & "|Dim Usage:(3)"
   Case 4
    strSearch = RGSignature & "|Dim Usage:(3) on (2)"
   Case 5
    DimUsageDeleteNonProblem
  End Select
  If Index <> 5 Then
    mObjDoc.InitiateSearch strSearch
  End If

End Sub

Private Sub mnuElseOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = SUGGESTION_MSG & "'Else"
   Case 1
    strSearch = SUGGESTION_MSG & "'ElseIf'"
   Case 2
    strSearch = SUGGESTION_MSG & "'Else'"
   Case 3
    strSearch = "Short Curcuit"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuFindShow_Click()

  mnuColOpt(0).Checked = bShowProject
  mnuColOpt(1).Checked = bShowComponent
  mnuColOpt(2).Checked = bShowCompLineNo
  mnuColOpt(3).Checked = bShowRoutine
  mnuColOpt(4).Checked = bShowProcLineNo

End Sub

Private Sub mnuGeneralOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = UPDATED_MSG
   Case 1
    strSearch = SUGGESTION_MSG
   Case 2
    strSearch = WARNING_MSG
   Case 3
    strSearch = RISK_MSG
   Case 4
    strSearch = PREVIOUSCODE_MSG
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuGotoOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = WARNING_MSG & "This GoTo label is orphaned"
   Case 1
    strSearch = WARNING_MSG & "This GoTo command has no target"
   Case 2
    strSearch = WARNING_MSG & "This 'GoTo 0'"
   Case 3
    strSearch = WARNING_MSG & "This GoTo might be replaced"
   Case 4
    strSearch = WARNING_MSG & "This GoTo label is the target of a GoTo which"
   Case 5
    strSearch = WARNING_MSG & "This GoTo jumps your code into"
   Case 6
    strSearch = WARNING_MSG & "This GoTo jumps your code out"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuParam_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = WARNING_MSG & "Unused Parameter"
   Case 1
    strSearch = WARNING_MSG & "Untyped Parameters"
   Case 2
    strSearch = SUGGESTION_MSG & "Insert 'ByVal '"
   Case 3
    strSearch = WARNING_MSG & "'ByVal ' inserted"
   Case 4
    strSearch = WARNING_MSG & "Function changed to Sub"
   Case 5
    strSearch = WARNING_MSG & "Function Typed automatically"
   Case 6
    strSearch = WARNING_MSG & "Large Code procedure ("
   Case 7
    strSearch = WARNING_MSG & "Large Control procedure ("
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnupopControlLev2_Click(Index As Integer)

  Select Case Index
   Case 0
    bShowctrlPRoject = Not bShowctrlPRoject
    mnupopControlLev2(Index).Checked = bShowctrlPRoject
   Case 1
    bShowctrlComponent = Not bShowctrlComponent
    mnupopControlLev2(Index).Checked = bShowctrlComponent
   Case 2
    bShowctrlCaption = Not bShowctrlCaption
    mnupopControlLev2(Index).Checked = bShowctrlCaption
  End Select
  AllControlsColumnWidths
  'CtrlGridReSize

End Sub

Private Sub mnupopControlLev3_Click(Index As Integer)

  Dim I As Long

  lsvAllControls.SortKey = Index '+ 1
  For I = mnupopControlLev3.LBound To mnupopControlLev3.UBound
    mnupopControlLev3(I).Checked = Index = I
  Next I

End Sub

Private Sub mnuPopControls_Click()

  mnupopControlLev2(0).Checked = bShowctrlPRoject
  mnupopControlLev2(1).Checked = bShowctrlComponent
  mnupopControlLev2(2).Checked = bShowctrlName

End Sub

Private Sub mnuPopUpDeleteOpt_Click(Index As Integer)

  mObjDoc.DoDelete Index

End Sub

Private Sub mnuStructural_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = "Structure Error"
   Case 1
    strSearch = WARNING_MSG & "Scope Changed to"
   Case 2
    strSearch = WARNING_MSG & "Scope Too Large"
   Case 3
    If FixData(UpdateInteger2Long).FixLevel > CommentOnly Then
      strSearch = WARNING_MSG & "Integer "
     Else
      strSearch = SUGGESTION_MSG & "Integer "
    End If
   Case 4
    strSearch = WARNING_MSG & "Duplicated Name"
   Case 5
    strSearch = RISK_MSG & " Turns off 'On Error Resume Next' in routine (Good coding but may not be what you want) "
   Case 6
    strSearch = SUGGESTION_MSG & "Code used to"
   Case 7
    strSearch = RGSignature & "Structure Expanded."
   Case 8
    strSearch = RGSignature & "Pleonasm Removed"
   Case 9
    strSearch = RGSignature & "For-Variable Inserted"
   Case 10
    strSearch = WARNING_MSG & "Empty String "
   Case 11
    strSearch = RISK_MSG & " Load/UnLoad"
   Case 12
    strSearch = SUGGESTION_MSG & "Hard-Coded Paths"
   Case 13
    strSearch = RGSignature & "DefType no longer needed."
   Case 14
    strSearch = WARNING_MSG & "Line number"
   Case 15
    strSearch = "Obsolete Code"
   Case 16
    strSearch = WARNING_MSG & "Unnecessary 'Exit"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuUntyped_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = ": UnTyped"
   Case 1
    strSearch = ": UnTyped Dim"
   Case 2
    strSearch = RGSignature & "Untyped Variable. Will behave"
   Case 3
    strSearch = UPDATED_MSG & "UnTyped Const with"
   Case 4
    strSearch = WARNING_MSG & "Untyped Parameters"
   Case 5
    strSearch = WARNING_MSG & "UnTyped Function"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuUnused_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = WARNING_MSG & "Unused "
   Case 1
    strSearch = WARNING_MSG & "Unused Variable"
   Case 2
    strSearch = WARNING_MSG & "Unused Dim"
   Case 3
    strSearch = WARNING_MSG & "Unused Declare"
   Case 4
    strSearch = WARNING_MSG & "Unused Type"
   Case 5
    strSearch = WARNING_MSG & "Unused Enum"
   Case 6
    strSearch = WARNING_MSG & "Unused Sub"
   Case 7
    strSearch = WARNING_MSG & "Unused Function"
   Case 8
    strSearch = WARNING_MSG & "Unused Property"
   Case 9
    strSearch = WARNING_MSG & "Unused Parameter"
   Case 10
    strSearch = WARNING_MSG & "Unused Control Code"
   Case 11
    strSearch = WARNING_MSG & "Empty '"
   Case 12
    strSearch = WARNING_MSG & "Unneeded "
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub mnuViewSortOpt_Click(Index As Integer)

  mObjDoc.DoSort Index

End Sub

Private Sub mnuWithOpt_Click(Index As Integer)

  Dim strSearch As String

  Select Case Index
   Case 0
    strSearch = RGSignature & "Auto-inserted With End...With Structure"
   Case 1
    strSearch = SUGGESTION_MSG & "Possible Start:"
   Case 2
    strSearch = SUGGESTION_MSG & "Possible End:"
   Case 3
    strSearch = "'APPROVED(Y )"
   Case 4
    strSearch = RGSignature & "APPROVED With End...With Structure"
   Case 5
    strSearch = SUGGESTION_MSG & "'With' structure contains"
   Case 6
    strSearch = RGSignature & "Purified With"
  End Select
  mObjDoc.InitiateSearch strSearch

End Sub

Private Sub tbsFileTool_Click()

  Do_tbsFileToolClick

End Sub

Private Sub tbsModule_Click()

  If PrevTab <> 0 Then
    fraModule(PrevTab).Visible = False
  End If
  With tbsModule
    fraModule(.SelectedItem.Index).Top = .ClientTop
    fraModule(.SelectedItem.Index).Left = .ClientLeft
    fraModule(.SelectedItem.Index).Caption = vbNullString
    fraModule(.SelectedItem.Index).Visible = True
    PrevTab = .SelectedItem.Index
    If .SelectedItem.Index = 2 Then
      Generate_ProjectArray True
     Else
      Generate_ModuleArray True
    End If
  End With

End Sub

Private Sub txtCtrlNewName_Change()

  If Not AutoReName Then
    cmdCtrlChange.Enabled = (LCase$(txtCtrlNewName.Text) <> LCase$(CntrlDesc(LngCurrentControl).CDName))
    cmdCtrlChange.Enabled = txtCtrlNewName.Text <> CntrlDesc(LngCurrentControl).CDName
  End If
  If txtCtrlNewName.Text <> lstPrefixSuggest.Text Then
    lstPrefixSuggest.ListIndex = -1
  End If

End Sub

Private Sub txtCtrlNewName_GotFocus()

  SetNewName txtCtrlNewName.Text

End Sub

Private Sub txtCtrlNewName_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    If cmdCtrlChange.Enabled Then
      KeyAscii = 0
      DO_cmdCtrlChangeClick
    End If
  End If

End Sub

Private Sub txtModuleEdit_Change(Index As Integer)

  Dim Selected As Long
  Dim ROnly    As Boolean

  Selected = ModDescMember(lsvAllModules.SelectedItem.SubItems(1))
  If Selected > -1 Then
    ROnly = ModDesc(Selected).MDReadOnly
  End If
  cmdEditMod(0).Enabled = txtModuleEdit(0).Text <> lsvAllModules.SelectedItem.SubItems(2) And Selected > -1 And Not ROnly
  cmdEditMod(1).Enabled = txtModuleEdit(1).Text <> lsvAllModules.SelectedItem.SubItems(1) And Selected > -1 And Not ROnly
  cmdEditMod(2).Enabled = Not SmartLeft(txtModuleEdit(1).Text, txtModuleEdit(0).Text & ".") And Selected > -1 And Not ROnly
  cmdEditMod(3).Enabled = cmdEditMod(2).Enabled
  lstSuggestModName.Enabled = Selected > -1 And Not ROnly
  txtModuleEdit(0).Enabled = Not ROnly
  txtModuleEdit(1).Enabled = Not ROnly
  lblReadOnlyMod.Visible = ROnly

End Sub

Private Sub txtModuleEdit_KeyPress(Index As Integer, _
                                   KeyAscii As Integer)

  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Select Case Index
     Case 0
      txtModuleEdit(0).Text = lsvAllModules.SelectedItem.SubItems(1)
     Case 1
      txtModuleEdit(1).Text = lsvAllModules.SelectedItem.SubItems(2)
     Case 2
      txtModuleEdit(2).Text = lsvAllModules.SelectedItem.Text
    End Select
   ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If Index < 3 Then
      cmdEditMod(Index).Value = True
    End If
  End If

End Sub

Private Sub txtProjectEdit_Change(Index As Integer)

  Dim Proj  As VBProject
  Dim LItem As ListItem
  Dim ROnly As Boolean

  Set LItem = lsvAllProjects.SelectedItem
  If Not LItem Is Nothing Then
    Set Proj = VBInstance.VBProjects.Item(LItem.Text)
    ROnly = IsFileReadOnly(Proj.FileName)
  End If
  cmdEditProj(0).Enabled = txtProjectEdit(0).Text <> lsvAllProjects.SelectedItem.Text And Not ROnly
  cmdEditProj(1).Enabled = txtProjectEdit(1).Text <> lsvAllProjects.SelectedItem.SubItems(1) And Not ROnly
  cmdEditProj(2).Enabled = Not SmartLeft(txtProjectEdit(1).Text, txtProjectEdit(0).Text & ".")
  cmdEditProj(3).Enabled = cmdEditProj(2).Enabled
  lstSuggestProjName.Enabled = Not ROnly
  txtProjectEdit(0).Enabled = Not ROnly
  txtProjectEdit(1).Enabled = Not ROnly
  lblReadOnlyPrj.Visible = ROnly

End Sub

Private Sub txtProjectEdit_KeyPress(Index As Integer, _
                                    KeyAscii As Integer)

  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Select Case Index
     Case 0
      txtProjectEdit(0).Text = lsvAllModules.SelectedItem.Text
     Case 1
      txtProjectEdit(1).Text = lsvAllModules.SelectedItem.SubItems(1)
     Case 2
      txtProjectEdit(2).Text = FileNameOnly(VBInstance.VBProjects.Item(lsvAllModules.SelectedItem.Text).FileName)
    End Select
   ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If Index < 3 Then
      cmdEditProj(Index).Value = True
    End If
  End If

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:11:41 AM) 10 + 1042 = 1052 Lines Thanks Ulli for inspiration and lots of code.

