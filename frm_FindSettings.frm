VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_FindSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7335
   Icon            =   "frm_FindSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleMode       =   0  'User
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSetting 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "Apply"
      Height          =   315
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame fraprop 
      Caption         =   "Settings"
      Height          =   7095
      Index           =   1
      Left            =   1560
      TabIndex        =   37
      Top             =   240
      Width           =   10575
      Begin VB.PictureBox picCFXPBugFixSettings 
         BorderStyle     =   0  'None
         Height          =   6330
         Index           =   13
         Left            =   240
         ScaleHeight     =   6330
         ScaleWidth      =   9600
         TabIndex        =   38
         Top             =   360
         Width           =   9600
         Begin VB.Frame fraSubSetting 
            Caption         =   "General"
            Height          =   6255
            Index           =   0
            Left            =   600
            TabIndex        =   46
            Top             =   0
            Width           =   6375
            Begin VB.PictureBox picCFXPBugFixSettings 
               BorderStyle     =   0  'None
               Height          =   5895
               Index           =   6
               Left            =   120
               ScaleHeight     =   5895
               ScaleWidth      =   6135
               TabIndex        =   47
               Top             =   240
               Width           =   6135
               Begin VB.Frame fraSetting 
                  Caption         =   "Tool Behaviour"
                  Height          =   1335
                  Left            =   0
                  TabIndex        =   74
                  Top             =   0
                  Width           =   1935
                  Begin VB.PictureBox picCFXPBugFixSettings 
                     BorderStyle     =   0  'None
                     Height          =   975
                     Index           =   4
                     Left            =   120
                     ScaleHeight     =   975
                     ScaleWidth      =   1695
                     TabIndex        =   75
                     Top             =   240
                     Width           =   1695
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Create a Back Up"
                        Height          =   195
                        Index           =   3
                        Left            =   0
                        TabIndex        =   98
                        Top             =   780
                        Value           =   1  'Checked
                        Width           =   1695
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Low CPU Usage"
                        Height          =   195
                        Index           =   6
                        Left            =   0
                        TabIndex        =   79
                        ToolTipText     =   "Slower (Some portables have heat problems while processing large files)"
                        Top             =   390
                        Width           =   1575
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Stay On Top"
                        Height          =   195
                        Index           =   4
                        Left            =   0
                        TabIndex        =   78
                        ToolTipText     =   "Code Fixer stays on top while processing"
                        Top             =   0
                        Width           =   1575
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Visible Scan"
                        Height          =   195
                        Index           =   5
                        Left            =   0
                        TabIndex        =   77
                        ToolTipText     =   "Code Panes show when being fixed"
                        Top             =   195
                        Width           =   1575
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Auto Read-Write"
                        Height          =   195
                        Index           =   7
                        Left            =   0
                        TabIndex        =   76
                        ToolTipText     =   "Set all files to Read-Write (allows editing)"
                        Top             =   585
                        Width           =   1575
                     End
                  End
               End
               Begin VB.Frame fraButtons 
                  Caption         =   "Quick Sets"
                  Height          =   1335
                  Left            =   4440
                  TabIndex        =   66
                  Top             =   0
                  Width           =   1575
                  Begin VB.PictureBox picCFXPBugFixSettings 
                     BorderStyle     =   0  'None
                     Height          =   1080
                     Index           =   11
                     Left            =   100
                     ScaleHeight     =   1080
                     ScaleWidth      =   1380
                     TabIndex        =   67
                     Top             =   175
                     Width           =   1375
                     Begin VB.PictureBox picCFXPBugFixSettings 
                        BorderStyle     =   0  'None
                        Height          =   975
                        Index           =   3
                        Left            =   20
                        ScaleHeight     =   975
                        ScaleWidth      =   1335
                        TabIndex        =   68
                        Top             =   40
                        Width           =   1335
                        Begin VB.OptionButton optLevelSetting 
                           Caption         =   "User"
                           Enabled         =   0   'False
                           Height          =   195
                           Index           =   4
                           Left            =   0
                           TabIndex        =   73
                           Top             =   780
                           Width           =   1095
                        End
                        Begin VB.OptionButton optLevelSetting 
                           Caption         =   "Maximum"
                           Height          =   195
                           Index           =   3
                           Left            =   0
                           TabIndex        =   72
                           Top             =   585
                           Width           =   1095
                        End
                        Begin VB.OptionButton optLevelSetting 
                           Caption         =   "Average"
                           Height          =   195
                           Index           =   2
                           Left            =   0
                           TabIndex        =   71
                           Top             =   390
                           Value           =   -1  'True
                           Width           =   1095
                        End
                        Begin VB.OptionButton optLevelSetting 
                           Caption         =   "Minimum"
                           Height          =   195
                           Index           =   1
                           Left            =   0
                           TabIndex        =   70
                           Top             =   195
                           Width           =   1095
                        End
                        Begin VB.OptionButton optLevelSetting 
                           Caption         =   "All Off"
                           Height          =   195
                           Index           =   0
                           Left            =   0
                           TabIndex        =   69
                           Top             =   0
                           Width           =   1095
                        End
                     End
                  End
               End
               Begin VB.CommandButton cmdStartFromSettings 
                  Caption         =   "Start"
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   65
                  Top             =   5400
                  Width           =   1335
               End
               Begin VB.Frame fraOldReg 
                  Caption         =   "Old Settings"
                  Height          =   615
                  Left            =   4440
                  TabIndex        =   61
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   1575
                  Begin VB.PictureBox picCFXPBugFixSettings 
                     BorderStyle     =   0  'None
                     Height          =   365
                     Index           =   5
                     Left            =   100
                     ScaleHeight     =   360
                     ScaleWidth      =   1380
                     TabIndex        =   62
                     Top             =   175
                     Width           =   1375
                     Begin VB.CommandButton cmdOldReg 
                        Caption         =   "Delete"
                        Height          =   255
                        Index           =   0
                        Left            =   755
                        TabIndex        =   64
                        ToolTipText     =   "Remove all old settings"
                        Top             =   40
                        Width           =   615
                     End
                     Begin VB.CommandButton cmdOldReg 
                        Caption         =   "Apply"
                        Height          =   255
                        Index           =   1
                        Left            =   20
                        TabIndex        =   63
                        ToolTipText     =   "Copy Previous Version's settings to current version."
                        Top             =   40
                        Width           =   615
                     End
                  End
               End
               Begin VB.Frame fraWarnings 
                  Caption         =   "Warning MsgBoxes"
                  Height          =   975
                  Left            =   0
                  TabIndex        =   57
                  ToolTipText     =   "These are MegBoxes you might like to stop appearing"
                  Top             =   1560
                  Width           =   1935
                  Begin VB.CheckBox chkUser 
                     Caption         =   "Not a project"
                     Height          =   195
                     Index           =   13
                     Left            =   120
                     TabIndex        =   60
                     Top             =   435
                     Width           =   1455
                  End
                  Begin VB.CheckBox chkUser 
                     Caption         =   "Control Tab"
                     Height          =   195
                     Index           =   14
                     Left            =   120
                     TabIndex        =   59
                     Top             =   630
                     Width           =   1455
                  End
                  Begin VB.CheckBox chkUser 
                     Caption         =   "Large code"
                     Height          =   195
                     Index           =   12
                     Left            =   120
                     TabIndex        =   58
                     Top             =   240
                     Width           =   1455
                  End
               End
               Begin VB.Frame fraModules 
                  Caption         =   "Modules (Uncheck = Read Only)"
                  Height          =   2175
                  Left            =   0
                  TabIndex        =   55
                  Top             =   3600
                  Width           =   3735
                  Begin MSComctlLib.ListView lsvModNames 
                     Height          =   1815
                     Left            =   120
                     TabIndex        =   56
                     Top             =   240
                     Width           =   3495
                     _ExtentX        =   6165
                     _ExtentY        =   3201
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     Checkboxes      =   -1  'True
                     FullRowSelect   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483624
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Key             =   "module"
                        Text            =   "Module"
                        Object.Width           =   2540
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Key             =   "project"
                        Text            =   "Project"
                        Object.Width           =   2540
                     EndProperty
                  End
               End
               Begin VB.Frame fraComments 
                  Caption         =   "Comments"
                  Height          =   1335
                  Left            =   4080
                  TabIndex        =   49
                  Top             =   3720
                  Width           =   2055
                  Begin VB.CheckBox chkUser 
                     Alignment       =   1  'Right Justify
                     Caption         =   " Usage"
                     Height          =   195
                     Index           =   15
                     Left            =   120
                     TabIndex        =   54
                     ToolTipText     =   "Include large number of usage statistics"
                     Top             =   1020
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkUser 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Ignore Comment Only"
                     Height          =   195
                     Index           =   10
                     Left            =   120
                     TabIndex        =   53
                     ToolTipText     =   "Don't include comments unless they involve fixes"
                     Top             =   825
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkUser 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Structural"
                     Height          =   195
                     Index           =   2
                     Left            =   120
                     TabIndex        =   52
                     ToolTipText     =   "Adds extra comments"
                     Top             =   630
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkUser 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Previous Code"
                     Height          =   195
                     Index           =   1
                     Left            =   120
                     TabIndex        =   51
                     ToolTipText     =   "Include original code in comments"
                     Top             =   435
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkUser 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Verbose"
                     Height          =   195
                     Index           =   0
                     Left            =   120
                     TabIndex        =   50
                     ToolTipText     =   "Extended comments which explain the main comment in detail"
                     Top             =   240
                     Width           =   1815
                  End
               End
               Begin VB.Frame fraLayout 
                  Caption         =   "Layout"
                  Height          =   3615
                  Left            =   2040
                  TabIndex        =   48
                  Top             =   0
                  Width           =   2055
                  Begin VB.PictureBox picCFXPBugFixSettings 
                     BorderStyle     =   0  'None
                     Height          =   3360
                     Index           =   9
                     Left            =   100
                     ScaleHeight     =   3360
                     ScaleWidth      =   1860
                     TabIndex        =   81
                     Top             =   180
                     Width           =   1855
                     Begin VB.Frame fraCaseElseIndent 
                        Caption         =   "Case indent"
                        Height          =   855
                        Index           =   1
                        Left            =   0
                        TabIndex        =   121
                        Top             =   1440
                        Width           =   1815
                        Begin VB.PictureBox picCFXPBugFixSettings 
                           BorderStyle     =   0  'None
                           Height          =   650
                           Index           =   12
                           Left            =   100
                           ScaleHeight     =   645
                           ScaleWidth      =   1620
                           TabIndex        =   122
                           Top             =   175
                           Width           =   1615
                           Begin VB.OptionButton OptCaseIndent 
                              Caption         =   "Full Indent"
                              Height          =   195
                              Index           =   2
                              Left            =   20
                              TabIndex        =   125
                              Top             =   430
                              Width           =   1095
                           End
                           Begin VB.OptionButton OptCaseIndent 
                              Caption         =   "1/2 indent"
                              Height          =   195
                              Index           =   1
                              Left            =   20
                              TabIndex        =   124
                              Top             =   235
                              Value           =   -1  'True
                              Width           =   1575
                           End
                           Begin VB.OptionButton OptCaseIndent 
                              Caption         =   "None"
                              Height          =   195
                              Index           =   0
                              Left            =   20
                              TabIndex        =   123
                              Top             =   40
                              Width           =   1575
                           End
                        End
                     End
                     Begin MSComCtl2.UpDown UpDLongProc 
                        Height          =   285
                        Left            =   1485
                        TabIndex        =   95
                        ToolTipText     =   "Number of code lines before the Large Procedure comment is inserted"
                        Top             =   2880
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   503
                        _Version        =   393216
                        Value           =   40
                        AutoBuddy       =   -1  'True
                        BuddyControl    =   "txtLongProc"
                        BuddyDispid     =   196626
                        OrigLeft        =   1440
                        OrigTop         =   2040
                        OrigRight       =   1695
                        OrigBottom      =   2295
                        Max             =   150
                        Min             =   40
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   0
                        Enabled         =   -1  'True
                     End
                     Begin VB.TextBox txtLongProc 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   1080
                        TabIndex        =   94
                        Text            =   "50"
                        Top             =   2880
                        Width           =   390
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Indent Comments"
                        Height          =   195
                        Index           =   11
                        Left            =   20
                        TabIndex        =   87
                        ToolTipText     =   "Indent with surrounding code/No indent for comments"
                        Top             =   445
                        Width           =   1815
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Preserve Blanks"
                        Height          =   195
                        Index           =   9
                        Left            =   20
                        TabIndex        =   86
                        ToolTipText     =   "If source contained blank lines keep them"
                        Top             =   235
                        Width           =   1815
                     End
                     Begin VB.CheckBox chkUser 
                        Caption         =   "Space Separators"
                        Height          =   195
                        Index           =   8
                        Left            =   20
                        TabIndex        =   85
                        ToolTipText     =   "Use blank lines in formatting"
                        Top             =   40
                        Width           =   1815
                     End
                     Begin VB.Frame fraCaseElseIndent 
                        Caption         =   "Else indent"
                        Height          =   855
                        Index           =   0
                        Left            =   20
                        TabIndex        =   84
                        Top             =   645
                        Width           =   1815
                        Begin VB.PictureBox picCFXPBugFixSettings 
                           BorderStyle     =   0  'None
                           Height          =   650
                           Index           =   10
                           Left            =   100
                           ScaleHeight     =   645
                           ScaleWidth      =   1620
                           TabIndex        =   89
                           Top             =   175
                           Width           =   1615
                           Begin VB.OptionButton OptElseIndent 
                              Caption         =   "None"
                              Height          =   195
                              Index           =   0
                              Left            =   20
                              TabIndex        =   92
                              Top             =   40
                              Width           =   1575
                           End
                           Begin VB.OptionButton OptElseIndent 
                              Caption         =   "1/2 indent"
                              Height          =   195
                              Index           =   1
                              Left            =   20
                              TabIndex        =   91
                              Top             =   235
                              Value           =   -1  'True
                              Width           =   1575
                           End
                           Begin VB.OptionButton OptElseIndent 
                              Caption         =   "Full Indent"
                              Height          =   195
                              Index           =   2
                              Left            =   20
                              TabIndex        =   90
                              Top             =   430
                              Width           =   1095
                           End
                        End
                     End
                     Begin VB.TextBox txtLongLine 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   1095
                        TabIndex        =   83
                        Text            =   "100"
                        Top             =   2520
                        Width           =   390
                     End
                     Begin MSComCtl2.UpDown UpdLongLine 
                        Height          =   285
                        Left            =   1485
                        TabIndex        =   82
                        ToolTipText     =   "Set length of lines for line continuation formatting"
                        Top             =   2520
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   503
                        _Version        =   393216
                        Value           =   100
                        AutoBuddy       =   -1  'True
                        BuddyControl    =   "txtLongLine"
                        BuddyDispid     =   196628
                        OrigLeft        =   1455
                        OrigTop         =   1680
                        OrigRight       =   1710
                        OrigBottom      =   1965
                        Max             =   150
                        Min             =   60
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   0
                        Enabled         =   -1  'True
                     End
                     Begin VB.Label lblFindSettings 
                        Caption         =   "Long Routine (40-150 Lines)"
                        Height          =   375
                        Index           =   1
                        Left            =   0
                        TabIndex        =   93
                        Top             =   2880
                        Width           =   1095
                     End
                     Begin VB.Label lblFindSettings 
                        Caption         =   "Long Line (60-150 Char)"
                        Height          =   375
                        Index           =   0
                        Left            =   15
                        TabIndex        =   88
                        Top             =   2400
                        Width           =   1095
                     End
                  End
               End
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Declaration Tools"
            Height          =   735
            Index           =   1
            Left            =   6720
            TabIndex        =   45
            Top             =   0
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   0
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "ReSutructure"
            Height          =   615
            Index           =   2
            Left            =   6840
            TabIndex        =   44
            Top             =   720
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   1
               Left            =   120
               TabIndex        =   115
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Locals"
            Height          =   615
            Index           =   4
            Left            =   6840
            TabIndex        =   43
            Top             =   2160
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   3
               Left            =   240
               TabIndex        =   117
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Suggestions"
            Height          =   615
            Index           =   5
            Left            =   6840
            TabIndex        =   42
            Top             =   2760
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   4
               Left            =   240
               TabIndex        =   118
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Unused Detection"
            Height          =   615
            Index           =   6
            Left            =   6600
            TabIndex        =   41
            Top             =   3600
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   5
               Left            =   240
               TabIndex        =   119
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Format"
            Height          =   615
            Index           =   7
            Left            =   6600
            TabIndex        =   40
            Top             =   4320
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   6
               Left            =   240
               TabIndex        =   120
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin VB.Frame fraSubSetting 
            Caption         =   "Parameters"
            Height          =   615
            Index           =   3
            Left            =   6840
            TabIndex        =   39
            Top             =   1440
            Width           =   6375
            Begin prj_CodeFix3.PropertyList ctlSettingsList 
               Height          =   2160
               Index           =   2
               Left            =   120
               TabIndex        =   116
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   4260
               ListColor       =   -2147483624
               Cols            =   9
               Rows            =   6
               FixedRows       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FormatString    =   "Property|Value"
            End
         End
         Begin MSComctlLib.TabStrip tbsSettings 
            Height          =   5775
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   10186
            MultiRow        =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   8
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "General"
                  Key             =   "general"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Declaration"
                  Key             =   "declaration"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Restructure"
                  Key             =   "restructure"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Parameter"
                  Key             =   "parameter"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Locals"
                  Key             =   "locals"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Suggest"
                  Key             =   "Suggestion"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Unused"
                  Key             =   "Unused"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Format"
                  Key             =   "format"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame fraprop 
      Caption         =   "Find"
      Height          =   3615
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   8775
      Begin VB.PictureBox picCFXPBugFixSettings 
         BorderStyle     =   0  'None
         Height          =   3480
         Index           =   8
         Left            =   120
         ScaleHeight     =   3480
         ScaleWidth      =   8580
         TabIndex        =   4
         Top             =   120
         Width           =   8580
         Begin VB.CheckBox ChkTBarButton 
            Caption         =   "Show Toolbar Button"
            Height          =   195
            Left            =   3120
            TabIndex        =   99
            ToolTipText     =   "Change will not occur until VB is restarted"
            Top             =   3085
            Width           =   1935
         End
         Begin VB.Frame fraToolBar_SearchInput 
            Caption         =   "ToolBar/Search Input"
            Height          =   495
            Left            =   20
            TabIndex        =   32
            ToolTipText     =   "Set position of Toolbar on Find Tool"
            Top             =   2925
            Width           =   2655
            Begin VB.PictureBox picCFXPBugFixSettings 
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   7
               Left            =   100
               ScaleHeight     =   240
               ScaleWidth      =   2460
               TabIndex        =   33
               Top             =   175
               Width           =   2455
               Begin VB.OptionButton optInputPos 
                  Caption         =   "Bottom"
                  Height          =   195
                  Index           =   1
                  Left            =   1340
                  TabIndex        =   35
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton optInputPos 
                  Caption         =   "Top"
                  Height          =   195
                  Index           =   0
                  Left            =   20
                  TabIndex        =   34
                  Top             =   45
                  Width           =   1095
               End
            End
         End
         Begin VB.Frame fraSettings 
            Caption         =   "Colours"
            Height          =   3015
            Index           =   3
            Left            =   5550
            TabIndex        =   24
            Top             =   -30
            Width           =   3015
            Begin VB.PictureBox picCFXPBugFixSettings 
               BorderStyle     =   0  'None
               Height          =   2760
               Index           =   0
               Left            =   100
               ScaleHeight     =   2760
               ScaleWidth      =   2820
               TabIndex        =   25
               Top             =   175
               Width           =   2815
               Begin VB.Frame fraSettings 
                  Caption         =   "Selection"
                  Height          =   855
                  Index           =   5
                  Left            =   20
                  TabIndex        =   31
                  Top             =   885
                  Width           =   1335
                  Begin VB.PictureBox picCFXPBugFixfrm_FindSettings 
                     BorderStyle     =   0  'None
                     Height          =   605
                     Index           =   1
                     Left            =   100
                     ScaleHeight     =   600
                     ScaleWidth      =   1140
                     TabIndex        =   108
                     Top             =   175
                     Width           =   1135
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Text"
                        Height          =   255
                        Index           =   2
                        Left            =   20
                        TabIndex        =   110
                        Top             =   40
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Back"
                        Height          =   255
                        Index           =   3
                        Left            =   20
                        TabIndex        =   109
                        Top             =   280
                        Width           =   1095
                     End
                  End
               End
               Begin VB.PictureBox picCFXPBugFixSettings 
                  BorderStyle     =   0  'None
                  Height          =   2760
                  Index           =   1
                  Left            =   0
                  ScaleHeight     =   2760
                  ScaleWidth      =   2820
                  TabIndex        =   26
                  Top             =   -25
                  Width           =   2820
                  Begin VB.Frame fraSettings 
                     Caption         =   "General"
                     Height          =   855
                     Index           =   6
                     Left            =   0
                     TabIndex        =   28
                     Top             =   0
                     Width           =   1335
                     Begin VB.PictureBox picCFXPBugFixfrm_FindSettings 
                        BorderStyle     =   0  'None
                        Height          =   605
                        Index           =   2
                        Left            =   100
                        ScaleHeight     =   600
                        ScaleWidth      =   1140
                        TabIndex        =   111
                        Top             =   175
                        Width           =   1135
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Text"
                           Height          =   255
                           Index           =   0
                           Left            =   20
                           TabIndex        =   113
                           Top             =   40
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Back"
                           Height          =   255
                           Index           =   1
                           Left            =   20
                           TabIndex        =   112
                           Top             =   280
                           Width           =   1095
                        End
                     End
                  End
                  Begin VB.Frame fraSettings 
                     Caption         =   "Default"
                     Height          =   2175
                     Index           =   4
                     Left            =   1440
                     TabIndex        =   27
                     Top             =   0
                     Width           =   1320
                     Begin VB.PictureBox picCFXPBugFixfrm_FindSettings 
                        BorderStyle     =   0  'None
                        Height          =   1925
                        Index           =   0
                        Left            =   100
                        ScaleHeight     =   1920
                        ScaleWidth      =   1125
                        TabIndex        =   100
                        Top             =   175
                        Width           =   1120
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Replacing"
                           Height          =   255
                           Index           =   9
                           Left            =   20
                           TabIndex        =   107
                           Top             =   1600
                           Width           =   1095
                        End
                        Begin VB.Label lblSettings 
                           Alignment       =   2  'Center
                           Caption         =   "Back Colours"
                           Height          =   255
                           Left            =   20
                           TabIndex        =   106
                           Top             =   280
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Text"
                           Height          =   255
                           Index           =   4
                           Left            =   20
                           TabIndex        =   105
                           Top             =   40
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "No Find"
                           Height          =   255
                           Index           =   8
                           Left            =   20
                           TabIndex        =   104
                           Top             =   1285
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Pattern Find"
                           Height          =   255
                           Index           =   7
                           Left            =   20
                           TabIndex        =   103
                           Top             =   1030
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Searching"
                           Height          =   255
                           Index           =   6
                           Left            =   20
                           TabIndex        =   102
                           Top             =   760
                           Width           =   1095
                        End
                        Begin VB.Label LblColour 
                           Alignment       =   2  'Center
                           BorderStyle     =   1  'Fixed Single
                           Caption         =   "Standard"
                           Height          =   255
                           Index           =   5
                           Left            =   20
                           TabIndex        =   101
                           Top             =   520
                           Width           =   1095
                        End
                     End
                  End
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Restore User"
                     Height          =   255
                     Index           =   10
                     Left            =   135
                     TabIndex        =   30
                     Top             =   2400
                     Width           =   1095
                  End
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Default"
                     Height          =   255
                     Index           =   11
                     Left            =   1560
                     TabIndex        =   29
                     Top             =   2400
                     Width           =   1095
                  End
               End
            End
         End
         Begin VB.Frame fraSettings 
            Caption         =   "Replace"
            Height          =   1095
            Index           =   2
            Left            =   35
            TabIndex        =   20
            Top             =   1770
            Width           =   2655
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Replace To Search"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Blank Warning"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "  Show Filter Warning"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame fraSettings 
            Caption         =   " Found Grid Appearance"
            Height          =   1695
            Index           =   0
            Left            =   15
            TabIndex        =   10
            Top             =   -30
            Width           =   5415
            Begin VB.CheckBox ChkRemFilters 
               Caption         =   "Remember Filters"
               Height          =   195
               Left            =   2640
               TabIndex        =   19
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CheckBox ChkSelectWhole 
               Caption         =   "Find select whole line"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox ChkAutoSelectedText 
               Caption         =   "Auto Selected Text"
               Height          =   195
               Left            =   2640
               TabIndex        =   17
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Component Lines"
               Height          =   195
               Index           =   2
               Left            =   2640
               TabIndex        =   16
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Procedure  Lines"
               Height          =   195
               Index           =   4
               Left            =   2640
               TabIndex        =   15
               Top             =   720
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Grid Lines"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   14
               Top             =   1080
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Procedure Name"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Component (If more than one)"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Project (If more than one)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraSettings 
            Caption         =   "Search History Size (20 - 200)"
            Height          =   1095
            Index           =   1
            Left            =   2795
            TabIndex        =   6
            Top             =   1770
            Width           =   2655
            Begin VB.PictureBox picCFXPBugFixSettings 
               BorderStyle     =   0  'None
               Height          =   840
               Index           =   2
               Left            =   100
               ScaleHeight     =   840
               ScaleWidth      =   2460
               TabIndex        =   7
               Top             =   175
               Width           =   2460
               Begin MSComCtl2.UpDown UpdHistory 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   97
                  Top             =   120
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   503
                  _Version        =   393216
                  Value           =   30
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtHistSize"
                  BuddyDispid     =   196642
                  OrigLeft        =   1200
                  OrigTop         =   480
                  OrigRight       =   1455
                  OrigBottom      =   735
                  Increment       =   10
                  Max             =   200
                  Min             =   20
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   0
                  Enabled         =   -1  'True
               End
               Begin VB.TextBox txtHistSize 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   720
                  TabIndex        =   96
                  Text            =   "30"
                  Top             =   120
                  Width           =   480
               End
               Begin VB.CheckBox ChkSaveHistory 
                  Caption         =   "Save"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   9
                  Top             =   600
                  Width           =   735
               End
               Begin VB.CommandButton cmdClearHistory 
                  Caption         =   "Clear"
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   8
                  Top             =   480
                  Width           =   735
               End
            End
         End
         Begin VB.CheckBox ChkLaunchStartup 
            Caption         =   "Launch On Startup"
            Height          =   195
            Left            =   5780
            TabIndex        =   5
            ToolTipText     =   "Find Tool appears in VB IDE during VB start up"
            Top             =   3085
            Width           =   1935
         End
      End
   End
   Begin MSComDlg.CommonDialog cdlSettings 
      Left            =   7440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip tbsSettingOuter 
      Height          =   4215
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find Tool"
            Key             =   "findtool"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Fixer Tools"
            Key             =   "cftools"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_FindSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' saftey value for frx problems detected
Private bRestoring                     As Boolean
'Safety switch for Apply and using control box to close
Private ApplyClicked                   As Boolean
'Preserves values at form load
Private OrigShowProj                   As Boolean
Private OrigShowComp                   As Boolean
Private OrigShowRout                   As Boolean
Private OrigLaunchStart                As Boolean
Private OrigbuttonOnBar                As Boolean
Private OrigRemFilt                    As Boolean
Private OrigSaveHist                   As Boolean
Private OrigbFindSelectWholeLine       As Boolean
Private OrighistDeep                   As Long
Private OrigElseIndent                 As Long
Private OrigCaseIndent                 As Long
Private OrigLineLength                 As Long
'
Private ColourTextForeUnDo             As Long
Private ColourTextBackUnDo             As Long
Private ColourFindSelectForeUnDo       As Long
Private ColourFindSelectBackUnDo       As Long
Private ColourHeadDefaultUnDo          As Long
Private ColourHeadWorkUnDo             As Long
Private ColourHeadPatternUnDo          As Long
Private ColourHeadNoFindUnDo           As Long
Private ColourHeadForeUnDo             As Long
Private ColourHeadReplaceUnDo          As Long

Private Function Chk2Bool(Chk As CheckBox) As Boolean

  'generic checkbox to boolean value

  Chk2Bool = Chk.Value = 1

End Function

Private Sub ChkAutoSelectedText_Click()

  bAutoSelectText = Chk2Bool(ChkLaunchStartup)

End Sub

Private Sub ChkLaunchStartup_Click()

  If Not bLoadingSettings Then
    bLaunchOnStart = Chk2Bool(ChkLaunchStartup)
  End If

End Sub

Private Sub ChkRemFilters_Click()

  If Not bLoadingSettings Then
    bRemFilters = Chk2Bool(ChkRemFilters)
  End If

End Sub

Private Function ChkReplace2Bool(intIndex As Long) As Boolean

  ChkReplace2Bool = ChkReplace(intIndex).Value = 1

End Function

Private Sub ChkReplace_Click(Index As Integer)

  'replace switches

  If Not bLoadingSettings Then
    bFilterWarning = ChkReplace2Bool(0)
    bBlankWarning = ChkReplace2Bool(1)
    bReplace2Search = ChkReplace2Bool(2)
  End If

End Sub

Private Sub ChkSaveHistory_Click()

  If Not bLoadingSettings Then
    bSaveHistory = Chk2Bool(ChkSaveHistory)
  End If

End Sub

Private Sub ChkSelectWhole_Click()

  If Not bLoadingSettings Then
    bFindSelectWholeLine = Chk2Bool(ChkSelectWhole)
  End If

End Sub

Private Function ChkShow2Bool(intIndex As Long) As Boolean

  ChkShow2Bool = ChkShow(intIndex).Value = vbChecked

End Function

Private Sub ChkShow_Click(Index As Integer)

  'Found Appearance switches

  If Not bLoadingSettings Then
    bShowProject = ChkShow2Bool(0)
    bShowComponent = ChkShow2Bool(1)
    bShowCompLineNo = ChkShow2Bool(2)
    bShowRoutine = ChkShow2Bool(3)
    bShowProcLineNo = ChkShow2Bool(4)
    bGridlines = ChkShow2Bool(5)
  End If

End Sub

Private Sub ChkTBarButton_Click()

  If Not bLoadingSettings Then
    bToolBarButton = Chk2Bool(ChkTBarButton)
    If Not bRestoring Then
      mObjDoc.Safe_MsgBox "Toolbar Button change will occur next time VB starts.", vbInformation
    End If
  End If

End Sub

Private Sub chkUser_Click(Index As Integer)

  Select Case Index
   Case 3
    If chkUser(3).Value = vbChecked Then
      If bModDescExists Then
        DoABackUp
      End If
    End If
   Case 5
    'a crash here means that cls_byteCheckBox has lost its magic
    'Reset 'value ' ProcedureID=Default
    SetTopMost Me, Xcheck(XStayOnTop)
  End Select

End Sub

Private Sub cmdClearHistory_Click()

  mObjDoc.ClearHistory

End Sub

Private Sub cmdOldReg_Click(Index As Integer)

  Select Case Index
   Case 0 'del old
    OldSettingsKill
   Case 1 'update old
    OldSettingsUpdate
  End Select

End Sub

Private Sub CmdSetting_Click(Index As Integer)

  ' OK/Apply/Cancel buttons

  ApplyClicked = False
  Select Case Index
   Case 0
    SaveUserSettings
    mObjDoc.ApplyChanges
    ApplyClicked = True
    Me.Hide
   Case 1
    SaveUserSettings
    mObjDoc.ApplyChanges
    ApplyClicked = True
   Case 2
    RestoreOriginals
    Me.Hide
  End Select
  SaveFormPosition Me

End Sub

Private Sub cmdStartFromSettings_Click()

  frm_FindSettings.Hide
  mObjDoc.DoFixes

End Sub

Private Sub ctlSettingsList_Change(Index As Integer)

  Dim arrTmp As Variant

  If Len(UserSettings) Then
    arrTmp = Split(UserSettings, "|")
    arrTmp(Index) = ctlSettingsList(Index).DataString
    UserSettings = Join(arrTmp, "|")
  End If

End Sub

Private Sub DoFixSettingClick()

  Dim I As Long

  SetTopMost Me, True
  With tbsSettings
    For I = 0 To 7
      fraSubSetting(I).Visible = .SelectedItem.Index - 1 = I
    Next I
    fraSubSetting(.SelectedItem.Index - 1).Move .ClientLeft, .ClientTop
    If .SelectedItem.Index > 1 Then
    End If
    fraSubSetting(.SelectedItem.Index - 1).Visible = True
    For I = 0 To 6
      SettingsEngine I, UserSettings
    Next I
  End With

End Sub

Private Sub DoSetClick()

  Dim CFrame As Frame
  Dim I      As Long

  Set CFrame = fraprop(tbsSettingOuter.SelectedItem.Index - 1)
  With Me
    .Width = CFrame.Width
    .Height = CFrame.Height + CmdSetting(0).Height * 4
    tbsSettingOuter.Move 0, 0, CFrame.Width, CFrame.Height + CmdSetting(0).Height
    .Height = CFrame.Height + CmdSetting(0).Height * 3.5
    .Refresh
  End With
  For I = 0 To fraprop.Count - 1
    fraprop(I).Caption = vbNullString
    fraprop(I).Visible = False
  Next I
  CFrame.Move 0, tbsSettingOuter.ClientTop
  CFrame.Visible = True
  CmdSetting(0).Left = (Me.Width - (CmdSetting(0).Width * 3)) / 2
  CmdSetting(1).Left = CmdSetting(0).Left + CmdSetting(1).Width
  CmdSetting(2).Left = CmdSetting(1).Left + CmdSetting(1).Width
  CmdSetting(0).Top = tbsSettingOuter.Height + tbsSettingOuter.Top
  CmdSetting(1).Top = tbsSettingOuter.Height + tbsSettingOuter.Top
  CmdSetting(2).Top = tbsSettingOuter.Height + tbsSettingOuter.Top
  '  If tbsSettingOuter.SelectedItem.Index - 1 = 1 Then
  DoFixSettingClick
  '  End If

End Sub

Private Sub Form_Activate()

  On Error Resume Next
  HideInitiliser
  'set safety values for Cancel button
  OrigbFindSelectWholeLine = bFindSelectWholeLine
  OrigShowProj = bShowProject
  OrigShowComp = bShowComponent
  OrigShowRout = bShowRoutine
  OrigLaunchStart = bLaunchOnStart
  OrigbuttonOnBar = bToolBarButton
  OrigLineLength = lngLineLength
  OrigElseIndent = lngElseIndent
  OrigCaseIndent = lngCaseIndent
  '
  OrigRemFilt = bRemFilters
  OrigSaveHist = bSaveHistory
  OrighistDeep = HistDeep
  DoSetClick
  On Error GoTo 0

End Sub

Private Sub Form_Deactivate()

  DoSetClick

End Sub

Private Sub Form_Load()

  Dim fra As Frame

  On Error Resume Next
  If Not (bInitializing Or bAddinTerminate) Then
    bVeryLargeMsgShow = False
    If LaunchTool Then
      HideInitiliser
      'SliderHistory.LargeChange = 20
      For Each fra In fraprop
        fra.Caption = vbNullString
        fra.BorderStyle = 0
      Next fra
      LoadFormPosition Me
      mObjDoc.ColoursApply
      ColourTextForeUnDo = ColourTextFore
      ColourTextBackUnDo = ColourTextBack
      ColourFindSelectForeUnDo = ColourFindSelectFore
      ColourFindSelectBackUnDo = ColourFindSelectBack
      ColourHeadDefaultUnDo = ColourHeadDefault
      ColourHeadWorkUnDo = ColourHeadWork
      ColourHeadPatternUnDo = ColourHeadPattern
      ColourHeadNoFindUnDo = ColourHeadNoFind
      ColourHeadForeUnDo = ColourHeadFore
      ColourHeadReplaceUnDo = ColourHeadReplace
      Me.Caption = "Properties " & AppDetails
      On Error GoTo 0
    End If
    DoSetClick
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Not ApplyClicked Then
    ' keeps changes if user clicks 'Apply' then uses CaptionBar 'X' button to close
    'otherwise restore
    RestoreOriginals
  End If
  SaveFormPosition Me
  Me.Hide

End Sub

Private Sub LblColour_Click(Index As Integer)

  Select Case Index
   Case 11 'Standard colours
    mObjDoc.ColoursStandard
   Case 10 'Undo
    ColourTextFore = ColourTextForeUnDo
    ColourTextBack = ColourTextBackUnDo
    ColourFindSelectFore = ColourFindSelectForeUnDo
    ColourFindSelectBack = ColourFindSelectBackUnDo
    ColourHeadDefault = ColourHeadDefaultUnDo
    ColourHeadWork = ColourHeadWorkUnDo
    ColourHeadPattern = ColourHeadPatternUnDo
    ColourHeadNoFind = ColourHeadNoFindUnDo
    ColourHeadFore = ColourHeadForeUnDo
    ColourHeadReplace = ColourHeadReplaceUnDo
   Case Else
    With cdlSettings
      .Flags = cdlCCRGBInit Or cdlCCFullOpen
      'set current color as default
      Select Case Index
       Case 0
        .Color = ColourTextFore
       Case 1
        .Color = ColourTextBack
       Case 2
        .Color = ColourFindSelectFore
       Case 3
        .Color = ColourFindSelectBack
       Case 4
        .Color = ColourHeadFore
       Case 5
        .Color = ColourHeadDefault
       Case 6
        .Color = ColourHeadWork
       Case 7
        .Color = ColourHeadPattern
       Case 8
        .Color = ColourHeadNoFind
       Case 8
        .Color = ColourHeadReplace
      End Select
      .ShowColor
      'apply new or default colour
      If Not .CancelError Then
        Select Case Index
         Case 0
          ColourTextFore = .Color
         Case 1
          ColourTextBack = .Color
         Case 2
          ColourFindSelectFore = .Color
         Case 3
          ColourFindSelectBack = .Color
         Case 4
          ColourHeadFore = .Color
         Case 5
          ColourHeadDefault = .Color
         Case 6
          ColourHeadWork = .Color
         Case 7
          ColourHeadPattern = .Color
         Case 8
          ColourHeadNoFind = .Color
         Case 9
          ColourHeadReplace = .Color
        End Select
      End If
    End With
  End Select
  mObjDoc.ColoursApply

End Sub

Private Sub lsvModNames_ItemCheck(ByVal Item As MSComctlLib.ListItem)

  Dim I As Long

  lsvModNames.SelectedItem = lsvModNames.ListItems.Item(Item.Index)
  Item.Checked = Item.Checked And Item.ListSubItems(1).Tag <> "ReadOnly"
  For I = 1 To UBound(ModDesc)
    ModDesc(I).MDDontTouch = UnTouchable(ModDesc(I).MDName)
  Next I

End Sub

Private Sub OptCaseIndent_Click(Index As Integer)

  lngCaseIndent = Index

End Sub

Private Sub OptElseIndent_Click(Index As Integer)

  lngElseIndent = Index

End Sub

Private Sub optInputPos_Click(Index As Integer)

  If Not bLoadingSettings Then
    bLocTop = optInputPos(0).Value = True
    mObjDoc.DoResize
  End If

End Sub

Private Sub optLevelSetting_Click(Index As Integer)

  StandardSettings CLng(Index)

End Sub

Private Sub RestoreOriginals()

  If Not bAddinTerminate Then
    bRestoring = True 'stops warning message from button on toolbar checkbox
    ChkSelectWhole.Value = Bool2Lng(OrigbFindSelectWholeLine)
    ChkRemFilters.Value = Bool2Lng(OrigRemFilt)
    ChkLaunchStartup.Value = Bool2Lng(OrigLaunchStart)
    ChkTBarButton.Value = Bool2Lng(OrigbuttonOnBar)
    OptElseIndent(OrigElseIndent).Value = True
    OptCaseIndent(OrigCaseIndent).Value = True
    UpdLongLine.Value = OrigLineLength
    UpdHistory.Value = OrighistDeep
    ChkShow(0).Value = Bool2Lng(OrigShowProj)
    ChkShow(1).Value = Bool2Lng(OrigShowComp)
    ChkShow(2).Value = Bool2Lng(OrigShowRout)
    ChkSaveHistory.Value = Bool2Lng(OrigSaveHist)
    bRestoring = False
  End If

End Sub

Private Sub tbsSettingOuter_Click()

  DoSetClick

End Sub

Private Sub tbsSettings_Click()

  DoFixSettingClick

End Sub

Private Sub tbsSettings_GotFocus()

  DoFixSettingClick

End Sub

Private Sub txtHistSize_Change()

  With txtHistSize
    If Val(.Text) >= 60 Then
      If Val(.Text) <= 150 Then
        HistDeep = Val(.Text)
      End If
    End If
  End With 'txtHistSize

End Sub

Private Sub txtLongLine_Change()

  With txtLongLine
    If Val(.Text) >= 60 Then
      If Val(.Text) <= 150 Then
        lngLineLength = Val(.Text)
      End If
    End If
  End With 'txtLongLine

End Sub

Private Sub txtLongProc_Change()

  With txtLongProc
    If Val(.Text) >= 40 Then
      If Val(.Text) <= 150 Then
        lngBigProcLines = Val(.Text)
      End If
    End If
  End With 'txtLongProc

End Sub

Private Sub UpdHistory_Change()

  HistDeep = UpdHistory.Value

End Sub

Private Sub UpdLongLine_Change()

  lngLineLength = UpdLongLine.Value

End Sub

Private Sub UpDLongProc_Change()

  lngBigProcLines = UpDLongProc.Value

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:11:17 AM) 33 + 507 = 540 Lines Thanks Ulli for inspiration and lots of code.
