VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_HourGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Source I think VBPJ
'''Set and forget hourglass handling
'''Call this class by placing the following lines
'''at the top of a procedure:
''Dim MyhourGlass as CHourGlass
''Set MyHourGlass= new chourglass
'''It works by triggering Class_Initialize when it gets scope
'''and Class_Terminate when the routine goes out of scope
Option Explicit

Private Sub Class_Initialize()

  Screen.MousePointer = vbHourglass

End Sub

Private Sub Class_Terminate()

  Screen.MousePointer = vbDefault

End Sub

':)Code Fixer V3.0.9 (25/03/2005 4:11:44 AM) 9 + 15 = 24 Lines Thanks Ulli for inspiration and lots of code.
