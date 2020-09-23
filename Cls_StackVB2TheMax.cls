VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_StackVB2TheMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''CStack - a class module for implementing
''Last-In-First-Out (stack) structures
''Date: 7/21/2001
''Versions: VB5 VB6 Level: Intermediate
''Author: The VB2TheMax Team
'--------------------------------------------------
'
'  The CSTACK class
'
' Usage:
'    Dim st As New CStack
'
'    ' push two values on the stack
'    st.Push 1234
'    st.Push 4567
'    ' display number of elements in the stack
'    Debug.Print "Count = " & st.Count
'    ' peek at the element on top of stack
'    Debug.Print "Peek = " & st.Peek
'    ' print the elements on the stack while popping them
'    Do While st.Count
'        Debug.Print "TOS item = " & st.Pop
'    Loop
'
'---------------------------------------------------
' this is the collection that holds the values
Private colValues     As New Collection

Public Function Count() As Long

  ' Return the number of values in the stack

  Count = colValues.Count

End Function

Public Function Pop() As Variant

  ' Pop a value off the stack - raise error if stack is emtpy

  Pop = colValues.Item(colValues.Count)
  colValues.Remove colValues.Count

End Function

Public Sub Push(lngVal As Variant)

  ' add a new value onto the stack

  colValues.Add lngVal

End Sub

''Public Function Peek() As Variant
''' Return the value on top of the stack, without popping it
''' raise error if stack is empty
''Peek = colValues.Item(colValues.Count)
''End Function
''

':)Code Fixer V3.0.9 (25/03/2005 4:11:34 AM) 28 + 35 = 63 Lines Thanks Ulli for inspiration and lots of code.
