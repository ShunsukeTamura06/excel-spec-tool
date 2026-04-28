Attribute VB_Name = "Module1"
Option Explicit

Sub Hello()
    MsgBox "Hello, world"
End Sub

Function Square(x As Long) As Long
    Square = x * x
End Function

Property Get TheAnswer() As Integer
    TheAnswer = 42
End Property
