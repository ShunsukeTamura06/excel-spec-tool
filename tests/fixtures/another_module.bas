Attribute VB_Name = "Module2"
Option Explicit

Public Sub UpdateDaily()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub

Private Function IsValid(s As String) As Boolean
    IsValid = Len(s) > 0
End Function
