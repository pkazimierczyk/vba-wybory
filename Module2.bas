Attribute VB_Name = "Module2"
Option Explicit

Sub czy_liczba(dana, komunikat)
    
If dana = "" Then
    Exit Sub
End If

If Not IsNumeric(dana) Then
    MsgBox (komunikat)
    dana.Value = ""
Else
    If Int(dana) - (dana) <> 0 Or dana < 0 Then
    MsgBox (komunikat)
    dana.Value = ""
    End If
End If

End Sub
