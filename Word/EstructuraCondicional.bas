Attribute VB_Name = "EstructuraCondicional"
Sub Ejemplo1()
    a = 12
    If a = 8 Then
        MsgBox "a es igual a 8"
    End If
End Sub
Sub Ejemplo2()
    a = 12
    If a = 8 Then
        MsgBox "a es IGUAL a 8"
    Else
        MsgBox "a es DISTINTO de 8"
    End If
End Sub

Sub Ejemplo3()
    a = 12
    If a = 8 Then
        MsgBox "a es IGUAL a 8"
    ElseIf a = 12 Then
        MsgBox "a es IGUAL a 12"
    Else
        MsgBox "a es DISTINTO de 8 y de 12"
    End If
End Sub
