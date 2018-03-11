Attribute VB_Name = "Vectores"
Sub MiPrimerArray()
    
    'Declaracion
    Dim miArray(3) As String
    
    'Asignacion
    miArray(0) = "lunes"
    miArray(1) = "martes"
    miArray(2) = "miércoles"

    'Acceso
    MsgBox miArray(0) & " " & miArray(1) & " " & miArray(2)

End Sub

Sub redimensionarArray()
'Redimensiona el array incluir el jueves y el viernes
    
    ReDim miArray(5)
    miArray(0) = "lunes"
    miArray(1) = "martes"
    miArray(2) = "miércoles"
    miArray(3) = "jueves"
    miArray(4) = "viernes"

    MsgBox miArray(0) & " " & miArray(1) & " " & miArray(2) & " " & miArray(3) & " " & miArray(4)

End Sub

Sub PreservarArray()

    ReDim miArray(3)
    miArray(0) = "lunes"
    miArray(1) = "martes"
    miArray(2) = "miércoles"

    ReDim Preserve miArray(3)
    miArray(3) = "jueves"
    miArray(4) = "viernes"

    MsgBox miArray(0) & " " & miArray(1) & " " & miArray(2) & " " & miArray(3) & " " & miArray(4)
End Sub
