Attribute VB_Name = "Procedimientos"
Sub ProcedimientoMensaje()

    MsgBox "Esto es un procedimiento sin par�metros"

End Sub

Sub DatosPersonales(nombre As String, edad As Integer, ciudad As String)
    
    MsgBox nombre & " " & edad & " " & ciudad
    
End Sub

Sub UsarProcedimiento()

    DatosPersonales "Eva", 25, "Granada"

End Sub

Sub UsarProcedimiento1()
'
' Ejemplo de uso de la funci�n Call
'

    Call DatosPersonales("Eva", 25, "Granada")

End Sub


'Ejercicios
'Crea un procedimiento que utiliza un procedimiento dentro de otro procedimiento
'Crea un procedimiento dependiendo de una condici�n, utilizando una estructura de control if - then
'Crea un procedimiento que recibe par�metros num�ricos y los multiplica
'Crea un procedimiento que recibe un par�metro de tipo boolean
 
