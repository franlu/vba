Attribute VB_Name = "Procedimientos"
Sub ProcedimientoMensaje()

    MsgBox "Esto es un procedimiento sin parámetros"

End Sub

Sub DatosPersonales(nombre As String, edad As Integer, ciudad As String)
    
    MsgBox nombre & " " & edad & " " & ciudad
    
End Sub

Sub UsarProcedimiento()

    DatosPersonales "Eva", 25, "Granada"

End Sub

Sub UsarProcedimiento1()
'
' Ejemplo de uso de la función Call
'

    Call DatosPersonales("Eva", 25, "Granada")

End Sub


'Ejercicios
'Crea un procedimiento que utiliza un procedimiento dentro de otro procedimiento
'Crea un procedimiento dependiendo de una condición, utilizando una estructura de control if - then
'Crea un procedimiento que recibe parámetros numéricos y los multiplica
'Crea un procedimiento que recibe un parámetro de tipo boolean
 
