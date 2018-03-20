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
	Call MayorEdad(20)
	Call Multiplicar(2,5)
	Call MostrarMensaje(True, "Esto es una cadena de texto")

End Sub

'Ejercicios
'Crea un procedimiento que utiliza un procedimiento dentro de otro procedimiento
'Crea un procedimiento dependiendo de una condición, utilizando una estructura de control if - then
'Crea un procedimiento que recibe parámetros numéricos y los multiplica
'Crea un procedimiento que recibe un parámetro de tipo boolean

Sub MayorEdad(edad As Byte)
'
' Muestra un mensaje en función del parámetro que recibe
'
	If edad >=18 then
		MsgBox "Eres MAYOR de edad"
	else
		MsgBox "Eres MENOR de edad"
	End if

End Sub

Sub Multiplicar(n1 As Byte, n2 As Byte)
'
' Muestra un mensaje con el resultado de multiplicar el contenidos de los parametros
'
	MsgBox n1 & " x " & n2 & " = " & n1*n2

End Sub
 
Sub MostrarMensaje(bandera As Boolean, mensaje As String)

	If bandera Then
		MsgBox mensaje
	Else
		MsgBox "La bandera no me permite mostrar el mensaje."
	End If

End Sub
 
 
 
 
 
 
 
 
 
 
 
 
 
