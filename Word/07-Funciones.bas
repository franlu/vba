Attribute VB_Name = "Funciones"

Function Resta(Valor1 As Integer, Valor2 As Integer) As Long
'
' Resta los valores pasados como parametros y devuelve el resultado
'
    Resta = Valor1 - Valor2
    
End Function

Function Suma(Valor1 As Integer, Valor2 As Integer) As Long
'
' Suma dos valores pasados como parametros y devuelve el resultado
'
    Suma = Valor1 + Valor2
    
End Function

Sub MostrarTotal()
'
' Ejemplo de uso de una funci�n
'
    Dim a As Integer
    Dim b As Integer
        
    a = 1
    b = 1
        
    MsgBox "La suma es " & Suma(a, b), vbOKOnly + vbInformation
	MsgBox "La suma es " & Suma(a, b), vbOKOnly + vbExclamation
	MsgBox "La suma es " & Suma(a, b), vbOKOnly + vbCritical
	MsgBox "La suma es " & Suma(a, b), vbOKOnly + vbQuestion
	
End Sub

Function Revision(AnioFab As Integer) As String
	If AnioFab > Year(Now) Then
		Revision = "A�o incorrecto."
	ElseIf AnioFab <= Year(Now) - 3 Then
		Revision = "S�"
	Else
		Revision = "No"
	End If
End Function

Sub UsarFunciones()

	MsgBox Resta(10,5)
	MsgBox Suma(2,2)
	
	Dim resultado As Long
	resultado = Resta(5,3)
	MsgBox resultado
	
End Sub

'Ejercicios
'1. Funci�n de tipo integer que recibe 2 par�metros
'2  Funci�n de tipo boolean que es ejecutada de acuerdo a una condici�n If.
 