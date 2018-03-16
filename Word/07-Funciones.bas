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

Sub CalcularTotal()
'
' Ejemplo de uso de una función
'
    Dim a As Integer
    Dim b As Integer
        
    a = 1
    b = 1
        
    MsgBox "La suma es " & Suma(a, b)

End Sub

'Ejercicios
'1. Función de tipo integer que recibe 2 parámetros
'2  Función de tipo boolean que es ejecutada de acuerdo a una condición If.
 