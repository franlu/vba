Attribute VB_Name = "EstructuraCondicional"

Sub Ejemplo1()
' Ejemplo de uso de la estructura condicional simple
	Dim a As Integer
    a = 12
    If a = 8 Then
        MsgBox "El valor de a es igual a 8"
    End If
	MsgBox "Hemos llegado al final del módulo"
End Sub

Sub Ejemplo2()
' Ejemplo de uso de la estructura condicional
    Dim a As Integer
	a = 10
    If a = (8+2) Then
        MsgBox "a es IGUAL a 8"
    Else
        MsgBox "a es DISTINTO de 8"
    End If
End Sub

Sub Ejemplo3()
' Ejemplo de uso de la estructura condicional anidada
	Dim a As Integer
    a = 15
    If a = 8 Then
        MsgBox "a es IGUAL a 8"
    ElseIf a = 12 Then
        MsgBox "a es IGUAL a 12"
    Else
        MsgBox "a es DISTINTO de 8 y de 12"
    End If
End Sub

Sub RecogidaTexto()
'
'       Solicita al usuario una cadena de texto
'       la escribe y hace un retorno de carro (Intro)
'
        texto = InputBox("¿Qué quieres que escriba?", "Recogida de texto")
        Selection.TypeText Text:=texto
        Selection.TypeParagraph
End Sub

Sub MayorEdad()
'
'       Muestra un mensaje diciendo si eres mayor de edad o no
'
        Dim Edad As Byte
        Edad = InputBox("¿Qué edad tienes?", "RECOGIDA DE EDAD")
        
        If Edad >= 18 Then
                MsgBox("Eres MAYOR de edad.")
        Else
                MsgBox("Eres MENOR de edad.")
        End If
        
End Sub

Sub Edad()
'
'       Muestra un mensaje en función de la edad introducida.
'
        Dim Edad As Integer
        Edad = InputBox("¿Qué edad tienes?", "RECOGIDA DE EDAD")
        
        If Edad < 13 Then MsgBox ("Eres Niño/a")
        If Edad >= 13 And Edad < 18 Then MsgBox ("Eres Adolescente")
        If Edad >= 18 And Edad < 30 Then MsgBox ("Eres Joven")
        If Edad >= 30 And Edad < 65 Then MsgBox ("Eres Adulto/a")
        If Edad >= 65 And Edad < 100 Then MsgBox ("Eres Jubilado/a")
        If Edad >= 100 Then MsgBox ("ERES MATUSALEN")
        
End Sub

Sub Edad2()
'	Muestra un mensaje en función de la edad introducida.
'	Se utiliza la estructura condicional anidada.
End Sub