Attribute VB_Name = "EstructuraCondicional"

Const NINIO As Byte = 13
Const ADOLESCENTE As Byte = 18
Const JOVEN As Byte = 30
Const ADULTO As Byte = 65
Const JUBILADO As Byte = 100

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
    If a = (8 + 2) Then
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
                MsgBox ("Eres MAYOR de edad.")
        Else
                MsgBox ("Eres MENOR de edad.")
        End If
        
End Sub

Sub Edad()
'
'Muestra un mensaje en función de la edad introducida.
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


'       Muestra un mensaje en función de la edad introducida.
'       Se utiliza la estructura condicional anidada.
        
        Dim Edad As Integer
        Edad = InputBox("¿Qué edad tienes?", "RECOGIDA DE EDAD")
        
        If Edad < 13 Then
                MsgBox ("Eres Niño/a")
        ElseIf Edad < 18 Then
                MsgBox ("Eres Adolescente")
        ElseIf Edad < 30 Then
                MsgBox ("Eres Joven")
        ElseIf Edad < 65 Then
                MsgBox ("Eres Adulto/a")
        ElseIf Edad < 100 Then
                MsgBox ("Eres Jubilado/a")
        Else
                MsgBox ("ERES MATUSALEN")
        End If
End Sub

Sub Edad3()
'       Muestra un mensaje en función de la edad introducida.
'       Se utiliza la estructura condicional anidada.
'       Utilizamos constantes
        
        Dim Edad As Integer
        Edad = InputBox("�Qu� edad tienes?", "RECOGIDA DE EDAD")
        
        If Edad < NINIO Then
                MsgBox ("Eres Ni�oo/a")
        ElseIf Edad < ADOLESCENTE Then
                MsgBox ("Eres Adolescente")
        ElseIf Edad < JOVEN Then
                MsgBox ("Eres Joven")
        ElseIf Edad < ADULTO Then
                MsgBox ("Eres Adulto/a")
        ElseIf Edad < JUBILADO Then
                MsgBox ("Eres Jubilado/a")
        Else
                MsgBox ("ERES MATUSALEN")
        End If
End Sub

Sub NumeroMayor()
'
' Compara el contenido de dos variables y muestra un mensaje
' indicando la variable que contiene el mayor valor.
'

Dim numero1 As Integer
Dim numero2 As Integer

' Mensaje de aviso al usuario
MsgBox ("Los valores permitidos son de -32768 a 32767")

' Recoger valores que introduce el usuario
numero1 = InputBox("Escribe un n�mero", "Calcular Mayor")
numero2 = InputBox("Escribe un n�mero", "Calcular Mayor")

'Calcular cual es el mayor de los dos numeros
If numero1 > numero2 Then
    MsgBox ("El primer valor introducido es el mayor.")
ElseIf numero1 = numero2 Then
    MsgBox ("Los valores son iguales.")
Else
    MsgBox ("El segundo valor introducido es el mayor.")
End If

End Sub






























