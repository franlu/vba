Attribute VB_Name = "01"
Sub RecogidaTexto()
'
'       Solicita al usuario una cadena de texto
'       la escribe y hace un retorno de carro (Intro)
'
        texto = InputBox("¿Qué quieres que escriba?", "Recogida de texto")
        Selection.TypeText Text:=texto
        Selection.TypeParagraph
End Sub

Sub SentenciaIF()
'
'       Muestra un mensaje diciendo si eres mayor de edad o no
'
        Dim Edad As Integer
        Edad = InputBox("¿Qué edad tienes?", "RECOGIDA DE EDAD")
        
        If Edad < 18 Then
                MsgBox ("Eres MENOR de edad.")
        Else
                MsgBox ("Eres MAYOR de edad.")
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

Sub BucleFor()
'
'       Solicita al usuario un numero y un texto
'       a continuacion escribe el texto el numero de veces
'       que indique el usuario.
'
        numero = InputBox("¿Dí el numero de veces que quieres que lo escriba?", "Nº de Veces")
        texto = InputBox("¿Qué quieres que escriba?", "Recogida de texto")
                For I = 1 To numero
                Selection.TypeText Text:=texto
                Selection.TypeParagraph
                Next
End Sub

Sub MostrarMensaje()
'
'       Solicita al usuario un numero y un texto
'       a continuacion escribe el texto el numero de veces
'       que indique el usuario.
'
        numero = InputBox("¿Dí el numero de veces que quieres que lo escriba?", "Nº de Veces")
        texto = InputBox("¿Qué quieres que escriba?", "Recogida de texto")
                For I = 1 To numero
                Selection.TypeText Text:=texto
                Selection.TypeParagraph
                Next
                MsgBox ("Acaba de finalizar la macro.")
End Sub

Sub TransponerPalabraDerecha()
'
'       Traspone una palabra hacia la derecha
'
        'Tecla F8 selecciona una palabra
        Selection.Extend
        Selection.Extend
        Selection.EscapeKey
        Selection.Cut
        Selection.MoveRight Unit:=wdWord, Count:=1
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Selection.MoveLeft Unit:=wdWord, Count:=1
End Sub
