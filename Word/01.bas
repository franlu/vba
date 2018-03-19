Attribute VB_Name = "01"

Sub BucleFor()
'
'       Solicita al usuario un numero y un texto
'       a continuacion escribe el texto el numero de veces
'       que indique el usuario.
'
	Dim numero As Byte
	Dim texto As String
	
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


