Attribute VB_Name = "01"

Sub BucleFor()
'
'       Solicita al usuario un numero y un texto
'       a continuacion escribe el texto el numero de veces
'       que indique el usuario.
'
	Dim numero As Byte
	Dim texto As String
	
	numero = InputBox("�D� el numero de veces que quieres que lo escriba?", "N� de Veces")
	texto = InputBox("�Qu� quieres que escriba?", "Recogida de texto")
		
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
        numero = InputBox("�D� el numero de veces que quieres que lo escriba?", "N� de Veces")
        texto = InputBox("�Qu� quieres que escriba?", "Recogida de texto")
                For I = 1 To numero
                Selection.TypeText Text:=texto
                Selection.TypeParagraph
                Next
                MsgBox ("Acaba de finalizar la macro.")
End Sub


