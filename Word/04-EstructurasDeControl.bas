Attribute VB_Name = "EstructurasDeControl"

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

Sub SelectionWith()
'
'       Solicita al usuario una cadena texto y establece
'       las propiedades del objeto Selection antes de escribirlo
'       en el documento.
'
	Dim texto As String
	
	texto = InputBox("¿Qué quieres que escriba?", "Recogida de texto")
	
	With Selection
		.Font.Bold = True
		.Font.Name = "Arial"
		.Font.ColorIndex = wdDarkBlue
		.ParagraphFormat.Alignment = wdAlignParagraphCenter
		.ParagraphFormat.SpaceAfter = 0
	End With
	
	Selection.TypeText Text:=texto
	Selection.TypeParagraph
	
End Sub

Sub ForToNext()
  Dim fila As Integer
  For contador = 1 To 10
    fila = contador
  Next
  MsgBox "Se alcanzó el valor " & fila
End Sub

Sub DoLoopUntil()
'
' DoLoopUntil Macro
'
'
  Dim contador As Integer
  Dim numero As Integer
  numero = 9
  Do Until numero = 10
   If numero = 0 Then Exit Do
     numero = numero - 1
     contador = contador + 1
   Loop
  MsgBox "Se alcanzó el valor " & numero & " " & contador

End Sub

Sub DoWhileLoop()
  Dim escribir As Integer
  escribir = 1
  Do While escribir = 7
    MsgBox "Escribir = " & escribir
    escribir = escribir + 1
  Loop
  MsgBox "Se acabo el bucle, Escribir = " & escribir
End Sub

Sub WhileWend()
  Dim a As Integer
  a = 0
  While (a < 13)
    a = a + 1
  Wend
  MsgBox "Se alcanzó el valor " & a
End Sub

Sub InfiniteLoop()
	Dim x As Integer
	x = 1
	Do
		Application.StatusBar = _
		"La aplicación ha entrado en un bucle infinito: " & x
		x = x + 1
	Loop
End Sub