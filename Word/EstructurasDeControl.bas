Attribute VB_Name = "EstructurasDeControl"
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

Sub ForToNext()
  Dim fila As Integer
  For contador = 1 To 10
    fila = contador
  Next
  MsgBox "Se alcanzó el valor " & fila
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