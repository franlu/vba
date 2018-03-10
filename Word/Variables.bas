Attribute VB_Name = "Variables"

Sub Variables()
'
'Declara una variable de cada tipo, le asigna un valor
'y lo imprime en el documento actual
'

Dim bool As Boolean 'True o False
Dim by As Byte ' 0 y 255
Dim fecha As Date '#01/01/1900#
Dim entero As Integer ' -32768 a 32767
Dim enteroLargo As Long ' 2.147.483.648 a 2.147.483.647
Dim doble As Double '-1.79769313486232e308 a 1.79769313486232e308
Dim cadena As String

bool = False
by = 255
fecha = #1/1/2018#
entero = 4069
enteroLargo = 20000000
doble = 1.79769313486232E+300
cadena = "Esto es una cadena de texto"

Selection.TypeText Text:=bool
Selection.TypeParagraph

Selection.TypeText Text:=by
Selection.TypeParagraph

Selection.TypeText Text:=fecha
Selection.TypeParagraph

Selection.TypeText Text:=entero
Selection.TypeParagraph

Selection.TypeText Text:=enteroLargo
Selection.TypeParagraph

Selection.TypeText Text:=doble
Selection.TypeParagraph

Selection.TypeText Text:=cadena
Selection.TypeParagraph

MsgBox entero
MsgBox cadena

End Sub
