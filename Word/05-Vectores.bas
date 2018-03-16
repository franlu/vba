Attribute VB_Name = "Vectores"
Sub MiPrimerArray()
' Ejemplo de uso de un vector
    
    'Declaracion
    Dim miArray(3) As String
    
    'Asignacion
    miArray(0) = "lunes"
    miArray(1) = "martes"
    miArray(2) = "miércoles"

    'Acceso
    MsgBox miArray(0) & " " & miArray(1) & " " & miArray(2)
	
	'Limites
	MsgBox UBound(miArray)
	MsgBox LBound(miArray)

End Sub

Sub redimensionarArray()
'Redimensiona un vector y muestra su contenido en un mensaje
    
	Dim Meses(3) As String
	
	Meses(0) = "Enero"
    Meses(1) = "Febrero"
    Meses(2) = "Marzo"
    
	MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
	
	ReDim Meses(12)
    MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)

End Sub

Sub PreservarArray()
'Rediminsiona un vector preservando su contenido

    Dim Meses(3) As String
	Meses(0) = "Enero"
    Meses(1) = "Febrero"
    Meses(2) = "Marzo"
    
	MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
	
	ReDim Preserve(12)
	Meses(3) = "Abril"
    Meses(4) = "Mayo"
    Meses(5) = "Junio"
	Meses(6) = "Julio"
    Meses(7) = "Agosto"
    Meses(8) = "Septiembre"
	Meses(9) = "Octubre"
    Meses(10) = "Noviembre"
    Meses(11) = "Diciembre"
	
    MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
	MsgBox Meses(3) & " " & Meses(4) & " " & Meses(5)
End Sub

Sub recorrerArray()
'
'Imprime el contenido de un vector en un documento de word
'
	Dim Andalucia(8) As String
	
	Andalucia(0) = "Almeria"
	Andalucia(1) = "Cádiz"
	Andalucia(2) = "Cordoba"
	Andalucia(3) = "Granada"
	Andalucia(4) = "Huelva"
	Andalucia(5) = "Jaen"
	Andalucia(6) = "Málaga"
	Andalucia(7) = "Sevilla"
	
	For i=0 To UBound(Andalucia)
		Selection.TypeText Text:=Andalucia(i)
	    Selection.TypeParagraph
	Next
End Sub