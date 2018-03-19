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

Sub RellenarAleatorio()
'
' Rellena el vector con numeros aleatorios, tamaño del vector lo indica el usuario.
' Uso de la funcion Rnd
'
	Dim tamanio As Byte
	tamanio = InputBox("Escribe el tamaño del vector (entre 0 y 255)", "Rellenar Vector")
	
	ReDim vector(tamanio) As Double
       
    For i = 0 To UBound(vector)
        vector(i) = Rnd
        MsgBox "Posición: " & i & " y su " & "contenido es: " & vector(i)
    Next
		
End Sub

Sub Euromillon()
'Genera una combianción de numeros aleatorios  para el juego del euromillon.
'Muestra un mensaje por cada número.
 Dim numeros(4) As Byte
 Dim estrellas(1) As Byte
  
 For i=0 To Ubound(numeros)
	numeros(i) = (50 * Rnd)+1
	MsgBox numeros(i)
 Next
 
 mensaje = mensaje & " Estrellas: "
 For i=0 To Ubound(estrellas)
	estrellas(i) = (12 * Rnd)+1
	MsgBox estrellas(i)
 Next
 
End Sub

Sub Euromillon1()
'
'Genera una combiancion para el juego del euromillon.
'La combinación se muestra en la misma ventana de mensaje.
'
 Dim numeros(4) As Byte
 Dim estrellas(1) As Byte
 Dim mensaje As String 
 
	 mensaje = "Números: "
	 For i=0 To Ubound(numeros)
		numeros(i) = (50 * Rnd)+1
		mensaje = mensaje & numeros(i) & " "
	 Next
	 
	 mensaje = mensaje & " Estrellas: "
	 For i=0 To Ubound(estrellas)
		estrellas(i) = (12 * Rnd)+1
		mensaje = mensaje & estrellas(i) & " "
	 Next
	 
 MsgBox mensaje
 
End Sub

Sub Buscar()
'Rellena un vector con numeros aleatorios. Pide al usuario un valor, lo busca dentro del vector.
'Muestra un mensaje indicando en la posicion que se encuentra.
End Sub


