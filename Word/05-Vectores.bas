Attribute VB_Name = "Vectores"
Sub MiPrimerArray()
    
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
'Redimensiona el array
    
	Dim Meses(3) As String
	Meses(0) = "Enero"
    Meses(1) = "Febrero"
    Meses(2) = "Marzo"
    
	MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
	
	ReDim Meses(12)
    MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)

End Sub

Sub PreservarArray()

    Dim Meses(3) As String
	Meses(0) = "Enero"
    Meses(1) = "Febrero"
    Meses(2) = "Marzo"
    
	MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
	
	ReDim Preserve(12)
    MsgBox Meses(0) & " " & Meses(1) & " " & Meses(2)
End Sub

Sub recorrerArray()
'Recorrer Array

End Sub