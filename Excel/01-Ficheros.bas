Attribute VB_Name = "Ficheros"

Sub btnLeer_Haga_clic_en()
' Lee el fichero coordenadas.txt y escribe en la hoja de calculo los valores leidos
    
    Dim fichero As String
    Dim texto As String
    Dim lineaTexto As String
    Dim posLatitud As Integer
    Dim posLongitud As Integer

    fichero = Application.GetOpenFilename()
    ' fichero = "C:\coordenadas.txt"
    
    Open fichero For Input As #1
    
    Do Until EOF(1)
        Line Input #1, lineaTexto
        texto = texto & lineaTexto
    Loop
    
    Close #1

    posLatitud = InStr(texto, "latitud")
    posLongitud = InStr(texto, "longitud")
    
    Range("A1").Value = Mid(texto, posLatitud + 9, 5)
    Range("A2").Value = Mid(texto, posLongitud + 10, 5)
    
End Sub

Sub btnEscribir_Haga_clic_en()
' Escribre en en el fichero ventas.csv los valores que contiene el rango seleccionado


    Dim fichero As String
    Dim rango As Range
    Dim valorCelda As Variant
    Dim i As Integer
    Dim j As Integer
    
    fichero = Application.DefaultFilePath & "\ventas.csv"
     
    Set rango = Selection
    
    Open fichero For Output As #1
    
    For i = 1 To rango.Rows.Count
        For j = 1 To rango.Columns.Count
      
            valorCelda = rango.Cells(i, j).Value
            If j = rango.Columns.Count Then
                Write #1, valorCelda
            Else
                Write #1, valorCelda, ' separar el valor mediante la coma
            End If
        
        Next j
    Next i
    
    Close #1
    
End Sub
