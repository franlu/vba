Attribute VB_Name = "ImportarHoja"

Public Sub importarHoja()
' Importar todas las hojas de los libros existentes en un directorio a un único libro.
    
    Dim directorio As String
    Dim fichero As String
    Dim hoja As Worksheet
    Dim total As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    directorio = "c:\libros\"
    fichero = Dir(directorio & "*.xl??")
    
    Do While fichero <> ""
        MsgBox "Importando hojas desde: " & fichero
        Workbooks.Open (directorio & fichero)
        For Each hoja In Workbooks(fichero).Worksheets
            total = Workbooks("02-Importar-hojas.xlsm").Worksheets.Count
            Workbooks(fichero).Worksheets(hoja.Name).Copy _
            after:=Workbooks("02-Importar-hojas.xlsm").Worksheets(total)
        Next hoja
        Workbooks(fichero).Close
        fichero = Dir() 'Si no hay mas ficheros Dir() devuelve cadena vacía
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
