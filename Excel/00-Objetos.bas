Attribute VB_Name = "Objetos"
' Objetos para trabajar con Excel
'   Application
'   Workbooks
'   ActiveWorkbook
'   WorkSheet
'   ActiveCell


Public Sub Nuevo_libro()
' Crear un nuevo libro
    
    Workbooks.Add

'Workbooks.Add Template:= "\\servidor\plantilla\excel\Balance.xlt"
'Workbooks.Add Template:= "C:\empresa\trabajadores.xlsx"
'Workbooks.Add Template:=xlWBATChart

'ActiveWorkbook.Save
'ActiveWorkbook.SaveAs FileName:="Salarios.xlsx”

'Workbooks.Open Filename:= "C:\empresa\inventario.xlsx”

End Sub

Sub Nuevo_libro_12_hojas()
' Crear un libro con 12 hojas

    Dim nhojas As Integer
    
    'Numero de hojas que contiene un nuevo libro
    nhojas = Application.SheetsInNewWorkbook
    
    'Se actualiza a 12 en numero de hojas del nuevo libro
    Application.SheetsInNewWorkbook = 12
    
    'Crear el nuevo libro
    Workbooks.Add
    
    'Volvemos a dejar 3 hojas para el nuevo libro
    Application.SheetsInNewWorkbook = nhojas
End Sub

Sub Nuevo_libro_v2(nh As Integer)
'Crea un nuevo libro con un número determinado de hojas

	Dim nh As Integer
    
    nhojas = Application.SheetsInNewWorkbook
       
    Application.SheetsInNewWorkbook = nh
    Workbooks.Add
       
    Application.SheetsInNewWorkbook = nhojas

End Sub

Sub Guardar_todos_libros()
' Guardar todos los libros que contiene la colección Workbooks
    
    Dim miLibro As Workbook
    For Each miLibro In Workbooks
        miLibro.Save
    Next miLibro
    
End Sub

Sub Libro_activo()
' Comprobar si hay algún libro abierto

    If ActiveWorkbook Is Nothing Then
        MsgBox "Por favor crea un nuevo libro antes de usar esta macro." _
        & vbCr & vbCr & "La macro ha finalizado.", _
        vbOKOnly + vbExclamation, "No hay libro abierto."
    End If
    
End Sub

Sub Usar_libro_activo()
' Mostrar el nombre del libro activo
    
    Dim miLibro As Workbooks
    Set miLibro = ActiveWorkbook
    
    MsgBox miLibro.Name

End Sub


Sub Aniadir_hoja()
' Añadir una hoja al primer libro, antes de la primera hoja

    Dim hoja As Worksheet
    Set hoja = Workbooks(1).Sheets.Add(before:=Sheets(1))
    hoja.Name = "Balance"
    
End Sub


Sub Eliminar_hoja()
' Eliminar la hoja de un libro
    
	Dim miLibro As Workbook
	Dim hoja As Worksheet
	
	Set miLibro = ActiveWorkbook
    Set hoja = Workbooks(1).Sheets.Add(before:=Sheets(1))
    
	hoja.Name = "Ingresos"
    miLibro.Sheets("Ingresos").Delete
	
End Sub

Sub Anio()
' Generar un libro con 12 hojas, el nombre de cada hoja es el nombre de los meses del año
	
	Dim nhojas As Integer
    
    nhojas = Application.SheetsInNewWorkbook
       
    Application.SheetsInNewWorkbook = 12
    
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:="Anio.xlsx"
    
    Application.SheetsInNewWorkbook = nhojas
    
    ActiveWorkbook.Sheets(1).Name = "Enero"
    ActiveWorkbook.Sheets(2).Name = "Febrero"
    ActiveWorkbook.Sheets(3).Name = "Marzo"
	ActiveWorkbook.Sheets(4).Name = "Abril"
    ActiveWorkbook.Sheets(5).Name = "Mayo"
    ActiveWorkbook.Sheets(6).Name = "Junio"
	ActiveWorkbook.Sheets(7).Name = "Julio"
    ActiveWorkbook.Sheets(8).Name = "Agosto"
    ActiveWorkbook.Sheets(9).Name = "Septiembre"
	ActiveWorkbook.Sheets(10).Name = "Octubre"
    ActiveWorkbook.Sheets(11).Name = "Noviembre"
    ActiveWorkbook.Sheets(12).Name = "Diciembre"

End Sub

Sub AnioV2()
' Generar un libro con 12 hojas, el nombre de cada hoja es el nombre de los meses del año
    
    Dim nhojas As Integer
    Dim meses(12) As String
    
    nhojas = Application.SheetsInNewWorkbook
       
    Application.SheetsInNewWorkbook = 12
    
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:="AnioV2.xlsx"
    
    Application.SheetsInNewWorkbook = nhojas
    
    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "Septiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        ActiveWorkbook.Sheets(i).Name = meses(i - 1)
    Next

End Sub

Sub Celda_Activa()
' Mostrar un mensaje con la dirección de la celda activa en el documento activo
    MsgBox ActiveCell.Address
End Sub

Sub Asignar_Valor()
' Establecer un valor para la celda activa

    ActiveCell.Value = 25
    MsgBox ActiveCell.Value
    
End Sub
