Private Sub btnClean_Click()

Dim objeto As Control

For Each objeto In fr1.Controls
    objeto.Value = False
Next

For Each objeto In fr2.Controls
    objeto.Value = False
Next

For Each objeto In ufEncuesta1.mpEncuesta.Pages(1).Controls
    If (objeto.Name Like "txt*") Then
        objeto.Text = ""
    End If
Next

MsgBox "La encuesta ha quedado vacía.", vbInformation, "Encuesta"

End Sub

Private Sub btnSave_Click()



Sheets(2).Activate
'Hoja2.Activate

For Each objeto In fr1.Controls
    If objeto.Value = True Then
        Cells(1, 1).Value = Mid(objeto.Name, Len(objeto.Name), 1)
    End If
Next

Dim completo As Boolean
completo = True

For Each objeto In ufEncuesta1.mpEncuesta.Pages(1).Controls
    If (objeto.Name Like "txt*") Then
        If (objeto.Text = "") Then
            MsgBox "Debes rellenar todos los cuadros de texto", vbExclamation, "Atención"
            completo = False
            Exit For
        End If
    End If
Next

If completo Then

    Cells(2, 1) = txtMegusta1.Text
    Cells(2, 2) = txtMegusta2.Text
    Cells(2, 3) = txtMegusta3.Text
    
    Cells(2, 4) = txtNomegusta1.Text
    Cells(2, 5) = txtNomegusta2.Text
    Cells(2, 6) = txtNomegusta3.Text
    
    Cells(2, 7) = txtCambio1.Text
    Cells(2, 8) = txtCambio2.Text
    Cells(2, 9) = txtCambio3.Text
    
    Dim fecha As String
    fecha = Now
    ' 01/05/2018 10:48:23
    
    fecha = Replace(fecha, "/", "-")
    ' 01-05-2018 10:48:23
    fecha = Replace(fecha, ":", "-")
    ' 01-05-2018 10-48-23
    fecha = Replace(fecha, " ", "-")
    ' 01-05-2018-10-48-23
    
    ActiveWorkbook.SaveAs Filename:="EncuestaAlumno-" & fecha & ".xlsm"
    
    MsgBox "Gracias por realizar la encuesta", vbInformation, "Fin de la encuesta"
    
	
	Sheets(2).Activate
    
    ' Limpiar la hoja de Resultados
    Cells(2, 1) = ""
    Cells(2, 2) = ""
    Cells(2, 3) = ""
    
    Cells(2, 4) = ""
    Cells(2, 5) = ""
    Cells(2, 6) = ""
    
    Cells(2, 7) = ""
    Cells(2, 8) = ""
    Cells(2, 9) = ""
    
    'Limpiar el formulario antes de finalizar
    Call btnClean_Click
	
    ufEncuesta1.Hide
    
End If

End Sub
