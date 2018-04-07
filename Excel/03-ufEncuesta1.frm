VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncuesta1 
   Caption         =   "Evaluación del perfil personal del profesor/a"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8715
   OleObjectBlob   =   "03-ufEncuesta1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ufEncuesta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClean_Click()

Dim oFrame As Control
Dim oOptionButton As Control
Dim oTextBox As Control

For Each oFrame In ufEncuesta1.mpEncuesta.Pages(0).Controls
    If (oFrame.Name Like "fr*") Then
        For Each oOptionButton In oFrame.Controls
            oOptionButton.Value = False
        Next
    End If
Next

For Each oTextBox In ufEncuesta1.mpEncuesta.Pages(1).Controls
    If (oTextBox.Name Like "txt*") Then
        oTextBox.Text = ""
    End If
Next

MsgBox "La encuesta ha quedado vacía.", vbInformation, "Encuesta"

End Sub

Private Sub btnSave_Click()

Dim i As Byte
Dim completo As Boolean
Dim fecha As String

Sheets(2).Activate
'Hoja2.Activate

i = 1
For Each oFrame In ufEncuesta1.mpEncuesta.Pages(0).Controls
    If (oFrame.Name Like "fr*") Then
        For Each oOptionButton In oFrame.Controls
            Cells(i, 1).Value = Mid(oOptionButton.Name, Len(oOptionButton.Name), 1)
        Next
        i = i + 1
    End If
Next

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
    
    fecha = Now
    fecha = Replace(fecha, "/", "-")
    fecha = Replace(fecha, ":", "-")
    fecha = Replace(fecha, " ", "-")
    
    ActiveWorkbook.SaveAs Filename:="Resultado-" & fecha & ".xlsm"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    MsgBox "Gracias por realizar la encuesta", vbInformation, "Fin de la encuesta"
    
    ufEncuesta1.Hide
    
End If

End Sub
