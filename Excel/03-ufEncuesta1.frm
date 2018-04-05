VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncuesta1 
   Caption         =   "Evaluaci�n del perfil personal del profesor/a"
   ClientHeight    =   11565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8715
   OleObjectBlob   =   "ufEncuesta1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ufEncuesta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

MsgBox "La encuesta ha quedado vac�a.", vbInformation, "Encuesta"

End Sub

Private Sub btnSave_Click()


Sheets(2).Activate
'Hoja2.Activate

For Each objeto In fr1.Controls
    If objeto.Value = True Then
        Cells(1, 1).Value = Mid(objeto.Name, Len(objeto.Name), 1)
    End If
Next

Cells(2, 1) = txtMegusta1.Text
Cells(2, 2) = txtMegusta2.Text
Cells(2, 3) = txtMegusta3.Text

Cells(2, 4) = txtNomegusta1.Text
Cells(2, 5) = txtNomegusta2.Text
Cells(2, 6) = txtNomegusta3.Text

Cells(2, 7) = txtCambio1.Text
Cells(2, 8) = txtCambio2.Text
Cells(2, 9) = txtCambio3.Text


MsgBox "Gracias por realizar la encuesta", vbInformation, "Fin de la encuesta"

ufEncuesta1.Hide

End Sub
