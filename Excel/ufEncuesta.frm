VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEncuesta 
   Caption         =   "Evaluación del perfil personal del profesor/a"
   ClientHeight    =   15585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   OleObjectBlob   =   "ufEncuesta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ufEncuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()

Dim elegido As Integer

For i = 0 To fr1.Controls.Count
    If fr1.Controls(i).Value = True Then
        elegido = i
    End If
Next
MsgBox elegido

End Sub

Private Sub Label44_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
