VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

 ' Propiedades
 Public Marca As String
 Public Modelo As String
 Public Combustible As String
 Public Motor As String
 Public Puertas As Integer

' Métodos
Sub ArrancarMotor()
    MsgBox "El motor esta encendido"
End Sub

Sub PararMotor()
    MsgBox "El motor esta parado"
End Sub

Sub Acelerar()
    MsgBox "La velocidad se ha incrementado."
End Sub

Sub Frenar()
  MsgBox "La velocidad se ha reducido"
End Sub
