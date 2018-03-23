VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Libro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vTitulo As String
Dim vAutor As String
Dim vPaginas As Long
Dim vPrecio As Double
Dim vFechaPublicacion As Date
Dim vDisponible As Boolean


Public Property Let Titulo(t As String)
    vTitulo = t
End Property
Public Property Get Titulo() As String
    Titulo = vTitulo
End Property
Public Property Let Disponible(b As Boolean)
    vDisponible = b
End Property
Public Property Get Disponible() As Boolean
    Disponible = vDisponible
End Property


Sub MostrarInfo()

    Dim mensaje As String
    
    mensaje = "Título:" & vbTab & vTitulo & vbCr
    mensaje = mensaje & "Autor:" & vbTab & vAutor & vbCr
    mensaje = mensaje & "Paginas:" & vbTab & vPaginas & vbCr
    mensaje = mensaje & "Precio:" & vbTab & vPrecio & " €" & vbCr
    mensaje = mensaje & "Fecha de publicación:" & vbTab & vFechaPublicacion & vbCr
    
    If vDisponible Then
        mensaje = mensaje & vbCr & "DISPONIBLE"
        MsgBox mensaje, vbOKOnly + vbInformation, "Datos de libro"
    End If
End Sub

