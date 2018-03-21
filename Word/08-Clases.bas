Attribute VB_Name = "Clases"




Sub UsarClase1()

    Dim milibro As New Libro
    
    milibro.Titulo = "Quijote"
    milibro.Disponible = True
    
    milibro.MostrarInfo
    
        
End Sub

Sub UsarClase2()

    Dim micoche As New Coche
    
    With micoche
        .Marca = "Chrysler"
        .Modelo = "Voyager"
        .Combustible = "Diesel"
        .Motor = "2.8 Litros"
        .Puertas = 5
    End With
       
    micoche.ArrancarMotor
    
End Sub
