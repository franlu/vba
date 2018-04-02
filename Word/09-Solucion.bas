Attribute VB_Name = "Repaso"

Public Sub Circulo()
'
' Pedir al usuario el radio de un círculo y muestrar un mensaje
' con su perímetro y su área.
'
    Const PI As Double = 3.1416
    
    Dim radio As Double
    Dim perimetro As Double
    Dim area As Double
    Dim mensaje As String
    
    radio = InputBox("Introduzca el Radio del Círculo", "Círculo")
    
    perimetro = 2 * PI * radio
    area = PI * (radio ^ 2)
    mensaje = "Para un círculo con Radio = " & radio & " Cm. Obtenemos: " & vbCr _
                & vbCr + vbTab & "Área = " & area _
                & vbCr + vbTab & "Perímetro = " & perimetro
                
    MsgBox mensaje
    
End Sub

Sub parteDecimal()
'
' Mostrar en un mensaje la parte decimal de un número real introducido por el usuario.
'
    Dim valorDecimal As Double
    Dim parteDecimal As Double
    Dim mensaje As String
    
    valorDecimal = InputBox("Introduce un valor decimal...")
    
    parteDecimal = valorDecimal - Int(valorDecimal)
    mensaje = "La parte decimal de " & valorDecimal & "es: " & parteDecimal
    
    MsgBox mensaje
    
End Sub

'
' 3. Escriba un programa que pida al usuario dos palabras, y que indique cuál de ellas es
'    la más larga y por cuántas letras lo es.
'
' 4. Escriba un programa que pida peso y altura al usuario. Calcule su IMC.
'    Muestre un mensaje según la clasificación de la OMS.
'
' 5. Escriba un programa que muestre la tabla de multiplicar del 1 al 10 del número ingresado
'    por el usuario.
'
' 6. Escriba un programa que genere todas las potencias de 2, desde la 0-ésima hasta la
'    ingresada por el usuario.
'
' 7. Escriba un programa que reciba como entrada las longitudes de los dos catetos a y b
'    de un triángulo rectángulo, y que entregue como salida el largo de la hipotenusa c del
'    triangulo, dado por el teorema de pitagoras c2 = a2 + b2.
