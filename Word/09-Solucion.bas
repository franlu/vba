Attribute VB_Name = "Repaso"

Public Sub Circulo()
'
' Pedir al usuario el radio de un c�rculo y muestrar un mensaje
' con su per�metro y su �rea.
'
    Const PI As Double = 3.1416
    
    Dim radio As Double
    Dim perimetro As Double
    Dim area As Double
    Dim mensaje As String
    
    radio = InputBox("Introduzca el Radio del C�rculo", "C�rculo")
    
    perimetro = 2 * PI * radio
    area = PI * (radio ^ 2)
    mensaje = "Para un c�rculo con Radio = " & radio & " Cm. Obtenemos: " & vbCr _
                & vbCr + vbTab & "�rea = " & area _
                & vbCr + vbTab & "Per�metro = " & perimetro
                
    MsgBox mensaje
    
End Sub

Sub parteDecimal()
'
' Mostrar en un mensaje la parte decimal de un n�mero real introducido por el usuario.
'
    Dim valorDecimal As Double
    Dim parteDecimal As Double
    Dim mensaje As String
    
    valorDecimal = InputBox("Introduce un valor decimal...")
    
    parteDecimal = valorDecimal - Int(valorDecimal)
    mensaje = "La parte decimal de " & valorDecimal & "es: " & parteDecimal
    
    MsgBox mensaje
    
End Sub

Sub palabras()
'
' Indicar cu�l de dos palabras es la m�s larga y por cu�ntas letras lo es.
'
	Dim palabra1 As String
	Dim palabra2 As String
	Dim diferencia As String
	
	palabra1 = InputBox("Introduce la primera palabra...")
	palabra2 = InputBox("Introduce la segunda palabra...")
	
	diferencia = Abs(palabra1 - palabra2)
	
	If Len(palabra1) > Len(palabra2) Then
		MsgBox "La primera palabra es mas larga por " & diferencia & " letras."
	ElseIf Len(palabra1) < Len(palabra2)
		MsgBox "La segunda palabra es mas larga por " & diferencia & " letras."
	Else
		MsgBox "Las dos palabras tienen la misma longitud."
	End If
	
End Sub

' 4. Escriba un programa que pida peso y altura al usuario. Calcule su IMC.
'    Muestre un mensaje seg�n la clasificaci�n de la OMS.
'
' 5. Escriba un programa que muestre la tabla de multiplicar del 1 al 10 del n�mero ingresado
'    por el usuario.
'
' 6. Escriba un programa que genere todas las potencias de 2, desde la 0-�sima hasta la
'    ingresada por el usuario.
'
' 7. Escriba un programa que reciba como entrada las longitudes de los dos catetos a y b
'    de un tri�ngulo rect�ngulo, y que entregue como salida el largo de la hipotenusa c del
'    triangulo, dado por el teorema de pitagoras c2 = a2 + b2.
