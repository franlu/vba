Attribute VB_Name = "00"
Sub MiPrimeraMacro()
'
' MiPrimeraMacro Macro
'   Escribe un nombre en color rojo
'   Fuente Verdana
'   Negrita
'   Cursiva
'   Tamaño 14
'
    Selection.TypeText Text:="Fco Javier"
    Selection.MoveLeft Unit:=wdCharacter, Count:=10, Extend:=wdExtend
    Selection.Font.Bold = wdToggle
    Selection.Font.Italic = wdToggle
    Selection.Font.Color = wdColorRed
    Selection.Font.Size = 14
End Sub

Sub MiSegundaMacro()
'
' MiSegundaMacro Macro
'   Insertar un WordArt
'
    ActiveDocument.Shapes.AddTextEffect(msoTextEffect13, "Macros", "Impact", _
        44#, msoFalse, msoFalse, 230.6, 286.5).Select
    Selection.InlineShapes(1).TextEffect.PresetShape = _
        msoTextEffectShapeCurveDown
    Selection.InlineShapes(1).Line.DashStyle = msoLineLongDash
End Sub

Sub MiTerceraMacro()
'
' MiTerceraMacro Macro
'   Escribe un nombre completo
'   Fuente: Arial
'   Color: Azul Oscuro
'   Negrita
'   Subrayado
'
    Selection.TypeText Text:="Eva Cuación Inmediata"
    Selection.MoveLeft Unit:=wdCharacter, Count:=21, Extend:=wdExtend
    Selection.Font.Name = "Arial"
    Selection.Font.Color = -738148353
    Selection.Font.Bold = wdToggle
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineSingle
End Sub

Sub MiCuartaMacro()
'
' MiCuartaMacro Macro
'   Inserta una tabla.
'   Inserta la fecha en una celda cualquiera.
'   Centrar el contenido de la celda.
'   Aplicar dobles bordes sobre esa celda.
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Tabla con cuadrícula" Then
            .Style = "Tabla con cuadrícula"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    Selection.InsertDateTime DateTimeFormat:="dd/MM/yyyy", InsertAsField:= _
        False, DateLanguage:=wdSpanishModernSort, CalendarType:=wdCalendarWestern _
        , InsertAsFullWidth:=False
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    With Selection.Cells
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub

Sub WebAyuntamiento()
'
' Hipervinculo Macro
'   Inserta un hipervinculo a la web del ayuntamiento
'
'
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        "http://www.granada.org", SubAddress:="", ScreenTip:="", TextToDisplay:= _
        "http://www.granada.org"
End Sub
