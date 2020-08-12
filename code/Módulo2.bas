Attribute VB_Name = "Módulo2"
Option Explicit
Dim ematrix(1 To 5, 1 To 2) As Double
Dim rectas(1 To 4, 1 To 2) As Long
Dim i, j As Integer
Dim restimado(1 To 7, 1 To 2), raceptable(1 To 7, 1 To 2) As Double
Dim rectasriesgo(1 To 7, 1 To 2) As Double
Dim interseccion(1 To 4, 1 To 7, 1 To 2) As Double
Dim crrf(1 To 7) As Double
Dim SLT(1 To 7) As Integer
Dim dominios(1 To 7) As String
Dim micelda As Range
Dim ebitda As Long

Private Sub datos_entrada()

ebitda = Worksheets("ESCENARIOS").Range("C17").Value

'definir puntos de curva de riesgo tolerable

ematrix(1, 1) = 0.03 * ebitda
ematrix(2, 1) = 0.05 * ebitda
ematrix(3, 1) = 0.11 * ebitda
ematrix(4, 1) = 0.2 * ebitda
ematrix(5, 1) = ebitda
ematrix(1, 2) = 0.8
ematrix(2, 2) = 0.4
ematrix(3, 2) = 0.15
ematrix(4, 2) = 0.07
ematrix(5, 2) = 0

'definir rectas de curva de riesgo tolerable(pendiente y término independiente)

For i = 1 To 4
    rectas(i, 1) = (ematrix(i + 1, 1) - ematrix(i, 1)) / (ematrix(i + 1, 2) - ematrix(i, 2))
    rectas(i, 2) = ematrix(i, 1) - rectas(i, 1) * ematrix(i, 2)
Next i

'Riesgo estimado obtenido en la simulación
For i = 1 To 7
    restimado(i, 1) = Worksheets("RESULTADOS").Cells(7 + i, 4).Value
    restimado(i, 2) = Worksheets("RESULTADOS").Cells(7 + i, 5).Value
Next i




End Sub


Private Sub calculo_slt()


Set micelda = Worksheets("ESCENARIOS").Range("C17")
If IsEmpty(micelda) Then
    MsgBox "No se ha introducido el valor del EBITDA"
    Exit Sub
End If
If Not IsNumeric(micelda) Then
    MsgBox "No se ha introducido el valor numérico en el  EBITDA"
    Exit Sub
End If

Call datos_entrada

'rectas riesgo (pendiente y término independiente)

For i = 1 To 7
    rectasriesgo(i, 1) = 0.3 * ebitda / 0.8
    rectasriesgo(i, 2) = restimado(i, 1) - rectasriesgo(i, 1) * restimado(i, 2)
Next i

' Puntos de intersección entre rectas riesgo y rectas de riesgo tolerable

For i = 1 To 4
    For j = 1 To 7
        interseccion(i, j, 1) = (rectasriesgo(j, 2) - rectas(i, 2)) / (rectas(i, 1) - rectasriesgo(j, 1))
        interseccion(i, j, 2) = rectasriesgo(j, 1) * interseccion(i, j, 1) + rectasriesgo(j, 2)
    Next j
Next i


'Calculo de riesgo aceptable, puntos de intersección válidos
For i = 1 To 4
    For j = 1 To 7
            If interseccion(i, j, 1) > ematrix(i + 1, 2) And interseccion(i, j, 1) < ematrix(i, 2) And interseccion(i, j, 2) > ematrix(i, 1) And interseccion(i, j, 2) < ematrix(i + 1, 1) Then
                raceptable(j, 1) = interseccion(i, j, 2)
                raceptable(j, 2) = interseccion(i, j, 1)
            End If
    Next j
Next i


'Igualar unidades de impacto y probabilidad. Cálculo de crrf

For j = 1 To 7
    restimado(j, 1) = restimado(j, 1) / (0.3 * ebitda / 5)
    restimado(j, 2) = restimado(j, 2) / 0.16
    raceptable(j, 1) = raceptable(j, 1) / (0.3 * ebitda / 5)
    raceptable(j, 2) = raceptable(j, 2) / 0.16
    crrf(j) = (restimado(j, 1) * restimado(j, 2)) / (raceptable(j, 1) * raceptable(j, 2))
Next j


'Cálculo de SLT
For j = 1 To 7
    Select Case crrf(j)
        Case Is < 1.25
            SLT(j) = 0
        Case Is < 2.25
            SLT(j) = 1
        Case Is < 3.25
            SLT(j) = 2
        Case Is < 4.25
            SLT(j) = 3
        Case Else
            SLT(j) = 4
    End Select
    Set micelda = Worksheets("RESULTADOS").Cells(7 + j, 6)
    With micelda
        .Value = SLT(j)
        .NumberFormat = "0"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
Next j

End Sub


Private Sub grafico()

Call datos_entrada

Dim grafico As ChartObject
Dim wks As Worksheet
Dim cht As Chart



Set wks = Worksheets(hoja)

wks.Range("AA1").Value = ematrix(1, 2)
wks.Range("AB1").Value = ematrix(2, 2)
wks.Range("AC1").Value = ematrix(3, 2)
wks.Range("AD1").Value = ematrix(4, 2)
wks.Range("AE1").Value = ematrix(5, 2)
wks.Range("AA2").Value = rectas(1, 1) * wks.Range("AA1").Value + rectas(1, 2)
wks.Range("AB2").Value = rectas(1, 1) * wks.Range("AB1").Value + rectas(1, 2)
wks.Range("AC2").Value = rectas(2, 1) * wks.Range("AC1").Value + rectas(2, 2)
wks.Range("AD2").Value = rectas(3, 1) * wks.Range("AD1").Value + rectas(3, 2)
wks.Range("AE2").Value = rectas(4, 1) * wks.Range("AE1").Value + rectas(4, 2)

Set grafico = wks.ChartObjects.Add(Left:=20, Width:=400, Top:=distance, Height:=300)
grafico.name = "grafico_riesgo"

grafico.Chart.ChartType = xlXYScatterLines
grafico.Chart.SetSourceData Source:=wks.Range("AA1:AE2")
grafico.Chart.SeriesCollection(1).Smooth = True
grafico.Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleNone


With grafico.Chart.PlotArea.Format.Fill
    .ForeColor.RGB = RGB(255, 0, 0)
    .TwoColorGradient msoGradientDiagonalDown, 1
    .GradientStops(2).Color = RGB(0, 255, 0)
    .GradientStops(2).Position = 0.99
    .GradientStops.Insert RGB(255, 255, 0), 0.5
End With

Set cht = wks.ChartObjects("grafico_riesgo").Chart

For i = 1 To 7
    With cht.SeriesCollection.NewSeries
        .name = Worksheets("RESULTADOS").Cells(7 + i, 2).Value
        .Values = restimado(i, 1)
        .XValues = restimado(i, 2)
        .HasDataLabels = True
        .MarkerSize = 8
        .MarkerStyle = xlMarkerStyleCircle
        .Smooth = True
        With .DataLabels
            .ShowLegendKey = False
            .ShowSeriesName = True
            .ShowValue = False
        End With
        
    End With
Next i

With cht.Axes(xlCategory)
    .MaximumScale = 0.8
    .TickLabels.NumberFormat = "0 %"
    .HasMajorGridlines = True
    .MajorUnit = 0.16
End With

With cht.Axes(xlValue)
    .MaximumScale = 0.3 * ebitda
    .TickLabels.NumberFormat = "#.##0 €"
    .HasMajorGridlines = True
    .MajorUnit = 0.06 * ebitda
End With

cht.HasLegend = False

cht.HasTitle = True
cht.ChartTitle.Text = "Representación gráfica de los escenarios"
With cht.ChartTitle.Font
    .Size = 14
    .Bold = True
    .Color = RGB(0, 0, 0)
End With

End Sub


