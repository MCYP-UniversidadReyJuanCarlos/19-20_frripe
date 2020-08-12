Attribute VB_Name = "Módulo1"
Option Explicit
Dim dominios(1 To 7) As String
Dim name As String
Dim h As Long
Dim i, j, k, l As Integer



Private Sub variable()

For i = 1 To 7
    dominios(i) = Worksheets("RESULTADOS").Cells(i + 7, 2).Value
Next i

End Sub




Private Sub filter()

Call variable

Worksheets("ESCENARIOS").Activate
Worksheets("ESCENARIOS").Range("AA1") = "DOMINIO"

For i = 1 To 7
    Worksheets("ESCENARIOS").Range("AA2") = dominios(i)
    name = dominios(i)
    Worksheets("ESCENARIOS").Range("A20:H320").AdvancedFilter _
    Action:=xlFilterCopy, _
    CriteriaRange:=Range("AA1:AA2"), _
    CopyToRange:=Worksheets(name).Range("A1")
    name = dominios(i)
    Worksheets(name).Range("1:1").Font.Bold = True
Next i

Worksheets("ESCENARIOS").Range("AA1:AA2").ClearContents
End Sub

Private Sub clear()


Call variable

For i = 1 To 7
    name = dominios(i)
    Worksheets(name).Cells.clear
Next i
    Worksheets("RESULTADOS").Range("D8:G15").ClearContents


End Sub


Private Sub simulacion()

Dim sim As Variant
Dim Nfilas As Integer
Dim simulacion1() As Double
Dim simulacion2() As Double
Dim perdidain, perdidare, probabilidad1, probabilidad2, probabilidad3, probabilidad4, probabilidad5 As Variant
Dim titulo As Variant
Dim inputs() As Double
Dim total1, total2 As Long
Dim micelda As Range


Call variable
Call clear
Call filter

Worksheets("RESULTADOS").Activate


sim = InputBox("Número de iteraciones:")
If sim = "" Then
    Exit Sub
End If
If Not IsNumeric(sim) Then
    Exit Sub
End If


ReDim Preserve simulacion2(1 To 2, 1 To sim)

l = 8

For k = 1 To 7
    name = dominios(k)
    Nfilas = contarfilas(name)
    perdidain = 0
    perdidare = 0
    
    If Nfilas > 1 Then
        Nfilas = Nfilas - 1
        ReDim Preserve inputs(1 To 4, 1 To Nfilas)
        ReDim Preserve simulacion1(1 To 4, 1 To Nfilas)
        total1 = 0
        For i = 1 To Nfilas
            inputs(1, i) = Log(Worksheets(name).Cells(i + 1, 4)) 'LN lim inferior
            inputs(2, i) = Log(Worksheets(name).Cells(i + 1, 5)) 'LN lim superior
            inputs(3, i) = (inputs(1, i) + inputs(2, i)) / 2  ' media
            inputs(4, i) = (inputs(2, i) - inputs(1, i)) / 3.29 ' desviación típica
        Next i
        For h = 1 To sim
            For j = 1 To Nfilas
                simulacion1(1, j) = Rnd()
                simulacion1(2, j) = Rnd()
                'simulación pérdida inherente
                If simulacion1(1, j) < Worksheets(name).Cells(j + 1, 6).Value Then
                    simulacion1(3, j) = WorksheetFunction.LogNorm_Inv(Arg1:=simulacion1(2, j), Arg2:=inputs(3, j), Arg3:=inputs(4, j))
                Else
                    simulacion1(3, j) = 0
                End If
                'simulación pérdida residual
                If Not IsEmpty(Worksheets(name).Cells(j + 1, 8).Value) Then
                    If simulacion1(1, j) < (1 - Worksheets(name).Cells(j + 1, 8).Value) * Worksheets(name).Cells(j + 1, 6).Value Then
                        simulacion1(4, j) = WorksheetFunction.LogNorm_Inv(Arg1:=simulacion1(2, j), Arg2:=inputs(3, j), Arg3:=inputs(4, j))
                    Else
                        simulacion1(4, j) = 0
                    End If
                End If
                
                
                total1 = total1 + simulacion1(3, j)
                total2 = total2 + simulacion1(4, j)
                
                ' Devolver datos
                For i = 1 To 3
                    Worksheets(name).Cells(j + 6 + Nfilas, i + 1).Offset(Nfilas * (h - 1), 0) = simulacion1(i, j)
                Next i
                    Worksheets(name).Cells(j + 6 + Nfilas, 6).Offset(Nfilas * (h - 1), 0) = simulacion1(4, j)
                            
            Next j
            simulacion2(1, h) = total1
            simulacion2(2, h) = total2
            total1 = 0
            total2 = 0
            perdidain = perdidain + simulacion2(1, h)
            perdidare = perdidare + simulacion2(2, h)
          
            ' Devolver datos
            Worksheets(name).Cells(7 + Nfilas, 1).Offset(Nfilas * (h - 1), 0) = h
            Worksheets(name).Cells(7 + Nfilas, 5).Offset(Nfilas * (h - 1), 0) = simulacion2(1, h)
            Worksheets(name).Cells(7 + Nfilas, 7).Offset(Nfilas * (h - 1), 0) = simulacion2(2, h)
        
        Next h
            perdidain = perdidain / sim
            perdidare = perdidare / sim
    Else
    'No hay definidos escenarios
    perdidain = 0
    perdidare = 0
    End If
    
    If perdidare = 0 Then
        perdidare = "N/A"
    Else
        perdidare = WorksheetFunction.Round(perdidare, 2)
    End If
    perdidain = WorksheetFunction.Round(perdidain, 2)
    

    Worksheets(name).Range("D:D").NumberFormat = "#,##0.00 €"
    Worksheets(name).Range("E:E").NumberFormat = "#,##0.00 €"
    Worksheets(name).Range("F:F").NumberFormat = "#,##0.00 €"
    Worksheets(name).Range("G:G").NumberFormat = "#,##0.00 €"
    Worksheets(name).Range("F2:F" & Nfilas + 1).NumberFormat = "0 %"
    
    
    ' Resultado de la iteración
    Set micelda = Worksheets(name).Cells(3 + Nfilas, 2)
    With micelda
        .Value = name
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    
    Set micelda = Worksheets(name).Cells(3 + Nfilas, 3)
    With micelda
        .Value = "Pérdida Inherente Media:"
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)
    End With
    
        Set micelda = Worksheets(name).Cells(4 + Nfilas, 3)
    With micelda
        .Value = "Pérdida Residual Media:"
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)
    End With
    
    Set micelda = Worksheets(name).Cells(3 + Nfilas, 4)
    With micelda
        .Value = perdidain
        .Font.Bold = True
        .NumberFormat = "#,##0.00 €"
        .WrapText = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    
    Set micelda = Worksheets(name).Cells(4 + Nfilas, 4)
    With micelda
        .Value = perdidare
        .Font.Bold = True
        .NumberFormat = "#,##0.00 €"
        .WrapText = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    
    ' Probabilidad ponderada de los escenarios
    probabilidad1 = 0
    probabilidad2 = 0
    For i = 1 To Nfilas
        Set micelda = Worksheets(name).Cells(1 + i, 6)
        probabilidad1 = probabilidad1 + inputs(3, i) * micelda
        probabilidad2 = probabilidad2 + inputs(3, i)
    Next i
    
    probabilidad3 = probabilidad1 / probabilidad2
    
    Set micelda = Worksheets(name).Cells(3 + Nfilas, 5)
    With micelda
        .Value = probabilidad3
        .Font.Bold = True
        .NumberFormat = "0 %"
        .WrapText = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    
  
    ' Tiltulo de la tabla
    
    titulo = Array("Iteración", "Aleat Prob", "Aleat Impacto", "Pérdida Inherente", "Suma", "Pérdida Residual", "Suma")
    
    For i = 0 To 6
    Set micelda = Worksheets(name).Cells(6 + Nfilas, i + 1)
    With micelda
        .Value = titulo(i)
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Next i
    

    
    Set micelda = Worksheets("RESULTADOS").Cells(l, 4)
    With micelda
        .Value = perdidain
        .Font.Bold = True
        .NumberFormat = "#,##0.00 €"
        .WrapText = True
    End With
    
    Set micelda = Worksheets("RESULTADOS").Cells(l, 5)
    With micelda
        .Value = probabilidad3
        .Font.Bold = False
        .NumberFormat = "0 %"
        .WrapText = True
    End With
    
    Set micelda = Worksheets("RESULTADOS").Cells(l, 7)
    With micelda
        .Value = perdidare
        .Font.Bold = True
        .NumberFormat = "#,##0.00 €"
        .WrapText = True
    End With
    
    l = l + 1
    
Next k

Set micelda = Worksheets("RESULTADOS").Cells(l, 4)
With micelda
    .Value = sim
    .WrapText = True
End With



    
End Sub

Function contarfilas(name As String) As Integer
Dim last_row As Integer
    contarfilas = Worksheets(name).Cells(Rows.Count, 1).End(xlUp).Row
End Function










