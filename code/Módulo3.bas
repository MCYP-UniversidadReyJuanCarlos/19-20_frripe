Attribute VB_Name = "Módulo3"
Option Explicit

Public hoja As String
Public distance As String


Dim name, name2 As String
Dim resultado(1 To 7, 1 To 6) As Variant
Dim micelda As Range
Dim i, j, k, contarfilas, escenarios, iteraciones As Integer

Private Sub informe1()

    hoja = "INFORME " & Year(Now) & " " & Month(Now) & " " & Day(Now) & " " & Hour(Now) & " " & Minute(Now)

    Worksheets.Add After:=Worksheets("RESULTADOS")
    
    ActiveSheet.name = hoja
    
    For i = 1 To 7
        For j = 1 To 6
            resultado(i, j) = Worksheets("RESULTADOS").Cells(i + 7, j + 1).Value
        Next j
    Next i
    
    iteraciones = Worksheets("RESULTADOS").Cells(15, 4).Value
    
    
    Set micelda = Worksheets(hoja).Range("A2:G2")
    name = Worksheets("ESCENARIOS").Cells(16, 2) & " " & Worksheets("ESCENARIOS").Cells(16, 3)
  
    With micelda
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 14
        .Font.Underline = True
        .Font.name = "Arial"
    End With
   
    Set micelda = ActiveCell.Offset(3, 0)
    micelda.Select
    
    name = "1. Nivel de seguridad objetivo"
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    Set micelda = ActiveCell.Offset(2, 1)
    micelda.Select
    
    name = "NS-O = ("
    
    For i = 1 To 6
        name = name & CStr(resultado(i, 5)) & ", "
    Next i
    name = name & CStr(resultado(7, 5)) + ")"
            
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    Set micelda = ActiveCell.Offset(2, 0)
    micelda.Select
      
    name = "NS-O = ("
    For i = 1 To 6
        name = name + resultado(i, 1) + ", "
    Next i
    name = name + resultado(7, 1) + ")"
      
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = False
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    Set micelda = ActiveCell.Offset(1, -1)
    micelda.Select
      
    For i = 1 To 7
        Set micelda = ActiveCell.Offset(1, 0)
        micelda.Select
        name = resultado(i, 1) & ": "
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Value = name
            .Font.Bold = False
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
        End With
    Next i
       
    Set micelda = ActiveCell.Offset(-7, 1)
    micelda.Select
       
    For i = 1 To 7
        Set micelda = ActiveCell.Offset(1, 0)
        micelda.Select
        name = resultado(i, 2)
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Value = name
            .Font.Bold = False
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
        End With
    Next i
    
    Set micelda = ActiveCell.Offset(3, -1)
    micelda.Select
    

End Sub

Private Sub informe2()


    name = "2. Escenarios evaluados"
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    For i = 1 To 7
        Set micelda = ActiveCell.Offset(1, 0)
        micelda.Select
        name = "2." & i & " " & resultado(i, 2) & " (" & resultado(i, 1) & ")"
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Value = name
            .Font.Bold = True
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
        End With
        
        Set micelda = ActiveCell.Offset(2, 0)
        micelda.Select
        
        Worksheets("ESCENARIOS").Activate
        Worksheets("ESCENARIOS").Range("AA1") = "DOMINIO"
        Worksheets("ESCENARIOS").Range("AA2") = resultado(i, 1)
        name2 = resultado(i, 1)
        Worksheets("ESCENARIOS").Range("A20:H320").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Range("AA1:AA2"), _
        CopyToRange:=micelda
        Worksheets("ESCENARIOS").Range("AA1:AA2").ClearContents

        Worksheets(hoja).Activate
        
        contarfilas = Worksheets(hoja).Cells(Rows.Count, 1).End(xlUp).Row
        escenarios = contarfilas - Selection.Row
        
        Set micelda = Range(ActiveCell, ActiveCell.Offset(2 + escenarios, 0))
        micelda.Select
        Selection.Copy
        
        Set micelda = ActiveCell.Offset(2 + escenarios, 0)
        micelda.Select
        ActiveSheet.Paste
        
        Set micelda = Range(ActiveCell.Offset(-2, 6), ActiveCell.Offset(-2 - escenarios, 6))
        micelda.Select
        Selection.Cut
        
        Set micelda = ActiveCell.Offset(2 + escenarios, -5)
        micelda.Select
        ActiveSheet.Paste
        
        Set micelda = Range(ActiveCell.Offset(-2, 6), ActiveCell.Offset(-2 - escenarios, 6))
        micelda.Select
        Selection.Cut
        
        Set micelda = ActiveCell.Offset(2 + escenarios, -4)
        micelda.Select
        ActiveSheet.Paste
        
        Set micelda = Range(ActiveCell.Offset(-2, 0), ActiveCell.Offset(-2 - escenarios, 2))
        micelda.Select
        Selection.Cut
        
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        ActiveSheet.Paste
        
        Set micelda = Range(ActiveCell.Offset(0, -2), ActiveCell.Offset(escenarios, -2))
        micelda.Select
        Selection.Cut
        
        Set micelda = ActiveCell.Offset(0, -1)
        micelda.Select
        ActiveSheet.Paste
        
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 2))
        micelda.Select
        Selection.Merge
        
        For j = 1 To escenarios
            Set micelda = ActiveCell.Offset(1, 0)
            micelda.Select
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 2))

            micelda.Select
            Selection.Merge

        Next j
        
        Set micelda = ActiveCell.Offset(2, 0)
        micelda.Select
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 1))
        micelda.Select
        Selection.Merge
        Selection.Value = "COSTE MEDIDAS COMPENSATORIAS"
        
        For j = 1 To escenarios
            Set micelda = ActiveCell.Offset(1, 0)
            micelda.Select
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 1))
            micelda.Select
            Selection.Merge
        Next j
        
        Set micelda = ActiveCell.Offset(-escenarios, 1)
        micelda.Select
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 1))
        micelda.Select
        Selection.Merge
        Selection.Value = "EFECTIVIDAD DE LAS MEDIDAS"
        
        For j = 1 To escenarios
            Set micelda = ActiveCell.Offset(1, 0)
            micelda.Select
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 1))
            micelda.Select
            Selection.Merge
        Next j
        
        Set micelda = ActiveCell.Offset(-2 - 2 * escenarios, -3)
        micelda.Select
        Set micelda = ActiveCell.Offset(0, 4)
        micelda.Select
        Selection.Value = "LÍM INF"
        
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        Selection.Value = "LÍM SUP"
        
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        Selection.Value = "PROB"
        
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, -6))
        micelda.Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Interior.Color = xlNone
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThick
            .RowHeight = 40
        End With
        
        For j = 1 To escenarios
            Set micelda = ActiveCell.Offset(1, 0)
            micelda.Select
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 6))
            micelda.Select
            With Selection
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .RowHeight = 70
            End With
        Next j
        
        Set micelda = ActiveCell.Offset(2, 0)
        micelda.Select
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
        micelda.Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = True
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThick
            .RowHeight = 40
        End With
        For j = 1 To escenarios
            Set micelda = ActiveCell.Offset(1, 0)
            micelda.Select
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
            micelda.Select
            With Selection
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            Selection.EntireRow.AutoFit
        Next j
        Set micelda = ActiveCell.Offset(1, 0)
        micelda.Select


Next i
        
Set micelda = ActiveCell.Offset(2, 0)
micelda.Select

End Sub


Private Sub informe3()


    name = "3. Valoración cuantitativa de los escenarios en cada dominio (Pérdida Inherente)"
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    Set micelda = ActiveCell.Offset(2, 0)
    micelda.Select
    Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
    micelda.Select
    Selection.Merge
    With Selection
        .Value = "DOMINIO"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    
    Set micelda = ActiveCell.Offset(0, 1)
    micelda.Select
    With Selection
        .Value = "PÉRDIDA MEDIA INHERENTE"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    Set micelda = ActiveCell.Offset(0, 1)
    micelda.Select
    With Selection
        .Value = "PROB"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    Set micelda = ActiveCell.Offset(1, -6)
    micelda.Select

    For j = 1 To 7
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
        micelda.Select
        Selection.Merge
        With Selection
            .Value = resultado(j, 2) & " (" & resultado(j, 1) & ")"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        With Selection
            .Value = resultado(j, 3)
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = False
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .NumberFormat = "#,##0 €"
        End With
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        With Selection
            .Value = resultado(j, 4)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = False
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .NumberFormat = "0 %"
        End With
        Set micelda = ActiveCell.Offset(1, -6)
        micelda.Select
     Next j
     
     
     Set micelda = ActiveCell.Offset(1, 0)
     micelda.Select
     With Selection
        .Value = "Nota: Valores obtenidos a partir de una simulación de " & iteraciones & " iteraciones"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Size = 8
        .Font.name = "Arial"
    End With
    Set micelda = ActiveCell.Offset(3, 0)
    micelda.Select
    
    distance = Selection.Top
    
    Application.Run "Módulo2.grafico"
    
    Set micelda = ActiveCell.Offset(23, 0)
    micelda.Select

End Sub

Private Sub informe4()


    name = "4. Valoración cuantitativa de los escenarios en cada dominio (Pérdida Residual)"
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    Set micelda = ActiveCell.Offset(2, 0)
    micelda.Select
    Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
    micelda.Select
    Selection.Merge
    With Selection
        .Value = "DOMINIO"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    
    Set micelda = ActiveCell.Offset(0, 1)
    micelda.Select
    With Selection
        .Value = "PÉRDIDA MEDIA RESIDUAL"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    Set micelda = ActiveCell.Offset(0, 1)
    micelda.Select
    With Selection
        .Value = "PROB"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = False
        .Font.name = "Arial"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .RowHeight = 40
    End With
    Set micelda = ActiveCell.Offset(1, -6)
    micelda.Select

    For j = 1 To 7
        Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 4))
        micelda.Select
        Selection.Merge
        With Selection
            .Value = resultado(j, 2) & " (" & resultado(j, 1) & ")"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        With Selection
            .Value = resultado(j, 6)
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = False
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .NumberFormat = "#,##0 €"
        End With
        Set micelda = ActiveCell.Offset(0, 1)
        micelda.Select
        With Selection
            .Value = resultado(j, 4)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = False
            .Font.Size = 11
            .Font.name = "Calibri"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .NumberFormat = "0 %"
        End With
        Set micelda = ActiveCell.Offset(1, -6)
        micelda.Select
     Next j
     
     
     Set micelda = ActiveCell.Offset(1, 0)
     micelda.Select
     With Selection
        .Value = "Nota: Valores obtenidos a partir de una simulación de " & iteraciones & " iteraciones"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Size = 8
        .Font.name = "Arial"
    End With
    Set micelda = ActiveCell.Offset(3, 0)
    micelda.Select

End Sub

Private Sub informe5()


    name = "5. Requisitos según la IEC-62443 asociados al Nivel de Seguridad Objetivo"
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
        .Font.Underline = False
        .Font.name = "Arial"
    End With
    
    For i = 1 To 7
      
        Set micelda = ActiveCell.Offset(1, 0)
        micelda.Select
    
        name = "5." & i & " " & resultado(i, 2) & " (" & resultado(i, 1) & ")"
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Value = name
            .Font.Bold = True
            .Font.Size = 10
            .Font.Underline = False
            .Font.name = "Arial"
        End With
        
        Set micelda = ActiveCell.Offset(2, 0)
        micelda.Select
    
        If resultado(i, 5) = 0 Then
            ActiveCell.Value = "No aplican requisitos de ciberseguridad al ser el nivel de seguridad objetivo cero."
            Set micelda = ActiveCell.Offset(2, 0)
        Else
            Worksheets("REQUISITOS").Range("A1").CurrentRegion.AutoFilter Field:=3, Criteria1:=resultado(i, 1) 'selecciona el dominio
            Select Case resultado(i, 5)
                Case Is = 1
                    Worksheets("REQUISITOS").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:="=1"
                Case Is = 2
                    Worksheets("REQUISITOS").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:=Array("1", "2"), Operator:=xlFilterValues
                Case Is = 3
                    Worksheets("REQUISITOS").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:=Array("1", "2", "3"), Operator:=xlFilterValues
                Case Is = 4
                    Worksheets("REQUISITOS").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:=Array("1", "2", "3", "4"), Operator:=xlFilterValues
            End Select
                Worksheets("REQUISITOS").Range("A1").CurrentRegion.Resize(, 2).Copy
                Worksheets(hoja).Select
                ActiveCell.Select
                ActiveSheet.Paste
                
            Set micelda = Range(ActiveCell, ActiveCell.Offset(0, 6))
            micelda.Select
            Application.DisplayAlerts = False
            Selection.Merge
            Application.DisplayAlerts = True
            With Selection
                .Value = "REQUISITOS"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Font.Bold = True
                .Font.Size = 10
                .Font.Underline = False
                .Font.name = "Arial"
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThick
                .RowHeight = 40
            End With
            Set micelda = ActiveCell.End(xlDown).Offset(2, 0)
        End If
 
        micelda.Select
    Next i
    Set micelda = ActiveCell.Offset(-2, 0)
    micelda.Select
End Sub

Private Sub informe()

Application.ScreenUpdating = False

Call informe1
ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Call informe2
ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Call informe3
ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Call informe4
ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Call informe5

ActiveWorkbook.Worksheets(hoja).PageSetup.CenterFooter = "Página &P"

Set micelda = Range(ActiveCell, Range("G1"))
micelda.Select
ActiveSheet.PageSetup.PrintArea = Selection.Address

ActiveWindow.DisplayGridlines = False

Set micelda = Range("A1")
micelda.Select

Application.ScreenUpdating = True
Application.CutCopyMode = False
Application.Dialogs(xlDialogPrintPreview).Show

End Sub

