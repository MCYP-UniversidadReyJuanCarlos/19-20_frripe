VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "ESTIMACIÓN DE PROBABILIDAD DE LOS ESCENARIOS"
   ClientHeight    =   10275
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   19500
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public si_vul, numero_escenario As String
    Public rango_vul As Range
    


Private Sub ComboBox1_Change()
    Dim ocupadas As String
    Dim k, i, j As Integer
    
    Me.Label2.Visible = True
    
    For k = 1 To 30
        Me.Controls("CheckBox" & k).Visible = False
        Me.Controls("CheckBox" & k).Caption = " "
        Me.Controls("CheckBox" & k).Value = False
    Next k
    
    'Mostrar dominio
    ocupadas = Application.CountIf(Worksheets("ESCENARIOS").Range("C21:C320"), "<>")
    For i = 1 To ocupadas
        If Worksheets("ESCENARIOS").Cells(20 + i, 3).Value = Me.ComboBox1.Value Then
            Me.TextBox1.Text = Worksheets("ESCENARIOS").Cells(20 + i, 2).Value
            numero_escenario = i
        End If
    Next i
    
    'Mostrar Vulnerabilidades
    Select Case Me.TextBox1.Value
        Case Is = "IA"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("E6:E44")
        Case Is = "AU"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("F6:F44")
        Case Is = "IS"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("G6:G44")
        Case Is = "CD"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("H6:H44")
        Case Is = "RD"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("I6:I44")
        Case Is = "RE"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("J6:J44")
        Case Is = "DR"
            Set rango_vul = Worksheets("VULNERABILIDADES").Range("K6:K44")
    End Select
    
    'Mostrar Checkbox
    si_vul = Application.CountIf(rango_vul, "SI")
    For k = 1 To si_vul
        Me.Controls("CheckBox" & k).Visible = True
    Next k
  
    'Mostrar Vulnerabilidades
    j = 1
    For i = 1 To 39
        If rango_vul.Cells(i, 1) = "SI" Then
            Controls("CheckBox" & j).Caption = Worksheets("VULNERABILIDADES").Cells(i + 5, 3).Value
            j = j + 1
        End If
    Next i
    
    'Habilitar botones
    Me.CommandButton1.Enabled = True
    Me.CommandButton2.Enabled = True
    
    
End Sub


Private Sub CommandButton1_Click()
    Dim Probabilidad As Double
    Dim cuenta, i As Integer
    
    cuenta = 0

    For i = 1 To 30
        If Controls("CheckBox" & i) = True Then
            cuenta = cuenta + 1
        End If
    Next
    Probabilidad = Application.WorksheetFunction.Round((cuenta / si_vul * 80), 0)
    
  
    
    If cuenta = 0 Then
       MsgBox "Es necesario seleccionar las vulnerabilidades asociadas al escanario para realizar la estimación de probabilidad"
    Else
       MsgBox "La probabilidad estimada del escenario es: " & Probabilidad & " %"
       Worksheets("ESCENARIOS").Cells(20 + UserForm4.numero_escenario, 6).Value = Probabilidad * 0.01
    End If
    
End Sub

Private Sub CommandButton2_Click()
    UserForm5.Show
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Salir_Click()
    Unload Me
End Sub


Private Sub UserForm_Activate()
    Dim i As Integer
    
    Worksheets("ESCENARIOS").Activate
    Dim ocupadas As String
    Dim rango, celda As Range
    ocupadas = Application.CountIf(Worksheets("ESCENARIOS").Range("C21:C320"), "<>")
    'Comprobar dominios
    Set rango = Worksheets("ESCENARIOS").Range("$B$21", Cells(20 + ocupadas, 3))
    
    i = 0
    For Each celda In rango
        If IsEmpty(celda.Value) Then
            i = i + 1
        End If
    Next celda
    
    If i > 0 Then
        MsgBox "Para estimar la probabilidad es necesario asignar todos los escenarios a un dominio"
    Else
        'Listar escenarios
        Set rango = Worksheets("ESCENARIOS").Range("$C$21", Cells(20 + ocupadas, 3))

        For Each celda In rango
            ComboBox1.AddItem celda.Value
        Next celda
    End If
   
       
End Sub

Private Sub UserForm_Initialize()
    'Esconder todos los checkbox
    Dim k As Integer
    For k = 1 To 30
        Me.Controls("CheckBox" & k).Visible = False
    Next k
 
    Me.Label2.Visible = False
    
    Me.CommandButton1.Caption = "ESTIMAR" & vbCrLf & "PROBABILIDAD"
    Me.CommandButton2.Caption = "DESCRIPCIÓN" & vbCrLf & "VULNERABILIDADES"
    Me.CommandButton1.Enabled = False
    Me.CommandButton2.Enabled = False
End Sub



