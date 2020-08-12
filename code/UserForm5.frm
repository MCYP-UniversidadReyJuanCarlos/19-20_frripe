VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "DESCRIPCIÓN DE VULNERABILIDADES"
   ClientHeight    =   12405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22650
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label39_Click()

End Sub

Private Sub Salir_Click()

    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim k, i, j As Integer
    For k = 1 To 40
        Me.Controls("Label" & k).Visible = False
        Me.Controls("Label" & k).Caption = " "
    Next k
    
    If UserForm4.si_vul > 10 Then
        Me.Label39.Visible = True
        Me.Label39.Caption = "Vulnerabilidad"
        Me.Label40.Visible = True
        Me.Label40.Caption = "Descripción"
    End If
    
    
    'Mostrar Listbox
    For k = 1 To 2 * UserForm4.si_vul
        UserForm5.Controls("Label" & k).Visible = True
    Next k

    'Mostrar Vulnerabilidades
    j = 1
    For i = 1 To 39
        If UserForm4.rango_vul.Cells(i, 1) = "SI" Then
            UserForm5.Controls("Label" & j).Caption = Worksheets("VULNERABILIDADES").Cells(i + 5, 3).Value
            j = j + 1
            UserForm5.Controls("Label" & j).Caption = Worksheets("VULNERABILIDADES").Cells(i + 5, 4).Value
            j = j + 1
        End If
    Next i
End Sub
