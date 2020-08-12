Attribute VB_Name = "Módulo4"
Private Sub boton4(control As IRibbonControl)

    Application.Run "Módulo3.informe"

End Sub

Private Sub boton1(control As IRibbonControl)
     
    UserForm4.Show
    
End Sub

Private Sub boton2(control As IRibbonControl)
    Application.Run "Módulo1.simulacion"
    
    Application.Run "Módulo2.calculo_slt"
End Sub

Private Sub CheckBox3(control As IRibbonControl, pressed As Boolean)
    Select Case pressed
        Case True
            Worksheets("IA").Visible = True
            Worksheets("AU").Visible = True
            Worksheets("IS").Visible = True
            Worksheets("CD").Visible = True
            Worksheets("RD").Visible = True
            Worksheets("RE").Visible = True
            Worksheets("DR").Visible = True
        Case False
            Worksheets("IA").Visible = False
            Worksheets("AU").Visible = False
            Worksheets("IS").Visible = False
            Worksheets("CD").Visible = False
            Worksheets("RD").Visible = False
            Worksheets("RE").Visible = False
            Worksheets("DR").Visible = False
    End Select
    

End Sub


