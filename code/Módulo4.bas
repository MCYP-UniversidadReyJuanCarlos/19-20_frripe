Attribute VB_Name = "M�dulo4"
Private Sub boton4(control As IRibbonControl)

    Application.Run "M�dulo3.informe"

End Sub

Private Sub boton1(control As IRibbonControl)
     
    UserForm4.Show
    
End Sub

Private Sub boton2(control As IRibbonControl)
    Application.Run "M�dulo1.simulacion"
    
    Application.Run "M�dulo2.calculo_slt"
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


