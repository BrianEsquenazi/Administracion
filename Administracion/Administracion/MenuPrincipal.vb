Public Class MenuPrincipal
    Dim forms As New List(Of Form)

    Private Sub abrir(ByVal form As Form)
        forms.Add(form)
        form.Show()
    End Sub

    Private Sub IngresosDeCuentasContablesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresosDeCuentasContablesToolStripMenuItem.Click
        abrir(New CuentaContableABM)
    End Sub

    Private Sub IngresoDeBancosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeBancosToolStripMenuItem.Click
        abrir(New BancosABM)
    End Sub

    Private Sub IngresoDeCambiosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeCambiosToolStripMenuItem.Click
        abrir(New TipoCambioABM)
    End Sub

    Private Sub IngresoDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeProveedoresToolStripMenuItem.Click
        abrir(New ProveedoresABM)
    End Sub

    Private Sub IngresoDeRubrosDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeRubrosDeProveedoresToolStripMenuItem.Click
        abrir(New RubrosProveedorABM)
    End Sub

    Private Sub btnCambio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCambio.Click
        Dim msgResult = vbYes
        If forms.Any(Function(form) form.Visible) Then
            msgResult = MsgBox("¿Se cerrarán todos los formularios abiertos, está seguro que desea cambiar de empresa?", vbYesNo, "Cambiar de Empresa")
        End If
        If msgResult = vbYes Then
            forms.ForEach(Sub(form) form.Dispose())
            Login.Show()
            Close()
        End If
    End Sub

    Private Sub PruebaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepositosToolStripMenuItem.Click
        abrir(New Depositos)
    End Sub

    Private Sub CargarInteresesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarInteresesToolStripMenuItem.Click
        abrir(New CargaIntereses)
    End Sub

    Private Sub SifereToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SifereToolStripMenuItem.Click
        abrir(New ProcesoSifere)
    End Sub

    Private Sub RetencionEsOpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetencionEsOpToolStripMenuItem.Click
        abrir(New ProcesoRetencionesPagos)
    End Sub

    Private Sub RetencionesRecibvosaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetencionesRecibvosaToolStripMenuItem.Click
        abrir(New ProcesoReteRecibos)
    End Sub

    Private Sub FinDelSistemaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinDelSistemaToolStripMenuItem.Click
        Close()
        End
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        abrir(New ProcesoPercepciones)
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        abrir(New CierreMes)
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        abrir(New DepuraCtaCte)
    End Sub

    Private Sub ConsultaDeCuentaCorrientePorPantallaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConsultaDeCuentaCorrientePorPantallaToolStripMenuItem.Click
        abrir(New CuentaCorrientePantalla)
    End Sub

    Private Sub SaldoDeCuentaCorrienteDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaldoDeCuentaCorrienteDeProveedoresToolStripMenuItem.Click
        abrir(New ListadoSaldosCuentaCorrienteProveedores)
    End Sub

    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        abrir(New ListadoIvaCompras)
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        abrir(New ListadoAsientoResumen)
    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        abrir(New ListadoProyeccionCobros)
    End Sub

    Private Sub IngresoDeNovedadesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeNovedadesToolStripMenuItem.Click
        abrir(New Compras)
    End Sub
End Class