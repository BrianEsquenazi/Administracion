Public Class MenuPrincipal
    Private Sub IngresosDeCuentasContablesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresosDeCuentasContablesToolStripMenuItem.Click
        CuentaContableABM.Show()
    End Sub

    Private Sub IngresoDeBancosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeBancosToolStripMenuItem.Click
        BancosABM.Show()
    End Sub

    Private Sub IngresoDeCambiosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeCambiosToolStripMenuItem.Click
        TipoCambioABM.Show()
    End Sub

    Private Sub IngresoDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeProveedoresToolStripMenuItem.Click
        ProveedoresABM.Show()
    End Sub

    Private Sub IngresoDeRubrosDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeRubrosDeProveedoresToolStripMenuItem.Click
        RubrosProveedorABM.Show()
    End Sub

    Private Sub btnCambio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCambio.Click
        Login.Show()
        Close()
    End Sub

    Private Sub PruebaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PruebaToolStripMenuItem.Click
        Depositos.Show()
    End Sub

    Private Sub CargarInteresesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarInteresesToolStripMenuItem.Click
        CargaIntereses.Show()
    End Sub


    Private Sub SifereToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SifereToolStripMenuItem.Click
        ProcesoSifere.Show()
    End Sub

    Private Sub RetencionEsOpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetencionEsOpToolStripMenuItem.Click
        ProcesoRetencionesPagos.Show()
    End Sub
    
    Private Sub RetencionesRecibvosaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetencionesRecibvosaToolStripMenuItem.Click
        ProcesoReteRecibos.Show()
    End Sub

    Private Sub FinDelSistemaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinDelSistemaToolStripMenuItem.Click
        Close()
        End
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        ProcesoPercepciones.Show()
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        CierreMes.Show()
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        DepuraCtaCte.Show()
    End Sub

    Private Sub ConsultaDeCuentaCorrientePorPantallaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConsultaDeCuentaCorrientePorPantallaToolStripMenuItem.Click
        CuentaCorrientePantalla.Show()
    End Sub

    Private Sub SaldoDeCuentaCorrienteDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaldoDeCuentaCorrienteDeProveedoresToolStripMenuItem.Click
        ListadoSaldosCuentaCorrienteProveedores.Show()
    End Sub

    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        ListadoIvaCompras.Show()
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        ListadoAsientoResumen.Show()
    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        ListadoProyeccionCobros.Show()
    End Sub

    Private Sub IngresoDeNovedadesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeNovedadesToolStripMenuItem.Click
        Compras.Show()
    End Sub
End Class