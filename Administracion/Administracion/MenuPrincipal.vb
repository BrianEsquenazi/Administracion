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
        Form1.Show()
        'Login.Show()
        'Close()
    End Sub
End Class