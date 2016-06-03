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

<<<<<<< HEAD
    Private Sub IngresoDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeProveedoresToolStripMenuItem.Click
        ProveedoresABM.Show()
=======
    Private Sub IngresoDeRubrosDeProveedoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IngresoDeRubrosDeProveedoresToolStripMenuItem.Click
        RubrosProveedorABM.Show()
>>>>>>> 6fab320be96482a5e6d6ac36a9f377e9e6370fbf
    End Sub
End Class