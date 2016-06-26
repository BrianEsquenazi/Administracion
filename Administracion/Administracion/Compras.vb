Public Class Compras

    Private Sub Compras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTab(Me)
    End Sub

    Private Sub txtNumero_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNumero.Leave
        'Complete with zeros
    End Sub

    Private Sub txtPunto_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPunto.Leave
        'Complete with zeros
    End Sub
End Class