Imports ClasesCompartidas

Public Class Pagos

    Private Sub Pagos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbTipo.SelectedIndex = 0
        Dim commonEventHandler As New CommonEventsHandler
        commonEventHandler.setIndexTab(Me)
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtObservaciones_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtObservaciones.Leave
        gridPagos.CurrentCell = gridPagos.Rows(0).Cells(0)
        gridPagos.Select()
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        txtProveedor.Text = proveedor.id
        txtRazonSocial.Text = proveedor.razonSocial
    End Sub

    Private Sub mostrarBanco(ByVal banco As Banco)
        txtBanco.Text = banco.id
        txtNombreBanco.Text = banco.nombre
    End Sub

    Private Sub txtProveedor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProveedor.Leave
        Dim proveedor = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
        If Not IsNothing(proveedor) Then
            mostrarProveedor(proveedor)
        Else
            txtRazonSocial.Text = ""
        End If
    End Sub

    Private Sub txtBanco_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBanco.KeyDown
        If e.KeyValue = Keys.Enter Then
            txtBanco_Leave(sender, Nothing)
        End If
    End Sub

    Private Sub txtBanco_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBanco.Leave
        Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(txtBanco.Text)
        If Not IsNothing(banco) Then
            mostrarBanco(banco)
        Else
            txtNombreBanco.Text = ""
        End If
    End Sub

    Private Sub txtOrdenPago_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrdenPago.Leave
        txtOrdenPago.Text = ceros(txtOrdenPago.Text, 6)
    End Sub
End Class