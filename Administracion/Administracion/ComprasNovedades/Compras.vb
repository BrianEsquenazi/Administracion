Imports ClasesCompartidas

Public Class Compras

    Private Sub Compras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTab(Me)
    End Sub

    Private Sub txtNumero_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNumero.Leave
        txtNumero.Text = ceros(txtNumero.Text, 8)
    End Sub

    Private Sub txtPunto_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPunto.Leave
        txtPunto.Text = ceros(txtPunto.Text, 4)
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtTipo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTipo.Leave
        Dim tipo As Integer = Val(txtTipo.Text)
        If 1 <= tipo And tipo <= 3 Then
            cmbTipo.SelectedIndex = tipo - 1
        Else
            cmbTipo.SelectedIndex = -1
        End If
    End Sub

    Private Sub cmbTipo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTipo.SelectedIndexChanged
        If cmbTipo.SelectedIndex <> -1 Then
            txtTipo.Text = cmbTipo.SelectedIndex + 1
        End If
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        If Not proveedor.estaDefinidoCompleto Then
            proveedor = DAOProveedor.buscarProveedorPorCodigo(proveedor.id)
        End If
        txtNombreProveedor.Text = proveedor.razonSocial
        txtCAI.Text = proveedor.cai
        txtVtoCAI.Text = proveedor.vtoCAI
    End Sub

    Private Sub txtCodigoProveedor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoProveedor.Leave
        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtCodigoProveedor.Text)
        If Not IsNothing(proveedor) Then
            mostrarProveedor(proveedor)
        Else
            txtNombreProveedor.Text = ""
            txtCAI.Text = ""
            txtVtoCAI.Text = "  /  /    "
        End If
    End Sub

    Private Sub cmbFormaPago_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbFormaPago.KeyDown
        If e.KeyValue = Keys.Enter Then
            cmbFormaPago_SelectedIndexChanged(sender, Nothing)
        End If
    End Sub

    Private Sub cmbFormaPago_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFormaPago.Leave
        cmbFormaPago_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub cmbFormaPago_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFormaPago.SelectedIndexChanged
        txtParidad.Empty = cmbFormaPago.SelectedIndex <> 2
        If txtParidad.Empty Then
            If txtParidad.EnterIndex <> -1 Then
                txtNeto.EnterIndex = txtParidad.EnterIndex
                txtParidad.EnterIndex = -1
            End If
        Else
            txtNeto_Enter(sender, e)
        End If
    End Sub

    Private Sub txtNeto_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNeto.Enter
        If txtParidad.EnterIndex = -1 Then
            txtParidad.EnterIndex = txtNeto.EnterIndex
            txtNeto.EnterIndex = txtNeto.EnterIndex + 1
        End If
    End Sub
End Class