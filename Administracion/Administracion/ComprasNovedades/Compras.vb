Imports ClasesCompartidas

Public Class Compras

    Dim diasPlazo As Integer = 0

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
        diasPlazo = CustomConvert.toIntOrZero(proveedor.diasPlazo)
    End Sub

    Private Sub txtCodigoProveedor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoProveedor.Leave
        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtCodigoProveedor.Text)
        If Not IsNothing(proveedor) Then
            mostrarProveedor(proveedor)
        Else
            txtNombreProveedor.Text = ""
            txtCAI.Text = ""
            txtVtoCAI.Text = "  /  /    "
            diasPlazo = 0
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

    Private Sub txtImporte_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNeto.Leave, txtIVARG.Leave, txtPercIB.Leave, txtNoGravado.Leave, txtIVA27.Leave, txtIVA21.Leave, txtIVA10.Leave
        txtTotal.Text = Math.Round((txtNeto.Text) + asDouble(txtIVA21.Text) + asDouble(txtIVARG.Text) + asDouble(txtIVA27.Text) + asDouble(txtPercIB.Text) + asDouble(txtNoGravado.Text) + asDouble(txtIVA10.Text), 2)
    End Sub

    Private Function asDouble(ByVal text As String)
        Return CustomConvert.toDoubleOrZero(text)
    End Function

    Private Function validarCampos() As Boolean
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(CustomConvert.toIntOr(txtPunto.Text, 0) <> 0, "El campo " & CustomLabel6.Text & " no puede ser cero")
        validador.alsoValidate(CustomConvert.toIntOr(txtNumero.Text, 0) <> 0, "El campo " & CustomLabel7.Text & " no puede ser cero")

        Return validador.flush
    End Function

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If validarCampos() Then
            'Agregar
        End If
    End Sub

    Private Sub txtFechaEmision_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFechaEmision.Leave
        Try
            Dim fecha As Date = Convert.ToDateTime(txtFechaEmision.Text)
            txtFechaIVA.Text = fecha.ToShortDateString
            txtFechaVto1.Text = fecha.AddDays(diasPlazo).ToShortDateString()
            txtFechaVto2.Text = txtFechaVto1.Text
        Catch ex As Exception

        End Try
    End Sub
End Class