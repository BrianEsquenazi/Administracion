Imports ClasesCompartidas

Public Class Compras

    Dim diasPlazo As Integer = 0
    Dim letrasValidas As New List(Of String) From {"A", "B", "C", "X", "M", "I"}

    Private Sub Compras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim commonEventsHandler As New CommonEventsHandler
        commonEventsHandler.setIndexTab(Me)
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
            txtParidad.Enabled = False
            txtParidad.Text = ""
        Else
            txtParidad.Enabled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIVARG.Leave, txtPercIB.Leave, txtNoGravado.Leave, txtIVA27.Leave, txtIVA21.Leave, txtIVA10.Leave
        Dim total As Double = asDouble(txtIVA21.Text) + asDouble(txtIVARG.Text) + asDouble(txtIVA27.Text) + asDouble(txtPercIB.Text) + asDouble(txtNoGravado.Text) + asDouble(txtIVA10.Text)
        If Not chkSoloIVA.Checked Then
            total += asDouble(txtNeto.Text)
        End If
        txtTotal.Text = Math.Round(total, 2)
    End Sub

    Private Sub txtNeto_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNeto.Leave
        txtIVA21.Text = Math.Round(asDouble(txtNeto.Text) * 0.21, 2)
        txtImporte_Leave(sender, e)
    End Sub

    Private Function asDouble(ByVal text As String)
        Return CustomConvert.toDoubleOrZero(text)
    End Function

    Private Function validarCampos() As Boolean
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(CustomConvert.toIntOr(txtPunto.Text, 0) <> 0, "El campo " & CustomLabel6.Text & " no puede ser cero")
        validador.alsoValidate(CustomConvert.toIntOr(txtNumero.Text, 0) <> 0, "El campo " & CustomLabel7.Text & " no puede ser cero")
        validador.alsoValidate(letrasValidas.Contains(txtLetra.Text) Or txtLetra.Text = "", "El valor ingresado (" & txtLetra.Text & ") no es una letra válida")
        validador.alsoValidate(DAOCompras.mesAbierto(txtFechaEmision.Text), "El mes de la fecha de emisión: " & txtFechaEmision.Text & " se encuentra cerrado según el sistema")

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

    Private Sub chkSoloIVA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoloIVA.CheckedChanged
        txtImporte_Leave(sender, e)
    End Sub

    Private Sub txtLetra_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLetra.TextChanged
        txtLetra.Text = txtLetra.Text.ToUpper
        If txtLetra.Text = "C" Then
            txtIVA21.Enabled = False
            txtIVARG.Enabled = False
            txtIVA27.Enabled = False
            txtPercIB.Enabled = False
            txtNoGravado.Enabled = False
            txtIVA10.Enabled = False
        Else
            txtIVA21.Enabled = True
            txtIVARG.Enabled = True
            txtIVA27.Enabled = True
            txtPercIB.Enabled = True
            txtNoGravado.Enabled = True
            txtIVA10.Enabled = True
        End If
        txtLetra.Select(txtLetra.Text.Count, 1)
    End Sub

    Private Sub txtNeto_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class