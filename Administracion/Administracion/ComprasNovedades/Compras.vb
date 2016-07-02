﻿Imports ClasesCompartidas

Public Class Compras

    Dim diasPlazo As Integer = 0
    Dim letrasValidas As New List(Of String) From {"A", "B", "C", "X", "M", "I"}
    Dim proveedor As Proveedor

    Private Sub Compras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim gridBuilder As New GridBuilder(gridAsientos)

        gridBuilder.addTextColumn(0, "Cuenta")
        gridBuilder.addTextColumn(1, "Descripción")
        gridBuilder.addPositiveFloatColumn(2, "Débito")
        gridBuilder.addPositiveFloatColumn(3, "Crédito")

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
            txtTipo.Text = "0" & cmbTipo.SelectedIndex + 1
        End If
        gridAsientos.Rows.Clear()
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        gridAsientos.Rows.Clear()
        chkSoloIVA.Checked = False
        optEfectivo.Checked = True
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
        proveedor = DAOProveedor.buscarProveedorPorCodigo(txtCodigoProveedor.Text)
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
        Dim total As Double = asDouble(txtIVA21.Text) + asDouble(txtIVARG.Text) + asDouble(txtIVA27.Text) + asDouble(txtPercIB.Text) + asDouble(txtNoGravado.Text) + asDouble(txtIVA10.Text) + asDouble(txtNeto.Text)
        txtTotal.Text = Math.Round(total, 2)
    End Sub

    Private Sub txtNeto_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNeto.Leave
        If txtIVA21.Enabled Then
            txtIVA21.Text = Math.Round(asDouble(txtNeto.Text) * 0.21, 2)
        End If
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
        validador.alsoValidate(DAOCierreMes.mesAbierto(txtFechaEmision.Text), "El mes de la fecha de emisión: " & txtFechaEmision.Text & " se encuentra cerrado según el sistema")
        validador.alsoValidate(gridAsientos.Rows.Count > 1, "No fue generado el asiento. No se puede confirmar")
        validador.alsoValidate(lblCredito.Text = lblDebito.Text, "El asiento se encuentra desbalanceado. Hay una diferencia de: " & Math.Abs(asDouble(lblCredito.Text) - asDouble(lblDebito.Text)))
        validador.alsoValidate(asientosCorrectos(), "El asiento se encuentra en un estado inválido, puede que falte asignar alguna cuenta")
        validador.alsoValidate(valoresDebeYHaberCorrectos(), "Una entrada del asiento tiene valores inválidos de Débito y/o Crédito")
        validador.alsoValidate(asDouble(lblDebito.Text) = asDouble(txtTotal.Text), "El total del asiento contable tiene que ser igual al importe total")

        Return validador.flush
    End Function

    Private Function valoresDebeYHaberCorrectos()
        Dim estado As Boolean = True
        For Each row As DataGridViewRow In gridAsientos.Rows
            estado = estado And (asDouble(row.Cells(2).Value) = 0 Xor asDouble(row.Cells(3).Value) = 0) _
                And asDouble(row.Cells(2).Value) >= 0 And asDouble(row.Cells(3).Value) >= 0
        Next
        Return estado
    End Function

    Private Function asientosCorrectos()
        Dim estado As Boolean = True
        For Each row As DataGridViewRow In gridAsientos.Rows
            estado = estado And row.Cells(1).Value <> ""
        Next
        Return estado
    End Function

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If validarCampos() Then
            actualizarProveedor()
            Dim compra As Compra = crearCompra()
            txtNroInterno.Text = compra.nroInterno
            DAOCompras.agregarCompra(compra)
            MsgBox("El número de interno asignado es: " & compra.nroInterno)
            'Agregar
        End If
    End Sub

    Private Function crearCompra() As Compra
        Dim compra As New Compra(DAOCompras.siguienteNumeroDeInterno(), proveedor, txtTipo.Text, cmbTipo.SelectedValue, cmbFormaPago.SelectedIndex,
                                 tipoPago(), txtLetra.Text, txtNumero.Text, txtFechaEmision.Text, txtFechaIVA.Text, txtFechaVto1.Text, txtFechaVto2.Text,
                                 asDouble(txtParidad.Text), asDouble(txtNeto.Text), asDouble(txtIVA21.Text), asDouble(txtIVARG.Text), asDouble(txtIVA27.Text),
                                 asDouble(txtPercIB.Text), asDouble(txtNoGravado.Text), asDouble(txtIVA10.Text), asDouble(txtTotal.Text), chkSoloIVA.Checked)
        'TODO AGREGAR LOS ASIENTOS
        Return compra
    End Function

    Private Function tipoPago() As Integer
        If optEfectivo.Checked Then : Return 1 : End If
        If optCtaCte.Checked Then : Return 2 : End If
        Return 3
    End Function

    Private Sub actualizarProveedor()
        proveedor.cai = txtCAI.Text
        proveedor.vtoCAI = txtVtoCAI.Text
        If IsNothing(proveedor.cuenta) Then
            proveedor.cuenta = DAOProveedor.cuentaDefault
        End If
        DAOProveedor.agregarProveedor(proveedor)
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
        If chkSoloIVA.Checked Then
            txtNeto.Text = 0
        End If
        txtNeto.Enabled = Not chkSoloIVA.Checked
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
            txtIVA21.Text = "0"
            txtIVARG.Text = "0"
            txtIVA27.Text = "0"
            txtPercIB.Text = "0"
            txtNoGravado.Text = "0"
            txtIVA10.Text = "0"
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

    Private Sub txtDespacho_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDespacho.Leave
        Dim cuenta As CuentaContable
        If IsNothing(proveedor) OrElse IsNothing(proveedor.cuenta) Then
            cuenta = DAOProveedor.cuentaDefault
        Else
            cuenta = proveedor.cuenta
        End If

        crearAsientoContableUsando(cuenta)
        gridAsientos.CurrentCell = gridAsientos.Rows(gridAsientos.Rows.Count - 1).Cells(0)
        gridAsientos.Select()
    End Sub

    Private Function esNotaDeCredito()
        Return txtTipo.Text = "3"
    End Function

    Private Function cuentaIVACredito() As CuentaContable
        Return DAOCuentaContable.IVACredito()
    End Function

    Private Function cuentaIVADebito() As CuentaContable
        Return DAOCuentaContable.IVADebito()
    End Function

    Private Function cuentaIngresosBrutos() As CuentaContable
        Return DAOCuentaContable.ingresosBrutos
    End Function

    Private Function cuentaIVARG3337() As CuentaContable
        Return DAOCuentaContable.IVARG3337
    End Function

    Private Sub crearAsientoContableUsando(ByVal cuenta As CuentaContable)
        gridAsientos.Rows.Clear()
        Dim total As Double = asDouble(txtTotal.Text)
        Dim sumaIvas As Double = asDouble(txtIVA10.Text) + asDouble(txtIVA21.Text) + asDouble(txtIVA27.Text)
        Dim ivaRG3337 As Double = asDouble(txtIVARG.Text)
        Dim ingresosBrutos As Double = asDouble(txtPercIB.Text)
        Dim diferencia As Double = total - sumaIvas - ingresosBrutos - ivaRG3337

        If esNotaDeCredito() Then
            If total <> 0 Then : gridAsientos.Rows.Add(cuenta.id, cuenta.descripcion, total, 0)
            End If
            If sumaIvas <> 0 Then : gridAsientos.Rows.Add(cuentaIVADebito.id, cuentaIVACredito.descripcion, 0, sumaIvas)
            End If
            If ivaRG3337 <> 0 Then : gridAsientos.Rows.Add(cuentaIVARG3337.id, cuentaIVARG3337.descripcion, 0, ivaRG3337)
            End If
            If ingresosBrutos <> 0 Then : gridAsientos.Rows.Add(cuentaIngresosBrutos.id, cuentaIngresosBrutos.descripcion, 0, ingresosBrutos)
            End If
            gridAsientos.Rows.Add("", "", 0, diferencia)
        Else
            If total <> 0 Then : gridAsientos.Rows.Add(cuenta.id, cuenta.descripcion, 0, total)
            End If
            If sumaIvas <> 0 Then : gridAsientos.Rows.Add(cuentaIVACredito.id, cuentaIVACredito.descripcion, sumaIvas, 0)
            End If
            If ivaRG3337 <> 0 Then : gridAsientos.Rows.Add(cuentaIVARG3337.id, cuentaIVARG3337.descripcion, ivaRG3337, 0)
            End If
            If ingresosBrutos <> 0 Then : gridAsientos.Rows.Add(cuentaIngresosBrutos.id, cuentaIngresosBrutos.descripcion, ingresosBrutos, 0)
            End If
            gridAsientos.Rows.Add("", "", diferencia, 0)
        End If

        calcularAsiento()
    End Sub

    Private Sub calcularAsiento()
        Dim valorDebe As Double = 0
        Dim valorHaber As Double = 0
        For Each row As DataGridViewRow In gridAsientos.Rows
            valorDebe += asDouble(row.Cells(2).Value)
            valorHaber += asDouble(row.Cells(3).Value)
        Next
        lblDebito.Text = valorDebe
        lblCredito.Text = valorHaber
    End Sub

    Private Sub gridAsientos_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridAsientos.CellValueChanged
        calcularAsiento()
        If e.RowIndex < 0 Or e.ColumnIndex < 0 Then
            Exit Sub
        End If
        If e.ColumnIndex = 0 Then
            Dim cuenta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(gridAsientos.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            If Not IsNothing(cuenta) Then
                gridAsientos.Rows(e.RowIndex).Cells(1).Value = cuenta.descripcion
            Else
                gridAsientos.Rows(e.RowIndex).Cells(1).Value = ""
            End If
        End If
    End Sub
End Class