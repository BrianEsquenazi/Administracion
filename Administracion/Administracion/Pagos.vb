Imports ClasesCompartidas

Public Class Pagos

    Dim queryController As QueryController
    Dim pagos As New List(Of DetalleCompraCuentaCorriente)
    Dim cheques As New List(Of Cheque)
    Dim chequeRow As Integer = -1
    Dim bancoOrden As Banco
    Dim proveedorOrden As Proveedor

    Private Sub Pagos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbTipo.SelectedIndex = 0
        lstSeleccion.Items.Add(New QueryController("Proveedores", AddressOf DAOProveedor.buscarProveedorPorNombre, AddressOf mostrarProveedor))
        lstSeleccion.Items.Add(New QueryController("Cuentas Corrientes", AddressOf cuentasCorrientesDelProveedorActual, AddressOf mostrarCuentaCorriente, False))
        lstSeleccion.Items.Add(New QueryController("Cheques Terceros", AddressOf DAODeposito.buscarCheques, AddressOf mostrarCheque, False))
        lstSeleccion.SelectedIndex = 0

        Dim gridPagosBuilder As New GridBuilder(gridPagos)
        gridPagosBuilder.addTextColumn(0, "Tipo")
        gridPagosBuilder.addTextColumn(1, "Letra")
        gridPagosBuilder.addNumericColumn(2, "Punto")
        gridPagosBuilder.addNumericColumn(3, "Número")
        gridPagosBuilder.addFloatColumn(4, "Importe")
        gridPagosBuilder.addTextColumn(5, "Descripción")

        Dim gridFormasBuilder As New GridBuilder(gridFormaPagos)
        gridFormasBuilder.addTextColumn(0, "Tipo", False)
        gridFormasBuilder.addNumericColumn(1, "Número")
        gridFormasBuilder.addDateColumn(2, "Fecha")
        gridFormasBuilder.addNumericColumn(3, "Banco")
        gridFormasBuilder.addTextColumn(4, "Nombre")
        gridFormasBuilder.addFloatColumn(5, "Importe")

        Dim commonEventHandler As New CommonEventsHandler
        commonEventHandler.setIndexTab(Me)
        txtFecha.Text = Date.Today.ToShortDateString
        txtFechaParidad.Text = Date.Today.ToShortDateString
    End Sub

    Private Function validarDatos() As Boolean
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(consistenciaEntreProveedorYGrillas(), "Algunos campos de las grillas no coinciden con el proveedor que se desea grabar")
        validador.alsoValidate(bancosValidos(), "Algunos campos de la grilla de forma de pagos no tienen un banco válido asignado")
        validador.alsoValidate(noHayDiferencia(), "Hay una diferencia de " & lblDiferencia.Text)
        validador.alsoValidate(hayMovimientos(), "No se registró ningún pago")

        Return validador.flush
    End Function

    Private Function hayMovimientos()
        Return CustomConvert.toDoubleOrZero(lblPagos.Text) <> 0
    End Function

    Private Function noHayDiferencia()
        Return CustomConvert.toDoubleOrZero(lblDiferencia.Text) = 0
    End Function

    Private Function bancosValidos()
        For Each row As DataGridViewRow In gridFormaPagos.Rows
            If Not row.IsNewRow And row.Cells(0).Value = "02" Then
                Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(row.Cells(3).Value)
                If IsNothing(banco) Then : Return False : End If
            End If
        Next
        Return True
    End Function

    Private Function consistenciaEntreProveedorYGrillas()
        Return pagos.All(Function(cuenta) cuenta.proveedor.id = txtProveedor.Text)
    End Function

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtObservaciones_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtObservaciones.Leave
        If gridPagos.Rows.Count = 0 Then
            lstSeleccion.SelectedIndex = 1
            lstSeleccion_DoubleClick(Nothing, Nothing)
        Else
            gridPagos.CurrentCell = gridPagos.Rows(0).Cells(4)
            gridPagos.Select()
            gridPagos.Focus()
        End If
    End Sub

    Private Function cuentasCorrientesDelProveedorActual()
        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
        If IsNothing(proveedor) Then
            Return New List(Of CtaCteProveedor)
        Else
            Return DAOCtaCteProveedor.cuentasSinSaldar(proveedor)
        End If
    End Function

    Private Sub mostrarCuentaCorriente(ByVal cuenta As DetalleCompraCuentaCorriente)
        If pagos.Any(Function(pagoExistente) cuenta.igualA(pagoExistente)) Then
            Exit Sub
        End If
        pagos.Add(cuenta)
        gridPagos.Rows.Add(cuenta.tipo, cuenta.letra, cuenta.punto, cuenta.numero, CustomConvert.toStringWithTwoDecimalPlaces(cuenta.saldo), "Pago Factura Nro " & CustomConvert.toIntOrZero(cuenta.numero))
    End Sub

    Private Sub mostrarOrdenDePago(ByVal orden As OrdenPago)
        If IsNothing(orden) Then : Exit Sub : End If
        'btnLimpiar.PerformClick()
        txtOrdenPago.Text = orden.nroOrden
        txtFecha.Text = orden.fecha
        txtObservaciones.Text = orden.observaciones
        txtFechaParidad.Text = orden.fechaParidad
        mostrarProveedor(orden.proveedor)
        mostrarBanco(orden.banco)
        txtGanancias.Text = CustomConvert.toStringWithTwoDecimalPlaces(orden.retGanancias)
        txtIBCiudad.Text = CustomConvert.toStringWithTwoDecimalPlaces(orden.retIBCiudad)
        txtIngresosBrutos.Text = CustomConvert.toStringWithTwoDecimalPlaces(orden.retIB)
        txtIVA.Text = CustomConvert.toStringWithTwoDecimalPlaces(orden.retIVA)
        mostrarPagos(orden.pagos)
        mostrarFormaPagos(orden.formaPagos)
        mostrarTipo(orden.tipo)
    End Sub

    Private Sub mostrarTipo(ByVal tipo As Integer)
        Select Case tipo
            Case 1
                optCtaCte.Checked = True
            Case 3
                optChequeRechazado.Checked = True
            Case 4
                optAnticipos.Checked = True
            Case 5
                optTransferencias.Checked = True
            Case Else
                optVarios.Checked = True
        End Select
    End Sub

    Private Sub mostrarPagos(ByVal pagos As List(Of Pago))
        gridPagos.Rows.Clear()
        For Each pago As Pago In pagos
            gridPagos.Rows.Add(pago.tipo, pago.letra, pago.punto, pago.numero, pago.importe, pago.descripcion)
        Next
        sumarImportes()
    End Sub

    Private Sub mostrarFormaPagos(ByVal formaPagos As List(Of FormaPago))
        gridFormaPagos.Rows.Clear()
        For Each formaPago As FormaPago In formaPagos
            gridFormaPagos.Rows.Add(formaPago.tipo, formaPago.numero, formaPago.fecha, formaPago.banco, formaPago.nombre, formaPago.importe)
        Next
        sumarImportes()
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        If IsNothing(proveedor) Then : Exit Sub : End If
        txtProveedor.Text = proveedor.id
        txtRazonSocial.Text = proveedor.razonSocial
    End Sub

    Private Sub mostrarBanco(ByVal banco As Banco)
        If IsNothing(banco) Then : Exit Sub : End If
        txtBanco.Text = banco.id
        txtNombreBanco.Text = banco.nombre
    End Sub

    Private Sub mostrarCuentaContable(ByVal cuenta As CuentaContable)
        'TODO
    End Sub

    Private Sub mostrarCheque(ByVal cheque As Cheque)
        If cheques.Any(Function(chequeExistente) cheque.igualA(chequeExistente)) Then
            chequeRow = -1
            gridFormaPagos.Select()
            Exit Sub
        End If
        cheques.Add(cheque)
        If chequeRow <> -1 Then
            gridFormaPagos.Rows(chequeRow).Cells(0).Value = "03"
            gridFormaPagos.Rows(chequeRow).Cells(1).Value = cheque.numero
            gridFormaPagos.Rows(chequeRow).Cells(2).Value = cheque.fecha
            gridFormaPagos.Rows(chequeRow).Cells(3).Value = ""
            gridFormaPagos.Rows(chequeRow).Cells(4).Value = cheque.banco
            gridFormaPagos.Rows(chequeRow).Cells(5).Value = CustomConvert.toStringWithTwoDecimalPlaces(cheque.importe)
            gridFormaPagos.CurrentCell = gridFormaPagos.Rows(chequeRow + 1).Cells(0)
            gridFormaPagos.Select()
            chequeRow = -1
        Else
            gridFormaPagos.Rows.Add("03", cheque.numero, cheque.fecha, "", cheque.banco, CustomConvert.toStringWithTwoDecimalPlaces(cheque.importe))
        End If
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
        mostrarOrdenDePago(DAOPagos.buscarOrdenPorNumero(txtOrdenPago.Text))
    End Sub

    Private Sub lstSeleccion_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstSeleccion.DoubleClick
        queryController = lstSeleccion.SelectedItem
        lstSeleccion.Visible = False
        lstConsulta.Visible = True
        txtConsulta.Visible = queryController.usesQueryText
        If txtConsulta.Visible Then
            lstConsulta.Height = 108
            lstConsulta.Top = 38
        Else
            lstConsulta.Height = lstSeleccion.Height
            lstConsulta.Top = lstSeleccion.Top
        End If
        lstConsulta.DataSource = queryController.query.Invoke("")
        txtConsulta.Focus()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click
        lstConsulta.Visible = False
        txtConsulta.Visible = False
        lstSeleccion.Visible = True
    End Sub

    Private Sub txtConsulta_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsulta.KeyDown
        If e.KeyValue = Keys.Enter Then
            lstConsulta.DataSource = queryController.query.Invoke(txtConsulta.Text)
            e.Handled = True
        End If
    End Sub

    Private Sub lstConsulta_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstConsulta.DoubleClick
        queryController.showMethod.Invoke(lstConsulta.SelectedValue)
        lstConsulta.Visible = False
        txtConsulta.Visible = False
        txtConsulta.Text = ""
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        gridPagos.Rows.Clear()
        pagos.Clear()
        gridFormaPagos.Rows.Clear()
        cheques.Clear()
    End Sub

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If validarDatos() Then
            Dim siguienteNumero As Integer = DAOPagos.siguienteNumeroDeOrden()

            Dim pago As OrdenPago = New OrdenPago(siguienteNumero, tipoOrden, CustomConvert.toDoubleOrZero(txtParidad.Text),
                                                  CustomConvert.toDoubleOrZero(txtTotal.Text), CustomConvert.toDoubleOrZero(txtIVA.Text),
                                                  CustomConvert.toDoubleOrZero(txtIngresosBrutos.Text), CustomConvert.toDoubleOrZero(txtIBCiudad.Text),
                                                  CustomConvert.toDoubleOrZero(txtGanancias.Text), txtFecha.Text, txtFechaParidad.Text, txtObservaciones.Text,
                                                  bancoOrden, proveedorOrden)
            DAOPagos.agregarPago(pago)
        End If
    End Sub

    Private Function tipoOrden()
        If optCtaCte.Checked Then : Return 1 : End If
        If optVarios.Checked Then : Return 2 : End If
        If optChequeRechazado.Checked Then : Return 3 : End If
        If optAnticipos.Checked Then : Return 4 : End If
        If optTransferencias.Checked Then : Return 5 : End If
        Return Nothing
    End Function

    Private Sub txtFecha_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFecha.Leave
        txtFechaParidad.Text = txtFecha.Text
    End Sub

    Private Sub eventoSegunTipoEnFormaDePagoPara(ByVal val As Integer, ByVal rowIndex As Integer, ByVal columnIndex As Integer)
        Dim nombre As String = ""
        Dim column As Integer = columnIndex
        Select Case val
            Case 1
                nombre = "Efectivo"
                column = 5
            Case 2
                column = 1
            Case 3
                chequeRow = rowIndex
                lstSeleccion.SelectedIndex = 2
                lstSeleccion_DoubleClick(Nothing, Nothing)
                Exit Sub
            Case 5
                nombre = "US$"
                column = 5
            Case 6
                nombre = "Varios"
                column = 5
            Case Else
                Exit Sub
        End Select
        gridFormaPagos.CurrentCell.Value = ceros(val.ToString, 2)
        gridFormaPagos.Rows(rowIndex).Cells(4).Value = nombre
        gridFormaPagos.CurrentCell = gridFormaPagos.Rows(rowIndex).Cells(column)
    End Sub

    Private Sub gridFormaPagos_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormaPagos.CellLeave
        If e.ColumnIndex = 3 And e.RowIndex > -1 Then
            Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(gridFormaPagos.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            If Not IsNothing(banco) Then
                gridFormaPagos.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value = banco.nombre
            End If
        End If
    End Sub

    Private Sub gridFormaPagos_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormaPagos.CellValueChanged
        sumarImportes()
        llenarConCerosNumero()
    End Sub

    Private Sub llenarConCerosNumero()
        For Each row As DataGridViewRow In gridFormaPagos.Rows
            If row.Cells(1).Value <> "" Then
                If row.Cells(1).Value.ToString.Length > 8 Then
                    row.Cells(1).Value = ""
                Else
                    row.Cells(1).Value = ceros(row.Cells(1).Value, 8)
                End If
            End If
        Next
    End Sub

    Private Sub gridFormaPagos_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridFormaPagos.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim iCol = gridFormaPagos.CurrentCell.ColumnIndex
            Dim iRow = gridFormaPagos.CurrentCell.RowIndex
            If iCol = 0 And iRow > -1 Then
                Dim val = gridFormaPagos.Rows(iRow).Cells(iCol).Value
                eventoSegunTipoEnFormaDePagoPara(CustomConvert.toIntOrZero(val), iRow, iCol)
            End If
        End If
    End Sub

    Private Sub gridPagos_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridPagos.CellValueChanged
        sumarImportes()
    End Sub

    Private Sub sumarImportes()
        Dim pagos As Double = 0
        Dim formaPagos As Double = 0
        Dim total As Double = 0

        total = CustomConvert.toDoubleOrZero(txtIVA.Text) + CustomConvert.toDoubleOrZero(txtGanancias.Text) + CustomConvert.toDoubleOrZero(txtIBCiudad.Text) +
            CustomConvert.toDoubleOrZero(txtIngresosBrutos.Text)

        For Each row As DataGridViewRow In gridPagos.Rows
            If Not row.IsNewRow Then
                pagos += CustomConvert.toDoubleOrZero(row.Cells(4).Value)
            End If
        Next

        For Each row As DataGridViewRow In gridFormaPagos.Rows
            If Not row.IsNewRow Then
                formaPagos += CustomConvert.toDoubleOrZero(row.Cells(5).Value)
            End If
        Next
        txtTotal.Text = CustomConvert.toStringWithTwoDecimalPlaces(total)
        lblPagos.Text = CustomConvert.toStringWithTwoDecimalPlaces(pagos)
        lblFormaPagos.Text = CustomConvert.toStringWithTwoDecimalPlaces(formaPagos + total)
        lblDiferencia.Text = CustomConvert.toStringWithTwoDecimalPlaces(pagos - formaPagos - total)
    End Sub

    Private Sub gridPagos_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles gridPagos.RowsAdded
        sumarImportes()
    End Sub

    Private Sub gridFormaPagos_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles gridFormaPagos.RowsAdded
        sumarImportes()
    End Sub

    Private Sub gridPagos_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles gridPagos.UserDeletedRow
        If e.Row.Cells(0).Value <> "" Then
            Dim detalle As DetalleCompraCuentaCorriente = pagos.Find(Function(pago) pago.tipo = e.Row.Cells(0).Value And pago.letra = e.Row.Cells(1).Value And
                                                                         pago.punto = e.Row.Cells(2).Value And pago.numero = e.Row.Cells(3).Value And
                                                                         pago.saldo = CustomConvert.toDoubleOrZero(e.Row.Cells(4).Value))
            If Not IsNothing(detalle) Then
                pagos.Remove(detalle)
            End If
        End If
        sumarImportes()
    End Sub

    Private Sub gridFormaPagos_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles gridFormaPagos.UserDeletedRow
        If e.Row.Cells(0).Value = "03" Then
            Dim chequeABorrar As Cheque = cheques.Find(Function(cheque) cheque.numero = e.Row.Cells(1).Value And cheque.fecha = e.Row.Cells(2).Value And cheque.banco = e.Row.Cells(4).Value And cheque.importe = e.Row.Cells(5).Value)
            If Not IsNothing(chequeABorrar) Then
                cheques.Remove(chequeABorrar)
            End If
        End If
        sumarImportes()
    End Sub

    Private Sub optTransferencias_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTransferencias.CheckedChanged
        If optTransferencias.Checked Then
            txtBanco.Enabled = True
            txtBanco.Empty = False
            txtNombreBanco.Empty = False
            txtBanco.Text = ""
            txtBanco.Focus()
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "", "")
            gridPagos.Columns(4).ReadOnly = False
            gridPagos.Columns(5).ReadOnly = False
        End If
    End Sub

    Private Sub optAnticipos_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAnticipos.CheckedChanged
        If optAnticipos.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "0,00", txtRazonSocial.Text)
            gridPagos.Columns(4).ReadOnly = True
            gridPagos.Columns(5).ReadOnly = True
        End If
    End Sub

    Private Sub optVarios_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVarios.CheckedChanged
        If optVarios.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "", txtRazonSocial.Text)
            gridPagos.Columns(4).ReadOnly = False
            gridPagos.Columns(5).ReadOnly = True
        End If
    End Sub

    Private Sub optChequeRechazado_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChequeRechazado.CheckedChanged
        If optChequeRechazado.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "", "")
            gridPagos.Columns(4).ReadOnly = False
            gridPagos.Columns(5).ReadOnly = False
        End If
    End Sub

    Private Sub optCtaCte_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCtaCte.CheckedChanged
        If optCtaCte.Checked Then
            gridPagos.Rows.Clear()
            Try
                gridPagos.Columns(4).ReadOnly = True
                gridPagos.Columns(5).ReadOnly = True
            Catch ex As ArgumentOutOfRangeException
            End Try
        End If
    End Sub
End Class