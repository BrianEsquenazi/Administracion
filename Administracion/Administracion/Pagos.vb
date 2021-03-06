﻿Imports ClasesCompartidas

Public Class Pagos

    Dim queryController As QueryController
    Dim pagos As New List(Of DetalleCompraCuentaCorriente)
    Dim cheques As New List(Of Cheque)
    Dim chequeRow As Integer = -1
    Dim bancoOrden As Banco
    Dim proveedorOrden As Proveedor
    Dim commonEventHandler As New CommonEventsHandler

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

        commonEventHandler.setIndexTab(Me)
        btnLimpiar.PerformClick()
    End Sub

    Private Function validarDatos() As Boolean
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(consistenciaEntreProveedorYGrillas(), "Algunos campos de las grillas no coinciden con el proveedor que se desea grabar")
        validador.alsoValidate(bancosValidos(), "Algunos campos de la grilla de forma de pagos no tienen un banco válido asignado")
        validador.alsoValidate(noHayDiferencia(), "Hay una diferencia de " & lblDiferencia.Text)
        validador.alsoValidate(hayMovimientos(), "No se registró ningún pago")
        validador.alsoValidate(CustomConvert.toIntOrZero(txtOrdenPago.Text) = 0, "No se puede hacer el alta, el registro ya existe")

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
            lstSeleccion_Click(Nothing, Nothing)
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
        If (LTrim(txtParidad.Text) = "") Then
            MessageBox.Show("No hay paridad informada")
        Else
            If pagos.Any(Function(pagoExistente) cuenta.igualA(pagoExistente)) Then
                Exit Sub
            End If
            pagos.Add(cuenta)
            gridPagos.Rows.Add(cuenta.tipo, cuenta.letra, cuenta.punto, cuenta.numero, CustomConvert.toStringWithTwoDecimalPlaces(cuenta.saldo), "Pago Factura Nro " & CustomConvert.toIntOrZero(cuenta.numero))

            If cuenta.esClausulaDolar Then
                generarNota(cuenta)
            End If
        End If
    End Sub

    Private Sub generarNota(ByVal cuenta As DetalleCompraCuentaCorriente)
        Dim resto As Double

        resto = (cuenta.montoDolar() * CustomConvert.toStringWithTwoDecimalPlaces(txtParidad.Text)) - CustomConvert.toStringWithTwoDecimalPlaces(cuenta.saldo)



        Select Case resto
            Case Is < 0
                gridPagos.Rows.Add("03", cuenta.letra, cuenta.punto, "99999999", CustomConvert.toStringWithTwoDecimalPlaces(resto), "N/C por Diferencia de Cambio")
            Case Is > 0
                gridPagos.Rows.Add("02", cuenta.letra, cuenta.punto, "99999999", CustomConvert.toStringWithTwoDecimalPlaces(resto), "N/D por Diferencia de Cambio")
            Case Else
                'ENTRA ACA SI ES IGUAL A CERO Y NO SE DEBE HACER NADA'
        End Select

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
        txtParidad.Text = orden.paridad
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

    Private Sub txtProveedor_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProveedor.KeyDown
        If e.KeyValue = Keys.Enter Then
            Dim proveedor = DAOProveedor.buscarProveedorPorCodigo(ceros(txtProveedor.Text, 11))
            If Not IsNothing(proveedor) Then
                mostrarProveedor(proveedor)
            Else
                txtRazonSocial.Text = ""
                MessageBox.Show("El proveedor ingresado es inexistente")
                txtProveedor.Focus()
            End If
        End If
    End Sub

    'Private Sub txtProveedor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProveedor.Leave
    '    Dim proveedor = DAOProveedor.buscarProveedorPorCodigo(ceros(txtProveedor.Text, 11))
    '    If Not IsNothing(proveedor) Then
    '        mostrarProveedor(proveedor)
    '    Else
    '        txtRazonSocial.Text = ""
    '        MessageBox.Show("El proveedor ingresado es inexistente")
    '    End If
    'End Sub

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

    Private Sub lstSeleccion_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSeleccion.Click
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

    Private Sub lstConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstConsulta.Click
        queryController.showMethod.Invoke(lstConsulta.SelectedValue)
        If queryController.text = "Proveedores" Then
            lstConsulta.Visible = False
        End If
        txtConsulta.Visible = False
        txtConsulta.Text = ""
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        txtIBCiudad.Text = "0,00"
        txtIngresosBrutos.Text = "0,00"
        txtGanancias.Text = "0,00"
        txtIVA.Text = "0,00"
        txtFecha.Text = Date.Today.ToShortDateString
        txtFechaParidad.Text = Date.Today.ToShortDateString
        gridPagos.Rows.Clear()
        pagos.Clear()
        gridFormaPagos.Rows.Clear()
        cheques.Clear()
        lstSeleccion.Visible = False
        lstConsulta.Visible = False
        txtConsulta.Visible = False
        traerParidad(txtFechaParidad.Text)

    End Sub

    Private Sub traerParidad(ByVal fecha As String)
        txtParidad.Text = SQLConnector.executeProcedureWithReturnValue("get_paridad", fecha).ToString()
    End Sub

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If validarDatos() Then
            Dim siguienteNumero As Integer = DAOPagos.siguienteNumeroDeOrden()
            bancoOrden = DAOBanco.buscarBancoPorCodigo(txtBanco.Text)
            proveedorOrden = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
            txtOrdenPago.Text = siguienteNumero
            Dim pago As OrdenPago = New OrdenPago(siguienteNumero, tipoOrden, CustomConvert.toDoubleOrZero(txtParidad.Text),
                                                  CustomConvert.toDoubleOrZero(txtTotal.Text), CustomConvert.toDoubleOrZero(txtIVA.Text),
                                                  CustomConvert.toDoubleOrZero(txtIngresosBrutos.Text), CustomConvert.toDoubleOrZero(txtIBCiudad.Text),
                                                  CustomConvert.toDoubleOrZero(txtGanancias.Text), txtFecha.Text, txtFechaParidad.Text, txtObservaciones.Text,
                                                  bancoOrden, proveedorOrden)
            crearNotasCreditoDebito()
            pago.pagos = crearPagos()
            pago.formaPagos = crearFormaPagos()
            DAOPagos.agregarPago(pago)
            MsgBox("El número de orden asignado es: " & siguienteNumero)
            btnLimpiar.PerformClick()
        End If
    End Sub

    Private Sub crearNotasCreditoDebito()
        Dim pagos As New List(Of Pago)
        Dim ultimoNumero As String = 0
        Dim tipoDoc As String = ""
        Dim neto As Double

        For Each row As DataGridViewRow In gridPagos.Rows
            If (Not row.IsNewRow And (Convert.ToString(row.Cells(3).Value) = "99999999")) Then

                If Convert.ToString(row.Cells(0).Value) = "02" Then
                    tipoDoc = "ND"
                Else
                    tipoDoc = "NC"
                End If

                neto = CustomConvert.toStringWithTwoDecimalPlaces(row.Cells(4).Value) / 1.21

                Dim interno As Integer = DAOCompras.siguienteNumeroDeInterno()
                Dim compra As New Compra(
                                         interno,
                                         DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text),
                                         Convert.ToString(row.Cells(0).Value),
                                         tipoDoc,
                                         "1",
                                         "2",
                                         Convert.ToString(row.Cells(1).Value),
                                         Convert.ToString(row.Cells(2).Value),
                                         ultimoNumero,
                                         txtFecha.Text,
                                         txtFecha.Text,
                                         txtFecha.Text,
                                         txtFecha.Text,
                                         CustomConvert.toStringWithTwoDecimalPlaces(txtParidad.Text),
                                         neto,
                                         CustomConvert.toStringWithTwoDecimalPlaces(row.Cells(4).Value) - neto,
                                         0,
                                         0,
                                         0,
                                         0,
                                         0,
                                         CustomConvert.toStringWithTwoDecimalPlaces(row.Cells(4).Value),
                                         0,
                                         "",
                                         "")
                crearImputaciones(compra)
                DAOCompras.agregarCompra(compra)
                DAOCompras.agregarDatosCuentaCorriente(compra)

            End If
            ultimoNumero = Convert.ToString(row.Cells(3).Value)
        Next
    End Sub

    Private Sub crearImputaciones(ByVal compra As Compra)
        Dim imputaciones As New List(Of Imputac)
        Dim debitoProv, debitoIva, debitoCuenta, creditoProv, creditoIva, creditoCuenta As Double
        Dim cuenta As Integer = 0
        debitoProv = 0
        debitoIva = debitoCuenta = creditoProv = creditoIva = creditoCuenta = debitoProv
        If compra.tipoDocumentoDescripcion = "ND" Then
            debitoCuenta = compra.neto
            debitoIva = compra.iva21
            creditoProv = compra.total
            cuenta = 6107
        Else
            creditoCuenta = compra.neto
            creditoIva = compra.iva21
            debitoProv = compra.total
            cuenta = 7308
        End If
        'For Each row As DataGridViewRow In gridAsientos.Rows
        'If Not row.IsNewRow Then

        imputaciones.Add(New Imputac(compra.fechaEmision, debitoProv, creditoProv, compra.proveedor.id.ToString, 2001, compra.nroInterno,
                                     compra.punto, compra.numero, compra.despacho, compra.letra, compra.tipoDocumento, "01"))
        imputaciones.Add(New Imputac(compra.fechaEmision, debitoIva, creditoIva, compra.proveedor.id.ToString, 151, compra.nroInterno,
                                     compra.punto, compra.numero, compra.despacho, compra.letra, compra.tipoDocumento, "02"))
        imputaciones.Add(New Imputac(compra.fechaEmision, debitoCuenta, creditoCuenta, compra.proveedor.id.ToString, cuenta, compra.nroInterno,
                             compra.punto, compra.numero, compra.despacho, compra.letra, compra.tipoDocumento, "03"))
        'End If
        'Next

        compra.agregarImputaciones(imputaciones)
    End Sub

    Private Function crearPagos()
        Dim pagos As New List(Of Pago)
        Dim ultimoNumero As String = ""

        For Each row As DataGridViewRow In gridPagos.Rows
            If Not row.IsNewRow Then
                If (Convert.ToString(row.Cells(3).Value) = "99999999") Then
                    pagos.Add(New Pago(Convert.ToString(row.Cells(0).Value), Convert.ToString(row.Cells(1).Value), Convert.ToString(row.Cells(2).Value), ultimoNumero,
                    Convert.ToString(row.Cells(5).Value), CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
                Else
                    pagos.Add(New Pago(Convert.ToString(row.Cells(0).Value), Convert.ToString(row.Cells(1).Value), Convert.ToString(row.Cells(2).Value), Convert.ToString(row.Cells(3).Value),
                    Convert.ToString(row.Cells(5).Value), CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
                End If

            End If
            ultimoNumero = Convert.ToString(row.Cells(3).Value)
        Next
        Return pagos
    End Function

    Private Function crearFormaPagos()
        Dim formaPagos As New List(Of FormaPago)
        For Each row As DataGridViewRow In gridFormaPagos.Rows
            If Not row.IsNewRow Then
                formaPagos.Add(New FormaPago(Convert.ToString(row.Cells(0).Value), CustomConvert.toIntOrZero(Convert.ToString(row.Cells(3).Value)), Convert.ToString(row.Cells(1).Value),
                                             Convert.ToString(row.Cells(2).Value), Convert.ToString(row.Cells(4).Value), CustomConvert.toDoubleOrZero(row.Cells(5).Value)))
            End If
        Next
        Return formaPagos
    End Function

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
                column = 4
            Case 2
                column = 1
            Case 3
                chequeRow = rowIndex
                lstSeleccion.SelectedIndex = 2
                lstSeleccion_Click(Nothing, Nothing)
                Exit Sub
            Case 5
                nombre = "US$"
                column = 4
            Case 6
                nombre = "Varios"
                column = 4
            Case Else
                Exit Sub
        End Select
        gridFormaPagos.CurrentCell.Value = ceros(val.ToString, 2)

        gridFormaPagos.Rows(rowIndex).Cells(4).Value = nombre
        gridFormaPagos.CurrentCell = gridFormaPagos.Rows(rowIndex).Cells(column)
    End Sub

    Private Sub gridFormaPagos_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormaPagos.CellLeave
        If e.ColumnIndex = 3 And e.RowIndex > -1 Then
            If gridFormaPagos.Rows(e.RowIndex).Cells(0).Value = "02" Then
                Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(gridFormaPagos.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                If Not IsNothing(banco) Then
                    gridFormaPagos.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value = banco.nombre
                End If
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
            gridPagos.Columns(5).ReadOnly = False
        End If
    End Sub

    Private Sub optAnticipos_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAnticipos.CheckedChanged
        If optAnticipos.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "0,00", txtRazonSocial.Text)
            gridPagos.Columns(5).ReadOnly = True
        End If
    End Sub

    Private Sub optVarios_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVarios.CheckedChanged
        If optVarios.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "", txtRazonSocial.Text)
            gridPagos.Columns(5).ReadOnly = True
        End If
    End Sub

    Private Sub optChequeRechazado_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChequeRechazado.CheckedChanged
        If optChequeRechazado.Checked Then
            gridPagos.Rows.Clear()
            gridPagos.Rows.Add("", "", "", "", "", "")
            gridPagos.Columns(5).ReadOnly = False
        End If
    End Sub

    Private Sub optCtaCte_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCtaCte.CheckedChanged
        If optCtaCte.Checked Then
            gridPagos.Rows.Clear()
            Try
                gridPagos.Columns(5).ReadOnly = True
            Catch ex As ArgumentOutOfRangeException
            End Try
        End If
    End Sub

    Private Sub txtFechaParidad_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFechaParidad.Leave
        traerParidad(txtFechaParidad.Text())
    End Sub

    Private Sub txtFechaParidad_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub lstSeleccion_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class