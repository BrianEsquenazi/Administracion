Imports ClasesCompartidas

Public Class Recibos

    Dim queryController As QueryController
    Dim commonEventsHandler As New CommonEventsHandler

    Private Sub Recibos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        commonEventsHandler.setIndexTab(Me)
        lstSeleccion.Items.Add(New QueryController("Clientes", AddressOf DAOCliente.buscarClientePorNombre, AddressOf mostrarCliente))
        lstSeleccion.Items.Add(New QueryController("Cuentas Contables", AddressOf DAOCuentaContable.buscarCuentaContablePorDescripcion, AddressOf mostrarCuenta))
        lstSeleccion.SelectedIndex = 0

        Dim gridBuilder As New GridBuilder(gridFormasPago)
        gridBuilder.addTextColumn(0, "Tipo")
        gridBuilder.addNumericColumn(1, "Número/Cta")
        gridBuilder.addDateColumn(2, "Fecha")
        gridBuilder.addTextColumn(3, "Banco")
        gridBuilder.addStrictlyPositiveFloatColumn(4, "Importe")

        Dim gridBuilder2 As New GridBuilder(gridPagos)
        gridBuilder2.addTextColumn(0, "Tipo")
        gridBuilder2.addTextColumn(1, "Letra")
        gridBuilder2.addNumericColumn(2, "Punto")
        gridBuilder2.addNumericColumn(3, "Número")
        gridBuilder2.addStrictlyPositiveFloatColumn(4, "Importe")

        btnLimpiar.PerformClick()
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
        lstConsulta.Visible = False
        txtConsulta.Visible = False
        txtConsulta.Text = ""
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        setDefaults()
    End Sub

    Private Sub setDefaults()
        txtFecha.Text = Date.Today.ToShortDateString
        gridFormasPago.Rows.Clear()
        gridPagos.Rows.Clear()
    End Sub

    Private Sub eventoSegunTipoEnFormaDePagoPara(ByVal val As Integer, ByVal rowIndex As Integer, ByVal columnIndex As Integer)
        Dim column As Integer = columnIndex
        Select Case val
            Case 1, 4
                column = 4
            Case 2
                column = 1
            Case Else
                Exit Sub
        End Select
        gridFormasPago.CurrentCell.Value = ceros(val.ToString, 2)
        gridFormasPago.CurrentCell = gridFormasPago.Rows(rowIndex).Cells(column)
    End Sub

    Private Sub gridFormasPago_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFormasPago.CellValueChanged
        sumarValores()
        agregarClienteABanco()
    End Sub

    Private Sub sumarValores()
        Dim total As Double = 0
        total += CustomConvert.toDoubleOrZero(txtRetGanancias.Text)
        total += CustomConvert.toDoubleOrZero(txtRetIB.Text)
        total += CustomConvert.toDoubleOrZero(txtRetIva.Text)
        total += CustomConvert.toDoubleOrZero(txtRetSuss.Text)
        For Each row As DataGridViewRow In gridFormasPago.Rows
            total += CustomConvert.toDoubleOrZero(row.Cells(4).Value)
        Next
        lblTotalFormasPago.Text = CustomConvert.toStringWithTwoDecimalPlaces(total)
    End Sub

    Private Sub gridFormasPago_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridFormasPago.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim iCol = gridFormasPago.CurrentCell.ColumnIndex
            Dim iRow = gridFormasPago.CurrentCell.RowIndex
            If iCol = 0 And iRow > -1 Then
                Dim val = gridFormasPago.Rows(iRow).Cells(iCol).Value
                eventoSegunTipoEnFormaDePagoPara(CustomConvert.toIntOrZero(val), iRow, iCol)
            End If
        End If
    End Sub

    Private Sub agregarClienteABanco()
        For Each row As DataGridViewRow In gridFormasPago.Rows
            If row.Cells(3).Value <> "" Then
                If row.Cells(3).Value.ToString.Length > 20 Or row.Cells(3).Value.ToString.Contains("/") Then
                    'row.Cells(3).Value = ""
                Else
                    row.Cells(3).Value = row.Cells(3).Value.ToString & clienteSinCeros()
                End If
            End If
        Next
    End Sub

    Private Function clienteSinCeros()
        Try
            Dim cliente As String = txtCliente.Text
            Dim numero As Integer = CustomConvert.toIntOrZero(cliente.Substring(1, cliente.Count - 1))
            Return "/" & cliente.First & numero
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtRecibo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRecibo.Leave
        txtRecibo.Text = ceros(txtRecibo.Text, 6)
        mostrarRecibo(DAORecibo.buscarRecibo(txtRecibo.Text))
    End Sub

    Private Sub mostrarRecibo(ByVal recibo As Recibo)
        If IsNothing(recibo) Then
            'Cleanner.cleanWithoutChangeFocus(Me)
            'setDefaults()
        Else
            txtRecibo.Text = recibo.codigo
            txtFecha.Text = recibo.fecha
            mostrarCliente(recibo.cliente)
            txtRetGanancias.Text = recibo.retGanancias
            txtRetIB.Text = recibo.retIB
            txtRetIva.Text = recibo.retIVA
            txtRetSuss.Text = recibo.retSuss
            txtTotal.Text = recibo.total
            txtParidad.Text = recibo.paridad
            mostrarPagos(recibo.pagos)
            mostrarFormasPago(recibo.formasPago)
        End If
    End Sub

    Private Sub mostrarReciboProvisorio(ByVal recibo As ReciboProvisorio)
        If Not IsNothing(recibo) Then
            txtRecibo.Text = recibo.codigo
            txtFecha.Text = recibo.fecha
            mostrarCliente(recibo.cliente)
            txtRetGanancias.Text = recibo.retGanancias
            txtRetIB.Text = recibo.retIB
            txtRetIva.Text = recibo.retIVA
            txtRetSuss.Text = recibo.retSuss
            txtTotal.Text = recibo.total
            txtParidad.Text = recibo.paridad
            mostrarFormasPago(recibo.formasPago)
        End If
    End Sub

    Private Sub mostrarPagos(ByVal pagos As List(Of Pago))
        gridPagos.Rows.Clear()
        For Each pago As Pago In pagos
            gridPagos.Rows.Add(pago.tipo, pago.letra, pago.punto, pago.numero, pago.importe)
        Next
    End Sub

    Private Sub mostrarFormasPago(ByVal formasPago As List(Of FormaPago))
        gridFormasPago.Rows.Clear()
        For Each forma As FormaPago In formasPago
            gridFormasPago.Rows.Add(forma.tipo, forma.numero, forma.fecha, forma.nombre, forma.importe)
        Next
    End Sub

    Private Sub mostrarCliente(ByVal cliente As Cliente)
        If IsNothing(cliente) Then
            txtNombre.Text = ""
        Else
            txtCliente.Text = cliente.id
            txtNombre.Text = cliente.razon
        End If
    End Sub

    Private Sub mostrarCuenta(ByVal cuenta As CuentaContable)
        If IsNothing(cuenta) Then
            txtNombreCuenta.Text = ""
        Else
            txtCuenta.Text = cuenta.id
            txtNombreCuenta.Text = cuenta.descripcion
        End If
    End Sub

    Private Sub txtCliente_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCliente.Leave
        mostrarCliente(DAOCliente.buscarClientePorCodigo(txtCliente.Text))
    End Sub

    Private Sub txtProvi_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProvi.Leave
        mostrarReciboProvisorio(DAORecibo.buscarReciboProvisorio(txtProvi.Text))
    End Sub
End Class