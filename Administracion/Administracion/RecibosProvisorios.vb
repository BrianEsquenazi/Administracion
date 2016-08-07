Imports ClasesCompartidas

Public Class RecibosProvisorios

    Dim queryController As QueryController
    Dim commonEventsHandler As New CommonEventsHandler

    Private Sub RecibosProvisorios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        commonEventsHandler.setIndexTab(Me)
        lstSeleccion.Items.Add(New QueryController("Clientes", AddressOf DAOCliente.buscarClientePorNombre, AddressOf mostrarCliente))
        lstSeleccion.SelectedIndex = 0

        Dim gridBuilder As New GridBuilder(gridRecibos)
        gridBuilder.addTextColumn(0, "Tipo")
        gridBuilder.addNumericColumn(1, "Número/Cta")
        gridBuilder.addDateColumn(2, "Fecha")
        gridBuilder.addTextColumn(3, "Banco")
        gridBuilder.addStrictlyPositiveFloatColumn(4, "Importe")
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
        gridRecibos.Rows.Clear()
    End Sub

    Private Sub eventoSegunTipoEnFormaDePagoPara(ByVal val As Integer, ByVal rowIndex As Integer, ByVal columnIndex As Integer)
        Dim column As Integer = columnIndex
        Select Case val
            Case 1
                column = 4
            Case 2
                column = 1
            Case Else
                Exit Sub
        End Select
        gridRecibos.CurrentCell.Value = ceros(val.ToString, 2)
        gridRecibos.CurrentCell = gridRecibos.Rows(rowIndex).Cells(column)
    End Sub

    Private Sub gridRecibos_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridRecibos.CellValueChanged
        sumarValores()
        agregarClienteABanco()
    End Sub

    Private Sub sumarValores()
        Dim total As Double = 0
        total += CustomConvert.toDoubleOrZero(txtRetGanancias.Text)
        total += CustomConvert.toDoubleOrZero(txtRetIB.Text)
        total += CustomConvert.toDoubleOrZero(txtRetIva.Text)
        total += CustomConvert.toDoubleOrZero(txtRetSuss.Text)
        For Each row As DataGridViewRow In gridRecibos.Rows
            total += CustomConvert.toDoubleOrZero(row.Cells(4).Value)
        Next
        lblTotal.Text = CustomConvert.toStringWithTwoDecimalPlaces(total)
    End Sub

    Private Sub gridRecibos_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridRecibos.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim iCol = gridRecibos.CurrentCell.ColumnIndex
            Dim iRow = gridRecibos.CurrentCell.RowIndex
            If iCol = 0 And iRow > -1 Then
                Dim val = gridRecibos.Rows(iRow).Cells(iCol).Value
                eventoSegunTipoEnFormaDePagoPara(CustomConvert.toIntOrZero(val), iRow, iCol)
            End If
        End If
    End Sub

    Private Sub agregarClienteABanco()
        For Each row As DataGridViewRow In gridRecibos.Rows
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
        mostrarRecibo(DAORecibo.buscarReciboProvisorio(txtRecibo.Text))
    End Sub

    Private Sub mostrarRecibo(ByVal reciboProvisorio As ReciboProvisorio)
        If IsNothing(reciboProvisorio) Then
            'Cleanner.cleanWithoutChangeFocus(Me)
            'setDefaults()
        Else
            txtRecibo.Text = reciboProvisorio.codigo
            txtFecha.Text = reciboProvisorio.fecha
            mostrarCliente(reciboProvisorio.cliente)
            txtRetGanancias.Text = reciboProvisorio.retGanancias
            txtRetIB.Text = reciboProvisorio.retIB
            txtRetIva.Text = reciboProvisorio.retIVA
            txtRetSuss.Text = reciboProvisorio.retSuss
            txtTotal.Text = reciboProvisorio.total
            txtParidad.Text = reciboProvisorio.paridad
            mostrarFormasPago(reciboProvisorio.formasPago)
        End If
    End Sub

    Private Sub mostrarFormasPago(ByVal formasPago As List(Of FormaPago))
        gridRecibos.Rows.Clear()
        For Each forma As FormaPago In formasPago
            gridRecibos.Rows.Add(forma.tipo, forma.numero, forma.fecha, forma.nombre, forma.importe)
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

    Private Sub txtCliente_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCliente.Leave
        mostrarCliente(DAOCliente.buscarClientePorCodigo(txtCliente.Text))
    End Sub

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(CustomConvert.toDoubleOrZero(lblTotal.Text) = CustomConvert.toDoubleOrZero(txtTotal.Text), "La suma de los importes de la tabla no coincide con lo informado en el total")

        If validador.flush Then
            Dim recibo As New ReciboProvisorio(txtRecibo.Text, txtFecha.Text, DAOCliente.buscarClientePorCodigo(txtCliente.Text),
                                               CustomConvert.toDoubleOrZero(txtRetGanancias.Text), CustomConvert.toDoubleOrZero(txtRetIB.Text),
                                               CustomConvert.toDoubleOrZero(txtRetIva.Text), CustomConvert.toDoubleOrZero(txtRetSuss.Text),
                                               CustomConvert.toDoubleOrZero(txtParidad.Text), CustomConvert.toDoubleOrZero(txtTotal.Text))
            recibo.formasPago = crearFormasPago()
            DAORecibo.agregarReciboProvisorio(recibo)
        End If
    End Sub

    Private Function crearFormasPago() As List(Of FormaPago)
        Dim formasPago As New List(Of FormaPago)
        For Each row As DataGridViewRow In gridRecibos.Rows
            If Not row.IsNewRow Then
                formasPago.Add(New FormaPago(row.Cells(0).Value, 0, asString(row.Cells(1).Value), asString(row.Cells(2).Value), asString(row.Cells(3).Value), CustomConvert.toDoubleOrZero(row.Cells(4).Value)))
            End If
        Next
        Return formasPago
    End Function

    Private Function asString(ByVal value)
        If IsNothing(value) Then
            Return ""
        Else
            Return value.ToString
        End If
    End Function

    Private Sub txtRetencion_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetGanancias.Leave, txtRetSuss.Leave, txtRetIva.Leave, txtRetIB.Leave
        sumarValores()
    End Sub
End Class