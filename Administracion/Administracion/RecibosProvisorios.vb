Imports ClasesCompartidas

Public Class RecibosProvisorios

    Dim commonEventsHandler As New CommonEventsHandler

    Private Sub RecibosProvisorios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        commonEventsHandler.setIndexTab(Me)

        Dim gridBuilder As New GridBuilder(gridRecibos)
        gridBuilder.addTextColumn(0, "Tipo")
        gridBuilder.addNumericColumn(1, "Número/Cta")
        gridBuilder.addDateColumn(2, "Fecha")
        gridBuilder.addTextColumn(3, "Banco")
        gridBuilder.addStrictlyPositiveFloatColumn(4, "Importe")
        btnLimpiar.PerformClick()
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
        mostrarRecibo(DAORecibo.buscarReciboProvisorio(txtRecibo.Text))
    End Sub

    Private Sub mostrarRecibo(ByVal reciboProvisorio As ReciboProvisorio)
        If IsNothing(reciboProvisorio) Then
            Cleanner.cleanWithoutChangeFocus(Me)
            setDefaults()
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
            'Dim recibo As New ReciboProvisorio(txtRecibo.Text, txtFecha.Text, DAOCliente.buscarClientePorCodigo(txtCliente.Text), 
            MsgBox("Agregaste")
        End If
    End Sub
End Class