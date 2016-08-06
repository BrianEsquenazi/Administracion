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
        txtFecha.Text = Date.Today.ToShortDateString
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
End Class