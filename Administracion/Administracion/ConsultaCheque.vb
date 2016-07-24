Imports ClasesCompartidas

Public Class ConsultaCheque

    Private Sub btnProceso_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProceso.Click
        If cmbTipo.SelectedIndex = 0 Then
            For Each row In SQLConnector.retrieveDataTable("get_carga_cheques_terceros", txtCheque.Text).Rows
                gridCheque.Rows.Add(row("Numero2").ToString,
                                    ("Banco2").ToString,
                                    row("Importe2").ToString,
                                    row("Fecha").ToString,
                                    row("Fecha2").ToString,
                                    row("Recibo").ToString,
                                    row("Cliente").ToString)
            Next
        Else
            For Each row In SQLConnector.retrieveDataTable("get_carga_cheques_propios", txtCheque.Text).Rows
                gridCheque.Rows.Add(row("Numero2").ToString,
                                    ("Banco2").ToString,
                                    row("Importe2").ToString,
                                    row("Fecha").ToString,
                                    row("Fecha2").ToString,
                                    row("Recibo").ToString,
                                    row("Proveedor").ToString)
            Next
        End If

    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub
End Class