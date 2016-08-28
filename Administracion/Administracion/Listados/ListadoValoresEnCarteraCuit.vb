Imports ClasesCompartidas
Imports System.IO

Public Class ListadoValoresEnCarteraCuit

    Private Sub ListadoValoresEnCarteraCuit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtDesdeFecha.Text = "  /  /    "
        txthastafecha.Text = "  /  /    "

        opcImpesora.Checked = True

    End Sub
    Private Sub txtdesdefecha_KeyPress(ByVal sender As Object, _
               ByVal e As System.Windows.Forms.KeyPressEventArgs) _
               Handles txtDesdeFecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtDesdeFecha.Text) = "S" Then
                txthastafecha.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeFecha.Text = "  /  /    "
            Me.txtDesdeFecha.SelectionStart = 0
        End If
    End Sub

    Private Sub txthastafecha_KeyPress(ByVal sender As Object, _
                ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                Handles txthastafecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txthastafecha.Text) = "S" Then
                txtDesdeFecha.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txthastafecha.Text = "  /  /    "
            Me.txthastafecha.SelectionStart = 0
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click


        'Dim varUno As String

        'Dim varEmpresa As String
        'Dim varFormula As String
        'Dim x As Char = Chr(34)
        'Dim varDesdefechaOrd, varHastafechaOrd As String
        'Dim varDesdeCliente, varHastaCliente As String

        'SQLConnector.retrieveDataTable("limpiar_valcar")

        'varEmpresa = "Surfactan S.A."

        'varDesdefechaOrd = ordenaFecha(txtDesdeFecha.Text)
        'varHastafechaOrd = ordenaFecha(txthastafecha.Text)

        'If LTrim(RTrim(txtCliente.Text)) <> "" Then
        '    varDesdeCliente = txtCliente.Text
        '    varHastaCliente = txtCliente.Text
        'Else
        '    varDesdeCliente = ""
        '    varHastaCliente = "Z99999"
        'End If

        'Dim tabla As DataTable

        'tabla = SQLConnector.retrieveDataTable("buscar_cheques_valcar", varDesdefechaOrd, varHastafechaOrd, varDesdeCliente, varHastaCliente)

        'For Each row As DataRow In tabla.Rows

        '    Dim CampoRecibos As New LeeRecibosValcar(row.Item(0), row.Item(1), row.Item(2),
        '                                   row.Item(3), row.Item(4), row.Item(5),
        '                                   row.Item(6), row.Item(7), row.Item(8),
        '                                   row.Item(9), row.Item(10), row.Item(11), row.Item(12),
        '                                   row.Item(13), row.Item(14), row.Item(15), row.Item(16), row.Item(17),
        '                                   row.Item(18), row.Item(19), row.Item(20), row.Item(21), row.Item(22), row.Item(23))

        '    varEmpresa = 1

        '    SQLConnector.executeProcedure("alta_valcar", CampoRecibos.recibo, CampoRecibos.cliente, CampoRecibos.numero2, CampoRecibos.Banco2, varImpo1, varImpo2, varImpo3,
        '                                  varImpo4, varImpo5, vartitulo1, vartitulo2, vartitulo3, vartitulo4, vartitulo5)

        'Next





        ''Dim txtdada As Double
        ''txtdada = SQLConnector.executeProcedureWithReturnValue("get_saldo_inicial_pagos", txtDesdefechaOrd, txtHastafechaOrd, txtDesdeBanco.Text, txtHastaBanco.Text)



        'varUno = "{Valcar.recibo} in 0 to 999999"
        'varFormula = varUno

        'Dim viewer As New ReportViewer("Listado de Valores en Cartera", Globals.reportPathWithName("wvalcarnet.rpt"), varFormula)

        'If opcPantalla.Checked = True Then
        '    viewer.Show()
        'Else
        '    viewer.imprimirReporte()
        'End If






    End Sub
End Class