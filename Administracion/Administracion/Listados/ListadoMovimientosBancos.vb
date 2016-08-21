Imports ClasesCompartidas
Imports System.IO

Public Class ListadoMovimientosBancos

    Dim txtVectorBanco(1000) As String

    Private Sub ListadoMovimientosBancos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtDesdeFecha.Text = "  /  /    "
        txthastafecha.Text = "  /  /    "

        txtDesdeBanco.Text = "0"
        txtHastaBanco.Text = "9999"

        opcPantalla.Checked = False
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
                txtDesdeBanco.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txthastafecha.Text = "  /  /    "
            Me.txthastafecha.SelectionStart = 0
        End If
    End Sub

    Private Sub txtdesdebanco_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtDesdeBanco.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtHastaBanco.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeBanco.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txthastabanco_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtHastaBanco.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtDesdeFecha.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtHastaBanco.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub


    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Size = New System.Drawing.Size(470, 490)

        lstAyuda.DataSource = DAOBanco.buscarBancoPorNombre("")

        txtAyuda.Text = ""
        txtAyuda.Visible = True
        lstAyuda.Visible = True

        txtAyuda.Focus()

    End Sub

    Private Sub txtAyuda_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtAyuda.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            lstAyuda.DataSource = DAOBanco.buscarBancoPorNombre(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtAyuda.Text = ""
            lstAyuda.DataSource = DAOBanco.buscarBancoPorNombre(txtAyuda.Text)
        End If
    End Sub

    Private Sub mostrarbanco(ByVal banco As Banco)
        txtDesdeBanco.Text = banco.id
        txtHastaBanco.Text = banco.id
        txtDesdeBanco.Focus()
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarbanco(lstAyuda.SelectedValue)
        REM txtDesdeProveedor.Text = lstAyuda.SelectedValue.id
    End Sub


    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click

        Dim txtUno As String

        Dim txtEmpresa As String
        Dim txtFormula As String
        Dim x As Char = Chr(34)
        Dim txtDesdefechaOrd, txtHastafechaOrd

        Dim txtBancoCodigo As Integer
        Dim txtBancoCuenta As String

        Dim txtRenglon As Integer
        Dim txtTitulo As String
        Dim txtTituloList As String
        Dim txtVarios As String

        Dim txtDebito, txtCredito As Double
        Dim txtAcredita, txtAcreditaOrd As String


        SQLConnector.retrieveDataTable("limpiar_movban")

        txtEmpresa = "Surfactan S.A."

        txtDesdefechaOrd = ordenaFecha(txtDesdeFecha.Text)
        txtHastafechaOrd = ordenaFecha(txthastafecha.Text)

        Dim tablaII As DataTable
        tablaII = SQLConnector.retrieveDataTable("buscar_banco_por_nombre", "")

        For Each row As DataRow In tablaII.Rows

            Dim CamposBanco As New LeeBanco(row.Item(0), row.Item(1), row.Item(2))

            txtBancoCuenta = CamposBanco.Cuenta
            txtBancoCodigo = CamposBanco.banco

            txtVectorBanco(txtBancoCodigo) = txtBancoCuenta

        Next


        txtRenglon = 0





        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_pagos_Movban", txtDesdefechaOrd, txtHastafechaOrd, txtDesdeBanco.Text, txtHastaBanco.Text)

        For Each row As DataRow In tabla.Rows

            Dim CampoPagos As New LeePagosMovBan(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString,
                                            row.Item(3), row.Item(4).ToString, row.Item(5).ToString,
                                            row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString,
                                            row.Item(9), row.Item(10), row.Item(11), row.Item(12),
                                            row.Item(13), row.Item(14))



            txtRenglon = txtRenglon + 1

            txtTitulo = "Pagos"
            txtEmpresa = 1
            txtTituloList = "Surfactan S.A."
            txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

            txtAcredita = CampoPagos.fecha2
            txtAcreditaOrd = CampoPagos.fechaord2
            txtDebito = 0
            txtCredito = CampoPagos.importe2

            SQLConnector.executeProcedure("alta_movban", txtRenglon, CampoPagos.banco2, CampoPagos.fecha, CampoPagos.fechaord, txtAcredita, txtAcreditaOrd, CampoPagos.observaciones,
                                          CampoPagos.numero2, txtDebito, txtCredito, CampoPagos.orden, txtEmpresa, txtTitulo, txtTituloList, CampoPagos.proveedor)


        Next






        tabla = SQLConnector.retrieveDataTable("buscar_depositos_Movban", txtDesdefechaOrd, txtHastafechaOrd, txtDesdeBanco.Text, txtHastaBanco.Text)

        For Each row As DataRow In tabla.Rows

            Dim CampoPagos As New LeePagosMovBan(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString,
                                            row.Item(3), row.Item(4).ToString, row.Item(5).ToString,
                                            row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString,
                                            row.Item(9), row.Item(10), row.Item(11), row.Item(12),
                                            row.Item(13), row.Item(14))



            txtRenglon = txtRenglon + 1

            txtTitulo = "Pagos"
            txtEmpresa = 1
            txtTituloList = "Surfactan S.A."
            txtVarios = "Desde el " + txtDesdeFecha.Text + " hasta el " + txthastafecha.Text

            txtAcredita = CampoPagos.fecha2
            txtAcreditaOrd = CampoPagos.fechaord2
            txtDebito = 0
            txtCredito = CampoPagos.importe2

            SQLConnector.executeProcedure("alta_movban", txtRenglon, CampoPagos.banco2, CampoPagos.fecha, CampoPagos.fechaord, txtAcredita, txtAcreditaOrd, CampoPagos.observaciones,
                                          CampoPagos.numero2, txtDebito, txtCredito, CampoPagos.orden, txtEmpresa, txtTitulo, txtTituloList, CampoPagos.proveedor)


        Next





        'Dim txtdada As Double
        'txtdada = SQLConnector.executeProcedureWithReturnValue("get_saldo_inicial_pagos", txtDesdefechaOrd, txtHastafechaOrd, txtDesdeBanco.Text, txtHastaBanco.Text)



        txtUno = "{Movban.Banco} in " + txtDesdeBanco.Text + " to " + txtHastaBanco.Text
        txtFormula = txtUno

        Dim viewer As New ReportViewer("Listado de Movimientos Bancarios", Globals.reportPathWithName("wMovbannet.rpt"), txtFormula)

        If opcPantalla.Checked = True Then
            viewer.Show()
        Else
            viewer.imprimirReporte()
        End If

    End Sub
End Class