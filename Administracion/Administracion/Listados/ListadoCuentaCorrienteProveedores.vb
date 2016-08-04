﻿Imports ClasesCompartidas
Imports System.IO

Public Class ListadoCuentaCorrienteProveedores

    Private Sub ListadoCuentaCorrienteProveedores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDesdeProveedor.Text = "0"
        txtHastaProveedor.Text = "99999999999"
        opcPantalla.Checked = False
        opcImpesora.Checked = True
        opcPendiente.Checked = True
        opcCompleto.Checked = False
    End Sub

    Private Sub txtdesdeproveedor_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtDesdeProveedor.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtHastaProveedor.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeProveedor.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txthastaproveedor_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtHastaProveedor.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtDesdeProveedor.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtHastaProveedor.Text = ""
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

        Me.Size = New System.Drawing.Size(550, 510)

        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")

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
            lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtAyuda.Text = ""
        End If
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        txtDesdeProveedor.Text = proveedor.id
        txtHastaProveedor.Text = proveedor.id
        txtDesdeProveedor.Focus()
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarProveedor(lstAyuda.SelectedValue)
    End Sub

    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click

        Dim txtUno As String

        Dim txtFormula As String
        Dim x As Char = Chr(34)
        Dim WSuma As Double
        Dim WOrden As Integer
        Dim txtCorte As String = ""
        Dim txtLLave As Integer = 0
        Dim txtEmpresa As String

        SQLConnector.retrieveDataTable("limpiar_impCtaCtePrvNet")


        REM Reviso el cual esta checkeado asi le pongo los valores a Tipo
        Dim WTipo As Char
        WTipo = "T"
        If (opcPendiente.Checked) Then
            WTipo = "P"
        End If

        txtEmpresa = "Surfactan S.A."

        REM dada fix CAMBIAR Al uso de dao!!
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta_corriente_proveedores_desdehasta", txtDesdeProveedor.Text, txtHastaProveedor.Text, WTipo)

        For Each row As DataRow In tabla.Rows

            Dim CCPrv As New CtaCteProveedoresDeudaDesdeHasta(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString, row.Item(3).ToString, row.Item(4), row.Item(5), row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString, row.Item(10), row.Item(11).ToString, row.Item(12).ToString)

            If txtLLave = 0 Then
                txtLLave = 1
                txtCorte = CCPrv.Proveedor
                WSuma = 0
                WOrden = 0
            End If

            If txtCorte <> CCPrv.Proveedor Then
                txtCorte = CCPrv.Proveedor
                WSuma = 0
                WOrden = 0
            End If

            WSuma = WSuma + CCPrv.saldo
            WOrden = WOrden + 1
            SQLConnector.executeProcedure("alta_impCtaCtePrvNet", CCPrv.Clave, CCPrv.Proveedor, CCPrv.Tipo, CCPrv.letra, CCPrv.punto, CCPrv.numero, CCPrv.total, CCPrv.saldo, CCPrv.fecha, CCPrv.vencimiento, CCPrv.VencimientoII, CCPrv.Impre, CCPrv.nroInterno, txtEmpresa, WSuma, WOrden, "", "", "", "", 0, 0, 0, 0, 0)

        Next

        txtUno = "{ImpCtaCtePrvNet.Proveedor} in " + x + "0" + x + " to " + x + "99999999999" + x
        txtFormula = txtUno

        Dim viewer As New ReportViewer("Listado de Corriente de Proveedres", "c:\Crystal\wccprvnet.rpt", txtFormula)

        If opcPantalla.Checked = True Then
            viewer.Show()
        Else
            viewer.imprimirReporte()
        End If





        REM borrar   impctacteprv




        'Dim tabla As DataTable
        'tabla = SQLConnector.retrieveDataTable("procesoReteIb", ordDesde, ordHasta)

        'For Each row As DataRow In tabla.Rows

        '    Dim CamposReteIb As New ProcesoReteIb(row.Item(0).ToString, row.Item(1), row.Item(2).ToString, row.Item(3).ToString, row.Item(4).ToString, row.Item(5).ToString, row.Item(6).ToString)

        '    WCuit = sacaguiones(CamposReteIb.cuit)
        '    WFecha = CamposReteIb.fecha
        '    WSucursal = "0001"
        '    WOrden = ceros(CamposReteIb.orden, 8)
        '    WImporte = ceros(formatonumerico(redondeo(CamposReteIb.retotra), "########0.#0", "."), 11)

        'Next







    End Sub
End Class