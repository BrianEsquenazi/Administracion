﻿Imports ClasesCompartidas
Imports System.IO

Public Class ListadoProyeccionCobros

    Private Sub ListadoProyeccionCobros_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtAyuda.Text = ""
        txtDesdeProveedor.Text = "0"
        txtHastaProveedor.Text = "99999999999"
        txtFecha1.Text = "  /  /    "
        txtFecha2.Text = "  /  /    "
        txtFecha3.Text = "  /  /    "
        txtFecha4.Text = "  /  /    "
        opcPantalla.Checked = False
        opcImpesora.Checked = True
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
            txtFecha1.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtHastaProveedor.Text = ""
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtfecha1_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtFecha1.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFecha1.Text) = "S" Then
                txtFecha2.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFecha1.Text = "  /  /    "
        End If
    End Sub

    Private Sub txtfecha2_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtFecha2.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFecha2.Text) = "S" Then
                txtFecha3.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFecha2.Text = "  /  /    "
        End If
    End Sub

    Private Sub txtfecha3_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtFecha3.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFecha3.Text) = "S" Then
                txtFecha4.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFecha3.Text = "  /  /    "
        End If
    End Sub

    Private Sub txtfecha4_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtFecha4.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFecha4.Text) = "S" Then
                txtDesdeProveedor.Focus()
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFecha4.Text = "  /  /    "
        End If
    End Sub





    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Size = New System.Drawing.Size(460, 535)

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

        Dim txtEmpresa As String
        Dim txtFormula As String
        Dim x As Char = Chr(34)
        Dim txtfechaOrd1, txtfechaOrd2, txtfechaOrd3, txtfechaOrd4 As String
        Dim txtcompara As String
        Dim txtimpo1, txtimpo2, txtimpo3, txtimpo4, txtimpo5 As Double



        SQLConnector.retrieveDataTable("limpiar_impCtaCtePrvNet")

        txtEmpresa = "Surfactan S.A."

        txtfechaOrd1 = ordenaFecha(txtFecha1.Text)
        txtfechaOrd2 = ordenaFecha(txtFecha2.Text)
        txtfechaOrd3 = ordenaFecha(txtFecha3.Text)
        txtfechaOrd4 = ordenaFecha(txtFecha4.Text)


        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta_corriente_proveedores_desdehasta", txtDesdeProveedor.Text, txtHastaProveedor.Text, "P")

        For Each row As DataRow In tabla.Rows

            Dim CCPrv As New CtaCteProveedoresDeudaDesdeHasta(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString, row.Item(3).ToString, row.Item(4), row.Item(5), row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString, row.Item(10), row.Item(11).ToString, row.Item(12).ToString)

            txtcompara = ordenaFecha(CCPrv.vencimiento)
            txtimpo1 = 0
            txtimpo2 = 0
            txtimpo3 = 0
            txtimpo4 = 0
            txtimpo5 = 0

            If txtcompara <= txtfechaOrd1 Then
                txtimpo1 = txtimpo1 + CCPrv.saldo
            Else
                If txtcompara <= txtfechaOrd2 Then
                    txtimpo2 = txtimpo2 + CCPrv.saldo
                Else
                    If txtcompara <= txtfechaOrd3 Then
                        txtimpo3 = txtimpo3 + CCPrv.saldo
                    Else
                        If txtcompara <= txtfechaOrd4 Then
                            txtimpo4 = txtimpo4 + CCPrv.saldo
                        Else
                            txtimpo5 = txtimpo5 + CCPrv.saldo
                        End If
                    End If
                End If
            End If

            SQLConnector.executeProcedure("alta_impCtaCtePrvNet", CCPrv.Clave, CCPrv.Proveedor, CCPrv.Tipo, CCPrv.letra, CCPrv.punto, CCPrv.numero, CCPrv.total, CCPrv.saldo, CCPrv.fecha, CCPrv.vencimiento, CCPrv.VencimientoII, CCPrv.Impre, CCPrv.nroInterno, txtEmpresa, 0, 0, txtFecha1.Text, txtFecha2.Text, txtFecha3.Text, txtFecha4.Text, txtimpo1, txtimpo2, txtimpo3, txtimpo4, txtimpo5)

        Next










        txtUno = "{ImpCtaCtePrvNet.Proveedor} in " + x + "0" + x + " to " + x + "99999999999" + x
        txtFormula = txtUno

        Dim viewer As New ReportViewer("Proyeccion de Cobros de Corriente de Proveedres", Globals.reportPathWithName("wProyccprvnet.rpt"), txtFormula)

        If opcPantalla.Checked = True Then
            viewer.Show()
        Else
            viewer.imprimirReporte()
        End If



    End Sub
End Class