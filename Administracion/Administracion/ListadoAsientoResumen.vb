﻿Imports ClasesCompartidas
Imports System.IO

Public Class ListadoAsientoResumen

    Dim WParametro As String

    Private Sub ListadoAsientoResumen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtDesdeFecha.Text = "  /  /    "
        txthastafecha.Text = "  /  /    "

        txtDesdeCuenta.Text = ""
        txtHastaCuenta.Text = ""


        TipoListado.Items.Clear()
        TipoListado.Items.Add("Completo")
        TipoListado.Items.Add("Resumido")
        TipoListado.SelectedIndex = 0

        opcPantalla.Checked = False
        opcImpesora.Checked = True

    End Sub

    Private Sub txtdesdefecha_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtDesdeFecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txthastafecha.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeFecha.Text = "  /  /    "
        End If
    End Sub

    Private Sub txthastafecha_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txthastafecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtDesdeCuenta.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txthastafecha.Text = "  /  /    "
        End If
    End Sub

    Private Sub txtdesdecuenta_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtDesdeCuenta.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtHastaCuenta.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtDesdeCuenta.Text = ""
        End If
    End Sub

    Private Sub txthastacuenta_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtHastaCuenta.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txtDesdeFecha.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtHastaCuenta.Text = ""
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Hide()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Size = New System.Drawing.Size(500, 520)

        txtAyuda.Text = ""
        lstAyuda.DataSource = DAOCuentaContable.buscarCuentaContablePorDescripcion("")


        txtAyuda.Visible = True
        lstAyuda.Visible = True

        txtAyuda.Focus()

    End Sub

    Private Sub txtAyuda_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtAyuda.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            lstAyuda.DataSource = DAOCuentaContable.buscarCuentaContablePorDescripcion(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtAyuda.Text = ""
        End If
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        WParametro = ""
        WParametro = lstAyuda.SelectedItem.ToString

        REM Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorNombre(WParametro)
        REM txtDesdeProveedor.Text = CampoProveedor.razonSocial

    End Sub

End Class