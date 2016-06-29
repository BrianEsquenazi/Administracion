Imports ClasesCompartidas
Imports System.IO

Public Class ListadoSaldosCuentaCorrienteProveedores

    Dim WParametro As String

    Private Sub LitadoSaldosCuentaCorrienteProveedores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtAyuda.Text = ""
        txtDesdeProveedor.Text = ""
        txtHastaProveedor.Text = ""
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
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Hide()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        Me.Size = New System.Drawing.Size(460, 400)

        txtAyuda.Text = ""
        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")


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

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        WParametro = ""
        WParametro = lstAyuda.SelectedItem.ToString

        REM Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorNombre(WParametro)
        REM txtDesdeProveedor.Text = CampoProveedor.razonSocial

    End Sub

End Class