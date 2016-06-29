Public Class ListadoIvaCompras

    

    Private Sub ListadoIvaCompras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtDesdeFecha.Text = "  /  /    "
        txthastafecha.Text = "  /  /    "

        TipoListado.Items.Clear()

        TipoListado.Items.Clear()
        TipoListado.Items.Add("C/Apertura")
        TipoListado.Items.Add("S/Apertura")

        TipoListado.SelectedIndex = 0

    End Sub

    Private Sub txtdesdefecha_KeyPress(ByVal sender As Object, _
                ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                Handles txtdesdefecha.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            txthastafecha.Focus()
            REM lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            REM txtAyuda.Text = ""
        End If
    End Sub

  
    Private Sub ListadoIvaCompras_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Hide()
        MenuPrincipal.Show()
    End Sub
    
    
    
End Class