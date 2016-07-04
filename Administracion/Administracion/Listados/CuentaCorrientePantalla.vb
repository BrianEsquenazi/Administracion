Imports ClasesCompartidas
Imports System.IO

Public Class CuentaCorrientePantalla

    Dim dataGridBuilder As GridBuilder
    Dim aa As String

    'Private Sub txtproveedor_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        Dim CampoProveedor As Banco = DAOProveedor.buscarProveedorPorCodigo("1")
    '        ProveedorRazon.Text = CampoProveedor.nombre
    '    End If
    'End Sub


    Private Sub txtproveedor_KeyPress(ByVal sender As Object, _
                    ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                    Handles txtProveedor.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
            txtRazon.Text = CampoProveedor.razonSocial
            Call Proceso()
            txtRazon.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtRazon.Focus()
        End If
    End Sub

    Private Sub CuentaCorrientePantalla_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dataGridBuilder = New GridBuilder(GRilla)

        dataGridBuilder.addTextColumn(0, "Tipo")
        dataGridBuilder.addTextColumn(1, "Letra")
        dataGridBuilder.addDateColumn(2, "Punto")
        dataGridBuilder.addTextColumn(3, "Numero")
        dataGridBuilder.addTextColumn(4, "Importe")
        dataGridBuilder.addTextColumn(5, "Saldo")
        dataGridBuilder.addTextColumn(6, "Fecha")
        dataGridBuilder.addTextColumn(7, "Vencimiento")

    End Sub

    Private Sub Proceso()
        'DAOCtaCteProveedor.buscardeuda(txtProveedor.Text).ForEach(Sub(ctacteprv) GRilla.Rows.Add(ctacteprv.Tipo, ctacteprv.letra, ctacteprv.punto, ctacteprv.numero, ctacteprv.total, ctacteprv.saldo, ctacteprv.fecha, ctacteprv.vencimiento))

        GRilla.AllowUserToAddRows = False
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        boxPantallaProveedores.Visible = True
        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")

        txtAyuda.Text = ""
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
        txtProveedor.Text = proveedor.id
        txtRazon.Text = proveedor.razonSocial
        boxPantallaProveedores.Visible = False
        Call Proceso()
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarProveedor(lstAyuda.SelectedValue)
    End Sub
  
    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

End Class