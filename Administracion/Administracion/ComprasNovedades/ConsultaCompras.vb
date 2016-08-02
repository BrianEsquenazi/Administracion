Public Class ConsultaCompras

    Dim duenio As Compras
    Dim query As QueryFunction
    Dim showMethod As ShowMethod
    Dim onlyProveedores As Boolean

    Public Sub New(ByVal form As Form, Optional ByVal isForProveedores As Boolean = False)
        InitializeComponent()
        duenio = form
        onlyProveedores = isForProveedores
    End Sub

    Private Sub lstSeleccion_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSeleccion.Click
        If lstSeleccion.SelectedItem = "Proveedores" Then
            query = AddressOf DAOProveedor.buscarProveedorPorNombre
            showMethod = AddressOf duenio.mostrarProveedor
        Else
            query = AddressOf DAOCuentaContable.buscarCuentaContablePorDescripcion
            showMethod = AddressOf duenio.mostrarCuentaContable
        End If
        lstConsulta.DataSource = query.Invoke("")
        txtConsulta.Visible = True
        lstConsulta.Visible = True
        lstSeleccion.Visible = False
        Me.Size = New System.Drawing.Size(400, 300)
        txtConsulta.Focus()
    End Sub

    Private Sub ConsultaCompras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Size = New System.Drawing.Size(200, 100)
        If onlyProveedores Then
            lstSeleccion.SelectedItem = "Proveedores"
            lstSeleccion_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub txtConsulta_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsulta.KeyDown
        If e.KeyValue = Keys.Enter Then
            lstConsulta.DataSource = query.Invoke(txtConsulta.Text)
        End If
    End Sub

    Private Sub lstConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstConsulta.Click
        showMethod.Invoke(lstConsulta.SelectedValue)
        Close()
    End Sub
End Class