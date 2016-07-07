Imports ClasesCompartidas

Public Class Pagos

    Dim queryController As QueryController

    Private Sub Pagos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbTipo.SelectedIndex = 0
        lstSeleccion.Items.Add(New QueryController("Proveedores", AddressOf DAOProveedor.buscarProveedorPorNombre, AddressOf mostrarProveedor))
        lstSeleccion.Items.Add(New QueryController("Cuentas Corrientes", AddressOf cuentasCorrientesDelProveedorActualSegunDescripcion, AddressOf mostrarProveedor))
        lstSeleccion.Items.Add(New QueryController("Cheques Terceros", AddressOf DAODeposito.buscarCheques, AddressOf mostrarCheque))
        lstSeleccion.Items.Add(New QueryController("Cuentas Contables", AddressOf DAOCuentaContable.buscarCuentaContablePorDescripcion, AddressOf mostrarCuentaContable))
        lstSeleccion.SelectedIndex = 0
        Dim commonEventHandler As New CommonEventsHandler
        commonEventHandler.setIndexTab(Me)
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub txtObservaciones_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtObservaciones.Leave
        gridPagos.CurrentCell = gridPagos.Rows(0).Cells(0)
        gridPagos.Select()
    End Sub

    Private Function cuentasCorrientesDelProveedorActualSegunDescripcion(ByVal description As String)
        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
        If IsNothing(proveedor) Then
            Return New List(Of CtaCteProveedor)
        Else
            Return DAOCtaCteProveedor.buscarCuentas()
        End If
    End Function

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        txtProveedor.Text = proveedor.id
        txtRazonSocial.Text = proveedor.razonSocial
    End Sub

    Private Sub mostrarBanco(ByVal banco As Banco)
        txtBanco.Text = banco.id
        txtNombreBanco.Text = banco.nombre
    End Sub

    Private Sub mostrarCuentaContable(ByVal cuenta As CuentaContable)
        'TODO
    End Sub

    Private Sub mostrarCheque(ByVal cheque As Cheque)
        'TODO
    End Sub

    Private Sub txtProveedor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProveedor.Leave
        Dim proveedor = DAOProveedor.buscarProveedorPorCodigo(txtProveedor.Text)
        If Not IsNothing(proveedor) Then
            mostrarProveedor(proveedor)
        Else
            txtRazonSocial.Text = ""
        End If
    End Sub

    Private Sub txtBanco_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBanco.KeyDown
        If e.KeyValue = Keys.Enter Then
            txtBanco_Leave(sender, Nothing)
        End If
    End Sub

    Private Sub txtBanco_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBanco.Leave
        Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(txtBanco.Text)
        If Not IsNothing(banco) Then
            mostrarBanco(banco)
        Else
            txtNombreBanco.Text = ""
        End If
    End Sub

    Private Sub txtOrdenPago_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrdenPago.Leave
        txtOrdenPago.Text = ceros(txtOrdenPago.Text, 6)
    End Sub

    Private Sub lstSeleccion_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstSeleccion.DoubleClick
        queryController = lstSeleccion.SelectedItem
        lstSeleccion.Visible = False
        lstConsulta.Visible = True
        txtConsulta.Visible = True
        lstConsulta.DataSource = queryController.query.Invoke("")
        txtConsulta.Focus()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click
        lstSeleccion.Visible = True
    End Sub

    Private Sub txtConsulta_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsulta.KeyDown
        If e.KeyValue = Keys.Enter Then
            lstConsulta.DataSource = queryController.query.Invoke(txtConsulta.Text)
            e.Handled = True
        End If
    End Sub

    Private Sub lstConsulta_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstConsulta.DoubleClick
        queryController.showMethod.Invoke(lstConsulta.SelectedValue)
        lstConsulta.Visible = False
        txtConsulta.Visible = False
        txtConsulta.Text = ""
    End Sub
End Class