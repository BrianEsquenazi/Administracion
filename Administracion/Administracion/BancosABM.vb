Imports ClasesCompartidas

Public Class BancosABM

    Private Sub limpiarCampos()
        Cleanner.clean(Me)
        ocultarQueries()
    End Sub

    Private Function validarCampos(ByVal agregar As Boolean)
        Dim validador As New Validator
        validador.validarPositivo(txtCodigo.Text, "código", Short.MaxValue)
        If agregar Then
            validador.validarNoVacio(txtNombre.Text, "nombre")
            validador.validarNoVacio(txtCuenta.Text, "cuenta")
        End If
        Return validador.flush()
    End Function

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If validarCampos(True) Then
            Dim cuenta As New CuentaContable(txtCuenta.Text, txtDescripcion.Text)
            Dim banco As New Banco(txtCodigo.Text, txtNombre.Text, cuenta)
            DAOBanco.agregarBanco(banco)
            limpiarCampos()
        End If
    End Sub

    Private Sub BancosABM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTab(Me)
        ocultarQueries()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If validarCampos(False) Then
            Dim cuenta As New CuentaContable(txtCuenta.Text, txtDescripcion.Text)
            Dim banco As New Banco(txtCodigo.Text, txtNombre.Text, cuenta)
            DAOBanco.eliminarBanco(banco)
            limpiarCampos()
        End If
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        mostrarQueries()
        cargarListaSegun("")
        txtQuery.Focus()
    End Sub

    Private Sub pantallaQuery(ByVal textBoxVisible As Boolean, ByVal listVisible As Boolean, ByVal height As Integer)
        txtQuery.Visible = textBoxVisible
        lstQuery.Visible = listVisible
        Me.Height = height
    End Sub

    Private Sub ocultarQueries()
        pantallaQuery(False, False, 240)
        txtQuery.Text = ""
    End Sub

    Private Sub mostrarQueries()
        pantallaQuery(True, True, 505)
    End Sub

    Private Sub txtQuery_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQuery.KeyDown
        If e.KeyCode = Keys.Enter Then
            cargarListaSegun(ActiveControl.Text)
        End If
    End Sub

    Private Sub cargarListaSegun(ByVal stringBusqueda As String)
        Dim bancos As List(Of Banco)

        bancos = DAOBanco.buscarBancoPorNombre(stringBusqueda)
        lstQuery.DataSource = bancos

        mostrarQueries()
    End Sub

    Private Sub lstQuery_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstQuery.DoubleClick
        Dim banco As Banco = lstQuery.SelectedItem
        txtCodigo.Text = banco.id
        txtNombre.Text = banco.nombre
        txtCuenta.Text = banco.cuenta.id
        txtDescripcion.Text = banco.cuenta.descripcion
        ocultarQueries()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClean.Click
        limpiarCampos()
    End Sub

    Private Sub txtCodigo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigo.Leave
        Dim banco As Banco = DAOBanco.buscarBancoPorCodigo(txtCodigo.Text)
        If Not IsNothing(banco) Then
            txtNombre.Text = banco.nombre
            txtCuenta.Text = banco.cuenta.id
            txtDescripcion.Text = banco.cuenta.descripcion
        End If
    End Sub

    Private Sub txtCuenta_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCuenta.Leave
        Dim cuenta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(txtCuenta.Text)
        If Not IsNothing(cuenta) Then
            txtDescripcion.Text = cuenta.descripcion
        Else
            txtDescripcion.Text = ""
        End If
    End Sub
End Class