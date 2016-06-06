Imports ClasesCompartidas

Public Class ProveedoresABM

    Dim organizadorABM As New FormOrganizer(Me, 800, 600)
    Dim observaciones As String
    Dim cufes As New List(Of Tuple(Of String, String)) From {Tuple.Create("", ""), Tuple.Create("", ""), Tuple.Create("", "")}

    Private Sub ProveedoresABM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbProvincia.DataSource = DAOProveedor.listarProvincias
        cmbRubro.DataSource = DAORubroProveedor.buscarRubroProveedorPorDescripcion("")
        cmbRegion.SelectedIndex = 0
        cmbCondicionIB1.SelectedIndex = 0
        cmbCondicionIB2.SelectedIndex = 0
        cmbCategoria2.SelectedIndex = 0

        organizadorABM.addControls({txtCodigo, txtRazonSocial}, txtDireccion, txtLocalidad, {cmbProvincia, txtCodigoPostal, cmbRegion}, {txtTelefono, txtDiasPlazo}, txtEmail, {txtObservaciones, txtCUIT}, {cmbTipoProveedor, cmbIVA}, txtCuenta, txtCheque, {cmbCondicionIB1, txtNroIB, txtPorcelProv, txtPorcelCABA}, {cmbRubro, txtNroSEDRONAR1}, {cmbCategoria1, cmbInscripcionIB})
        organizadorABM.addCompactedControls({txtCAI, txtCAIVto}, cmbCertificados, cmbEstado, cmbClasificacion, {btnObservaciones, btnCUFE})
        organizadorABM.addAnnexedControls(New List(Of CustomControl) From {txtCuentaDescripcion, cmbCondicionIB2, txtNroSEDRONAR2, cmbCategoria2, txtCategoria, txtCertificados, txtClasificacion})
        organizadorABM.setAddButtonClick(AddressOf agregar)
        organizadorABM.setDeleteButtonClick(AddressOf borrar)
        organizadorABM.setDefaultCleanButtonClick()
        organizadorABM.setDefaultCloseButtonClick()
        organizadorABM.setListButtonClick(AddressOf listado)
        organizadorABM.addQueryFunction(AddressOf DAOProveedor.buscarProveedorPorNombre, "Proveedores", AddressOf mostrarProveedor)
        organizadorABM.addQueryFunction(AddressOf DAOCuentaContable.buscarCuentaContablePorDescripcion, "Cuentas Contables", AddressOf mostrarCuenta)
        organizadorABM.compactOrganize()
    End Sub

    Private Sub agregar()
        MsgBox("Agregaste")
    End Sub

    Private Sub borrar()
        MsgBox("Borraste")
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        txtCodigo.Text = proveedor.id
        txtRazonSocial.Text = proveedor.razonSocial
    End Sub

    Private Sub mostrarCuenta(ByVal cuenta As CuentaContable)
        txtCuenta.Text = cuenta.id
        txtCuentaDescripcion.Text = cuenta.descripcion
    End Sub

    Private Sub listado()

    End Sub

    Private Sub btnObservaciones_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnObservaciones.Click
        Dim formularioObservaciones As New ObservacionesProveedor()

        formularioObservaciones.CustomTextBox1.Text = observaciones
        If formularioObservaciones.ShowDialog(Me) = DialogResult.OK Then
            observaciones = formularioObservaciones.CustomTextBox1.Text
        End If
    End Sub

    Private Sub btnCUFE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCUFE.Click
        Dim formularioCUFE As New CUFEProveedor()

        formularioCUFE.txtCUFE1.Text = cufes(0).Item1
        formularioCUFE.txtCUFE1Fecha.Text = cufes(0).Item2
        formularioCUFE.txtCUFE2.Text = cufes(1).Item1
        formularioCUFE.txtCUFE2Fecha.Text = cufes(1).Item2
        formularioCUFE.txtCUFE3.Text = cufes(2).Item1
        formularioCUFE.txtCUFE3Fecha.Text = cufes(2).Item2
        If formularioCUFE.ShowDialog(Me) = DialogResult.OK Then
            cufes.Clear()
            cufes.Add(Tuple.Create(formularioCUFE.txtCUFE1.Text, formularioCUFE.txtCUFE1Fecha.Text))
            cufes.Add(Tuple.Create(formularioCUFE.txtCUFE2.Text, formularioCUFE.txtCUFE2Fecha.Text))
            cufes.Add(Tuple.Create(formularioCUFE.txtCUFE3.Text, formularioCUFE.txtCUFE3Fecha.Text))
        End If
    End Sub

    Private Sub txtCuenta_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCuenta.Leave
        Dim cuenta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(txtCuenta.Text)
        If Not IsNothing(cuenta) Then
            txtCuentaDescripcion.Text = cuenta.descripcion
        Else
            txtCuentaDescripcion.Text = ""
        End If
    End Sub

    Private Sub txtCodigo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodigo.Leave
        Dim proveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtCodigo.Text)
        If Not IsNothing(proveedor) Then
            txtRazonSocial.Text = proveedor.razonSocial
        Else
            txtRazonSocial.Text = ""
        End If
    End Sub
End Class