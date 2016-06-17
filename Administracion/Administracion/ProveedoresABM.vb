﻿Imports ClasesCompartidas

Public Class ProveedoresABM

    Dim organizadorABM As New FormOrganizer(Me, 800, 600)
    Dim observaciones As String
    Dim cufe1 As Tuple(Of String, String) = Tuple.Create("", "")
    Dim cufe2 As Tuple(Of String, String) = Tuple.Create("", "")
    Dim cufe3 As Tuple(Of String, String) = Tuple.Create("", "")

    Private Sub ProveedoresABM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbProvincia.DataSource = DAOProveedor.listarProvincias
        cmbRubro.DisplayMember = "ToString"
        cmbRubro.ValueMember = "valueMember"
        cmbRubro.DataSource = DAORubroProveedor.buscarRubroProveedorPorDescripcion("")
        cmbRegion.SelectedIndex = 0
        cmbCondicionIB1.SelectedIndex = 0
        cmbCondicionIB2.SelectedIndex = 0
        cmbCategoria2.SelectedIndex = 0

        organizadorABM.addControls({txtCodigo, txtRazonSocial}, txtDireccion, txtLocalidad, {cmbProvincia, txtCodigoPostal, cmbRegion}, {txtTelefono, txtDiasPlazo}, txtEmail, {txtObservaciones, txtCUIT}, {cmbTipoProveedor, cmbIVA}, txtCuenta, txtCheque, {cmbCondicionIB1, txtNroIB, txtPorcelProv, txtPorcelCABA}, {cmbRubro, txtNroSEDRONAR1}, {cmbCategoria1, cmbInscripcionIB})
        organizadorABM.addCompactedControls({txtCAI, txtCAIVto}, cmbCertificados, cmbEstado, cmbCalificacion, {btnObservaciones, btnCUFE})
        organizadorABM.addAnnexedControls(New List(Of CustomControl) From {txtCuentaDescripcion, cmbCondicionIB2, txtNroSEDRONAR2, cmbCategoria2, txtCategoria, txtCertificados, txtCalificacion})
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
        If Not proveedor.estaDefinidoCompleto Then
            proveedor = DAOProveedor.buscarProveedorPorCodigo(proveedor.id)
        End If
        txtCodigo.Text = proveedor.id
        txtRazonSocial.Text = proveedor.razonSocial
        txtDireccion.Text = proveedor.direccion
        txtLocalidad.Text = proveedor.localidad
        cmbProvincia.SelectedIndex = proveedor.provincia
        txtCodigoPostal.Text = proveedor.codPostal
        cmbRegion.SelectedIndex = proveedor.region
        txtTelefono.Text = proveedor.telefono
        txtDiasPlazo.Text = proveedor.diasPlazo
        txtEmail.Text = proveedor.email
        txtObservaciones.Text = proveedor.observaciones
        txtCUIT.Text = proveedor.cuit
        cmbTipoProveedor.SelectedIndex = proveedor.tipo
        cmbIVA.SelectedIndex = proveedor.codIva
        mostrarCuenta(proveedor.cuenta)
        txtCheque.Text = proveedor.nombreCheque
        cmbCondicionIB1.SelectedIndex = proveedor.condicionIB1
        'cmbCondicionIB2.SelectedIndex = proveedor.condicionIB2
        txtNroIB.Text = proveedor.numeroIB
        txtPorcelProv.Text = proveedor.porceIBProvincia
        txtPorcelCABA.Text = proveedor.porceIBCABA
        mostrarRubro(proveedor.rubro)
        txtNroSEDRONAR1.Text = proveedor.numeroSEDRONAR
        txtNroSEDRONAR2.Text = proveedor.vtoSEDRONAR
        cmbCategoria1.SelectedIndex = proveedor.categoria
        cmbCategoria2.SelectedIndex = proveedor.categoriaCalif
        txtCategoria.Text = proveedor.vtoCategoria
        cmbInscripcionIB.SelectedIndex = proveedor.tipoInscripcionIB
        txtCAI.Text = proveedor.cai
        txtCAIVto.Text = proveedor.vtoCAI
        cmbCertificados.SelectedIndex = proveedor.certificados
        txtCertificados.Text = proveedor.vtoCertificados
        cmbEstado.SelectedIndex = proveedor.estado
        cmbCalificacion.SelectedIndex = proveedor.calificacion
        txtCalificacion.Text = proveedor.vtoCalificacion

        observaciones = proveedor.observacionCompleta
        cufe1 = Tuple.Create(proveedor.cufe1, proveedor.vtoCUFE1)
        cufe2 = Tuple.Create(proveedor.cufe2, proveedor.vtoCUFE2)
        cufe3 = Tuple.Create(proveedor.cufe3, proveedor.vtoCUFE3)
    End Sub

    Private Sub mostrarCuenta(ByVal cuenta As CuentaContable)
        If Not IsNothing(cuenta) Then
            txtCuenta.Text = cuenta.id
            txtCuentaDescripcion.Text = cuenta.descripcion
        End If
    End Sub

    Private Sub mostrarRubro(ByVal rubro As RubroProveedor)
        If Not IsNothing(rubro) Then
            cmbRubro.SelectedValue = rubro.codigo
        Else
            cmbRubro.SelectedValue = -1
        End If
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

        formularioCUFE.txtCUFE1.Text = cufe1.Item1
        formularioCUFE.txtCUFE1Fecha.Text = cufe1.Item2
        formularioCUFE.txtCUFE2.Text = cufe2.Item1
        formularioCUFE.txtCUFE2Fecha.Text = cufe2.Item2
        formularioCUFE.txtCUFE3.Text = cufe3.Item1
        formularioCUFE.txtCUFE3Fecha.Text = cufe3.Item2()
        If formularioCUFE.ShowDialog(Me) = DialogResult.OK Then
            cufe1 = Tuple.Create(formularioCUFE.txtCUFE1.Text, formularioCUFE.txtCUFE1Fecha.Text)
            cufe2 = Tuple.Create(formularioCUFE.txtCUFE2.Text, formularioCUFE.txtCUFE2Fecha.Text)
            cufe3 = Tuple.Create(formularioCUFE.txtCUFE3.Text, formularioCUFE.txtCUFE3Fecha.Text)
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