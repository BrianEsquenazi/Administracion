﻿Imports ClasesCompartidas

Public Class ProveedoresABM

    Dim organizadorABM As New FormOrganizer(Me, 800, 600)

    Private Sub ProveedoresABM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbProvincia.DataSource = DAOProveedor.listarProvincias
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
End Class