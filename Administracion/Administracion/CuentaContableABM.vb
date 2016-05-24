﻿Imports ClasesCompartidas

Public Class CuentaContableABM

    Private Function validarCampos(ByVal validarDescripcion As Boolean)
        Dim validador As New Validator
        validador.validarNoVacio(txtCodigo.Text, "código")
        If validarDescripcion Then
            validador.validarNoVacio(txtDescripcion.Text, "descripción")
        End If
        Return validador.flush()
    End Function

    Private Function validarCampos()
        Return validarCampos(True)
    End Function

    Private Function validarCodigo()
        Return validarCampos(False)
    End Function

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If validarCampos() Then
            Dim cuenta As New CuentaContable(txtCodigo.Text, txtDescripcion.Text)
            DAOCuentaContable.agregarCuentaContable(cuenta)
            limpiarCampos()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If validarCodigo() Then
            Dim cuenta As New CuentaContable(txtCodigo.Text, txtDescripcion.Text)
            DAOCuentaContable.eliminarCuentaContable(cuenta)
            limpiarCampos()
        End If
    End Sub

    Private Sub CuentaContableABM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ocultarQueries()
        CommonEventsHandler.setIndexTab(Me)
        txtCodigo.Focus()
    End Sub

    Private Sub lstQuery_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cuenta As CuentaContable = lstQuery.SelectedItem
        txtCodigo.Text = cuenta.id
        txtDescripcion.Text = cuenta.descripcion
        ocultarQueries()
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        mostrarSoloTxtQuery()
    End Sub

    Private Sub txtQuery_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQuery.KeyDown
        If e.KeyCode = Keys.Enter Then
            cargarListaSegun(ActiveControl.Text)
        End If
    End Sub

    Private Sub pantallaQuery(ByVal textBoxVisible As Boolean, ByVal listVisible As Boolean, ByVal height As Integer)
        txtQuery.Visible = textBoxVisible
        lstQuery.Visible = listVisible
        Me.Height = height
    End Sub

    Private Sub mostrarSoloTxtQuery()
        pantallaQuery(True, False, 250)
        txtQuery.Text = ""
        txtQuery.Focus()
    End Sub

    Private Sub ocultarQueries()
        pantallaQuery(False, False, 220)
        txtQuery.Text = ""
    End Sub

    Private Sub mostrarQueries()
        pantallaQuery(True, True, 485)
    End Sub

    Private Sub cargarListaSegun(ByVal stringBusqueda As String)
        Dim cuentas As List(Of CuentaContable)

        cuentas = DAOCuentaContable.buscarCuentaContablePorDescripcion(stringBusqueda)

        lstQuery.DisplayMember = "nombre"
        lstQuery.ValueMember = "codigo"
        lstQuery.DataSource = cuentas

        mostrarQueries()
    End Sub

    Private Sub limpiarCampos()
        Cleanner.clean(Me)
        ocultarQueries()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClean.Click
        limpiarCampos()
    End Sub

    Private Sub txtCodigo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodigo.KeyDown
        If e.KeyValue = Keys.Enter Then
            Dim cuenta As CuentaContable = DAOCuentaContable.buscarCuentaContablePorCodigo(txtCodigo.Text)
            If Not IsNothing(cuenta) Then
                txtDescripcion.Text = cuenta.descripcion
            End If
        End If
    End Sub
End Class
