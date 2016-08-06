Imports ClasesCompartidas

Public Class AplicacionComprobantes

    Private Sub txtProveedor_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProveedor.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            btnProceso.PerformClick()
        End If
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        If IsNothing(proveedor) Then : Exit Sub : End If
        txtProveedor.Text = proveedor.id
        txtRazon.Text = proveedor.razonSocial
        Proceso()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click
        lstAyuda.Visible = True
        txtAyuda.Visible = True
        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")
        txtAyuda.Focus()
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarProveedor(lstAyuda.SelectedValue)
    End Sub

    Private Sub txtAyuda_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAyuda.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre(txtAyuda.Text)
        End If
    End Sub

    Private Sub txtAyuda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnProceso_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProceso.Click
        Dim proveedor = DAOProveedor.buscarProveedorPorCodigo(ceros(txtProveedor.Text, 11))
        If Not IsNothing(proveedor) Then
            mostrarProveedor(proveedor)
            ' cargar tabla
            '                ctacteprv()
            '                Tipo
            '                Letra
            '                Punto
            '                Numero
            '                Fecha
            '                Importe
            '                Saldo
            '                Aplica

            '1 y 2 positivo
            '3 y 5 negativo
        Else
            txtProveedor.Text = ""
            txtRazon.Text = ""
            txtProveedor.Focus()
        End If
    End Sub

    Private Sub dtgCuentas_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtgCuentas.KeyDown
        If (e.KeyCode = Keys.Enter And dtgCuentas.CurrentCell.ColumnIndex = 7) Then

            If IsNothing(dtgCuentas.Rows(dtgCuentas.CurrentCell.RowIndex).Cells(dtgCuentas.CurrentCell.ColumnIndex).Value) Then
                dtgCuentas.Rows(dtgCuentas.CurrentCell.RowIndex).Cells(dtgCuentas.CurrentCell.ColumnIndex).Value = 0
            End If

            Dim tipo As String = dtgCuentas.Rows(dtgCuentas.CurrentCell.RowIndex).Cells(0).Value
            Dim saldoNuevo As Double
            Dim valorAplica As Double = Convert.ToDouble(dtgCuentas.Rows(dtgCuentas.CurrentCell.RowIndex).Cells(dtgCuentas.CurrentCell.ColumnIndex).Value)

            Select Case tipo
                Case "01", "02"
                    saldoNuevo = Convert.ToDouble(txtSaldo.Text) + valorAplica
                Case "03", "05"
                    saldoNuevo = Convert.ToDouble(txtSaldo.Text) - valorAplica
            End Select

            txtSaldo.Text = saldoNuevo.ToString
        End If
    End Sub

    Private Sub AplicacionComprobantes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim gridFormasBuilder As New GridBuilder(dtgCuentas)
        gridFormasBuilder.addTextColumn(0, "Tipo", True)
        gridFormasBuilder.addTextColumn(1, "Letra", True)
        gridFormasBuilder.addTextColumn(2, "Punto", True)
        gridFormasBuilder.addTextColumn(3, "Numero", True)
        gridFormasBuilder.addDateColumn(4, "Fecha", True)
        gridFormasBuilder.addFloatColumn(5, "Importe", True)
        gridFormasBuilder.addFloatColumn(6, "Saldo", True)
        gridFormasBuilder.addFloatColumn(7, "Aplica", True)
    End Sub

    Private Sub Proceso()

        txtAyuda.Visible = False
        lstAyuda.Visible = False

        Dim WRenglon As Integer
        Dim WSuma As Double

        dtgCuentas.Rows.Clear()
        dtgCuentas.Rows.Add()
        WRenglon = 0

        REM Reviso el cual esta checkeado asi le pongo los valores a Tipo

        REM dada fix CAMBIAR Al uso de dao!!
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta_corriente_proveedores_deuda", txtProveedor.Text, "P")

        For Each row As DataRow In tabla.Rows

            Dim CamposCtaCtePrv As New CtaCteProveedoresDeuda(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString, row.Item(3).ToString, row.Item(4), row.Item(5), row.Item(6).ToString, row.Item(7).ToString)


            dtgCuentas.Rows.Add()

            dtgCuentas.Item(0, WRenglon).Value = CamposCtaCtePrv.Tipo
            dtgCuentas.Item(1, WRenglon).Value = CamposCtaCtePrv.letra
            dtgCuentas.Item(2, WRenglon).Value = CamposCtaCtePrv.punto
            dtgCuentas.Item(3, WRenglon).Value = CamposCtaCtePrv.numero
            dtgCuentas.Item(4, WRenglon).Value = CamposCtaCtePrv.fecha
            dtgCuentas.Item(5, WRenglon).Value = formatonumerico(CamposCtaCtePrv.total, "########0.#0", ".")
            dtgCuentas.Item(6, WRenglon).Value = formatonumerico(CamposCtaCtePrv.saldo, "########0.#0", ".")


            WRenglon = WRenglon + 1
            WSuma = WSuma + CamposCtaCtePrv.saldo

        Next

        dtgCuentas.AllowUserToAddRows = False
        txtAyuda.Text = ""

        txtSaldo.Text = "0.00"

        dtgCuentas.CurrentCell = dtgCuentas.Item(7, 0)
        dtgCuentas.Rows(0).Cells(7).Selected = True
        dtgCuentas.Focus()



    End Sub

    Private Sub btnGraba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGraba.Click
        Dim saldo As Double = Convert.ToDouble(txtSaldo.Text)

        If saldo <> 0 Then
            MsgBox("El saldo actual no esta equilibrado")
        Else

        End If
    End Sub
End Class
