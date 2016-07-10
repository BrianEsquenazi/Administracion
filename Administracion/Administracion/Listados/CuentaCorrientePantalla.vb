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

        GRilla.Columns(0).Width = 50
        GRilla.Columns(1).Width = 50
        GRilla.Columns(2).Width = 50
        GRilla.Columns(3).Width = 100
        GRilla.Columns(4).Width = 100
        GRilla.Columns(5).Width = 100
        GRilla.Columns(6).Width = 100
        GRilla.Columns(7).Width = 100

        opcPendiente.Checked = True
        opcCompleto.Checked = False

    End Sub

    Private Sub Proceso()

        Dim WRenglon As Integer

        GRilla.Rows.Clear()
        GRilla.Rows.Add()
        WRenglon = 0

        REM Reviso el cual esta checkeado asi le pongo los valores a Tipo
        Dim WTipo As Char
        WTipo = "T"

        If (opcPendiente.Checked) Then
            WTipo = "P"
        End If

        REM dada fix CAMBIAR Al uso de dao!!
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta_corriente_proveedores_deuda", txtProveedor.Text, WTipo)

        For Each row As DataRow In tabla.Rows

            Dim CamposCtaCtePrv As New CtaCteProveedoresDeuda(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString, row.Item(3).ToString, row.Item(4), row.Item(5), row.Item(6).ToString, row.Item(7).ToString)

            WRenglon = WRenglon + 1

            GRilla.Rows.Add()

            GRilla.Item(0, WRenglon).Value = CamposCtaCtePrv.Tipo
            GRilla.Item(1, WRenglon).Value = CamposCtaCtePrv.letra
            GRilla.Item(2, WRenglon).Value = CamposCtaCtePrv.punto
            GRilla.Item(3, WRenglon).Value = CamposCtaCtePrv.numero
            GRilla.Item(4, WRenglon).Value = CamposCtaCtePrv.total
            GRilla.Item(5, WRenglon).Value = CamposCtaCtePrv.saldo
            GRilla.Item(6, WRenglon).Value = CamposCtaCtePrv.fecha
            GRilla.Item(7, WRenglon).Value = CamposCtaCtePrv.vencimiento



        Next


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