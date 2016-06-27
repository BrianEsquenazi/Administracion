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
                    Handles Proveedor.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            aa = Proveedor.Text
            Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(aa)
            ProveedorRazon.Text = CampoProveedor.razonSocial
            Call Proceso()
            ProveedorRazon.Focus()
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            ProveedorRazon.Focus()
        End If
    End Sub

    Private Sub CuentaCorrientePantalla_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dataGridBuilder = New GridBuilder(GRilla)

        dataGridBuilder.addTextColumn(0, "Codigo")
        dataGridBuilder.addTextColumn(1, "Razón Social")
        dataGridBuilder.addDateColumn(2, "Vto CAI")
        dataGridBuilder.addTextColumn(3, "Razón Social")

    End Sub

    Private Sub GRilla_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GRilla.CellContentClick

    End Sub

    Private Sub Proceso()

        Dim orddesde As String
        Dim ordhasta As String

        orddesde = "20160301"
        ordhasta = "20160631"

        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("procesoSifere", ordDesde, ordHasta)

        For Each row As DataRow In tabla.Rows

            Dim CamposImputac As New Imputac(row.Item(0).ToString, row.Item(1), row.Item(2).ToString, row.Item(3).ToString, row.Item(4).ToString, row.Item(5).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString)

            dataGridBuilder.addColumn(row.Item(0).ToString, row.Item(1), row.Item(2))


        Next

        MsgBox("Proceso Finalizado", MsgBoxStyle.Information)



    End Sub


End Class