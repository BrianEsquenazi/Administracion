Imports ClasesCompartidas

Public Class Form1
    Dim dataGridBuilder As GridBuilder

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim prov As Proveedor = DAOProveedor.buscarProveedorPorCodigo("10057766117")
        dataGridBuilder = New GridBuilder(DataGridView1)
        dataGridBuilder.addTextColumn(0, "Codigo")
        dataGridBuilder.addTextColumn(1, "Razón Social")
        dataGridBuilder.addDateColumn(2, "Vto CAI")
        dataGridBuilder.addTextColumn(3, "Razón Social")
        'Se agregarían las columnas numéricas (en un futuro próximo) y los formatos, tal cual tienen los texts ahora
        DataGridView1.Rows.Add(prov.id, prov.razonSocial, prov.vtoCAI.ToString, prov.rubro.ToString)
    End Sub

End Class