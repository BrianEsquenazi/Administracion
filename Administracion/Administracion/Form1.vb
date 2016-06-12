Public Class Form1
    Dim dataGridBuilder As GridBuilder

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dataGridBuilder = New GridBuilder(DataGridView1)
        dataGridBuilder.addTextColumn(0, "Nombre")
        dataGridBuilder.addTextColumn(1, "Apellido")
        dataGridBuilder.addDateColumn(2, "Fecha Ejemplo")
        'Se agregarían las columnas numéricas (en un futuro próximo) y los formatos, tal cual tienen los texts ahora
    End Sub

End Class