Imports ClasesCompartidas

Public Class CargaIntereses

    Dim dataGridBuilder As GridBuilder

    Private Sub CargaIntereses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gridCtaCte.DataSource = SQLConnector.retrieveDataTable("get_carga_intereses")

        dataGridBuilder = New GridBuilder(gridCtaCte)
        dataGridBuilder.addDateColumn(0, "FechaOriginal")
        dataGridBuilder.addTextColumn(1, "DesProveOriginal")
        dataGridBuilder.addTextColumn(2, "FacturaOriginal")
        dataGridBuilder.addTextColumn(3, "Cuota")
        dataGridBuilder.addDateColumn(4, "fecha")
        dataGridBuilder.addFloatColumn(5, "Saldo")
        dataGridBuilder.addFloatColumn(6, "Intereses")
        dataGridBuilder.addFloatColumn(7, "IvaIntereses")
        dataGridBuilder.addTextColumn(8, "Referencia")
        dataGridBuilder.addTextColumn(9, "Clave")
        dataGridBuilder.addTextColumn(10, "NroInterno")


        
    End Sub
End Class

