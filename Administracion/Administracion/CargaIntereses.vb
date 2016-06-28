Imports ClasesCompartidas

Public Class CargaIntereses

    Dim dataGridBuilder As GridBuilder

    Private Sub CargaIntereses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DAOCtaCteProveedor.buscarCuentas().ForEach(Sub(cuenta) gridCtaCte.Rows.Add(cuenta.fechaOriginal))
        REM gridCtaCte.Items.Add(Cheque))

        gridCtaCte.AllowUserToAddRows = False
        gridCtaCte.Rows(1).Cells(1).ReadOnly = False
        gridCtaCte.Columns(7).ReadOnly = False
        gridCtaCte.Columns(8).ReadOnly = False
        gridCtaCte.Columns(9).Visible = False
        gridCtaCte.Columns(10).Visible = False

        'gridCtaCte.DataSource = SQLConnector.retrieveDataTable("get_carga_interesesssssss")
        'REM AGREGAR POR DAO CON FOREACH Y ADD ROW DIRECTO (EN DEPOSITOS ESTA HECHO)

        'dataGridBuilder = New GridBuilder(gridCtaCte)

        'dataGridBuilder.addDateColumn(0, "FechaOriginal")
        'dataGridBuilder.addTextColumn(1, "DesProveOriginal")
        'dataGridBuilder.addTextColumn(2, "FacturaOriginal")
        'dataGridBuilder.addTextColumn(3, "Cuota")
        'dataGridBuilder.addDateColumn(4, "Vencimiento")
        'dataGridBuilder.addFloatColumn(5, "Saldo")
        'dataGridBuilder.addFloatColumn(6, "Intereses")
        'dataGridBuilder.addFloatColumn(7, "IvaIntereses")
        'gridCtaCte.Columns(7).ReadOnly = False
        'dataGridBuilder.addTextColumn(8, "Referencia")
        'gridCtaCte.Columns(8).ReadOnly = False
        'dataGridBuilder.addTextColumn(9, "Clave")
        'gridCtaCte.Columns(9).Visible = False
        'dataGridBuilder.addTextColumn(10, "NroInterno")
        'gridCtaCte.Columns(10).Visible = False

    End Sub

End Class

