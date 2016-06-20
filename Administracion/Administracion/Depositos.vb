Imports ClasesCompartidas

Public Class Depositos
    Dim dataGridBuilder As GridBuilder
    Dim showFunction As ShowMethod

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dataGridBuilder = New GridBuilder(gridCheques)
        dataGridBuilder.addTextColumn(0, "Tipo")
        dataGridBuilder.addTextColumn(1, "Número")
        dataGridBuilder.addDateColumn(2, "Fecha")
        dataGridBuilder.addPositiveFloatColumn(3, "Importe")
        lstSeleccion.SelectedIndex = 0
        Me.Width = formNormalWidth()
        CommonEventsHandler.setIndexTab(Me)
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        gridCheques.Rows.Clear()
    End Sub

    Private Sub mostrarSeleccionDeConsulta()
        lstConsulta.Visible = False
        lstSeleccion.Visible = True
        Me.Width = formWithListsWidth()
    End Sub

    Private Function formNormalWidth()
        Return 480
    End Function

    Private Function formWithListsWidth()
        Return 800
    End Function

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click
        mostrarSeleccionDeConsulta()
    End Sub

    Private Sub mostrarBanco(ByVal banco As Banco)
        txtCodigoBanco.Text = banco.id
        txtDescripcionBanco.Text = banco.nombre
    End Sub

    Private Sub mostrarCheque(ByVal cheque As Cheque)
        'Agregarlo al grid
    End Sub

    Private Sub lstSeleccion_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSeleccion.DoubleClick
        If lstSeleccion.SelectedItem = "Bancos" Then
            showFunction = AddressOf mostrarBanco
            lstConsulta.DataSource = DAOBanco.buscarBancoPorNombre("")
        Else
            showFunction = AddressOf mostrarCheque
            lstConsulta.DataSource = DAODeposito.buscarCheques()
        End If
        lstSeleccion.Visible = False
        lstConsulta.Visible = True
    End Sub

    Private Sub lstConsulta_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstConsulta.DoubleClick
        showFunction.Invoke(lstConsulta.SelectedValue)
        If lstSeleccion.SelectedItem = "Bancos" Then
            lstConsulta.Visible = False
            Me.Width = formNormalWidth()
        Else
            lstConsulta.Items.Remove(lstConsulta.SelectedItem)
        End If
    End Sub

    Private Sub lstConsulta_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class