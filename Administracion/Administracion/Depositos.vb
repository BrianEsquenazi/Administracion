Imports ClasesCompartidas

Public Class Depositos
    Dim dataGridBuilder As GridBuilder
    Dim showFunction As ShowMethod
    Dim cheques As New List(Of Cheque)

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

    Private Function sumaImportes()
        Dim valorImportes As Double = 0
        For Each row As DataGridViewRow In gridCheques.Rows
            valorImportes += row.Cells(3).Value
        Next
        Return valorImportes
    End Function

    Private Function validarCampos()
        Dim validador As New Validator

        validador.validate(Me)
        validador.alsoValidate(CustomConvert.toDoubleOrZero(CustomTextBox1.Text) = sumaImportes(), "El campo importe tiene que ser igual a la suma de los cheques (" & sumaImportes() & ")")
        validador.alsoValidate(cheques.Count = gridCheques.Rows.Count, "La cantidad de cheques registrados no coincide con la cantidad de filas de la tabla")

        Return validador.flush
    End Function

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Close()
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
        gridCheques.Rows.Clear()
        cheques.Clear()
        Me.Width = formNormalWidth()
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
        If Not cheques.Any(Function(otroCheque) otroCheque.igualA(cheque)) Then
            cheques.Add(cheque)
            gridCheques.Rows.Add(cheque.tipo, cheque.numero, cheque.fecha, cheque.importe)
        End If
    End Sub

    Private Sub lstSeleccion_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSeleccion.DoubleClick
        If lstSeleccion.SelectedItem = "Bancos" Then
            showFunction = AddressOf mostrarBanco
            lstConsulta.DataSource = DAOBanco.buscarBancoPorNombre("")
        Else
            showFunction = AddressOf mostrarCheque
            lstConsulta.DataSource = Nothing
            DAODeposito.buscarCheques().ForEach(Sub(cheque) lstConsulta.Items.Add(cheque))
        End If
        lstSeleccion.Visible = False
        lstConsulta.Visible = True
    End Sub

    Private Sub lstConsulta_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstConsulta.DoubleClick
        showFunction.Invoke(lstConsulta.SelectedItem)
        If lstSeleccion.SelectedItem = "Bancos" Then
            lstConsulta.Visible = False
            Me.Width = formNormalWidth()
        Else
            lstConsulta.Items.Remove(lstConsulta.SelectedItem)
        End If
    End Sub

    Private Sub gridCheques_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles gridCheques.UserDeletingRow
        Dim chequeABorrar As Cheque = cheques.Find(Function(cheque) cheque.numero = e.Row.Cells(1).Value And cheque.fecha = e.Row.Cells(2).Value And cheque.importe = e.Row.Cells(3).Value)
        If Not IsNothing(chequeABorrar) Then
            cheques.Remove(chequeABorrar)
        End If
    End Sub

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If validarCampos() Then
            'agregar
        End If
    End Sub
End Class