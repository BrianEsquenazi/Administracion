Imports ClasesCompartidas

Public Class TipoCambioABM

    Dim organizadorABM As New FormOrganizer(Me, 350, 600)

    'Private Sub ocultarQueries()
    '    pantallaQuery(False, False, 240)
    '    txtQuery.Text = ""
    'End Sub

    'Private Sub pantallaQuery(ByVal textBoxVisible As Boolean, ByVal listVisible As Boolean, ByVal height As Integer)
    '    txtQuery.Visible = textBoxVisible
    '    lstQuery.Visible = listVisible
    '    Me.Height = height
    'End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTab(Me)
        organizadorABM.addControls(New List(Of CustomControl) From {txtFecha, txtParidad})
        organizadorABM.setAddButtonClick(AddressOf btnAddClick)
        organizadorABM.setDeleteButtonClick(AddressOf btnDeleteClick)
        organizadorABM.setDefaultCleanButtonClick()
        organizadorABM.setDefaultCloseButtonClick()
        organizadorABM.setListButtonClick(AddressOf btnListClick)
        organizadorABM.setQueryButtonClick(AddressOf btnQueryClick)
        organizadorABM.organize()
    End Sub

    Private Sub btnAddClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim cambio As New TipoDeCambio(txtFecha.Text, txtParidad.Text)
        DAOTipoCambio.agregarTipoCambio(cambio)
    End Sub

    Private Sub btnDeleteClick(ByVal sender As Object, ByVal e As EventArgs)
        DAOTipoCambio.eliminarTipoCambio(txtFecha.Text)
    End Sub

    Private Sub btnQueryClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de consulta")
    End Sub

    Private Sub btnListClick(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Private Sub txtFecha_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFecha.Leave
        Dim cambio As TipoDeCambio = DAOTipoCambio.buscarTipoCambioPorFecha(txtFecha.Text)
        If Not IsNothing(cambio) Then
            txtFecha.Text = cambio.fecha
            txtParidad.Text = cambio.paridad
        End If
    End Sub
End Class