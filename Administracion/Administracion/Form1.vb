Public Class Form1

    Dim organizadorABM As New FormOrganizer(Me, 800, 600)

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTab(Me)
        organizadorABM.addControls(New List(Of CustomControl) From {txtCodigo, txtDescripcion})
        organizadorABM.setAddButtonClick(AddressOf btnAddClick)
        organizadorABM.setDeleteButtonClick(AddressOf btnDeleteClick)
        organizadorABM.setCleanButtonClick(AddressOf btnCleanClick)
        organizadorABM.setCloseButtonClick(AddressOf btnCloseClick)
        organizadorABM.setListButtonClick(AddressOf btnListClick)
        organizadorABM.setQueryButtonClick(AddressOf btnQueryClick)
        organizadorABM.organize()
    End Sub

    Private Sub btnAddClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de agregar")
    End Sub

    Private Sub btnDeleteClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de borrar")
    End Sub

    Private Sub btnCleanClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de limpiar")
    End Sub

    Private Sub btnQueryClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de consulta")
    End Sub

    Private Sub btnListClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de lista")
    End Sub

    Private Sub btnCloseClick(ByVal sender As Object, ByVal e As EventArgs)
        MsgBox("Clickeaste el botón de cerrar")
    End Sub
End Class