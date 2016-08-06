Public Class RecibosProvisorios

    Dim commonEvent As CommonEventsHandler

    Private Sub RecibosProvisorios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        Cleanner.clean(Me)
    End Sub
End Class