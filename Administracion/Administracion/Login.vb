Imports ClasesCompartidas

Public Class Login
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbEntity.DataSource = Globals.connectionStringNames()
    End Sub

    Private Function validarCampos()
        Dim validador As New Validator
        validador.validarNoVacio(cmbEntity.Text, False, "Empresa")
        Return validador.flush()
    End Function

    Private Sub btnAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        If validarCampos() Then
            Globals.empresa = cmbEntity.Text
            MenuPrincipal.Show()
            Close()
        End If
    End Sub
End Class