Public Class CUFEProveedor

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        If validarCampos() Then
            btnClose.PerformClick()
        End If
    End Sub

    Private Function validarCampos()
        Dim validador As New Validator
        Me.Controls.OfType(Of CustomTextBox).ToList.ForEach(Sub(control) validador.validate(control.Text, control.Validator, control.Empty, descripcionPara(control.LabelAssociationKey)))
        Return validador.flush
    End Function

    Private Function descripcionPara(ByVal index As Integer)
        Return Me.Controls.OfType(Of CustomLabel).ToList.Find(Function(label) label.ControlAssociationKey = index).Text
    End Function

    Private Sub CUFEProveedor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CommonEventsHandler.setIndexTabNotCRUDForm(Me)
    End Sub
End Class