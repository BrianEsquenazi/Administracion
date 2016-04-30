Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim martin As CuentaContable
        martin = New CuentaContable(Val(TextBox1.Text))

        MsgBox(martin.hola(TextBox2.Text))
    End Sub
End Class
