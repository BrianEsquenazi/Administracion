Public Class Banco
    Private id As Integer
    Private nombre As String
    Private cuenta As CuentaContable

    Public Sub New(ByVal codigo As Integer, ByVal descripcion As String, ByVal cuentaContable As CuentaContable)
        id = codigo
        nombre = descripcion
        cuenta = cuentaContable
    End Sub

End Class
