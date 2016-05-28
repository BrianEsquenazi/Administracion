Public Class Banco
    Public id As Short
    Public nombre As String
    Public cuenta As CuentaContable

    Public Sub New(ByVal codigo As Short, ByVal descripcion As String, ByVal cuentaContable As CuentaContable)
        id = codigo
        nombre = descripcion
        cuenta = cuentaContable
    End Sub

    Public ReadOnly Property descripcion() As String
        Get
            Return nombre
        End Get
    End Property

    Public ReadOnly Property codigo() As String
        Get
            Return id
        End Get
    End Property

End Class
