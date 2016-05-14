Public Class CuentaContable
    Public id As String
    Public descripcion As String

    Public Sub New(ByVal identificador As Integer, ByVal desc As String)
        id = identificador
        descripcion = desc
    End Sub

    Public ReadOnly Property nombre() As String
        Get
            Return descripcion
        End Get
    End Property

    Public ReadOnly Property codigo() As String
        Get
            Return id
        End Get
    End Property
End Class
