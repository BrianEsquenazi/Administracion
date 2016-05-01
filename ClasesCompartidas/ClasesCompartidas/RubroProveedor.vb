Public Class RubroProveedor
    Private id As Integer
    Private descripcion As String

    Public Sub New(ByVal codigo As Integer, ByVal nombre As String)
        id = codigo
        descripcion = nombre
    End Sub
End Class
