Public Class Proveedor
    Public id As String
    Public razonSocial, direccion, codPostal, localidad, provincia As String

    Public Sub New(ByVal codigo As String, ByVal nombre As String)
        id = codigo
        razonSocial = nombre
    End Sub

    Public Overrides Function ToString() As String
        Return razonSocial
    End Function
End Class
