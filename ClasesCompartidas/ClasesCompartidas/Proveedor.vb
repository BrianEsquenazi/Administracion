Public Class Proveedor
    Public id As Long
    Public razonSocial, direccion, codPostal, localidad, provincia As String

    Public Sub New(ByVal codigo As Long, ByVal nombre As String)
        id = codigo
        razonSocial = nombre
    End Sub

    Public Overrides Function ToString() As String
        Return razonSocial
    End Function
End Class
