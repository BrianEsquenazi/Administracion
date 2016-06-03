Public Class RubroProveedor
    Public codigo As Integer
    Public descripcion As String

    Public Sub New(ByVal cod As Integer, ByVal nombre As String)
        codigo = cod
        descripcion = nombre
    End Sub

    Public Overrides Function ToString() As String
        Return descripcion
    End Function
End Class
