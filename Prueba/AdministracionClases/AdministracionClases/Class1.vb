Public Class CuentaContable

    Private numeroCuenta As Integer

    Public Sub New(ByVal valor As Integer)
        numeroCuenta = valor
    End Sub

    Public Function hola(ByVal value As String)
        Return "Hola manola " & value & " mi número de cuenta es: " & numeroCuenta
    End Function

End Class
