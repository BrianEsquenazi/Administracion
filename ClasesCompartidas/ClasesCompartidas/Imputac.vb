Public Class Imputac

    Public fechaord As String
    Public debito As Double
    Public proveedor As String
    Public cuenta As String
    Public nrointerno As String
    Public punto As String
    Public numero As String
    Public despacho As String
    Public cuit As String

    Public Sub New(ByVal fechaord2 As String, ByVal debito2 As Double, ByVal proveedor2 As String, ByVal cuenta2 As String, ByVal nrointerno2 As String, ByVal punto2 As String, ByVal numero2 As String, ByVal despacho2 As String, ByVal cuit2 As String)

        fechaord = fechaord2
        debito = debito2
        proveedor = proveedor2
        cuenta = cuenta2
        nrointerno = nrointerno2
        punto = punto2
        numero = numero2
        despacho = despacho2
        cuit = cuit2

    End Sub

End Class
