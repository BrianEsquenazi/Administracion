Public Class DetalleCompraCuentaCorriente
    Public tipo, punto, numero, letra, fecha, impresion As String
    Public saldo, total As Double
    Public numInterno As Integer
    Public proveedor As Proveedor

    Public Sub New(ByVal tipoDoc As String, ByVal tipoImpre As String, ByVal punt As String, ByVal nro As String, ByVal letraString As String,
                   ByVal fechaString As String, ByVal restante As Double, ByVal valorTotal As Double,
                   ByVal interno As String, ByVal prov As Proveedor)
        tipo = tipoDoc
        impresion = tipoImpre
        punto = punt
        numero = nro
        letra = letraString
        fecha = fechaString
        saldo = restante
        total = valorTotal
        numInterno = interno
        proveedor = prov
    End Sub

    Public Overrides Function ToString() As String
        Return impresion & " - " & letra & " - " & punto & " - " & numero & " - " & fecha & " - " & asDoubleString(saldo).PadLeft(10, "_")
    End Function

    Private Function asDoubleString(ByVal value) As String
        Dim originalString As String = value.ToString
        If originalString.IndexOf(",") = -1 Then
            Return originalString & "," & "".PadLeft(2, "0")
        Else
            Dim difference As Integer = originalString.Count - originalString.IndexOf(",") - 1
            If difference < 2 Then
                Return originalString & "".PadLeft(2 - difference, "0")
            End If
        End If
        Return originalString
        Return value.ToString
    End Function

End Class
