Public Class Compra
    Public nroInterno As Integer
    Public proveedor As Proveedor
    Public tipoDocumento, formaPago, tipoPago As Integer
    Public tipoDocumentoDescripcion, letra, numero, fechaEmision, fechaIVA, fechaVto1, fechaVto2, remito, despacho As String
    Public paridad, neto, iva21, ivaRG, iva27, percibidoIB, exento, iva105, total As Double
    Public soloIVA As Integer

    Public Sub New(ByVal interno As Integer, ByVal prov As Proveedor, ByVal tipoDoc As Integer, ByVal tipoDocDesc As String,
                   ByVal forma As Integer, ByVal tipo As Integer, ByVal letraDoc As String, ByVal num As String, ByVal emision As String,
                   ByVal fecIVA As String, ByVal vto1 As String, ByVal vto2 As String, ByVal cotizacion As Double, ByVal importeNeto As Double,
                   ByVal importeIVA21 As Double, ByVal importeIVARG As Double, ByVal importeIVA27 As Double, ByVal importeIB As Double,
                   ByVal importeExento As Double, ByVal importeIVA105 As Double, ByVal importeTotal As Double, ByVal ivaSolo As Integer)
        nroInterno = interno
        proveedor = prov
        tipoDocumento = tipoDoc
        tipoDocumentoDescripcion = tipoDocDesc
        formaPago = forma
        tipoPago = tipo
        letra = letraDoc
        numero = num
        fechaEmision = emision
        fechaIVA = fecIVA
        fechaVto1 = vto1
        fechaVto2 = vto2
        paridad = cotizacion
        neto = importeNeto
        iva21 = importeIVA21
        iva27 = importeIVA27
        ivaRG = importeIVARG
        iva105 = importeIVA105
        percibidoIB = importeIB
        exento = importeExento
        total = importeTotal
        soloIVA = ivaSolo
    End Sub

    Public Function codigoProveedor() As String
        If IsNothing(proveedor) Then
            Return ""
        End If
        Return proveedor.id
    End Function
End Class
