Imports ClasesCompartidas

Public Class DAOCompras

    Private Shared Function crearCompra(ByVal row As DataRow)
        Dim compra As Compra
        compra = New Compra(asInt(row("NroInterno")), DAOProveedor.buscarProveedorPorCodigo(row("Proveedor").ToString), asInt(row("Tipo")), "", asInt(row("Pago")), asInt(row("Contado")), row("Letra").ToString,
                                row("Punto").ToString, row("Numero").ToString, asDate(row("Fecha")), asDate(row("Periodo")), asDate(row("Vencimiento")), asDate(row("Vencimiento1")), asDouble(row("Paridad")),
                                asDouble(row("Neto")), asDouble(row("Iva21")), asDouble(row("Iva5")), asDouble(row("Iva27")), asDouble(row("Ib")), asDouble(row("Exento")), asDouble(row("Iva105")),
                                0, asBool(row("SoloIva")), row("Remito").ToString, row("Despacho").ToString)
            Return compra
    End Function

    Private Shared Function asDate(ByVal value)
        Return CustomConvert.asTextDate(value.ToString)
    End Function

    Private Shared Function asInt(ByVal value)
        Return CustomConvert.toIntOrZero(value.ToString)
    End Function

    Private Shared Function asDouble(ByVal value)
        Return CustomConvert.toDoubleOrZero(value.ToString)
    End Function

    Private Shared Function asBool(ByVal value)
        Return CustomConvert.toBoolOrFalse(value)
    End Function

    Public Shared Function buscarCompraPorCodigo(ByVal codigo As String)
        Dim row As DataRow
        Try
            row = SQLConnector.retrieveDataTable("get_compra_por_codigo", CustomConvert.toIntOrZero(codigo)).Rows(0)
        Catch ex As Exception
            Return Nothing
        End Try
        Return crearCompra(row)
    End Function

    Public Shared Sub agregarCompra(ByVal compra As Compra)
        SQLConnector.executeProcedure("alta_iva_compra", compra.nroInterno, compra.codigoProveedor, compra.tipoDocumento, compra.letra, compra.punto, compra.numero, compra.fechaEmision,
                                      compra.fechaVto1, compra.fechaVto2, compra.fechaIVA, compra.neto, compra.iva21, compra.ivaRG, compra.iva27,
                                      compra.percibidoIB, compra.exento, compra.tipoPago, compra.tipoDocumentoDescripcion, compra.paridad,
                                      compra.formaPago, compra.proveedor.cai, compra.proveedor.vtoCAI, compra.iva105, compra.despacho, compra.remito, compra.soloIVA)
        agregarImputaciones(compra.imputaciones)
    End Sub

    Private Shared Sub agregarImputaciones(ByVal imputaciones As List(Of Imputac))
        imputaciones.ForEach(Sub(imputacion) SQLConnector.executeProcedure("alta_imputacion", imputacion.clave, imputacion.tipoMovimiento, imputacion.proveedor, imputacion.tipoComprobante,
                                                                           imputacion.letra, imputacion.punto, imputacion.numero, imputacion.renglon,
                                                                           imputacion.fechaord, "", imputacion.cuenta, imputacion.debito, imputacion.credito, imputacion.nrointerno))
    End Sub

    Public Shared Function siguienteNumeroDeInterno() As Long
        Return SQLConnector.retrieveDataTable("get_siguiente_numero_interno").Rows(0)(0)
    End Function
End Class
