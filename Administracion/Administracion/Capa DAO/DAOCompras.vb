Imports ClasesCompartidas

Public Class DAOCompras

    Public Shared Sub agregarCompra(ByVal compra As Compra)
        SQLConnector.executeProcedure("alta_iva_compra", compra.nroInterno, compra.codigoProveedor, compra.tipoDocumento, compra.letra, compra.punto, compra.numero, compra.fechaEmision,
                                      compra.fechaVto1, compra.fechaVto2, compra.fechaIVA, compra.neto, compra.iva21, compra.ivaRG, compra.iva27,
                                      compra.percibidoIB, compra.exento, compra.tipoPago, compra.tipoDocumentoDescripcion, compra.paridad,
                                      compra.formaPago, compra.proveedor.cai, compra.proveedor.vtoCAI, compra.iva105, compra.despacho, compra.remito, compra.soloIVA)
    End Sub

    Public Shared Function siguienteNumeroDeInterno() As Long
        Return SQLConnector.retrieveDataTable("get_siguiente_numero_interno").Rows(0)(0)
    End Function
End Class
