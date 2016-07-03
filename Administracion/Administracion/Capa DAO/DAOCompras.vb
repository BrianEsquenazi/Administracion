Imports ClasesCompartidas

Public Class DAOCompras

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
