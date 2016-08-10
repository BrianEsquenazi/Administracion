﻿Imports ClasesCompartidas

Public Class DAORecibo

    Public Shared Function existeReciboProvisorio(ByVal codigo As String)
        Return SQLConnector.checkIfExists("get_recibo_provisorio", codigo)
    End Function

    Public Shared Function existeRecibo(ByVal codigo As String)
        Return SQLConnector.checkIfExists("get_recibo", codigo)
    End Function

    Public Shared Sub agregarRecibo(ByVal recibo As Recibo)
        Dim renglon As Integer = 1
        For Each formaPago As FormaPago In recibo.formasPago
            SQLConnector.executeProcedure("alta_recibo_forma_pago", recibo.codigo, ceros(renglon, 2), recibo.codigoCliente, recibo.fecha,
                                        recibo.retGanancias, recibo.retIVA, recibo.retIB, recibo.retSuss, ceros(formaPago.tipo, 2),
                                        formaPago.numero, formaPago.fecha, formaPago.nombre, formaPago.importe, recibo.total,
                                        recibo.paridad, recibo.observaciones, recibo.tipo, recibo.codigoCuenta)
            renglon += 1
        Next
        renglon = 1
        For Each pago As Pago In recibo.pagos
            SQLConnector.executeProcedure("alta_recibo_pago", recibo.codigo, ceros(renglon, 2), recibo.codigoCliente, recibo.fecha,
                                        recibo.retGanancias, recibo.retIVA, recibo.retIB, recibo.retSuss, ceros(pago.tipo, 2),
                                        pago.letra, pago.punto, pago.numero, pago.importe, recibo.total,
                                        recibo.paridad, recibo.observaciones, recibo.tipo, recibo.codigoCuenta)
            renglon += 1
        Next
    End Sub

    Public Shared Sub actualizarReciboProvisorio(ByVal codProvisorio As String, ByVal codRecibo As String)
        SQLConnector.executeProcedure("actualizar_recibo_provisorio", codProvisorio, codRecibo)
    End Sub

    Public Shared Sub agregarReciboProvisorio(ByVal recibo As ReciboProvisorio)
        Dim renglon As Integer = 1
        For Each formaPago As FormaPago In recibo.formasPago
            SQLConnector.executeProcedure("alta_recibo_provisorio", recibo.codigo, ceros(renglon, 2), recibo.codigoCliente, recibo.fecha,
                                        recibo.retGanancias, recibo.retIVA, recibo.retIB, recibo.retSuss, ceros(formaPago.tipo, 2),
                                        formaPago.numero, formaPago.fecha, formaPago.nombre, formaPago.importe, recibo.total,
                                        recibo.paridad)
            renglon += 1
        Next
    End Sub

    Private Shared Function crearFormaPago(ByVal rowA As DataRow)
        Return New FormaPago(rowA("Tipo2").ToString, 0, rowA("Numero2").ToString, rowA("Fecha2").ToString, rowA("banco2").ToString,
                            asDouble(rowA("Importe2")))
    End Function

    Private Shared Function crearPago(ByVal rowA As DataRow)
        Return New Pago(rowA("Tipo1").ToString, rowA("Letra1").ToString, rowA("Punto1").ToString, rowA("Numero1").ToString,
                        "", asDouble(rowA("Importe1")))
    End Function

    Public Shared Function buscarRecibo(ByVal codRecibo As String)
        Try
            Dim tabla As DataTable = SQLConnector.retrieveDataTable("get_recibo", codRecibo)
            Dim row As DataRow = tabla.Rows(0)
            Dim recibo As New Recibo(row("Recibo").ToString, CustomConvert.asTextDate(row("Fecha").ToString), DAOCliente.buscarClientePorCodigo(row("Cliente").ToString),
                                        asDouble(row("RetGanancias")), asDouble(row("RetOtra")), asDouble(row("RetIva")), asDouble(row("RetSuss")), asDouble(row("Paridad")),
                                        0, DAOCuentaContable.buscarCuentaContablePorCodigo(row("Cuenta").ToString), row("Observaciones").ToString, CustomConvert.toIntOrZero(row("TipoRec")))
            Dim formasPago As New List(Of FormaPago)
            Dim pagos As New List(Of Pago)
            For Each rowA As DataRow In tabla.Rows
                If rowA("TipoReg").ToString = "2" Then
                    formasPago.Add(crearFormaPago(rowA))
                Else
                    pagos.Add(crearPago(rowA))
                End If
            Next
            recibo.formasPago = formasPago
            recibo.pagos = pagos
            Return recibo
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function buscarReciboProvisorio(ByVal codRecibo As String)
        Try
            Dim tabla As DataTable = SQLConnector.retrieveDataTable("get_recibo_provisorio", codRecibo)
            Dim row As DataRow = tabla.Rows(0)
            Dim recibo As ReciboProvisorio = New ReciboProvisorio(row("Recibo").ToString, CustomConvert.asTextDate(row("Fecha").ToString), DAOCliente.buscarClientePorCodigo(row("Cliente").ToString),
                                        asDouble(row("RetGanancias")), asDouble(row("RetOtra")), asDouble(row("RetIva")), asDouble(row("RetSuss")), asDouble(row("Paridad")),
                                        0)
            Dim formasPago As New List(Of FormaPago)
            For Each rowA As DataRow In tabla.Rows
                If rowA("TipoReg").ToString = "2" Then
                    formasPago.Add(crearFormaPago(rowA))
                End If
            Next
            recibo.formasPago = formasPago
            Return recibo
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function asDouble(ByVal value)
        Return CustomConvert.toDoubleOrZero(value)
    End Function
End Class
