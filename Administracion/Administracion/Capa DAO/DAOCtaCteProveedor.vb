Imports ClasesCompartidas

Public Class DAOCtaCteProveedor


    Public Shared Sub modificarCuentaSi(ByVal cuenta As DataGridViewRow)
        Dim saldoNuevo, intereses, iva As Decimal

        intereses = Convert.ToDecimal(cuenta.Cells("Intereses").Value)
        iva = Convert.ToDecimal(cuenta.Cells("IvaIntereses").Value)

        If (intereses <> 0 Or iva <> 0) Then

            saldoNuevo = Convert.ToDecimal(cuenta.Cells("Saldo").Value) + intereses + iva
            SQLConnector.executeProcedure("modificar_carga_intereses", cuenta.Cells("clave").Value, saldoNuevo, intereses, iva, cuenta.Cells("Referencia").Value)
        End If
    End Sub

    Public Shared Function buscarCuentas() As List(Of CtaCteProveedor)
        Dim cuentas As New List(Of CtaCteProveedor)
        For Each row In SQLConnector.retrieveDataTable("get_carga_intereses").Rows
            cuentas.Add(New CtaCteProveedor(row("FechaOriginal").ToString, row("DesProveOriginal").ToString, row("FacturaOriginal").ToString, row("Cuota").ToString, row("fecha").ToString, row("Saldo"), row("Intereses"), row("IvaIntereses"), row("Referencia").ToString, row("Clave").ToString, row("NroInterno").ToString))
        Next
        Return cuentas
    End Function

    Public Shared Function cuentasSinSaldar(ByVal proveedor As Proveedor) As List(Of DetalleCompraCuentaCorriente)
        Dim cuentas As New List(Of DetalleCompraCuentaCorriente)
        For Each row In SQLConnector.retrieveDataTable("get_cuentas_sin_saldar", proveedor.id).Rows
            cuentas.Add(New DetalleCompraCuentaCorriente(row("Impre").ToString, row("Punto").ToString, row("Numero").ToString,
                                                         row("Letra").ToString, row("fecha").ToString, row("Saldo"),
                                                         row("Total"), row("NroInterno").ToString, proveedor))
        Next
        Return cuentas
    End Function


    'Public Shared Function buscardeuda(ByVal proveedor As String)
    '    Dim ctacteprv As New List(Of CtaCteProveedoresDeuda)
    '    For Each row In SQLConnector.retrieveDataTable("buscar_Cuenta_Corriente_Proveedores_deuda", proveedor).Rows
    '        ctacteprv.Add(New CtaCteProveedoresDeuda(row("Tipo").ToString, row("letra").ToString, row("punto").ToString, row("numero").ToString, row("total"), row("saldo"), row("fecha").ToString, row("vencimiento").ToString))
    '    Next
    '    Return ctacteprv
    'End Function

End Class
