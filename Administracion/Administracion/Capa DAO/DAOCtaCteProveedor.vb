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
End Class
