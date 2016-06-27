Imports ClasesCompartidas

Public Class DAOCtaCteProveedor

    Public Shared Function buscarCuentas() As List(Of CtaCteProveedor)
        Dim cuentas As New List(Of CtaCteProveedor)
        For Each row In SQLConnector.retrieveDataTable("get_carga_intereses").Rows
            cuentas.Add(New CtaCteProveedor(row("FechaOriginal").ToString, row("DesProveOriginal").ToString, row("FacturaOriginal").ToString, row("Cuota").ToString, row("fecha").ToString, row("Saldo"), row("Intereses"), row("IvaIntereses"), row("Referencia").ToString, row("Cuota").ToString, row("NroInterno").ToString))
        Next
        Return cuentas
    End Function
End Class
