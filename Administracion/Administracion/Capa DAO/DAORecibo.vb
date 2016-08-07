Imports ClasesCompartidas

Public Class DAORecibo

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
                    formasPago.Add(New FormaPago(rowA("Tipo2").ToString, 0, rowA("Numero2").ToString, rowA("Fecha2").ToString, rowA("banco2").ToString,
                                                 asDouble(rowA("Importe2"))))
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
