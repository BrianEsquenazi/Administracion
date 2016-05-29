Imports ClasesCompartidas

Public Class DAOTipoCambio

    Public Shared Sub agregarTipoCambio(ByVal cambio As TipoDeCambio)
        SQLConnector.executeProcedure("alta_tipo_cambio", cambio.fecha, cambio.paridad)
    End Sub

    Public Shared Sub eliminarTipoCambio(ByVal fecha As String)
        SQLConnector.executeProcedure("baja_tipo_cambio", fecha)
    End Sub

    Public Shared Function buscarTipoCambioPorFecha(ByVal fecha As String)
        Try
            Dim cambios As New List(Of TipoDeCambio)
            Dim tabla As DataTable
            tabla = SQLConnector.retrieveDataTable("buscar_tipo_cambio_por_fecha", fecha)
            If tabla.Rows.Count > 0 Then
                Return New TipoDeCambio(tabla(0)("fecha"), tabla(0)("Cambio"))
            Else
                Return Nothing
            End If
        Catch e As Exception
            Return Nothing
        End Try
    End Function
End Class
