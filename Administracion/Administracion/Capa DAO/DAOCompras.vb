Imports ClasesCompartidas

Public Class DAOCompras

    Public Shared Function mesAbierto(ByVal fecha As String)
        Try
            Dim mes As Integer = Convert.ToDateTime(fecha).Month
            Return True
            'Return SQLConnector.checkIfExists("get_mes_abierto", mes)
        Catch ex As FormatException
            Return True
        End Try
    End Function

End Class
