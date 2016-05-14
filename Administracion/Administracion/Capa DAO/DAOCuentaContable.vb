Imports ClasesCompartidas

Public Class DAOCuentaContable

    Public Shared Sub agregarCuentaContable(ByVal cuenta As CuentaContable)
        SQLConnector.executeProcedure("alta_cuenta", cuenta.id, cuenta.descripcion, 1, 1)
    End Sub

    Public Shared Sub eliminarCuentaContable(ByVal cuenta As CuentaContable)
        SQLConnector.executeProcedure("baja_cuenta", cuenta.id)
    End Sub

    Public Shared Function buscarCuentaContablePorDescripcion(ByVal descripcion As String)
        Dim cuentas As New List(Of CuentaContable)
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta", descripcion)
        For Each cuenta As DataRow In tabla.Rows
            cuentas.Add(New CuentaContable(cuenta("cuenta"), cuenta("descripcion")))
        Next
        Return cuentas
    End Function


End Class
