Imports ClasesCompartidas

Public Class DAOCierreMes

    Public Shared Sub agregarCierremes(ByVal cuenta As CuentaContable)
        SQLConnector.executeProcedure("alta_cierre", cuenta.id, cuenta.descripcion, 1, 1)
    End Sub

    Public Shared Function buscarCierre(ByVal codigo As String)
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_cuenta_por_codigo", codigo)
        If tabla.Rows.Count > 0 Then
            Return New CuentaContable(tabla(0)("cuenta"), tabla(0)("descripcion"))
        Else
            Return Nothing
        End If
    End Function

End Class
