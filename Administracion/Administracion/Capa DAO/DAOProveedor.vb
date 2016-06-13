Imports ClasesCompartidas

Public Class DAOProveedor

    Public Shared Function listarProvincias() As List(Of Provincia)
        Dim provincias As New List(Of Provincia)
        For Each row In SQLConnector.retrieveDataTable("get_provincias").Rows
            provincias.Add(New Provincia(row("codigo"), row("nombre")))
        Next
        Return provincias
    End Function

    Public Shared Function buscarProveedorPorNombre(ByVal nombre As String)
        Dim proveedores As New List(Of Proveedor)
        Dim tabla As DataTable
        tabla = SQLConnector.retrieveDataTable("buscar_proveedor_por_nombre", nombre)
        For Each proveedor As DataRow In tabla.Rows
            proveedores.Add(New Proveedor(proveedor("codigo"), proveedor("nombre")))
        Next
        Return proveedores
    End Function

    Public Shared Function buscarProveedorPorCodigo(ByVal codigo As String)
        Try
            Dim tabla As DataTable
            tabla = SQLConnector.retrieveDataTable("buscar_proveedor_por_codigo", codigo)
            If tabla.Rows.Count > 0 Then
                Return New Proveedor(codigo, tabla(0)("nombre"))
            Else
                Return Nothing
            End If
        Catch e As Exception
            Return Nothing
        End Try
    End Function

End Class
