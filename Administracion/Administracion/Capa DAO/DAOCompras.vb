Imports ClasesCompartidas

Public Class DAOCompras

    Public Shared Sub agregarCompra(ByVal compra As Compra)
        'SQLConnector.executeProcedure("alta_compra", compra)
    End Sub

    Public Shared Function siguienteNumeroDeInterno() As Long
        Return SQLConnector.retrieveDataTable("get_siguiente_numero_interno").Rows(0)(0)
    End Function
End Class
