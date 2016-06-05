Imports ClasesCompartidas

Public Class DAOProveedor

    Public Function listado_provincias() As List(Of Provincia)
        Dim provincias As New List(Of Provincia)
        For Each row In SQLConnector.retrieveDataTable("listado_provincias").Rows
            provincias.Add(New Provincia(row("codigo"), row("nombre")))
        Next
        Return provincias
    End Function


End Class
