Imports System.Configuration

Public Class Globals
    Shared Function getConnectionString()
        Return ConfigurationManager.ConnectionStrings("administracion").ConnectionString
    End Function
End Class
