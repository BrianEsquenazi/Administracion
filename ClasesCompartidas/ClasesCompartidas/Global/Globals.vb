Imports System.Configuration

Public Class Globals
    Shared Function getConnectionString()
        Return ConfigurationManager.ConnectionStrings("GD1C2015").ConnectionString
    End Function
End Class
