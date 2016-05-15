Imports System.Configuration

Public Class Globals
    Public Shared empresa As String

    Shared Function getConnectionString() As String
        If empresa Is Nothing Then
            Throw New ApplicationException("No fue seleccionada la empresa")
        Else
            Return ConfigurationManager.ConnectionStrings(empresa).ConnectionString
        End If
    End Function

    Public Shared Function connectionStringNames() As List(Of String)
        Dim connections As New List(Of String)
        For index As Integer = 1 To ConfigurationManager.ConnectionStrings.Count - 1
            connections.Add(ConfigurationManager.ConnectionStrings(index).Name)
        Next
        Return connections
    End Function
End Class
