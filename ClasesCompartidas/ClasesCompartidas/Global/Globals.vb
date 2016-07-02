﻿Imports System.Configuration

Public Class Globals
    Public Shared empresa As String

    Public Shared Function reportPathWithName(ByVal reportName As String)
        Dim path As String
        Try
            path = ConfigurationManager.AppSettings("reportsLocation")
        Catch ex As Exception
            Throw New ApplicationException("Problemas obteniendo la variable de entorno 'reportsLocation' del app.config")
        End Try

        If path = "" Then
            Throw New ApplicationException("La variable de entorno 'reportsLocation' no está definida en el app.config")
        End If
        Return path & reportName
    End Function

    Public Shared Function connectionWorksFor(ByVal connectionName As String)
        Try
            Return SQLConnector.tryConnection(ConfigurationManager.ConnectionStrings(connectionName).ConnectionString)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function getConnectionString() As String
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
