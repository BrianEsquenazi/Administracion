﻿Public Class CuentaContable
    Public id As String
    Public descripcion As String

    Public Sub New(ByVal identificador As String, ByVal desc As String)
        id = identificador
        descripcion = desc
    End Sub

    Public Overrides Function ToString() As String
        Return descripcion
    End Function
End Class
