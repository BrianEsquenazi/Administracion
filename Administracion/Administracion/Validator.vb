Public Class Validator

    Private warnings As String

    Public Sub New()
        warnings = ""
    End Sub

    Public Function flush()
        If warnings <> "" Then
            MsgBox(warnings, MsgBoxStyle.Exclamation, "No se puede confirmar la operación")
            Return False
        End If
        Return True
    End Function

    Public Sub validarNoVacio(ByVal valor As String, ByVal descripcionCampo As String)
        If valor = "" Then
            agregarWarning("El campo " & descripcionCampo & " no puede ser vacío")
        End If
    End Sub

    Private Sub agregarWarning(ByVal warning As String)
        warnings = warnings & vbCrLf & warning
    End Sub

End Class
