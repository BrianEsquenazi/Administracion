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

    Public Sub validarFecha(ByVal valor As String, ByVal descripcionCampo As String)
        Dim res As Date
        validarNoVacio(valor, descripcionCampo)
        If Not Date.TryParse(valor, res) Then
            agregarWarning("El campo " & descripcionCampo & " debe ser una fecha válida")
        End If
    End Sub

    Public Sub validarNoVacio(ByVal valor As String, ByVal descripcionCampo As String)
        If valor = "" Then
            agregarWarning("El campo " & descripcionCampo & " no puede ser vacío")
        End If
    End Sub

    Public Sub validarPositivo(ByVal valor As String, ByVal descripcionCampo As String, ByVal max As Double)
        If Not validarNumerico(valor, descripcionCampo) OrElse (valor > max Or valor < 0) Then
            agregarWarning("El campo " & descripcionCampo & " debe estar entre 0 y " & max)
        End If
    End Sub

    Public Function validarNumerico(ByVal valor As String, ByVal descripcionCampo As String)
        validarNoVacio(valor, descripcionCampo)
        If Not IsNumeric(valor) Then
            agregarWarning("El campo " & descripcionCampo & " debe ser numérico")
            Return False
        End If
        Return True
    End Function

    Private Sub agregarWarning(ByVal warning As String)
        warnings = warnings & vbCrLf & warning
    End Sub

End Class
