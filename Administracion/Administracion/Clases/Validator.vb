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

    Public Sub validate(ByVal value As String, ByVal type As Integer, ByVal emptyPermitted As Boolean, ByVal description As String)
        Select Case type
            Case ValidatorType.NotEmpty
                validarNoVacio(value, emptyPermitted, description)
            Case ValidatorType.Numeric
                validarNumerico(value, emptyPermitted, description)
            Case ValidatorType.Positive
                validarPositivo(value, emptyPermitted, description)
            Case ValidatorType.PositiveWithMax
                validarPositivo(value, emptyPermitted, description, Double.MaxValue)
            Case ValidatorType.DateFormat
                validarFecha(value, emptyPermitted, description)
        End Select
    End Sub

    Public Function validarNoVacio(ByVal valor As String, ByVal emptyPermitted As Boolean, ByVal descripcionCampo As String)
        If Not emptyPermitted And valor = "" Then
            agregarWarning("El campo " & descripcionCampo & " no puede ser vacío")
            Return False
        Else
            If valor = "" Then : Return True
            Else : Return False
            End If
        End If
    End Function

    Public Sub validarFecha(ByVal valor As String, ByVal emptyPermitted As Boolean, ByVal descripcionCampo As String)
        Dim res As Date
        If Not validarNoVacio(valor, emptyPermitted, descripcionCampo) Then
            If Not Date.TryParse(valor, res) Then
                agregarWarning("El campo " & descripcionCampo & " debe ser una fecha válida")
            End If
        End If
    End Sub

    Public Sub validarPositivo(ByVal valor As String, ByVal emptyPermitted As Boolean, ByVal descripcionCampo As String)
        If validarNumerico(valor, emptyPermitted, descripcionCampo) Then
            If valor < 0 Then
                agregarWarning("El campo " & descripcionCampo & " debe ser positivo")
            End If
        End If
    End Sub

    Public Sub validarPositivo(ByVal valor As String, ByVal emptyPermitted As Boolean, ByVal descripcionCampo As String, ByVal max As Double)
        If validarNumerico(valor, emptyPermitted, descripcionCampo) Then
            If valor > max Or valor < 0 Then
                agregarWarning("El campo " & descripcionCampo & " debe estar entre 0 y " & max)
            End If
        End If
    End Sub

    Public Function validarNumerico(ByVal valor As String, ByVal emptyPermitted As Boolean, ByVal descripcionCampo As String)
        If Not validarNoVacio(valor, emptyPermitted, descripcionCampo) Then
            If Not IsNumeric(valor) Then
                agregarWarning("El campo " & descripcionCampo & " debe ser numérico")
                Return False
            End If
            Return True
        End If
        Return False
    End Function

    Private Sub agregarWarning(ByVal warning As String)
        warnings = warnings & vbCrLf & warning
    End Sub

End Class
