Public Class Cheque
    Public tipo, numero, fecha As String
    Public importe As Double


    Public Sub New(ByVal unTipo As Integer, ByVal nro As String, ByVal fechaCheque As String, ByVal valor As Double)
        tipo = unTipo
        numero = nro
        fecha = fechaCheque
        importe = valor
    End Sub

End Class
