Public Class FormaPago
    Public banco As Integer
    Public tipo, numero, fecha, nombre As String
    Public importe As Double

    Public Sub New(ByVal tipoForma As String, ByVal codBanco As Integer, ByVal numeroCheque As String,
                   ByVal fechaCheque As String, ByVal descripcion As String, ByVal valor As Double)
        tipo = tipoForma
        banco = codBanco
        numero = numeroCheque
        fecha = fechaCheque
        nombre = descripcion
        importe = valor
    End Sub
End Class
