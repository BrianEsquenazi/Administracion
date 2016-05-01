Public Class TipoDeCambio
    Private fecha As Date
    Private valor As Double

    Public Sub New(ByVal dia As Date, ByVal precio As Double)
        fecha = dia
        valor = precio
    End Sub

End Class
