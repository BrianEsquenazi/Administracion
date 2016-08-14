Public Class LeeDepositos

    Public fechaord, deposito, tipo2, numero2, Fecha As String
    Public importe2 As Double
    Public Banco As Integer


    Public Sub New(ByVal fechaord2 As String, ByVal deposito2 As String, ByVal tipo22 As String, ByVal numero22 As String,
                   ByVal Fecha2 As String, ByVal importe22 As Double, ByVal Banco2 As Integer)
        fechaord = fechaord2
        deposito = deposito2
        tipo2 = tipo22
        numero2 = numero22
        Fecha = fecha2
        importe2 = importe22
        Banco = Banco2
    End Sub

End Class
