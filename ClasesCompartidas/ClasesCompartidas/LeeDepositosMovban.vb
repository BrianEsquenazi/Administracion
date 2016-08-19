Public Class LeeDepositosMovban

    Public fechaord, Fecha, Acredita, AcreditaOrd, Deposito As String
    Public importe As Double
    Public Banco, Renglon As Integer


    Public Sub New(ByVal Banco2 As String, ByVal deposito2 As String, ByVal Renglon2 As String,
                   ByVal Fecha2 As String, ByVal FechaOrd2 As String, ByVal acredita2 As String, ByVal Acreditaord2 As String,
                   ByVal importe2 As Double)

        Banco = Banco2
        Deposito = deposito2
        Renglon = Renglon2
        Fecha = Fecha2
        fechaord = FechaOrd2
        Acredita = acredita2
        AcreditaOrd = Acreditaord2
        importe = importe2

    End Sub

End Class
