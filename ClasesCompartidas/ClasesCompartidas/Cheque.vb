Public Class Cheque
    Implements ItemDeposito

    Public identificacion, fechaCheque, banco, clave As String
    Public valorImporte As Double


    Public Sub New(ByVal nro As String, ByVal fechaSegunCheque As String, ByVal valor As Double, ByVal codBanco As String, ByVal clav As String)
        identificacion = nro
        fechaCheque = fechaSegunCheque
        valorImporte = valor
        banco = codBanco
        clave = clav
    End Sub


    Public Function igualA(ByVal otroCheque As Cheque)
        Return identificacion = otroCheque.identificacion And
        fechaCheque = otroCheque.fechaCheque And
        banco = otroCheque.banco And
        valorImporte = otroCheque.valorImporte And
        clave = otroCheque.clave
    End Function

    Public Overrides Function ToString() As String
        Return fechaCheque & " - " & identificacion & " - " & banco & " - " & valorImporte
    End Function

    Public Function tipo() As Integer Implements ItemDeposito.tipo
        Return 3
    End Function

    Public Function importe() As Double Implements ItemDeposito.importe
        Return valorImporte
    End Function

    Public Function numero() As String Implements ItemDeposito.numero
        Return identificacion
    End Function

    Public Function nombre() As String Implements ItemDeposito.nombre
        Return banco
    End Function

    Public Function fecha() As String Implements ItemDeposito.fecha
        Return fechaCheque
    End Function
End Class
