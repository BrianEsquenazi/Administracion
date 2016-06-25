Public Class Cheque
    Public numero, fecha, banco, clave As String
    Public importe As Double


    Public Sub New(ByVal nro As String, ByVal fechaCheque As String, ByVal valor As Double, ByVal codBanco As String, ByVal clav As String)
        numero = nro
        fecha = fechaCheque
        importe = valor
        banco = codBanco
        clave = clav
    End Sub


    Public Function igualA(ByVal otroCheque As Cheque)
        Return numero = otroCheque.numero And
        fecha = otroCheque.fecha And
        banco = otroCheque.banco And
        importe = otroCheque.importe And
        clave = otroCheque.clave
    End Function

    Public Overrides Function ToString() As String
        Return fecha & " - " & numero & " - " & banco & " - " & importe
    End Function
End Class
