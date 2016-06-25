Imports ClasesCompartidas

Public Class DAODeposito

    Public Shared Function buscarCheques() As List(Of Cheque)
        Dim cheques As New List(Of Cheque)
        For Each row In SQLConnector.retrieveDataTable("get_cheque_en_cartera").Rows
            cheques.Add(New Cheque(row("Numero2").ToString, row("Fecha2").ToString, row("Importe2"), row("banco2"), row("Clave")))
        Next
        Return cheques
    End Function

End Class
