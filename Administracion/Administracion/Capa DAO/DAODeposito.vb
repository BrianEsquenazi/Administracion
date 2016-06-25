﻿Imports ClasesCompartidas

Public Class DAODeposito

    Public Shared Sub agregarDeposito(ByVal deposito As Deposito, ByVal gridRows As DataGridViewRowCollection)
        deposito.agregarItems(createItems(gridRows))
        agregarDeposito(deposito)
    End Sub

    Public Shared Sub agregarDeposito(ByVal deposito As Deposito, ByVal cheques As List(Of ItemDeposito))
        deposito.agregarItems(cheques)
        agregarDeposito(deposito)
    End Sub

    Private Shared Sub agregarDeposito(ByVal deposito As Deposito)
        For Each item As ItemDeposito In deposito.items
            Dim renglon As String = indiceComoString(item, deposito.items)
            SQLConnector.executeProcedure("alta_deposito", deposito.numero & renglon, deposito.numero, renglon, deposito.banco.id, deposito.fecha, deposito.importeTotal, deposito.fechaAcreditacion, item.tipo, item.numero, item.fecha, item.importe, item.nombre)
        Next
    End Sub

    Private Shared Function indiceComoString(ByVal item As ItemDeposito, ByVal items As List(Of ItemDeposito))
        Dim indice As String
        Dim index As Integer = items.IndexOf(item) + 1
        If index > 9 Then
            indice = index.ToString
        Else
            indice = "0" & index.ToString
        End If
        Return indice
    End Function


    Public Shared Function buscarCheques() As List(Of Cheque)
        Dim cheques As New List(Of Cheque)
        For Each row In SQLConnector.retrieveDataTable("get_cheque_en_cartera").Rows
            cheques.Add(New Cheque(row("Numero2").ToString, row("Fecha2").ToString, row("Importe2"), row("banco2"), row("Clave")))
        Next
        Return cheques
    End Function

    Private Shared Function createItems(ByVal gridRows As DataGridViewRowCollection)
        Dim items As New List(Of ItemDeposito)
        For Each row As DataGridViewRow In gridRows
            If Not row.IsNewRow Then
                items.Add(New Efectivo(row.Cells(0).Value, row.Cells(4).Value))
            End If
        Next
        Return items
    End Function
End Class
