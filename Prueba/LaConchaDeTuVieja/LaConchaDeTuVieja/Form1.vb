Imports ClasesCompartidas
Imports System.Data.SqlClient
Imports System.Configuration

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conexion As SqlConnection = New SqlConnection
        Dim comando As SqlCommand = New SqlCommand

        Try
            SQLConnector.conexionSql(conexion, comando)
            'Creo un procedure mágico que suma 3 a cualquier valor que le pase
            comando.CommandText = "IF EXISTS(SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID('PROCEDURE_PRUEBA'))" & vbCrLf & _
                "DROP PROCEDURE PROCEDURE_PRUEBA"
            comando.ExecuteNonQuery()
            comando.CommandText = "CREATE PROCEDURE PROCEDURE_PRUEBA(@VAL INT) AS" & vbCrLf & _
                "BEGIN" & vbCrLf & _
                "DECLARE @CANTIDAD INT" & vbCrLf & _
                "SET @CANTIDAD = @VAL + 3" & vbCrLf & _
                "RETURN @CANTIDAD" & vbCrLf & _
                "END"
            comando.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            MsgBox(SQLConnector.executeProcedureWithReturnValue("PROCEDURE_PRUEBA", Val(TextBox1.Text)))
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Fijate si buildeaste el proyecto de clases compartidas y después buildeaste éste. De nada ;)")
        End Try
    End Sub
End Class
