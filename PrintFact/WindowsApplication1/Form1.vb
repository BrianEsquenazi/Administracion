Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class frmPrint
    Dim cryRpt As New ReportDocument

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim stConexion As String
        Dim connectionDB As SqlConnection = New SqlConnection()

        'Try
        Me.Visible = False

        stConexion = ""

        'Cargo el reporte.
        cryRpt.Load(Application.StartupPath & "\imprefactura.rpt")

        'Obtengo los datos de conexion desde el crytal report.
        For Each CrTable In cryRpt.Database.Tables
            stConexion = "Server=" & DirectCast(DirectCast(DirectCast(CrTable, CrystalDecisions.CrystalReports.Engine.Table).LogOnInfo, CrystalDecisions.Shared.TableLogOnInfo).ConnectionInfo, CrystalDecisions.Shared.ConnectionInfo).ServerName _
            & ";Database=" & DirectCast(DirectCast(DirectCast(CrTable, CrystalDecisions.CrystalReports.Engine.Table).LogOnInfo, CrystalDecisions.Shared.TableLogOnInfo).ConnectionInfo, CrystalDecisions.Shared.ConnectionInfo).DatabaseName _
            & ";User Id=" & DirectCast(DirectCast(DirectCast(CrTable, CrystalDecisions.CrystalReports.Engine.Table).LogOnInfo, CrystalDecisions.Shared.TableLogOnInfo).ConnectionInfo, CrystalDecisions.Shared.ConnectionInfo).UserID _
            & ";Password=" & DirectCast(DirectCast(DirectCast(CrTable, CrystalDecisions.CrystalReports.Engine.Table).LogOnInfo, CrystalDecisions.Shared.TableLogOnInfo).ConnectionInfo, CrystalDecisions.Shared.ConnectionInfo).UserID & ";Trusted_Connection=True;"
        Next

        'Abro conexion con la base de datos.
        connectionDB.ConnectionString = "Data Source=(LOCAL)\SQLSERVER2008;Initial Catalog=surfactanSA;User ID=usuarioadmin; Password=usuarioadmin" 'stConexion
        connectionDB.Open()

        'Creo adaptador para leer la tabla.
        Dim adp As SqlDataAdapter = New SqlDataAdapter("select * from imprefactura", connectionDB)
        Dim ds As DataSet = New DataSet()
        adp.Fill(ds)

        CrystalReportViewer1.ReportSource = cryRpt
        CrystalReportViewer1.Refresh()

        'cryRpt.PrintToPrinter(1, False, 1, 1)

        Dim CrExportOptions As ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

        'Armo el nombre del archivo.
        CrDiskFileDestinationOptions.DiskFileName = "c:\FcElectronica\FC-0009-" & ds.Tables(0).Rows(0)("numero").ToString() & ".pdf"

        CrExportOptions = cryRpt.ExportOptions
        With CrExportOptions
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
            .DestinationOptions = CrDiskFileDestinationOptions
            .FormatOptions = CrFormatTypeOptions
        End With

        'Exporta y crea el pdf.
        cryRpt.Export()

        'Se envia el mail.
        If setEmailSend("Factura electronia", "Sres: " & ds.Tables(0).Rows(0)("razon").ToString() & vbCrLf & "Por la presente se envia la factura electronica." & vbCrLf & "Saludos" & vbCrLf & "Surfactan SA", _
                     ds.Tables(0).Rows(0)("email").ToString(), "", _
                     CrDiskFileDestinationOptions.DiskFileName, "") = True Then

            'Se cambia el estado.
            Dim Comando As SqlCommand = New SqlCommand("update imprefactura set estado='S'", connectionDB)
            Comando.ExecuteNonQuery()

        End If

        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        'End Try
    End Sub

    Private Function setEmailSend(ByVal sSubject As String, ByVal sBody As String, _
                             ByVal sTo As String, ByVal sCC As String, _
                             ByVal sFilename As String, ByVal sDisplayname As String) As Boolean

        Try
            Dim oApp As Outlook._Application
            oApp = New Outlook.Application

            Dim oMsg As Outlook._MailItem
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)

            oMsg.Subject = sSubject
            oMsg.Body = sBody

            oMsg.To = sTo
            oMsg.CC = sCC


            Dim strS As String = sFilename
            Dim strN As String = sDisplayname
            If sFilename <> "" Then
                Dim sBodyLen As Integer = Int(sBody.Length)
                Dim oAttachs As Outlook.Attachments = oMsg.Attachments
                Dim oAttach As Outlook.Attachment

                oAttach = oAttachs.Add(strS, , sBodyLen, strN)

            End If

            oMsg.Send()

            setEmailSend = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        setEmailSend = False
    End Function
End Class
