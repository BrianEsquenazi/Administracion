Imports ClasesCompartidas
Imports System.IO

Public Class ProcesoSifere

    Dim nombreArchivo As String

    Dim ordDesde As String
    Dim ordHasta As String

    Dim WCodigo As String
    Dim WNumero As String
    Dim WPunto As String
    Dim WImpoIva As String
    Dim WCuit As String
    Dim WDespacho As String
    Dim WBanco As Integer
    Dim WFecha As String
    Dim WLetra As String

    Private Sub ProcesoSifere_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtDesde.Text = "  /  /    "
        txtHasta.Text = "  /  /    "

        TipoProceso.Items.Clear()
        TipoProceso.Items.Add("No Aduana")
        TipoProceso.Items.Add("Aduana")
        TipoProceso.SelectedIndex = 0

    End Sub

    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click

        'WBanco = 12
        'Dim campobanco As Banco = DAOBanco.buscarBancoPorCodigo(WBanco)
        'wnombre = campobanco.nombre

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            nombreArchivo = FolderBrowserDialog1.SelectedPath
        End If

        If Trim(nombreArchivo) <> "" Then
            nombreArchivo = nombreArchivo + "\" + txtNombre.Text + ".txt"
        End If

        File.Create(nombreArchivo).Dispose()

        Dim escritor As New System.IO.StreamWriter(nombreArchivo)

        ordDesde = ordenaFecha(txtDesde.Text)
        ordHasta = ordenaFecha(txtHasta.Text)

        Select Case TipoProceso.SelectedIndex
            Case 0
                Dim tabla As DataTable
                tabla = SQLConnector.retrieveDataTable("procesoSifere", ordDesde, ordHasta)

                For Each row As DataRow In tabla.Rows

                    Dim CamposImputac As New Imputac(row.Item(0).ToString, row.Item(1), row.Item(2).ToString, row.Item(3).ToString, row.Item(4).ToString, row.Item(5).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString)

                    If ProveedorAduana(CamposImputac.proveedor) = "N" Then

                        WCodigo = CodigoSifere(CamposImputac.cuenta)
                        If Val(WCodigo) <> 0 Then

                            WPunto = ceros(CamposImputac.punto, 4)
                            WNumero = ceros(CamposImputac.numero, 8)
                            WImpoIva = ceros(formatonumerico(redondeo(CamposImputac.debito), "########0.#0", ","), 11)
                            WCuit = leederecha(CamposImputac.cuit, 13)
                            WFecha = CamposImputac.fechaord
                            WLetra = CamposImputac.letra

                            escritor.Write(WCodigo + WCuit + WFecha + WPunto + WNumero + "F" + WLetra + WImpoIva + vbCrLf)

                        End If

                    End If

                Next

                escritor.Close()

                MsgBox("Proceso Finalizado de Sifere (No aduana)", MsgBoxStyle.Information)


            Case Else
                Dim tabla As DataTable
                tabla = SQLConnector.retrieveDataTable("procesoSifere", ordDesde, ordHasta)

                For Each row As DataRow In tabla.Rows

                    Dim CamposImputac As New Imputac(row.Item(0).ToString, row.Item(1), row.Item(2).ToString, row.Item(3).ToString, row.Item(4).ToString, row.Item(5).ToString, row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString)

                    If ProveedorAduana(CamposImputac.proveedor) = "S" Then

                        WCodigo = CodigoSifere(CamposImputac.cuenta)
                        If Val(WCodigo) <> 0 Then

                            WPunto = ceros(CamposImputac.punto, 4)
                            WNumero = ceros(CamposImputac.numero, 8)

                            WDespacho = agregaespacios(CamposImputac.despacho, 20)
                            WImpoIva = ceros(formatonumerico(redondeo(CamposImputac.debito), "########0.#0", ","), 11)
                            WCuit = leederecha(CamposImputac.cuit, 13)
                            WFecha = CamposImputac.fechaord

                            escritor.Write(WCodigo + WCuit + WFecha + WNumero + WImpoIva + vbCrLf)

                        End If

                    End If

                Next

                escritor.Close()

                MsgBox("Proceso Finalizado de Sifere Aduana", MsgBoxStyle.Information)

        End Select

    End Sub

    Private Sub txtDesde_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDesde.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtHasta.Focus()
        End If
        'If e.KeyCode = Keys.Escape Then
        '    txtDesde.Text = "  /  /    "
        'End If
    End Sub

    Private Sub txtHasta_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDesde.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtNombre.Focus()
        End If
    End Sub

    Private Sub txtnombre_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDesde.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtDesde.Focus()
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Hide()
        MenuPrincipal.Show()
    End Sub

    Private Sub CustomLabel4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomLabel4.Click

    End Sub
    Private Sub TipoProceso_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TipoProceso.SelectedIndexChanged

    End Sub
    Private Sub txtNombre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNombre.TextChanged

    End Sub
    Private Sub CustomLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomLabel3.Click

    End Sub
    Private Sub txtDesde_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles txtDesde.MaskInputRejected

    End Sub
    Private Sub CustomLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomLabel2.Click

    End Sub
    Private Sub txtHasta_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles txtHasta.MaskInputRejected

    End Sub
    Private Sub CustomLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomLabel1.Click

    End Sub
End Class