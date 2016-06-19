Imports ClasesCompartidas
Imports System.IO

Public Class ProcesoRetencionesPagos

    Dim nombreArchivo As String

    Dim ordDesde As String
    Dim ordHasta As String


    Dim WCuit As String
    Dim WFecha As String
    Dim WSucursal As String
    Dim WOrden As String
    Dim WImporte As String


    Private Sub ProcesoRetencionesPagos_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtDesde.Text = "  /  /    "
        txtHasta.Text = "  /  /    "

        TipoProceso.Items.Clear()
        TipoProceso.Items.Add("Ingresos Brutos")
        TipoProceso.Items.Add("Ganancias")

        TipoProceso.SelectedIndex = 0

    End Sub


    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            nombreArchivo = FolderBrowserDialog1.SelectedPath
        End If

        REM XNombre = WDir + "\AR-30610524598-" + Nombre.text + "-6-LOTE1.txt"
        If Trim(nombreArchivo) <> "" Then
            nombreArchivo = nombreArchivo + "\AR-30549165083-" + txtNombre.Text + "-6-LOTE1.txt"
        End If

        File.Create(nombreArchivo).Dispose()

        Dim escritor As New System.IO.StreamWriter(nombreArchivo)

        ordDesde = ordenaFecha(txtDesde.Text)
        ordHasta = ordenaFecha(txtHasta.Text)

        Select Case TipoProceso.SelectedIndex
            Case 0
                Dim tabla As DataTable
                tabla = SQLConnector.retrieveDataTable("procesoReteIb", ordDesde, ordHasta)

                For Each row As DataRow In tabla.Rows

                    Dim CamposReteIb As New ProcesoReteIb(row.Item(0).ToString, row.Item(1), row.Item(2).ToString, row.Item(3).ToString, row.Item(4).ToString, row.Item(5).ToString, row.Item(6).ToString)

                    WCuit = sacaguiones(CamposReteIb.cuit)
                    WFecha = CamposReteIb.fecha
                    WSucursal = "0001"
                    WImporte = ceros(formatonumerico(redondeo(CamposReteIb.retotra), "########0.#0", "."), 11)

                    escritor.Write(WCuit + WFecha + WSucursal + WOrden + WImporte + "A" + vbCrLf)

                Next

                escritor.Close()

                MsgBox("Proceso Finalizado", MsgBoxStyle.Information)


            Case Else

        End Select

    End Sub

    Private Sub txtDesde_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            txtHasta.Focus()
        End If
        'If e.KeyCode = Keys.Escape Then
        '    txtDesde.Text = "  /  /    "
        'End If
    End Sub

    Private Sub txtHasta_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
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



   
End Class