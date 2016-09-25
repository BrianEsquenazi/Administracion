Imports ClasesCompartidas
Imports System.IO

Public Class ListadoCuentaCorrienteProveedoresSelectivo

    Dim varRenglon As Integer
    Dim varTotal, varSaldo, varTotalUs, varSaldoUs, varSaldoOriginal, varDife, varParidad, varParidadTotal As Double
    Dim varPago As Integer

    Private Sub ListadoCuentaCorrienteProveedoresSelectivo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDesdeProveedor.Text = ""
        txtFechaEmision.Text = "  /  /    "
        varRenglon = 0
        opcPantalla.Checked = False
        opcImpesora.Checked = True
    End Sub

    Private Sub txtfechaemision_KeyPress(ByVal sender As Object, _
               ByVal e As System.Windows.Forms.KeyPressEventArgs) _
               Handles txtFechaEmision.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If ValidaFecha(txtFechaEmision.Text) = "S" Then

                Dim CampoTipoCambio As TipoDeCambio = DAOTipoCambio.buscarTipoCambioPorFecha(txtFechaEmision.Text)
                If IsNothing(CampoTipoCambio) Then
                    MsgBox("Paridad Inexistente")
                    txtFechaEmision.Focus()
                Else
                    varparidadtotal = CampoTipoCambio.paridad
                    txtDesdeProveedor.Focus()
                End If
            End If
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtFechaEmision.Text = "  /  /    "
            Me.txtFechaEmision.SelectionStart = 0
        End If
    End Sub

    Private Sub txtdesdeproveedor_KeyPress(ByVal sender As Object, _
                    ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                    Handles txtDesdeProveedor.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            ' DADA que no rompa cuando el codigo no existe y usar la funcion "ceros" para completar??
            Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtDesdeProveedor.Text)
            If IsNothing(CampoProveedor) Then
                MsgBox("Proveedor incorrecto")
                txtDesdeProveedor.Focus()
            Else
                GRilla.Rows.Add()
                GRilla.Item(0, varRenglon).Value = txtDesdeProveedor.Text
                GRilla.Item(1, varRenglon).Value = CampoProveedor.razonSocial
                varRenglon = varRenglon + 1
                GRilla.CurrentCell = GRilla(0, 0)

                txtDesdeProveedor.Text = ""
                txtRazon.Text = ""
                txtDesdeProveedor.Focus()
            End If

        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtRazon.Focus()
        End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
        MenuPrincipal.Show()
    End Sub

    Private Sub btnConsulta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsulta.Click

        lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre("")

        txtAyuda.Text = ""
        txtAyuda.Visible = True
        lstAyuda.Visible = True

        txtAyuda.Focus()

    End Sub

    Private Sub txtAyuda_KeyPress(ByVal sender As Object, _
                   ByVal e As System.Windows.Forms.KeyPressEventArgs) _
                   Handles txtAyuda.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            lstAyuda.DataSource = DAOProveedor.buscarProveedorPorNombre(txtAyuda.Text)
        ElseIf e.KeyChar = Convert.ToChar(Keys.Escape) Then
            e.Handled = True
            txtAyuda.Text = ""
        End If
    End Sub

    Private Sub mostrarProveedor(ByVal proveedor As Proveedor)
        txtDesdeProveedor.Text = proveedor.id
        Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(txtDesdeProveedor.Text)
        If IsNothing(CampoProveedor) Then
            MsgBox("Proveedor incorrecto")
            txtDesdeProveedor.Focus()
        Else
            GRilla.Rows.Add()
            GRilla.Item(0, varRenglon).Value = txtDesdeProveedor.Text
            GRilla.Item(1, varRenglon).Value = CampoProveedor.razonSocial
            varRenglon = varRenglon + 1
            GRilla.CurrentCell = GRilla(0, 0)

            txtDesdeProveedor.Text = ""
            txtRazon.Text = ""
            txtAyuda.Text = ""
            txtAyuda.Focus()
        End If
    End Sub

    Private Sub lstAyuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAyuda.Click
        mostrarProveedor(lstAyuda.SelectedValue)
    End Sub


    Private Sub btnAcepta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcepta.Click

        Dim txtUno, txtDos As String
        Dim txtFormula As String
        Dim x As Char = Chr(34)

        Dim WOrden As Integer
        Dim txtEmpresa As String

        Dim varOrdFecha As String
        Dim varCiclo As Integer
        Dim varPorce As Double
        Dim varAcumulado As Double
        Dim varProveedor, varLetra As String
        Dim varNeto, varIva, varIva5, varIva27, varIva105, varIb, varExento, varTotalTrabajo As Double
        Dim varRetIb, varRetIva, varRetGan, varAcumulaIb, varRete As Double
        Dim varPorceIb, varPorceIbCaba As Double
        Dim varTipoIbCaba, varTipoIva, varTipoPrv, varTipoIb As Integer
        Dim varPago, varEmpresa As Integer
        Dim varAcumulaNeto, varAcumulaNetoII, varAcumulaIva As Double
        Dim varFecha As String

        SQLConnector.retrieveDataTable("limpiar_impCtaCtePrvNet")

        txtEmpresa = "Surfactan S.A."
        varEmpresa = 1

        varOrdFecha = ordenaFecha(txtFechaEmision.Text)

        For varCiclo = 0 To varRenglon

            varProveedor = GRilla.Item(0, varCiclo).Value

            If LTrim(RTrim(varProveedor)) <> "" Then

                varAcumulado = 0
                varAcumulaIva = 0
                varAcumulaNeto = 0

                Dim tabla As DataTable
                tabla = SQLConnector.retrieveDataTable("buscar_cuenta_corriente_proveedores_selectivo", varProveedor)

                For Each row As DataRow In tabla.Rows

                    Dim CCPrv As New CtaCteProveedoresDeudaDesdeHastaII(row.Item(0).ToString, row.Item(1).ToString, row.Item(2).ToString, row.Item(3).ToString, row.Item(4), row.Item(5), row.Item(6).ToString, row.Item(7).ToString, row.Item(8).ToString, row.Item(9).ToString, row.Item(10), row.Item(11).ToString, row.Item(12).ToString, row.Item(13), row.Item(14))

                    varPago = CCPrv.pago
                    varParidad = CCPrv.paridad
                    varfecha = CCPrv.fecha

                    If varPago <> 2 Then
                        varTotal = CCPrv.total
                        varSaldo = CCPrv.saldo
                        varTotalUs = 0
                        varSaldoUs = 0
                        varSaldoOriginal = 0
                        varDife = 0
                    Else
                        varTotal = (CCPrv.total / varParidad) * varParidadTotal
                        varSaldo = (CCPrv.saldo / varParidad) * varParidadTotal
                        varTotalUs = (CCPrv.total / varParidad)
                        varSaldoUs = (CCPrv.saldo / varParidad)
                        varSaldoOriginal = CCPrv.saldo
                        varDife = varSaldo - CCPrv.saldo
                    End If

                    redondeo(varTotal)
                    redondeo(varSaldo)

                    varAcumulado = varAcumulado + varSaldo

                    If varTotal = varSaldo Then
                        varPorce = 1
                    Else
                        varPorce = varSaldo / varTotal
                    End If

                    varNeto = 0
                    varIva = 0
                    varIva5 = 0
                    varIva27 = 0
                    varIva105 = 0
                    varIb = 0
                    varExento = 0
                    varTotalTrabajo = 0
                    varLetra = ""

                    Dim CampoProveedor As Proveedor = DAOProveedor.buscarProveedorPorCodigo(CCPrv.Proveedor)
                    If IsNothing(CampoProveedor) Then
                        REM no existe
                    Else
                        varTipoIb = CampoProveedor.condicionIB1
                        varTipoIbCaba = CampoProveedor.condicionIB2
                        varTipoIva = CampoProveedor.codIva
                        varTipoPrv = CampoProveedor.tipo + 1
                        varPorceIb = CampoProveedor.porceIBProvincia
                        varPorceIbCaba = CampoProveedor.porceIBCABA

                        'WTipoIb = RstProveedor!CodIb
                        'WTipoIbCaba = RstProveedor!CodIbCaba
                        'WTipoiva = RstProveedor!Iva
                        'WTipoprv = Val(RstProveedor!Tipo) + 1
                        'WPorceIb = IIf(IsNull(RstProveedor!PorceIb), "0", RstProveedor!PorceIb)
                        'WPorceIbCaba = IIf(IsNull(RstProveedor!PorceIbCaba), "0", RstProveedor!PorceIbCaba)
                    End If

                    Dim compra As Compra = DAOCompras.buscarCompraPorCodigo(CCPrv.nroInterno)
                    If IsNothing(compra) Then
                        REM no existe
                    Else
                        varLetra = compra.letra
                        varNeto = compra.neto
                        varIva = compra.iva21
                        varIva5 = compra.ivaRG
                        varIva27 = compra.iva27
                        varIva105 = compra.iva105
                        varIb = compra.percibidoIB
                        varExento = compra.exento
                        varTotalTrabajo = varNeto + varIva + varIva5 + varIva27 + varIva105 + varIb + varExento
                        varLetra = compra.letra
                        varPago = compra.tipoPago
                    End If

                    varRetIb = 0
                    varRetIva = 0
                    varRetGan = 0
                    varAcumulaIb = 0

                    If varTotalTrabajo <> 0 Then
                        varAcumulaNetoII = varNeto * varPorce
                    Else
                        If varTipoIva = 2 Then
                            varAcumulaNetoII = (varSaldo / 1.21)
                        Else
                            varAcumulaNetoII = varSaldo
                        End If
                    End If

                    If varPago = 2 Then
                        varAcumulaNetoII = varAcumulaNetoII + (varDife / 1.21)
                    End If
                    varAcumulaNeto = varAcumulaNeto + varAcumulaNetoII

                    If varTipoIb = 0 Or varTipoIb = 1 Then
                        varRete = varAcumulaNeto * (varPorceIb / 100)
                        varAcumulaIb = varAcumulaIb + redondeo(varRete)
                        varRetIb = redondeo(varAcumulaIb)
                    End If

                    If varTipoIbCaba = 3 Or varTipoIbCaba = 4 Or varPorceIbCaba <> 0 Then
                        If varTipoIbCaba <> 2 Then
                            If varEmpresa = 1 Then
                                If varAcumulaNeto >= 300 Then
                                    If varPorceIbCaba <> 0 Then
                                        varRete = varAcumulaNeto * (varPorceIbCaba / 100)
                                    Else
                                        If varTipoIbCaba = 3 Then
                                            varRete = varAcumulaNeto * (3 / 100)
                                        Else
                                            varRete = varAcumulaNeto * (4.5 / 100)
                                        End If
                                    End If
                                End If
                                varAcumulaIb = varAcumulaIb + redondeo(varRete)
                                varRetIb = varAcumulaIb
                            End If
                        End If
                    End If
                    Stop



                    varOrdFecha = leederecha(ordenaFecha(varFecha), 6)
                    Dim CampoAcumulado As LeeAcumulado = DaoAcumulado.buscarAcumulado(varProveedor, varOrdFecha)


                    varRetGan = CaculoRetencionGanancia(varTipoPrv, varAcumulaNeto, CampoAcumulado.neto, CampoAcumulado.retenido, CampoAcumulado.anticipo, CampoAcumulado.bruto, CampoAcumulado.iva)

                    If varLetra = "M" Then
                        If varNeto >= 1000 Then
                            varAcumulaIva = varAcumulaIva + varIva
                        End If
                        varRetIb = varRetIb + varAcumulaIva
                    End If

                    '!acuneto = !Acumulado - WRetIb - WRetgan
                    '!Nombre = WNombre
                    '!Cheque = WCheque
                    '!ReteIb = WRetIb
                    '!ReteGan = WRetgan




                    SQLConnector.executeProcedure("alta_impCtaCtePrvNet", CCPrv.Clave, CCPrv.Proveedor, CCPrv.Tipo, CCPrv.letra, CCPrv.punto, CCPrv.numero, varTotal, varSaldo, CCPrv.fecha, CCPrv.vencimiento, txtFechaEmision.Text, CCPrv.Impre, CCPrv.nroInterno, txtEmpresa, varAcumulado, WOrden, txtFechaEmision.Text, "", "", "", varParidadTotal, varSaldoOriginal, varDife, 0, 0, "", 0, 0, 0, varParidad, varTotalUs, varSaldoUs, 0, 0)

                Next

            End If

        Next

        Stop

        txtUno = "{impCtaCtePrvNet.Proveedor} in " + x + "" + x + " to " + x + "ZZZZZZZZZZZ" + x
        txtDos = " and {impCtaCtePrvNet.Saldo} <> 0.00"
        txtFormula = txtUno + txtDos

        'Dim viewer As New ReportViewer("Listado de Corriente de Proveedres Selectivo", Globals.reportPathWithName("wccprvfecnet.rpt"), txtFormula)

        'If opcPantalla.Checked = True Then
        '    viewer.Show()
        'Else
        '    viewer.imprimirReporte()
        'End If


    End Sub

End Class