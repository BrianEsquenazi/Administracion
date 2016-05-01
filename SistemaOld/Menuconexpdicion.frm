VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Cotizaciones"
   ClientHeight    =   7815
   ClientLeft      =   2280
   ClientTop       =   780
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Informe 
         Caption         =   "Ingreso de Informe de Recepcion"
      End
      Begin VB.Menu Hoja 
         Caption         =   "Ingreso de Hoja de Produccion"
      End
      Begin VB.Menu Movvar 
         Caption         =   "Ingreso de Movimientos Varios"
      End
      Begin VB.Menu Pedeti 
         Caption         =   "Emision de Etiquetas de Exportacion"
      End
      Begin VB.Menu Movguia 
         Caption         =   "Ingreso de Guias de Traslado Interno"
      End
      Begin VB.Menu Prestamo 
         Caption         =   "Prestamos entre plantas"
      End
      Begin VB.Menu Actualiza 
         Caption         =   "Actualizacion de Pedidos"
      End
      Begin VB.Menu ModPedExp 
         Caption         =   "Actualizacion de Pedidos de Exportacion"
      End
      Begin VB.Menu Solic 
         Caption         =   "Ingreso de Solicitud de Pedido de Compra"
      End
      Begin VB.Menu Entdev 
         Caption         =   "Entrada de Devolucion de Mercaderia"
      End
      Begin VB.Menu SolHoja 
         Caption         =   "Ingreso de Solicitud de Hoja de Produccion"
      End
      Begin VB.Menu Centro 
         Caption         =   "Verificacion de Pedidos"
      End
      Begin VB.Menu Insumo 
         Caption         =   "Ingreso de Solicitud de Compras de Insumos"
      End
      Begin VB.Menu VerificaPedido 
         Caption         =   "Verificacion de Pedidos Pendientes"
      End
      Begin VB.Menu HojaProduccion 
         Caption         =   "Actualizacion de Hojas de Produccion"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu FichaMp 
         Caption         =   "Listado de Ficha de Stock de M.P."
      End
      Begin VB.Menu FiechaPt 
         Caption         =   "Listado de Ficha de Stock de P.T."
      End
      Begin VB.Menu ConsFichaMp 
         Caption         =   "Consulta de Ficha de Stock M.P."
      End
      Begin VB.Menu ConFichaPt 
         Caption         =   "Consulta de Ficha de Stock P.T."
      End
      Begin VB.Menu Eti1 
         Caption         =   "Emision de Etiquetas"
      End
      Begin VB.Menu Eti4 
         Caption         =   "Emision de Etiquetas  M.P. de Reventa"
      End
      Begin VB.Menu Lotemat 
         Caption         =   "Listado de Ficha de Lote de Materia Prima"
      End
      Begin VB.Menu Loteter 
         Caption         =   "Listado de Ficha de Lote de Producto Terminado"
      End
      Begin VB.Menu Pedpen 
         Caption         =   "Listado de Pedidos Pendientes"
      End
      Begin VB.Menu Pedpenter 
         Caption         =   "Listado de Pedidos Pendientes por Producto Terminado"
      End
      Begin VB.Menu EtiVerde 
         Caption         =   "Emision de Etiquetas Verdes"
      End
   End
   Begin VB.Menu asdasdsa 
      Caption         =   "Procesos"
      Begin VB.Menu FinCot 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAtributo As Recordset
Dim spAtributo As String
Dim Atri(10, 100) As Integer

Private Sub ActuaCargaSolicitud_Click()
    PrgBajaSolicitud.Show
End Sub

Private Sub Actualiza_Click()
    PrgModpedNuevo.Show
End Sub

Private Sub ActualizaInforme_Click()
    PrgActualizaInforme.Show
End Sub

Private Sub ALtaProt_Click()
    PrgAltaProyec.Show
End Sub

Private Sub AnalisisOrden_Click()
    PrgAnalisisOrden.Show
End Sub

Private Sub AnalisisMp_Click()
    PrgAnalisisMp.Show
End Sub

Private Sub AnalisisPt_Click()
    PrgAnalisisPt.Show
End Sub

Private Sub Arti_Click()
    PrgArti.Show
End Sub

Private Sub Califica_Click()
    PrgCalifica.Show
End Sub

Private Sub CalificaGrilla_Click()
    PrgCalificaGrilla.Show
End Sub

Private Sub CargaSolicitud_Click()
    PrgCargaSolicitud.Show
End Sub

Private Sub Centro_Click()
    PrgCentro.Show
End Sub

Private Sub CierreStk_Click()
    OPEN_FILE_InveMp
    OPEN_FILE_InvePt
    PrgCierre.Show
End Sub

Private Sub Command1_Click()

    Stop

    ZSql = ""
    ZSql = ZSql + "DELETE CtaCte"
    ZSql = ZSql + " Where OrdFecha < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Saldo = 0 "
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE CtaCtePrv"
    ZSql = ZSql + " Where OrdFecha < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Saldo = 0 "
    spCtactePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtactePrv, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Depositos"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spDeposito = ZSql
    Set RstDeposito = db.OpenRecordset(spDeposito, dbOpenSnapshot, dbSQLPassThrough)
        
            
    ZSql = ""
    ZSql = ZSql + "DELETE Estadistica"
    ZSql = ZSql + " Where OrdFecha < " + "'" + "20070101" + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Guia"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Saldo = 0 "
    spGuia = ZSql
    Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Hoja"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Saldo = 0 "
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)

    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Imputac"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spImputac = ZSql
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)

    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Informe"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Laudo"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Saldo = 0 "
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Movlab"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spMovlab = ZSql
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Movvar"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spMovvar = ZSql
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Orden"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Pagos"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spPago = ZSql
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Pedido"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "DELETE Recibos"
    ZSql = ZSql + " Where FechaOrd < " + "'" + "20070101" + "'"
    ZSql = ZSql + " and Estado2 <> 'P'"
    spRecibo = ZSql
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)

    Stop

End Sub

Private Sub CompMat_Click()
    PrgOrdart.Show
End Sub

Private Sub Compo_Click()
    PrgCompos.Show
End Sub

Private Sub Compos1_Click()
    PrgCompos1.Show
End Sub

Private Sub CompPrv_Click()
    PrgOrdprv.Show
End Sub

Private Sub ConArtCon_Click()
    PrgOrdartCon.Show
End Sub

Private Sub ConFichaPt_Click()
    PrgConsFicTer.Show
End Sub

Private Sub ConsFichaMp_Click()
    PrgConsFicMat.Show
End Sub

Private Sub ConsFicMatAnt_Click()
    PrgConsFicMatAnt.Show
End Sub

Private Sub ConsFicMAtDy_Click()
    If Val(WEmpresa) = 6 Then
        PrgConsFicMatDyTransito.Show
            Else
        PrgConsFicMatDy.Show
    End If
End Sub

Private Sub ConsFicTerAnt_Click()
    PrgConsFicTerAnt.Show
End Sub

Private Sub ConsumoArt_Click()
    PrgConsumoArt.Show
End Sub

Private Sub ConsumoMatII_Click()
    PrgConsumoArtII.Show
End Sub

Private Sub Consumoter_Click()
    PrgConsumoTer.Show
End Sub

Private Sub ConsumoTerII_Click()
    PrgConsumoTerII.Show
End Sub

Private Sub Costo_Click()
    PrgCosto.Show
End Sub

Private Sub Cotart_Click()
    PrgCotart.Show
End Sub

Private Sub Cotiza_Click()
    PrgCoti.Show
End Sub

Private Sub Cotprv_Click()
    PrgCoTPRV.Show
End Sub

Private Sub deputa_Click()
    PrgDepuraSaldos.Show
End Sub

Private Sub DepuraMp_Click()
    PrgDepuraMp.Show
End Sub

Private Sub DepuraPt_Click()
    PrgDepuraPt.Show
End Sub

Private Sub DepuraSaldosInforme_Click()
    PrgDepuraSaldoInforme.Show
End Sub

Private Sub DepuraSaldosOrden_Click()
    PrgDepuraSaldoOrden.Show
End Sub

Private Sub DispoPt_Click()
    PrgDisponiblePt.Show
End Sub

Private Sub DispoPtOtro_Click()
    PrgDisponiblePtOtro.Show
End Sub

Private Sub Efluentes_Click()
    PrgEfluentes.Show
End Sub

Private Sub Entdev_Click()
    PrgEntdev.Show
End Sub

Private Sub Envases_Click()

  PrgEnv.Show

End Sub

Private Sub Eti1_Click()
    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa
    PrgEti3.Show
End Sub

Private Sub Eti4_Click()
    Rem PrgEti4.Show
    PrgEti5.Show
End Sub

Private Sub FGH_Click()
    Form1.Show
End Sub

Private Sub EtiVerde_Click()
    PrgEtiVerde.Show
End Sub

Private Sub FichaMp_Click()
    PrgFicmat.Show
End Sub

Private Sub FiechaPt_Click()
    PrgFicter.Show
End Sub

Private Sub FinCot_Click()
    Close
    End
End Sub

Private Sub HistoriaArticulo_Click()
    PrgHistoriaArticulo.Show
End Sub

Private Sub HistoriaTerminado_Click()
    PrgHistoriaTerminado.Show
End Sub

Private Sub Hoja_Click()
    If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
        PrgHoja.Show
    End If
End Sub

Private Sub HojaProduccion_Click()
    PrgHojaProduccion.Show
End Sub

Private Sub Homologa_Click()
    PrgHomologaProve.Show
End Sub

Private Sub Informe_Click()
    PrgInforme.Show
End Sub

Private Sub Insumo_Click()
    PrgInsumos.Show
End Sub

Private Sub Laudo_Click()
    Prglaudo.Show
End Sub

Private Sub ListaCalifica_Click()
    PrgListaCalifica.Show
End Sub

Private Sub LIstaCarga_Click()
    PrgListaCarga.Show
End Sub

Private Sub ListaCheckList_Click()
    PrgListaCheckList.Show
End Sub

Private Sub ListaDerechos_Click()
    PrgListaDerechos.Show
End Sub

Private Sub ListaImpoFecha_Click()
    PrgListaImpoFecha.Show
End Sub

Private Sub ListaImpoFechaII_Click()
    PrgListaImpoFechaII.Show
End Sub

Private Sub ListaProydyiiPanta_Click()
    PrgListProyDyIIPanta.Show
End Sub

Private Sub LIstaProyec_Click()
    PrgListproyec.Show
End Sub

Private Sub ListArt1_Click()
    PrgValuaMat.Show
End Sub

Private Sub Listcomp_Click()
    PrgListcomp.Show
End Sub

Private Sub Listcomp1_Click()
    PrgListcomp1.Show
End Sub

Private Sub ListCont_Click()
    PrgControl.Show
End Sub

Private Sub ListCot_Click()
    PrgListcot.Show
End Sub

Private Sub ListDispoDy_Click()
    PrgListDispoDy.Show
End Sub

Private Sub ListEnv1_Click()
    PrgListmov1.Show
End Sub

Private Sub ListEnv2_Click()
    PrgListmov2.Show
End Sub

Private Sub ListFicMatDy_Click()
    PrgFicmatDy.Show
End Sub

Private Sub Listhoja_Click()
    PrgListhoja.Show
End Sub

Private Sub ListInf_Click()
    PrgListinf.Show
End Sub

Private Sub ListInfImporta_Click()
    PrgListinfImporta.Show
End Sub

Private Sub LIstInfPend_Click()
    prglistinfpend.Show
End Sub

Private Sub Listmat1_Click()
    PrgListmat1.Show
End Sub

Private Sub Listmat2_Click()
    PrgStkmat.Show
End Sub

Private Sub ListOrd_Click()
    PrgListOrd.Show
End Sub

Private Sub Listpres_Click()
    PrgListpres.Show
End Sub

Private Sub ListProyDWII_Click()
    PrgListProyDwII.Show
End Sub

Private Sub ListProydyii_Click()
    PrgListProyDyII.Show
End Sub

Private Sub ListStkDy_Click()
    PrgStockConsol.Show
End Sub

Private Sub ListStkDyPedido_Click()
    PrgListStkDyPedido.Show
End Sub

Private Sub ListStkFamDq_Click()
    PrgListStkFamDq.Show
End Sub

Private Sub ListStkFamDs_Click()
    PrgListStkFamDs.Show
End Sub

Private Sub ListStkFamDw_Click()
    PrgListStkFamDw.Show
End Sub

Private Sub ListStkFamDy_Click()
    PrgListStkFamDy.Show
End Sub

Private Sub Listter_Click()
    PrgListter.Show
End Sub

Private Sub ListTer1_Click()
    PrgLister1.Show
End Sub

Private Sub Lotemat_Click()
    PrgLotemat.Show
End Sub

Private Sub Loteter_Click()
    PrgLoteter.Show
End Sub

Private Sub Metodos_Click()
    PrgMetodos.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimoPlanta.Show
End Sub

Private Sub Minimo1_Click()
    PrgMinimoConsol.Show
End Sub

Private Sub Minter_Click()
    PrgMinTerPlanta.Show
End Sub

Private Sub MInter1_Click()
    PrgMinTerConsol.Show
End Sub

Private Sub MiraInsumos_Click()
    ZMuestraCol = 1
    ZMuestraRow = 1
    ZMuestraTopRow = 1
    PrgMiraInsumos.Show
End Sub

Private Sub mirainsumosiii_Click()
    WTextoSolicitante = ""
    PrgMiraInsumosII.Show
End Sub

Private Sub mirasol_Click()
    PrgMIrasol.Show
End Sub

Private Sub Mirasolguia_Click()
    PrgMiraSolGuia.Show
End Sub

Private Sub Modifcolor_Click()
    ProcesoActivate = 0
    PrgModifColor.Show
End Sub

Private Sub ModPedExp_Click()
    PrgModpedExp.Show
End Sub

Private Sub MovEnv_Click()
    PrgMovEnv.Show
End Sub

Private Sub Movgas_Click()
    PrgMovgas.Show
End Sub

Private Sub Movguia_Click()
    PrgMovguia.Show
End Sub

Private Sub Movvar_Click()
    PrgMovvar.Show
End Sub

Private Sub Movvar1_Click()
    PrgMovvar1.Show
End Sub

Private Sub Movvar2_Click()
    PrgMovvar2.Show
End Sub

Private Sub Orden_Click()
    WProcesoOrden = 0
    PrgOrden.Show
End Sub

Private Sub Orden1_Click()
    PrgOrdPenPrv.Show
End Sub

Private Sub Orden2_Click()
    PrgOrdPenArt.Show
End Sub

Private Sub OrdenAnual_Click()
    PrgOrdenAnual.Show
End Sub

Private Sub OrdenDy_Click()
    PrgOrdenDy.Show
End Sub

Private Sub OrdPrnDy_Click()
    PrgOrdPenDy.Show
End Sub

Private Sub OrdPrnDyII_Click()
    PrgOrdPenDyII.Show
End Sub

Private Sub pasadtdy_Click()
    PrgPasaPtDy.Show
End Sub

Private Sub pasadw_Click()
    PrgPasaDW.Show
End Sub

Private Sub PasaDy_Click()
    PrgPasaDy.Show
End Sub

Private Sub Pedeti_Click()
    PrgPedeti.Show
End Sub

Private Sub PedPen_Click()
    PrgPedPen.Show
End Sub

Private Sub PedPenDy_Click()
    PrgPedPendy.Show
End Sub

Private Sub Pedpenter_Click()
    PrgPedPenTer.Show
End Sub

Private Sub Posicion_Click()
    PrgPosicion.Show
End Sub

Private Sub Prestamo_Click()
    PrgPrestamo.Show
End Sub

Private Sub PrgActuaFactuExpoasd_Click()
    PrgActuaFactuexpo.Show
End Sub

Private Sub Proc1_Click()
    PrgProc1.Show
End Sub

Private Sub Proc101_Click()
    PrgProc101.Show
End Sub

Private Sub Proc102_Click()
    PrgProc102.Show
End Sub

Private Sub Proc11_Click()
    PrgProc11.Show
End Sub

Private Sub Proc2_Click()
    PrgProc2.Show
End Sub

Private Sub Proc9_Click()
    PrgProc9.Show
End Sub

Private Sub ProcesoFecha_Click()
    PrgProcesoFecha.Show
End Sub

Private Sub ProcFabrica_Click()
    PrgCargaIV.Show
End Sub

Private Sub ProcHoja_Click()
    PrgProchoja.Show
End Sub

Private Sub ProcHojaEspecif_Click()
    PrgProchojaEspecif.Show
End Sub

Private Sub prove_Click()
    PrgProve.Show
End Sub

Private Sub ProveAdicional_Click()
    PrgProveAdicional.Show
End Sub

Private Sub Proy_Click()
    PrgProyec.Show
End Sub

Private Sub Sedronar_Click()
    PrgSedronar.Show
End Sub

Private Sub ProyStkDy_Click()
    PrgListProyDy.Show
End Sub

Private Sub RepPedpen_Click()
    PrgProcPedPen.Show
End Sub

Private Sub reprocesoactuaorden_Click()
    PrgReprocesoActuaOrdenPartida.Show
End Sub

Private Sub Seguimiento_Click()
    PrgSeguimiento.Show
End Sub

Private Sub SolGuia_Click()
    PrgSolGuia.Show
End Sub

Private Sub SolHoja_Click()
    PrgSolHoja.Show
End Sub

Private Sub Solic_Click()
    PrgSolic.Show
End Sub

Private Sub Solicitud_Click()
    PrgSolicitud.Show
End Sub

Private Sub SolZona_Click()
    PrgCargaZona.Show
End Sub

Private Sub StkMinDy_Click()
    PrgListMinDy.Show
End Sub

Private Sub Terminado_Click()
    PrgTermi.Show
End Sub

Private Sub Cambio_Click()
    frmLoginCotiza.Show
End Sub

Private Sub Fin_Click()
    Menu.WindowState = 1
End Sub

Private Sub Form_Activate()

    Rem If MiRuta = "" Then
    Rem     MiRuta = CurDir + "\"
    Rem     MiRutaII = Left$(CurDir, 1)
    Rem End If
    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
    If WEmpresa = "" Then
    
        WEmpresa = "0001"
        Rem Empresa.Show
        Rem Empresa.SetFocus
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With

            Else
            
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With
        
    End If
    
End Sub

Private Sub Ultima_Click()
    PrgUltima.Show
End Sub

Private Sub valo1_Click()
    PrgStock1Otro.Show
End Sub

Private Sub Valo2_Click()
    PrgStock2Otro.Show
End Sub

Private Sub ValuacionPartida_Click()
    PrgValuaMatPartida.Show
End Sub

Private Sub Verifica_Click()
    PrgVerifica.Show
End Sub

Private Sub VerificaPedido_Click()
    PrgVerificaPedido.Show
End Sub

Private Sub verilot1_Click()
    PrgVerilot1.Show
End Sub

Private Sub Verilot2_Click()
    PrgVerilot2.Show
End Sub

Private Sub verilot3_Click()
    PrgVerilot3.Show
End Sub

Private Sub verio_Click()
    PrgVeriSaldosInforme.Show
End Sub
