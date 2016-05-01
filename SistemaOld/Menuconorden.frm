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
      Begin VB.Menu Arti 
         Caption         =   "Ingreso de Materias Primas"
      End
      Begin VB.Menu Terminado 
         Caption         =   "Ingreso de Producto Terminado"
      End
      Begin VB.Menu Prove 
         Caption         =   "Ingreso de Proveedores"
      End
      Begin VB.Menu ProveAdicional 
         Caption         =   "Datos Adicionales de Proveedores"
      End
      Begin VB.Menu Homologa 
         Caption         =   "Homologacion de Muestras "
      End
      Begin VB.Menu Posicion 
         Caption         =   "Ingreso de Posicion Arncelaria"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Cotiza 
         Caption         =   "Ingreso de Cotizaciones"
      End
      Begin VB.Menu Orden 
         Caption         =   "Emision de Ordenes de Compra"
      End
      Begin VB.Menu Movvar 
         Caption         =   "Ingreso de Movimientos Varios"
      End
      Begin VB.Menu Solic 
         Caption         =   "Ingreso de Solicitud de Pedido de Compra"
      End
      Begin VB.Menu Mirasol 
         Caption         =   "Consulta de Solicitiudes de Pedido de Compra"
      End
      Begin VB.Menu Insumo 
         Caption         =   "Ingreso de Solicitud de Compras de Insumos"
      End
      Begin VB.Menu mirainsumos 
         Caption         =   "Consulta de Solicitudes de Compras de Insumos"
      End
      Begin VB.Menu mirainsumosiii 
         Caption         =   "Consulta de Solicitudes por Origen"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu ListCot 
         Caption         =   "Listado de Cotizaciones"
      End
      Begin VB.Menu ListOrd 
         Caption         =   "Listado de Ordenes de Compra"
      End
      Begin VB.Menu CotPrv 
         Caption         =   "Listado de Cotizaciones por Proveedor"
      End
      Begin VB.Menu CotArt 
         Caption         =   "Listado de Cotizaciones por Articulo"
      End
      Begin VB.Menu Orden1 
         Caption         =   "Listado de O/C Pend. por Proveedor"
      End
      Begin VB.Menu Orden2 
         Caption         =   "Listado de O/C Pend. por Articulos"
      End
      Begin VB.Menu Listmat1 
         Caption         =   "Listado de Materia Prima"
      End
      Begin VB.Menu Listmat2 
         Caption         =   "Listado de Materia Prima ( Stock )"
      End
      Begin VB.Menu Listter 
         Caption         =   "Listado de Producto Terminado (Stock)"
      End
      Begin VB.Menu Minimo 
         Caption         =   "Listado de Materia Prima (Minimo)"
      End
      Begin VB.Menu MInter 
         Caption         =   "Listado de Producto Terminado (Minimo)"
      End
      Begin VB.Menu FichaMp 
         Caption         =   "Listado de Ficha de Stock de M.P."
      End
      Begin VB.Menu FiechaPt 
         Caption         =   "Listado de Ficha de Stock de P.T."
      End
      Begin VB.Menu Minimo1 
         Caption         =   "Listado de Materia Prima (Minimo Consolidado)"
      End
      Begin VB.Menu MInter1 
         Caption         =   "Listado de Producto Terminado (Minimo Consolidado)"
      End
   End
   Begin VB.Menu n 
      Caption         =   "Listados "
      Begin VB.Menu CompPrv 
         Caption         =   "Listado de Compras por Proveedor"
      End
      Begin VB.Menu CompMat 
         Caption         =   "Listado de Compras por Materia Prima"
      End
      Begin VB.Menu ConArtCon 
         Caption         =   "Listado de Compras por Materia Prima Concolidada"
      End
      Begin VB.Menu LIstInfPend 
         Caption         =   "Listado de Informe de Recepcion Pendientes de Aprobacion"
      End
      Begin VB.Menu ListCont 
         Caption         =   "Listado de Control de Ordenes"
      End
      Begin VB.Menu ConsFichaMp 
         Caption         =   "Consulta de Ficha de Stock M.P."
      End
      Begin VB.Menu ConFichaPt 
         Caption         =   "Consulta de Ficha de Stock P.T."
      End
      Begin VB.Menu Ultima 
         Caption         =   "Listado de Ultima Compra de Materia Prima"
      End
      Begin VB.Menu ConsFicMatAnt 
         Caption         =   "Consulta de Ficha de Materia Prima Historica"
      End
      Begin VB.Menu ConsFicTerAnt 
         Caption         =   "Consulta de Ficha de Producto Terminado Historico"
      End
      Begin VB.Menu ListInfImporta 
         Caption         =   "Listado de Analisis de Ordenes de Compra de Importacion"
      End
      Begin VB.Menu HistoriaTerminado 
         Caption         =   "Listado de Verificacion de Ultimos Movimientos de P.T."
      End
      Begin VB.Menu HistoriaArticulo 
         Caption         =   "Listado de Verificacion de Ultimos Movimientos de M.P."
      End
   End
   Begin VB.Menu dy 
      Caption         =   "Listados de Reventa"
      Begin VB.Menu ConsFicMAtDy 
         Caption         =   "Consulta de Fichas de Stock"
      End
      Begin VB.Menu ListFicMatDy 
         Caption         =   "Listado de Fichas de Stock"
      End
      Begin VB.Menu ListStkDy 
         Caption         =   "Listado de Stock"
      End
   End
   Begin VB.Menu procesos 
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
            Rem Open "dada.txt" For Output As #1
    Open "http://clientes.gb2.com.ar/demo/archivos/descargas25.txt" For Output As #1

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
    
    XOperador = Str$(WOperador)
    XProceso = "1"
    WAtributo1 = "00000000000000000000000000000000000000000000"
    WAtributo2 = "00000000000000000000000000000000000000000000"
    WAtributo3 = "00000000000000000000000000000000000000000000"
    WAtributo4 = "00000000000000000000000000000000000000000000"
    WAtributo5 = "00000000000000000000000000000000000000000000"
    WAtributo6 = "00000000000000000000000000000000000000000000"
    WAtributo7 = "00000000000000000000000000000000000000000000"
    WAtributo8 = "00000000000000000000000000000000000000000000"
    WAtributo9 = "00000000000000000000000000000000000000000000"
    WAtributo10 = "00000000000000000000000000000000000000000000"
    
    XParam = "'" + XOperador + "','" _
                 + XProceso + "'"
    spAtributo = "ConsultaAtributo " + XParam
    Set rstAtributo = db.OpenRecordset(spAtributo, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtributo.RecordCount > 0 Then
        WAtributo1 = rstAtributo!Atributo1 + "00000000000000000000000000000"
        WAtributo2 = rstAtributo!Atributo2 + "00000000000000000000000000000"
        WAtributo3 = rstAtributo!Atributo3 + "00000000000000000000000000000"
        WAtributo4 = rstAtributo!Atributo4 + "00000000000000000000000000000"
        WAtributo5 = rstAtributo!Atributo5 + "00000000000000000000000000000"
        WAtributo6 = rstAtributo!Atributo6 + "00000000000000000000000000000"
        WAtributo7 = rstAtributo!Atributo7 + "00000000000000000000000000000"
        WAtributo8 = rstAtributo!Atributo8 + "00000000000000000000000000000"
        WAtributo9 = rstAtributo!Atributo9 + "00000000000000000000000000000"
        WAtributo10 = rstAtributo!Atributo10 + "00000000000000000000000000000"
        rstAtributo.Close
    End If
    
    For Ciclo = 1 To 10
        Select Case Ciclo
            Case 1
                Auxiliar = WAtributo1
            Case 2
                Auxiliar = WAtributo2
            Case 3
                Auxiliar = WAtributo3
            Case 4
                Auxiliar = WAtributo4
            Case 5
                Auxiliar = WAtributo5
            Case 6
                Auxiliar = WAtributo6
            Case 7
                Auxiliar = WAtributo7
            Case 8
                Auxiliar = WAtributo8
            Case 9
                Auxiliar = WAtributo9
            Case 10
                Auxiliar = WAtributo10
            Case Else
        End Select
        For Ciclo1 = 1 To 31
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
        Next Ciclo1
    Next Ciclo
            
    Menu.Arti.Enabled = Atri(1, 1)
    Menu.Terminado.Enabled = Atri(1, 2)
    Menu.Prove.Enabled = Atri(1, 3)
    Menu.ProveAdicional.Enabled = Atri(1, 4)
    Menu.Homologa.Enabled = Atri(1, 5)
    Menu.Posicion.Enabled = Atri(1, 6)
    
    Menu.Cotiza.Enabled = Atri(2, 1)
    Menu.Orden.Enabled = Atri(2, 2)
    Menu.Movvar.Enabled = Atri(2, 3)
    Menu.Solic.Enabled = Atri(2, 4)
    Menu.Mirasol.Enabled = Atri(2, 5)
    Menu.Insumo.Enabled = Atri(2, 6)
    Menu.mirainsumos.Enabled = Atri(2, 7)
    Menu.mirainsumosiii.Enabled = Atri(2, 8)
    
    Menu.ListCot.Enabled = Atri(3, 1)
    Menu.ListOrd.Enabled = Atri(3, 2)
    Menu.CotPrv.Enabled = Atri(3, 3)
    Menu.CotArt.Enabled = Atri(3, 4)
    Menu.Orden1.Enabled = Atri(3, 5)
    Menu.Orden2.Enabled = Atri(3, 6)
    Menu.Listmat1.Enabled = Atri(3, 7)
    Menu.Listmat2.Enabled = Atri(3, 8)
    Menu.Listter.Enabled = Atri(3, 9)
    Menu.Minimo.Enabled = Atri(3, 10)
    Menu.MInter.Enabled = Atri(3, 11)
    Menu.FichaMp.Enabled = Atri(3, 12)
    Menu.FiechaPt.Enabled = Atri(3, 13)
    Menu.Minimo1.Enabled = Atri(3, 14)
    Menu.MInter1.Enabled = Atri(3, 15)
    
    Menu.CompPrv.Enabled = Atri(4, 1)
    Menu.CompMat.Enabled = Atri(4, 2)
    Menu.ConArtCon.Enabled = Atri(4, 3)
    Menu.LIstInfPend.Enabled = Atri(4, 4)
    Menu.ListCont.Enabled = Atri(4, 5)
    Menu.ConsFichaMp.Enabled = Atri(4, 6)
    Menu.ConFichaPt.Enabled = Atri(4, 7)
    Menu.Ultima.Enabled = Atri(4, 8)
    Menu.ConsFicMatAnt.Enabled = Atri(4, 9)
    Menu.ConsFicTerAnt.Enabled = Atri(4, 10)
    Menu.ListInfImporta.Enabled = Atri(4, 11)
    Menu.HistoriaTerminado.Enabled = Atri(4, 12)
    Menu.HistoriaArticulo.Enabled = Atri(4, 13)
    
    Menu.ConsFicMAtDy.Enabled = Atri(5, 1)
    Menu.ListFicMatDy.Enabled = Atri(5, 2)
    Menu.ListStkDy.Enabled = Atri(5, 3)
    
    Menu.FinCot.Enabled = 1

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
