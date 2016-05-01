VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Cotizaciones"
   ClientHeight    =   7815
   ClientLeft      =   2280
   ClientTop       =   780
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   7680
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
      Begin VB.Menu Envases 
         Caption         =   "Ingreso de Envases"
      End
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
      Begin VB.Menu Efluentes 
         Caption         =   "Ingreso de Efluentes de Lavado"
      End
      Begin VB.Menu Homologa 
         Caption         =   "Homologacion de Muestras "
      End
      Begin VB.Menu reprocesoactuaorden 
         Caption         =   "Reproceso"
      End
      Begin VB.Menu Posicion 
         Caption         =   "Ingreso de Posicion Arncelaria"
      End
      Begin VB.Menu PrgActuaFactuExpoasd 
         Caption         =   "Actualizacion de Partidas de Facturas de Exportacion"
      End
      Begin VB.Menu Trazabilidad 
         Caption         =   "Trazabilidad de Factura de Exportacion"
      End
      Begin VB.Menu TrazabilidadTraspa 
         Caption         =   "Trazabilidad de Factura de Exportacion (exportacion)"
      End
      Begin VB.Menu asdfas 
         Caption         =   "Centro de Control de Importaciones"
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
      Begin VB.Menu Informe 
         Caption         =   "Ingreso de Informe de Recepcion"
      End
      Begin VB.Menu Laudo 
         Caption         =   "Ingreso de Laudo de Liberacion"
      End
      Begin VB.Menu Hoja 
         Caption         =   "Ingreso de Hoja de Produccion"
      End
      Begin VB.Menu Movvar 
         Caption         =   "Ingreso de Movimientos Varios"
      End
      Begin VB.Menu MovEnv 
         Caption         =   "Ingreso y Egreso de Envases"
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
      Begin VB.Menu Mirasol 
         Caption         =   "Consulta de Solicitiudes de Pedido de Compra"
      End
      Begin VB.Menu Entdev 
         Caption         =   "Entrada de Devolucion de Mercaderia"
      End
      Begin VB.Menu SolHoja 
         Caption         =   "Ingreso de Solicitud de Hoja de Produccion"
      End
      Begin VB.Menu ModifColor 
         Caption         =   "Actualizacion de Pedidos de Colorantes  DY / DW / DS"
      End
      Begin VB.Menu Centro 
         Caption         =   "Verificacion de Pedidos"
      End
      Begin VB.Menu SolGuia 
         Caption         =   "Ingreso de Solidictud de Guia de Traslado Interno"
      End
      Begin VB.Menu Mirasolguia 
         Caption         =   "Consultas de Guias de Traslado Interno"
      End
      Begin VB.Menu ActualizaInforme 
         Caption         =   "Actualizacion de Informes de Recepcion de M.P. de Reventa"
      End
      Begin VB.Menu DepuraMp 
         Caption         =   "Depuracion de Restos de Materia Prima"
      End
      Begin VB.Menu DepuraPt 
         Caption         =   "Depuracion de Restos de Productos Terminados"
      End
      Begin VB.Menu ALtaProt 
         Caption         =   "Ingreso de Proyeccion de Ventas Anuales de M.P. de Reventa"
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
      Begin VB.Menu VerificaPedido 
         Caption         =   "Verificacion de Pedidos Pendientes"
      End
      Begin VB.Menu Solicitud 
         Caption         =   "Ingreso de Solicitud de Fondos"
         Visible         =   0   'False
      End
      Begin VB.Menu CargaSolicitud 
         Caption         =   "Carga de Solicitud de Produccion"
      End
      Begin VB.Menu ActuaCargaSolicitud 
         Caption         =   "Recepcion de Produccion"
      End
      Begin VB.Menu HojaProduccion 
         Caption         =   "Actualizacion de Hojas de Produccion"
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
      Begin VB.Menu ListArt1 
         Caption         =   "Listado de Valuacion de Materia Prima"
      End
      Begin VB.Menu ListTer1 
         Caption         =   "Listado de Valuacion de Producto Terminado"
      End
      Begin VB.Menu Minimo 
         Caption         =   "Listado de Materia Prima (Minimo)"
      End
      Begin VB.Menu MInter 
         Caption         =   "Listado de Producto Terminado (Minimo)"
      End
      Begin VB.Menu Compo 
         Caption         =   "Listado de Composicion"
      End
      Begin VB.Menu Proy 
         Caption         =   "Listado de Proyeccion de Entradas"
      End
      Begin VB.Menu FichaMp 
         Caption         =   "Listado de Ficha de Stock de M.P."
      End
      Begin VB.Menu FiechaPt 
         Caption         =   "Listado de Ficha de Stock de P.T."
      End
      Begin VB.Menu Movvar1 
         Caption         =   "Listado de Movimientos Varios de Materia Prima"
      End
      Begin VB.Menu Movvar2 
         Caption         =   "Listado de Movimientos Varios de Producto Terminado"
      End
      Begin VB.Menu Listhoja 
         Caption         =   "Listado de Hojas de Produccion"
      End
      Begin VB.Menu Minimo1 
         Caption         =   "Listado de Materia Prima (Minimo Consolidado)"
      End
      Begin VB.Menu MInter1 
         Caption         =   "Listado de Producto Terminado (Minimo Consolidado)"
      End
      Begin VB.Menu ConsumoTer 
         Caption         =   "Listado de Consumo de Productos Terminados (H.P.)"
      End
      Begin VB.Menu ConsumoArt 
         Caption         =   "Listado de Consumo de Materia Prima  (H.P.)"
      End
      Begin VB.Menu DispoPt 
         Caption         =   "Listado de Disponibilidad de Producto Terminado"
      End
      Begin VB.Menu LIstaCarga 
         Caption         =   "Listado de Solicitudes de Produccion Pendientes de Entrega"
      End
      Begin VB.Menu DispoPtOtro 
         Caption         =   "Listado de Disponibilidad de Producto Terminado (Pellital)"
      End
      Begin VB.Menu AnalisisMp 
         Caption         =   "Analisis de Materia Prima"
      End
      Begin VB.Menu AnalisisPt 
         Caption         =   "Analisis de Producto Terminado"
      End
      Begin VB.Menu ListaImpoFecha 
         Caption         =   "Listado de Valorizacion de Importaciones a Fecha"
      End
      Begin VB.Menu Importa 
         Caption         =   "Comparativo de Importaciones/Exportacioes entre Periodos"
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
         Caption         =   "Listado de Compras por Materia Prima Consolidada"
      End
      Begin VB.Menu ListInf 
         Caption         =   "Listado de Informe de Recepcion"
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
      Begin VB.Menu Eti1 
         Caption         =   "Emision de Etiquetas"
      End
      Begin VB.Menu Eti4 
         Caption         =   "Emision de Etiquetas  M.P. de Reventa"
      End
      Begin VB.Menu ListEnv1 
         Caption         =   "Listado de Envases por Cliente"
      End
      Begin VB.Menu ListEnv2 
         Caption         =   "Listado de Envases por Envases"
      End
      Begin VB.Menu Verifica 
         Caption         =   "Listado de Verificacion de Correlatividades"
      End
      Begin VB.Menu Listcomp 
         Caption         =   "Listado de Componentes de Formulas (M.P.)"
      End
      Begin VB.Menu Listcomp1 
         Caption         =   "Listado de Componentes de Formulas (P.T.)"
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
      Begin VB.Menu Costo 
         Caption         =   "Listado de Costos por Partida"
      End
      Begin VB.Menu Valo1 
         Caption         =   "Listado de Valorizacion de Stock de M.P. a fecha"
      End
      Begin VB.Menu Valo2 
         Caption         =   "Listado de Valorizacion de Stock de P.T. a fecha"
      End
      Begin VB.Menu ConsFicMatAnt 
         Caption         =   "Consulta de Ficha de Materia Prima Historica"
      End
      Begin VB.Menu ConsFicTerAnt 
         Caption         =   "Consulta de Ficha de Producto Terminado Historico"
      End
      Begin VB.Menu Pedpenter 
         Caption         =   "Listado de Pedidos Pendientes por Producto Terminado"
      End
      Begin VB.Menu ConsumoMatII 
         Caption         =   "Listado de Analisis de Consumo de Materia Prima (Total Salidas)"
      End
      Begin VB.Menu ConsumoTerII 
         Caption         =   "Listado de Analisis de Consumo de Producto Terminado  (Total Salidas)"
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
      Begin VB.Menu EtiVerde 
         Caption         =   "Emision de Etiquetas Verdes"
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
      Begin VB.Menu PedPenDy 
         Caption         =   "Listado de Pedido Pendiente de M.P. de Reventa"
      End
      Begin VB.Menu ListStkFamDy 
         Caption         =   "Listado de Stock por Familia (DY)"
      End
      Begin VB.Menu ListDispoDy 
         Caption         =   "Listado de Disponibilidad de Stock"
      End
      Begin VB.Menu StkMinDy 
         Caption         =   "Listado de Stock Minimo de M.P. de Reventa"
      End
      Begin VB.Menu ProyStkDy 
         Caption         =   "Proyeccion de Stock"
      End
      Begin VB.Menu OrdPrnDy 
         Caption         =   "Listado de Ordenes Compras Pendientes por Planta (Orden)"
      End
      Begin VB.Menu LIstaProyec 
         Caption         =   "Listado de Proyeccion de Materia Prima de Reventa"
      End
      Begin VB.Menu OrdenAnual 
         Caption         =   "Listado de Ordenes de Compra Anuales"
      End
      Begin VB.Menu ListProydyii 
         Caption         =   "Listado de Proyeccion de Ordenes de Compra"
      End
      Begin VB.Menu ListaProydyiiPanta 
         Caption         =   "Actualizacion de Proyeccion de Ordenes de Compra"
      End
      Begin VB.Menu OrdPrnDyII 
         Caption         =   "Listado de Ordenes Compras Pendientes de DY / DW / DS  (Articulo)"
      End
      Begin VB.Menu SolZona 
         Caption         =   "Ingreso de Solciitud de Mercaderia a Zona Franca"
      End
      Begin VB.Menu ListStkFamDw 
         Caption         =   "Listado de Stock por Familia (DW)"
      End
      Begin VB.Menu OrdenDy 
         Caption         =   "Ingreso de Pedidos de Importaciones"
      End
      Begin VB.Menu ListStkDyPedido 
         Caption         =   "Listado de Analisis de Cortes de Pedido"
      End
      Begin VB.Menu ListStkFamDs 
         Caption         =   "Listado de Stock por Familia (DS)"
      End
      Begin VB.Menu ListStkFamDq 
         Caption         =   "Listado de Stock por Familia (DQ)"
      End
      Begin VB.Menu Seguimiento 
         Caption         =   "Seguimiento de Ordenes de Compra de Importacion"
      End
      Begin VB.Menu DepuraSaldosOrden 
         Caption         =   "Depuracion de Saldos de Ordenes de Compra"
      End
      Begin VB.Menu DepuraSaldosInforme 
         Caption         =   "Depuracion de Saldos de Informes de Recepcion"
      End
      Begin VB.Menu ListaDerechos 
         Caption         =   "Listado de % de Derechos"
      End
      Begin VB.Menu ValuacionPartida 
         Caption         =   "Calculo de Costo Promedio"
      End
      Begin VB.Menu ListaImpoFechaII 
         Caption         =   "Listado de Valorizacion de Importaciones a Fecha (Dy/Ds/Dw)"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu CierreStk 
         Caption         =   "Cierre del Stock"
         Enabled         =   0   'False
      End
      Begin VB.Menu Proc1 
         Caption         =   "Reproceso de Materia Prima"
      End
      Begin VB.Menu Proc2 
         Caption         =   "Reproceso de Producto Terminado"
      End
      Begin VB.Menu RepPedpen 
         Caption         =   "Reproceso de Pedidos"
      End
      Begin VB.Menu PasaDy 
         Caption         =   "Pasa de CO a DY/DW"
      End
      Begin VB.Menu ProcesoFecha 
         Caption         =   "Reproceso de Fechas de Laudos y Hojas"
      End
      Begin VB.Menu Proc11 
         Caption         =   "Generacion de NK y Re"
      End
      Begin VB.Menu Proc101 
         Caption         =   "Verificacion de MP"
      End
      Begin VB.Menu Proc102 
         Caption         =   "Veriricacion de PT"
      End
      Begin VB.Menu Verilot1 
         Caption         =   "Verificacion de Lotes de M.P."
      End
      Begin VB.Menu Verilot2 
         Caption         =   "Verificacion de Lotes de P.T."
      End
      Begin VB.Menu ProcHoja 
         Caption         =   "Verificacion de Hojas de Produccion"
      End
      Begin VB.Menu FinCot 
         Caption         =   "Fin del Sistema"
      End
      Begin VB.Menu pasadw 
         Caption         =   "Pasa Pt a DW"
      End
      Begin VB.Menu pasadtdy 
         Caption         =   "Pasa PT a DY"
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

Private Sub asdfas_Click()
    PrgCentroImportacion.Show
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
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    
    
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
    If Val(Wempresa) = 6 Then
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
    Rem If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
        PrgHoja.Show
    Rem End If
End Sub

Private Sub HojaProduccion_Click()
    PrgHojaProduccion.Show
End Sub

Private Sub Homologa_Click()
    PrgHomologaProve.Show
End Sub

Private Sub Importa_Click()
    PrgImporta.Show
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

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    Rem If MiRuta = "" Then
    Rem     MiRuta = CurDir + "\"
    Rem     MiRutaII = Left$(CurDir, 1)
    Rem End If
    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
    If Wempresa = "" Then
    
        Wempresa = "0001"
        Rem Empresa.Show
        Rem Empresa.SetFocus
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With

            Else
            
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
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
        WAtributo2 = rstAtributo!Atributo2 + "000000000000000000000000000000000"
        WAtributo3 = rstAtributo!Atributo3 + "000000000000000000000000000000000"
        WAtributo4 = rstAtributo!Atributo4 + "000000000000000000000000000000000"
        WAtributo5 = rstAtributo!Atributo5 + "000000000000000000000000000000000"
        WAtributo6 = rstAtributo!Atributo6 + "000000000000000000000000000000000"
        WAtributo7 = rstAtributo!Atributo7 + "000000000000000000000000000000000"
        WAtributo8 = rstAtributo!Atributo8 + "000000000000000000000000000000000"
        WAtributo9 = rstAtributo!Atributo9 + "000000000000000000000000000000000"
        WAtributo10 = rstAtributo!Atributo10 + "0000000000000000000000000000000"
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
        For Ciclo1 = 1 To 33
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
   Next Ciclo1
    Next Ciclo
            
    Menu.Envases.Enabled = Atri(1, 1)
    Menu.Arti.Enabled = Atri(1, 2)
    Menu.Terminado.Enabled = Atri(1, 3)
    Menu.Prove.Enabled = Atri(1, 4)
    Menu.Efluentes.Enabled = Atri(1, 5)
    Menu.Homologa.Enabled = Atri(1, 6)
    Rem by nan
    Menu.ProveAdicional.Enabled = Atri(1, 7)
    Menu.reprocesoactuaorden.Enabled = Atri(1, 8)
    Menu.Posicion.Enabled = Atri(1, 9)
    Menu.PrgActuaFactuExpoasd.Enabled = Atri(1, 10)
    Rem end by nan
    Menu.Cotiza.Enabled = Atri(2, 1)
    Menu.Orden.Enabled = Atri(2, 2)
    Menu.Informe.Enabled = Atri(2, 3)
    Menu.Laudo.Enabled = Atri(2, 4)
    Menu.Hoja.Enabled = Atri(2, 5)
    Menu.Movvar.Enabled = Atri(2, 6)
    Menu.MovEnv.Enabled = Atri(2, 7)
    Menu.Pedeti.Enabled = Atri(2, 8)
    Menu.Movguia.Enabled = Atri(2, 9)
    Menu.Prestamo.Enabled = Atri(2, 10)
    Menu.Actualiza.Enabled = Atri(2, 11)
    Menu.ModPedExp.Enabled = Atri(2, 12)
    Menu.Solic.Enabled = Atri(2, 13)
    Menu.Mirasol.Enabled = Atri(2, 14)
    Menu.Entdev.Enabled = Atri(2, 15)
    Menu.SolHoja.Enabled = Atri(2, 16)
    Menu.ModifColor.Enabled = Atri(2, 17)
    Menu.Centro.Enabled = Atri(2, 18)
    Menu.SolGuia.Enabled = Atri(2, 19)
    Menu.Mirasolguia.Enabled = Atri(2, 20)
    Menu.ActualizaInforme.Enabled = Atri(2, 21)
    Menu.DepuraMp.Enabled = Atri(2, 22)
    Menu.DepuraPt.Enabled = Atri(2, 23)
    Menu.ALtaProt.Enabled = Atri(2, 24)
    Menu.Insumo.Enabled = Atri(2, 25)
    Menu.mirainsumos.Enabled = Atri(2, 26)
    Menu.mirainsumosiii.Enabled = Atri(2, 27)
    Menu.VerificaPedido.Enabled = Atri(2, 28)
    Menu.CargaSolicitud.Enabled = Atri(2, 29)
    Menu.ActuaCargaSolicitud.Enabled = Atri(2, 30)
    Menu.HojaProduccion.Enabled = Atri(2, 31)
    
    Menu.ListCot.Enabled = Atri(3, 1)
    Menu.ListOrd.Enabled = Atri(3, 2)
    Menu.Cotprv.Enabled = Atri(3, 3)
    Menu.Cotart.Enabled = Atri(3, 4)
    Menu.Orden1.Enabled = Atri(3, 5)
    Menu.Orden2.Enabled = Atri(3, 6)
    Menu.Listmat1.Enabled = Atri(3, 7)
    Menu.Listmat2.Enabled = Atri(3, 8)
    Menu.Listter.Enabled = Atri(3, 9)
    Menu.ListArt1.Enabled = Atri(3, 10)
    Menu.ListTer1.Enabled = Atri(3, 11)
    Menu.Minimo.Enabled = Atri(3, 12)
    Menu.MInter.Enabled = Atri(3, 13)
    Menu.Compo.Enabled = Atri(3, 14)
    Menu.Proy.Enabled = Atri(3, 15)
    Menu.FichaMp.Enabled = Atri(3, 16)
    Menu.FiechaPt.Enabled = Atri(3, 17)
    Menu.Movvar1.Enabled = Atri(3, 18)
    Menu.Movvar2.Enabled = Atri(3, 19)
    Menu.Listhoja.Enabled = Atri(3, 20)
    Menu.Minimo1.Enabled = Atri(3, 21)
    Menu.MInter1.Enabled = Atri(3, 22)
    Menu.ConsumoTer.Enabled = Atri(3, 23)
    Menu.ConsumoArt.Enabled = Atri(3, 24)
    Menu.DispoPt.Enabled = Atri(3, 25)
    Menu.LIstaCarga.Enabled = Atri(3, 26)
    Menu.DispoPtOtro.Enabled = Atri(3, 27)
    Menu.AnalisisMp.Enabled = Atri(3, 28)
    Menu.AnalisisPt.Enabled = Atri(3, 29)
    Menu.ListaImpoFecha.Enabled = Atri(3, 30)
    
    Menu.CompPrv.Enabled = Atri(4, 1)
    Menu.CompMat.Enabled = Atri(4, 2)
    Menu.ConArtCon.Enabled = Atri(4, 3)
    Menu.ListInf.Enabled = Atri(4, 4)
    Menu.LIstInfPend.Enabled = Atri(4, 5)
    Menu.ListCont.Enabled = Atri(4, 6)
    Menu.ConsFichaMp.Enabled = Atri(4, 7)
    Menu.ConFichaPt.Enabled = Atri(4, 8)
    Menu.Ultima.Enabled = Atri(4, 9)
    Menu.Eti1.Enabled = Atri(4, 10)
    Menu.Eti4.Enabled = Atri(4, 11)
    Menu.ListEnv1.Enabled = Atri(4, 12)
    Menu.ListEnv2.Enabled = Atri(4, 13)
    Menu.Verifica.Enabled = Atri(4, 14)
    Menu.Listcomp.Enabled = Atri(4, 15)
    Menu.Listcomp1.Enabled = Atri(4, 16)
    Menu.Lotemat.Enabled = Atri(4, 17)
    Menu.Loteter.Enabled = Atri(4, 18)
    Menu.PedPen.Enabled = Atri(4, 19)
    Menu.Costo.Enabled = Atri(4, 20)
    Menu.Valo1.Enabled = Atri(4, 21)
    Menu.Valo2.Enabled = Atri(4, 22)
    Menu.ConsFicMatAnt.Enabled = Atri(4, 23)
    Menu.ConsFicTerAnt.Enabled = Atri(4, 24)
    Menu.Pedpenter.Enabled = Atri(4, 25)
    Menu.ConsumoMatII.Enabled = Atri(4, 26)
    Menu.ConsumoTerII.Enabled = Atri(4, 27)
    Menu.ListInfImporta.Enabled = Atri(4, 28)
    Menu.HistoriaTerminado.Enabled = Atri(4, 29)
    Menu.HistoriaArticulo.Enabled = Atri(4, 30)
    Menu.EtiVerde.Enabled = Atri(4, 31)
    
    Menu.ConsFicMAtDy.Enabled = Atri(5, 1)
    Menu.ListFicMatDy.Enabled = Atri(5, 2)
    Menu.ListStkDy.Enabled = Atri(5, 3)
    Menu.PedPenDy.Enabled = Atri(5, 4)
    Menu.ListStkFamDy.Enabled = Atri(5, 5)
    Menu.ListDispoDy.Enabled = Atri(5, 6)
    Menu.StkMinDy.Enabled = Atri(5, 7)
    Menu.ProyStkDy.Enabled = Atri(5, 8)
    Menu.OrdPrnDy.Enabled = Atri(5, 9)
    Menu.LIstaProyec.Enabled = Atri(5, 10)
    Menu.OrdenAnual.Enabled = Atri(5, 11)
    Menu.ListProydyii.Enabled = Atri(5, 12)
    Menu.ListaProydyiiPanta.Enabled = Atri(5, 13)
    Menu.OrdPrnDyII.Enabled = Atri(5, 14)
    Menu.SolZona.Enabled = Atri(5, 15)
    Menu.ListStkFamDw.Enabled = Atri(5, 16)
    Menu.OrdenDy.Enabled = Atri(5, 17)
    Menu.ListStkDyPedido.Enabled = Atri(5, 18)
    Menu.ListStkFamDs.Enabled = Atri(5, 19)
    Menu.ListStkFamDq.Enabled = Atri(5, 20)
    Rem by nan
    Menu.Seguimiento.Enabled = Atri(5, 21)
    Menu.DepuraSaldosOrden.Enabled = Atri(5, 22)
    Menu.DepuraSaldosInforme.Enabled = Atri(5, 23)
    Menu.ListaDerechos.Enabled = Atri(5, 24)
    Menu.ValuacionPartida.Enabled = Atri(5, 25)
    Menu.ListStkFamDq.Enabled = Atri(5, 26)
    Menu.Seguimiento.Enabled = Atri(5, 27)
    Menu.DepuraSaldosInforme.Enabled = Atri(5, 28)
    Menu.ListaDerechos.Enabled = Atri(5, 29)
    Menu.ValuacionPartida.Enabled = Atri(5, 30)
    Menu.ListaImpoFechaII.Enabled = Atri(5, 31)
    
    
    Rem fin by nan
    Menu.CierreStk.Enabled = Atri(6, 1)
    Menu.Proc1.Enabled = Atri(6, 2)
    Menu.Proc2.Enabled = Atri(6, 3)
    Menu.RepPedpen.Enabled = Atri(6, 4)
    Menu.PasaDy.Enabled = Atri(6, 5)
    Menu.ProcesoFecha.Enabled = Atri(6, 6)
    Menu.Proc11.Enabled = Atri(6, 7)
    Menu.Proc101.Enabled = Atri(6, 8)
    Menu.Proc102.Enabled = Atri(6, 9)
    Menu.Verilot1.Enabled = Atri(6, 10)
    Menu.Verilot2.Enabled = Atri(6, 11)
    Menu.ProcHoja.Enabled = Atri(6, 12)
    Menu.FinCot.Enabled = 1
    Menu.pasadw.Enabled = Atri(6, 14)
    Menu.pasadtdy.Enabled = Atri(6, 15)

End Sub

Private Sub Trazabilidad_Click()
    PrgTrazabilidad.Show
End Sub

Private Sub TrazabilidadTraspa_Click()
    PrgTrazabilidadTraspa.Show
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
