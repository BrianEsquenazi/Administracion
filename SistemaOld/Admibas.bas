Attribute VB_Name = "SISTEMA"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Rem Global Const FILENAME = "LOCALIDAD.MDB"
Global rstPedido2 As Recordset
Global oldempresa As String
Global salidita As String
Global Responsable As String
Global solicit As String
Global vercolumna As Integer
Global wgraba2 As String
Global Ingreso As String
Global cont As String
Global cliente2 As String
Global descliente2 As String
Global contras As String
Global PATH_PROG As String
Global coderr As Integer
Global Ds(30) As Integer
Global Const FILE_TYPE = ""
Global Lote As String
Global TipoImpre As String
Global XIndice As Integer
Global Text As String
Global Auxi As String
Global Auxi1 As String
Global Auxi2 As String
Global Auxi3 As String
Global Validate As String
Global Cicla As Integer
Global WAuxi As Integer
Global XCol As Integer
Global XRow As Integer
Global Existe As String
Global Renglon As Integer
Global WProveedor As String
Global WTipo As String
Global WLetra As String
Global WPunto As String
Global WNumero As String
Global XProveedor As String
Global XTipo As String
Global XLetra As String
Global XPunto As String
Global XNumero As String
Global WImpo As Double
Global WCtaConcepto As String
Global Inicial As Double
Global Wempresa As String
Global WEmpresaVerifica As String
Global WEmpresaRevalida As String
Global WNombreEmpresa As String
Global PCliente As String
Global PTipo As String
Global PTerminado As String
Global PLote As String
Global WRecibo As String
Global WOPago As String
Global WXPed As String
Global WXSol As String
Global Pasalote As String
Global DbConnect$
Global DSN$
Global UID$
Global PWD$
Global DSQ$
Global ProcesoActivate As Integer
Global WFuncion As Integer
Global Wpasamuestra As String
Global WMuestra As String
Global WAtraso As String
Global WOt As String
Global WPosi1 As Integer
Global WPosi2 As Integer
Global WPosi3 As Integer
Global WOrigenPosi As Integer
Global TraspaDatos(100, 10) As String
Global PasaEmpresa As String
Global ColumnaOpcion As Integer
Global WOperador As Integer
Global WClaveOperador As String
Global WSectorOperador As String
Global WTransporteOperador As String
Global WMateriaOperador As String
Global WTextoSolicitante As String
Global MiRuta As String
Global MiRutaII As String
Global XEmpresa As String
Global WSalidaError As String
Global ZMuestraCol As Integer
Global ZMuestraRow As Integer
Global ZMuestraTopRow As Integer
Global ZSql As String
Global ZAno As String
Global ZTerminado As String
Global ZVersion As String
Global ZVersionII As String
Global WProcesoOrden As Integer
Global WProveedorOrden As String
Global WDesProveedorOrden As String
Global WVectorOrden(100, 10) As String
Global ZLotePruart As String
Global ZArticuloPruart As String
Global ZHojaProceso As String
Global ZEtapaProceso As String
Global ZTerminadoProceso As String
Global ZCantidadProceso As String
Global ZOperarioProceso As String
Global ZProgramaOrigen As String
Global ZLoteRevalida As String
Global ZFechaRevalida As String
Global ZFechaHoja As String
Global ZFechaVencimiento As String
Global ZArticuloRevalida As String
Global ZDesArticuloRevalida As String
Global ZNroRevalida As String
Global ZZOrigenProceso As Integer
Global ZCargaDesvio(100, 20) As String

Global WWClaveOperador As String

Global WWWPasaFoto As String
Global WPasaOrden As String
Global WPasaOrigen As Integer
Global WPasaInforme As String
Global WPasaProveedor As String
Global WPasaMoneda As Integer

Global WPasaNroInterno As String

Global WPasaEmpresa As Integer
Global WPasaCarpeta As Integer
Global WPasaUltimoFob As Double
Global WPasaFactor As Double
Global WPasaUltimoCosto As Double
Global WPasaUltimoTipo As Integer

Global ZZIb As Integer
Global ZZCodigoResponsable As Integer
Global WPasaTipo As String
Global WPasaAno As String
Global WPasaNumero As String
Global WPasaFob As String
Global WPasaCliente As String
Global WPasaTerminado As String
Global WPasaLote As String
Global WPasaCantidad As Double
Global WPasaTipoPedido As String

Global ZZProcesoFactura As Integer

Global ZZPasaProceso As Integer
Global ZZPasaProcesoAltaAgenda As Integer
Global ZZPasaProcesoFechaAgenda As String

Global ZZPasaFila As Integer
Global ZZPasaColumna As Integer
Global ZZPasaProcesoActualiza As String

Global ZLugarAtraso As Integer
Global ZHastaAtraso As Integer
Global ZAtraso(100, 10) As String
Global ZLLave As String
Global ZZPeligrosa As String
Global ZZPasaHoja As String
Global ZZPasaFecha As String
Global ZZPasaRemito As String
Global ZZPasaProveedor As String
Global ZZPasaOrden As String
Global ZZPasaEmpresa As String
Global ZZPasaDatos(1000, 10) As String
Global ZZTrabajaLote(100) As String
Global ZZPasaTerminado As String
Global ZZPasaCantidad As Double
Global ZZPasaClave As String
Global ZZPasaTipoPedido As String
Global ZZPasaAgenda As Integer
Global ZZPasaPaso As String
Global ZZPasaCliente As String

Global ZZPasaReciboRecibo As String
Global ZZPasaReciboCliente As String
Global ZZPasaReciboImpo As String

Global ZZOperadorResponsable As String
Global ZZOperadorResponsableNombre As String

Global EmpresaActual As String

Dim rstListaPtVencido As Recordset
Dim spListaPtVencido As String
Dim rstPosicion As Recordset
Dim spPosicion As String


'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------

Global DbsEmpresa As Database
Global DbsAdminis As Database
Global DbsVentas As Database
Global DbsCotiza As Database
Global DbsAuxi As Database
Global DbsAuxi1 As Database
Global DbsLaboratorio As Database
Global DbsCotizaciones As Database
Global DbsInve As Database

'definicion de tablas de base de datos de empresa

Global rstEmpresa As Recordset

'definicion de tablas de base de datos  de administracion
Global rstgraficinver As Recordset

Global rstSedronarProceso As Recordset
Global rstSedro As Recordset
Global rstCuenta As Recordset
Global RstProve As Recordset
Global rstBanco As Recordset
Global rstRecibos As Recordset
Global rstImputac As Recordset
Global rstCtaCtePrv As Recordset
Global rstImpCtaCtePrv As Recordset
Global rstRevisa As Recordset
Global rstIvaComp As Recordset
Global rstPagos As Recordset
Global rstpasamuestra As Recordset
Global rstImpreRec As Recordset
Global rstImpreDep As Recordset
Global rstImprePago As Recordset
Global rstImpreRetIb As Recordset
Global rstImpreRetGan As Recordset
Global rstImprePed As Recordset
Global rstImprePedCol As Recordset
Global rstDepositos As Recordset
Global rstMovban As Recordset
Global rstPruebas As Recordset
Global rstValcar As Recordset
Global rstStockDy As Recordset
Global rstInve As Recordset
Global rstIva As Recordset
Global rstAuxi As Recordset
Global rstAuxi1 As Recordset
Global rstAuxiliar As Recordset
Global rstProceso1 As Recordset
Global rstPosdat As Recordset
Global rstRetencion As Recordset
Global rstListado1 As Recordset
Global rstAduana As Recordset
Global rstImpcyb As Recordset
Global rstFraseH As Recordset
Global rstFraseP As Recordset
Global rstDatosEtiqueta As Recordset
Global rstEmail As Recordset
Global rstpruefarma As Recordset
'definicion de tablas de base de datos de ventas

Global rstRubro As Recordset
Global rstVendedores As Recordset
Global rstPago As Recordset
Global rstCambios As Recordset
Global rstLineas As Recordset
Global rstEnvases As Recordset
Global rstTerminado As Recordset
Global rstClientes As Recordset
Global rstCash As Recordset
Global rstPrecios As Recordset
Global rstComposicion As Recordset
Global rstArticulo As Recordset
Global rstPedido As Recordset
Global rstPresta As Recordset
Global rstCtacte As Recordset
Global rstCtacte8 As Recordset
Global rstImpCtaCte As Recordset
Global rstImpCtaCteProy As Recordset
Global rstProyectaDias As Recordset
Global rstTrazabilidad As Recordset
Global rstEsta As Recordset
Global rstEsta1 As Recordset
Global rstEsta2 As Recordset
Global rstEsta8 As Recordset
Global rstEstaAnu As Recordset
Global rstEstaComando As Recordset
Global rstEstaAnuClie As Recordset
Global rstNumero As Recordset
Global rstNumero8 As Recordset
Global rstDesccomp As Recordset
Global rstTemperatura0 As Recordset
Global rstTemperatura1 As Recordset
Global rstTemperatura2 As Recordset
Global rstTemperatura3 As Recordset
Global rstTemperatura4 As Recordset
Global rstTemperatura5 As Recordset
Global rstTemperatura6 As Recordset
Global rstTemperatura7 As Recordset
Global rstEstadoReactor0 As Recordset
Global rstEstadoReactor1 As Recordset
Global rstEstadoReactor2 As Recordset
Global rstEstadoReactor3 As Recordset
Global rstEstadoReactor4 As Recordset
Global rstEstadoReactor5 As Recordset
Global rstEstadoReactor6 As Recordset
Global rstEstadoReactor7 As Recordset
Global rstFichaMat As Recordset
Global rstFichaTer As Recordset
Global rstFichaCon As Recordset
Global rstRanking As Recordset
Global rstMoreno As Recordset
Global rstMp As Recordset
Global rstPt As Recordset

'definicion de tablas de base de datos de ventas

Global rstCotiza As Recordset
Global rstCotiza1 As Recordset
Global rstOrden As Recordset
Global rstSolic As Recordset
Global rstWOrden As Recordset
Global rstInforme As Recordset
Global rstLaudo As Recordset
Global rstMovvar As Recordset
Global rstMovenv As Recordset
Global rstHoja As Recordset
Global rstEtiqueta As Recordset
Global rstEtiquetaII As Recordset
Global rstEtiquetaIII As Recordset
Global rstEtiquetaIV As Recordset
Global rstMoratoria As Recordset
Global rstProcesos As Recordset
Global rstImporta As Recordset
Global rstImpreEtiDy As Recordset
Global rstPedeti As Recordset
Global rstLiscot As Recordset
Global rstProyec As Recordset
Global rstControl As Recordset
Global rstVerifica As Recordset
Global rstFichaenv As Recordset
Global rstAvance As Recordset

'definicion de tablas de base de datos de laboartorio

Global rstEnsayos As Recordset
Global rstEspecificaciones As Recordset
Global rstEspecif As Recordset
Global rstPrueba As Recordset
Global rstPrueter As Recordset
Global rstMovlab As Recordset

Global rstInveMp As Recordset
Global rstInvePT As Recordset

'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------

Global Const TABLA_Empresa = "Empresa"

Global Const TABLA_SedronarProceso = "SedronarProceso"
Global Const TABLA_Sedro = "Sedro"
Global Const TABLA_Cuenta = "Cuenta"
Global Const TABLA_Prove = "Prove"
Global Const TABLA_Banco = "Banco"
Global Const TABLA_Recibos = "Recibos"
Global Const TABLA_Imputac = "Imputac"
Global Const TABLA_CtaCtePrv = "CtaCtePrv"
Global Const TABLA_ImpCtaCtePrv = "ImpCtaCtePrv"
Global Const TABLA_Revisa = "Revisa"
Global Const TABLA_IvaComp = "Ivacomp"
Global Const TABLA_Pagos = "Pagos"
Global Const TABLA_Depositos = "Depositos"
Global Const TABLA_Movban = "Movban"
Global Const TABLA_Pruebas = "Pruebas"
Global Const TABLA_Valcar = "Valcar"
Global Const TABLA_StockDy = "StockDy"
Global Const TABLA_Inve = "Inve"
Global Const TABLA_Iva = "Iva"
Global Const TABLA_Auxi = "Auxi"
Global Const TABLA_Auxi1 = "Auxi1"
Global Const TABLA_Auxiliar = "Auxiliar"
Global Const TABLA_Proceso1 = "Proceso1"
Global Const TABLA_Posdat = "Posdat"
Global Const TABLA_Retencion = "Retencion"
Global Const TABLA_Listado1 = "Listado1"
Global Const TABLA_Aduana = "Aduana"
Global Const TABLA_Impcyb = "Impcyb"
Global Const TABLA_FraseH = "FraseH"
Global Const TABLA_FraseP = "FraseP"
Global Const TABLA_DatosEtiqueta = "DatosEtiqueta"
Global Const TABLA_Email = "Email"
Global Const TABLA_Toto = "Toto"

Global Const TABLA_pasamuestra = "pasamuestra"
Global Const TABLA_Rubro = "Rubro"
Global Const TABLA_Vendedores = "Vendedores"
Global Const TABLA_ImpreRec = "ImpreRec"
Global Const TABLA_ImpreDep = "ImpreDep"
Global Const TABLA_ImprePago = "ImprePago"
Global Const TABLA_ImpreRetIb = "ImpreRetIb"
Global Const TABLA_ImpreRetGan = "ImpreRetGan"
Global Const TABLA_ImprePed = "ImprePed"
Global Const TABLA_ImprePedCol = "ImprePedCol"
Global Const TABLA_Pago = "Pago"
Global Const TABLA_Cambios = "Cambios"
Global Const TABLA_LINEAS = "Lineas"
Global Const TABLA_ENVASES = "Envases"
Global Const TABLA_TERMINADO = "Terminado"
Global Const TABLA_Clietes = "Clientes"
Global Const TABLA_Cash = "Cash"
Global Const TABLA_Precios = "Precios"
Global Const TABLA_COMPOSICION = "Composicion"
Global Const TABLA_Articulo = "Articulo"
Global Const TABLA_Pedido = "Pedido"
Global Const TABLA_Presta = "Presta"
Global Const TABLA_CtaCte = "CtaCte"
Global Const TABLA_CtaCte8 = "CtaCte8"
Global Const TABLA_ImpCtaCte = "ImpCtaCte"
Global Const TABLA_ImpCtaCteProy = "ImpCtaCteproy"
Global Const TABLA_ProyectaDias = "ProyectaDias"
Global Const TABLA_Trazabilidad = "Trazabilidad"
Global Const TABLA_Esta = "Esta"
Global Const TABLA_Esta1 = "Esta1"
Global Const TABLA_Esta2 = "Esta2"
Global Const TABLA_Esta8 = "Esta8"
Global Const TABLA_EstaAnu = "EstaAnu"
Global Const TABLA_EstaComando = "EstaComando"
Global Const TABLA_EstaAnuClie = "EstaAnuClie"
Global Const TABLA_Numero = "Numero"
Global Const TABLA_Numero8 = "Numero8"
Global Const TABLA_DescComp = "DescComp"
Global Const TABLA_Temperatura0 = "Temperatura0"
Global Const TABLA_Temperatura1 = "Temperatura1"
Global Const TABLA_Temperatura2 = "Temperatura2"
Global Const TABLA_Temperatura3 = "Temperatura3"
Global Const TABLA_Temperatura4 = "Temperatura4"
Global Const TABLA_Temperatura5 = "Temperatura5"
Global Const TABLA_Temperatura6 = "Temperatura6"
Global Const TABLA_Temperatura7 = "Temperatura7"
Global Const TABLA_EstadoReactor0 = "EstadoReactor0"
Global Const TABLA_EstadoReactor1 = "EstadoReactor1"
Global Const TABLA_EstadoReactor2 = "EstadoReactor2"
Global Const TABLA_EstadoReactor3 = "EstadoReactor3"
Global Const TABLA_EstadoReactor4 = "EstadoReactor4"
Global Const TABLA_EstadoReactor5 = "EstadoReactor5"
Global Const TABLA_EstadoReactor6 = "EstadoReactor6"
Global Const TABLA_EstadoReactor7 = "EstadoReactor7"
Global Const TABLA_FichaMat = "FichaMat"
Global Const TABLA_FichaTer = "FichaTer"
Global Const TABLA_FichaCon = "FichaCon"
Global Const TABLA_Ranking = "Ranking"
Global Const TABLA_Moreno = "Moreno"
Global Const TABLA_Mp = "Mp"
Global Const TABLA_Pt = "Pt"

Global Const TABLA_Cotiza1 = "Cotiza1"
Global Const TABLA_Cotiza = "Cotiza"
Global Const TABLA_Solic = "Solic"
Global Const TABLA_Orden = "Orden"
Global Const TABLA_WOrden = "WOrden"
Global Const TABLA_Informe = "Informe"
Global Const TABLA_Laudo = "Laudo"
Global Const TABLA_Movvar = "Movvar"
Global Const TABLA_MovEnv = "MovEnv"
Global Const TABLA_Hoja = "Hoja"
Global Const TABLA_Etiqueta = "Etiqueta"
Global Const TABLA_EtiquetaII = "EtiquetaII"
Global Const TABLA_EtiquetaIII = "EtiquetaIII"
Global Const TABLA_EtiquetaIV = "EtiquetaIV"
Global Const TABLA_Moratoria = "Moratoria"
Global Const TABLA_Procesos = "Procesos"
Global Const TABLA_Importa = "Importa"
Global Const TABLA_ImpreEtiDy = "ImpreEtiDy"
Global Const TABLA_Pedeti = "pedeti"
Global Const TABLA_Liscot = "Liscot"
Global Const TABLA_Proyec = "Proyec"
Global Const TABLA_Control = "Control"
Global Const TABLA_Verifica = "Verifica"
Global Const TABLA_Fichaenv = "Fichaenv"

Global Const TABLA_ENSAYOS = "ENSAYOS"
Global Const TABLA_Especificaciones = "Especificaciones"
Global Const TABLA_Especif = "Especif"
Global Const TABLA_PRUEBA = "PRUEBA"
Global Const TABLA_LOTE = "LOTE"
Global Const TABLA_PrueTer = "PrueTer"
Global Const TABLA_Movlab = "Movlab"

Global Const TABLA_InveMp = "InveMp"
Global Const TABLA_InvePt = "InvePt"
Global Const TABLA_prueterfarma = "prueterfarma"



'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE Vendedor
'--------------------------------------------------------
 
 Global Const Codigo = "CODIGO"
 Global Const Descripcion = "DESCRIPCION"
 
Sub OPEN_FILE_Empresa()
    Set DbsEmpresa = OpenDatabase("Empresa.mdb", False, False, FILE_TYPE)
    Set rstEmpresa = DbsEmpresa.OpenRecordset("Empresa")
End Sub
 
Sub OPEN_FILE_Cuenta()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCuenta = DbsAdminis.OpenRecordset("Cuenta")
End Sub

Sub OPEN_FILE_Prove()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set RstProve = DbsAuxi.OpenRecordset("Proveedor")
End Sub

Sub OPEN_FILE_Banco()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstBanco = DbsAdminis.OpenRecordset("Banco")
End Sub

Sub OPEN_FILE_Recibos()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRecibos = DbsAdminis.OpenRecordset("Recibos")
End Sub

Sub OPEN_FILE_Imputac()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImputac = DbsAuxi.OpenRecordset("Imputac")
End Sub

Sub OPEN_FILE_CtaCtePrv()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCtaCtePrv = DbsAdminis.OpenRecordset("CtaCtePrv")
End Sub

Sub OPEN_FILE_ImpCtaCtePrv()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpCtaCtePrv = DbsAuxi.OpenRecordset("CtaCtePrv")
End Sub

Sub OPEN_FILE_Revisa()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstRevisa = DbsAuxi.OpenRecordset("Revisa")
End Sub

Sub OPEN_FILE_Ivacomp()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstIvaComp = DbsAdminis.OpenRecordset("IvaComp")
End Sub

Sub OPEN_FILE_Pagos()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPagos = DbsAdminis.OpenRecordset("Pagos")
End Sub

Sub OPEN_FILE_Depositos()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstDepositos = DbsAdminis.OpenRecordset("Depositos")
End Sub

Sub OPEN_FILE_Movban()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstMovban = DbsAuxi.OpenRecordset("Movban")
End Sub

Sub OPEN_FILE_Pruebas()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstPruebas = DbsAuxi.OpenRecordset("Pruebas")
End Sub

Sub OPEN_FILE_Valcar()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstValcar = DbsAuxi.OpenRecordset("Valcar")
End Sub

Sub OPEN_FILE_StockDy()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstStockDy = DbsAuxi.OpenRecordset("StockDy")
End Sub

Sub OPEN_FILE_Inve()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstInve = DbsAuxi.OpenRecordset("Inve")
End Sub

Sub OPEN_FILE_Iva()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstIva = DbsAuxi.OpenRecordset("Iva")
End Sub

Sub OPEN_FILE_Auxi()
    Set DbsAdminis = OpenDatabase(Wempresa + "admi.mdb", False, False, FILE_TYPE)
    Set rstAuxi = DbsAdminis.OpenRecordset("Auxiliar")
End Sub

Sub OPEN_FILE_Auxiliar()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstAuxiliar = DbsAuxi.OpenRecordset("Auxiliar")
End Sub

Sub OPEN_FILE_Proceso1()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstProceso1 = DbsAdminis.OpenRecordset("Proceso1")
End Sub

Sub OPEN_FILE_Posdat()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstPosdat = DbsAuxi.OpenRecordset("Posdat")
End Sub

Sub OPEN_FILE_Retencion()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRetencion = DbsAdminis.OpenRecordset("Retencion")
End Sub

Sub OPEN_FILE_Listado1()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstListado1 = DbsAuxi.OpenRecordset("Listado1")
End Sub

Sub OPEN_FILE_Aduana()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstAduana = DbsAuxi.OpenRecordset("Aduana")
End Sub

Sub OPEN_FILE_Impcyb()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpcyb = DbsAuxi.OpenRecordset("Impcyb")
End Sub

Sub OPEN_FILE_FraseH()
    Set DbsAuxi = OpenDatabase("Eti.mdb", False, False, FILE_TYPE)
    Set rstFraseH = DbsAuxi.OpenRecordset("FraseH")
End Sub

Sub OPEN_FILE_FraseP()
    Set DbsAuxi = OpenDatabase("Eti.mdb", False, False, FILE_TYPE)
    Set rstFraseP = DbsAuxi.OpenRecordset("FraseP")
End Sub

Sub OPEN_FILE_Email()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstEmail = DbsAuxi.OpenRecordset("Email")
End Sub

Sub OPEN_FILE_Toto()
    Set DbsAdminis = OpenDatabase(Wempresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstToto = DbsAdminis.OpenRecordset("Toto")
End Sub

Sub OPEN_FILE_Rubro()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstRubro = DbsVentas.OpenRecordset("Rubro")
End Sub

Sub OPEN_FILE_Vendedores()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstVendedores = DbsVentas.OpenRecordset("Vendedores")
End Sub

Sub OPEN_FILE_ImpreRec()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpreRec = DbsVentas.OpenRecordset("ImpreRec")
End Sub

Sub OPEN_FILE_ImprePago()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImprePago = DbsVentas.OpenRecordset("ImprePago")
End Sub

Sub OPEN_FILE_ImpreRetIb()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpreRetIb = DbsVentas.OpenRecordset("ImpreRetIb")
End Sub

Sub OPEN_FILE_ImpreRetGan()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpreRetGan = DbsVentas.OpenRecordset("ImpreRetGan")
End Sub

Sub OPEN_FILE_ImpreDep()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpreDep = DbsVentas.OpenRecordset("ImpreDep")
End Sub

Sub OPEN_FILE_ImprePed()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImprePed = DbsVentas.OpenRecordset("ImprePed")
End Sub

Sub OPEN_FILE_ImprePedCol()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImprePedCol = DbsVentas.OpenRecordset("ImprePedCol")
End Sub

Sub OPEN_FILE_pasamuestra()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstpasamuestra = DbsVentas.OpenRecordset("Muestra")
End Sub

Sub OPEN_FILE_Pago()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstPago = DbsVentas.OpenRecordset("Pago")
End Sub

Sub OPEN_FILE_Cambios()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstCambios = DbsVentas.OpenRecordset("Cambios")
End Sub

Sub OPEN_FILE_LINEAS()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstLineas = DbsVentas.OpenRecordset("Lineas")
End Sub

Sub OPEN_FILE_ENVASES()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstEnvases = DbsVentas.OpenRecordset("Envases")
End Sub

Sub OPEN_FILE_TERMINADO()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstTerminado = DbsVentas.OpenRecordset("Terminado")
End Sub

Sub OPEN_FILE_Clientes()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstClientes = DbsVentas.OpenRecordset("Cliente")
End Sub

Sub OPEN_FILE_CASH()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstCash = DbsAuxi.OpenRecordset("Cash")
End Sub

Sub OPEN_FILE_Precios()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstPrecios = DbsVentas.OpenRecordset("Precios")
End Sub

Sub OPEN_FILE_Composicion()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstComposicion = DbsVentas.OpenRecordset("Composicion")
End Sub

Sub OPEN_FILE_Articulo()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstArticulo = DbsVentas.OpenRecordset("Articulo")
End Sub

Sub OPEN_FILE_Presta()
    Set DbsVentas = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstPresta = DbsVentas.OpenRecordset("Prestamo")
End Sub

Sub OPEN_FILE_Pedido2()
    Set DbsVentas = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
    Set rstPedido2 = DbsVentas.OpenRecordset("Pedido")
End Sub

Sub OPEN_FILE_Pedido()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstPedido = DbsVentas.OpenRecordset("Pedido")
End Sub

Sub OPEN_FILE_Ctacte()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstCtacte = DbsVentas.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Ctacte8()
    Select Case Val(Wempresa)
        Case 1
            Set DbsVentas = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstCtacte8 = DbsVentas.OpenRecordset("Ctacte1")
        Case 8
            Set DbsVentas = OpenDatabase("F:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstCtacte8 = DbsVentas.OpenRecordset("Ctacte1")
        Case Else
            Set DbsVentas = OpenDatabase("c:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstCtacte8 = DbsVentas.OpenRecordset("Ctacte1")
    End Select
End Sub

Sub OPEN_FILE_ImpCtacte()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpCtaCte = DbsAuxi.OpenRecordset("ImpCtacte")
End Sub

Sub OPEN_FILE_ImpCtacteProy()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstImpCtaCteProy = DbsAuxi.OpenRecordset("ImpCtacteProy")
End Sub

Sub OPEN_FILE_ProyectaDias()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstProyectaDias = DbsAuxi.OpenRecordset("ProyectaDias")
End Sub

Sub OPEN_FILE_Trazabilidad()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstTrazabilidad = DbsAuxi.OpenRecordset("Trazabilidad")
End Sub

Sub OPEN_FILE_Esta()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstEsta = DbsAuxi.OpenRecordset("Estadistica")
End Sub

Sub OPEN_FILE_Esta1()
    Select Case Val(Wempresa)
        Case 1
            Set DbsAuxi = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta1 = DbsAuxi.OpenRecordset("Estadistica")
        Case 8
            Rem dada
            Set DbsAuxi = OpenDatabase("e:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta1 = DbsAuxi.OpenRecordset("Estadistica")
        Case Else
            Set DbsAuxi = OpenDatabase("d:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta1 = DbsAuxi.OpenRecordset("Estadistica")
    End Select
End Sub

Sub OPEN_FILE_Esta2()
    Select Case Val(Wempresa)
        Case 1
            Set DbsAuxi = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta2 = DbsAuxi.OpenRecordset("Estadif")
        Case 8
            Rem dada
            Set DbsAuxi = OpenDatabase("e:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta2 = DbsAuxi.OpenRecordset("Estadif")
        Case Else
            Set DbsAuxi = OpenDatabase("d:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta2 = DbsAuxi.OpenRecordset("Estadif")
    End Select
End Sub

Sub OPEN_FILE_Esta8()
    Select Case Val(Wempresa)
        Case 1
            Set DbsAuxi = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta8 = DbsAuxi.OpenRecordset("Esta1")
        Case 8
            Set DbsAuxi = OpenDatabase("e:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta8 = DbsAuxi.OpenRecordset("Esta1")
        Case Else
            Set DbsAuxi = OpenDatabase("c:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstEsta8 = DbsAuxi.OpenRecordset("Esta1")
    End Select
End Sub

Sub OPEN_FILE_EstaAnu()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstEstaAnu = DbsAuxi.OpenRecordset("EstaAnu")
End Sub

Sub OPEN_FILE_EstaComando()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstEstaComando = DbsAuxi.OpenRecordset("EstaComando")
End Sub

Sub OPEN_FILE_EstaAnuClie()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstEstaAnuClie = DbsAuxi.OpenRecordset("EstaAnuClie")
End Sub

Sub OPEN_FILE_Numero()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstNumero = DbsVentas.OpenRecordset("Numero")
End Sub

Sub OPEN_FILE_Numero8()
    Select Case Val(Wempresa)
        Case 1
            Set DbsVentas = OpenDatabase("\\193.168.0.2\f\vb\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstNumero8 = DbsVentas.OpenRecordset("Numero")
        Case 8
            Set DbsVentas = OpenDatabase("F:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstNumero8 = DbsVentas.OpenRecordset("Numero")
        Case Else
            Set DbsVentas = OpenDatabase("c:\vende\0001vend.mdb", False, False, FILE_TYPE)
            Set rstNumero8 = DbsVentas.OpenRecordset("Numero")
    End Select
End Sub

Sub OPEN_FILE_DescComp()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstDesccomp = DbsVentas.OpenRecordset("DescComp")
End Sub

Sub OPEN_FILE_Temperatura0()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura0 = DbsAuxi.OpenRecordset("TempReac0")
End Sub

Sub OPEN_FILE_Temperatura1()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura1 = DbsAuxi.OpenRecordset("TempReac1")
End Sub

Sub OPEN_FILE_Temperatura2()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura2 = DbsAuxi.OpenRecordset("TempReac2")
End Sub

Sub OPEN_FILE_Temperatura3()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura3 = DbsAuxi.OpenRecordset("TempReac3")
End Sub

Sub OPEN_FILE_Temperatura4()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura4 = DbsAuxi.OpenRecordset("TempReac4")
End Sub

Sub OPEN_FILE_Temperatura5()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura5 = DbsAuxi.OpenRecordset("TempReac5")
End Sub

Sub OPEN_FILE_Temperatura6()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura6 = DbsAuxi.OpenRecordset("TempReac6")
End Sub

Sub OPEN_FILE_Temperatura7()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstTemperatura7 = DbsAuxi.OpenRecordset("TempReac7")
End Sub

Sub OPEN_FILE_EstadoReactor0()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor0 = DbsAuxi.OpenRecordset("EstReac0")
End Sub

Sub OPEN_FILE_EstadoReactor1()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor1 = DbsAuxi.OpenRecordset("EstReac1")
End Sub

Sub OPEN_FILE_EstadoReactor2()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor2 = DbsAuxi.OpenRecordset("EstReac2")
End Sub

Sub OPEN_FILE_EstadoReactor3()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor3 = DbsAuxi.OpenRecordset("EstReac3")
End Sub

Sub OPEN_FILE_EstadoReactor4()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor4 = DbsAuxi.OpenRecordset("EstReac4")
End Sub

Sub OPEN_FILE_EstadoReactor5()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor5 = DbsAuxi.OpenRecordset("EstReac5")
End Sub

Sub OPEN_FILE_EstadoReactor6()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor6 = DbsAuxi.OpenRecordset("EstReac6")
End Sub

Sub OPEN_FILE_EstadoReactor7()
    Rem Set DbsAuxi = OpenDatabase("\\193.168.0.126\Principal\Base.mdb", False, False, FILE_TYPE)
    Set DbsAuxi = OpenDatabase("Base.mdb", False, False, FILE_TYPE)
    Set rstEstadoReactor7 = DbsAuxi.OpenRecordset("EstReac7")
End Sub

Sub OPEN_FILE_FichaMat()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstFichaMat = DbsAuxi.OpenRecordset("FichaMat")
End Sub

Sub OPEN_FILE_Sedro()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstSedro = DbsAuxi.OpenRecordset("Sedronar")
End Sub

Sub OPEN_FILE_SedronarProceso()
    Select Case Val(Wempresa)
        Case 1
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoI")
        Case 2
            Set DbsAuxi = OpenDatabase("SedronarII.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoI")
        Case 3
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoII")
        Case 4
            Set DbsAuxi = OpenDatabase("SedronarII.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoII")
        Case 5
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoIII")
        Case 7
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoV")
        Case 8
            Set DbsAuxi = OpenDatabase("SedronarII.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoIII")
        Case 9
            Set DbsAuxi = OpenDatabase("SedronarII.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoV")
        Case 10
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoVI")
        Case Else
            Set DbsAuxi = OpenDatabase("Sedronar.mdb", False, False, FILE_TYPE)
            Set rstSedronarProceso = DbsAuxi.OpenRecordset("SedronarProcesoVII")
    End Select
End Sub

Sub OPEN_FILE_FichaTer()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstFichaTer = DbsAuxi.OpenRecordset("FichaTer")
End Sub

Sub OPEN_FILE_FichaCon()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstFichaCon = DbsAuxi.OpenRecordset("FichaCon")
End Sub

Sub OPEN_FILE_Ranking()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstRanking = DbsAuxi.OpenRecordset("Ranking")
End Sub

Sub OPEN_FILE_Moreno()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstMoreno = DbsAuxi.OpenRecordset("Moreno")
End Sub

Sub OPEN_FILE_Mp()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstMp = DbsVentas.OpenRecordset("Mp")
End Sub

Sub OPEN_FILE_Pt()
    Set DbsVentas = OpenDatabase(Wempresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstPt = DbsVentas.OpenRecordset("Pt")
End Sub

Sub OPEN_FILE_Cotiza()
    Set DbsCotizaciones = OpenDatabase("Cotiza.mdb", False, False, FILE_TYPE)
    Set rstCotiza = DbsCotizaciones.OpenRecordset("Cotiza")
End Sub

Sub OPEN_FILE_Cotiza1()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstCotiza1 = DbsCotiza.OpenRecordset("Cotiza")
End Sub

Sub OPEN_FILE_Orden()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstOrden = DbsCotiza.OpenRecordset("Orden")
End Sub

Sub OPEN_FILE_WOrden()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstWOrden = DbsAuxi.OpenRecordset("WOrden")
End Sub

Sub OPEN_FILE_Informe()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstInforme = DbsCotiza.OpenRecordset("Informe")
End Sub

Sub OPEN_FILE_LAUDO()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstLaudo = DbsCotiza.OpenRecordset("Laudo")
End Sub

Sub OPEN_FILE_Movvar()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstMovvar = DbsCotiza.OpenRecordset("Movvar")
End Sub

Sub OPEN_FILE_MovEnv()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstMovenv = DbsCotiza.OpenRecordset("MovEnv")
End Sub

Sub OPEN_FILE_Hoja()
    Set DbsCotiza = OpenDatabase(Wempresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstHoja = DbsCotiza.OpenRecordset("Hoja")
End Sub

Sub OPEN_FILE_Etiqueta()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstEtiqueta = DbsAuxi.OpenRecordset("Etiqueta")
End Sub

Sub OPEN_FILE_EtiquetaII()
    Set DbsAuxi = OpenDatabase("etiprueba.mdb", False, False, FILE_TYPE)
    Set rstEtiquetaII = DbsAuxi.OpenRecordset("Etiqueta")
End Sub

Sub OPEN_FILE_EtiquetaIII()
    Set DbsAuxi = OpenDatabase("etiprueba.mdb", False, False, FILE_TYPE)
    Set rstEtiquetaIII = DbsAuxi.OpenRecordset("EtiquetaII")
End Sub

Sub OPEN_FILE_EtiquetaIV()
    Set DbsAuxi = OpenDatabase("etiprueba.mdb", False, False, FILE_TYPE)
    Set rstEtiquetaIV = DbsAuxi.OpenRecordset("EtiquetaIII")
End Sub

Sub OPEN_FILE_Moratoria()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstMoratoria = DbsAuxi.OpenRecordset("Moratoria")
End Sub

Sub OPEN_FILE_Procesos()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstProcesos = DbsAuxi.OpenRecordset("Procesos")
End Sub

Sub OPEN_FILE_Importa()
    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
        Set DbsAuxi = OpenDatabase("0001auxi.mdb", False, False, FILE_TYPE)
            Else
        Set DbsAuxi = OpenDatabase("0002auxi.mdb", False, False, FILE_TYPE)
    End If
    Set rstImporta = DbsAuxi.OpenRecordset("Importa")
End Sub

Sub OPEN_FILE_ImpreEtiDy()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstImpreEtiDy = DbsAuxi.OpenRecordset("ImpreEtiDy")
End Sub

Sub OPEN_FILE_Pedeti()
    Set DbsAuxi = OpenDatabase(Wempresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstPedeti = DbsAuxi.OpenRecordset("Pedeti")
End Sub

Sub OPEN_FILE_Liscot()
    Set DbsAuxi1 = OpenDatabase("0001Auxi.mdb", False, False, FILE_TYPE)
    Set rstLiscot = DbsAuxi1.OpenRecordset("Listcot")
End Sub

Sub OPEN_FILE_Proyec()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstProyec = DbsAuxi.OpenRecordset("Proyec")
End Sub

Sub OPEN_FILE_Control()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstControl = DbsAuxi.OpenRecordset("Control")
End Sub

Sub OPEN_FILE_Verifica()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstVerifica = DbsAuxi.OpenRecordset("Verifica")
End Sub

Sub OPEN_FILE_FichaEnv()
    Set DbsAuxi = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstFichaenv = DbsAuxi.OpenRecordset("Fichaenv")
End Sub

Sub OPEN_FILE_ENSAYOS()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEnsayos = DbsLaboratorio.OpenRecordset("ENSAYOS")
End Sub
 
Sub OPEN_FILE_Especificaciones()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEspecificaciones = DbsLaboratorio.OpenRecordset("Especificaciones")
End Sub

Sub OPEN_FILE_Especif()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEspecif = DbsLaboratorio.OpenRecordset("Especif")
End Sub
 
Sub OPEN_FILE_PRUEBA()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstPrueba = DbsLaboratorio.OpenRecordset("PRUEBA")
End Sub

Sub OPEN_FILE_PrueTer()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstPrueter = DbsLaboratorio.OpenRecordset("PrueTer")
End Sub

Sub OPEN_FILE_Movlab()
    Set DbsLaboratorio = OpenDatabase(Wempresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstMovlab = DbsLaboratorio.OpenRecordset("Movlab")
End Sub

Sub OPEN_FILE_InveMp()
    Set DbsInve = OpenDatabase("Inve.mdb", False, False, FILE_TYPE)
    Set rstInveMp = DbsInve.OpenRecordset("InveMp")
End Sub

Sub OPEN_FILE_InvePt()
    Set DbsInve = OpenDatabase("Inve.mdb", False, False, FILE_TYPE)
    Set rstInvePT = DbsInve.OpenRecordset("InvePt")
End Sub

Sub NumbersOnly(T As Control, KeyAscii As Integer)
'This Sub allows only the digits 0 to 9, an initial minus sign and one period.

If KeyAscii < Asc(" ") Then     ' Is this Control char?
    Exit Sub                    ' Yes, let it pass
End If
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
     'don't discard it
ElseIf KeyAscii = Asc(".") Then 'if its a period
     If InStr(1, T, ".") Then 'if there is already a period
          KeyAscii = 0   'discard it
     End If
ElseIf KeyAscii = Asc("-") And T.SelStart = 0 Then
     'keep it, it's an initial minus sign
Else
    KeyAscii = 0  ' Discard all other characters
End If
'Now prevent any characters in front of a minus sign
If Mid$(T.Text, T.SelStart + T.SelLength + 1, 1) = "-" Then
    KeyAscii = 0   ' Discard characters before -
End If

End Sub



Sub Errores(coderr As Integer, Archivo As String, Mensaje As String)

    e = coderr
    Select Case e
        Case 3021
            m$ = Mensaje$
            A% = MsgBox(m$, 0, "Archivo de " + Archivo)
        Case Else
            m$ = Mensaje$
            A% = MsgBox(m$, 0, Archivo)
    End Select
    
End Sub

Sub Ceros(Campo As String, largo As Integer)

    L% = 1
    cadena$ = ""
    While L% <= Len(Campo) And L% > 0
        If Mid$(Campo, L%, 1) <> Chr$(32) Then cadena$ = cadena$ + Mid$(Campo, L%, 1)
        L% = L% + 1
    Wend
    Campo = Right$(String$(40, "0") + cadena$, largo)
    
End Sub

Sub Calcula_CostoFactura(Producto As String, Costo As Double)

    Dim VectorCalculo(100, 2) As String
    Dim AuxiliarCalculo(100, 7) As String
    
    Erase AuxiliarCalculo
    Erase VectorCalculo
    
    Renglon = 0
    
    If Left$(Producto, 2) = "PT" Or Left$(Producto, 2) = "PE" Or Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
    
        If Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
            Producto = "PT" + Mid$(Producto, 3, 10)
        End If
        
        VectorCalculo(1, 1) = Producto
        VectorCalculo(1, 2) = "1"
        Costo = 0
        Lugar = 1
        Cicla = 0
        
        Do
            Cicla = Cicla + 1
            If VectorCalculo(Cicla, 1) <> "" Then
            
                Entra = "S"
        
                spComposicion = "ConsultaComposicionProducto " + "'" + VectorCalculo(Cicla, 1) + "'"
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                
                If rstComposicion.RecordCount > 0 Then
                With rstComposicion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Entra = "N"
                            
                            Tipo = rstComposicion!Tipo
                            Articulo1 = rstComposicion!Articulo1
                            Articulo2 = rstComposicion!Articulo2
                            Cantidad = rstComposicion!Cantidad
                            
                            Rem If Left$(Articulo1, 2) = "DW" Then
                            Rem     Tipo = "T"
                            Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                            Rem End If
                            
                            Select Case Tipo
                                Case "T"
                                    If Producto <> Articulo2 Then
                                        Lugar = Lugar + 1
                                        VectorCalculo(Lugar, 1) = Articulo2
                                        VectorCalculo(Lugar, 2) = Str$(Cantidad * Val(VectorCalculo(Cicla, 2)))
                                    End If
                                Case "M"
                                    Renglon = Renglon + 1
                                    AuxiliarCalculo(Renglon, 1) = Articulo1
                                    AuxiliarCalculo(Renglon, 2) = Cantidad
                                    AuxiliarCalculo(Renglon, 3) = VectorCalculo(Cicla, 2)
                                Case Else
                            End Select
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstComposicion.Close
                End If
                
                Rem If Entra = "S" And Left$(VectorCalculo(Cicla, 1), 2) = "DW" Then
                Rem     Renglon = Renglon + 1
                Rem     AuxiliarCalculo(Renglon, 1) = Left$(VectorCalculo(Cicla, 1), 3) + Right$(VectorCalculo(Cicla, 1), 7)
                Rem     AuxiliarCalculo(Renglon, 2) = 1
                Rem     AuxiliarCalculo(Renglon, 3) = VectorCalculo(Cicla, 2)
                Rem End If
                
                    Else
                    
                Exit Do
                
            End If
            
        Loop
        
        If Renglon > 0 Then
                        
            For Da = 1 To Renglon
                Articulo = AuxiliarCalculo(Da, 1)
                Cantidad = AuxiliarCalculo(Da, 2)
                XVectorCalculo = AuxiliarCalculo(Da, 3)
                
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVectorCalculo))
                    Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVectorCalculo))
                    rstArticulo.Close
                End If
            Next Da
            
                Else
                
            XArti = Left$(Producto, 3) + Right$(Producto, 7)
            spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Costo = rstArticulo!Costo2
                rstArticulo.Close
            End If
        
        End If
        
            Else
            
        XArti = Left$(Producto, 3) + Right$(Producto, 7)
        spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Costo = rstArticulo!Costo2
            rstArticulo.Close
        End If

    End If
    
End Sub


