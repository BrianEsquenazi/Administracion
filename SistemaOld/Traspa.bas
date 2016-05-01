Attribute VB_Name = "SISTEMA"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Rem Global Const FILENAME = "LOCALIDAD.MDB"

Global PATH_PROG As String
Global coderr As Integer
Global Ds(30) As Integer
Global Const FILE_TYPE = ""
Global Text As String
Global Auxi As String
Global Auxi1 As String
Global Auxi2 As String
Global Validate As String
Global Cicla As Integer
Global WAuxi As Integer
Global XCol As Integer
Global XRow As Integer
Global Existe As String
Global Renglon As Integer
Global Inicial As Double
Global WEmpresa As String
Global XEmpresa As String
Global WFecha As String
Global WFectraspaso As String
Global WHastaEmpresa As Integer
Global WDesdeEmpresa As Integer
Global DbConnect$
Global DSN$
Global UID$
Global PWD$
Global DSQ$
Global WXPed As String
Global ZSql As String

'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------


Global DbsEmpresa As Database
Global DbsVentas As Database
Global DbsCotiza As Database
Global DbsLaboratorio As Database
Global DbsTraspa As Database
Global DbsOtra As Database

Global rstEmpresa As Recordset

'definicion de tablas de base de datos de ventas

Global rstTerminado As Recordset
Global rstClientes As Recordset
Global rstClientes2 As Recordset
Global rstPrecios As Recordset
Global rstComposicion As Recordset
Global rstArticulo As Recordset
Global rstCtacte As Recordset
Global rstCtaCte2 As Recordset
Global rstCtaCte4 As Recordset
Global rstEstadistica As Recordset
Global rstDesccomp As Recordset

'definicion de tablas de base de datos de ventas

Global rstCotiza As Recordset
Global rstOrden As Recordset
Global rstInforme As Recordset
Global rstLaudo As Recordset
Global rstMovvar As Recordset
Global rstHoja As Recordset
Global rstEtiqueta As Recordset

'definicion de tablas de base de datos de laboartorio

Global rstEnsayos As Recordset
Global rstEspecificaciones As Recordset
Global rstEspecif As Recordset
Global rstPrueba As Recordset
Global rstPrueter As Recordset
Global rstMovlab As Recordset

'definicion de tablas de traspado de datos




Global rstWTerminado As Recordset
Global rstWClientes As Recordset
Global rstWProveedor As Recordset
Global rstWPrecios As Recordset
Global rstWComposicion As Recordset
Global rstWArticulo As Recordset
Global rstWCtaCte As Recordset
Global rstWCtaCte2 As Recordset
Global rstWCtaCte4 As Recordset
Global rstWEstadistica As Recordset
Global rstWDescComp As Recordset
Global rstWInforme As Recordset
Global rstWLaudo As Recordset
Global rstWMovvar As Recordset
Global rstWMovguia As Recordset
Global rstWHoja As Recordset
Global rstWEnsayos As Recordset
Global rstWEspecificaciones As Recordset
Global rstWEspecif As Recordset
Global rstWSolic As Recordset
Global rstWPrueba As Recordset
Global rstWPrueter As Recordset
Global rstWMovlab As Recordset

'definicion de tablas de traspado de datos

Global rstWCotiza As Recordset
Global rstWOrden As Recordset

'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------



Global Const TABLA_Empresa = "Empresa"

Global Const TABLA_TERMINADO = "Terminado"
Global Const TABLA_Clientes = "Clientes"
Global Const TABLA_Clientes2 = "Clientes2"
Global Const TABLA_Precios = "Precios"
Global Const TABLA_COMPOSICION = "Composicion"
Global Const TABLA_Articulo = "Articulo"
Global Const TABLA_CtaCte = "CtaCte"
Global Const TABLA_CtaCte2 = "CtaCte2"
Global Const TABLA_CtaCte4 = "CtaCte4"
Global Const TABLA_Estadistica = "Estadsitica"
Global Const TABLA_DescComp = "DescComp"
Global Const TABLA_Etiqueta = "Etiqueta"

Global Const TABLA_Cotiza = "Cotiza"
Global Const TABLA_Orden = "Orden"
Global Const TABLA_Informe = "Informe"
Global Const TABLA_Laudo = "Laudo"
Global Const TABLA_Movvar = "Movvar"
Global Const TABLA_Hoja = "Hoja"

Global Const TABLA_ENSAYOS = "ENSAYOS"
Global Const TABLA_ESPECIFICACIONES = "ESPECIFICAICONES"
Global Const TABLA_ESPECIF = "ESPECIF"
Global Const TABLA_PRUEBA = "PRUEBA"
Global Const TABLA_PrueTer = "PrueTer"
Global Const TABLA_Movlab = "Movlab"

Global Const TABLA_WTERMINADO = "WTerminado"
Global Const TABLA_WClietes = "WClientes"
Global Const TABLA_WProveedor = "WProveedor"
Global Const TABLA_WPrecios = "WPrecios"
Global Const TABLA_WCOMPOSICION = "WComposicion"
Global Const TABLA_WArticulo = "WArticulo"
Global Const TABLA_WCtaCte = "WCtaCte"
Global Const TABLA_WCtaCte2 = "WCtaCte2"
Global Const TABLA_WCtaCte4 = "WCtaCte4"
Global Const TABLA_WEstadistica = "WEstadsitica"
Global Const TABLA_WDescComp = "WDescComp"
Global Const TABLA_WInforme = "WInforme"
Global Const TABLA_WLaudo = "WLaudo"
Global Const TABLA_WMovvar = "WMovvar"
Global Const TABLA_WMovguia = "WMovguia"
Global Const TABLA_WHoja = "WHoja"
Global Const TABLA_WENSAYOS = "WENSAYOS"
Global Const TABLA_WESPECIFICACIONES = "WESPECIFICAICONES"
Global Const TABLA_WESPECIF = "WESPECIF"
Global Const TABLA_WSolic = "WSolic"
Global Const TABLA_WPRUEBA = "WPRUEBA"
Global Const TABLA_WPrueTer = "WPrueTer"
Global Const TABLA_WMovlab = "WMovlab"

Global Const TABLA_WCotiza = "WCotiza"
Global Const TABLA_WOrden = "WOrden"


Sub OPEN_FILE_Empresa()
    Set DbsEmpresa = OpenDatabase("Empresa.mdb", False, False, FILE_TYPE)
    Set rstEmpresa = DbsEmpresa.OpenRecordset("Empresa")
End Sub
 
Sub OPEN_FILE_TERMINADO()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstTerminado = DbsVentas.OpenRecordset("Terminado")
End Sub

Sub OPEN_FILE_Clientes()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstClientes = DbsVentas.OpenRecordset("Cliente")
End Sub

Sub OPEN_FILE_Clientes2()
    Set DbsVentas = OpenDatabase("0002vent.mdb", False, False, FILE_TYPE)
    Set rstClientes2 = DbsVentas.OpenRecordset("Cliente")
End Sub

Sub OPEN_FILE_Precios()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstPrecios = DbsVentas.OpenRecordset("Precios")
End Sub

Sub OPEN_FILE_Composicion()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstComposicion = DbsVentas.OpenRecordset("Composicion")
End Sub

Sub OPEN_FILE_Articulo()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstArticulo = DbsVentas.OpenRecordset("Articulo")
End Sub

Sub OPEN_FILE_Ctacte()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstCtacte = DbsVentas.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Ctacte2()
    Set DbsVentas = OpenDatabase("0002vent.mdb", False, False, FILE_TYPE)
    Set rstCtaCte2 = DbsVentas.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Ctacte4()
    Set DbsVentas = OpenDatabase("0004vent.mdb", False, False, FILE_TYPE)
    Set rstCtaCte4 = DbsVentas.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Estadistica()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstEstadistica = DbsVentas.OpenRecordset("Estadistica")
End Sub

Sub OPEN_FILE_DescComp()
    Set DbsVentas = OpenDatabase(WEmpresa + "vent.mdb", False, False, FILE_TYPE)
    Set rstDesccomp = DbsVentas.OpenRecordset("DescComp")
End Sub

Sub OPEN_FILE_Cotiza()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstCotiza = DbsCotiza.OpenRecordset("Cotiza")
End Sub

Sub OPEN_FILE_Etiqueta()
    Set DbsAuxi = OpenDatabase(WEmpresa + "auxi.mdb", False, False, FILE_TYPE)
    Set rstEtiqueta = DbsAuxi.OpenRecordset("Etiqueta")
End Sub

Sub OPEN_FILE_Orden()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstOrden = DbsCotiza.OpenRecordset("Orden")
End Sub

Sub OPEN_FILE_Informe()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstInforme = DbsCotiza.OpenRecordset("Informe")
End Sub

Sub OPEN_FILE_LAUDO()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstLaudo = DbsCotiza.OpenRecordset("Laudo")
End Sub

Sub OPEN_FILE_Movvar()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstMovvar = DbsCotiza.OpenRecordset("Movvar")
End Sub

Sub OPEN_FILE_Hoja()
    Set DbsCotiza = OpenDatabase(WEmpresa + "Coti.mdb", False, False, FILE_TYPE)
    Set rstHoja = DbsCotiza.OpenRecordset("Hoja")
End Sub

Sub OPEN_FILE_ENSAYOS()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEnsayos = DbsLaboratorio.OpenRecordset("ENSAYOS")
End Sub

Sub OPEN_FILE_ESPECIFICACIONES()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEspecificaciones = DbsLaboratorio.OpenRecordset("ESPECIFICACIONES")
End Sub

Sub OPEN_FILE_ESPECIF()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstEspecif = DbsLaboratorio.OpenRecordset("ESPECIF")
End Sub

Sub OPEN_FILE_PRUEBA()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstPrueba = DbsLaboratorio.OpenRecordset("PRUEART")
End Sub

Sub OPEN_FILE_PrueTer()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstPrueter = DbsLaboratorio.OpenRecordset("PrueTer")
End Sub

Sub OPEN_FILE_Movlab()
    Set DbsLaboratorio = OpenDatabase(WEmpresa + "labo.mdb", False, False, FILE_TYPE)
    Set rstMovlab = DbsLaboratorio.OpenRecordset("Movlab")
End Sub

Sub OPEN_FILE_WTERMINADO()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWTerminado = DbsTraspa.OpenRecordset("WTerminado")
End Sub

Sub OPEN_FILE_WClientes()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWClientes = DbsTraspa.OpenRecordset("WCliente")
End Sub

Sub OPEN_FILE_WProveedor()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWProveedor = DbsTraspa.OpenRecordset("WProveedor")
End Sub

Sub OPEN_FILE_WPrecios()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWPrecios = DbsTraspa.OpenRecordset("WPrecios")
End Sub

Sub OPEN_FILE_WComposicion()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWComposicion = DbsTraspa.OpenRecordset("WComposicion")
End Sub

Sub OPEN_FILE_WArticulo()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWArticulo = DbsTraspa.OpenRecordset("WArticulo")
End Sub

Sub OPEN_FILE_WCtacte()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWCtaCte = DbsTraspa.OpenRecordset("WCtacte")
End Sub

Sub OPEN_FILE_WCtacte2()
    Set DbsTraspa = OpenDatabase("0002tras.mdb", False, False, FILE_TYPE)
    Set rstWCtaCte2 = DbsTraspa.OpenRecordset("WCtacte")
End Sub

Sub OPEN_FILE_WCtacte4()
    Set DbsTraspa = OpenDatabase("0004tras.mdb", False, False, FILE_TYPE)
    Set rstWCtaCte4 = DbsTraspa.OpenRecordset("WCtacte")
End Sub

Sub OPEN_FILE_WEstadistica()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWEstadistica = DbsTraspa.OpenRecordset("WEstadistica")
End Sub

Sub OPEN_FILE_WDescComp()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWDescComp = DbsTraspa.OpenRecordset("WDescComp")
End Sub

Sub OPEN_FILE_WInforme()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWInforme = DbsTraspa.OpenRecordset("WInforme")
End Sub

Sub OPEN_FILE_WLAUDO()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWLaudo = DbsTraspa.OpenRecordset("WLaudo")
End Sub

Sub OPEN_FILE_WMovvar()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWMovvar = DbsTraspa.OpenRecordset("WMovvar")
End Sub

Sub OPEN_FILE_WMovguia()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWMovguia = DbsTraspa.OpenRecordset("WMovguia")
End Sub

Sub OPEN_FILE_WHoja()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWHoja = DbsTraspa.OpenRecordset("WHoja")
End Sub

Sub OPEN_FILE_WENSAYOS()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWEnsayos = DbsTraspa.OpenRecordset("WENSAYOS")
End Sub

Sub OPEN_FILE_WESPECIFICACIONES()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWEspecificaciones = DbsTraspa.OpenRecordset("WESPECIFICACIONES")
End Sub

Sub OPEN_FILE_WESPECIF()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWEspecif = DbsTraspa.OpenRecordset("WESPECIF")
End Sub

Sub OPEN_FILE_WSolic()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWSolic = DbsTraspa.OpenRecordset("WSolic")
End Sub

Sub OPEN_FILE_WPRUEBA()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWPrueba = DbsTraspa.OpenRecordset("WPRUEART")
End Sub

Sub OPEN_FILE_WPrueTer()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWPrueter = DbsTraspa.OpenRecordset("WPrueTer")
End Sub

Sub OPEN_FILE_WMovlab()
    Set DbsTraspa = OpenDatabase(WEmpresa + "tras.mdb", False, False, FILE_TYPE)
    Set rstWMovlab = DbsTraspa.OpenRecordset("WMovlab")
End Sub

Sub OPEN_FILE_WCotiza()
    Set DbsOtra = OpenDatabase(WEmpresa + "otra.mdb", False, False, FILE_TYPE)
    Set rstWCotiza = DbsOtra.OpenRecordset("WCotiza")
End Sub

Sub OPEN_FILE_WOrden()
    Set DbsOtra = OpenDatabase(WEmpresa + "otra.mdb", False, False, FILE_TYPE)
    Set rstWOrden = DbsOtra.OpenRecordset("WOrden")
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
            a% = MsgBox(m$, 0, "Archivo de " + Archivo$)
        Case Else
            m$ = Mensaje$
            a% = MsgBox(m$, 0, "Archivo de Vendedor")
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


