Attribute VB_Name = "SISTEMA"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Rem Global Const FILENAME = "LOCALIDAD.MDB"
Global Const FILE_TYPE = ""
Global PATH_PROG As String
Global coderr As Integer
Global TipoImpre As String
Global XIndice As Integer
Global text As String
Global Auxi As String
Global Auxi1 As String
Global Lote As String
Global Auxi2 As String
Global Validate As String
Global Cicla As Integer
Global WAuxi As Integer
Global WFecha As String
Global WVendedor As String
Global XCol As Integer
Global XRow As Integer
Global Existe As String
Global WUnidad As String
Global WTipo As String
Global WPunto As String
Global WLetra As String
Global WNumero As String
Global WCliente As String
Global Fecha As String
Global WVencimiento As String
Global WTipoIva As String
Global WProvincia As String
Global WNeto As Double
Global WIva1 As Double
Global WIva2 As Double
Global WTotal As Double
Global WConsecionaria As String
Global WDesde As String
Global Whasta As String
Global WAno As String
Global WMes As String
Global WDia As String


'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------

Global DbsVentas As Database
Global rstVendedor As Recordset
Global rstTipo As Recordset
Global rstModelo As Recordset
Global rstUbicacion As Recordset
Global rstUnidad As Recordset
Global rstProvincia As Recordset
Global rstCiudad As Recordset
Global rstCliente As Recordset
Global rstResultado As Recordset
Global rstVisitas As Recordset
Global rstPedido As Recordset
Global rstCtaCte As Recordset
Global rstIvaven As Recordset
Global rstConsecionaria As Recordset
Global rstDescComp As Recordset
Global rstOrigen As Recordset
Global rstListado1 As Recordset
Global rstCliRandon As Recordset

'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------
Global Const TABLA_Vendedor = "Vendedor"
Global Const TABLA_Tipo = "ClientesICAICONES"
Global Const TABLA_Modelo = "Modelo"
Global Const TABLA_Ubicacion = "Ubicacion"
Global Const TABLA_Unidad = "Unidad"
Global Const TABLA_Provincia = "Provincia"
Global Const TABLA_Ciudad = "Ciudad"
Global Const TABLA_Clientes = "Clientes"
Global Const TABLA_Resultado = "Resultado"
Global Const TABLA_Visitas = "Visitas"
Global Const TABLA_Pedido = "Pedido"
Global Const TABLA_CtaCte = "CtaCte"
Global Const TABLA_Ivaven = "IvaVen"
Global Const TABLA_Consecionaria = "Consecionaria"
Global Const TABLA_DescComp = "DescComp"
Global Const TABLA_Origen = "Origen"
Global Const TABLA_Listado1 = "Listado1"
Global Const TABLA_CliRandon = "CliRandon"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE Vendedor
'--------------------------------------------------------
 
 Global Const Codigo = "CODIGO"
 Global Const Descripcion = "DESCRIPCION"
 

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE Tipo
'--------------------------------------------------------
 
Rem  Global Const ESPEPRODUCTO = "ESPEPRODUCTO"
Rem Global Const Ensayo1 = "ENSAYO1"
Rem Global Const Valor1 = "VALOR1"
Rem Global Const Ensayo2 = "ENSAYO2"
Rem Global Const valor2 = "VALOR2"
Rem Global Const Ensayo3 = "ENSAYO3"
Rem Global Const Valor3 = "VALOR3"
Rem Global Const Ensayo4 = "ENSAYO4"
Rem Global Const valor4 = "VALOR4"
Rem Global Const Ensayo5 = "ENSAYO5"
Rem Global Const valor5 = "VALOR5"
Rem Global Const Ensayo6 = "ENSAYO6"
Rem Global Const valor6 = "VALOR6"
Rem Global Const Ensayo7 = "ENSAYO7"
Rem Global Const valor7 = "VALOR7"
Rem Global Const Ensayo8 = "ENSAYO8"
Rem Global Const valor8 = "VALOR8"
Rem Global Const Ensayo9 = "ENSAYO9"
Rem Global Const valor9 = "VALOR9"
Rem Global Const Ensayo10 = "ENSAYO10"
Rem Global Const valor10 = "VALOR10"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE Modelo
'--------------------------------------------------------

 Rem Global Const PRODUCTO = "PRODUCTO"
 Rem Global Const DESCRIPCION = "DESCRIPCION"
 
 
 
Sub OPEN_FILE_Vendedor()
    Set DbsVentas = OpenDatabase("ventas.mdb", False, False, FILE_TYPE)
    Set rstVendedor = DbsVentas.OpenRecordset("Vendedor")
End Sub
 

Sub OPEN_FILE_Tipo()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstTipo = DbsVentas.OpenRecordset("Tipo")
End Sub

Sub OPEN_FILE_Modelo()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstModelo = DbsVentas.OpenRecordset("Modelo")
End Sub

Sub OPEN_FILE_Ubicacion()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstUbicacion = DbsVentas.OpenRecordset("Ubicacion")
End Sub

Sub OPEN_FILE_Unidad()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstUnidad = DbsVentas.OpenRecordset("Unidad")
End Sub

Sub OPEN_FILE_Provincia()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstProvincia = DbsVentas.OpenRecordset("Provincia")
End Sub
Sub OPEN_FILE_Ciudad()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstCiudad = DbsVentas.OpenRecordset("Ciudad")
End Sub

Sub OPEN_FILE_Clientes()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstCliente = DbsVentas.OpenRecordset("Cliente")
End Sub

Sub OPEN_FILE_Resultado()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstResultado = DbsVentas.OpenRecordset("Resultado")
End Sub

Sub OPEN_FILE_Visitas()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstVisitas = DbsVentas.OpenRecordset("Visitas")
End Sub

Sub OPEN_FILE_Pedido()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstPedido = DbsVentas.OpenRecordset("Pedido")
End Sub

Sub OPEN_FILE_CtaCte()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstCtaCte = DbsVentas.OpenRecordset("CtaCte")
End Sub

Sub OPEN_FILE_IvaVen()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstIvaven = DbsVentas.OpenRecordset("IvaVen")
End Sub

Sub OPEN_FILE_Consecionaria()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstConsecionaria = DbsVentas.OpenRecordset("Consecionaria")
End Sub

Sub OPEN_FILE_DescComp()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstDescComp = DbsVentas.OpenRecordset("DescComp")
End Sub

Sub OPEN_FILE_Origen()
    Set DbsVentas = OpenDatabase("Ventas.mdb", False, False, FILE_TYPE)
    Set rstOrigen = DbsVentas.OpenRecordset("Origen")
End Sub

Sub OPEN_FILE_Listado1()
    Set DbsVentas = OpenDatabase("ventas.mdb", False, False, FILE_TYPE)
    Set rstListado1 = DbsVentas.OpenRecordset("Listado1")
End Sub

Sub OPEN_FILE_CliRandon()
    Set DbsVentas = OpenDatabase("c:\Archivos de Programa\Sistema Mercadologico Randon\ventas.mdb", False, False, FILE_TYPE)
    Set rstCliRandon = DbsVentas.OpenRecordset("Cliente")
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
If Mid$(T.text, T.SelStart + T.SelLength + 1, 1) = "-" Then
    KeyAscii = 0   ' Discard characters before -
End If
End Sub

Sub Errores(coderr As Integer, Archivo As String, Mensaje As String)

    e = coderr
    Select Case e
        Case 3021
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de " + Archivo$)
        Case Else
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de Vendedor")
    End Select
    
End Sub

Sub Ceros(Campo As String, Largo As Integer)

    L% = 1
    cadena$ = ""
    While L% <= Len(Campo) And L% > 0
        If Mid$(Campo, L%, 1) <> Chr$(32) Then cadena$ = cadena$ + Mid$(Campo, L%, 1)
        L% = L% + 1
    Wend
    Campo = Right$(String$(40, "0") + cadena$, Largo)
    
End Sub
