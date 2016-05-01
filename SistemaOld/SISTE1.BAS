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
Global Text As String
Global Auxi As String
Global Auxi1 As String
Global Lote As String
Global Auxi2 As String
Global Validate As String

'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------

Global DbsLaboratorio As Database
Global rstEnsayos As Recordset
Global rstEspecificaciones As Recordset
Global rstProductos As Recordset
Global rstLineas As Recordset
Global rstEnvases As Recordset
Global rstComposicion As Recordset
Global rstTerminado As Recordset
Global rstEspecif As Recordset
Global rstPrueba As Recordset
Global rstLaudo As Recordset
Global rstPrueter As Recordset

'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------
Global Const TABLA_ENSAYOS = "ENSAYOS"
Global Const TABLA_ESPECIFICACIONES = "ESPECIFICAICONES"
Global Const TABLA_PRODUCTOS = "PRODUCTOS"
Global Const TABLA_LINEAS = "LINEAS"
Global Const TABLA_ENVASES = "ENVASES"
Global Const TABLA_COMPOSICION = "COMPOSICION"
Global Const TABLA_TERMINADO = "TERMINADO"
Global Const TABLA_ESPECIF = "ESPECIF"
Global Const TABLA_PRUEBA = "PRUEBA"
Global Const TABLA_LOTE = "LOTE"
Global Const TABLA_PrueTer = "PrueTer"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE ENSAYOS
'--------------------------------------------------------
 
 Global Const Codigo = "CODIGO"
 Global Const Descripcion = "DESCRIPCION"
 

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE ESPECIFICACIONES
'--------------------------------------------------------
 
 Global Const ESPEPRODUCTO = "ESPEPRODUCTO"
 Global Const Ensayo1 = "ENSAYO1"
 Global Const Valor1 = "VALOR1"
 Global Const Ensayo2 = "ENSAYO2"
 Global Const valor2 = "VALOR2"
 Global Const Ensayo3 = "ENSAYO3"
 Global Const Valor3 = "VALOR3"
 Global Const Ensayo4 = "ENSAYO4"
 Global Const valor4 = "VALOR4"
 Global Const Ensayo5 = "ENSAYO5"
 Global Const valor5 = "VALOR5"
 Global Const Ensayo6 = "ENSAYO6"
 Global Const valor6 = "VALOR6"
 Global Const Ensayo7 = "ENSAYO7"
 Global Const valor7 = "VALOR7"
 Global Const Ensayo8 = "ENSAYO8"
 Global Const valor8 = "VALOR8"
 Global Const Ensayo9 = "ENSAYO9"
 Global Const valor9 = "VALOR9"
 Global Const Ensayo10 = "ENSAYO10"
 Global Const valor10 = "VALOR10"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE PRODUCTOS
'--------------------------------------------------------

 Rem Global Const PRODUCTO = "PRODUCTO"
 Rem Global Const DESCRIPCION = "DESCRIPCION"
 
 
 
Sub OPEN_FILE_ENSAYOS()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstEnsayos = DbsLaboratorio.OpenRecordset("ENSAYOS")
End Sub
 

Sub OPEN_FILE_ESPECIFICACIONES()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstEspecificaciones = DbsLaboratorio.OpenRecordset("ESPECIFICACIONES")
End Sub

Sub OPEN_FILE_PRODUCTOS()
    Set DbsLaboratorio = OpenDatabase("0001vent.mdb", False, False, FILE_TYPE)
    Set rstProductos = DbsLaboratorio.OpenRecordset("Articulo")
End Sub

Sub OPEN_FILE_LINEAS()
    Set DbsLaboratorio = OpenDatabase("0001vent.mdb", False, False, FILE_TYPE)
    Set rstLineas = DbsLaboratorio.OpenRecordset("LINEAS")
End Sub

Sub OPEN_FILE_ENVASES()
    Set DbsLaboratorio = OpenDatabase("0001vent.mdb", False, False, FILE_TYPE)
    Set rstEnvases = DbsLaboratorio.OpenRecordset("ENVASES")
End Sub

Sub OPEN_FILE_TERMINADO()
    Set DbsLaboratorio = OpenDatabase("0001vent.mdb", False, False, FILE_TYPE)
    Set rstTerminado = DbsLaboratorio.OpenRecordset("TERMINADO")
End Sub

Sub OPEN_FILE_ESPECIF()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstEspecif = DbsLaboratorio.OpenRecordset("ESPECIF")
End Sub

Sub OPEN_FILE_PRUEBA()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstPrueba = DbsLaboratorio.OpenRecordset("PRUEART")
End Sub

Sub OPEN_FILE_LAUDO()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstLaudo = DbsLaboratorio.OpenRecordset("LAUDO")
End Sub

Sub OPEN_FILE_PrueTer()
    Set DbsLaboratorio = OpenDatabase("0001labo.mdb", False, False, FILE_TYPE)
    Set rstPrueter = DbsLaboratorio.OpenRecordset("PrueTer")
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
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de " + Archivo$)
        Case Else
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de Ensayos")
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
