Attribute VB_Name = "SISTEMA"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Rem Global Const FILENAME = "LOCALIDAD.MDB"

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
Global WEmpresa As String
Global PCliente As String
Global PTipo As String
Global PTerminado As String
Global PLote As String
Global WRecibo As String
Global WXPed As String
Global Pasalote As String
Global DbConnect$
Global DSN$
Global UID$
Global PWD$
Global DSQ$

'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------

Global DbsEmpresa As Database
Global DbsAdminis As Database
Global DbsVentas As Database
Global DbsCotiza As Database
Global DbsAuxi As Database
Global DbsLaboratorio As Database
Global DbsCotizaciones As Database
Global DbsInve As Database

'definicion de tablas de base de datos de empresa

Global rstEmpresa As Recordset

'definicion de tablas de base de datos  de administracion

 

Sub OPEN_FILE_Empresa()
    Set DbsEmpresa = OpenDatabase("Empresa.mdb", False, False, FILE_TYPE)
    Set rstEmpresa = DbsEmpresa.OpenRecordset("Empresa")
End Sub
 

Sub OPEN_FILE_Auxiliar()
    Set DbsAuxi = OpenDatabase(WEmpresa + "Auxi.mdb", False, False, FILE_TYPE)
    Set rstAuxiliar = DbsAuxi.OpenRecordset("Auxiliar")
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
            A% = MsgBox(m$, 0, "Archivo de " + Archivo$)
        Case Else
            m$ = Mensaje$
            A% = MsgBox(m$, 0, "Archivo de Vendedor")
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


