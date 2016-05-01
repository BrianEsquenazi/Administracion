Attribute VB_Name = "Module3"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Type OrientStructure
  Orientation As Long
  Pad As String * 16
End Type

Global PATH_PROG As String
Global QUIT_IN_PROGRESS As Integer
Global Const FILENAME = "SYSTEM.MDB"
Global Const FILE_TYPE = ""
Global Const FILENAME_REPORTE_SUMAS_SALDOS = "SALDOS.MDB"

'Cantidad máxima de entradas que puede tener un asiento
Global Const MAX_ENTRADAS_POR_ASIENTO = 100

'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------
 Global DB_EMPRESAS   As Database
 Global DB_LEYENDAS   As Database
 Global DB_CONTABLES  As Database
 Global DB_ASIENTOS   As Database
 Global DB_TIPOS      As Database
 Global DB_INDICES    As Database
 Global DB_COEF       As Database
 Global DB_CAMBIO     As Database
 Global DB_AJUSTE     As Database

 Global DN_EMPRESAS   As DYNASET
 Global DN_LEYENDAS   As DYNASET
 Global DN_CONTABLES  As DYNASET
 Global DN_ASIENTOS   As DYNASET
 Global DN_TIPOS      As DYNASET
 Global DN_INDICES    As DYNASET
 Global DN_COEF       As DYNASET
 Global DN_CAMBIO     As DYNASET
 Global DN_AJUSTE     As DYNASET

'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------
  Global Const TABLA_EMPRESAS = "EMPRESAS"
  Global Const TABLA_TIPO_CAMBIO = "TIPOS DE CAMBIO"
  Global Const TABLA_NOMBRE_INDICES = "NOMBRE DE LOS INDICES"
  Global Const TABLA_LEYENDAS = "LEYENDAS"
  Global Const TABLA_CUENTAS_CONTABLES = "CUENTAS CONTABLES"
  Global Const TABLA_COEFICIENTES = "COEFICIENTES"
  Global Const TABLA_ASIENTOS_TIPO = "ASIENTOS TIPO"
  Global Const TABLA_ASIENTOS_CONTABLES = "ASIENTOS CONTABLES"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE EMPRESAS
'--------------------------------------------------------
Global Const CODIGO = "CODIGO"
Global Const Nombre = "NOMBRE"
Global Const CUIT = "CUIT"
Global Const INICIO = "INICIO"
Global Const FIN = "FIN"
Global Const DIRECCION = "DIRECCION"
Global Const LOCALIDAD = "LOCALIDAD"
Global Const PCIA = "PROVINCIA"
Global Const ACTIVIDAD = "ACTIVIDAD"
Global Const TIENE_SUCURSALES = "TIENE_SUCURSALES"
Global Const MONEDA = "MONEDA"

'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE TIPOS DE CAMBIO
'--------------------------------------------------------
Global Const FECHA_CAMBIO = "FECHA"
Global Const VALOR_CAMBIO = "VALOR"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE NOMBRES DE LOS INDICES
'--------------------------------------------------------------
Global Const INDEX_CODE = "CODIGO_INDICE"
Global Const INDEX_NAME = "NOMBRE_INDICE"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES A LA TABLA DE LEYENDAS
'--------------------------------------------------------------
Global Const CODIGO_LEYENDA = "CODIGO_LEYENDA"
Global Const LEYENDA = "NOMBRE_LEYENDA"
Global Const LEYENDA_RPT_EMPRESA = "LEYENDA_NOMBRE_EMPRESA"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES A LA TABLA DE COEFICIENTES
'--------------------------------------------------------------
Global Const COEF_INDICE = "INDICE"
Global Const COEF_MES = "MES"
Global Const COEF_AÑO = "ANIO"
Global Const COEF_IMPORTE = "IMPORTE"
Global Const RPT_COEF_INDICE = "RPT_INDICE"
Global Const RPT_COEF_EMPRESA = "RPT_EMPRESA"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES A LA TABLA DE ASIENTOS TIPO
'--------------------------------------------------------------
Global Const AS_TIPO = "AS_TIPO"
Global Const AS_RENGLON = "AS_RENGLON"
Global Const AS_CUENTA = "AS_CUENTA"
Global Const AS_DESCRIPCION = "AS_DESCRIPCION"
Global Const AS_LEYENDA = "AS_LEYENDA"
Global Const AS_SUCURSAL = "AS_SUCURSAL"
Global Const AS_RPT_NOMBRECUENTA = "AS_RPT_NOMBRE_CUENTA"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES A LA TABLA DE ASIENTOS CONTABLES
'--------------------------------------------------------------
Global Const CC_ASIENTO = "NUMERO_ASIENTO"
Global Const CC_RENGLON = "RENGLON_ASIENTO"
Global Const CC_FECHA = "FECHA_ASIENTO"
Global Const CC_TIPO = "TIPO_ASIENTO"
Global Const CC_CUENTA = "CUENTA_ASIENTO"
Global Const CC_DESCRIPCION = "DESCRIPCION_ASIENTO"
Global Const CC_OBSERVACIONES = "OBSERVACIONES_ASIENTO"
Global Const CC_DEBITO = "DEBITO_PESOS"
Global Const CC_CREDITO = "CREDITO_PESOS"
Global Const CC_DDEBITO = "DEBITO_DOLARES"
Global Const CC_DCREDITO = "CREDITO_DOLARES"
Global Const CC_INICIAL = "ASIENTO_INICIAL"
Global Const CC_SUCURSAL = "CC_SUCURSAL"
Global Const CC_REPORTE_NOMBRECUENTA = "RPT_NOM_CUENTA"  'SE USA SÓLO PARA
Global Const CC_REPORTE_SALDO_INI = "RPT_SALDO_INI"  'EL REPORTE
Global Const CC_RPT_NOMBRE_EMPRESA = "RPT_ASIENTO_EMPRESA"

'--------------------------------------------------------------
' CAMPOS CORRESPONDIENTES A LA TABLA DE CUENTAS CONTABLES
'--------------------------------------------------------------
 Global Const CUENTA = "CUENTA"
 Global Const DESCRIPCION = "DESCRIPCION"
 Global Const Nivel = "NIVEL"
 Global Const IMPUTABLE = "IMPUTABLE"
 Global Const INICIAL = "INICIAL"
 Global Const DINICIAL = "DINICIAL"
 Global Const CINICIAL = "CINICIAL"
 Global Const CONT_SUCURSAL = "CONT_SUCURSAL"
 Global Const CONT_RPT_EMPRESA = "CONT_RPT_EMPRESA"

 ' DEBITOS Y CREDITOS EN PESOS
 Global Const DEBITO1 = "DEBITO1"
 Global Const DEBITO2 = "DEBITO2"
 Global Const DEBITO3 = "DEBITO3"
 Global Const DEBITO4 = "DEBITO4"
 Global Const DEBITO5 = "DEBITO5"
 Global Const DEBITO6 = "DEBITO6"
 Global Const DEBITO7 = "DEBITO7"
 Global Const DEBITO8 = "DEBITO8"
 Global Const DEBITO9 = "DEBITO9"
 Global Const DEBITO10 = "DEBITO10"
 Global Const DEBITO11 = "DEBITO11"
 Global Const DEBITO12 = "DEBITO12"
 
 Global Const CREDITO1 = "CREDITO1"
 Global Const CREDITO2 = "CREDITO2"
 Global Const CREDITO3 = "CREDITO3"
 Global Const CREDITO4 = "CREDITO4"
 Global Const CREDITO5 = "CREDITO5"
 Global Const CREDITO6 = "CREDITO6"
 Global Const CREDITO7 = "CREDITO7"
 Global Const CREDITO8 = "CREDITO8"
 Global Const CREDITO9 = "CREDITO9"
 Global Const CREDITO10 = "CREDITO10"
 Global Const CREDITO11 = "CREDITO11"
 Global Const CREDITO12 = "CREDITO12"


 ' DEBITOS Y CREDITOS EN DOLARES
 Global Const DDEBITO1 = "DDEBITO1"
 Global Const DDEBITO2 = "DDEBITO2"
 Global Const DDEBITO3 = "DDEBITO3"
 Global Const DDEBITO4 = "DDEBITO4"
 Global Const DDEBITO5 = "DDEBITO5"
 Global Const DDEBITO6 = "DDEBITO6"
 Global Const DDEBITO7 = "DDEBITO7"
 Global Const DDEBITO8 = "DDEBITO8"
 Global Const DDEBITO9 = "DDEBITO9"
 Global Const DDEBITO10 = "DDEBITO10"
 Global Const DDEBITO11 = "DDEBITO11"
 Global Const DDEBITO12 = "DDEBITO12"
 
 Global Const DCREDITO1 = "DCREDITO1"
 Global Const DCREDITO2 = "DCREDITO2"
 Global Const DCREDITO3 = "DCREDITO3"
 Global Const DCREDITO4 = "DCREDITO4"
 Global Const DCREDITO5 = "DCREDITO5"
 Global Const DCREDITO6 = "DCREDITO6"
 Global Const DCREDITO7 = "DCREDITO7"
 Global Const DCREDITO8 = "DCREDITO8"
 Global Const DCREDITO9 = "DCREDITO9"
 Global Const DCREDITO10 = "DCREDITO10"
 Global Const DCREDITO11 = "DCREDITO11"
 Global Const DCREDITO12 = "DCREDITO12"
              
'--------------------------------------------------------
' Constantes para Imprimir Reportes
'--------------------------------------------------------
Global DESTINO_REPORTE As Integer  'Destino de los Reportes
Global Const REPORTE_IMPRESORA = 1 'IMPRESORA
Global Const REPORTE_PANTALLA = 0  'PANTALLA
Global Const CANCEL_REPORT = 255   'CANCELAR REPORTE


 Global Const PESOS = 0
 Global Const DOLAR = 1
 Global Const COMBINADOS = 2

 Global LOCK_CAPS As Integer
 Global LOCK_NUM As Integer
 Global LOCK_SCROLL As Integer

 Type SALDOS_TYPE
   CUENTA As String * 20
   Nivel As Integer
   Nombre As String * 30
 End Type

Declare Function FindWindow% Lib "USER" (ByVal LPCLASSNAME As Any, ByVal LPCAPTION As Any)
Declar Function Escape% Lib "GDI" (ByVal hdc%, ByVal nesc%, ByVal nlen%, lpdata As OrientStructure, lpout As Any)

Sub ACTUALIZAR_CUENTAS_CONTABLES()
                     
   Dim DEBI As String, CREDI As String, INI  As String, CUEN As String
   Dim SUCU As String, CON_SUCU As String, TIPO_ASI  As Integer, VALOR As Double

   '--------------------------------------------------------------------
   'Ver si la Empresa tiene sucursales y su tipo de Moneda
   '--------------------------------------------------------------------
   OPEN_EMPRESAS
   CAD = CODIGO + " = '" + LTrim$(RTrim$(MAIN.Label2.Caption)) + "'"
   DN_EMPRESAS.FindFirst CAD
   If Not DN_EMPRESAS.NoMatch Then
     CON_SUCU = DN_EMPRESAS(TIENE_SUCURSALES) & ""
     T = 0 & DN_EMPRESAS(MONEDA)
     If T <> 0 And T <> 1 And T <> 2 Then T = 0
     If CON_SUCU = "" Then CON_SUCU = "N"
   Else
     DN_EMPRESAS.Close
     DB_EMPRESAS.Close
     Beep
     MsgBox "El Código de la Empresa con la que está trabajando actualmente no fue encontrado.", 48, "Error"
     Exit Sub
   End If
   DN_EMPRESAS.Close
   DB_EMPRESAS.Close
   T = LTrim$(RTrim$(Str$(Val(T))))
   TIPO_ASI = T
   '--------------------------------------------------------------------

    '--------------------------------------------------------------------
    'Si la empresa usa el Sistema BiMonetario se debe
    'averiguar el Tipo de Cambio
    '--------------------------------------------------------------------
    If TIPO_ASI = COMBINADOS Then
       OPEN_TIPO_CAMBIO
       BUSCA = FECHA_CAMBIO + " = '" + LTrim$(RTrim$(ASIENTOS.Text2.Text)) + "'"
       DN_CAMBIO.FindFirst BUSCA
       If DN_CAMBIO.NoMatch Then
         DN_CAMBIO.Close
         DB_CAMBIO.Close
         Beep
         MsgBox "No se encontró el valor del dólar a la fecha " + LTrim$(RTrim$(ASIENTOS.Text2.Text)), 48, "Error"
         Exit Sub
       Else
         VALOR = 0
         V = LTrim$(RTrim$(DN_CAMBIO(VALOR_CAMBIO) & ""))
         If V = "" Then
           VALOR = 0
         Else
           VALOR = V
         End If
       End If
       DN_CAMBIO.Close
       DB_CAMBIO.Close
     End If
     '--------------------------------------------------------------------

   OPEN_FILE_CUENTAS_CONTABLES

   'Poner todos los saldos en cero.
   RESETEAR_CUENTAS
   

   'Crear una consulta que contenga todos los asientos contables
   'de la empresa
   Set DN_ASIENTOS = DB_ASIENTOS.CreateDynaset(TABLA_ASIENTOS_CONTABLES)


   If DN_ASIENTOS.RecordCount > 0 Then
     BeginTrans
     
     DN_ASIENTOS.MoveFirst
     Do While Not DN_ASIENTOS.EOF
       CUEN = LTrim$(RTrim$(DN_ASIENTOS(CC_CUENTA) & ""))
       SUCU = LTrim$(RTrim$(DN_ASIENTOS(CC_SUCURSAL) & ""))
       INI = LTrim$(RTrim$(DN_ASIENTOS(CC_INICIAL) & ""))

       '-------------------------------------------------------------
       'VER EL TIPO DE ASIENTO DE LA EMPRESA
       '-------------------------------------------------------------
       Select Case TIPO_ASI
         Case Is = PESOS
           DEBI = DN_ASIENTOS(CC_DEBITO) & ""
           CREDI = DN_ASIENTOS(CC_CREDITO) & ""
         
         Case Is = DOLAR
           DEBI = DN_ASIENTOS(CC_DDEBITO) & ""
           CREDI = DN_ASIENTOS(CC_DCREDITO) & ""
         
         Case Is = COMBINADOS
           'Tomar el Valor en Pesos
           DEBI = DN_ASIENTOS(CC_DEBITO) & ""
           CREDI = DN_ASIENTOS(CC_CREDITO) & ""

         Case Else
           DEBI = "0"
           CREDI = "0"
       End Select
       '-------------------------------------------------------------

       If CUEN <> "" And Len(SUCU) = 2 Then
          BUSCA = CUENTA + " = '" + CUEN + "' AND " + CONT_SUCURSAL + " = '" + SUCU + "'"
          DN_CONTABLES.FindFirst BUSCA
          
          If DN_CONTABLES.NoMatch Then
            Beep
            MsgBox "La Cuenta Nº [" + SUCU + CUEN + "] no fue encontrada.", 48, "Error"
          Else
            GRABAR_CUENTA CUEN, SUCU, INI, DEBI, CREDI, TIPO_ASI, VALOR
          End If
       Else
         Beep
         MsgBox "Error Interno", 48, "Actualización de Cuentas Contables"
       End If
       DN_ASIENTOS.MoveNext
     Loop
     
     CommitTrans
   Else
     Beep
     MsgBox "No se encontraron Asientos Contables", 64, "Información"
   End If

   DN_CONTABLES.Close
   DB_CONTABLES.Close
End Sub

Sub BORRAR_MONTO(NUM_CUENTA As String, ASIENTO_INICIAL As String, NUM_MES As Integer, MONTO_DEBE As String, MONTO_HABER As String)

   Const PESOS = 0
   Const DOLAR = 1
   Const COMBINADOS = 2

   Dim T As String, CON_SUCU As String, CAD As String, TIPO_ASI As Integer, M As Integer
   Dim MONTO_CUENTA_DEBE As String, MONTO_CUENTA_HABER As String, VALOR As String

   '--------------------------------------------------------------------
   'Buscar el Tipo de Moneda de la Empresa
   'y ver si tiene sucursales.
   '--------------------------------------------------------------------
   OPEN_EMPRESAS
   CAD = CODIGO + " = '" + LTrim$(RTrim$(MAIN.Label2.Caption)) + "'"
   DN_EMPRESAS.FindFirst CAD
   If Not DN_EMPRESAS.NoMatch Then
     CON_SUCU = DN_EMPRESAS(TIENE_SUCURSALES) & ""
     T = DN_EMPRESAS(MONEDA) & ""
     If T <> "0" And T <> "1" And T <> "2" Then T = "0"
     If CON_SUCU = "" Then CON_SUCU = "N"
   Else
     DN_EMPRESAS.Close
     DB_EMPRESAS.Close
     Beep
     MsgBox "El Código de la Empresa con la que está trabajando actualmente no fue encontrado.", 48, "Error"
     Exit Sub
   End If
   DN_EMPRESAS.Close
   DB_EMPRESAS.Close
   T = Str$(Val(T))
   '--------------------------------------------------------------------
   TIPO_ASI = Val(T)

   If CON_SUCU = "S" Then
     NUM_CUENTA = Mid$(NUM_CUENTA, 3, 20)
     NUM_SUCU = Mid$(NUM_CUENTA, 1, 2)
   Else
     NUM_SUCU = "00"
   End If

   NUM_CUENTA = LTrim$(RTrim$(NUM_CUENTA))
   NUM_SUCU = LTrim$(RTrim$(NUM_SUCU))

   If Len(NUM_SUCU) = 2 And NUM_CUENTA <> "" Then
       If TIPO_ASI = COMBINADOS Then
         '--------------------------------------------------------------------
         'Si los asientos son combinados se debe buscar el tipo
         'de cambio para la fecha del asiento
         '--------------------------------------------------------------------
          OPEN_TIPO_CAMBIO
          BUSCA = FECHA_CAMBIO + " = '" + LTrim$(RTrim$(ASIENTOS.Text2.Text)) + "'"
          DN_CAMBIO.FindFirst BUSCA
          If DN_CAMBIO.NoMatch Then
            DN_CAMBIO.Close
            DB_CAMBIO.Close
            Beep
            MsgBox "No se pudo encontrar el valor del Dólar a la fecha " + LTrim$(RTrim$(ASIENTOS.Text2.Text)), 48, "Tipo de Cambio No Encontrado"
            Exit Sub
          Else
            VALOR = LTrim$(RTrim$(DN_CAMBIO(VALOR_CAMBIO) & ""))
          End If
          DN_CAMBIO.Close
          DB_CAMBIO.Close
         '--------------------------------------------------------------------
       End If
      
      '------------------------
      'Buscar el Nº de Cuenta
      '------------------------
      BUSCA = CUENTA + " = '" + NUM_CUENTA + "' AND " + CONT_SUCURSAL + " = '" + NUM_SUCU + "'"
      DN_CONTABLES.FindFirst BUSCA
      If Not DN_CONTABLES.NoMatch Then
          
          DN_CONTABLES.Edit

          
          '-----------------------------------------------------------
          'Si el Asiento es Inicial...
          '-----------------------------------------------------------
          If INI = "S" Then
             Select Case TIPO_ASI
                Case Is = PESOS
                   M = 0
                   MONTO_INI = LTrim$(RTrim$(DN_CONTABLES(INICIAL) & ""))
                   If MONTO_INI = "" Then M = 0
                   M = M - Val(MONTO_DEBE) - Val(MONTO_HABER)
                   DN_CONTABLES(INICIAL) = LTrim$(RTrim$(Str$(M)))

                Case Is = DOLAR
                   M = 0
                   MONTO_INI = LTrim$(RTrim$(DN_CONTABLES(DINICIAL) & ""))
                   If MONTO_INI = "" Then M = 0
                   M = M - Val(MONTO_DEBE) - Val(MONTO_HABER)
                   DN_CONTABLES(DINICIAL) = LTrim$(RTrim$(Str$(M)))

                Case Is = COMBINADOS
                   M = 0
                   MONTO_INI = LTrim$(RTrim$(DN_CONTABLES(CINICIAL) & ""))
                   If MONTO_INI = "" Then M = 0
                   M = M - Val(MONTO_DEBE) - Val(MONTO_HABER)
                   DN_CONTABLES(CINICIAL) = LTrim$(RTrim$(Str$(M)))

             End Select
          End If
          '-----------------------------------------------------------
          
          Select Case NUM_MES
            Case Is = 1
               If TIPO_ASI = PESOS Then
                 'Leer los valores grabados en disco
                 MONTO_CUENTA_DEBE = DN_CONTABLES(DEBITO1) & ""
                 MONTO_CUENTA_HABER = DN_CONTABLES(CREDITO1) & ""

                 'Restar los montos
                 MONTO_CUENTA_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_DEBE) - Val(MONTO_DEBE))))
                 MONTO_CUENTA_HABER = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_HABER) - Val(MONTO_HABER))))

                 'Actualizar registro
                 DN_CONTABLES(DEBITO1) = MONTO_CUENTA_DEBE
                 DN_CONTABLES(CREDITO1) = MONTO_CUENTA_HABER
                 '----------------------------------------------------------

               ElseIf TIPO_ASI = DOLAR Then
                 'Leer los valores grabados en disco
                 MONTO_CUENTA_DEBE = DN_CONTABLES(DDEBITO1) & ""
                 MONTO_CUENTA_HABER = DN_CONTABLES(DCREDITO1) & ""

                 'Restar los montos
                 MONTO_CUENTA_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_DEBE) - Val(MONTO_DEBE))))
                 MONTO_CUENTA_HABER = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_HABER) - Val(MONTO_HABER))))

                 'Actualizar registro
                 DN_CONTABLES(DDEBITO1) = MONTO_CUENTA_DEBE
                 DN_CONTABLES(DCREDITO1) = MONTO_CUENTA_HABER
                 '----------------------------------------------------------

               
               ElseIf TIPO_ASI = COMBINADOS Then
                 '----------------------------------------------------------
                 'Para valores en Pesos
                 '----------------------------------------------------------
                 MONTO_CUENTA_DEBE = DN_CONTABLES(DEBITO1) & ""
                 MONTO_CUENTA_HABER = DN_CONTABLES(CREDITO1) & ""
                 MONTO_CUENTA_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_DEBE) - Val(MONTO_DEBE))))
                 MONTO_CUENTA_HABER = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_HABER) - Val(MONTO_HABER))))
                 DN_CONTABLES(DEBITO1) = MONTO_CUENTA_DEBE
                 DN_CONTABLES(CREDITO1) = MONTO_CUENTA_HABER
                 
                 '----------------------------------------------------------
                 'Valor Equivalente en Dólares
                 '----------------------------------------------------------
                 DEBI_DOLAR = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_DEBE) / Val(VALOR))))
                 CREDI_DOLAR = LTrim$(RTrim$(Str$(Val(MONTO_CUENTA_HABER) / Val(VALOR))))
                 DN_CONTABLES(DDEBITO1) = MONTO_CUENTA_DEBE
                 DN_CONTABLES(DCREDITO1) = MONTO_CUENTA_HABER
                 '----------------------------------------------------------
               End If

          End Select

          DN_CONTABLES.Update
      End If
   End If

End Sub

Sub BUSCAR_EMPRESA(COD_EMPRESA As String)
 
 COD_EMPRESA = LTrim$(RTrim$(COD_EMPRESA))

 If COD_EMPRESA <> "" Then
   CADENA = CODIGO + " = '" + COD_EMPRESA + "'"

   DN_EMPRESAS.FindFirst CADENA
   
   If Not DN_EMPRESAS.NoMatch Then VER_EMPRESA
 End If
End Sub

Function CERO(S As String) As String
 If Len(S) = 1 Then S = "0" + S
 CERO = S
End Function

Sub CREAR_FILE_AJUSTE()
  
  Dim DB As Database, TB As New TableDef
  Dim F1 As New Field
  Dim F2 As New Field
  Dim F3 As New Field
  Dim F4 As New Field
  Dim F5 As New Field
  Dim F6 As New Field
  Dim F7 As New Field
  Dim F8 As New Field


  Dim ARCHIVO As String, COD_EMPRESA  As String


  COD_EMPRESA = LTrim$(RTrim$(MAIN.Label2.Caption))

  WAIT1.Label1.Caption = "Creando Reporte Balance Ajustado" + Chr(13) + Chr(10) + "para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = GET_FILENAME(COD_EMPRESA)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = "BALANCE"
F1.Name = "AJUSTE_SUCURSAL"
F1.Type = DB_TEXT
F1.Size = 2
TB.Fields.Append F1

F2.Name = "N_CUENTA"
F2.Type = DB_TEXT
F2.Size = 10
TB.Fields.Append F2

F3.Name = "DESCRI"
F3.Type = DB_TEXT
F3.Size = 30
TB.Fields.Append F3

F4.Name = "HISTORICO"
F4.Type = DB_TEXT
F4.Size = 10
TB.Fields.Append F4

F5.Name = "AJUSTADO"
F5.Type = DB_TEXT
F5.Size = 10
TB.Fields.Append F5

F6.Name = "DIFERENCIA"
F6.Type = DB_TEXT
F6.Size = 10
TB.Fields.Append F6

F7.Name = "EMPR"
F7.Type = DB_TEXT
F7.Size = 30
TB.Fields.Append F7

F8.Name = "FECHA"
F8.Type = DB_TEXT
F8.Size = 7
TB.Fields.Append F8

'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1


End Sub

Sub CREAR_FILE_ASIENTOS_CONTABLES(COD_EMPRESA As String)

  Dim DB As Database, TB As New TableDef
  Dim IDX As New Index, IDX2 As New Index, IDX3 As New Index

  Dim F1 As New Field
  Dim F2 As New Field
  Dim F3 As New Field
  Dim F4 As New Field
  Dim F5 As New Field
  Dim F6 As New Field
  Dim F7 As New Field
  Dim F8 As New Field
  Dim F9 As New Field
  Dim F10 As New Field
  Dim F11 As New Field
  Dim F12 As New Field
  Dim F13 As New Field
  Dim F14 As New Field
  Dim F15 As New Field
  Dim F16 As New Field

  Dim ARCHIVO As String

  WAIT1.Label1.Caption = "Creando la estructura de Asientos Contables" + Chr(13) + Chr(10) + " para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = GET_FILENAME(COD_EMPRESA)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de ASIENTOS CONTABLES y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_ASIENTOS_CONTABLES
F1.Name = CC_ASIENTO
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = CC_RENGLON
F2.Type = DB_TEXT
F2.Size = 2
TB.Fields.Append F2

F3.Name = CC_FECHA
F3.Type = DB_TEXT
F3.Size = 10
TB.Fields.Append F3

F4.Name = CC_TIPO
F4.Type = DB_TEXT
F4.Size = 1
TB.Fields.Append F4

F5.Name = CC_OBSERVACIONES
F5.Type = DB_TEXT
F5.Size = 40
TB.Fields.Append F5

F6.Name = CC_CUENTA
F6.Type = DB_TEXT
F6.Size = 10
TB.Fields.Append F6

F7.Name = CC_DESCRIPCION
F7.Type = DB_TEXT
F7.Size = 20
TB.Fields.Append F7

F8.Name = CC_DEBITO
F8.Type = DB_TEXT
F8.Size = 10
TB.Fields.Append F8

F9.Name = CC_CREDITO
F9.Type = DB_TEXT
F9.Size = 10
TB.Fields.Append F9

F10.Name = CC_DDEBITO
F10.Type = DB_TEXT
F10.Size = 10
TB.Fields.Append F10

F11.Name = CC_DCREDITO
F11.Type = DB_TEXT
F11.Size = 10
TB.Fields.Append F11

F12.Name = CC_INICIAL
F12.Type = DB_TEXT
F12.Size = 1
TB.Fields.Append F12

F13.Name = CC_SUCURSAL
F13.Type = DB_TEXT
F13.Size = 2
TB.Fields.Append F13

F14.Name = CC_REPORTE_NOMBRECUENTA
F14.Type = DB_TEXT
F14.Size = 40
TB.Fields.Append F14

F15.Name = CC_REPORTE_SALDO_INI
F15.Type = DB_TEXT
F15.Size = 10
TB.Fields.Append F15

F16.Name = CC_RPT_NOMBRE_EMPRESA
F16.Type = DB_TEXT
F16.Size = 30
TB.Fields.Append F16

'CREAR LOS INDICES
IDX.Name = "ASIENTO_RENGLON"
IDX.Unique = False  'True
IDX.Primary = False
IDX.Fields = CC_ASIENTO + ";" + CC_RENGLON
TB.Indexes.Append IDX

IDX2.Name = "CUENTA_SUCURSAL_ASIENTO"
IDX2.Unique = False
IDX2.Primary = False
IDX2.Fields = CC_CUENTA + ";" + CC_SUCURSAL + ";" + CC_ASIENTO
TB.Indexes.Append IDX2

IDX3.Name = "SUCURSAL_CUENTA_ASIENTO"
IDX3.Unique = False
IDX3.Primary = False
IDX3.Fields = CC_SUCURSAL + ";" + CC_CUENTA + ";" + CC_ASIENTO
TB.Indexes.Append IDX3
'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1
End Sub

Sub CREAR_FILE_ASIENTOS_TIPO(COD_EMPRESA As String)

'Este procedimiento crea la estructura de la base de datos para la empresa
'que posee el código que se pasa como parámetro

  Dim DB As Database, TB As New TableDef, IDX As New Index
  Dim F1 As New Field
  Dim F2 As New Field
  Dim F3 As New Field
  Dim F4 As New Field
  Dim F5 As New Field
  Dim F6 As New Field
  Dim F7 As New Field
  
  Dim ARCHIVO As String

  WAIT1.Label1.Caption = "Creando la estructura de Asientos Tipo" + Chr(13) + Chr(10) + " para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = GET_FILENAME(COD_EMPRESA)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de ASIENTOS TIPO y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_ASIENTOS_TIPO
F1.Name = AS_TIPO
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = AS_RENGLON
F2.Type = DB_TEXT
F2.Size = 2
TB.Fields.Append F2

F3.Name = AS_CUENTA
F3.Type = DB_TEXT
F3.Size = 10
TB.Fields.Append F3

F4.Name = AS_DESCRIPCION
F4.Type = DB_TEXT
F4.Size = 50
TB.Fields.Append F4

F5.Name = AS_LEYENDA
F5.Type = DB_TEXT
F5.Size = 20
TB.Fields.Append F5

F6.Name = AS_SUCURSAL
F6.Type = DB_TEXT
F6.Size = 2
TB.Fields.Append F6

F7.Name = AS_RPT_NOMBRECUENTA
F7.Type = DB_TEXT
F7.Size = 40
TB.Fields.Append F7
'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1
End Sub

Sub CREAR_FILE_COEF(COD_EMPRESA As String)

'Este procedimiento crea la estructura de la base de datos para la empresa
'que posee el código que se pasa como parámetro

  Dim DB As Database, TB As New TableDef
  Dim F1 As New Field
  Dim F2 As New Field
  Dim F3 As New Field
  Dim F4 As New Field
  Dim F5 As New Field
  Dim F6 As New Field

  Dim ARCHIVO As String

  WAIT1.Label1.Caption = "Creando la estructura de Coeficientes" + Chr(13) + Chr(10) + "para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = GET_FILENAME(COD_EMPRESA)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de COEFICIENTES y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_COEFICIENTES
F1.Name = COEF_INDICE
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = COEF_AÑO
F2.Type = DB_TEXT
F2.Size = 2
TB.Fields.Append F2

F3.Name = COEF_MES
F3.Type = DB_TEXT
F3.Size = 2
TB.Fields.Append F3

F4.Name = COEF_IMPORTE
F4.Type = DB_DOUBLE
TB.Fields.Append F4

F5.Name = RPT_COEF_INDICE
F5.Type = DB_TEXT
F5.Size = 30
TB.Fields.Append F5

F6.Name = RPT_COEF_EMPRESA
F6.Type = DB_TEXT
F6.Size = 30
TB.Fields.Append F6

'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1

End Sub

Sub CREAR_FILE_CUENTAS_CONTABLES(CLAVE As String)

 WAIT1.Label1.Caption = "Creando la Estructura de Cuentas Contables" + Chr(13) + Chr(10) + " para la empresa " + MAIN.Label3.Caption
 WAIT1.Show
 WAIT1.Label1.Refresh

Dim DB As Database, TB As New TableDef
Dim ARCHIVO As String

Dim IDX2 As New Index
Dim IDX3 As New Index

Dim F1 As New Field
Dim F2 As New Field
Dim F3 As New Field
Dim F4 As New Field
Dim F5 As New Field
Dim F6 As New Field
Dim F7 As New Field
Dim F8 As New Field
Dim F9 As New Field
Dim F10 As New Field
Dim F11 As New Field
Dim F12 As New Field
Dim F13 As New Field
Dim F14 As New Field
Dim F15 As New Field
Dim F16 As New Field
Dim F17 As New Field
Dim F18 As New Field
Dim F19 As New Field
Dim F20 As New Field
Dim F21 As New Field
Dim F22 As New Field
Dim F23 As New Field
Dim F24 As New Field
Dim F25 As New Field
Dim F26 As New Field
Dim F27 As New Field
Dim F28 As New Field
Dim F29 As New Field
Dim F30 As New Field
Dim F31 As New Field
Dim F32 As New Field
Dim F33 As New Field
Dim F34 As New Field
Dim F35 As New Field
Dim F36 As New Field
Dim F37 As New Field
Dim F38 As New Field
Dim F39 As New Field
Dim F40 As New Field
Dim F41 As New Field
Dim F42 As New Field
Dim F43 As New Field
Dim F44 As New Field
Dim F45 As New Field
Dim F46 As New Field
Dim F47 As New Field
Dim F48 As New Field
Dim F49 As New Field
Dim F50 As New Field
Dim F51 As New Field
Dim F52 As New Field
Dim F53 As New Field
Dim F54 As New Field
Dim F55 As New Field
Dim F56 As New Field
Dim F57 As New Field
     
  F$ = MAIN.Label2.Caption
  ARCHIVO = GET_FILENAME(F$)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de Cuentas Contables y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_CUENTAS_CONTABLES
F1.Name = CUENTA
F1.Type = DB_TEXT
F1.Size = 10
TB.Fields.Append F1

F2.Name = DESCRIPCION
F2.Type = DB_TEXT
F2.Size = 40
TB.Fields.Append F2

F3.Name = Nivel
F3.Type = DB_TEXT
F3.Size = 2
TB.Fields.Append F3

F4.Name = IMPUTABLE
F4.Type = DB_TEXT
F4.Size = 1
TB.Fields.Append F4

F5.Name = INICIAL
F5.Type = DB_TEXT
F5.Size = 20
TB.Fields.Append F5

F6.Name = DEBITO1
F6.Type = DB_DOUBLE
TB.Fields.Append F6

F7.Name = CREDITO1
F7.Type = DB_DOUBLE
TB.Fields.Append F7

F8.Name = DEBITO2
F8.Type = DB_DOUBLE
TB.Fields.Append F8

F9.Name = CREDITO2
F9.Type = DB_DOUBLE
TB.Fields.Append F9

F10.Name = DEBITO3
F10.Type = DB_DOUBLE
TB.Fields.Append F10

F11.Name = CREDITO3
F11.Type = DB_DOUBLE
TB.Fields.Append F11

F12.Name = DEBITO4
F12.Type = DB_DOUBLE
TB.Fields.Append F12

F13.Name = CREDITO4
F13.Type = DB_DOUBLE
TB.Fields.Append F13

F14.Name = DEBITO5
F14.Type = DB_DOUBLE
TB.Fields.Append F14

F15.Name = CREDITO5
F15.Type = DB_DOUBLE
TB.Fields.Append F15

F16.Name = DEBITO6
F16.Type = DB_DOUBLE
TB.Fields.Append F16

F17.Name = CREDITO6
F17.Type = DB_DOUBLE
TB.Fields.Append F17

F18.Name = DEBITO7
F18.Type = DB_DOUBLE
TB.Fields.Append F18

F19.Name = CREDITO7
F19.Type = DB_DOUBLE
TB.Fields.Append F19

F20.Name = DEBITO8
F20.Type = DB_DOUBLE
TB.Fields.Append F20

F21.Name = CREDITO8
F21.Type = DB_DOUBLE
TB.Fields.Append F21

F22.Name = DEBITO9
F22.Type = DB_DOUBLE
TB.Fields.Append F22

F23.Name = CREDITO9
F23.Type = DB_DOUBLE
TB.Fields.Append F23

F24.Name = DEBITO10
F24.Type = DB_DOUBLE
TB.Fields.Append F24

F53.Name = CREDITO10
F53.Type = DB_DOUBLE
TB.Fields.Append F53

F25.Name = DEBITO11
F25.Type = DB_DOUBLE
TB.Fields.Append F25

F26.Name = CREDITO11
F26.Type = DB_DOUBLE
TB.Fields.Append F26
                   
F27.Name = DEBITO12
F27.Type = DB_DOUBLE
TB.Fields.Append F27

F28.Name = CREDITO12
F28.Type = DB_DOUBLE
TB.Fields.Append F28

F29.Name = DINICIAL
F29.Type = DB_TEXT
F29.Size = 20
TB.Fields.Append F29

F30.Name = DDEBITO1
F30.Type = DB_DOUBLE
TB.Fields.Append F30

F31.Name = DCREDITO1
F31.Type = DB_DOUBLE
TB.Fields.Append F31

F32.Name = DDEBITO2
F32.Type = DB_DOUBLE
TB.Fields.Append F32

F33.Name = DCREDITO2
F33.Type = DB_DOUBLE
TB.Fields.Append F33

F34.Name = DDEBITO3
F34.Type = DB_DOUBLE
TB.Fields.Append F34

F35.Name = DCREDITO3
F35.Type = DB_DOUBLE
TB.Fields.Append F35

F36.Name = DDEBITO4
F36.Type = DB_DOUBLE
TB.Fields.Append F36

F37.Name = DCREDITO4
F37.Type = DB_DOUBLE
TB.Fields.Append F37

F38.Name = DDEBITO5
F38.Type = DB_DOUBLE
TB.Fields.Append F38

F39.Name = DCREDITO5
F39.Type = DB_DOUBLE
TB.Fields.Append F39

F40.Name = DDEBITO6
F40.Type = DB_DOUBLE
TB.Fields.Append F40

F41.Name = DCREDITO6
F41.Type = DB_DOUBLE
TB.Fields.Append F41

F42.Name = DDEBITO7
F42.Type = DB_DOUBLE
TB.Fields.Append F42

F43.Name = DCREDITO7
F43.Type = DB_DOUBLE
TB.Fields.Append F43

F44.Name = DDEBITO8
F44.Type = DB_DOUBLE
TB.Fields.Append F44

F45.Name = DCREDITO8
F45.Type = DB_DOUBLE
TB.Fields.Append F45

F46.Name = DDEBITO9
F46.Type = DB_DOUBLE
TB.Fields.Append F46

F47.Name = DCREDITO9
F47.Type = DB_DOUBLE
TB.Fields.Append F47

F48.Name = DDEBITO10
F48.Type = DB_DOUBLE
TB.Fields.Append F48

F54.Name = DCREDITO10
F54.Type = DB_DOUBLE
TB.Fields.Append F54


F49.Name = DDEBITO11
F49.Type = DB_DOUBLE
TB.Fields.Append F49

F50.Name = DCREDITO11
F50.Type = DB_DOUBLE
TB.Fields.Append F50
                   
F51.Name = DDEBITO12
F51.Type = DB_DOUBLE
TB.Fields.Append F51

F52.Name = DCREDITO12
F52.Type = DB_DOUBLE
TB.Fields.Append F52

F55.Name = CONT_RPT_EMPRESA
F55.Type = DB_TEXT
F55.Size = 30
TB.Fields.Append F55

F56.Name = CONT_SUCURSAL
F56.Type = DB_TEXT
F56.Size = 2
TB.Fields.Append F56

F57.Name = CINICIAL
F57.Type = DB_TEXT
F57.Size = 20
TB.Fields.Append F57

'CREAR LOS INDICES
IDX2.Name = "CUENTA_SUCURSAL"
IDX2.Unique = True
IDX2.Primary = True
IDX2.Fields = CUENTA + ";" + CONT_SUCURSAL
TB.Indexes.Append IDX2

IDX3.Name = "POR_CUENTA"
IDX3.Unique = False
IDX3.Primary = False
IDX3.Fields = CUENTA
TB.Indexes.Append IDX3

'-----------------------------------------------------------------------

DB.TableDefs.Append TB
DB.Close
 
WAIT1.Hide
Unload WAIT1

End Sub

Sub CREAR_FILE_EMPRESAS()

'Este procedimiento crea la estructura de la base de datos y la tabla
'de empresas.

Dim DB As Database, TB As New TableDef
Dim F1 As New Field
Dim F2 As New Field
Dim F3 As New Field
Dim F4 As New Field
Dim F5 As New Field
Dim F6 As New Field
Dim F7 As New Field
Dim F8 As New Field
Dim F9 As New Field
Dim F10 As New Field
Dim F11 As New Field

  WAIT1.Label1.Caption = "Creando Tabla de Empresas en el Archivo Maestro"
  WAIT1.Show
  WAIT1.Label1.Refresh
  
  If Not FILE_EXISTS(PATH_PROG + FILENAME) Then
    Set DB = CreateDatabase(PATH_PROG + FILENAME, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(PATH_PROG + FILENAME, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de Empresas y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_EMPRESAS
F1.Name = CODIGO
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = Nombre
F2.Type = DB_TEXT
F2.Size = 40
TB.Fields.Append F2

F3.Name = DIRECCION
F3.Type = DB_TEXT
F3.Size = 30
TB.Fields.Append F3

F4.Name = LOCALIDAD
F4.Type = DB_TEXT
F4.Size = 30
TB.Fields.Append F4
       
F5.Name = PCIA
F5.Type = DB_TEXT
F5.Size = 30
TB.Fields.Append F5

F6.Name = CUIT
F6.Type = DB_TEXT
F6.Size = 15
TB.Fields.Append F6

F7.Name = INICIO
F7.Type = DB_DATE
TB.Fields.Append F7

F8.Name = FIN
F8.Type = DB_DATE
TB.Fields.Append F8

F9.Name = ACTIVIDAD
F9.Type = DB_TEXT
F9.Size = 80
TB.Fields.Append F9

F10.Name = TIENE_SUCURSALES
F10.Type = DB_TEXT
F10.Size = 1
TB.Fields.Append F10

F11.Name = MONEDA
F11.Type = DB_BYTE
TB.Fields.Append F11
'-----------------------------------------------------------------------

DB.TableDefs.Append TB
DB.Close
WAIT1.Hide
Unload WAIT1
End Sub

Sub CREAR_FILE_INDICES()

Dim DB As Database, TB As New TableDef
Dim F1 As New Field
Dim F2 As New Field

  WAIT1.Label1.Caption = "Creando Tabla de Indices en el Archivo Maestro"
  WAIT1.Show
  WAIT1.Label1.Refresh
 
  If Not FILE_EXISTS(PATH_PROG + FILENAME) Then
    Set DB = CreateDatabase(PATH_PROG + FILENAME, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(PATH_PROG + FILENAME, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de INDICES y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_NOMBRE_INDICES
F1.Name = INDEX_CODE
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = INDEX_NAME
F2.Type = DB_TEXT
F2.Size = 40
TB.Fields.Append F2

'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close
WAIT1.Hide
Unload WAIT1

End Sub

Sub CREAR_FILE_LEYENDAS(COD_EMPRESA As String)

'Este procedimiento crea la estructura de la base de datos para la empresa
'que posee el código que se pasa como parámetro

  Dim DB As Database, TB As New TableDef, IDX As New Index
  Dim F1 As New Field
  Dim F2 As New Field
  Dim F3 As New Field

  Dim ARCHIVO As String

  WAIT1.Label1.Caption = "Creando la estructura de Leyendas" + Chr(13) + Chr(10) + " para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = GET_FILENAME(COD_EMPRESA)

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de LEYENDAS y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_LEYENDAS
F1.Name = CODIGO_LEYENDA
F1.Type = DB_TEXT
F1.Size = 4
TB.Fields.Append F1

F2.Name = LEYENDA
F2.Type = DB_TEXT
F2.Size = 20
TB.Fields.Append F2

F3.Name = LEYENDA_RPT_EMPRESA
F3.Type = DB_TEXT
F3.Size = 30
TB.Fields.Append F3
'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1

End Sub

Sub CREAR_FILE_SUMAS_SALDOS()
  
  Dim DB As Database, TB As New TableDef
  Dim F1 As New Field

  Dim ARCHIVO As String

  WAIT1.Label1.Caption = "Creando estructura de Sumas y Saldos" + Chr(13) + Chr(10) + "para la empresa " + MAIN.Label3.Caption
  WAIT1.Show
  WAIT1.Label1.Refresh

  ARCHIVO = PATH_PROG + FILENAME_REPORTE_SUMAS_SALDOS

  If Not FILE_EXISTS(ARCHIVO) Then
    Set DB = CreateDatabase(ARCHIVO, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(ARCHIVO, False, False)
  End If


TB.Name = TABLA_SUMAS_SALDOS
F1.Name = CAMPO_TEXTO
F1.Type = DB_TEXT
F1.Size = 80
TB.Fields.Append F1

'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close

WAIT1.Hide
Unload WAIT1
End Sub

Sub CREAR_FILE_TIPO_CAMBIO()

'Este procedimiento crea la estructura de la base de datos y la tabla
'de los TIPOS DE CAMBIO.

Dim DB As Database, TB As New TableDef
Dim F1 As New Field
Dim F2 As New Field

  WAIT1.Label1.Caption = "Creando Tabla para Tipos de Cambio en el Archivo Maestro"
  WAIT1.Show
  WAIT1.Label1.Refresh
  
 
 
  If Not FILE_EXISTS(PATH_PROG + FILENAME) Then
    Set DB = CreateDatabase(PATH_PROG + FILENAME, DB_LANG_GENERAL, DB_VERSION_10)
  Else
    Set DB = OpenDatabase(PATH_PROG + FILENAME, False, False)
  End If

'-----------------------------------------------------------------------
' Creación de la tabla de TIPOS DE CAMBIO y sus respectivos campos
'-----------------------------------------------------------------------
TB.Name = TABLA_TIPO_CAMBIO
F1.Name = FECHA_CAMBIO
F1.Type = DB_TEXT
F1.Size = 10
TB.Fields.Append F1

F2.Name = VALOR_CAMBIO
F2.Type = DB_DOUBLE
TB.Fields.Append F2

'-----------------------------------------------------------------------
DB.TableDefs.Append TB
DB.Close
WAIT1.Hide
Unload WAIT1
End Sub

Function CUIT_VALIDO(CUIT As String) As Integer
  Dim CICLO As Integer, SUMA As Integer, DA As Variant, RESTO As Integer
  Dim COMPARA As String, DIGITO As String, DIGI As Integer, RESULTADO As Integer
  
  ReDim UNO(1 To 10) As Integer
  ReDim DOS(1 To 10) As Integer


  For CICLO = 1 To 10
    UNO(CICLO) = Val(Mid$(CUIT, CICLO, 1))
  Next

  COMPARA = Right$(CUIT, 1)

  DOS(1) = 5
  DOS(2) = 4
  DOS(3) = 3
  DOS(4) = 2
  DOS(5) = 7
  DOS(6) = 6
  DOS(7) = 5
  DOS(8) = 4
  DOS(9) = 3
  DOS(10) = 2

  SUMA = 0

 For CICLO = 1 To 10
   SUMA = SUMA + (UNO(CICLO) * DOS(CICLO))
 Next

 DA = Int(SUMA / 11) * 11
 RESTO = SUMA - DA

 If RESTO = 0 Then
   DIGITO = "0"
 Else
   DIGI = 11 - RESTO
   DIGITO = Str$(DIGI)
   'CALL CEROS(DIGITO, 1)
 End If

 If COMPARA <> DIGITO Then
   RESULTADO = False
 Else
   RESULTADO = True
 End If

 CUIT_VALIDO = RESULTADO
End Function

Function DEBE_Y_HABER_IGUALES() As Integer

 'Esta función controla que las columnas del DEBE y el HABER
 'sean iguales.

  Dim ACUM_DEBE As Double, ACUM_HABER As Double, RESULTADO As Integer
  Dim A As Integer, DEBE As String, HABER As String

  ACUM_DEBE = 0
  ACUM_HABER = 0

  A = 0
  Do
    A = A + 1
    ASIENTOS.Grid1.Row = A
    'Nº DE CUENTA
    ASIENTOS.Grid1.Col = 0
    CUEN = LTrim$(RTrim$(ASIENTOS.Grid1.Text))

    'DEBE
    ASIENTOS.Grid1.Col = 3
    DEBE = LTrim$(RTrim$(ASIENTOS.Grid1.Text))
    
    'HABER
    ASIENTOS.Grid1.Col = 4
    HABER = LTrim$(RTrim$(ASIENTOS.Grid1.Text))

    If CUEN <> "" Then
      ACUM_DEBE = ACUM_DEBE + Val(DEBE)
      ACUM_HABER = ACUM_HABER + Val(HABER)
    End If
  Loop Until CUEN = ""
  

  'Es necesario hacer la conversión a cadena para que la función
  'devuelva el valor correcto.
  'No se puede hacer la comparación entre ACUM_DEBE y ACUM_HABER.
  '(compararlos da cualquier cosa)
   DEBE = LTrim$(RTrim$(Str$(ACUM_DEBE)))
   HABER = LTrim$(RTrim$(Str$(ACUM_HABER)))

  If DEBE = HABER Then
    RESULTADO = True
  Else
    RESULTADO = False
  End If

  ASIENTOS.Grid1.Col = 0
  ASIENTOS.Grid1.Row = 1

  ASIENTOS.RES_DEBE.Caption = DEBE
  ASIENTOS.RES_HABER.Caption = HABER

  DEBE_Y_HABER_IGUALES = RESULTADO
End Function

Function FECHA_MAYOR_IGUAL(F1 As String, F2 As String) As Integer
  
  Dim RESULTADO As Integer
  Dim D1 As String, D2 As String, M1 As String, M2 As String, A1 As String, A2 As String
  Dim FEC1 As String, FEC2 As String

  'ESTA FUNNCIÓN DEVUELVE "TRUE" SI LA FECHA "F1" >= "F2"

  FEC1 = Format$(F1, "DD/MM/YYYY")
  D1 = Mid$(FEC1, 1, 2)
  M1 = Mid$(FEC1, 4, 2)
  A1 = Mid$(FEC1, 7, 4)

  FEC2 = Format$(F2, "DD/MM/YYYY")
  D2 = Mid$(FEC2, 1, 2)
  M2 = Mid$(FEC2, 4, 2)
  A2 = Mid$(FEC2, 7, 4)

  RESULTADO = False
  
  AÑO_MAYOR = Val(A1) >= Val(A2)
  MES_MAYOR = Val(M1) >= Val(M2)
  DIA_MAYOR = Val(D1) >= Val(D2)
  

  If AÑO_MAYOR And MES_MAYOR And DIA_MAYOR Then
     RESULTADO = True
  
  ElseIf (Val(A1) = Val(A2) And Val(M1) = Val(M2)) And (Val(D1) < Val(D2)) Then RESULTADO = False
  
  ElseIf (Val(A1) = Val(A2) And Val(D1) = Val(D2)) And (Val(M1) < Val(M2)) Then RESULTADO = False
  
  ElseIf Val(A1) = Val(A2) And Val(M1) < Val(M2) And Val(D1) > Val(D2) Then RESULTADO = False

  ElseIf (Val(A1) = Val(A2) And Val(D1) < Val(D2)) And (Val(M1) < Val(M2)) Then RESULTADO = False

  ElseIf AÑO_MAYOR And MES_MAYOR And Not DIA_MAYOR Then RESULTADO = True

  ElseIf AÑO_MAYOR And MES_MAYOR Then
      If Val(D1) >= Val(D2) Then RESULTADO = True

  ElseIf AÑO_MAYOR And (Not MES_MAYOR Or Not DIA_MAYOR) Then RESULTADO = True

  End If


  FECHA_MAYOR_IGUAL = RESULTADO
End Function

Function FILE_EXISTS(ARCH As String) As Integer

 Dim T As String, R As Integer

    T$ = Dir(ARCH)
    If LTrim$(RTrim$(T)) = "" Then
      R = False
    Else
      R = True
    End If

 FILE_EXISTS = R
End Function

Function GET_FECHA(FECHA As String) As String
 ' ESTA FUNCIÓN TRANSFORMA LA FECHA CON FORMATO DDMMAA EN DD/MM/AA

 If Len(FECHA) <> 6 Then CAD$ = "ERROR"
 
 DIA = Mid$(FECHA, 1, 2)
 MES = Mid$(FECHA, 3, 2)
 AÑO = Mid$(FECHA, 5, 2)
 
 If Val(DIA) < 1 Or Val(DIA) > 31 Or Val(MES) < 1 Or Val(MES) > 12 Then CAD$ = "ERROR"

 
 If CAD$ <> "ERROR" Then
   If Val(AÑO) >= 80 Then
     SIGLO = "19"
   Else
     SIGLO = "20"
   End If

   CAD$ = DIA + "/" + MES + "/" + SIGLO + AÑO
 End If

 GET_FECHA = CAD$

End Function

Function GET_FILENAME(CODIGO_EMPRESA As String) As String
  GET_FILENAME = PATH_PROG + "EMPR" + CODIGO_EMPRESA + ".MDB"
End Function

Function GET_PCIA(PROVI As Integer) As String
  'ESTA FUNCIÓN DEVUELVE EL NOMBRE DE LA PCIA.
  'SEGÚN SU NUMERO

 ReDim P(23) As String * 30

 P(1) = "BUENOS AIRES"
 P(2) = "SANTA FE"
 P(3) = "LA PAMPA"
 P(4) = "CORDOBA"
 P(5) = "MENDOZA"
 P(6) = "SAN LUIS"
 P(7) = "SAN JUAN"
 P(8) = "CORRIENTES"
 P(9) = "MISIONES"
 P(10) = "ENTRE RÍOS"
 P(11) = "FORMOSA"
 P(12) = "SANTIAGO DEL ESTERO"
 P(13) = "CATAMARCA"
 P(14) = "SALTA"
 P(15) = "JUJUY"
 P(16) = "LA RIOJA"
 P(17) = "TUCUMÁN"
 P(18) = "NEUQUÉN"
 P(19) = "RÍO NEGRO"
 P(20) = "CHUBUT"
 P(21) = "SANTA CRUZ"
 P(22) = "TIERRA DEL FUEGO"

 GET_PCIA = P(PROVI)
End Function

Sub GRABAR_CUENTA(NUM_CUENTA As String, SUCU As String, INI As String, DEBE As String, HABER As String, TIPO_ASI As Integer, VALOR As Double)

  Dim MONTO_INI As Double, CAD1 As String, CAD2 As String

  MES = Val(LTrim$(RTrim$(Mid$(ASIENTOS.Text2.Text, 4, 2))))
  NUM_CUENTA = LTrim$(RTrim$(NUM_CUENTA))
  SUCU = LTrim$(RTrim$(SUCU))

  CAD1 = CUENTA + " = '" + NUM_CUENTA + "' AND " + CONT_SUCURSAL + " = '" + SUCU + "'"
  DN_CONTABLES.FindFirst CAD1
  If DN_CONTABLES.NoMatch Then
    Beep
    MsgBox "No se encontró la Cuenta Nº [" + SUCU + NUM_CUENTA + "]", 48, "Grabando Cuentas"
    Exit Sub
  End If

  NOM = LTrim$(RTrim$(DN_CONTABLES(DESCRIPCION) & ""))
  
  
  DN_CONTABLES.Edit

  '------------------------------------------------------------------
  'Si se trata de un Asiento Inicial...
  '------------------------------------------------------------------
  INI = LTrim$(RTrim$(UCase$(INI)))
  DN_ASIENTOS.Edit
  DN_ASIENTOS(CC_REPORTE_NOMBRECUENTA) = NOM

  If INI = "S" Then
     Select Case TIPO_ASI
       Case Is = PESOS
         MONTO_INI = Val(LTrim$(RTrim$(DN_CONTABLES(INICIAL) & "")))
         MONTO_INI = MONTO_INI + Val(DEBE) - Val(HABER)
         DN_CONTABLES(INICIAL) = LTrim$(RTrim$(Str$(MONTO_INI)))
         DN_ASIENTOS(CC_REPORTE_NOMBRECUENTA) = NOM

       Case Is = DOLAR
         MONTO_INI = Val(LTrim$(RTrim$(DN_CONTABLES(DINICIAL) & "")))
         MONTO_INI = MONTO_INI + Val(DEBE) - Val(HABER)
         DN_CONTABLES(DINICIAL) = LTrim$(RTrim$(Str$(MONTO_INI)))
         DN_ASIENTOS(CC_REPORTE_SALDO_INI) = DN_CONTABLES(DINICIAL) & ""

       Case Is = COMBINADOS
         MONTO_INI = Val(LTrim$(RTrim$(DN_CONTABLES(CINICIAL) & "")))
         MONTO_INI = MONTO_INI + Val(DEBE) - Val(HABER)
         DN_CONTABLES(CINICIAL) = LTrim$(RTrim$(Str$(MONTO_INI)))
         DN_ASIENTOS(CC_REPORTE_SALDO_INI) = DN_CONTABLES(CINICIAL) & ""

       Case Else
         Beep
         MsgBox "El tipo de Asiento no es reconocido por el Sistema.", 16, "Error en Datos"
         Exit Sub
     End Select
  End If
  DN_ASIENTOS.Update
  '------------------------------------------------------------------



  Select Case MES
    '------------------------------------------------------------------
    ' MES 1
    '------------------------------------------------------------------
    Case Is = 1
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO1) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO1) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO1) = MONTO_DEBE
           DN_CONTABLES(CREDITO1) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO1) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO1) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO1) = MONTO_DEBE
           DN_CONTABLES(DCREDITO1) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO1) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO1) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO1) = MONTO_DEBE
           DN_CONTABLES(CREDITO1) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO1) = MONTO_DEBE
           DN_CONTABLES(DCREDITO1) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 2
    '------------------------------------------------------------------
    Case Is = 2
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO2) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO2) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO2) = MONTO_DEBE
           DN_CONTABLES(CREDITO2) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO2) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO2) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO2) = MONTO_DEBE
           DN_CONTABLES(DCREDITO2) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO2) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO2) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO2) = MONTO_DEBE
           DN_CONTABLES(CREDITO2) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO2) = MONTO_DEBE
           DN_CONTABLES(DCREDITO2) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 3
    '------------------------------------------------------------------
    Case Is = 3
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO3) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO3) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO3) = MONTO_DEBE
           DN_CONTABLES(CREDITO3) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO3) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO3) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO3) = MONTO_DEBE
           DN_CONTABLES(DCREDITO3) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO3) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO3) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO3) = MONTO_DEBE
           DN_CONTABLES(CREDITO3) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO3) = MONTO_DEBE
           DN_CONTABLES(DCREDITO3) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 4
    '------------------------------------------------------------------
    Case Is = 4
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO4) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO4) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO4) = MONTO_DEBE
           DN_CONTABLES(CREDITO4) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO4) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO4) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO4) = MONTO_DEBE
           DN_CONTABLES(DCREDITO4) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO4) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO4) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO4) = MONTO_DEBE
           DN_CONTABLES(CREDITO4) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO4) = MONTO_DEBE
           DN_CONTABLES(DCREDITO4) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 5
    '------------------------------------------------------------------
    Case Is = 5
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO5) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO5) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO5) = MONTO_DEBE
           DN_CONTABLES(CREDITO5) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO5) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO5) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO5) = MONTO_DEBE
           DN_CONTABLES(DCREDITO5) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO5) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO5) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO5) = MONTO_DEBE
           DN_CONTABLES(CREDITO5) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO5) = MONTO_DEBE
           DN_CONTABLES(DCREDITO5) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 6
    '------------------------------------------------------------------
    Case Is = 6
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO6) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO6) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO6) = MONTO_DEBE
           DN_CONTABLES(CREDITO6) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO6) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO6) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO6) = MONTO_DEBE
           DN_CONTABLES(DCREDITO6) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO6) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO6) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO6) = MONTO_DEBE
           DN_CONTABLES(CREDITO6) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO6) = MONTO_DEBE
           DN_CONTABLES(DCREDITO6) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 7
    '------------------------------------------------------------------
    Case Is = 7
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO7) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO7) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO7) = MONTO_DEBE
           DN_CONTABLES(CREDITO7) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO7) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO7) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO7) = MONTO_DEBE
           DN_CONTABLES(DCREDITO7) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO7) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO7) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO7) = MONTO_DEBE
           DN_CONTABLES(CREDITO7) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO7) = MONTO_DEBE
           DN_CONTABLES(DCREDITO7) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 8
    '------------------------------------------------------------------
    Case Is = 8
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO8) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO8) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO8) = MONTO_DEBE
           DN_CONTABLES(CREDITO8) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO8) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO8) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO8) = MONTO_DEBE
           DN_CONTABLES(DCREDITO8) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO8) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO8) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO8) = MONTO_DEBE
           DN_CONTABLES(CREDITO8) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO8) = MONTO_DEBE
           DN_CONTABLES(DCREDITO8) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 9
    '------------------------------------------------------------------
    Case Is = 9
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO9) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO9) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO9) = MONTO_DEBE
           DN_CONTABLES(CREDITO9) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO9) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO9) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO9) = MONTO_DEBE
           DN_CONTABLES(DCREDITO9) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO9) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO9) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO9) = MONTO_DEBE
           DN_CONTABLES(CREDITO9) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO9) = MONTO_DEBE
           DN_CONTABLES(DCREDITO9) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 10
    '------------------------------------------------------------------
    Case Is = 10
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO10) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO10) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO10) = MONTO_DEBE
           DN_CONTABLES(CREDITO10) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO10) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO10) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO10) = MONTO_DEBE
           DN_CONTABLES(DCREDITO10) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO10) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO10) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO10) = MONTO_DEBE
           DN_CONTABLES(CREDITO10) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO10) = MONTO_DEBE
           DN_CONTABLES(DCREDITO10) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 11
    '------------------------------------------------------------------
    Case Is = 11
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO11) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO11) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO11) = MONTO_DEBE
           DN_CONTABLES(CREDITO11) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO11) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO11) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO11) = MONTO_DEBE
           DN_CONTABLES(DCREDITO11) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO11) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO11) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO11) = MONTO_DEBE
           DN_CONTABLES(CREDITO11) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO11) = MONTO_DEBE
           DN_CONTABLES(DCREDITO11) = MONTO_HABER
      End Select
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    ' MES 12
    '------------------------------------------------------------------
    Case Is = 12
      Select Case TIPO_ASI
         Case Is = PESOS
           MONTO_DEBE = DN_CONTABLES(DEBITO12) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO12) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO12) = MONTO_DEBE
           DN_CONTABLES(CREDITO12) = MONTO_HABER

         Case Is = DOLAR
           MONTO_DEBE = DN_CONTABLES(DDEBITO12) & ""
           MONTO_HABER = DN_CONTABLES(DCREDITO12) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DDEBITO12) = MONTO_DEBE
           DN_CONTABLES(DCREDITO12) = MONTO_HABER

         Case Is = COMBINADOS
           MONTO_DEBE = DN_CONTABLES(DEBITO12) & ""
           MONTO_HABER = DN_CONTABLES(CREDITO12) & ""
           MONTO_DEBE = LTrim$(RTrim$(Str$(Val(MONTO_DEBE) + Val(DEBE))))
           MONTO_HABER = LTrim$(RTrim$(Str$(Val(MONTO_HABER) + Val(HABER))))
           DN_CONTABLES(DEBITO12) = MONTO_DEBE
           DN_CONTABLES(CREDITO12) = MONTO_HABER

           'Equivalente en Dólares
           MONTO_DEBE = MONTO_DEBE / VALOR
           MONTO_HABER = MONTO_HABER / VALOR
           DN_CONTABLES(DDEBITO12) = MONTO_DEBE
           DN_CONTABLES(DCREDITO12) = MONTO_HABER
      End Select
    '------------------------------------------------------------------
  End Select

  DN_CONTABLES.Update

End Sub

Sub LIMPIAR_PLANILLA(VENTANA As Form, STARTING_ROW As Integer)

 'ESTE PROCEDIMIENTO LIMPIA EL OBJETO "GRID" PERTENECIENTE
 'A LA VENTANA QUE SE PASA COMO PARÁMETRO

  VENTANA!Grid1.FillStyle = 1
  VENTANA!Grid1.SelStartCol = 0
  VENTANA!Grid1.SelStartRow = STARTING_ROW
  VENTANA!Grid1.SelEndCol = VENTANA!Grid1.Cols - 1
  VENTANA!Grid1.SelEndRow = VENTANA!Grid1.Rows - 1
  VENTANA!Grid1.Text = ""
  VENTANA!Grid1.FillStyle = 0

End Sub

Sub LOAD_LISTA_ASIENTOS_CONTABLES()

  'Esta rutina carga el asiento nº "xxx" del periodo "dd/mm/yyyy" en la grilla.

  Dim CONTADOR As Integer
  Dim CUEN As String, LEYE As String, DEBI As String, CREDI As String, NOM As String
  Dim DOLAR As String, PESOS As String, RENG As String, FECH As String
  Dim DEBI_PESOS As String, DEBI_DOLAR As String, CREDI_PESOS As String
  Dim CREDI_DOLAR As String


  If ASIENTOS.Label5.Caption = "S" Then
    CON_SUCU = True
  Else
    CON_SUCU = False
  End If

  CLAVE = LTrim$(RTrim$(ASIENTOS.Text1.Text))
  FECHA_ASI = LTrim$(RTrim$(ASIENTOS.Text2.Text))
  FECHA_ASI = Format$(FECHA_ASI, "DD/MM/YYYY")


  OPEN_FILE_CUENTAS_CONTABLES
  
  FECHA_ASI = Mid$(FECHA_ASI, 4, 7)
  CADENA = "SELECT * FROM [" + TABLA_ASIENTOS_CONTABLES + "] WHERE " + CC_ASIENTO + " = '" + CLAVE + "'"
  CADENA = CADENA + " AND " + CC_FECHA + " >= '" + "01/" + FECHA_ASI + "' AND " + CC_FECHA + " <= '" + "31/" + FECHA_ASI + "'"
  
  Set DN_ASIENTOS = DB_ASIENTOS.CreateDynaset(CADENA)
  
  ASIENTOS.Grid1.Cols = 7
  ASIENTOS.Grid1.Rows = MAX_ENTRADAS_POR_ASIENTO
  ASIENTOS.Grid1.FixedRows = 1
  ASIENTOS.Grid1.FixedCols = 0
  
  LIMPIAR_PLANILLA ASIENTOS, 1

  SETGRID_ASIENTOS_CONTABLES 0, 0, "Cuenta"
  SETGRID_ASIENTOS_CONTABLES 0, 1, "Descripción"
  SETGRID_ASIENTOS_CONTABLES 0, 2, "Leyenda"
  SETGRID_ASIENTOS_CONTABLES 0, 3, "Débito"
  SETGRID_ASIENTOS_CONTABLES 0, 4, "Crédito"
  SETGRID_ASIENTOS_CONTABLES 0, 5, "Fecha"
  SETGRID_ASIENTOS_CONTABLES 0, 6, "Orden"
  

  If DN_ASIENTOS.RecordCount > 0 Then
     '------------------------------------------------------------------------
     'Actualización del Saldo Inicial
     'y del Nombre de la Empresa (este último sólo se usa en CRW)
     '------------------------------------------------------------------------
     BeginTrans
     DN_ASIENTOS.MoveFirst
     Do While Not DN_ASIENTOS.EOF
        RPTCUEN = DN_ASIENTOS(CC_CUENTA) & ""
        BUSCA_CUEN = CUENTA + " = '" + RPTCUEN + "' AND " + CONT_SUCURSAL + " = '" + LTrim$(RTrim$(DN_ASIENTOS(CC_SUCURSAL) & "")) + "'"
        DN_CONTABLES.FindFirst BUSCA_CUEN
        
        DN_ASIENTOS.Edit
        DN_ASIENTOS(CC_RPT_NOMBRE_EMPRESA) = LTrim$(RTrim$(MAIN.Label3.Caption))

        SALDO_INI = "0"
        If Not DN_CONTABLES.NoMatch Then
           TIPO_ASI = DN_ASIENTOS(CC_TIPO) & ""

           Select Case Val(TIPO_ASI)
             Case Is = 0
               SALDO_INI = DN_CONTABLES(INICIAL) & ""
             
             Case Is = 1
               SALDO_INI = DN_CONTABLES(DINICIAL) & ""
           
             Case Is = 3
               SALDO_INI = DN_CONTABLES(CINICIAL) & ""

           End Select
           
           If LTrim$(RTrim$(SALDO_INI)) = "" Then SALDO_INI = "0"
           DN_ASIENTOS(CC_REPORTE_SALDO_INI) = SALDO_INI
        Else
          Beep
          MsgBox "Error Interno. Nº de Cuenta [" + RPTCUEN + "] no encontrado.", 48, "Leyendo Asientos Contables"
        End If

        DN_ASIENTOS.Update
        DN_ASIENTOS.MoveNext
     Loop
    CommitTrans
    '-------------------------------------------------------------

     'Me aseguro de que los cambios ya fueron grabados
     DoEvents

     '------------------------------------------------------------------------
     'ACTUALIZAR EL Nº DE RENGLÓN  (Se debe hacer por separado, no se
     'puede poner en el loop anterior)
     '------------------------------------------------------------------------
     CONTADOR = 0
     BeginTrans
     DN_ASIENTOS.MoveFirst
     Do While Not DN_ASIENTOS.EOF
        CONTADOR = CONTADOR + 1
        DN_ASIENTOS.Edit
        DN_ASIENTOS(CC_RENGLON) = LTrim$(RTrim$(Str$(CONTADOR)))
        DN_ASIENTOS.Update
        DN_ASIENTOS.MoveNext
     Loop
    CommitTrans
    '-------------------------------------------------------------




    'ES NECESARIO CREAR NUEVAMENTE EL DYNASET
    'Para que los asientos aparezcan ordenados tal cual los ingresó
    'el usuario
    CADENA = "SELECT * FROM [" + TABLA_ASIENTOS_CONTABLES + "] WHERE " + CC_ASIENTO + " = '" + CLAVE + "'"
    CADENA = CADENA + " AND " + CC_FECHA + " >= '" + "01/" + FECHA_ASI + "' AND " + CC_FECHA + " <= '" + "31/" + FECHA_ASI + "' ORDER BY " + CC_RENGLON
    Set DN_ASIENTOS = DB_ASIENTOS.CreateDynaset(CADENA)

    CONTADOR = 0
    DN_ASIENTOS.MoveFirst
    
    INI = DN_ASIENTOS(CC_INICIAL) & ""
    OBS = DN_ASIENTOS(CC_OBSERVACIONES) & ""
    
    ASIENTOS.Text4.Text = INI
    ASIENTOS.Text9.Text = OBS

    Do While Not DN_ASIENTOS.EOF

       CONTADOR = CONTADOR + 1
       
       CUEN = DN_ASIENTOS(CC_CUENTA) & ""
       LEYE = DN_ASIENTOS(CC_DESCRIPCION) & ""
       OBS = DN_ASIENTOS(CC_OBSERVACIONES) & ""
       TIPO_ASI = DN_ASIENTOS(CC_TIPO) & ""
       FECH = DN_ASIENTOS(CC_FECHA) & ""
       ASIENTOS.Text2.Text = FECH
       INI = DN_ASIENTOS(CC_INICIAL) & ""
       RENG = DN_ASIENTOS(CC_RENGLON) & ""
       SUCU = DN_ASIENTOS(CC_SUCURSAL) & ""

       DEBI_PESOS = DN_ASIENTOS(CC_DEBITO) & ""
       DEBI_DOLAR = DN_ASIENTOS(CC_DDEBITO) & ""
       CREDI_PESOS = DN_ASIENTOS(CC_CREDITO) & ""
       CREDI_DOLAR = DN_ASIENTOS(CC_DCREDITO) & ""

       DEBI = ""
       CREDI = ""

       Select Case Val(TIPO_ASI)
         Case Is = 0
           'EN PESOS
            If DEBI_PESOS <> "" Then DEBI = DEBI_PESOS
            If CREDI_PESOS <> "" Then CREDI = CREDI_PESOS

         Case Is = 1
           'EN DOLARES
           If DEBI_DOLAR <> "" Then DEBI = DEBI_DOLAR
           If CREDI_DOLAR <> "" Then CREDI = CREDI_DOLAR

         Case Is = 2
           'COMBINADOS
            If DEBI_PESOS <> "" Then DEBI = DEBI_PESOS
            If CREDI_PESOS <> "" Then CREDI = CREDI_PESOS

       End Select
            
       If DEBI = "0" Then DEBI = ""
       If CREDI = "0" Then CREDI = ""


       CAD = CUENTA + " = '" + CUEN + "' AND " + CONT_SUCURSAL + " = '" + SUCU + "'"
       DN_CONTABLES.FindFirst CAD
       
       NOM = ""
       If Not DN_CONTABLES.NoMatch Then
          NOM = DN_CONTABLES(DESCRIPCION) & ""
       End If

       If CON_SUCU Then CUEN = SUCU + CUEN

       SETGRID_ASIENTOS_CONTABLES CONTADOR, 0, CUEN
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 1, NOM
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 2, LEYE
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 3, DEBI
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 4, CREDI
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 5, Format$(FECH, "MM/YYYY")
       SETGRID_ASIENTOS_CONTABLES CONTADOR, 6, RENG
       DN_ASIENTOS.MoveNext
    Loop

  Else
     ASIENTOS.Grid1.Rows = MAX_ENTRADAS_POR_ASIENTO
     ASIENTOS.Grid1.FixedRows = 1
     SETGRID_ASIENTOS_CONTABLES 1, 0, ""
     SETGRID_ASIENTOS_CONTABLES 1, 1, ""
     SETGRID_ASIENTOS_CONTABLES 1, 2, ""
     SETGRID_ASIENTOS_CONTABLES 1, 3, ""
     SETGRID_ASIENTOS_CONTABLES 1, 4, ""
     SETGRID_ASIENTOS_CONTABLES 1, 5, ""
     SETGRID_ASIENTOS_CONTABLES 1, 6, ""
  End If
  
  ASIENTOS.Grid1.ColWidth(0) = 1000
  ASIENTOS.Grid1.ColWidth(1) = 2600
  ASIENTOS.Grid1.ColWidth(2) = 2900
  ASIENTOS.Grid1.ColWidth(3) = 1300
  ASIENTOS.Grid1.ColWidth(4) = 1300
  ASIENTOS.Grid1.ColWidth(5) = 1
  ASIENTOS.Grid1.ColWidth(6) = 1
  
  ASIENTOS.Grid1.ColAlignment(3) = 1
  ASIENTOS.Grid1.ColAlignment(4) = 1
  ASIENTOS.Grid1.ColAlignment(5) = 2

  ASIENTOS.Grid1.Row = CONTADOR + 1
  ASIENTOS.Grid1.Col = 0

  If CONTADOR > 0 Then
    ASIENTOS.Grid1.SetFocus
  Else
    ASIENTOS.Text4.SetFocus
  End If
End Sub

Sub LOAD_LISTA_TIPOS()
  Dim CONTADOR As Integer
  Dim T As String, LEYE As String, CUEN As String
  Dim DESCRI As String, RENG As String, NOM_CUENTA As String


  OPEN_EMPRESAS
  COD_EMPRESA = MAIN.Label2.Caption
  BUSCA = CODIGO + " = '" + COD_EMPRESA + "'"
  DN_EMPRESAS.FindFirst BUSCA
  If DN_EMPRESAS.NoMatch Then
    DN_EMPRESAS.Close
    DB_EMPRESAS.Close
    Beep
    MsgBox "Error en el Archivo de Empresas. Verifique que los datos de la Empresa " + LTrim$(RTrim$(MAIN.Label3.Caption)) + " sean los correctos.", 16, "Error"
    Exit Sub
  Else
    FF = LTrim$(RTrim$(DN_EMPRESAS(TIENE_SUCURSALES) & ""))
    If FF = "" Then
      Beep
      MsgBox "Error en el Archivo de Empresas. Verifique que los datos de la Empresa " + LTrim$(RTrim$(MAIN.Label3.Caption)) + " sean los correctos.", 16, "Error"
      Exit Sub
    End If
  End If

  If FF = "S" Then
    CON_SUCU = True
  Else
    CON_SUCU = False
  End If
  
  TIP.Grid1.FixedRows = 1
  LIMPIAR_PLANILLA TIP, 1

  C = LTrim$(RTrim$(TIP.Text1.Text))
  
  If C = "" Then Exit Sub

  CADENA = "SELECT * FROM [" + TABLA_ASIENTOS_TIPO + "] WHERE " + AS_TIPO + " = '" + C + "' ORDER BY " + AS_CUENTA + ", " + AS_DESCRIPCION + ", " + AS_LEYENDA

  Set DN_TIPOS = DB_TIPOS.CreateDynaset(CADENA)

  TIP.Grid1.Rows = 3
  TIP.Grid1.Cols = 4

  SETGRID_TIPO 0, 0, "Cuenta"
  SETGRID_TIPO 0, 1, "Descripción del Asiento"
  SETGRID_TIPO 0, 2, "Leyenda"

  If DN_TIPOS.RecordCount > 0 Then

    'ACTUALIZAR EL CAMPO "RENGLÓN" Y CARGAR
    '       LOS DATOS EN LA PLANILLA

    OPEN_FILE_CUENTAS_CONTABLES
    
    DN_TIPOS.MoveFirst
    CONTADOR = 1
    BeginTrans
    
    Do While Not DN_TIPOS.EOF
      DN_TIPOS.Edit
      DN_TIPOS(AS_RENGLON) = CONTADOR

      NUMCUEN = DN_TIPOS(AS_CUENTA) & ""
      BUSCA = CUENTA + " = '" + NUMCUEN + "'"
      DN_CONTABLES.FindFirst BUSCA

      If Not DN_CONTABLES.NoMatch Then
        BUSCA = DN_CONTABLES(DESCRIPCION) & ""
        DN_TIPOS(AS_RPT_NOMBRECUENTA) = BUSCA
      End If

      CONTADOR = CONTADOR + 1
      DN_TIPOS.Update
      DN_TIPOS.MoveNext
    Loop


    CONTADOR = 1
    DN_TIPOS.MoveFirst
    Do While Not DN_TIPOS.EOF
      TIP.Grid1.Rows = CONTADOR + 1

      T = DN_TIPOS(AS_TIPO) & ""
      CUEN = DN_TIPOS(AS_CUENTA) & ""
      DESCRI = DN_TIPOS(AS_DESCRIPCION) & ""
      LEYE = DN_TIPOS(AS_LEYENDA) & ""
      RENG = DN_TIPOS(AS_RENGLON) & ""
      SUCU = DN_TIPOS(AS_SUCURSAL) & ""
      
      BUSCA = CUENTA + " = '" + CUEN + "'"
      DN_CONTABLES.FindFirst BUSCA
      If Not DN_CONTABLES.NoMatch Then
        NOM_CUENTA = DN_CONTABLES(DESCRIPCION) & ""
      Else
        NOM_CUENTA = ""
      End If

      If CON_SUCU Then CUEN = SUCU + CUEN

      SETGRID_TIPO CONTADOR, 0, CUEN
      SETGRID_TIPO CONTADOR, 1, NOM_CUENTA
      SETGRID_TIPO CONTADOR, 2, LEYE
      SETGRID_TIPO CONTADOR, 3, RENG
      
      CONTADOR = CONTADOR + 1
      DN_TIPOS.MoveNext
    Loop

    TIP.Grid1.Rows = CONTADOR
    CommitTrans
    
    DN_CONTABLES.Close
    DB_CONTABLES.Close
  Else
   TIP.Grid1.FixedRows = 1
   TIP.Grid1.Rows = 2
   SETGRID_TIPO 1, 0, ""
   SETGRID_TIPO 1, 1, ""
   SETGRID_TIPO 1, 2, ""
   SETGRID_TIPO 1, 3, ""
  End If
  TIP.Grid1.ColWidth(0) = 1000
  TIP.Grid1.ColWidth(1) = 3000
  TIP.Grid1.ColWidth(2) = 3000
  TIP.Grid1.ColWidth(3) = 1

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

Sub OPEN_EMPRESAS()
  On Error GoTo EMPRESA_ERRORES

  Set DB_EMPRESAS = OpenDatabase(PATH_PROG + FILENAME, False, False, FILE_TYPE)
  Set DN_EMPRESAS = DB_EMPRESAS.CreateDynaset(TABLA_EMPRESAS)
  Exit Sub

EMPRESA_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_EMPRESAS
        Resume

      Case Is = 3011
        DB_EMPRESAS.Close
        'No se encontró la TABLA especificada o
        CREAR_FILE_EMPRESAS
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select
End Sub

Sub OPEN_FILE_AJUSTE()
  On Error GoTo AJUSTE_ERRORES

  Set DB_AJUSTE = OpenDatabase(PATH_PROG + "EMPR" + LTrim$(RTrim$(MAIN.Label2.Caption)) + ".MDB", False, False, FILE_TYPE)
  Set DN_AJUSTE = DB_AJUSTE.CreateDynaset("BALANCE")

  Exit Sub

AJUSTE_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_AJUSTE
        Resume

      Case Is = 3011
        DB_AJUSTE.Close
        'No se encontró la TABLA especificada o
        CREAR_FILE_AJUSTE
        Set DB_AJUSTE = OpenDatabase(PATH_PROG + "EMPR" + LTrim$(RTrim$(MAIN.Label2.Caption)) + ".MDB", False, False, FILE_TYPE)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select

End Sub

Sub OPEN_FILE_ASIENTOS_CONTABLES()
  On Error GoTo ASIENTOS_ERRORES

  F$ = MAIN.Label2.Caption
  Set DB_ASIENTOS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
  Set DN_ASIENTOS = DB_ASIENTOS.CreateDynaset(TABLA_ASIENTOS_CONTABLES)
  Exit Sub

ASIENTOS_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_ASIENTOS_CONTABLES (F$)
        Resume

      Case Is = 3011
        DB_ASIENTOS.Close
        'No se encontró la TABLA especificada o
        CREAR_FILE_ASIENTOS_CONTABLES (F$)
        Set DB_ASIENTOS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select


End Sub

Sub OPEN_FILE_ASIENTOS_TIPO()
  On Error GoTo TIP_ERRORES

  F$ = MAIN.Label2.Caption
  Set DB_TIPOS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
  Set DN_TIPOS = DB_TIPOS.CreateDynaset(TABLA_ASIENTOS_TIPO)
  Exit Sub

TIP_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_ASIENTOS_TIPO (F$)
        Resume

      Case Is = 3011
        DB_TIPOS.Close
        'No se encontró la TABLA especificada o
        CREAR_FILE_ASIENTOS_TIPO (F$)
        Set DB_TIPOS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select

End Sub

Sub OPEN_FILE_COEFICIENTES()
  On Error GoTo COEF_ERRORES

  F$ = MAIN.Label2.Caption
  Set DB_COEF = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
  Set DN_COEF = DB_COEF.CreateDynaset(TABLA_COEFICIENTES)
  Exit Sub

COEF_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_COEF (F$)
        Resume

      Case Is = 3011
        DB_COEF.Close
        'No se encontró la TABLA especificada o
        CREAR_FILE_COEF (F$)
        Set DB_COEF = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select

End Sub

Sub OPEN_FILE_CUENTAS_CONTABLES()
  On Error GoTo CCONT_ERRORES

  F$ = MAIN.Label2.Caption
  Set DB_CONTABLES = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
  Set DN_CONTABLES = DB_CONTABLES.CreateDynaset(TABLA_CUENTAS_CONTABLES)
  Exit Sub

CCONT_ERRORES:
   E = Err
   Select Case E
      Case Is = 3011  'ERROR = "No se encontró la TABLA especificada"
        DB_CONTABLES.Close 'CERRAR
        CREAR_FILE_CUENTAS_CONTABLES (F$)  'CREAR
        'VOLVER A ABRIR
        Set DB_CONTABLES = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
        Resume

      Case Is = 3024
        'Tipo de Error = "El archivo no existe"
        CREAR_FILE_CUENTAS_CONTABLES (F$)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Cuentas Contables"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select
End Sub

Sub OPEN_FILE_LEYENDAS()
  On Error GoTo LEYENDA_ERRORES

  F$ = MAIN.Label2.Caption
  Set DB_LEYENDAS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
  Set DN_LEYENDAS = DB_LEYENDAS.CreateDynaset(TABLA_LEYENDAS)
  Exit Sub

LEYENDA_ERRORES:
   E = Err
   Select Case E
      
      Case Is = 3024
        'El archivo no existe.
        CREAR_FILE_LEYENDAS (F$)
        Resume

      Case Is = 3011
        DB_LEYENDAS.Close
        'No se encontró la TABLA especificada
        CREAR_FILE_LEYENDAS (F$)
        Set DB_LEYENDAS = OpenDatabase(GET_FILENAME(F$), False, False, FILE_TYPE)
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Leyendas"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select
End Sub

Sub OPEN_NOMBRE_INDICES()
  On Error GoTo NDX_ERROR

  Set DB_INDICES = OpenDatabase(PATH_PROG + FILENAME, False, False, FILE_TYPE)
  Set DN_INDICES = DB_INDICES.CreateDynaset(TABLA_NOMBRE_INDICES)
  Exit Sub

NDX_ERROR:
   E = Err
   Select Case E
      Case Is = 3011
        DB_INDICES.Close
        'ERROR = "No se encontró la TABLA especificada"
        CREAR_FILE_INDICES
        Set DB_INDICES = OpenDatabase(PATH_PROG + FILENAME, False, False, FILE_TYPE)
        Resume

      Case Is = 3024
        'Tipo de Error = "El archivo no existe"
        CREAR_FILE_INDICES
        Resume

      Case Else
        Beep
        MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
        End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select


End Sub

Sub OPEN_TIPO_CAMBIO()
  
  On Error GoTo CAMBIO_ERRORES
  Set DB_CAMBIO = OpenDatabase(PATH_PROG + FILENAME, False, False, FILE_TYPE)
  Set DN_CAMBIO = DB_CAMBIO.CreateDynaset(TABLA_TIPO_CAMBIO)
  Exit Sub

CAMBIO_ERRORES:
   E = Err
   Select Case E
      Case Is = 3011
        DB_CAMBIO.Close
        'ERROR = "No se encontró la TABLA especificada"
        CREAR_FILE_TIPO_CAMBIO
        Set DB_CAMBIO = OpenDatabase(PATH_PROG + FILENAME, False, False, FILE_TYPE)
        Resume

      Case Is = 3024
        'Tipo de Error = "El archivo no existe"
        CREAR_FILE_TIPO_CAMBIO
        Resume

       Case Else
         Beep
         MsgBox "Error Interno Nº " + Str$(Err) + " - " + Error$(Err) + " Sistema Abortado.", 16, "Error"
         End  'SALIR DEL SISTEMA
             'Y VOLVER A WINDOWS
  End Select


End Sub

Sub RESETEAR_CUENTAS()
   '--------------------------------------------------------------------
   'Poner en cero todas las Cuentas Contables
   '--------------------------------------------------------------------
   
   BeginTrans

   DN_CONTABLES.MoveFirst
   Do While Not DN_CONTABLES.EOF
     DN_CONTABLES.Edit

     DN_CONTABLES(DEBITO1) = 0
     DN_CONTABLES(CREDITO1) = 0
     DN_CONTABLES(DDEBITO1) = 0
     DN_CONTABLES(DCREDITO1) = 0
     
     DN_CONTABLES(DEBITO2) = 0
     DN_CONTABLES(CREDITO2) = 0
     DN_CONTABLES(DDEBITO2) = 0
     DN_CONTABLES(DCREDITO2) = 0

     DN_CONTABLES(DEBITO3) = 0
     DN_CONTABLES(CREDITO3) = 0
     DN_CONTABLES(DDEBITO3) = 0
     DN_CONTABLES(DCREDITO3) = 0
     
     DN_CONTABLES(DEBITO4) = 0
     DN_CONTABLES(CREDITO4) = 0
     DN_CONTABLES(DDEBITO4) = 0
     DN_CONTABLES(DCREDITO4) = 0

     DN_CONTABLES(DEBITO5) = 0
     DN_CONTABLES(CREDITO5) = 0
     DN_CONTABLES(DDEBITO5) = 0
     DN_CONTABLES(DCREDITO5) = 0
     
     DN_CONTABLES(DEBITO6) = 0
     DN_CONTABLES(CREDITO6) = 0
     DN_CONTABLES(DDEBITO6) = 0
     DN_CONTABLES(DCREDITO6) = 0

     DN_CONTABLES(DEBITO7) = 0
     DN_CONTABLES(CREDITO7) = 0
     DN_CONTABLES(DDEBITO7) = 0
     DN_CONTABLES(DCREDITO7) = 0
     
     DN_CONTABLES(DEBITO8) = 0
     DN_CONTABLES(CREDITO8) = 0
     DN_CONTABLES(DDEBITO8) = 0
     DN_CONTABLES(DCREDITO8) = 0
     
     DN_CONTABLES(DEBITO9) = 0
     DN_CONTABLES(CREDITO9) = 0
     DN_CONTABLES(DDEBITO9) = 0
     DN_CONTABLES(DCREDITO9) = 0

     DN_CONTABLES(DEBITO10) = 0
     DN_CONTABLES(CREDITO10) = 0
     DN_CONTABLES(DDEBITO10) = 0
     DN_CONTABLES(DCREDITO10) = 0

     DN_CONTABLES(DEBITO11) = 0
     DN_CONTABLES(CREDITO11) = 0
     DN_CONTABLES(DDEBITO11) = 0
     DN_CONTABLES(DCREDITO11) = 0

     DN_CONTABLES(DEBITO12) = 0
     DN_CONTABLES(CREDITO12) = 0
     DN_CONTABLES(DDEBITO12) = 0
     DN_CONTABLES(DCREDITO12) = 0

     'Poner en Cero los Saldos Iniciales
     DN_CONTABLES(INICIAL) = "0"
     DN_CONTABLES(DINICIAL) = "0"
     DN_CONTABLES(CINICIAL) = "0"

     DN_CONTABLES.Update
     DN_CONTABLES.MoveNext
   Loop
   CommitTrans
   '--------------------------------------------------------------------

End Sub

Sub SETGRID(FILA As Integer, COLUMNA As Integer, VALOR As String)
  LISTADOS.Grid1.Row = FILA
  LISTADOS.Grid1.Col = COLUMNA
  LISTADOS.Grid1.Text = VALOR
End Sub

Sub SETGRID_ASIENTOS_CONTABLES(FILA As Integer, COLUMNA As Integer, VALOR As String)
  ASIENTOS.Grid1.Row = FILA
  ASIENTOS.Grid1.Col = COLUMNA
  ASIENTOS.Grid1.Text = VALOR
End Sub

Sub SETGRID_TIPO(FILA As Integer, COLUMNA As Integer, VALOR As String)
  TIP.Grid1.Row = FILA
  TIP.Grid1.Col = COLUMNA
  TIP.Grid1.Text = VALOR
End Sub

Sub SHOW_LIST_EMPRESAS()
 If (DESTINO_REPORTE = REPORTE_PANTALLA) Or (DESTINO_REPORTE = REPORTE_IMPRESORA) Then
  LISTADOS.Report1.DataFiles(0) = PATH_PROG + "SYSTEM.MDB"
  LISTADOS.Report1.ReportFileName = PATH_PROG + "EMPRESAS.RPT"
  LISTADOS.Report1.SelectionFormula = "{EMPRESAS.CODIGO} > ''"
  LISTADOS.Report1.WindowTitle = "Listado de Empresas"
  LISTADOS.Report1.WindowTop = -3
  LISTADOS.Report1.WindowLeft = -3
  LISTADOS.Report1.WindowWidth = Screen.Width
  LISTADOS.Report1.WindowHeight = Screen.Height

  LISTADOS.Report1.SortFields(0) = "+{EMPRESAS.CODIGO}"
  LISTADOS.Report1.SortFields(1) = "+{EMPRESAS.NOMBRE}"

  LISTADOS.Report1.Destination = DESTINO_REPORTE
  LISTADOS.Report1.Action = 1
 End If
End Sub

Function SUMATORIA(FFF As Form, DESDE As Integer, HASTA As Integer, COLUMNA As Integer) As Double
  
 'ESTA FUNCIÓN SUMA UNA COLUMNA DESDE UNA FILA HASTA LA OTRA
  
  Dim ACUM As Double, A As Integer, NUM As String

  If COLUMNA > FFF!Grid1.Cols - 1 Then Exit Function
  If DESDE > FFF!Grid1.Rows - 1 Or HASTA > FFF!Grid1.Rows - 1 Then Exit Function
  If DESDE < 0 Then Exit Function

  FFF!Grid1.Col = COLUMNA
  ACUM = 0
  For A = DESDE To HASTA
    FFF!Grid1.Row = A
    NUM = FFF!Grid1.Text
    ACUM = ACUM + Val(NUM)
  Next

  SUMATORIA = ACUM
End Function

Sub TOTAL_CUENTA(PERIODO As Integer, TIPO As Integer, SUCU As String, CUEN As String, INI As String, DEBE As String, HABER As String, FINAL As String)

 '---------------------------------------------------------------------------
 ' Función
 '---------------------------------------------------------------------------
 'Esta rutina tiene por objetivo sumar los campos Debe, Haber e Inicial
 'de la Tabla de cuentas contables y devolver la sumatoria en las variables
 'que se pasan como parámetro.
 '---------------------------------------------------------------------------
 
 '---------------------------------------------------------------------------
 ' Parámetros
 '---------------------------------------------------------------------------
 ' PERIODO = Mes del cual se desea obtener el Sumas y Saldos
 ' TIPO = Indica si se deben tomar los valores de Pesos o Dólares
 ' SUCU = Nº de sucursal
 ' CUEN = Nº de Cuenta
 ' DEBE, HABER, FINAL, INI = La rutina devuelve en estas variables
 ' la suma de las columnas del Debe, Haber, Saldo inicial y Saldo Final.
 '---------------------------------------------------------------------------


 Dim DS As DYNASET
 Dim I As Double, D As Double, H As Double, F As Double
 
 CUEN = LTrim$(RTrim$(CUEN))
 SUCU = LTrim$(RTrim$(SUCU))

 Set DS = DB_CONTABLES.CreateDynaset("SELECT * FROM [" + TABLA_CUENTAS_CONTABLES + "] WHERE " + CONT_SUCURSAL + " = '" + SUCU + "' AND " + CUENTA + " LIKE '" + CUEN + "*'")

 I = 0
 F = 0
 D = 0
 H = 0
 
 If DS.RecordCount > 0 Then
   DS.MoveFirst
   Do While Not DS.EOF

   '---------------------------------------------------------------------------
   ' Obtener los valores del Débito y el Crédito según el Período
   ' elegido por el usuario
   '---------------------------------------------------------------------------
      Select Case PERIODO
         Case Is = 1
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO1)
             CREDI = 0 & DS(CREDITO1)
           Else
             DEBI = 0 & DS(DDEBITO1)
             CREDI = 0 & DS(DCREDITO1)
           End If

         Case Is = 2
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO2)
             CREDI = 0 & DS(CREDITO2)
           Else
             DEBI = 0 & DS(DDEBITO2)
             CREDI = 0 & DS(DCREDITO2)
           End If

         Case Is = 3
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO3)
             CREDI = 0 & DS(CREDITO3)
           Else
             DEBI = 0 & DS(DDEBITO3)
             CREDI = 0 & DS(DCREDITO3)
           End If

         Case Is = 4
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO4)
             CREDI = 0 & DS(CREDITO4)
           Else
             DEBI = 0 & DS(DDEBITO4)
             CREDI = 0 & DS(DCREDITO4)
           End If

         Case Is = 5
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO5)
             CREDI = 0 & DS(CREDITO5)
           Else
             DEBI = 0 & DS(DDEBITO5)
             CREDI = 0 & DS(DCREDITO5)
           End If

         Case Is = 6
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO6)
             CREDI = 0 & DS(CREDITO6)
           Else
             DEBI = 0 & DS(DDEBITO6)
             CREDI = 0 & DS(DCREDITO6)
           End If

         Case Is = 7
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO7)
             CREDI = 0 & DS(CREDITO7)
           Else
             DEBI = 0 & DS(DDEBITO7)
             CREDI = 0 & DS(DCREDITO7)
           End If

         Case Is = 8
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO8)
             CREDI = 0 & DS(CREDITO8)
           Else
             DEBI = 0 & DS(DDEBITO8)
             CREDI = 0 & DS(DCREDITO8)
           End If

         Case Is = 9
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO9)
             CREDI = 0 & DS(CREDITO9)
           Else
             DEBI = 0 & DS(DDEBITO9)
             CREDI = 0 & DS(DCREDITO9)
           End If

         Case Is = 10
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO10)
             CREDI = 0 & DS(CREDITO10)
           Else
             DEBI = 0 & DS(DDEBITO10)
             CREDI = 0 & DS(DCREDITO10)
           End If

         Case Is = 11
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO11)
             CREDI = 0 & DS(CREDITO11)
           Else
             DEBI = 0 & DS(DDEBITO11)
             CREDI = 0 & DS(DCREDITO11)
           End If

         Case Is = 12
           If TIPO = PESOS Then
             DEBI = 0 & DS(DEBITO12)
             CREDI = 0 & DS(CREDITO12)
           Else
             DEBI = 0 & DS(DDEBITO12)
             CREDI = 0 & DS(DCREDITO12)
           End If
      End Select

      '------------------------------------------------------------
      'Obtener el Saldo Inicial
      '------------------------------------------------------------
      If TIPO = PESOS Then
        SAL = LTrim$(RTrim$(DS(INICIAL) & ""))
      Else
        SAL = LTrim$(RTrim$(DS(DINICIAL) & ""))
      End If
      
      If SAL = "" Then
        SALDO_INI = 0
      Else
        SALDO_INI = Val(SAL)
      End If
   '---------------------------------------------------------------

     'Acumuladores
     I = I + SALDO_INI
     D = D + DEBI
     H = H + CREDI

     DS.MoveNext
   Loop

   F = (D - H) + I
   INI = LTrim$(RTrim$(Str$(I)))
   DEBE = LTrim$(RTrim$(Str$(D)))
   HABER = LTrim$(RTrim$(Str$(H)))
   FINAL = LTrim$(RTrim$(Str$(F)))
 Else
   INI = "0"
   DEBE = "0"
   HABER = "0"
   FINAL = "0"
 End If

End Sub

Sub VER_ASIENTO_TIPO()
  If Not DN_TIPOS.EOF Then
    TIP.Text1.Text = ""
    TIP.Text2.Text = ""
    TIP.Text3.Text = ""
    TIP.Text4.Text = ""
    TIP.Text5.Text = ""
    
    TIP.Text1.Text = DN_TIPOS(AS_TIPO) & ""
    TIP.Text2.Text = DN_TIPOS(AS_RENGLON) & ""
    TIP.Text3.Text = DN_TIPOS(AS_CUENTA) & ""
    TIP.Text4.Text = DN_TIPOS(AS_DESCRIPCION) & ""
    TIP.Text5.Text = DN_TIPOS(AS_LEYENDA) & ""
  End If
End Sub

Sub VER_CUENTA_CONTABLE()
  Dim CON_SUCU As Integer

  If CONTABLES.Label3.Caption = "S" Then
    CON_SUCU = True
  Else
    CON_SUCU = False
  End If

  If Not DN_CONTABLES.EOF Then
    CONTABLES.Text1.Text = ""
    CONTABLES.Text2.Text = ""
    CONTABLES.Text3.Text = ""
    CONTABLES.Text4.Text = ""
    CONTABLES.Check3D1.Value = False

    If CON_SUCU Then
      CAD1 = LTrim$(RTrim$(DN_CONTABLES(CUENTA))) & ""
      CAD2 = LTrim$(RTrim$(DN_CONTABLES(CONT_SUCURSAL))) & ""
      If CAD2 = "" Then CAD2 = "00"
      CONTABLES.Text1.Text = CAD2 + CAD1
    Else
      CONTABLES.Text1.Text = DN_CONTABLES(CUENTA) & ""
    End If

    CONTABLES.Text2.Text = DN_CONTABLES(DESCRIPCION) & ""
    CONTABLES.Text4.Text = DN_CONTABLES(IMPUTABLE) & ""
    CONTABLES.Text3.Text = DN_CONTABLES(Nivel) & ""

    If CONTABLES.Text4.Text = "S" Then CONTABLES.Check3D1.Value = True
  End If
End Sub

Sub VER_EMPRESA()
  If Not DN_EMPRESAS.EOF Then
    EMPRESA.Text1.Text = ""
    EMPRESA.Text2.Text = ""
    EMPRESA.Text3.Text = ""
    EMPRESA.Text4.Text = ""
    EMPRESA.Text5.Text = ""
    EMPRESA.Text6.Text = ""
    EMPRESA.Text7.Text = ""
    EMPRESA.Text8.Text = ""
    
    EMPRESA.Text1.Text = DN_EMPRESAS(CODIGO) & ""
    EMPRESA.Text3.Text = DN_EMPRESAS(Nombre) & ""
    EMPRESA.Text6.Text = DN_EMPRESAS(DIRECCION) & ""
    EMPRESA.Text8.Text = DN_EMPRESAS(LOCALIDAD) & ""
    EMPRESA.Text4.Text = DN_EMPRESAS(CUIT) & ""
    EMPRESA.Text9.Text = DN_EMPRESAS(TIENE_SUCURSALES) & ""
    EMPRESA.Combo1.Text = DN_EMPRESAS(PCIA) & ""

     Q1$ = DN_EMPRESAS(INICIO) & ""
     EMPRESA.Text2.Text = Format$(Q1$, "DD/MM/YYYY")

     Q2$ = DN_EMPRESAS(FIN) & ""
     EMPRESA.Text5.Text = Format$(Q2$, "DD/MM/YYYY")

     EMPRESA.Text7.Text = DN_EMPRESAS(ACTIVIDAD) & ""

     'A% = 0 & DN_EMPRESAS(PCIA)
     'If A% >= 0 And A% <= 22 Then EMPRESA.Combo1.ListIndex = A%

     A% = 0 & DN_EMPRESAS(MONEDA)
     If A% >= 0 And A% <= 2 Then EMPRESA.Combo2.ListIndex = A%

  End If
End Sub

Sub VER_LEYENDAS()
  If Not DN_LEYENDAS.EOF Then
    LEYENDAS.Text1.Text = DN_LEYENDAS(CODIGO_LEYENDA) & ""
    LEYENDAS.Text2.Text = DN_LEYENDAS(LEYENDA) & ""
  End If
End Sub

Sub VER_LISTA_CONTABLES()

  Dim CUEN As String
  Dim DESCRI As String
  Dim NIV As String
  Dim IMPU As String
  Dim CON_SUCU As Integer

  Dim A As Integer

  If CONTABLES.Label3.Caption = "S" Then
    CON_SUCU = True
  Else
    CON_SUCU = False
  End If

  Set DN_CONTABLES = DB_CONTABLES.CreateDynaset(TABLA_CUENTAS_CONTABLES)
  DN_CONTABLES.Sort = CONT_SUCURSAL + ", " + CUENTA
  Set DN_CONTABLES = DN_CONTABLES.CreateDynaset()

 
 If DN_CONTABLES.RecordCount = 0 Then
  Beep
  MsgBox "El archivo está vacío.", 48, "Listado"
  Exit Sub
 
 Else
  
  Set DN_CONTABLES = DB_CONTABLES.CreateDynaset(TABLA_CUENTAS_CONTABLES)
  DN_CONTABLES.Sort = CONT_SUCURSAL + ", " + CUENTA
  Set DN_CONTABLES = DN_CONTABLES.CreateDynaset()

  LISTADOS.Grid1.Rows = 1000
  LISTADOS.Grid1.Cols = 8
  
  LIMPIAR_PLANILLA LISTADOS, 1

  SETGRID 0, 0, "Cuenta"
  SETGRID 0, 1, "Descripción"
  SETGRID 0, 2, "Nivel"
  SETGRID 0, 3, "Imputable"

  A = 1
  Do While Not DN_CONTABLES.EOF
     CUEN = ""
     DESCRI = ""
     NIV = ""
     IMPU = ""

     CUEN = LTrim$(RTrim$(DN_CONTABLES(CUENTA) & ""))
     DESCRI = DN_CONTABLES(DESCRIPCION) & ""
     NIV = DN_CONTABLES(Nivel) & ""
     IMPU = DN_CONTABLES(IMPUTABLE) & ""
     SUCU = DN_CONTABLES(CONT_SUCURSAL) & ""

     If CON_SUCU Then CUEN = SUCU + CUEN
     SETGRID A, 0, CUEN
     SETGRID A, 1, DESCRI
     SETGRID A, 2, NIV
     SETGRID A, 3, IMPU
     A = A + 1
     DN_CONTABLES.MoveNext
  Loop
  LISTADOS.Grid1.ColWidth(0) = 1000
  LISTADOS.Grid1.ColWidth(1) = 4800
  LISTADOS.Grid1.ColWidth(2) = 1000
  LISTADOS.Grid1.ColWidth(3) = 1500
  LISTADOS.Grid1.ColAlignment(2) = 2
  LISTADOS.Grid1.ColAlignment(3) = 2
  
  LISTADOS.Show 1

End If

End Sub

Sub VER_LISTA_EMPRESAS()
  Dim SUCU As String

If DN_EMPRESAS.RecordCount = 0 Then
  Beep
  MsgBox "El archivo está vacío.", 48, "Listado"
  Exit Sub
Else
  LISTADOS.Caption = "Lista de Empresas"
  
  LISTADOS.Grid1.BorderStyle = 1
  LISTADOS.Grid1.Cols = 11
  LISTADOS.Grid1.Rows = 100
  LIMPIAR_PLANILLA LISTADOS, 1

  SETGRID 0, 0, "Código"
  SETGRID 0, 1, "Sucursal"
  SETGRID 0, 2, "Nombre"
  SETGRID 0, 3, "Dirección"
  SETGRID 0, 4, "Localidad"
  SETGRID 0, 5, "Provincia"
  SETGRID 0, 6, "Fecha de Inicio"
  SETGRID 0, 7, "Fecha de Cierre"
  SETGRID 0, 8, "CUIT"
  SETGRID 0, 9, "Actividad Económica"

  SQL1 = "SELECT * FROM " + TABLA_EMPRESAS + " ORDER BY " + CODIGO + ", " + Nombre
  Set DN_EMPRESAS = DB_EMPRESAS.CreateDynaset(SQL1)

  A% = 1
  DN_EMPRESAS.MoveFirst

  Do While Not DN_EMPRESAS.EOF
  
     N_CODIGO$ = DN_EMPRESAS(CODIGO) & ""
     N_NOMBRE$ = DN_EMPRESAS(Nombre) & ""
     
     N_DIRECCION$ = DN_EMPRESAS(DIRECCION) & ""
     N_LOCALIDAD$ = DN_EMPRESAS(LOCALIDAD) & ""
     N_CUIT$ = DN_EMPRESAS(CUIT) & ""
     Q1$ = DN_EMPRESAS(INICIO) & ""
     
     N_INI$ = GET_FECHA(Format$(Q1$, "DDMMYY"))
     Q2$ = DN_EMPRESAS(FIN) & ""
     N_FIN$ = GET_FECHA(Format$(Q2$, "DDMMYY"))
     N_RUBRO$ = DN_EMPRESAS(ACTIVIDAD) & ""



     N_PCIA$ = DN_EMPRESAS(PCIA) & ""
     SETGRID A%, 0, N_CODIGO$
     SETGRID A%, 2, N_NOMBRE$
     SETGRID A%, 3, N_DIRECCION$
     SETGRID A%, 4, N_LOCALIDAD$
     SETGRID A%, 5, N_PCIA$
     SETGRID A%, 6, N_INI$
     SETGRID A%, 7, N_FIN$
     SETGRID A%, 8, N_CUIT$
     SETGRID A%, 9, N_RUBRO$
     
     A% = A% + 1
     DN_EMPRESAS.MoveNext
  Loop

  For A% = 0 To LISTADOS.Grid1.Cols - 1
    LISTADOS.Grid1.ColAlignment(A%) = 0
  Next
                               
  LISTADOS.Grid1.ColWidth(0) = 800
    
  LISTADOS.Grid1.ColWidth(1) = 1

  LISTADOS.Grid1.ColWidth(2) = 3000
  LISTADOS.Grid1.ColWidth(3) = 3000
  LISTADOS.Grid1.ColWidth(4) = 3000
  LISTADOS.Grid1.ColWidth(5) = 3000
  LISTADOS.Grid1.ColWidth(6) = 2000
  LISTADOS.Grid1.ColWidth(7) = 2000
  LISTADOS.Grid1.ColWidth(8) = 2700
  LISTADOS.Grid1.ColWidth(9) = 15000

  LISTADOS.Show

End If
End Sub

Sub VER_LISTA_INDICES()
  Dim COD As String
  Dim NOM As String
  Dim A As Integer

If DN_INDICES.RecordCount = 0 Then
  Beep
  MsgBox "El archivo está vacío.", 48, "Listado"
  Exit Sub
Else
  SQL = "SELECT * FROM [" + TABLA_NOMBRE_INDICES + "] ORDER BY " + INDEX_CODE
  Set DN_INDICES = DB_INDICES.CreateDynaset(SQL)

  If DN_INDICES.RecordCount > 0 Then

    BeginTrans
    Do While Not DN_INDICES.EOF
       COD = LTrim$(RTrim$(DN_INDICES(INDEX_CODE) & ""))
       COD = Space$(4 - Len(COD)) + COD
       DN_INDICES.Edit
       DN_INDICES(INDEX_CODE) = COD
       DN_INDICES.Update
       DN_INDICES.MoveNext
    Loop
    CommitTrans

    SQL = "SELECT * FROM [" + TABLA_NOMBRE_INDICES + "] ORDER BY " + INDEX_CODE
    Set DN_INDICES = DB_INDICES.CreateDynaset(SQL)
    
    LISTADOS.Grid1.Cols = 10
    LISTADOS.Grid1.Rows = 100
    LIMPIAR_PLANILLA LISTADOS, 1

    SETGRID 0, 0, "Código"
    SETGRID 0, 1, "Nombre del Índice"

    A = 1
    Do While Not DN_INDICES.EOF
       COD = DN_INDICES(INDEX_CODE) & ""
       NOM = DN_INDICES(INDEX_NAME) & ""
       SETGRID A, 0, COD
       SETGRID A, 1, NOM
       A = A + 1
       DN_INDICES.MoveNext
    Loop
    LISTADOS.Grid1.ColWidth(0) = 1000
    LISTADOS.Grid1.ColWidth(1) = 5800

    LISTADOS.Show 1
  
    BeginTrans
    Do While Not DN_INDICES.EOF
       COD = LTrim$(RTrim$(DN_INDICES(INDEX_CODE) & ""))
       DN_INDICES.Edit
       DN_INDICES(INDEX_CODE) = COD
       DN_INDICES.Update
       DN_INDICES.MoveNext
    Loop
    CommitTrans
  
  End If
End If
End Sub

Sub VER_LISTA_LEYENDAS()
  Dim COD As String
  Dim NOM As String
  Dim A As Integer

If DN_LEYENDAS.RecordCount = 0 Then
  Beep
  MsgBox "El archivo está vacío.", 48, "Listado"
  Exit Sub
Else
  SQL1 = "SELECT * FROM [" + TABLA_LEYENDAS + "] ORDER BY " + CODIGO_LEYENDA
  Set DN_LEYENDAS = DB_LEYENDAS.CreateDynaset(SQL1)

  LIMPIAR_PLANILLA LISTADOS, 1
  LISTADOS.Grid1.Cols = 10
  LISTADOS.Grid1.Rows = 100
  SETGRID 0, 0, "Código"
  SETGRID 0, 1, "Leyenda"

  BeginTrans
  'Reordenar la lista de leyendas
  DN_LEYENDAS.MoveFirst
  Do While Not DN_LEYENDAS.EOF
     DN_LEYENDAS.Edit
     COD = LTrim$(RTrim$(DN_LEYENDAS(CODIGO_LEYENDA) & ""))
     COD = Space$(4 - Len(COD)) + COD
     DN_LEYENDAS(CODIGO_LEYENDA) = COD
     DN_LEYENDAS.Update
     DN_LEYENDAS.MoveNext
  Loop
  CommitTrans

  Set DN_LEYENDAS = DB_LEYENDAS.CreateDynaset(SQL1)
  
  A = 1
  DN_LEYENDAS.MoveFirst
  Do While Not DN_LEYENDAS.EOF
     COD = DN_LEYENDAS(CODIGO_LEYENDA) & ""
     NOM = DN_LEYENDAS(LEYENDA) & ""
     SETGRID A, 0, COD
     SETGRID A, 1, NOM
     A = A + 1
     DN_LEYENDAS.MoveNext
  Loop
  LISTADOS.Grid1.ColWidth(0) = 1000
  LISTADOS.Grid1.ColWidth(1) = 5800

  BeginTrans
  'Restaurar la lista de leyendas
  DN_LEYENDAS.MoveFirst
  Do While Not DN_LEYENDAS.EOF
     DN_LEYENDAS.Edit
     COD = LTrim$(RTrim$(DN_LEYENDAS(CODIGO_LEYENDA) & ""))
     DN_LEYENDAS(CODIGO_LEYENDA) = COD
     DN_LEYENDAS.Update
     DN_LEYENDAS.MoveNext
  Loop
  CommitTrans

  LISTADOS.Show 1

End If
End Sub

Sub VER_LISTA_TIPO_CAMBIO()

Dim FECHA As String, V As String

If DN_CAMBIO.RecordCount = 0 Then
  Beep
  MsgBox "El archivo está vacío.", 48, "Listado"
  Exit Sub
Else
  C = "SELECT * FROM [" + TABLA_TIPO_CAMBIO + "] ORDER BY " + FECHA_CAMBIO
  
  Set DN_CAMBIO = DB_CAMBIO.CreateDynaset(C)

  LISTADOS.Caption = "Listado de Tipos de Cambio"
  LISTADOS.Grid1.Cols = 20
  LISTADOS.Grid1.Rows = 1999

  SETGRID 0, 0, "FECHA"
  SETGRID 0, 1, "TIPO DE CAMBIO"

  A% = 1
  DN_CAMBIO.MoveFirst
  Do While Not DN_CAMBIO.EOF
    
    FECHA = DN_CAMBIO(FECHA_CAMBIO) & ""
    
    If IsNumeric(DN_CAMBIO(VALOR_CAMBIO)) Then
      S = VarType(DN_CAMBIO(VALOR_CAMBIO))
      If S >= 2 And S <= 6 Then
        V = Str$(DN_CAMBIO(VALOR_CAMBIO))
      Else
        V = "ERROR"
      End If
    End If
    
    SETGRID A%, 0, GET_FECHA(Format$(FECHA, "DDMMYY"))
    SETGRID A%, 1, V
    DN_CAMBIO.MoveNext
    A% = A% + 1
  Loop
  LISTADOS.Grid1.ColWidth(0) = 1500
  LISTADOS.Grid1.ColWidth(1) = 4000
  
  LISTADOS.Show
End If

End Sub

Sub VER_NOMBRE_INDICES()
  If Not DN_INDICES.EOF Then
    NDXNAMES.Text1.Text = DN_INDICES(INDEX_CODE) & ""
    NDXNAMES.Text2.Text = DN_INDICES(INDEX_NAME) & ""
  End If
End Sub

Sub VER_TIPO_CAMBIO()
  If Not DN_CAMBIO.EOF Then
    A1$ = DN_CAMBIO(FECHA_CAMBIO) & ""
    If A1$ <> "" Then A1$ = Format$(A1$, "DD/MM/YYYY")
    CAMBIO1.Text1.Text = A1$
    CAMBIO1.Text2.Text = DN_CAMBIO(VALOR_CAMBIO) & ""
  End If
End Sub

