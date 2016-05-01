VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgVariosEste 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Comprobantes por Concepto Varios de Exportacion"
   ClientHeight    =   8130
   ClientLeft      =   315
   ClientTop       =   405
   ClientWidth     =   11295
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11295
   Visible         =   0   'False
   Begin VB.CommandButton ReImpresionII 
      Caption         =   "ReImpresion Factura"
      Height          =   615
      Left            =   9720
      TabIndex        =   38
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Cae 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   36
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Dolar2 
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   5520
      Width           =   5055
   End
   Begin VB.TextBox Dolar1 
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   5160
      Width           =   5055
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Comprobante"
      Height          =   1095
      Left            =   9240
      TabIndex        =   28
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Credito 
         Caption         =   "Nota de Credito"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Debito 
         Caption         =   "Nota de Debito"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Factura 
         Caption         =   "Factura Varias"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta de Datos"
      Height          =   495
      Left            =   9840
      TabIndex        =   26
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglones"
      Height          =   495
      Left            =   9840
      TabIndex        =   25
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   4440
      Width           =   8655
      Begin VB.TextBox WDescripcion 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   27
         Text            =   " "
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox WImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   23
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   21
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      Height          =   495
      Left            =   9840
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   6600
      TabIndex        =   17
      Top             =   5160
      Width           =   2415
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9840
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9840
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1815
      Left            =   2280
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   9840
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   450
      Left            =   9840
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   9840
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3135
      Left            =   360
      OleObjectBlob   =   "variosExte.frx":0000
      TabIndex        =   3
      Top             =   1200
      Width           =   8655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1815
      ItemData        =   "variosExte.frx":09EA
      Left            =   120
      List            =   "variosExte.frx":09F1
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label23 
      Caption         =   "Cae"
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "Dolar"
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Comprobante"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgVariosEste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 2 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private Auxiliar(100, 2) As String
Private Articulo As String
Private Auxi As String
Private Auxi1 As String
Private Renglon As Integer
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstDesccomp As Recordset
Dim spDesccomp As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String
Dim Compara As Double
Private WCodIb As Integer
Private WImpoIb As Double
Dim WNro As String
Dim WImpresion(100, 10) As String
Private WTexto1 As String
Private WTexto2 As String

Dim ZZComprobante As Integer
Dim ZZCuit As String
Dim ZZPais As String
Dim ZZCuitII As String
Dim ZZRazon As String
Dim ZZDomicilio As String
Dim ZZFechaCae As String

Dim ZZGrabaFactura As String

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
        rstPago.Close
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias2 = rstPago!Dias
        WPlazo2 = rstPago!Plazo
        rstPago.Close
    End If
    
    Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)

End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    WDescripcion.Text = ""
    WImporte.Text = ""
    WLinea.Text = ""
    
    WDescripcion.SetFocus

End Sub

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.Visible = True
     
 End Sub



Private Sub Dolar1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dolar2.SetFocus
    End If
End Sub

Private Sub Dolar2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dolar1.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstClientes!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()
    
    WCol = DBGrid1.Col
    WRow = DBGrid1.Row
    
    DBGrid1.Col = WCol
    DBGrid1.Row = WRow
    
    DBGrid1.Col = 0
    WDescri = DBGrid1.Text
    
    DBGrid1.Col = 1
    WImporte = DBGrid1.Text
    
    If WDescri = "" And Val(WImporte) = 0 Then
        WDescripcion.Text = ""
        WLinea.Text = ""
            Else
        WLinea.Text = DBGrid1.Row + 1
        WDescripcion.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 0
    WDescripcion.Text = DBGrid1.Text

    DBGrid1.Col = 1
    If Val(DBGrid1.Text) <> 0 Then
        WImporte.Text = DBGrid1.Text
            Else
        WImporte.Text = ""
    End If
    
    WDescripcion.SetFocus
    
    If Fecha.Text = "  /  /    " Or Cliente.Text = "" Then
         Numero.SetFocus
    End If

End Sub

Private Sub Calcula_Click()

    WNeto = 0

    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 1
            WImporte = Val(DBGrid1.Text)
                    
            WNeto = WNeto + WImporte
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0
    
    WIva1 = 0
    WIva2 = 0
    WImpoIb = 0
    
    WTotal = WNeto
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()
    
    If Trim(Cae.Text) = "" Then
        ZZGrabaFactura = ""
        Call Calcula_Cae
        If ZZGrabaFactura <> "S" Then
            Exit Sub
        End If
    End If
    
    Pasa = "S"

    WPago1 = 1
    WPago2 = 1
    Call Calcula_FechaVto

    Cliente.Text = UCase(Cliente.Text)
        
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Call Calcula_Click
        
    If Factura.Value = True Then
        WTipo = "03"
        WImpre = "FV"
    End If
    If Debito.Value = True Then
        WTipo = "04"
        WImpre = "ND"
    End If
    If Credito.Value = True Then
        WTipo = "05"
        WImpre = "NC"
    End If
        
    WNumero = Numero.Text
    WRenglon = "01"
    WCliente = Cliente.Text
    WFecha = Fecha.Text
    WEstado = "0"
    Call Convierte_datos(Str$(Total), Auxi)
    If Credito.Value = False Then
        XTotal = Str$(WTotal)
        XTotalUs = Str$(WTotal)
        XSaldo = Str$(WTotal)
        XSaldoUs = Str$(WTotal)
        XNet = Str$(WNeto * Val(Paridad.Text))
        XIva1 = Str$(WIva1 * Val(Paridad.Text))
        XIva2 = Str$(WIva2 * Val(Paridad.Text))
        XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
        XSeguro = ""
        XFlete = ""
            Else
        XTotal = Str$(WTotal * -1)
        XTotalUs = Str$(WTotal * -1)
        XSaldo = Str$(WTotal * -1)
        XSaldoUs = Str$(WTotal * -1)
        XNet = Str$(WNeto * -1 * Val(Paridad.Text))
        XIva1 = Str$(WIva1 * -1 * Val(Paridad.Text))
        XIva2 = Str$(WIva2 * -1 * Val(Paridad.Text))
        XImpoIb = Str$(WImpoIb * -1 * Val(Paridad.Text))
        XSeguro = ""
        XFlete = ""
    End If
            
    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
    WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
    WPedido = ""
    WRemito = ""
    WOrden = ""
    WParidad = Paridad.Text
    WProvincia = WProvincia
    XVendedor = Str$(WVendedor)
    XRubro = Str$(WRubro)
    WComprobante = ""
    WAceptada = ""
    WCosto = ""
    WImporte1 = ""
    WImporte2 = ""
    WImporte3 = ""
    WImporte4 = ""
    WImporte5 = ""
    WImporte6 = ""
    WImporte7 = ""
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = WTipo + Auxi + "01"
    XEmpresa = "1"
    WDate = Date$
    WNroFactura = ""
    WNroRecibo = ""
        
    XParam = "'" + WClave + "','" _
                 + WTipo + "','" + WNumero + "','" _
                 + WRenglon + "','" + WCliente + "','" _
                 + WFecha + "','" + WEstado + "','" _
                 + Wvencimiento + "','" + WVencimiento1 + "','" _
                 + XTotal + "','" + XTotalUs + "','" _
                 + XSaldo + "','" + XSaldoUs + "','" _
                 + WOrdFecha + "','" + WOrdVencimiento + "','" _
                 + WOrdVencimiento1 + "','" + WImpre + "','" _
                 + XEmpresa + "','" _
                 + XNet + "','" + XIva1 + "','" _
                 + XIva2 + "','" + WPedido + "','" _
                 + WRemito + "','" + WOrden + "','" _
                 + WParidad + "','" + WProvincia + "','" _
                 + XVendedor + "','" + XRubro + "','" _
                 + WComprobante + "','" + WAceptada + "','" _
                 + WCosto + "','" _
                 + WImporte1 + "','" + WImporte2 + "','" _
                 + WImporte3 + "','" + WImporte4 + "','" _
                 + WImporte5 + "','" + WImporte6 + "','" _
                 + WImporte7 + "','" + WDate + "','" _
                 + XSeguro + "','" + XFlete + "','" _
                 + XImpoIb + "','" + WNroFactura + "','" _
                 + WNroRecibo + "'"
                        
    spCtacte = "AltaCtacteVarios " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    ZZImpreNumero = Auxi
        
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "Gastos = " + "'" + "0" + "',"
    ZSql = ZSql & "ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql & "Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql & "FechaCae = " + "'" + "  /  /    " + "',"
    ZSql = ZSql & "Marca = " + "'" + "" + "',"
    ZSql = ZSql & "Envio1 = " + "'" + "" + "',"
    ZSql = ZSql & "Envio2 = " + "'" + "" + "',"
    ZSql = ZSql & "Pago1 = " + "'" + "" + "',"
    ZSql = ZSql & "Pago2 = " + "'" + "" + "',"
    ZSql = ZSql & "NroOrden = " + "'" + "" + "',"
    ZSql = ZSql & "FecOrden = " + "'" + "" + "',"
    ZSql = ZSql & "Consignatario = " + "'" + "" + "',"
    ZSql = ZSql & "Cip = " + "'" + "" + "',"
    ZSql = ZSql & "CipLista = " + "'" + "" + "',"
    ZSql = ZSql & "Idioma = " + "'" + "" + "',"
    ZSql = ZSql & "ImpreDolar1 = " + "'" + Dolar1.Text + "',"
    ZSql = ZSql & "ImpreDolar2 = " + "'" + Dolar2.Text + "',"
    ZSql = ZSql & "ImpreTotal = " + "'" + "" + "',"
    ZSql = ZSql & "ImpreTotalBruto = " + "'" + "" + "',"
    ZSql = ZSql & "ImpreTotalNeto = " + "'" + "" + "'"
    ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
    ZZClaveCtaCte = WClave
        
        
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            WDescripcion = DBGrid1.Text
                    
            DBGrid1.Col = 1
            WImporte = DBGrid1.Text
                    
            DBGrid1.Col = 0
            WDescripcion = DBGrid1.Text
                    
            If WDescripcion <> "" Or Val(WImporte) <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                        
                If Factura.Value = True Then
                    WTipo = "03"
                End If
                If Debito.Value = True Then
                    WTipo = "04"
                End If
                If Credito.Value = True Then
                    WTipo = "05"
                End If
                        
                WNumero = Numero.Text
                WRenglon = Str$(Renglon)
                WImporte = WImporte
                XEmpresa = "1"
                    
                WClave = WTipo + Auxi1 + Auxi
                WDate = Date$
                    
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WNumero + "','" _
                        + WRenglon + "','" _
                        + WDescripcion + "','" _
                        + WImporte + "','" _
                        + XEmpresa + "','" _
                        + WDate + "'"
                        
                spDesccomp = "AltaDesccomp " + XParam
                Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
                        
                ZSql = ""
                ZSql = ZSql & "UPDATE Desccomp SET "
                ZSql = ZSql & "ClaveCtaCte = " + "'" + ZZClaveCtaCte + "',"
                ZSql = ZSql & "Cliente = " + "'" + Cliente.Text + "'"
                ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
                spDesccomp = ZSql
                Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
                        
            End If
                                        
        Next iRow
            
    Next a
        
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            rstNumero.Close
            WCodigo = "02"
            WNumero = Numero.Text
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            rstNumero.Close
            WCodigo = "08"
            WNumero = Numero.Text
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            rstNumero.Close
            WCodigo = "09"
            WNumero = Numero.Text
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
    End If
        
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Call Impresion
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
    
    WDescripcion.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cae.Text = ""
    Dolar1.Text = ""
    Dolar2.Text = ""
    
    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
  
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 1
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Total.Caption = ""
    Paridad.Text = ""
    
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    
    Factura.Value = True
    Debito.Value = False
    Credito.Value = False
    
    Graba.Enabled = True
    Borra.Enabled = True
    Ingresa.Enabled = True
    
    Numero.SetFocus

End Sub

Private Sub ReImpresionII_Click()
    Call Impresion
End Sub

Private Sub WDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.SetFocus
    End If
End Sub

Private Sub WImporte_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.Text = Pusing("###,###.##", WImporte.Text)
        Call Alta_Vector
        Call Ingresa_Click
        Call Calcula_Click
        WImporte.Text = ""
        WDescripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spClientes = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                WPago1 = 1
                WPago2 = 1
                WVendedor = rstClientes!vendedor
                WProvincia = rstClientes!Provincia
                WRubro = rstClientes!Rubro
                WCodIva = rstClientes!Iva
                Rem WCodIb = rstCliente!Ib
                WRazon = rstClientes!Razon
                WDireccion = rstClientes!Direccion
                WLocalidad = rstClientes!Localidad
                WPostal = rstClientes!Postal
                WCuit = rstClientes!Cuit
                WDirentrega = rstClientes!DirEntrega
                rstClientes.Close
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
            End If
            Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                        Call Calcula_Click
                        DBGrid1.Row = WRow

                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub

Private Sub Form_Load()

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 1, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 1
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 6000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i
 
    Rem DBGrid1.FirstRow = 0
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Row = 0
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cae.Text = ""
    Dolar1.Text = ""
    Dolar2.Text = ""
    
    WLinea.Text = ""
    WDescripcion.Text = ""
    WImporte.Text = ""
    Renglon = 0
    
    Factura.Value = True
    Debito.Value = False
    Credito.Value = False

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
     
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
     
    Numero.SetFocus
    
End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
            
        WAnterior = DBGrid1.Row
                
        DBGrid1.Col = 0
        DBGrid1.Text = WDescripcion.Text
            
        If Val(WImporte.Text) <> 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
                Else
            DBGrid1.Col = 1
            DBGrid1.Text = ""
        End If
            
        DBGrid1.Row = Renglon
        DBGrid1.Col = 0
            
            Else
                
        DBGrid1.Row = Val(WLinea.Text) - 1
            
        WAnterior = DBGrid1.Row
                
        DBGrid1.Col = 0
        DBGrid1.Text = WDescripcion.Text
            
        If Val(WImporte.Text) <> 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
                Else
            DBGrid1.Col = 1
            DBGrid1.Text = ""
        End If
            
        DBGrid1.Row = Renglon
        DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 1
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    
    If Factura.Value = True Then
        WTipo = "03"
    End If
    If Debito.Value = True Then
        WTipo = "04"
    End If
    If Credito.Value = True Then
        WTipo = "05"
    End If
    
    XParam = "'" + WTipo + "','" _
                + Numero.Text + "'"
    
    spDesccomp = "ConsultaDesccomp1 " + XParam
    Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
    If rstDesccomp.RecordCount > 0 Then
    
        With rstDesccomp
            .MoveFirst
            Do
                If .EOF = False Then
                
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                
                DBGrid1.Col = 0
                DBGrid1.Text = !Descripcion
                
                If !Importe <> 0 Then
                    DBGrid1.Col = 1
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Importe))
                        Else
                    DBGrid1.Col = 1
                    DBGrid1.Text = ""
                End If
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDesccomp.Close
    End If
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Call Calcula_Click
    
    Graba.Enabled = False
    Borra.Enabled = False
    Ingresa.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Factura.Value = True Then
            WTipo = "03"
        End If
        If Debito.Value = True Then
            WTipo = "04"
        End If
        If Credito.Value = True Then
            WTipo = "05"
        End If
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = WTipo + Auxi + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
                
                Fecha.Text = rstCtacte!Fecha
                Cliente.Text = rstCtacte!Cliente
                Vencimiento.Text = rstCtacte!Vencimiento
                Paridad.Text = rstCtacte!Paridad
                Cae.Text = IIf(IsNull(rstCtacte!Cae), "", rstCtacte!Cae)
                Dolar1.Text = IIf(IsNull(rstCtacte!ImpreDolar1), "", rstCtacte!ImpreDolar1)
                Dolar2.Text = IIf(IsNull(rstCtacte!ImpreDolar2), "", rstCtacte!ImpreDolar2)
                
                rstCtacte.Close
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    WPago1 = 1
                    WPago2 = 1
                    WVendedor = rstCliente!vendedor
                    WProvincia = rstCliente!Provincia
                    WRubro = rstCliente!Rubro
                    WCodIva = rstCliente!Iva
                    WCodIb = rstCliente!Ib
                    WRazon = rstCliente!Razon
                    WDireccion = rstCliente!Direccion
                    WLocalidad = rstCliente!Localidad
                    WPostal = rstCliente!Postal
                    WCuit = rstCliente!Cuit
                    WDirentrega = rstCliente!DirEntrega
                End If
                Call Proceso_Click
                    Else
                Rem .Index = "Numero"
                Rem .Seek "=", Val(Numero.Text)
                Rem If .NoMatch = False Then
                Rem     m$ = "Comprobante ya existente"
                Rem   A% = MsgBox(m$, 0, "Ingreso de comprobantes varias")
                Rem     Numero.SetFocus
                Rem        Else
                Rem    Graba.Enabled = True
                Rem    Borra.Enabled = True
                Rem    Ingresa.Enabled = True
                Rem    WNumero = Numero.Text
                Rem    Numero.Text = WNumero
                Rem    Fecha.SetFocus
                Rem End If
                Graba.Enabled = True
                Borra.Enabled = True
                Ingresa.Enabled = True
                WNumero = Numero.Text
                Numero.Text = WNumero
                Fecha.SetFocus
                
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Cliente.Text = rstCliente!Cliente
            DesCliente.Caption = rstCliente!Razon
            WPago1 = 1
            WPago2 = 1
            WVendedor = rstCliente!vendedor
            WProvincia = rstCliente!Provincia
            WRubro = rstCliente!Rubro
            WCodIva = rstCliente!Iva
            WCodIb = rstCliente!Ib
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            WDirentrega = rstCliente!DirEntrega
            rstCliente.Close
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Cliente.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de Comprobante varios")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Remito.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstCliente.Close
    End If
    End If

End Sub


Sub Impresion()
    
    Call Calcula_Barra
    If Factura.Value = True Then
        WTipo = "03"
    End If
    If Debito.Value = True Then
        WTipo = "04"
    End If
    If Credito.Value = True Then
        WTipo = "05"
    End If
        
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = WTipo + Auxi + "01"
        
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "ImpreTotalNeto= " + "'" + Total.Caption + "'"
    ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
    Listado.WindowTitle = "Factura Electronica"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Factura.Value = True Then
        ZZTipo = "03"
    End If
    If Debito.Value = True Then
        ZZTipo = "04"
    End If
    If Credito.Value = True Then
        ZZTipo = "05"
    End If

    Auxi1 = Trim(Str$(Val(Numero.Text)))
    Auxi2 = ZZTipo
    Call Ceros(Auxi2, 2)

    Uno = "{Desccomp.Numero} in " + Chr$(34) + Auxi1 + Chr$(34) + " to " + Chr$(34) + Auxi1 + Chr$(34)
    Dos = " and {Desccomp.Tipo} in " + Chr$(34) + Auxi2 + Chr$(34) + " to " + Chr$(34) + Auxi2 + Chr$(34)

    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Destination = 1
    
    Select Case Val(ZZTipo)
        Case 5
            Select Case Val(WEmpresa)
                Case 1
                    Listado.ReportFileName = "ImpreNotaExpoVarios.rpt"
                Case Else
                    Listado.ReportFileName = "ImpreNotaExpoVariosPelli.rpt"
            End Select
        Case Else
            Select Case Val(WEmpresa)
                Case 1
                    Listado.ReportFileName = "ImpreFacturaExpoVarios.rpt"
                Case Else
                    Listado.ReportFileName = "ImpreFacturaExpoVariosPelli.rpt"
            End Select
    End Select
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT DescComp.Tipo, DescComp.Numero, DescComp.Renglon, DescComp.Descripcion, DescComp.Importe, " _
            + "Cliente.Razon, Cliente.Direccion, Cliente.Localidad, " _
            + "CtaCte.fecha, CtaCte.TotalUs, CtaCte.Seguro, CtaCte.Flete, CtaCte.ImpreNumero, CtaCte.Cae, CtaCte.FechaCae, CtaCte.Marca, CtaCte.Envio1, CtaCte.Envio2, CtaCte.Pago1, CtaCte.Pago2, CtaCte.NroOrden, CtaCte.FecOrden, CtaCte.Consignatario, CtaCte.Cip, CtaCte.ImpreDolar1, CtaCte.ImpreDolar2, CtaCte.ImpreTotal, CtaCte.ImpreTotalBruto, CtaCte.ImpreTotalNeto, CtaCte.Gastos, CtaCte.ImpreBarra, CtaCte.ImpreBarraII " _
            + "From " _
            + DSQ + ".dbo.DescComp DescComp, " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.CtaCte CtaCte " _
            + "Where " _
            + "DescComp.Cliente = Cliente.Cliente AND " _
            + "DescComp.ClaveCtaCte = CtaCte.Clave AND " _
            + "DescComp.Numero >= '" + Auxi1 + "' AND " _
            + "DescComp.Numero <= '" + Auxi1 + "' AND " _
            + "DescComp.Tipo >= '" + Auxi2 + "' AND " _
            + "DescComp.Tipo <= '" + Auxi2 + "'"
    
    Listado.Connect = Connect()
    Listado.CopiesToPrinter = 2
    
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.Action = 1

End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = "dolares"
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    WTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    WTexto1 = WTexto1 + Space$(50)
    
    Pasa = 0
    
    For DA = 40 To 1 Step -1
        If Mid$(WTexto1, DA, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(WTexto1, DA, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next DA
    
    WTexto2 = Mid$(WTexto1, DA + 2, 35)
    WTexto1 = Left$(WTexto1, DA)
    
End Sub




Private Sub Calcula_Cae()
    
    Dim WSAA As Object, WSFEXv1 As Object
    Dim dst_cmp  As Integer
    
    
    
    On Error GoTo ManejoError
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEXv1
    tra = WSAA.CreateTRA("WSFEXv1")
    Debug.Print tra
    
    
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    ZPath = "c:\salva\"
    
    Select Case Val(WEmpresa)
        Case 1
            ZNombre = "surfa"
            ZCuit = "30549165083"
        Case Else
            ZNombre = "pellital"
            ZCuit = "30610524598"
    End Select
    
    

    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    
    
    ' Llamar al web service para autenticar:
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Producción



    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica de Exportación
    Set WSFEXv1 = CreateObject("WSFEXv1")
    
    
    
    ' Setear tocken y sing de autorización (pasos previos)
    WSFEXv1.Token = WSAA.Token
    WSFEXv1.Sign = WSAA.Sign
    
    
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEXv1.Cuit = ZCuit
    
    
    
    ' Conectar al Servicio Web de Facturación
    ok = WSFEXv1.Conectar("https://servicios1.afip.gov.ar/WSFEXv1/service.asmx") ' homologación
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEXv1.Dummy
    Debug.Print "appserver status", WSFEXv1.AppServerStatus
    Debug.Print "dbserver status", WSFEXv1.DbServerStatus
    Debug.Print "authserver status", WSFEXv1.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    If Factura.Value = True Then
        tipo_cbte = 19 ' FC Expo (ver tabla de parámetros)
    End If
    If Debito.Value = True Then
        tipo_cbte = 20 ' FC Expo (ver tabla de parámetros)
    End If
    If Credito.Value = True Then
        tipo_cbte = 21 ' FC Expo (ver tabla de parámetros)
    End If
    Select Case Val(WEmpresa)
        Case 1
            punto_vta = 6
        Case Else
            punto_vta = 3
    End Select
    
    
    ' Obtengo el último número de comprobante y le agrego 1
    
    Debug.Print WSFEXv1.XmlRequest
    Debug.Print WSFEXv1.XmlResponse
    
    
    Cbte_Nro = WSFEXv1.GetLastCMP(tipo_cbte, punto_vta) + 1 '16
    ZZComprobante = Cbte_Nro
    
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZCuit = rstCliente!Cuit
        ZZPais = Trim(IIf(IsNull(rstCliente!Pais), "0", rstCliente!Pais))
        ZZCuitII = Trim(IIf(IsNull(rstCliente!CuitII), "", rstCliente!CuitII))
        rstCliente.Close
    End If
    
    
    
    fecha_cbte = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    tipo_expo = 2 ' tipo de exportación (ver tabla de parámetros)
    permiso_existente = ""
    dst_cmp = Val(ZZPais)
    XXCliente = WRazon
    cuit_pais_cliente = ZZCuit
    domicilio_cliente = WDireccion
    id_impositivo = ZZCuitII
    Rem ZZCuitII
    moneda_id = "DOL" ' para reales, "DOL" o "PES" (ver tabla de parámetros)
    Rem moneda_ctz = 0.5   PARIDAD
    moneda_ctz = Val(Paridad.Text)
    obs_comerciales = "..."
    obs = "..."
    forma_pago = ""
    incoterms = "FOB"  ' (ver tabla de parámetros)
    incoterms_ds = ""
    idioma_cbte = 1  ' (ver tabla de parámetros)
    IMP_TOTAL = Total.Caption
   
    ' Creo una factura (internamente, no se llama al WebService):
    ok = WSFEXv1.CrearFactura(tipo_cbte, punto_vta, Cbte_Nro, fecha_cbte, _
            IMP_TOTAL, tipo_expo, permiso_existente, dst_cmp, _
            XXCliente, cuit_pais_cliente, domicilio_cliente, _
            id_impositivo, moneda_id, moneda_ctz, _
            obs_comerciales, obs, forma_pago, incoterms, _
            idioma_cbte, incoterms_ds)
    
    
    
    
    
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            WDescripcion = DBGrid1.Text
                    
            DBGrid1.Col = 1
            WImporte = DBGrid1.Text
                    
            If Val(WImporte) <> 0 Then
                    
                XXCodigo = ""
                XXDs = WDescripcion
                qty = "0"
                XXPrecio = "0"
                umed = 0 ' Ver tabla de parámetros (unidades de medida)
                IMP_TOTAL = WImporte ' importe total final del artículo
                Bonif = ""
                
                ' lo agrego a la factura (internamente, no se llama al WebService):
                ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL, Bonif)
                        
            End If
                                        
        Next iRow
            
    Next a
    
    
    
    
    ' Agrego un permiso (ver manual para el desarrollador)
    Rem id = "99999AAXX999999A"
    Rem dst = Val(ZZPais)
    Rem ok = WSFEXv1.AgregarPermiso(id, dst)
        
        
        
        
    ' Agrego un comprobante asociado (ver manual para el desarrollador)
    Rem tipo_cbte_asoc = 19
    Rem punto_vta_asoc = 2
    Rem cbte_nro_asoc = 1
    Rem ok = WSFEXv1.AgregarCmpAsoc(tipo_cbte_asoc, punto_vta_asoc, cbte_nro_asoc)
        
        
        
    'id = "99000000000100" ' número propio de transacción
    ' obtengo el último ID y le adiciono 1 (advertencia: evitar overflow!)
    id = CStr(CCur(WSFEXv1.GetLastID()) + 1)
    
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Cae = WSFEXv1.Authorize(id)
    Debug.Print WSFEXv1.XmlRequest
    Debug.Print WSFEXv1.XmlResponse
    Cae.Text = Cae
        
        
        
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Or WSFEXv1.Resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    ElseIf WSFEXv1.obs <> "" And WSFEXv1.obs <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    End If
    
    
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEXv1.XmlRequest
    Debug.Print WSFEXv1.XmlResponse
    
    MsgBox "Resultado:" & WSFEXv1.Resultado & " CAE: " & Cae & " Reproceso: " & WSFEXv1.Reproceso & " Obs: " & WSFEXv1.obs & " Nro: " & ZZComprobante, vbInformation + vbOKOnly
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEXv1.Eventos
        If evento <> "0: " Then
            MsgBox "Evento: " & evento, vbInformation
        End If
    Next
    
    ' Buscar la factura
    cae2 = WSFEXv1.GetCMP(tipo_cbte, punto_vta, Cbte_Nro)
    
    Debug.Print "Fecha Comprobante:", WSFEXv1.FechaCbte
    Debug.Print "Importe Total:", WSFEXv1.ImpTotal
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!"
            Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
        ZZGrabaFactura = "S"
    End If
    
    
    Exit Sub
    
ManejoError:
    ' Si hubo error:
    Debug.Print WSFEXv1.XmlRequest
    Debug.Print WSFEXv1.XmlResponse
    
    
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEXv1.XmlRequest
    Debug.Assert False

End Sub





Private Sub Calcula_Barra()
    
    Dim ZZCara(1000) As String
    
    ZZNumero = ""
    Select Case Val(WEmpresa)
        Case 1
            ZZNumero = "30549165083"
        Case Else
            ZZNumero = "30610524598"
    End Select
    
    ZZNumero = ZZNumero + "19"
    
    Select Case Val(WEmpresa)
        Case 1
            ZZNumero = ZZNumero + "0006"
        Case Else
            ZZNumero = ZZNumero + "0003"
    End Select
    
    ZZNumero = ZZNumero + Trim(Cae.Text)
    
    ZZFechaCae = DateAdd("d", 10, Fecha.Text)
    ZZOrdFechaCae = Right$(ZZFechaCae, 4) + Mid$(ZZFechaCae, 4, 2) + Left$(ZZFechaCae, 2)
    ZZNumero = ZZNumero + ZZOrdFechaCae
    
    ZZCara(0) = "!"
    ZZCara(1) = Chr$(34)
    ZZCara(2) = "#"
    ZZCara(3) = "$"
    ZZCara(4) = "%"
    ZZCara(5) = "&"
    ZZCara(6) = "?"
    ZZCara(7) = "("
    ZZCara(8) = ")"
    ZZCara(9) = "*"
    ZZCara(10) = "+"
    ZZCara(11) = ","
    ZZCara(12) = "-"
    ZZCara(13) = "."
    ZZCara(14) = "/"
    ZZCara(15) = "0"
    ZZCara(16) = "1"
    ZZCara(17) = "2"
    ZZCara(18) = "3"
    ZZCara(19) = "4"
    ZZCara(20) = "5"
    ZZCara(21) = "6"
    ZZCara(22) = "7"
    ZZCara(23) = "8"
    ZZCara(24) = "9"
    ZZCara(25) = ":"
    ZZCara(26) = ";"
    ZZCara(27) = "<"
    ZZCara(28) = "="
    ZZCara(29) = ">"
    ZZCara(30) = "?"
    ZZCara(31) = "@"
    ZZCara(32) = "A"
    ZZCara(33) = "B"
    ZZCara(34) = "C"
    ZZCara(35) = "D"
    ZZCara(36) = "E"
    ZZCara(37) = "F"
    ZZCara(38) = "G"
    ZZCara(39) = "H"
    ZZCara(40) = "I"
    ZZCara(41) = "J"
    ZZCara(42) = "K"
    ZZCara(43) = "L"
    ZZCara(44) = "M"
    ZZCara(45) = "N"
    ZZCara(46) = "O"
    ZZCara(47) = "P"
    ZZCara(48) = "Q"
    ZZCara(49) = "R"
    ZZCara(50) = "S"
    ZZCara(51) = "T"
    ZZCara(52) = "U"
    ZZCara(53) = "V"
    ZZCara(54) = "W"
    ZZCara(55) = "X"
    ZZCara(56) = "Y"
    ZZCara(57) = "Z"
    ZZCara(58) = "["
    ZZCara(59) = "\"
    ZZCara(60) = "]"
    ZZCara(61) = "^"
    ZZCara(62) = "_"
    ZZCara(63) = "`"
    ZZCara(64) = "a"
    ZZCara(65) = "b"
    ZZCara(66) = "c"
    ZZCara(67) = "d"
    ZZCara(68) = "e"
    ZZCara(69) = "f"
    ZZCara(70) = "g"
    ZZCara(71) = "h"
    ZZCara(72) = "i"
    ZZCara(73) = "j"
    ZZCara(74) = "k"
    ZZCara(75) = "l"
    ZZCara(76) = "m"
    ZZCara(77) = "n"
    ZZCara(78) = "o"
    ZZCara(79) = "p"
    ZZCara(80) = "q"
    ZZCara(81) = "r"
    ZZCara(82) = "s"
    ZZCara(83) = "t"
    ZZCara(84) = "u"
    ZZCara(85) = "v"
    ZZCara(86) = "w"
    ZZCara(87) = "x"
    ZZCara(88) = "y"
    ZZCara(89) = "z"
    ZZCara(90) = "¡"
    ZZCara(91) = "¢"
    ZZCara(92) = "£"
    ZZCara(93) = "¤"
    ZZCara(94) = "¥"
    ZZCara(95) = "¦"
    ZZCara(96) = "§"
    ZZCara(97) = "¨"
    ZZCara(98) = "©"
    ZZCara(99) = "ª"
    
    Rem ZZNumero = "3070306062119000260321213344273201008198"
    Rem ZZNumero = "000102030405060708091011121314151617181920"
    Rem ZZNumero = "2122232425262728293031323334353637383940"
    Rem ZZNumero = "4142434445464748495051525354555657585960"
    Rem ZZNumero = "6162636465666768697071727374757677787980"
    Rem ZZNumero = "81828384858687888990919293949596979899"
    Rem ZZNumero = "307030606211900026032121334427320100819"
    
    ZZSumaI = 0
    ZZSumaII = 0
    
    For Ciclo = 1 To 39 Step 2
        ZZSumaI = ZZSumaI + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    ZZSumaI = ZZSumaI * 3
    
    For Ciclo = 2 To 39 Step 2
        ZZSumaII = ZZSumaII + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    
    ZZSuma = ZZSumaI + ZZSumaII
    ZZVerifica = ZZSuma
    ZZDigi = 0
    
    Do
    
        ZZVerifi = Int(ZZVerifica / 10) * 10
        
        If ZZVerifi = ZZVerifica Then
            Exit Do
        End If
        
        ZZDigi = ZZDigi + 1
        
        ZZVerifica = ZZSuma + ZZDigi
        
    Loop
    
    ZZNumero = ZZNumero + Trim(Str$(ZZDigi))
    
    lccar = ""
    barralargo = ZZNumero
    
    For lni = 1 To Len(barralargo) Step 2
        ZZLugar = Val(Mid(barralargo, lni, 2))
        lccar = lccar + ZZCara(ZZLugar)
    Next
    
    Rem barralargo = "{" + lccar + "}"
    barralargo = "(" + lccar + ")"
    
    
    If Factura.Value = True Then
        ZZTipo = "03"
    End If
    If Debito.Value = True Then
        ZZTipo = "04"
    End If
    If Credito.Value = True Then
        ZZTipo = "05"
    End If
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    ZZImpreNumero = "0000" + Right$(Auxi, 4)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql & "FechaCae = " + "'" + ZZFechaCae + "',"
    ZSql = ZSql & "ImpreBarra = " + "'" + barralargo + "',"
    ZSql = ZSql & "ImpreBarraII = " + "'" + ZZNumero + "'"
    ZSql = ZSql & " Where Tipo = " + "'" + ZZTipo + "'"
    ZSql = ZSql & " and Numero = " + "'" + Numero.Text + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

End Sub



Private Sub Credito_Click()
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
End Sub

Private Sub Debito_Click()
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
End Sub


Private Sub Factura_Click()
    If Factura.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Debito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "08" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
    If Credito.Value = True Then
        spNumero = "ConsultaNumero " + "'" + "09" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            Numero.Text = rstNumero!Numero + 1
            rstNumero.Close
                Else
            Numero.Text = ""
        End If
    End If
End Sub

