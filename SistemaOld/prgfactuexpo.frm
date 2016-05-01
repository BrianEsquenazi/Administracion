VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactuexpo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos"
   ClientHeight    =   8310
   ClientLeft      =   225
   ClientTop       =   285
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.CommandButton ImpreCertificado 
      Caption         =   "Certificado Analisis"
      Height          =   615
      Left            =   10200
      TabIndex        =   71
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton REImpreIII 
      Caption         =   "ReImpresion Varios"
      Height          =   615
      Left            =   8880
      TabIndex        =   70
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox IdiomaII 
      Height          =   315
      Left            =   6480
      TabIndex        =   69
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox TipoPedido 
      Height          =   315
      Left            =   6360
      TabIndex        =   66
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton ReImpresionII 
      Caption         =   "ReImpresion Factura"
      Height          =   615
      Left            =   10200
      TabIndex        =   65
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton DatosAdicionales 
      Caption         =   "Datos Adicionales"
      Height          =   450
      Left            =   9120
      TabIndex        =   60
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Cae 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   57
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame CargaAdicional 
      Height          =   4815
      Left            =   480
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   6495
      Begin VB.ComboBox Idioma 
         Height          =   315
         Left            =   1080
         TabIndex        =   62
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton AceptaAdicional 
         Caption         =   "Confirma Datos"
         Height          =   570
         Left            =   3000
         TabIndex        =   61
         Top             =   3960
         Width           =   1215
      End
      Begin VB.ComboBox CipLista 
         Height          =   315
         Left            =   5040
         TabIndex        =   59
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Envio1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox Envio2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox Pago1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox Pago2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox NroOrden 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   44
         Text            =   " "
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Consignatario 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   42
         Text            =   " "
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Marca 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   41
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox Dolar2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   40
         Text            =   " "
         Top             =   3240
         Width           =   5055
      End
      Begin VB.TextBox Dolar1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   39
         Text            =   " "
         Top             =   2880
         Width           =   5055
      End
      Begin MSMask.MaskEdBox fecorden 
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label24 
         Caption         =   "Idioma"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label rrr 
         Caption         =   "Marca"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Envio"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Nro orden"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Orden"
         Height          =   375
         Left            =   3000
         TabIndex        =   52
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Consignatario"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Incoterms"
         Height          =   375
         Left            =   4080
         TabIndex        =   50
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Dolar"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.CommandButton ReImpresion 
      Caption         =   "ReImpresion Remito"
      Height          =   615
      Left            =   10200
      TabIndex        =   34
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton ConsultaPedido 
      Caption         =   "Consulta Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   33
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Graba1 
      Caption         =   "Fc. Exportacion"
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   29
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      Height          =   495
      Left            =   9120
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   8160
      TabIndex        =   21
      Top             =   3840
      Width           =   2535
      Begin VB.TextBox Descuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   67
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Gastos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   64
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Flete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   37
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Seguro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   36
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Flete"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Seguro"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Gastos"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Fob"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   20
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   8040
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   2280
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Remito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   14
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
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
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   5
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   10200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      ItemData        =   "prgfactuexpo.frx":0000
      Left            =   360
      List            =   "prgfactuexpo.frx":0007
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   6375
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreRemito.rpt"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "prgfactuexpo.frx":0015
      TabIndex        =   35
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label23 
      Caption         =   "Cae"
      Height          =   375
      Left            =   3360
      TabIndex        =   58
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Orden de compra"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Remito"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgFactuexpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
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
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private Dada As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WImpoIb As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private parcial As String
Private WSeguro As Double
Private WFlete As Double
Private WGastos As Double
Private WDescuento As Double
Private WTexto1 As String
Private WTexto2 As String
Private Auxiliar(100, 50) As String
Private AuxiliarII(100, 50) As String

Dim ZZControlLote(100, 60) As String
Dim ControlLote(12, 2) As String
Dim ControlEnvase(12, 2) As String

Dim ZZZZProceso As Integer
Dim ListaSga(1000) As String

Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstImpreRemito As Recordset
Dim spImpreRemio As String

Dim XParam As String
Dim WLote(12, 2) As String
Dim WImpresion(100, 10) As String
Dim XEnvase(100, 6) As String
Dim XCanti As String
Private WTipoPedido As String

Dim VectorCosto(100, 3) As String
Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZZClave As String
Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZFecha As String
Dim ZZNombre As String
Dim ZZDireccion As String
Dim ZZLocalidad As String
Dim ZZPedido As String
Dim ZZCliente As String
Dim ZZOrden As String
Dim ZZDescripcion As String
Dim ZZCantidad As String
Dim ZZRemito As String
Dim ZZDestino As Integer

Dim ZZVector(100, 10) As String
Dim ZZImpre(100, 10) As String
Dim ZZCampo1 As String
Dim ZZCampo2 As String
Dim ZLote6 As Double
Dim ZLote7 As Double
Dim ZLote8 As Double
Dim ZLote9 As Double
Dim ZLote10 As Double
Dim ZLote11 As Double
Dim ZLote12 As Double

Dim ZZComprobante As Integer
Dim ZZCuit As String
Dim ZZPais As String
Dim ZZCuitII As String
Dim ZZRazon As String
Dim ZZDomicilio As String
Dim ZZFechaCae As String

Dim ZZGrabaFactura As String

Dim ZImpreFicha(100) As String
Dim CargaEmpresa(12, 2) As String
Dim ZZLote As String
Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String
Dim ZZIntervencion As String
Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 6) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String


Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        Rem WDescuento = rstPago!Descuento
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

Private Sub AceptaAdicional_Click()
    CargaAdicional.Visible = False
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For a = 0 To 8
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            If Val(Cantidad) <> 0 Then
                WNeto = WNeto + (Val(Cantidad) * Val(Precio))
            End If
                    
        Next iRow
            
    Next a
    
    WNeto = 0
    
    For a = 0 To 8
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            If Val(Cantidad) <> 0 Then
                WNeto = WNeto + (Val(Cantidad) * Val(Precio))
            End If
                    
        Next iRow
            
    Next a
    
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0

    Rem WDescueneto = Descuento.te
    Rem If Val(Paridad.Text) <> 0 Then
    Rem     WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    
    Rem If WDescuento <> 0 Then
    Rem     Rem WImpoDto = WNeto * WDescuento / 100
    Rem     WImpoDto = WDescuento
    Rem     Call Redondeo(WImpoDto)
    Rem     WNeto = WNeto - WImpoDto
    Rem End If
    Rem
    Rem If WTasa <> 0 Then
    Rem     WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
    Rem     Call Redondeo(WImpoInteres)
    Rem     WNeto = WNeto + WImpoInteres
    Rem End If
    
    WIva1 = 0
    WIva2 = 0
    WImpoIb = 0
    
    Rem Select Case Val(WCodIva)
    Rem     Case 2
    Rem         WIva1 = WNeto * 0.21
    Rem         WIva2 = WNeto * 0.105
    Rem         Call Redondeo(WIva1)
    Rem         Call Redondeo(WIva2)
    Rem     Case 4
    Rem         WIva1 = 0
    Rem         WIva2 = 0
    Rem     Case Else
    Rem         WIva1 = WNeto * 0.21
    Rem         Call Redondeo(WIva1)
    Rem End Select
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Rem Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Rem Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Rem Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Rem Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Iva2.Caption = "0.00"
    End If
        
    If Val(Wempresa) = 1 Then
    
        Rem end by nan
        WTotal = WNeto
        Rem WNeto = WNeto - Val(Seguro.Text) - Val(Flete.Text) - Val(Gastos.Text)
        Rem  + Val(Descuento.Text)
        Rem by nan
        
        Rem OJO
        WTotal = XNeto - Val(Descuento.Text)
        Rem fin by nan
        If WNeto <> 0 Then
            Call Convierte1_datos(Str$(WNeto), Auxi)
            Neto.Caption = Pusing("###,###.##", Auxi)
                Else
            Neto.Caption = "0.00"
        End If
            
            
        Call Convierte1_datos(Str$(WTotal), Auxi)
        Total.Caption = Pusing("###,###.##", Auxi)

            Else
        
        WTotal = WNeto + Val(Seguro.Text) + Val(Flete.Text) + Val(Gastos.Text) - Val(Descuento.Text)
        
        If WNeto <> 0 Then
            Call Convierte1_datos(Str$(WNeto), Auxi)
            Neto.Caption = Pusing("###,###.##", Auxi)
                Else
            Neto.Caption = "0.00"
        End If
            
        Call Convierte1_datos(Str$(WTotal), Auxi)
        Total.Caption = Pusing("###,###.##", Auxi)

    End If

End Sub

Private Sub cmdClose_Click()

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    
    PrgFactuexpo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()
    Call Impresion_Remito
End Sub

Private Sub Command6_Click()

End Sub

Private Sub ConsultaPedido_Click()
    ZZProcesoFactura = 3
    PrgSeleccionaPedido.Show
End Sub

Private Sub DatosAdicionales_Click()
    CargaAdicional.Visible = True
    Marca.SetFocus
End Sub

Private Sub DBGrid1_DblClick()
        
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        Select Case rstPedido!TipoPedido
            Case 1
                ZZPasaTipoPedido = "CO"
            Case 3
                ZZPasaTipoPedido = "BI"
            Case 4
                ZZPasaTipoPedido = "FA"
            Case 5
                ZZPasaTipoPedido = "PG"
            Case Else
                ZZPasaTipoPedido = "PT"
        End Select
        rstPedido.Close
    End If
    
    DBGrid1.Col = 5
    ZZPasaClave = DBGrid1.Text
    DBGrid1.Col = 0
    ZZPasaTerminado = DBGrid1.Text
    DBGrid1.Col = 4
    ZZPasaCantidad = Val(DBGrid1.Text)
    ZSuma = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Clave = " + "'" + ZZPasaClave + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        XLote = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
        ZSuma = ZSuma + ZLote
        XLote = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
        ZSuma = ZSuma + ZLote
        XLote = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
        ZSuma = ZSuma + ZLote
        XLote = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
        ZSuma = ZSuma + ZLote
        XLote = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
        ZSuma = ZSuma + ZLote
                
        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
        
        If Len(Trim(WLoteAdicional)) = 98 Then
            XLote = Val(Mid$(WLoteAdicional, 9, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 23, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 37, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 51, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 65, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 79, 6))
            ZSuma = ZSuma + ZLote
            XLote = Val(Mid$(WLoteAdicional, 93, 6))
            ZSuma = ZSuma + ZLote
        End If
        
        rstEstadistica.Close
        
    End If
    
    If ZSuma <> 0 Then
        Exit Sub
    End If
    
    PrgModFactuExpo.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    
    If ZZProcesoFactura = 99 And Val(Pedido.Text) <> 0 Then
        Call Pedido_KeyPress(13)
        Call Fecha_Keypress(13)
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        Remito.SetFocus
    End If
    
End Sub



Private Sub Graba_Click()

    If Trim(Marca.Text) = "" Then
        m$ = "No se a informado marca"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If Trim(Envio1.Text) = "" And Trim(Envio2.Text) = "" Then
        m$ = "No se a informado instrucciones de envio"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If Trim(Pago1.Text) = "" And Trim(Pago2.Text) = "" Then
        m$ = "No se a informado condiciones de pago"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If Trim(Dolar1.Text) = "" And Trim(Dolar2.Text) = "" Then
        m$ = "No se a informado el importe en dolares"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If CipLista.ListIndex < 1 Then
        m$ = "Codigo de incoterms incorrecto"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If Idioma.ListIndex < 1 Then
        m$ = "Codigo idioma incorrexto"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    If Val(Flete.Text) = 0 Then
        T$ = "Factura de Exportacion"
        Estilo = Estilo = vbYesNo + vbCritical + vbDefaultButton2
        m$ = "No se informo importe de flete, desea continuar "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% <> 6 Then
            Exit Sub
        End If
    End If
    
    If Val(Seguro.Text) = 0 Then
        T$ = "Factura de Exportacion"
        Estilo = Estilo = vbYesNo + vbCritical + vbDefaultButton2
        m$ = "No se informo importe de seguro, desea continuar "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% <> 6 Then
            Exit Sub
        End If
    End If
    
    If Val(Gastos.Text) = 0 Then
        T$ = "Factura de Exportacion"
        Estilo = Estilo = vbYesNo + vbCritical + vbDefaultButton2
        m$ = "No se informo importe de Gastos, desea continuar "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% <> 6 Then
            Exit Sub
        End If
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZCuit = rstCliente!Cuit
        ZZPais = Trim(IIf(IsNull(rstCliente!Pais), "0", rstCliente!Pais))
        ZZCuitII = Trim(IIf(IsNull(rstCliente!CuitII), "", rstCliente!CuitII))
        rstCliente.Close
    End If
    
    If Trim(ZZCuit) = "" Then
        m$ = "No se a informado el numero de cuit"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End If
    
    Rem If Trim(ZZCuitII) = "" Then
    Rem     m$ = "No se a informado el identificacion tributario"
    Rem     G% = MsgBox(m$, 0, "Factura de Exportacion")
    Rem     Rem Exit Sub
    Rem End If
    
    
    
    
    Erase ZZVector
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    ZZImpreTotal = 0
    ZZImpreTotalBruto = 0
    ZZImpreTotalNeto = 0
    
    ZZDesde = 1
    
    For a = 0 To 8
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
    
        For iRow = 0 To 9
        
            WRenglon = WRenglon + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            If Trim(Articulo) <> "" Then
                ZZVector(WRenglon, 1) = Articulo
                ZZVector(WRenglon, 2) = Str$(Cantidad)
                ZZVector(WRenglon, 3) = Str$(Precio)
            End If
            
            ControlLote(1, 1) = ZZControlLote(WRenglon, 1)
            ControlLote(1, 2) = ZZControlLote(WRenglon, 2)
            ControlLote(2, 1) = ZZControlLote(WRenglon, 3)
            ControlLote(2, 2) = ZZControlLote(WRenglon, 4)
            ControlLote(3, 1) = ZZControlLote(WRenglon, 5)
            ControlLote(3, 2) = ZZControlLote(WRenglon, 6)
            ControlLote(4, 1) = ZZControlLote(WRenglon, 7)
            ControlLote(4, 2) = ZZControlLote(WRenglon, 8)
            ControlLote(5, 1) = ZZControlLote(WRenglon, 9)
            ControlLote(5, 2) = ZZControlLote(WRenglon, 10)
            ControlLote(6, 1) = ZZControlLote(WRenglon, 11)
            ControlLote(6, 2) = ZZControlLote(WRenglon, 12)
            ControlLote(7, 1) = ZZControlLote(WRenglon, 13)
            ControlLote(7, 2) = ZZControlLote(WRenglon, 14)
            ControlLote(8, 1) = ZZControlLote(WRenglon, 15)
            ControlLote(8, 2) = ZZControlLote(WRenglon, 16)
            ControlLote(9, 1) = ZZControlLote(WRenglon, 17)
            ControlLote(9, 2) = ZZControlLote(WRenglon, 18)
            ControlLote(10, 1) = ZZControlLote(WRenglon, 19)
            ControlLote(10, 2) = ZZControlLote(WRenglon, 20)
            ControlLote(11, 1) = ZZControlLote(WRenglon, 21)
            ControlLote(11, 2) = ZZControlLote(WRenglon, 22)
            ControlLote(12, 1) = ZZControlLote(WRenglon, 23)
            ControlLote(12, 2) = ZZControlLote(WRenglon, 24)
            
            ControlEnvase(1, 1) = ZZControlLote(WRenglon, 31)
            ControlEnvase(1, 2) = ZZControlLote(WRenglon, 32)
            ControlEnvase(2, 1) = ZZControlLote(WRenglon, 33)
            ControlEnvase(2, 2) = ZZControlLote(WRenglon, 34)
            ControlEnvase(3, 1) = ZZControlLote(WRenglon, 35)
            ControlEnvase(3, 2) = ZZControlLote(WRenglon, 36)
            ControlEnvase(4, 1) = ZZControlLote(WRenglon, 37)
            ControlEnvase(4, 2) = ZZControlLote(WRenglon, 38)
            ControlEnvase(5, 1) = ZZControlLote(WRenglon, 39)
            ControlEnvase(5, 2) = ZZControlLote(WRenglon, 40)
            ControlEnvase(6, 1) = ZZControlLote(WRenglon, 41)
            ControlEnvase(6, 2) = ZZControlLote(WRenglon, 42)
            ControlEnvase(7, 1) = ZZControlLote(WRenglon, 43)
            ControlEnvase(7, 2) = ZZControlLote(WRenglon, 44)
            ControlEnvase(8, 1) = ZZControlLote(WRenglon, 45)
            ControlEnvase(8, 2) = ZZControlLote(WRenglon, 46)
            ControlEnvase(9, 1) = ZZControlLote(WRenglon, 47)
            ControlEnvase(9, 2) = ZZControlLote(WRenglon, 48)
            ControlEnvase(10, 1) = ZZControlLote(WRenglon, 49)
            ControlEnvase(10, 2) = ZZControlLote(WRenglon, 50)
            ControlEnvase(11, 1) = ZZControlLote(WRenglon, 51)
            ControlEnvase(11, 2) = ZZControlLote(WRenglon, 52)
            ControlEnvase(12, 1) = ZZControlLote(WRenglon, 53)
            ControlEnvase(12, 2) = ZZControlLote(WRenglon, 54)
            
            SumaLote = 0
            SumaEnvase = 0
            SumaBruto = 0

            For Ciclo1 = 1 To 12
            
                If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                
                    SumaLote = SumaLote + Val(ControlLote(Ciclo1, 2))
                
                    ZZEnvase = ControlEnvase(Ciclo1, 1)
                    
                    If Val(ZZEnvase) <> 0 Then
                    
                        ZPeso = 0
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Envases"
                        ZSql = ZSql + " Where Envases = " + "'" + ZZEnvase + "'"
                        spEnvases = ZSql
                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnvases.RecordCount > 0 Then
                            ZPeso = IIf(IsNull(rstEnvases!Peso), "0", rstEnvases!Peso)
                            ZTipo = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
                            rstEnvases.Close
                        End If
                        
                        SumaEnvase = SumaEnvase + Val(ControlEnvase(Ciclo1, 2))
                        
                        ZZBruto = Val(ControlLote(Ciclo1, 2)) + (ZPeso * Val(ControlEnvase(Ciclo1, 2)))
                        SumaBruto = SumaBruto + ZZBruto
                        
                    End If
                    
                End If
                    
            Next Ciclo1
                        
            ZZImpre(WRenglon, 1) = SumaEnvase
            ZZImpre(WRenglon, 2) = ZTipo
            ZZImpre(WRenglon, 3) = Trim(Str$(ZZDesde)) + "/" + Trim(Str$(SumaEnvase))
            ZZImpre(WRenglon, 4) = Str$(SumaBruto)
            ZZImpre(WRenglon, 5) = Str$(SumaLote)
            
            ZZImpreTotal = ZZImpreTotal + SumaEnvase
            ZZImpreTotalBruto = ZZImpreTotalBruto + SumaBruto
            ZZImpreTotalNeto = ZZImpreTotalNeto + SumaLote
            
            ZZDesde = ZZDesde + SumaEnvase
            
            If TipoPedido.ListIndex = 0 Then
                If Cantidad <> SumaLote Then
                    m$ = Articulo + " Verifique la discriminacion de lotes"
                    G% = MsgBox(m$, 0, "Emision de facturas")
                    Exit Sub
                End If
            End If
            
        Next iRow
    Next a

    Call Calcula_Click
    
    Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
    Rem     WImporte = WNeto
    Rem     WNeto = WNeto / 1.21
    Rem     Call Redondeo(WNeto)
    Rem     WIva1 = WImporte - WNeto
    Rem     WIva2 = 0
    Rem End If
    
    Rem Call Graba_FE
    
    If Trim(Cae.Text) = "" Then
        ZZGrabaFactura = ""
        Call Calcula_Cae
        If ZZGrabaFactura <> "S" Then
            Exit Sub
        End If
    End If
    
    WTipo = "01"
    WNumero = Numero.Text
    WRenglon = "01"
    WCliente = Cliente.Text
    WFecha = Fecha.Text
    WEstado = "0"
    Rem Wvencimiento = Wvencimiento
    Rem WVencimiento1 = WVencimiento1
    Call Convierte_datos(Str$(Total), Auxi)
    XTotal = Str$(WTotal)
    XTotalUs = Str$(WTotal)
    XSaldo = Str$(WTotal)
    XSaldoUs = Str$(WTotal)
    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
    WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
    WImpre = "FC"
    XNet = Str$(WNeto * Val(Paridad.Text))
    XIva1 = Str$(WIva1 * Val(Paridad.Text))
    XIva2 = Str$(WIva2 * Val(Paridad.Text))
    XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
    XSeguro = Seguro.Text
    XFlete = Flete.Text
    XGastos = Gastos.Text
    XDescuento = Descuento.Text
    WPedido = Pedido.Text
    WRemito = Remito.Text
    WOrden = Orden.Text
    WParidad = Paridad.Text
    WProvincia = WProv
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
    WClave = "01" + Auxi + "01"
    XEmpresa = "1"
    WDate = Date$
    
    Call Numtolet
    
    WTexto1 = UCase(WTexto1)
    WTexto2 = UCase(WTexto2)
    
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    ZZImpreNumero = "0000" + Right$(Auxi, 4)
    ZZCae = Cae.Text
    ZZFechaCae = "  /  /    "
    ZZMarca = Marca.Text
    ZZEnvio1 = Envio1.Text
    ZZEnvio2 = Envio2.Text
    ZZPago1 = Pago1.Text
    ZZPago2 = Pago2.Text
    ZZNroOrden = NroOrden.Text
    ZZFecOrden = fecorden.Text
    ZZConsignatario = Consignatario.Text
    ZZCipLista = Str$(CipLista.ListIndex)
    ZZIdioma = Str$(Idioma.ListIndex)
    ZZCip = CipLista.Text
    ZZCip = ZZCip
    ZZImpreDolar1 = WTexto1
    ZZImpreDolar = WTexto2
    ZZImpreTotal = Str$(ZZImpreTotal)
    ZZImpreTotalBruto = Str$(ZZImpreTotalBruto)
    ZZImpreTotalNeto = Str$(ZZImpreTotalNeto)
    
    XParam = "'" + WClave + "','" _
                + WTipo + "','" + WNumero + "','" _
                + WRenglon + "','" + WCliente + "','" _
                + WFecha + "','" + WEstado + "','" _
                + Wvencimiento + "','" + WVencimiento1 + "','" _
                + XTotal + "','" + XTotalUs + "','" _
                + XSaldo + "','" + XSaldoUs + "','" _
                + WOrdFecha + "','" + WOrdVencimiento + "','" _
                + WOrdVencimiento1 + "','" + WImpre + "','" _
                + Wempresa + "','" _
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
                + XImpoIb + "'"
                    
    spCtacte = "AltaCtacte " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "Descuento = " + "'" + Descuento.Text + "',"
    ZSql = ZSql & "Gastos = " + "'" + Gastos.Text + "',"
    ZSql = ZSql & "ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql & "Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql & "FechaCae = " + "'" + "  /  /    " + "',"
    ZSql = ZSql & "Marca = " + "'" + Marca.Text + "',"
    ZSql = ZSql & "Envio1 = " + "'" + ZZEnvio1 + "',"
    ZSql = ZSql & "Envio2 = " + "'" + ZZEnvio2 + "',"
    ZSql = ZSql & "Pago1 = " + "'" + ZZPago1 + "',"
    ZSql = ZSql & "Pago2 = " + "'" + ZZPago2 + "',"
    ZSql = ZSql & "NroOrden = " + "'" + ZZNroOrden + "',"
    ZSql = ZSql & "FecOrden = " + "'" + ZZFecOrden + "',"
    ZSql = ZSql & "Consignatario = " + "'" + ZZConsignatario + "',"
    ZSql = ZSql & "Cip = " + "'" + ZZCip + "',"
    ZSql = ZSql & "CipLista = " + "'" + ZZCipLista + "',"
    ZSql = ZSql & "Idioma = " + "'" + ZZIdioma + "',"
    ZSql = ZSql & "ImpreDolar1 = " + "'" + Dolar1.Text + "',"
    ZSql = ZSql & "ImpreDolar2 = " + "'" + Dolar2.Text + "',"
    ZSql = ZSql & "ImpreTotal = " + "'" + ZZImpreTotal + "',"
    ZSql = ZSql & "ImpreTotalBruto = " + "'" + ZZImpreTotalBruto + "',"
    ZSql = ZSql & "ImpreTotalNeto = " + "'" + ZZImpreTotalNeto + "'"
    ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZClaveCtaCte = WClave
    
    Erase VectorCosto
    Erase Auxiliar
    Erase AuxiliarII
    
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    For a = 0 To 8
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
    
        For iRow = 0 To 9
        
            WRenglon = WRenglon + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
            
            DBGrid1.Col = 1
            ZZImpreTerminado = DBGrid1.Text
            
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            
            If Val(Wempresa) = 1 And IdiomaII.ListIndex = 1 Then
            
                XXXCodigo = Val(Mid$(Articulo, 4, 5))
                If XXXCodigo >= 25000 And XXXCodigo <= 25999 Then
                
                    If Left$(Articulo, 2) = "PT" Or Left$(Articulo, 2) = "PE" Then
                    
                        ZZDescripcionIngles = ""
                        ZZDescripcionInglesII = ""
                    
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZZDescripcionIngles = IIf(IsNull(rstTerminado!Descripcioningles), "", rstTerminado!Descripcioningles)
                            rstTerminado.Close
                        End If
                        
                        ClavePrecios = Cliente.Text + Articulo
                        spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPrecios.RecordCount > 0 Then
                            ZZDescripcionInglesII = rstPrecios!Descripcion
                            rstPrecios.Close
                        End If
                    
                        ZZImpreTerminado = ZZDescripcionIngles + " - " + ZZDescripcionInglesII
                        
                    End If
                    
                End If
            End If
            
            
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            
            ControlLote(1, 1) = ZZControlLote(WRenglon, 1)
            ControlLote(1, 2) = ZZControlLote(WRenglon, 2)
            ControlLote(2, 1) = ZZControlLote(WRenglon, 3)
            ControlLote(2, 2) = ZZControlLote(WRenglon, 4)
            ControlLote(3, 1) = ZZControlLote(WRenglon, 5)
            ControlLote(3, 2) = ZZControlLote(WRenglon, 6)
            ControlLote(4, 1) = ZZControlLote(WRenglon, 7)
            ControlLote(4, 2) = ZZControlLote(WRenglon, 8)
            ControlLote(5, 1) = ZZControlLote(WRenglon, 9)
            ControlLote(5, 2) = ZZControlLote(WRenglon, 10)
            ControlLote(6, 1) = ZZControlLote(WRenglon, 11)
            ControlLote(6, 2) = ZZControlLote(WRenglon, 12)
            ControlLote(7, 1) = ZZControlLote(WRenglon, 13)
            ControlLote(7, 2) = ZZControlLote(WRenglon, 14)
            ControlLote(8, 1) = ZZControlLote(WRenglon, 15)
            ControlLote(8, 2) = ZZControlLote(WRenglon, 16)
            ControlLote(9, 1) = ZZControlLote(WRenglon, 17)
            ControlLote(9, 2) = ZZControlLote(WRenglon, 18)
            ControlLote(10, 1) = ZZControlLote(WRenglon, 19)
            ControlLote(10, 2) = ZZControlLote(WRenglon, 20)
            ControlLote(11, 1) = ZZControlLote(WRenglon, 21)
            ControlLote(11, 2) = ZZControlLote(WRenglon, 22)
            ControlLote(12, 1) = ZZControlLote(WRenglon, 23)
            ControlLote(12, 2) = ZZControlLote(WRenglon, 24)
            
            ControlEnvase(1, 1) = ZZControlLote(WRenglon, 31)
            ControlEnvase(1, 2) = ZZControlLote(WRenglon, 32)
            ControlEnvase(2, 1) = ZZControlLote(WRenglon, 33)
            ControlEnvase(2, 2) = ZZControlLote(WRenglon, 34)
            ControlEnvase(3, 1) = ZZControlLote(WRenglon, 35)
            ControlEnvase(3, 2) = ZZControlLote(WRenglon, 36)
            ControlEnvase(4, 1) = ZZControlLote(WRenglon, 37)
            ControlEnvase(4, 2) = ZZControlLote(WRenglon, 38)
            ControlEnvase(5, 1) = ZZControlLote(WRenglon, 39)
            ControlEnvase(5, 2) = ZZControlLote(WRenglon, 40)
            ControlEnvase(6, 1) = ZZControlLote(WRenglon, 41)
            ControlEnvase(6, 2) = ZZControlLote(WRenglon, 42)
            ControlEnvase(7, 1) = ZZControlLote(WRenglon, 43)
            ControlEnvase(7, 2) = ZZControlLote(WRenglon, 44)
            ControlEnvase(8, 1) = ZZControlLote(WRenglon, 45)
            ControlEnvase(8, 2) = ZZControlLote(WRenglon, 46)
            ControlEnvase(9, 1) = ZZControlLote(WRenglon, 47)
            ControlEnvase(9, 2) = ZZControlLote(WRenglon, 48)
            ControlEnvase(10, 1) = ZZControlLote(WRenglon, 49)
            ControlEnvase(10, 2) = ZZControlLote(WRenglon, 50)
            ControlEnvase(11, 1) = ZZControlLote(WRenglon, 51)
            ControlEnvase(11, 2) = ZZControlLote(WRenglon, 52)
            ControlEnvase(12, 1) = ZZControlLote(WRenglon, 53)
            ControlEnvase(12, 2) = ZZControlLote(WRenglon, 54)
            
            lote1 = Val(ControlLote(1, 1))
            Canti1 = Val(ControlLote(1, 2))
            lote2 = Val(ControlLote(2, 1))
            Canti2 = Val(ControlLote(2, 2))
            lote3 = Val(ControlLote(3, 1))
            Canti3 = Val(ControlLote(3, 2))
            lote4 = Val(ControlLote(4, 1))
            Canti4 = Val(ControlLote(4, 2))
            lote5 = Val(ControlLote(5, 1))
            Canti5 = Val(ControlLote(5, 2))
            
            WLoteAdicional = ""
            For ZZCiclo = 6 To 12
                ZZCampo1 = ControlLote(ZZCiclo, 1)
                ZZCampo2 = ControlLote(ZZCiclo, 2)
                Call Ceros(ZZCampo1, 8)
                Call Ceros(ZZCampo2, 6)
                WLoteAdicional = WLoteAdicional + ZZCampo1 + ZZCampo2
            Next ZZCiclo
            
            Envase1 = Val(ControlEnvase(1, 1))
            CantiEnvase1 = Val(ControlEnvase(1, 2))
            Envase2 = Val(ControlEnvase(2, 1))
            CantiEnvase2 = Val(ControlEnvase(2, 2))
            Envase3 = Val(ControlEnvase(3, 1))
            CantiEnvase3 = Val(ControlEnvase(3, 2))
            Envase4 = Val(ControlEnvase(4, 1))
            CantiEnvase4 = Val(ControlEnvase(4, 2))
            Envase5 = Val(ControlEnvase(5, 1))
            CantiEnvase5 = Val(ControlEnvase(5, 2))
            
            WEnvAdicional = ""
            For ZZCiclo = 6 To 12
                ZZCampo1 = ControlEnvase(ZZCiclo, 1)
                ZZCampo2 = ControlEnvase(ZZCiclo, 2)
                Call Ceros(ZZCampo1, 4)
                Call Ceros(ZZCampo2, 4)
                WEnvAdicional = WEnvAdicional + ZZCampo1 + ZZCampo2
            Next ZZCiclo
            
            If Cantidad <> 0 Then
            
                If Left$(Articulo, 2) = "PT" Or Left$(Articulo, 2) = "PE" Then
                
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
                    
                        Else
                        
                    If Left$(Articulo, 2) = "DY" Then
                        WLinea = 16
                            Else
                        If Left$(Articulo, 2) = "DS" Then
                            WLinea = 16
                                Else
                            If Left$(Articulo, 2) = "DQ" Then
                                WLinea = 22
                                    Else
                                WLinea = 5
                            End If
                        End If
                    End If
                    
                End If
            
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
            
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                WTipo = "01"
                WNumero = Numero.Text
                XRenglon = Str$(Renglon)
                WArticulo = Articulo
                XCantidad = Str$(Cantidad)
                XPrecio = Str$(Precio * Val(Paridad.Text))
                XPrecioUs = Str$(Precio)
                XImporte = Str$(Precio * Cantidad * Val(Paridad.Text))
                XImporteUs = Str$(Precio * Cantidad)
                WCliente = Cliente.Text
                WParidad = Paridad.Text
                XVendedor = Str$(WVendedor)
                XRubro = Str$(WRubro)
                XLinea = Str$(WLinea)
                XCosto2 = ""
                XCosto1 = ""
                WCoeficiente = ""
                WPedido = Pedido.Text
                WFecha = Fecha.Text
                WImporte1 = ""
                WImporte2 = ""
                WImporte3 = ""
                WImporte4 = ""
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XArticulo = Left$(Articulo, 8)
                WRemito = Remito.Text
                WClave = "01" + Auxi1 + Auxi
                WDate = Date$
                XCanti = ""
                XImpo = ""
                XImpoUs = ""
                
                XMarca = ""
                If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                    Select Case WTipoPedido
                        Case "PG", "CO"
                            XMarca = ""
                        Case Else
                            XMarca = "X"
                    End Select
                End If
                
                WLote1 = Str$(lote1)
                WCanti1 = Str$(Canti1)
                WLote2 = Str$(lote2)
                WCanti2 = Str$(Canti2)
                Wlote3 = Str$(lote3)
                WCanti3 = Str$(Canti3)
                WLote4 = Str$(lote4)
                WCanti4 = Str$(Canti4)
                WLote5 = Str$(lote5)
                WCanti5 = Str$(Canti5)
                WTipoProDy = Left$(Articulo, 2)
                If WTipoProDy <> "PT" And WTipoProDy <> "PE" Then
                    XTipoproDy = "M"
                    XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                        Else
                    XTipoproDy = "T"
                    XArticuloDy = "  -   -   "
                End If
            
                XParam = "'" + WClave + "','" _
                            + WTipo + "','" + WNumero + "','" _
                            + XRenglon + "','" + WArticulo + "','" _
                            + XCantidad + "','" + XPrecio + "','" _
                            + XPrecioUs + "','" + XImporte + "','" _
                            + XImporteUs + "','" + WCliente + "','" _
                            + WParidad + "','" + XVendedor + "','" _
                            + XRubro + "','" + XLinea + "','" _
                            + XCosto1 + "','" + XCosto2 + "','" _
                            + WCoeficiente + "','" + WPedido + "','" _
                            + WFecha + "','" + WImporte1 + "','" _
                            + WImporte2 + "','" + WImporte3 + "','" _
                            + WImporte4 + "','" + WOrdFecha + "','" _
                            + XArticulo + "','" + WRemito + "','" _
                            + WDate + "','" + XCanti + "','" _
                            + XImpo + "','" + XImpoUs + "','" _
                            + XMarca + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + Wlote3 + "','" + WCanti3 + "','" _
                            + WLote4 + "','" + WCanti4 + "','" _
                            + WLote5 + "','" + WCanti5 + "','" _
                            + XTipoproDy + "','" + XArticuloDy + "'"
            
                spEstadistica = "AltaEstadistica " + XParam
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                XEnv1 = Str$(Envase1)
                XCantiEnv1 = Str$(CantiEnvase1)
                XEnv2 = Str$(Envase2)
                XCantiEnv2 = Str$(CantiEnvase2)
                XEnv3 = Str$(Envase3)
                XCantiEnv3 = Str$(CantiEnvase3)
                XEnv4 = Str$(Envase4)
                XCantiEnv4 = Str$(CantiEnvase4)
                XEnv5 = Str$(Envase5)
                XCantiEnv5 = Str$(CantiEnvase5)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Estadistica SET "
                ZSql = ZSql + " ClaveCtaCte = " + "'" + ZZClaveCtaCte + "',"
                ZSql = ZSql + " LoteAdicional = " + "'" + WLoteAdicional + "',"
                ZSql = ZSql + " EnvAdicional = " + "'" + WEnvAdicional + "',"
                ZSql = ZSql + " Env1 = " + "'" + XEnv1 + "',"
                ZSql = ZSql + " CantiEnv1 = " + "'" + XCantiEnv1 + "',"
                ZSql = ZSql + " Env2 = " + "'" + XEnv2 + "',"
                ZSql = ZSql + " CantiEnv2 = " + "'" + XCantiEnv2 + "',"
                ZSql = ZSql + " Env3 = " + "'" + XEnv3 + "',"
                ZSql = ZSql + " CantiEnv3 = " + "'" + XCantiEnv3 + "',"
                ZSql = ZSql + " Env4 = " + "'" + XEnv4 + "',"
                ZSql = ZSql + " CantiEnv4 = " + "'" + XCantiEnv4 + "',"
                ZSql = ZSql + " Env5 = " + "'" + XEnv5 + "',"
                ZSql = ZSql + " CantiEnv5 = " + "'" + XCantiEnv5 + "',"
                ZSql = ZSql + " ImpreTerminado = " + "'" + Left$(ZZImpreTerminado, 50) + "',"
                ZSql = ZSql + " ImpreCantidad = " + "'" + ZZImpre(WRenglon, 1) + "',"
                ZSql = ZSql + " ImpreTipo = " + "'" + ZZImpre(WRenglon, 2) + "',"
                ZSql = ZSql + " ImpreNumeros = " + "'" + ZZImpre(WRenglon, 3) + "',"
                ZSql = ZSql + " ImpreBruto = " + "'" + ZZImpre(WRenglon, 4) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                 
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                VectorCosto(Renglon, 1) = WArticulo
                VectorCosto(Renglon, 2) = WClave
                
                If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                    Select Case WTipoPedido
                        Case "FA", "PT", "BI"
                            XEmpresa = Wempresa
                            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                                Select Case WTipoPedido
                                    Case "PG", "CO"
                                        Wempresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "FA"
                                        Wempresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                        Wempresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End Select
                            End If
                    
                            XMarca = ""
                            XParam = "'" + WClave + "','" _
                                         + WTipo + "','" + WNumero + "','" _
                                         + XRenglon + "','" + WArticulo + "','" _
                                         + XCantidad + "','" + XPrecio + "','" _
                                         + XPrecioUs + "','" + XImporte + "','" _
                                         + XImporteUs + "','" + WCliente + "','" _
                                         + WParidad + "','" + XVendedor + "','" _
                                         + XRubro + "','" + XLinea + "','" _
                                         + XCosto1 + "','" + XCosto2 + "','" _
                                         + WCoeficiente + "','" + WPedido + "','" _
                                         + WFecha + "','" + WImporte1 + "','" _
                                         + WImporte2 + "','" + WImporte3 + "','" _
                                         + WImporte4 + "','" + WOrdFecha + "','" _
                                         + XArticulo + "','" + WRemito + "','" _
                                         + WDate + "','" + XCanti + "','" _
                                         + XImpo + "','" _
                                         + XImpoUs + "','" _
                                         + XMarca + "','" _
                                         + WLote1 + "','" + WCanti1 + "','" _
                                         + WLote2 + "','" + WCanti2 + "','" _
                                         + Wlote3 + "','" + WCanti3 + "','" _
                                         + WLote4 + "','" + WCanti4 + "','" _
                                         + WLote5 + "','" + WCanti5 + "','" _
                                         + XTipoproDy + "','" + XArticuloDy + "'"
                
                                    spEstadistica = "AltaEstadistica " + XParam
                                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "UPDATE Estadistica SET "
                                    ZSql = ZSql + " LoteAdicional = " + "'" + WLoteAdicional + "',"
                                    ZSql = ZSql + " EnvAdicional = " + "'" + WEnvAdicional + "',"
                                    ZSql = ZSql + " Env1 = " + "'" + XEnv1 + "',"
                                    ZSql = ZSql + " CantiEnv1 = " + "'" + XCantiEnv1 + "',"
                                    ZSql = ZSql + " Env2 = " + "'" + XEnv2 + "',"
                                    ZSql = ZSql + " CantiEnv2 = " + "'" + XCantiEnv2 + "',"
                                    ZSql = ZSql + " Env3 = " + "'" + XEnv3 + "',"
                                    ZSql = ZSql + " CantiEnv3 = " + "'" + XCantiEnv3 + "',"
                                    ZSql = ZSql + " Env4 = " + "'" + XEnv4 + "',"
                                    ZSql = ZSql + " CantiEnv4 = " + "'" + XCantiEnv4 + "',"
                                    ZSql = ZSql + " Env5 = " + "'" + XEnv5 + "',"
                                    ZSql = ZSql + " CantiEnv5 = " + "'" + XCantiEnv5 + "'"
                                    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                    
                                    spEstadistica = ZSql
                                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Call Conecta_Empresa
                        
                        Case Else
                    End Select
                End If
                
                WLote1 = ZZControlLote(WRenglon, 1)
                WCanti1 = ZZControlLote(WRenglon, 2)
                WLote2 = ZZControlLote(WRenglon, 3)
                WCanti2 = ZZControlLote(WRenglon, 4)
                Wlote3 = ZZControlLote(WRenglon, 5)
                WCanti3 = ZZControlLote(WRenglon, 6)
                WLote4 = ZZControlLote(WRenglon, 7)
                WCanti4 = ZZControlLote(WRenglon, 8)
                WLote5 = ZZControlLote(WRenglon, 9)
                WCanti5 = ZZControlLote(WRenglon, 10)
                WLote16 = ZZControlLote(WRenglon, 11)
                WCanti6 = ZZControlLote(WRenglon, 12)
                Wlote17 = ZZControlLote(WRenglon, 13)
                WCanti7 = ZZControlLote(WRenglon, 14)
                WLote18 = ZZControlLote(WRenglon, 15)
                WCanti8 = ZZControlLote(WRenglon, 16)
                Wlote19 = ZZControlLote(WRenglon, 17)
                WCanti9 = ZZControlLote(WRenglon, 18)
                Wlote110 = ZZControlLote(WRenglon, 19)
                WCanti10 = ZZControlLote(WRenglon, 20)
                WLote11 = ZZControlLote(WRenglon, 21)
                WCanti11 = ZZControlLote(WRenglon, 22)
                WLote12 = ZZControlLote(WRenglon, 23)
                WCanti12 = ZZControlLote(WRenglon, 24)
                
            
                Auxiliar(Renglon, 1) = Articulo
                Auxiliar(Renglon, 2) = Cantidad
                Auxiliar(Renglon, 3) = Precio
                Auxiliar(Renglon, 4) = WRenglon
                Auxiliar(Renglon, 5) = WLote1
                Auxiliar(Renglon, 6) = WCanti1
                Auxiliar(Renglon, 7) = WLote2
                Auxiliar(Renglon, 8) = WCanti2
                Auxiliar(Renglon, 9) = Wlote3
                Auxiliar(Renglon, 10) = WCanti3
                Auxiliar(Renglon, 11) = WLote4
                Auxiliar(Renglon, 12) = WCanti4
                Auxiliar(Renglon, 13) = WLote5
                Auxiliar(Renglon, 14) = WCanti5
                Auxiliar(Renglon, 15) = WLote6
                Auxiliar(Renglon, 16) = WCanti6
                Auxiliar(Renglon, 17) = WLote7
                Auxiliar(Renglon, 18) = WCanti7
                Auxiliar(Renglon, 19) = WLote8
                Auxiliar(Renglon, 20) = WCanti8
                Auxiliar(Renglon, 21) = WLote9
                Auxiliar(Renglon, 22) = WCanti9
                Auxiliar(Renglon, 23) = WLote10
                Auxiliar(Renglon, 24) = WCanti10
                Auxiliar(Renglon, 25) = WLote11
                Auxiliar(Renglon, 26) = WCanti11
                Auxiliar(Renglon, 27) = WLote12
                Auxiliar(Renglon, 28) = WCanti12
                Auxiliar(Renglon, 30) = ZZImpreTerminado

            
                AuxiliarII(Renglon, 1) = Articulo
                AuxiliarII(Renglon, 2) = Cantidad
                AuxiliarII(Renglon, 3) = Precio
                AuxiliarII(Renglon, 4) = WRenglon
                AuxiliarII(Renglon, 5) = WLote1
                AuxiliarII(Renglon, 6) = WCanti1
                AuxiliarII(Renglon, 7) = WLote2
                AuxiliarII(Renglon, 8) = WCanti2
                AuxiliarII(Renglon, 9) = Wlote3
                AuxiliarII(Renglon, 10) = WCanti3
                AuxiliarII(Renglon, 11) = WLote4
                AuxiliarII(Renglon, 12) = WCanti4
                AuxiliarII(Renglon, 13) = WLote5
                AuxiliarII(Renglon, 14) = WCanti5
                AuxiliarII(Renglon, 15) = WLote6
                AuxiliarII(Renglon, 16) = WCanti6
                AuxiliarII(Renglon, 17) = WLote7
                AuxiliarII(Renglon, 18) = WCanti7
                AuxiliarII(Renglon, 19) = WLote8
                AuxiliarII(Renglon, 20) = WCanti8
                AuxiliarII(Renglon, 21) = WLote9
                AuxiliarII(Renglon, 22) = WCanti9
                AuxiliarII(Renglon, 23) = WLote10
                AuxiliarII(Renglon, 24) = WCanti10
                AuxiliarII(Renglon, 25) = WLote11
                AuxiliarII(Renglon, 26) = WCanti11
                AuxiliarII(Renglon, 27) = WLote12
                AuxiliarII(Renglon, 28) = WCanti12
                AuxiliarII(Renglon, 30) = ZZImpreTerminado



            End If
                                
        Next iRow
        
    Next a
    
    For DA = 1 To Renglon
    
        Articulo = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        Precio = Auxiliar(DA, 3)
        WRenglon = Auxiliar(DA, 4)
        lote1 = Auxiliar(DA, 5)
        Cantidad1 = Auxiliar(DA, 6)
        lote2 = Auxiliar(DA, 7)
        Cantidad2 = Auxiliar(DA, 8)
        lote3 = Auxiliar(DA, 9)
        Cantidad3 = Auxiliar(DA, 10)
        lote4 = Auxiliar(DA, 11)
        Cantidad4 = Auxiliar(DA, 12)
        lote5 = Auxiliar(DA, 13)
        Cantidad5 = Auxiliar(DA, 14)
        lote6 = Auxiliar(DA, 15)
        Cantidad6 = Auxiliar(DA, 16)
        lote7 = Auxiliar(DA, 17)
        Cantidad7 = Auxiliar(DA, 18)
        lote8 = Auxiliar(DA, 19)
        Cantidad8 = Auxiliar(DA, 20)
        lote9 = Auxiliar(DA, 21)
        Cantidad9 = Auxiliar(DA, 22)
        lote10 = Auxiliar(DA, 23)
        Cantidad10 = Auxiliar(DA, 24)
        lote11 = Auxiliar(DA, 25)
        Cantidad11 = Auxiliar(DA, 26)
        lote12 = Auxiliar(DA, 27)
        Cantidad12 = Auxiliar(DA, 28)
        
        WLote(1, 1) = lote1
        WLote(1, 2) = Cantidad1
        WLote(2, 1) = lote2
        WLote(2, 2) = Cantidad2
        WLote(3, 1) = lote3
        WLote(3, 2) = Cantidad3
        WLote(4, 1) = lote4
        WLote(4, 2) = Cantidad4
        WLote(5, 1) = lote5
        WLote(5, 2) = Cantidad5
        WLote(6, 1) = lote6
        WLote(6, 2) = Cantidad6
        WLote(7, 1) = lote7
        WLote(7, 2) = Cantidad7
        WLote(8, 1) = lote8
        WLote(8, 2) = Cantidad8
        WLote(9, 1) = lote9
        WLote(9, 2) = Cantidad9
        WLote(10, 1) = lote10
        WLote(10, 2) = Cantidad10
        WLote(11, 1) = lote11
        WLote(11, 2) = Cantidad11
        WLote(12, 1) = lote12
        WLote(12, 2) = Cantidad12
        
        WTipoProDy = Left$(Articulo, 2)
        If WTipoProDy <> "PT" And WTipoProDy <> "PE" Then
            XTipoproDy = "M"
            XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                Else
            XTipoproDy = "T"
            XArticuloDy = "  -   -   "
        End If
        
        If XTipoproDy = "M" Then
        
            XEmpresa = Wempresa
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "FA"
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            End If
                
            spArticulo = "ConsultaArticulo " + "'" + XArticuloDy + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = XArticuloDy
                WPedido = Str$(rstArticulo!Venta - Cantidad)
                WSalidas = Str$(rstArticulo!Salidas + Cantidad)
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                        + WPedido + "','" _
                        + WSalidas + "','" _
                        + WDate + "'"
                spArticulo = "ModificaArticuloFacturas " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                    
            For Da2 = 1 To 12
                    
                lote1 = WLote(Da2, 1)
                Cantidad1 = WLote(Da2, 2)
                        
                If Val(lote1) <> 0 Then
                    
                    XParam = "'" + lote1 + "','" _
                            + XArticuloDy + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WClave = rstLaudo!Clave
                        WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad1))
                        WDate = Date$
                        rstLaudo.Close
                        
                        XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                        spLaudo = "ModificaLaudoSaldo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        
                                Else
                            
                        XParam = "'" + XArticuloDy + "','" _
                                    + lote1 + "'"
                        spMovguia = "ListaMovguiaLote " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WClave = rstMovguia!Clave
                            WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad1))
                            WDate = Date$
                            rstMovguia.Close
                            
                            XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                            spMovguia = "ModificaMovguiaSaldo " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                    End If
                End If
            Next Da2
                
            Call Conecta_Empresa
                    
                Else
                
            XEmpresa = Wempresa
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "FA"
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            End If
                
                        
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                WCodigo = Articulo
                WPedido = Str$(rstTerminado!Pedido - Cantidad)
                WSalidas = Str$(rstTerminado!Salidas + Cantidad)
                WDate = Date$
                
                WLinea = rstTerminado!Linea
                rstTerminado.Close
            
                XParam = "'" + WCodigo + "','" _
                        + WPedido + "','" _
                        + WSalidas + "','" _
                        + WDate + "'"
                                       
                spTerminado = "ModificaTerminadoFacturas " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
            For Da2 = 1 To 12
        
                If Val(WLote(Da2, 1)) <> 0 Then
                    Lote = WLote(Da2, 1)
                    Cantilote = WLote(Da2, 2)
                
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                + Articulo + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            WClave = rstHoja!Clave
                            WSaldo = Str$(rstHoja!Saldo - Cantilote)
                            WDate = Date$
                            rstHoja.Close
                        
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificaHojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                            
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                WSaldo = Str$(rstMovguia!Saldo - Cantilote)
                                WDate = Date$
                                rstMovguia.Close
                    
                                XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                        
                        End If
                    End If
                End If
        
            Next Da2
            
            Call Conecta_Empresa
            
        End If
        

        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
    
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        ClavePedido = Auxi + Auxi1
        
        XParam = "'" + Left$(ClavePedido, 6) + "','" _
                    + Right$(ClavePedido, 2) + "'"
        spPedido = "ConsultaPedido2 " + XParam
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            WFacturado = Str$(rstPedido!Facturado + Cantidad)
            If Val(WFacturado) > rstPedido!Cantidad Then
                WFacturado = Str$(rstPedido!Cantidad)
            End If
            rstPedido.Close
            XParam = "'" + ClavePedido + "','" _
                        + WFacturado + "'"
                                       
            spPedido = "ModificaPedidoFacturas " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        If XTipoproDy = "M" Then
        
            ClavePrecioMp = Cliente.Text + XArticuloDy
        
            spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePrecioMp + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
        
                WFecha1 = ""
                WFactura1 = ""
                WPrecio1 = ""
                WCantidad1 = ""
            
                WFecha2 = ""
                WFactura2 = ""
                WPrecio2 = ""
                WCantidad2 = ""
            
                WFecha3 = ""
                WFactura3 = ""
                WPrecio3 = ""
                WCantidad3 = ""
            
                WFecha4 = ""
                WFactura4 = ""
                WPrecio4 = ""
                WCantidad4 = ""
            
                WFecha5 = ""
                WFactura5 = ""
                WPrecio5 = ""
                WCantidad5 = ""
            
                If rstPreciosMp!Cantidad2 <> O Then
                    WFecha1 = rstPreciosMp!Fecha2
                    WFactura1 = rstPreciosMp!Factura2
                    WPrecio1 = Str$(rstPreciosMp!Precio2)
                    WCantidad1 = Str$(rstPreciosMp!Cantidad2)
                End If
                            
                If rstPreciosMp!Cantidad3 <> O Then
                    WFecha2 = rstPreciosMp!Fecha3
                    WFactura2 = rstPreciosMp!Factura3
                    WPrecio2 = Str$(rstPreciosMp!Precio3)
                    WCantidad2 = Str$(rstPreciosMp!Cantidad3)
                End If
                            
                If rstPreciosMp!Cantidad4 <> O Then
                    WFecha3 = rstPreciosMp!Fecha4
                    WFactura3 = rstPreciosMp!Factura4
                    WPrecio3 = Str$(rstPreciosMp!Precio4)
                    WCantidad3 = Str$(rstPreciosMp!Cantidad4)
                End If
                            
                If rstPreciosMp!Cantidad5 <> O Then
                    WFecha4 = rstPreciosMp!Fecha5
                    WFactura4 = rstPreciosMp!Factura5
                    WPrecio4 = Str$(rstPreciosMp!Precio5)
                    WCantidad4 = Str$(rstPreciosMp!Cantidad5)
                End If
                            
                WFecha5 = Fecha.Text
                WFactura5 = Numero.Text
                WPrecio5 = Str$(Precio)
                WCantidad5 = Str$(Cantidad)
                            
                WDate = Date$
            
                rstPreciosMp.Close
            
                XParam = "'" + ClavePrecioMp + "','" _
                        + WFecha1 + "','" _
                        + WFactura1 + "','" _
                        + WPrecio1 + "','" _
                        + WCantidad1 + "','" _
                        + WFecha2 + "','" _
                        + WFactura2 + "','" _
                        + WPrecio2 + "','" _
                        + WCantidad2 + "','" _
                        + WFecha3 + "','" _
                        + WFactura3 + "','" _
                        + WPrecio3 + "','" _
                        + WCantidad3 + "','" _
                        + WFecha4 + "','" _
                        + WFactura4 + "','" _
                        + WPrecio4 + "','" _
                        + WCantidad4 + "','" _
                        + WFecha5 + "','" _
                        + WFactura5 + "','" _
                        + WPrecio5 + "','" _
                        + WCantidad5 + "','" _
                        + WDate + "'"
                                       
                spPreciosMp = "ModificaPreciosFacturaMp " + XParam
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
                Else
            
            ClavePrecio = Cliente.Text + Articulo
        
            spPrecios = "ConsultaPrecios " + "'" + ClavePrecio + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
        
                WFecha1 = ""
                WFactura1 = ""
                WPrecio1 = ""
                WCantidad1 = ""
            
                WFecha2 = ""
                WFactura2 = ""
                WPrecio2 = ""
                WCantidad2 = ""
            
                WFecha3 = ""
                WFactura3 = ""
                WPrecio3 = ""
                WCantidad3 = ""
            
                WFecha4 = ""
                WFactura4 = ""
                WPrecio4 = ""
                WCantidad4 = ""
            
                WFecha5 = ""
                WFactura5 = ""
                WPrecio5 = ""
                WCantidad5 = ""
            
                If rstPrecios!Cantidad2 <> O Then
                    WFecha1 = rstPrecios!Fecha2
                    WFactura1 = rstPrecios!Factura2
                    WPrecio1 = Str$(rstPrecios!Precio2)
                    WCantidad1 = Str$(rstPrecios!Cantidad2)
                End If
                            
                If rstPrecios!Cantidad3 <> O Then
                    WFecha2 = rstPrecios!Fecha3
                    WFactura2 = rstPrecios!Factura3
                    WPrecio2 = Str$(rstPrecios!Precio3)
                    WCantidad2 = Str$(rstPrecios!Cantidad3)
                End If
                            
                If rstPrecios!Cantidad4 <> O Then
                    WFecha3 = rstPrecios!Fecha4
                    WFactura3 = rstPrecios!Factura4
                    WPrecio3 = Str$(rstPrecios!Precio4)
                    WCantidad3 = Str$(rstPrecios!Cantidad4)
                End If
                            
                If rstPrecios!Cantidad5 <> O Then
                    WFecha4 = rstPrecios!Fecha5
                    WFactura4 = rstPrecios!Factura5
                    WPrecio4 = Str$(rstPrecios!Precio5)
                    WCantidad4 = Str$(rstPrecios!Cantidad5)
                End If
                            
                WFecha5 = Fecha.Text
                WFactura5 = Numero.Text
                WPrecio5 = Str$(Precio)
                WCantidad5 = Str$(Cantidad)
                            
                WDate = Date$
            
                rstPrecios.Close
            
                XParam = "'" + ClavePrecio + "','" _
                        + WFecha1 + "','" _
                        + WFactura1 + "','" _
                        + WPrecio1 + "','" _
                        + WCantidad1 + "','" _
                        + WFecha2 + "','" _
                        + WFactura2 + "','" _
                        + WPrecio2 + "','" _
                        + WCantidad2 + "','" _
                        + WFecha3 + "','" _
                        + WFactura3 + "','" _
                        + WPrecio3 + "','" _
                        + WCantidad3 + "','" _
                        + WFecha4 + "','" _
                        + WFactura4 + "','" _
                        + WPrecio4 + "','" _
                        + WCantidad4 + "','" _
                        + WFecha5 + "','" _
                        + WFactura5 + "','" _
                        + WPrecio5 + "','" _
                        + WCantidad5 + "','" _
                        + WDate + "'"
                                       
                spPrecios = "ModificaPreciosFactura " + XParam
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
        
    Next DA
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Pedido SET "
    ZSql = ZSql & "MarcaFactura = " + "'" + "0" + "'"
    ZSql = ZSql & " Where Pedido = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
    spNumero = "ConsultaNumero " + "'" + "02" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        WCodigo = "02"
        WNumero = Numero.Text
        rstNumero.Close
        XParam = "'" + WCodigo + "','" _
                     + WNumero + "'"
        spNumero = "ModificaNumero " + XParam
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    For Ciclo = 1 To 99
    
        If VectorCosto(Ciclo, 1) <> "" Then
        
            ZZZProducto = VectorCosto(Ciclo, 1)
            ZZClave = VectorCosto(Ciclo, 2)
            
            ZZZCosto = 0
            Call Calcula_CostoFactura(ZZZProducto, ZZZCosto)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Estadistica SET "
            ZSql = ZSql + " Costo1 = " + "'" + Str$(ZZZCosto) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    Call Impresion_FE
    Call Impresion_Remito
    If Val(Wempresa) = 1 Then
        ZZDestino = 1
        Call Impresion_Varios
    End If
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Numero.SetFocus
        
End Sub

Private Sub Graba1_Click()
    Call Calcula_Click
    Call Impresion
    Numero.SetFocus
End Sub

Private Sub ImpreCertificado_Click()
    Call Impresion_Certificado_Pantalla
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    Cae.Text = ""
    CipLista.ListIndex = 0
    Idioma.ListIndex = 0
    TipoPedido.ListIndex = 0
    IdiomaII.ListIndex = 0
    
    For a = 0 To 8
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Rem Iva1.Caption = ""
    Rem Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Rem Dto.Caption = ""
    Rem Interes.Caption = ""
    Seguro.Text = ""
    Flete.Text = ""
    Gastos.Text = ""
    Descuento.Text = ""

    Marca.Text = ""
    Envio1.Text = ""
    Envio2.Text = ""
    Dolar1.Text = ""
    Dolar2.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    NroOrden.Text = ""
    fecorden.Text = "  /  /    "
    Consignatario.Text = ""
    
    spNumero = "ConsultaNumero " + "'" + "02" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 39 Then
        KeyCode = 13
    End If

    Select Case DBGrid1.Col
        Case 0, 1, 2, 3, 4
            Select Case KeyCode
                Case Else
                    Rem If KeyCode <> 0 Then Stop
            End Select
        Case Else
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

    TipoPedido.Clear
    
    TipoPedido.AddItem "Normal"
    TipoPedido.AddItem "S/Actualizar"

    TipoPedido.ListIndex = 0

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
    
    CipLista.Clear
    
    CipLista.AddItem ""
    CipLista.AddItem "EXW"
    CipLista.AddItem "FCA"
    CipLista.AddItem "FAS"
    CipLista.AddItem "FOB"
    CipLista.AddItem "CFR"
    CipLista.AddItem "CIF"
    CipLista.AddItem "CPT"
    CipLista.AddItem "CIP"
    CipLista.AddItem "DAF"
    CipLista.AddItem "DES"
    CipLista.AddItem "DEQ"
    CipLista.AddItem "DDA"
    CipLista.AddItem "DDP"
    CipLista.AddItem "DAP"
    
    CipLista.ListIndex = 0
    
    IdiomaII.Clear
    
    IdiomaII.AddItem "Fc.Caste"
    IdiomaII.AddItem "Fc.Ingles"
    
    IdiomaII.ListIndex = 0
    
    
    
    Idioma.Clear
    
    Idioma.AddItem ""
    Idioma.AddItem "Español"
    Idioma.AddItem "Ingles"
    
    Idioma.ListIndex = 0
    

    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 5, 0 To 99)

mTotalRows& = 100

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
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Parcial"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    
    Neto.Caption = ""
    Rem Iva1.Caption = ""
    Rem Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Rem Dto.Caption = ""
    Rem Interes.Caption = ""
    Seguro.Text = ""
    Flete.Text = ""
    Gastos.Text = ""
    Descuento.Text = ""
    
    Marca.Text = ""
    Envio1.Text = ""
    Envio2.Text = ""
    Dolar1.Text = ""
    Dolar2.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    NroOrden.Text = ""
    fecorden.Text = "  /  /    "
    Consignatario.Text = ""
    
    spNumero = "ConsultaNumero " + "'" + "02" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Numero.SetFocus
     
End Sub

Private Sub Proceso_Click()

    For a = 0 To 8
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    WSeguro = 0
    WFlete = 0
    WGastos = 0
    WDescuento = 0
    
    Erase Auxiliar
    Erase ZZControlLote
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
    
                    If TipoPedido.ListIndex = 0 Then
                        Canti = !Cantidad1
                            Else
                        Canti = !Cantidad - !Facturado
                    End If
                    
                
                    Select Case Mid$(!Terminado, 1, 2)
                        Case "Z2"
                            WSeguro = WSeguro + (!Precio * !Cantidad)
                                                            
                        Case "Z1"
                            WFlete = WFlete + (!Precio * !Cantidad)
                    
                        Case Else
                            Renglon = Renglon + 1
                            Rem Canti = Renglon
            
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = !Terminado
                            Auxi1 = !Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                            Rem DBGrid1.Text = Pusing("###,###.##", Str$(Canti))
                
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Precio))
                
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###.##", Str$(Canti))
                            
                            ZZControlLote(Renglon, 1) = Str$(!lote1)
                            ZZControlLote(Renglon, 2) = Str$(!CantiLote1)
                            ZZControlLote(Renglon, 3) = Str$(!lote2)
                            ZZControlLote(Renglon, 4) = Str$(!CantiLote2)
                            ZZControlLote(Renglon, 5) = Str$(!lote3)
                            ZZControlLote(Renglon, 6) = Str$(!CantiLote3)
                            ZZControlLote(Renglon, 7) = Str$(!lote4)
                            ZZControlLote(Renglon, 8) = Str$(!CantiLote4)
                            ZZControlLote(Renglon, 9) = Str$(!lote5)
                            ZZControlLote(Renglon, 10) = Str$(!CantiLote5)
                            ZLote6 = IIf(IsNull(!lote6), "0", !lote6)
                            ZCantiLote6 = IIf(IsNull(!CantiLote6), "0", !CantiLote6)
                            ZZControlLote(Renglon, 11) = Str$(ZLote6)
                            ZZControlLote(Renglon, 12) = Str$(ZCantiLote6)
                            ZLote7 = IIf(IsNull(!lote7), "0", !lote7)
                            ZCantiLote7 = IIf(IsNull(!CantiLote7), "0", !CantiLote7)
                            ZZControlLote(Renglon, 13) = Str$(ZLote7)
                            ZZControlLote(Renglon, 14) = Str$(ZCantiLote7)
                            ZLote8 = IIf(IsNull(!lote8), "0", !lote8)
                            ZCantiLote8 = IIf(IsNull(!CantiLote8), "0", !CantiLote8)
                            ZZControlLote(Renglon, 15) = Str$(ZLote8)
                            ZZControlLote(Renglon, 16) = Str$(ZCantiLote8)
                            ZLote9 = IIf(IsNull(!lote9), "0", !lote9)
                            ZCantiLote9 = IIf(IsNull(!CantiLote9), "0", !CantiLote9)
                            ZZControlLote(Renglon, 17) = Str$(ZLote9)
                            ZZControlLote(Renglon, 18) = Str$(ZCantiLote9)
                            ZLote10 = IIf(IsNull(!lote10), "0", !lote10)
                            ZCantiLote10 = IIf(IsNull(!CantiLote10), "0", !CantiLote10)
                            ZZControlLote(Renglon, 19) = Str$(ZLote10)
                            ZZControlLote(Renglon, 20) = Str$(ZCantiLote10)
                            ZLote11 = IIf(IsNull(!lote11), "0", !lote11)
                            ZCantiLote11 = IIf(IsNull(!CantiLote11), "0", !CantiLote11)
                            ZZControlLote(Renglon, 21) = Str$(ZLote11)
                            ZZControlLote(Renglon, 22) = Str$(ZCantiLote11)
                            ZLote12 = IIf(IsNull(!lote12), "0", !lote12)
                            ZCantiLote12 = IIf(IsNull(!CantiLote12), "0", !CantiLote12)
                            ZZControlLote(Renglon, 23) = Str$(ZLote12)
                            ZZControlLote(Renglon, 24) = Str$(ZCantiLote12)
                               
                            ZZControlLote(Renglon, 31) = Str$(!Env1)
                            ZZControlLote(Renglon, 32) = Str$(!CantiEnv1)
                            ZZControlLote(Renglon, 33) = Str$(!Env2)
                            ZZControlLote(Renglon, 34) = Str$(!CantiEnv2)
                            ZZControlLote(Renglon, 35) = Str$(!Env3)
                            ZZControlLote(Renglon, 36) = Str$(!CantiEnv3)
                            ZZControlLote(Renglon, 37) = Str$(!Env4)
                            ZZControlLote(Renglon, 38) = Str$(!CantiEnv4)
                            ZZControlLote(Renglon, 39) = Str$(!Env5)
                            ZZControlLote(Renglon, 40) = Str$(!CantiEnv5)
                            Rem ZZControlLote(Renglon, 41) = Str$(!Env6)
                            Rem ZZControlLote(Renglon, 42) = Str$(!CantiEnv6)
                            Rem ZZControlLote(Renglon, 43) = Str$(!Env7)
                            Rem ZZControlLote(Renglon, 44) = Str$(!CantiEnv7)
                            Rem ZZControlLote(Renglon, 45) = Str$(!Env8)
                            Rem ZZControlLote(Renglon, 46) = Str$(!CantiEnv8)
                            Rem ZZControlLote(Renglon, 47) = Str$(!Env9)
                            Rem ZZControlLote(Renglon, 48) = Str$(!CantiEnv9)
                            Rem ZZControlLote(Renglon, 49) = Str$(!Env10)
                            Rem ZZControlLote(Renglon, 50) = Str$(!CantiEnv10)
                            Rem ZZControlLote(Renglon, 51) = Str$(!Env11)
                            Rem ZZControlLote(Renglon, 52) = Str$(!CantiEnv11)
                            Rem ZZControlLote(Renglon, 53) = Str$(!Env12)
                            Rem ZZControlLote(Renglon, 54) = Str$(!CantiEnv12)

                            Auxiliar(Renglon, 1) = Auxi1
                            Auxiliar(Renglon, 2) = Canti
                            
                            XEnvase(Renglon, 1) = rstPedido!Envase1
                            XEnvase(Renglon, 2) = rstPedido!Canti1
                            XEnvase(Renglon, 3) = rstPedido!Envase2
                            XEnvase(Renglon, 4) = rstPedido!Canti2
                            XEnvase(Renglon, 5) = rstPedido!Envase3
                            XEnvase(Renglon, 6) = rstPedido!Canti3
                            
                    End Select
                                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    Seguro.Text = Str$(WSeguro)
    Seguro.Text = Pusing("###,###.##", Seguro.Text)
    Flete.Text = Str$(WFlete)
    Flete.Text = Pusing("###,###.##", Flete.Text)
    Gastos.Text = Str$(WGastos)
    Gastos.Text = Pusing("###,###.##", Gastos.Text)
    Descuento.Text = Str$(WDescuento)
    Descuento.Text = Pusing("###,###.##", Descuento.Text)
    
    WRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
        If Left$(Auxi1, 2) <> "PT" And Left$(Auxi1, 2) <> "PE" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                ClavePreciosMp = Cliente.Text + WArti
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    Precio = rstPreciosMp!Precio
                
                    rstPreciosMp.Close
                End If

                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    
                    rstArticulo.Close
                End If
        
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
        
                    DBGrid1.Col = 1
                    
                    DBGrid1.Text = rstPrecios!Descripcion
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    Precio = rstPrecios!Precio
                    rstPrecios.Close
                End If
        End Select

        If Val(Canti) <> 0 Then
            WNeto = WNeto + (Val(Canti) * Precio)
        End If
        
    Next DA
    
    Call Calcula_Click

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
    
    Graba.Enabled = True

End Sub

Private Sub Proceso1_Click()

    WNeto = 0

    For a = 0 To 8
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Erase AuxiliarII
    
    Renglon = 0
    Erase Auxiliar
    
    XParam = "'" + "01" + "','" _
                + Numero.Text + "'"
    
    spEstadistica = "ConsultaEstadistica1 " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        
        
        
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                   
                    
                    
                    
                    
                    
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                    
                    DBGrid1.Col = 1
                    DBGrid1.Text = IIf(IsNull(rstEstadistica!impreterminado), "", rstEstadistica!impreterminado)
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    Dada = Str$(rstEstadistica!PrecioUs)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                    
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstEstadistica!Clave
                
                    Dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("#,###.###", Dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!PrecioUs)
                    End If
            
                    AuxiliarII(Renglon, 1) = rstEstadistica!Articulo
                    AuxiliarII(Renglon, 2) = Str$(rstEstadistica!Cantidad)
                    AuxiliarII(Renglon, 3) = Str$(rstEstadistica!PrecioUs)
                    AuxiliarII(Renglon, 4) = Renglon
                    
                    AuxiliarII(Renglon, 5) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    AuxiliarII(Renglon, 6) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    AuxiliarII(Renglon, 7) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    AuxiliarII(Renglon, 8) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                    AuxiliarII(Renglon, 9) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    AuxiliarII(Renglon, 10) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                    AuxiliarII(Renglon, 11) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    AuxiliarII(Renglon, 12) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                    AuxiliarII(Renglon, 13) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    AuxiliarII(Renglon, 14) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                    AuxiliarII(Renglon, 30) = IIf(IsNull(rstEstadistica!impreterminado), "0", rstEstadistica!impreterminado)
                    
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                    
                    If Len(Trim(WLoteAdicional)) = 98 Then
                        AuxiliarII(Renglon, 15) = Mid$(WLoteAdicional, 1, 8)
                        AuxiliarII(Renglon, 16) = Mid$(WLoteAdicional, 9, 6)
                        AuxiliarII(Renglon, 17) = Mid$(WLoteAdicional, 15, 8)
                        AuxiliarII(Renglon, 18) = Mid$(WLoteAdicional, 23, 6)
                        AuxiliarII(Renglon, 19) = Mid$(WLoteAdicional, 29, 8)
                        AuxiliarII(Renglon, 20) = Mid$(WLoteAdicional, 37, 6)
                        AuxiliarII(Renglon, 21) = Mid$(WLoteAdicional, 43, 8)
                        AuxiliarII(Renglon, 22) = Mid$(WLoteAdicional, 51, 6)
                        AuxiliarII(Renglon, 23) = Mid$(WLoteAdicional, 57, 8)
                        AuxiliarII(Renglon, 24) = Mid$(WLoteAdicional, 65, 6)
                        AuxiliarII(Renglon, 25) = Mid$(WLoteAdicional, 71, 8)
                        AuxiliarII(Renglon, 26) = Mid$(WLoteAdicional, 79, 6)
                        AuxiliarII(Renglon, 27) = Mid$(WLoteAdicional, 85, 8)
                        AuxiliarII(Renglon, 28) = Mid$(WLoteAdicional, 93, 6)
                            Else
                        AuxiliarII(Renglon, 15) = ""
                        AuxiliarII(Renglon, 16) = ""
                        AuxiliarII(Renglon, 17) = ""
                        AuxiliarII(Renglon, 18) = ""
                        AuxiliarII(Renglon, 19) = ""
                        AuxiliarII(Renglon, 20) = ""
                        AuxiliarII(Renglon, 21) = ""
                        AuxiliarII(Renglon, 22) = ""
                        AuxiliarII(Renglon, 23) = ""
                        AuxiliarII(Renglon, 24) = ""
                        AuxiliarII(Renglon, 25) = ""
                        AuxiliarII(Renglon, 26) = ""
                        AuxiliarII(Renglon, 27) = ""
                        AuxiliarII(Renglon, 28) = ""
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = rstEstadistica!Clave
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    If ZZDada = 9999 Then
    
    XRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
        ZZClave = Auxiliar(DA, 2)
        
        If Left$(Auxi1, 2) <> "PT" And Left$(Auxi1, 2) <> "PE" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                    
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                    
                    DBGrid1.Col = 1
                    Rem DBGrid1.Text = "CALCIUM CITRATE TETRAHYDRAT ULTRADENSE"
                    DBGrid1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
                If Val(Wempresa) = 1 And IdiomaII.ListIndex = 1 Then
                
                    Articulo = Auxi1
                    XXXCodigo = Val(Mid$(Articulo, 4, 5))
                    If XXXCodigo >= 25000 And XXXCodigo <= 25999 Then
                    
                        If Left$(Articulo, 2) = "PT" Or Left$(Articulo, 2) = "PE" Then
                        
                            ZZDescripcionIngles = ""
                            ZZDescripcionInglesII = ""
                        
                            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                ZZDescripcionIngles = IIf(IsNull(rstTerminado!Descripcioningles), "", rstTerminado!Descripcioningles)
                                rstTerminado.Close
                            End If
                            
                            ClavePrecios = Cliente.Text + Articulo
                            spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                            If rstPrecios.RecordCount > 0 Then
                                ZZDescripcionInglesII = rstPrecios!Descripcion
                                rstPrecios.Close
                            End If
                        
                            ZZImpreTerminado = ZZDescripcionIngles + " - " + ZZDescripcionInglesII
                            DBGrid1.Col = 1
                            DBGrid1.Text = ZZImpreTerminado
                            
                        End If
                        
                    End If
                End If
                
        End Select
    Next DA
    
    End If
    
    
    
    If ZZDada = 9999 Then
    
    
    XRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
        Articulo = Auxi1
        ZZClave = Auxiliar(DA, 2)
    
        Renglon = Renglon + 1

        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
        
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        ZZDescripcionIngles = ""
        ZZDescripcionInglesII = ""
    
        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZZDescripcionIngles = IIf(IsNull(rstTerminado!Descripcioningles), "", rstTerminado!Descripcioningles)
            rstTerminado.Close
        End If
        
        ClavePrecios = Cliente.Text + Articulo
        spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            ZZDescripcionInglesII = rstPrecios!descripcionfarma
            rstPrecios.Close
        End If
    
        ZZImpreTerminado = Trim(ZZDescripcionIngles) + " - " + Trim(ZZDescripcionInglesII)
        DBGrid1.Col = 1
        DBGrid1.Text = ZZImpreTerminado
        
        ZSql = ""
        ZSql = ZSql & "UPDATE Estadistica SET "
        ZSql = ZSql & "ImpreTerminado = " + "'" + ZZImpreTerminado + "'"
        ZSql = ZSql & " Where Clave = " + "'" + ZZClave + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next DA
    
    
    End If
    
    
    
    
    
    
    
    
    
    Call Calcula_Click

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
    
    Graba.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
        
            Pedido.Text = rstCtacte!Pedido
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Remito.Text = rstCtacte!Remito
            Orden.Text = rstCtacte!Orden
            
            WParidad = IIf(IsNull(rstCtacte!Paridad), "0", rstCtacte!Paridad)
            Paridad.Text = Str$(WParidad)
            Paridad.Text = Pusing("#,###.###", Paridad.Text)
            
            WSeguro = IIf(IsNull(rstCtacte!Seguro), "0", rstCtacte!Seguro)
            WFlete = IIf(IsNull(rstCtacte!Flete), "0", rstCtacte!Flete)
            WGastos = IIf(IsNull(rstCtacte!Gastos), "0", rstCtacte!Gastos)
            WDescuento = IIf(IsNull(rstCtacte!Descuento), "0", rstCtacte!Descuento)
            Seguro.Text = Str$(WSeguro)
            Seguro.Text = Pusing("###,###.##", Seguro.Text)
            Flete.Text = Str$(WFlete)
            Flete.Text = Pusing("###,###.##", Flete.Text)
            Gastos.Text = Str$(WGastos)
            Gastos.Text = Pusing("###,###.##", Gastos.Text)
            Descuento.Text = Str$(WDescuento)
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
            
            Cae.Text = IIf(IsNull(rstCtacte!Cae), "", rstCtacte!Cae)
            Marca.Text = IIf(IsNull(rstCtacte!Marca), "", rstCtacte!Marca)
            Envio1.Text = IIf(IsNull(rstCtacte!Envio1), "", rstCtacte!Envio1)
            Envio2.Text = IIf(IsNull(rstCtacte!Envio2), "", rstCtacte!Envio2)
            Pago1.Text = IIf(IsNull(rstCtacte!Pago1), "", rstCtacte!Pago1)
            Pago2.Text = IIf(IsNull(rstCtacte!Pago2), "", rstCtacte!Pago2)
            NroOrden.Text = IIf(IsNull(rstCtacte!NroOrden), "", rstCtacte!NroOrden)
            fecorden.Text = IIf(IsNull(rstCtacte!fecorden), "  /  /    ", rstCtacte!fecorden)
            Consignatario.Text = IIf(IsNull(rstCtacte!Consignatario), "", rstCtacte!Consignatario)
            Dolar1.Text = IIf(IsNull(rstCtacte!ImpreDolar1), "", rstCtacte!ImpreDolar1)
            Dolar2.Text = IIf(IsNull(rstCtacte!ImpreDolar2), "", rstCtacte!ImpreDolar2)
            
            WCipLista = IIf(IsNull(rstCtacte!CipLista), "0", rstCtacte!CipLista)
            CipLista.ListIndex = WCipLista
            WIdioma = IIf(IsNull(rstCtacte!Idioma), "0", rstCtacte!Idioma)
            Idioma.ListIndex = WIdioma
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!vendedor
                WProv = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            
            Call Proceso1_Click
            
                    Else
                    
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Facturas")
            Rem     Numero.SetFocus
            Rem        Else
            Rem     WNumero = Numero.Text
            Rem    Rem Call Limpia_Click
            Rem    Numero.Text = WNumero
            Rem    Pedido.SetFocus
            Rem End If
            WNumero = Numero.Text
            Rem Call Limpia_Click
            Numero.Text = WNumero
            Pedido.SetFocus
                
        End If
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            Select Case rstPedido!TipoPedido
                Case 1
                    WTipoPedido = "CO"
                Case 3
                    WTipoPedido = "BI"
                Case 4
                    WTipoPedido = "FA"
                Case 5
                    WTipoPedido = "PG"
                Case Else
                    WTipoPedido = "PT"
            End Select
            rstPedido.Close
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!vendedor
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WProv = rstCliente!Provincia
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            Call Calcula_FechaVto
            Call Proceso_Click
            Fecha.SetFocus
        End If
    End If
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
                rstCambios.Close
                         Else
                 Paridad.Text = ""
            End If
            Rem Paridad.Text = "1"
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Remito.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub REImpreIII_Click()
    ZZDestino = 0
    Call Impresion_Varios
End Sub

Private Sub ReImpresion_Click()
    Call Impresion_Remito
End Sub

Private Sub ReImpresionII_Click()
    Call Impresion_FE
End Sub

Private Sub Seguro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Seguro.Text = Pusing("###,###.##", Seguro.Text)
        Call Calcula_Click
    End If
End Sub

Private Sub Flete_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Flete.Text = Pusing("###,###.##", Flete.Text)
        Call Calcula_Click
    End If
End Sub

Private Sub Gastos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Gastos.Text = Pusing("###,###.##", Gastos.Text)
        Call Calcula_Click
    End If
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento.Text = Pusing("###,###.##", Descuento.Text)
        Call Calcula_Click
    End If
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
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Orden.SetFocus
    End If
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion()

        Rem Open "LPT1" For Output As #99
        Open "dada.txt" For Output As #99

        Print #99, Chr$(27) + Chr$(40) + "19U";
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72)
        Print #99, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72)

        For XX = 1 To 1

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(18); Left$(fecorden.Text, 2);
        Print #99, Tab(21); Mid$(fecorden.Text, 4, 2);
        Print #99, Tab(24); Right$(fecorden.Text, 2);
        Print #99, Tab(27); Left$(NroOrden.Text, 6);
        Print #99, Tab(37); Consignatario.Text;
        Print #99, Tab(68); Left$(Fecha.Text, 2);
        Print #99, Tab(71); Mid$(Fecha.Text, 4, 2);
        Print #99, Tab(74); Right$(Fecha.Text, 2)
        Print #99, ""
        Print #99, ""

        Print #99, Tab(45); Envio1.Text
        Print #99, Tab(3); Left$(WRazon, 40);
        Print #99, Tab(45); Envio2.Text

        Print #99, ""
        Print #99, Tab(3); Left$(WDireccion, 40);
        Print #99, Tab(45); Pago1.Text
        Print #99, Tab(3); Left$(WLocalidad, 40);
        Print #99, Tab(45); Pago2.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(85); "USD"
        Print #99, ""
        Print #99, ""

        Suma1 = 0
        Suma2 = 0
        Suma3 = 0
        Erase WImpresion
        WRenglon = 0
        
        Impre = 0
        
        For a = 0 To 8
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                XProducto = DBGrid1.Text
                
                DBGrid1.Col = 1
                XDescri = DBGrid1.Text
                
                DBGrid1.Col = 3
                XPrecio = DBGrid1.Text
            
                DBGrid1.Col = 4
                XCantidad = DBGrid1.Text
                
                DBGrid1.Col = 5
                XCantidad1 = DBGrid1.Text
                
                DBGrid1.Col = 6
                XTipo = DBGrid1.Text
                
                DBGrid1.Col = 7
                XNumero = DBGrid1.Text
                
                DBGrid1.Col = 8
                XBruto = DBGrid1.Text
                
                WRenglon = WRenglon + 1
                
                WImpresion(WRenglon, 1) = XProducto
                WImpresion(WRenglon, 2) = XDescri
                WImpresion(WRenglon, 3) = ""
                WImpresion(WRenglon, 4) = XPrecio
                WImpresion(WRenglon, 5) = XCantidad
                WImpresion(WRenglon, 6) = XCantidad1
                WImpresion(WRenglon, 7) = XTipo
                WImpresion(WRenglon, 8) = XNumero
                WImpresion(WRenglon, 9) = XBruto
                    
            Next iRow
            
        Next a
        
        XPasa = 0
        
        
        For a = 1 To 99
        
                Producto = WImpresion(a, 1)
                Descri = WImpresion(a, 2)
                Precio = Val(Alinea("##,###.##", WImpresion(a, 4)))
                Cantidad = Val(WImpresion(a, 5))
                Cantidad1 = Val(WImpresion(a, 6))
                WTipo = WImpresion(a, 7)
                WNumero = WImpresion(a, 8)
                Bruto = Val(WImpresion(a, 9))
                    
                If Cantidad <> 0 Then
                
                        If WNumero = WImpresion(a + 1, 8) And XPasa = 0 Then
                        
                            Print #99, Tab(2); Alinea("###", Str$(1));
                            Print #99, Tab(8); WTipo;
                            Print #99, Tab(12); WNumero;
                            Print #99, Tab(22); "Palet conteniendo : ";
                            Print #99, Tab(60); Alinea("#####.##", Str$(24));
                            Suma2 = Suma2 + 24
                            
                            XPasa = 1
                            
                            XCanti = XEnvase(a, 2)
                            Call Ceros(XCanti, 2)
                            spEnvase = "ConsultaEnvases " + "'" + XEnvase(a, 1) + "'"
                            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                            If rstEnvase.RecordCount > 0 Then
                                WEnva = "(" + XCanti + "X" + Left$(rstEnvase!Descripcion, 10) + ")"
                                rstEnvase.Close
                                    Else
                                WEnva = ""
                            End If
                            
                            Print #99, Tab(22); Left$(Descri, 21); " "; WEnva;
                            Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                            Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                            Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                            Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                            
                            Suma1 = Suma1 + Cantidad1
                            Suma2 = Suma2 + Bruto
                            Suma3 = Suma3 + Cantidad
                            
                                    Else
                                    
                            If WNumero <> WImpresion(a + 1, 8) Then
                                XPasa = 0
                            End If
                                
                            If WNumero = WImpresion(a - 1, 8) Then
                            
                                XCanti = XEnvase(a, 2)
                                Call Ceros(XCanti, 2)
                                spEnvase = "ConsultaEnvases " + "'" + XEnvase(a, 1) + "'"
                                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                If rstEnvase.RecordCount > 0 Then
                                    WEnva = "(" + XCanti + "X" + Left$(rstEnvase!Descripcion, 10) + ")"
                                    rstEnvase.Close
                                        Else
                                    WEnva = ""
                                End If
                            
                                Print #99, Tab(22); Left$(Descri, 21); " "; WEnva;
                                Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                                Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                                Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                                Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                                    Else
                                Print #99, Tab(2); Alinea("###", Str$(Cantidad1));
                                Print #99, Tab(8); WTipo;
                                Print #99, Tab(12); WNumero;
                                Print #99, Tab(22); Left$(Descri, 37);
                                Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                                Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                                Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                                Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                            End If
                                
                            Suma1 = Suma1 + Cantidad1
                            Suma2 = Suma2 + Bruto
                            Suma3 = Suma3 + Cantidad
                            
                        End If
                
                        Impre = Impre + 1
                    
                End If
            
        Next a
        
        
        
        For DA = Impre To 23
            Print #99, ""
        Next DA

        Print #99, Tab(5); "Todas las disputas que puedan surgir en el presente contrato seran finalmente arregladas"
        Print #99, Tab(5); "de acuerdo a las Reglas de Conciliacion y Arbitraje  de  la  Camara  Internacional   de"
        Print #99, Tab(5); "Comercio por uno o mas arbitros de acuerdo de dichas reglas"
        Print #99, Tab(5); "INCOTERMS 1990";
        
        Call Numtolet
        
        WTexto1 = UCase(WTexto1)
        WTexto2 = UCase(WTexto2)

        Print #99, Tab(22); "Son Dolares estadounidenses"
        Print #99, Tab(20); WTexto1

        Print #99, Tab(2); Alinea("###", Str$(Suma1));
        Print #99, Tab(20); WTexto2;
        Print #99, Tab(60); Alinea("#####.#", Str$(Suma2));
        Print #99, Tab(68); Alinea("#####", Str$(Suma3));
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        If Val(Seguro.Text) <> 0 Then
                Print #99, Tab(83); Alinea("###,###.##", Seguro.Text)
                        Else
                Print #99, ""
        End If
        Print #99, ""
        If Val(Flete.Text) <> 0 Then
                Print #99, Tab(83); Alinea("###,###.##", Flete.Text)
                        Else
                Print #99, ""
        End If
        Print #99, Tab(6); Marca.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(83); Alinea("###,###.##", Total.Caption)

        Next XX
        
        Close #99
End Sub
        
Sub Impresion_Remito()

    m$ = "Coloque el remito"
    G% = MsgBox(m$, 0, "Impresion de remito")

    Impre = 0
    ZSql = "DELETE ImpreRemito"
    spImpreRemito = ZSql
    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
    
    For a = 0 To 8
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
            
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 0
            Producto = DBGrid1.Text
            
            DBGrid1.Col = 1
            Descri = DBGrid1.Text
            
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
        
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            If Cantidad <> 0 Then
            
                Impre = Impre + 1
                
                Auxi1 = Numero.Text
                Call Ceros(Auxi1, 6)
                Auxi2 = Impre
                Call Ceros(Auxi2, 2)
                    
                ZZClave = Auxi1 + Auxi2
                ZZNumero = Numero.Text
                ZZRenglon = Str$(Impre)
                ZZFecha = Fecha.Text
                ZZNombre = WRazon
                ZZDireccion = WDireccion
                ZZLocalidad = WLocalidad
                ZZPedido = Pedido.Text
                ZZCliente = Cliente.Text
                ZZOrden = Orden.Text
                ZZDescripcion = Left$(Descri, 50)
                ZZCantidad = Str$(Cantidad)
                ZZRemito = Remito.Text
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ImpreRemito ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Nombre ,"
                ZSql = ZSql + "Direccion ,"
                ZSql = ZSql + "Localidad ,"
                ZSql = ZSql + "Pedido ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZClave + "',"
                ZSql = ZSql + "'" + ZZNumero + "',"
                ZSql = ZSql + "'" + ZZRenglon + "',"
                ZSql = ZSql + "'" + ZZFecha + "',"
                ZSql = ZSql + "'" + ZZNombre + "',"
                ZSql = ZSql + "'" + ZZDireccion + "',"
                ZSql = ZSql + "'" + ZZLocalidad + "',"
                ZSql = ZSql + "'" + ZZPedido + "',"
                ZSql = ZSql + "'" + ZZCliente + "',"
                ZSql = ZSql + "'" + ZZOrden + "',"
                ZSql = ZSql + "'" + ZZDescripcion + "',"
                ZSql = ZSql + "'" + ZZCantidad + "',"
                ZSql = ZSql + "'" + ZZRemito + "')"
                
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
                
        Next iRow
        
    Next a
    
    For aa = Impre To 22
        
        Auxi1 = Numero.Text
        Call Ceros(Auxi1, 6)
        Auxi2 = aa
        Call Ceros(Auxi2, 2)
            
        ZZClave = Auxi1 + Auxi2
        ZZNumero = Numero.Text
        ZZRenglon = Str$(aa)
        ZZFecha = Fecha.Text
        ZZNombre = WRazon
        ZZDireccion = WDireccion
        ZZLocalidad = WLocalidad
        ZZPedido = Pedido.Text
        ZZCliente = Cliente.Text
        ZZOrden = Orden.Text
        ZZDescripcion = ""
        ZZCantidad = ""
        ZZRemito = Remito.Text
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreRemito ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Pedido ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Remito )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZFecha + "',"
        ZSql = ZSql + "'" + ZZNombre + "',"
        ZSql = ZSql + "'" + ZZDireccion + "',"
        ZSql = ZSql + "'" + ZZLocalidad + "',"
        ZSql = ZSql + "'" + ZZPedido + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZOrden + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZRemito + "')"
        
        spImpreRemito = ZSql
        Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
            
    Next aa
        
    Listado.WindowTitle = "Listado de Familia de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{ImpreRemito.Numero} in " + "0" + " to " + "999999"
    Listado.Destination = 1
    
    Listado.ReportFileName = "ImpreRemito.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ImpreRemito.Clave, ImpreRemito.Numero, ImpreRemito.Fecha, ImpreRemito.Nombre, ImpreRemito.Direccion, ImpreRemito.Localidad, ImpreRemito.Pedido, ImpreRemito.Cliente, ImpreRemito.Descripcion, ImpreRemito.Cantidad, ImpreRemito.Remito " _
            + "From " _
            + DSQ + ".dbo.ImpreRemito ImpreRemito " _
            + "Where " _
            + "ImpreRemito.Numero >= 0 AND " _
            + "ImpreRemito.Numero <= 999999"
    
    Listado.Connect = Connect()
    Listado.CopiesToPrinter = 2
    
    Listado.Destination = 1
    
    Listado.Action = 1

End Sub



Sub Impresion_FE()

    Call Calcula_Barra
        
    ZZImpreTotalNeto = 0
    For a = 0 To 8
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
    
        For iRow = 0 To 9
        
            WRenglon = WRenglon + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            ZZImpreTotalNeto = ZZImpreTotalNeto + Cantidad
            
        Next iRow
    Next a
        
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = "01" + Auxi + "01"
        
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "ImpreTotalNeto= " + "'" + Str$(ZZImpreTotalNeto) + "'"
    ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
    Listado.WindowTitle = "Factura Electronica"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Estadistica.Numero} in " + Numero.Text + " to " + Numero.Text
    Listado.Destination = 1
    
    Select Case Val(Wempresa)
        Case 1
            If IdiomaII.ListIndex = 1 Then
                Listado.ReportFileName = "ImpreFacturaExpoDolar.rpt"
                    Else
                Listado.ReportFileName = "ImpreFacturaExpo.rpt"
            End If
        Case Else
            Listado.ReportFileName = "ImpreFacturaExpoPelli.rpt"
    End Select
    Rem Listado.ReportFileName = "ImpreFacturaExpocristalia.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Estadistica.Numero, Estadistica.Cantidad, Estadistica.PrecioUs, Estadistica.ImporteUs, Estadistica.ImpreTerminado, Estadistica.ImpreCantidad, Estadistica.ImpreTipo, Estadistica.ImpreNumeros, Estadistica.ImpreBruto, " _
            + "CtaCte.fecha, CtaCte.TotalUs, CtaCte.Seguro, CtaCte.Flete, CtaCte.ImpreNumero, CtaCte.Cae, CtaCte.FechaCae, CtaCte.Marca, CtaCte.Envio1, CtaCte.Envio2, CtaCte.Pago1, CtaCte.Pago2, CtaCte.NroOrden, CtaCte.FecOrden, CtaCte.Consignatario, CtaCte.Cip, CtaCte.ImpreDolar1, CtaCte.ImpreDolar2, CtaCte.ImpreTotal, CtaCte.ImpreTotalBruto, CtaCte.ImpreTotalNeto, CtaCte.Gastos, CtaCte.ImpreBarra, CtaCte.ImpreBarraII, CtaCte.Descuento  " _
            + "Cliente.Razon, Cliente.Direccion, Cliente.Localidad " _
            + "From " _
            + DSQ + ".dbo.Estadistica Estadistica, " _
            + DSQ + ".dbo.CtaCte CtaCte, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Estadistica.ClaveCtaCte = CtaCte.Clave AND " _
            + "CtaCte.Cliente = Cliente.Cliente AND " _
            + "Estadistica.Numero >= " + Numero.Text + " AND " _
            + "Estadistica.Numero <= " + Numero.Text
    
    Listado.Connect = Connect()
    Listado.CopiesToPrinter = 2
    
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.Action = 1

End Sub



Private Sub Marca_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envio1.SetFocus
    End If
End Sub

Private Sub Envio1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envio2.SetFocus
    End If
End Sub

Private Sub Envio2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pago1.SetFocus
    End If
End Sub

Private Sub Pago1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pago2.SetFocus
    End If
End Sub

Private Sub Pago2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroOrden.SetFocus
    End If
End Sub

Private Sub NroOrden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fecorden.SetFocus
    End If
End Sub

Private Sub Fecorden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Consignatario.SetFocus
    End If
End Sub

Private Sub Consignatario_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dolar1.SetFocus
    End If
End Sub

Private Sub Dolar1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dolar2.SetFocus
    End If
End Sub

Private Sub Dolar2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Marca.SetFocus
    End If
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
    
    Dim WSAA As Object, WSFEX As Object
    Dim dst_cmp  As Integer
    
    
    
    On Error GoTo ManejoError
    
    
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEXv1
    tra = WSAA.CreateTRA("wsfex")
    Debug.Print tra
    
    
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    ZPath = "c:\salva\"
    
    Select Case Val(Wempresa)
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
    Set WSFEX = CreateObject("WSFEX")
   
    
    
    ' Setear tocken y sing de autorización (pasos previos)
    WSFEX.Token = WSAA.Token
    WSFEX.Sign = WSAA.Sign
    
    
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEX.Cuit = ZCuit
    
    
    
    ' Conectar al Servicio Web de Facturación
    ok = WSFEX.Conectar("https://servicios1.afip.gov.ar/WSFEX/service.asmx") ' homologación
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEX.Dummy
    Debug.Print "appserver status", WSFEX.AppServerStatus
    Debug.Print "dbserver status", WSFEX.DbServerStatus
    Debug.Print "authserver status", WSFEX.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = 19 ' FC Expo (ver tabla de parámetros)
    Select Case Val(Wempresa)
        Case 1
            punto_vta = 6
        Case Else
            punto_vta = 3
    End Select
    
    
    ' Obtengo el último número de comprobante y le agrego 1
    
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    
    
    Cbte_Nro = WSFEX.GetLastCMP(tipo_cbte, punto_vta) + 1 '16
    ZZComprobante = Cbte_Nro
    
    
    fecha_cbte = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    tipo_expo = 1 ' tipo de exportación (ver tabla de parámetros)
    permiso_existente = "N"
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
    forma_pago = Pago1.Text
    incoterms = CipLista.Text  ' (ver tabla de parámetros)
    incoterms_ds = ""
    idioma_cbte = Idioma.ListIndex  ' (ver tabla de parámetros)
    IMP_TOTAL = Total.Caption
   
    ' Creo una factura (internamente, no se llama al WebService):
    Rem ok = WSFEXv1.CrearFactura(tipo_cbte, punto_vta, Cbte_Nro, fecha_cbte, _
    REM         IMP_TOTAL, tipo_expo, permiso_existente, dst_cmp, _
    REM         XXCliente, cuit_pais_cliente, domicilio_cliente, _
    REM         id_impositivo, moneda_id, moneda_ctz, _
    REM         obs_comerciales, obs, forma_pago, incoterms, _
    REM         idioma_cbte, incoterms_ds)
    
    ' Creo una factura (internamente, no se llama al WebService):
    ok = WSFEX.CrearFactura(tipo_cbte, punto_vta, Cbte_Nro, fecha_cbte, _
            IMP_TOTAL, tipo_expo, permiso_existente, dst_cmp, _
            XXCliente, cuit_pais_cliente, domicilio_cliente, _
            id_impositivo, moneda_id, moneda_ctz, _
            obs_comerciales, obs, forma_pago, incoterms, _
            idioma_cbte)
    
    
    
    
    ' Agrego un item:
    
    For ZZCiclo = 1 To 99
    
        ZZArticulo = ZZVector(ZZCiclo, 1)
        ZZCantidad = ZZVector(ZZCiclo, 2)
        ZZPrecio = ZZVector(ZZCiclo, 3)
        
        If Trim(ZZArticulo) <> "" Then
    
            If Left$(ZZArticulo, 2) = "PT" Or Left$(ZZArticulo, 2) = "PE" Then
                ClavePrecios = Cliente.Text + ZZArticulo
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    ZZDescripcion = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                        Else
                WArti = Left$(ZZArticulo, 3) + Right$(ZZArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            End If
    
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            Rem If Val(IMP_TOTAL) <> 0 Then
                ' lo agrego a la factura (internamente, no se llama al WebService):
                Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            Rem End If
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
            
        End If
        
    Next ZZCiclo
    
    If Val(Wempresa) <> 1 Then
    
        If Val(Flete.Text) <> 0 Then
    
            ZZArticulo = "Flete"
            ZZCantidad = "1"
            ZZPrecio = Flete.Text
            ZZDescripcion = "Flete"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
        
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
        
        End If
        
        If Val(Seguro.Text) <> 0 Then
    
            ZZArticulo = "Seguro"
            ZZCantidad = "1"
            ZZPrecio = Seguro.Text
            ZZDescripcion = "Seguro"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        If Val(Gastos.Text) <> 0 Then
    
            ZZArticulo = "Gastos"
            ZZCantidad = "1"
            ZZPrecio = Gastos.Text
            ZZDescripcion = "Gastos"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        If Val(Descuento.Text) <> 0 Then
    
            ZZArticulo = "Dto"
            ZZCantidad = "1"
            ZZPrecio = Str$(Val(Gastos.Text) * -1)
            ZZDescripcion = "Descuento"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        
    End If
    
    
    
    
    
    Rem OJO
    Rem If Val(WEmpresa) = 1 Then
    Rem
    Rem     If Val(Gastos.Text) <> 0 Then
    Rem
    Rem         ZZArticulo = "Gastos"
    Rem         ZZCantidad = "1"
    Rem         ZZPrecio = Gastos.Text
    Rem         ZZDescripcion = "Gastos"
    Rem
    Rem         XXCodigo = ZZArticulo
    Rem         XXDs = ZZDescripcion
    Rem         qty = ZZCantidad
    Rem         XXPrecio = ZZPrecio
    Rem         umed = 1 ' Ver tabla de parámetros (unidades de medida)
    Rem         IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
    Rem         Bonif = ""
    Rem
    Rem         ' lo agrego a la factura (internamente, no se llama al WebService):
    Rem         ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL, Bonif)
    Rem     End If
    Rem
    Rem End If
    
    
    
    
    
    
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
    Rem id = CStr(CCur(WSFEXv1.GetLastID()) + 1)
    
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Rem Cae = WSFEXv1.Authorize(id)
    Rem Debug.Print WSFEXv1.XmlRequest
    Rem Debug.Print WSFEXv1.XmlResponse
    Rem Cae.Text = Cae
        
        
        
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    Rem If Cae = "" Or WSFEXv1.Resultado <> "A" Then
    Rem     MsgBox "No se asignó CAE (Rechazado). Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    Rem ElseIf WSFEXv1.obs <> "" And WSFEXv1.obs <> "00" Then
    Rem     MsgBox "Se asignó CAE pero con advertencias. Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    Rem End If
    
    
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Rem Debug.Print WSFEXv1.XmlRequest
    Rem Debug.Print WSFEXv1.XmlResponse
    
    Rem MsgBox "Resultado:" & WSFEXv1.Resultado & " CAE: " & Cae & " Reproceso: " & WSFEXv1.Reproceso & " Obs: " & WSFEXv1.obs & " Nro: " & ZZComprobante, vbInformation + vbOKOnly
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    Rem For Each evento In WSFEXv1.Eventos
    Rem     If evento <> "0: " Then
    Rem         MsgBox "Evento: " & evento, vbInformation
    Rem     End If
    Rem Next
    
    ' Buscar la factura
    Rem cae2 = WSFEXv1.GetCMP(tipo_cbte, punto_vta, Cbte_Nro)
    
    Rem Debug.Print "Fecha Comprobante:", WSFEXv1.FechaCbte
    Rem Debug.Print "Importe Total:", WSFEXv1.ImpTotal
    
    Rem If Cae <> cae2 Then
    Rem     MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!"
    Rem         Else
    Rem     MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    Rem     ZZGrabaFactura = "S"
    Rem End If
    
    
    
    
    
    
    
    
    'id = "99000000000100" ' número propio de transacción
    ' obtengo el último ID y le adiciono 1 (advertencia: evitar overflow!)
    id = CStr(CCur(WSFEX.GetLastID()) + 1)
    
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Cae = WSFEX.Authorize(id)
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    Cae.Text = Cae
        
        
        
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Or WSFEX.Resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Observación (motivos): " & WSFEX.obs, vbInformation + vbOKOnly
    ElseIf WSFEX.obs <> "" And WSFEX.obs <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Observación (motivos): " & WSFEX.obs, vbInformation + vbOKOnly
    End If
    
    
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    
    MsgBox "Resultado:" & WSFEX.Resultado & " CAE: " & Cae & " Reproceso: " & WSFEX.Reproceso & " Obs: " & WSFEX.obs & " Nro: " & ZZComprobante, vbInformation + vbOKOnly
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEX.Eventos
        If evento <> "0: " Then
            MsgBox "Evento: " & evento, vbInformation
        End If
    Next
    
    ' Buscar la factura
    cae2 = WSFEX.GetCMP(tipo_cbte, punto_vta, Cbte_Nro)
    
    Debug.Print "Fecha Comprobante:", WSFEX.FechaCbte
    Debug.Print "Importe Total:", WSFEX.ImpTotal
    
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
    Select Case Val(Wempresa)
        Case 1
            ZZNumero = "30549165083"
        Case Else
            ZZNumero = "30610524598"
    End Select
    
    ZZNumero = ZZNumero + "19"
    
    Select Case Val(Wempresa)
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
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    ZZImpreNumero = "0000" + Right$(Auxi, 4)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE CtaCte SET "
    ZSql = ZSql & "ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql & "FechaCae = " + "'" + ZZFechaCae + "',"
    ZSql = ZSql & "ImpreBarra = " + "'" + barralargo + "',"
    ZSql = ZSql & "ImpreBarraII = " + "'" + ZZNumero + "'"
    ZSql = ZSql & " Where Tipo = " + "'" + "01" + "'"
    ZSql = ZSql & " and Numero = " + "'" + Numero.Text + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

End Sub







Private Sub Impresion_Varios()

    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        WIdioma = IIf(IsNull(rstClientes!Idioma), "0", rstClientes!Idioma)
        rstClientes.Close
    End If
    
    Call Conecta_Empresa


    ZZVersion = 0
    FileCopy "W:\base.pdf", "c:\pdfprint\base.pdf"
    Kill "C:\pdfprint\*.pdf"
    ZZImprePdf = "N"
    
    Erase ZImpreFicha
    ZLugarFicha = 0
    
    For DA = 1 To 99
    
        Articulo = AuxiliarII(DA, 1)
        Cantidad = AuxiliarII(DA, 2)
        Precio = AuxiliarII(DA, 3)
        WRenglon = AuxiliarII(DA, 4)
        WLote1 = AuxiliarII(DA, 5)
        WCanti1 = AuxiliarII(DA, 6)
        WLote2 = AuxiliarII(DA, 7)
        WCanti2 = AuxiliarII(DA, 8)
        Wlote3 = AuxiliarII(DA, 9)
        WCanti3 = AuxiliarII(DA, 10)
        WLote4 = AuxiliarII(DA, 11)
        WCanti4 = AuxiliarII(DA, 12)
        WLote5 = AuxiliarII(DA, 13)
        WCanti5 = AuxiliarII(DA, 14)
        WLote6 = AuxiliarII(DA, 15)
        WCanti6 = AuxiliarII(DA, 16)
        WLote7 = AuxiliarII(DA, 17)
        WCanti7 = AuxiliarII(DA, 18)
        WLote8 = AuxiliarII(DA, 19)
        WCanti8 = AuxiliarII(DA, 20)
        WLote9 = AuxiliarII(DA, 21)
        WCanti9 = AuxiliarII(DA, 22)
        WLote10 = AuxiliarII(DA, 23)
        WCanti10 = AuxiliarII(DA, 24)
        WLote11 = AuxiliarII(DA, 25)
        WCanti11 = AuxiliarII(DA, 26)
        WLote12 = AuxiliarII(DA, 27)
        WCanti12 = AuxiliarII(DA, 28)
        ZZDescriArticulo = Trim(AuxiliarII(DA, 30))
        
        If Articulo = "" Then Exit For
    
        Rem
        Rem Hoja de Seguridad
        Rem
        
    
        Rem
        Rem Hoja de Seguridad
        Rem
        
        
        
        
        
        
        XTipoPro = "CO"
        If Left$(UCase(Articulo), 2) = "PT" Then
        
            XCodigo = Val(Mid$(Articulo, 4, 5))
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        XTipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            XTipoPro = "BI"
                                Else
                            If XCodigo >= 40000 And XCodigo <= 41000 Then
                                XTipoPro = "TA"
                                    Else
                                XTipoPro = "PT"
                            End If
                        End If
                    End If
                End If
            End If
        
            If Left$(Articulo, 2) = "YQ" Then
                XTipoPro = "PT"
            End If
            If Left$(Articulo, 2) = "YH" Then
                XTipoPro = "PT"
            End If
            If Left$(Articulo, 2) = "YP" Then
                XTipoPro = "PT"
            End If
            If Left$(Articulo, 2) = "YF" Then
                XTipoPro = "FA"
            End If
    
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZLinea = rstTerminado!Linea
                rstTerminado.Close
            End If
    
            Select Case ZLinea
                Case 8
                    XTipoPro = "PG"
                Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                    XTipoPro = "FA"
                Case Else
            End Select
        
        End If
        
        Select Case XTipoPro
            Case "CO", "FA"
                ZZZZProceso = 1
        
            Case Else
                ZZZZProceso = 0
                
        End Select
    
        ListaSga(1) = "PT-01903-100"
        ListaSga(2) = "PT-01921-100"
        ListaSga(3) = "PT-01922-100"
        ListaSga(4) = "PT-01923-100"
        ListaSga(5) = "PT-01960-075"
        ListaSga(6) = "PT-02009-100"
        ListaSga(7) = "PT-02025-100"
        ListaSga(8) = "PT-02029-078"
        ListaSga(9) = "PT-02029-083"
        ListaSga(10) = "PT-02049-100"
        ListaSga(11) = "PT-02051-100"
        ListaSga(12) = "PT-02587-200"
        ListaSga(13) = "PT-02600-100"
        ListaSga(14) = "PT-04170-100"
        ListaSga(15) = "PT-06304-100"
        ListaSga(16) = "PT-06600-100"
        ListaSga(17) = "PT-30410-100"
        ListaSga(18) = "PT-30516-100"
        
        For Ciclo = 1 To 1000
            If ListaSga(Ciclo) = UCase(Articulo) Then
                ZZZZProceso = 1
                Exit For
                    Else
                If ListaSga(Ciclo) = "" Then
                    Exit For
                End If
            End If
        Next Ciclo
        
        
        
        
        ZZRequiereCertificado = 0
        ZZRequiereMsds = 0
        ZZRequiereMsdsCada = 0
        ZZRequiereHoja = 0
        ZZBusqueda = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteEspecif"
        ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + Cliente.Text + "'"
        spClienteEspecif = ZSql
        Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteEspecif.RecordCount > 0 Then
            ZZRequiereCertificado = IIf(IsNull(rstClienteEspecif!RequiereCertificado), "0", rstClienteEspecif!RequiereCertificado)
            ZZRequiereMsds = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
            ZZRequiereMsdsCada = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
            ZZRequiereHoja = IIf(IsNull(rstClienteEspecif!RequiereHoja), "0", rstClienteEspecif!RequiereHoja)
            rstClienteEspecif.Close
        End If
        
        If ZZRequiereMsdsCada = 1 Then
            ZZBusqueda = "S"
                Else
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Cliente = " + "'" + Cliente.Text + "'"
            ZSql = ZSql + " and Estadistica.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " and Estadistica.Numero <> " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                rstEstadistica.Close
                    Else
                ZZBusqueda = "S"
            End If
        End If
        
        If ZZBusqueda = "S" Then
            
            If Left$(UCase(Articulo), 2) = "PT" Or Left$(UCase(Articulo), 2) = "PE" Then

                Es = ZZDescriArticulo
                x = ""
                For XX = 1 To Len(Es)
                    Y = Mid$(Es, XX, 1)
                    If Y <> " " And Y <> "/" Then
                        x = x + Y
                    End If
                Next
                ZZCodArt = x + Mid$(Articulo, 4, 5) + Right$(Articulo, 3)
                ZZCodArtII = ZZCodArt
                If WIdioma = 1 Then
                    ZZCodArt = ZZCodArt + "ing"
                End If
                
                    Else
                    
                ZZCodArt = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                ZZCodArtII = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                
            End If
            
            If ZZZZProceso = 1 Then
                ZZRuta = "w:\fds\fds" + ZZCodArt + ".PDF"
                    Else
                ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArt + ".PDF"
            End If
            ZZRutaCopia = "w:\MSDS EXPO\" + Numero.Text + " - MSDS" + ZZCodArt + ".PDF"
            ZZEstado = Dir(ZZRuta)
            ZZEstado = Trim(ZZEstado)
            If ZZEstado <> "" Then
                
                ZZNombreArchi = 1
                Do
                    Auxi = Str$(ZZNombreArchi)
                    Call Ceros(Auxi, 8)
                    ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                    
                    ZZRutaII = ZZNombreArchiII
                    ZZEstadoII = Dir(ZZRutaII)
                    ZZEstadoII = Trim(ZZEstadoII)
                    If ZZEstadoII = "" Then
                        Exit Do
                    End If
                    ZZNombreArchi = ZZNombreArchi + 1
                Loop
                
                FileCopy ZZRuta, ZZRutaII
                FileCopy ZZRuta, ZZRutaCopia
                Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                ZZImprePdf = "S"
                
                Rem TiempoPausa = 2 ' Asigna hora de inicio.
                Rem Inicio = Timer  ' Establece la hora de inicio.
                Rem Do While Timer < Inicio + TiempoPausa
                Rem     DoEvents    ' Cambia a otros procesos.
                Rem Loop
            
                Rem Select Case ZZVersion
                Rem     Case 1
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem     Case 2
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem     Case 3
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem     Case 4
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem     Case 5
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem     Case Else
                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Rem End Select
                
                Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                
                    Else
                    
                If ZZZZProceso = 1 Then
                    ZZRuta = "w:\fds\fds" + ZZCodArtII + ".PDF"
                        Else
                    ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArtII + ".PDF"
                End If
                ZZRutaCopia = "w:\MSDS EXPO\" + Numero.Text + " - MSDS" + ZZCodArtII + ".PDF"
                ZZEstado = Dir(ZZRuta)
                ZZEstado = Trim(ZZEstado)
                If ZZEstado <> "" Then
                    
                    ZZNombreArchi = 1
                    Do
                        Auxi = Str$(ZZNombreArchi)
                        Call Ceros(Auxi, 8)
                        ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                        
                        ZZRutaII = ZZNombreArchiII
                        ZZEstadoII = Dir(ZZRutaII)
                        ZZEstadoII = Trim(ZZEstadoII)
                        If ZZEstadoII = "" Then
                            Exit Do
                        End If
                        ZZNombreArchi = ZZNombreArchi + 1
                    Loop
                    
                    FileCopy ZZRuta, ZZRutaII
                    FileCopy ZZRuta, ZZRutaCopia
                
                    Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                    Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                    ZZImprePdf = "S"
                    
                    Rem TiempoPausa = 2 ' Asigna hora de inicio.
                    Rem Inicio = Timer  ' Establece la hora de inicio.
                    Rem Do While Timer < Inicio + TiempoPausa
                    Rem     DoEvents    ' Cambia a otros procesos.
                    Rem Loop
                
                    Rem Select Case ZZVersion
                    Rem     Case 1
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem     Case 2
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem     Case 3
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem     Case 4
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem     Case 5
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem     Case Else
                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Rem End Select
                    
                    Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                    
                        Else
                        
                    ZZAviso = "S"
                    If Left$(Articulo, 2) = "YQ" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "YH" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "YP" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "YF" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "ML" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "QC" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "ZE" Then
                        ZZAviso = "N"
                    End If
                    If Left$(Articulo, 2) = "ZT" Then
                        ZZAviso = "N"
                    End If
                        
                    If ZZZZProceso = 1 Then
                        If ZZAviso = "S" Then
                            m$ = "El FDS  (" + ZZCodArt + ")  no se ha encontrado"
                            a% = MsgBox(m$, 0, "Impresion de comprobantes varios")
                        End If
                            Else
                        If ZZAviso = "S" Then
                            m$ = "El MSDS  (" + ZZCodArt + ")  no se ha encontrado"
                            a% = MsgBox(m$, 0, "Impresion de comprobantes varios")
                        End If
                    End If
                
                End If
                
            End If
            
        End If
    
                    
        Rem
        Rem ficha de intevencion
        Rem
        If Left$(UCase(Articulo), 2) = "PT" Or Left$(UCase(Articulo), 2) = "PE" Then
                
            ZZIntervencion = ""
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                rstTerminado.Close
            End If
            
            If Val(ZZIntervencion) <> 0 Then
            
                ZEntraFicha = "S"
                For ZZCicloFicha = 1 To ZLugarFicha
                    If ZImpreFicha(ZZCicloFicha) = Val(ZZIntervencion) Then
                        ZEntraFicha = "N"
                        Exit For
                    End If
                Next ZZCicloFicha
                If ZEntraFicha = "S" Then
                    ZLugarFicha = ZLugarFicha + 1
                    ZImpreFicha(ZLugarFicha) = ZZIntervencion
                End If
                
            End If
            
                Else
                
            ZZCodArt = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
            ZZIntervencion = ""
            spArticulo = "ConsultaArticulo " + "'" + ZZCodArt + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                rstArticulo.Close
            End If
            
            If Val(ZZIntervencion) <> 0 Then
            
                ZEntraFicha = "S"
                For ZZCicloFicha = 1 To ZLugarFicha
                    If ZImpreFicha(ZZCicloFicha) = Val(ZZIntervencion) Then
                        ZEntraFicha = "N"
                        Exit For
                    End If
                Next ZZCicloFicha
                If ZEntraFicha = "S" Then
                    ZLugarFicha = ZLugarFicha + 1
                    ZImpreFicha(ZLugarFicha) = ZZIntervencion
                End If
                
            End If
                
        End If
                    
            
    
        Rem
        Rem certificado de analisis
        Rem
    
        For ZZCiclo = 1 To 12
            
            Select Case ZZCiclo
                Case 1
                    ZZLugar = 5
                Case 2
                    ZZLugar = 7
                Case 3
                    ZZLugar = 9
                Case 4
                    ZZLugar = 11
                Case 5
                    ZZLugar = 13
                Case 6
                    ZZLugar = 15
                Case 7
                    ZZLugar = 17
                Case 8
                    ZZLugar = 19
                Case 9
                    ZZLugar = 21
                Case 10
                    ZZLugar = 23
                Case 11
                    ZZLugar = 25
                Case Else
                    ZZLugar = 27
            End Select
            
            If Val(AuxiliarII(DA, ZZLugar)) <> 0 Then
        
                ZZEntra = "N"
        
                If Left$(UCase(Articulo), 2) = "PT" Then
                
                    XCodigo = Val(Mid$(Articulo, 4, 5))
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    If XCodigo >= 40000 And XCodigo <= 41000 Then
                                        XTipoPro = "TA"
                                            Else
                                        XTipoPro = "PT"
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    If Left$(Articulo, 2) = "YQ" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YH" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YP" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YF" Then
                        XTipoPro = "FA"
                    End If
                    If Left$(Articulo, 2) = "PE" Then
                        XTipoPro = "PT"
                    End If
            
                    ZLinea = 0
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
            
                    Select Case ZLinea
                        Case 8
                            XTipoPro = "PG"
                        Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                            XTipoPro = "FA"
                        Case Else
                    End Select
            
                    If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                    
                        XEmpresa = Wempresa
                        
                        Select Case Val(Wempresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = AuxiliarII(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = AuxiliarII(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = AuxiliarII(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = AuxiliarII(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                            
                        Erase ZOpcion
                        Erase ZValor
                        Erase ZEnsayo
                        Erase ZStd
                        Erase ZDescri
                        Erase ZDescriII
                            
                        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        ZVersion = 0
                        
                        ZZEntra = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM AltaCertificado"
                        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                        ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                        spAltaCertificado = ZSql
                        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstAltaCertificado.RecordCount > 0 Then
                            ZOpcion(1) = rstAltaCertificado!Opcion1
                            ZOpcion(2) = rstAltaCertificado!Opcion2
                            ZOpcion(3) = rstAltaCertificado!Opcion3
                            ZOpcion(4) = rstAltaCertificado!Opcion4
                            ZOpcion(5) = rstAltaCertificado!Opcion5
                            ZOpcion(6) = rstAltaCertificado!Opcion6
                            ZOpcion(7) = rstAltaCertificado!Opcion7
                            ZOpcion(8) = rstAltaCertificado!Opcion8
                            ZOpcion(9) = rstAltaCertificado!Opcion9
                            ZOpcion(10) = rstAltaCertificado!Opcion10
                            rstAltaCertificado.Close
                            ZZEntra = "S"
                        End If
                        
                        If ZZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                        End If
                        
                        If ZZEntra = "S" Then
                            If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                ZZEntra = "N"
                            End If
                        End If
                                
                        If ZZEntra = "N" Then
                            m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                        End If
                                        
                        If ZZEntra = "S" Then
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    CargaEmpresa(1, 1) = "0001"
                                    CargaEmpresa(1, 2) = "Empresa01"
                                    CargaEmpresa(2, 1) = "0003"
                                    CargaEmpresa(2, 2) = "Empresa03"
                                    CargaEmpresa(3, 1) = "0005"
                                    CargaEmpresa(3, 2) = "Empresa05"
                                    CargaEmpresa(4, 1) = "0006"
                                    CargaEmpresa(4, 2) = "Empresa06"
                                    CargaEmpresa(5, 1) = "0007"
                                    CargaEmpresa(5, 2) = "Empresa07"
                                    CargaEmpresa(6, 1) = "0010"
                                    CargaEmpresa(6, 2) = "Empresa10"
                                    CargaEmpresa(7, 1) = "0011"
                                    CargaEmpresa(7, 2) = "Empresa11"
                                    ZHasta1 = 7
                                Case Else
                                    CargaEmpresa(1, 1) = "0002"
                                    CargaEmpresa(1, 2) = "Empresa02"
                                    CargaEmpresa(2, 1) = "0004"
                                    CargaEmpresa(2, 2) = "Empresa04"
                                    CargaEmpresa(3, 1) = "0008"
                                    CargaEmpresa(3, 2) = "Empresa08"
                                    CargaEmpresa(4, 1) = "0009"
                                    CargaEmpresa(4, 2) = "Empresa09"
                                    ZHasta1 = 4
                            End Select
                                
                            For ZCiclo = 1 To ZHasta1
                                
                                Wempresa = CargaEmpresa(ZCiclo, 1)
                                txtOdbc = CargaEmpresa(ZCiclo, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        
                                If Val(ZLote) > 99999 Then
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 6)
                                        Else
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 5)
                                End If
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Prueter"
                                ZSql = ZSql + " Where Prueter.Lote = " + "'" + ZLote + "'"
                                spPrueter = ZSql
                                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPrueter.RecordCount > 0 Then
                                
                                    If Left$(rstPrueter!prueba, 1) = "2" Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                    If rstPrueter!Producto <> ZProducto Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                        
                                    WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
                                            
                                    ZValor(1) = rstPrueter!Valor1
                                    ZValor(2) = rstPrueter!valor2
                                    ZValor(3) = rstPrueter!Valor3
                                    ZValor(4) = rstPrueter!valor4
                                    ZValor(5) = rstPrueter!valor5
                                    ZValor(6) = rstPrueter!valor6
                                    ZValor(7) = rstPrueter!valor7
                                    ZValor(8) = rstPrueter!valor8
                                    ZValor(9) = rstPrueter!valor9
                                    ZValor(10) = rstPrueter!valor10
                                        
                                    rstPrueter.Close
                                    
                                    WFechaElaboracion = ""
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Hoja"
                                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZLote + "'"
                                    spHoja = ZSql
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                                        ZZHoja = rstHoja!Hoja
                                        ZZProducto = rstHoja!Producto
                                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                        ZZFecha = rstHoja!Fecha
                                        ZZMeses = ""
                                        rstHoja.Close
                                        
                                        spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                                            rstTerminado.Close
                                        End If
                                        
                                        If Val(ZZMeses) <> 0 Then
                                        
                                            If Val(ZZRevalida) <> 0 Then
                                                WVida = Val(ZZMesesRevalida)
                                                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                                WAno = Val(Right$(ZZFechaRevalida, 4))
                                                    Else
                                                WVida = Val(ZZMeses)
                                                WMes = Val(Mid$(ZZFecha, 4, 2))
                                                WAno = Val(Right$(ZZFecha, 4))
                                            End If
                                            
                                            For Ciclo = 1 To WVida
                                                WMes = WMes + 1
                                                If WMes > 12 Then
                                                    WAno = WAno + 1
                                                    WMes = 1
                                                End If
                                            Next Ciclo
                                            ZMes = Str$(WMes)
                                            ZAno = Str$(WAno)
                                            Call Ceros(ZMes, 2)
                                            Call Ceros(ZAno, 4)
                                            WFechaElaboracion = ZMes + "/" + ZAno
                                            
                                        End If
                                        
                                    End If
                                        
                                    If Left$(ZArticulo, 2) = "SE" Then
                                        WProducto = "SE" + Mid$(ZArticulo, 3, 10)
                                            Else
                                        WProducto = "PT" + Mid$(ZArticulo, 3, 10)
                                    End If
                                        
                                    Select Case Val(Wempresa)
                                        Case 1, 3, 5, 6, 7, 10, 11
                                            Wempresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            Wempresa = "0004"
                                            txtOdbc = "Empresa04"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                        
                                    LlamaImprime = "N"
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM EspecifUnificaVersion"
                                    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
                                    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                                    spEspecifUnificaVersion = ZSql
                                    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEspecifUnificaVersion.RecordCount > 0 Then
                                        With rstEspecifUnificaVersion
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    
                                                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                                                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                                                            
                                                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                                
                                                    End If
                                                        
                                                    If WDesde > WFechaord And LlamaImprime = "N" Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                    End If
                                                    
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstEspecifUnificaVersion.Close
                                    End If
                                    
                                    If LlamaImprime = "N" Then
                                    
                                        Sql1 = "Select *"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                
                                            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
                                            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
                                            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
                                            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
                                            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
                                            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
                                            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
                                            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
                                            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
                                            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
                                                                
                                            ZStd(1, 1) = rstEspecifUnifica!Valor1
                                            ZStd(2, 1) = rstEspecifUnifica!valor2
                                            ZStd(3, 1) = rstEspecifUnifica!Valor3
                                            ZStd(4, 1) = rstEspecifUnifica!valor4
                                            ZStd(5, 1) = rstEspecifUnifica!valor5
                                            ZStd(6, 1) = rstEspecifUnifica!valor6
                                            ZStd(7, 1) = rstEspecifUnifica!valor7
                                            ZStd(8, 1) = rstEspecifUnifica!valor8
                                            ZStd(9, 1) = rstEspecifUnifica!valor9
                                            ZStd(10, 1) = rstEspecifUnifica!valor10
                                                                
                                            ZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                                            ZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                                            ZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                                            ZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                                            ZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                                            ZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                                            ZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                                            ZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                                            ZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                                            ZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                                                                
                                            ZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                                            ZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                                            ZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                                            ZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                                            ZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                                            ZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                                            ZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                                            ZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                                            ZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                                            ZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                                    
                                            ZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                                            ZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!HAsta2), "", rstEspecifUnifica!HAsta2)
                                            ZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                                            ZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                                            ZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                                            ZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                                            ZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                                            ZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                                            ZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                                            ZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                                
                                            ZStd(1, 5) = IIf(IsNull(rstEspecifUnifica!Valor1Ing), "", rstEspecifUnifica!Valor1Ing)
                                            ZStd(2, 5) = IIf(IsNull(rstEspecifUnifica!Valor2Ing), "", rstEspecifUnifica!Valor2Ing)
                                            ZStd(3, 5) = IIf(IsNull(rstEspecifUnifica!Valor3Ing), "", rstEspecifUnifica!Valor3Ing)
                                            ZStd(4, 5) = IIf(IsNull(rstEspecifUnifica!Valor4Ing), "", rstEspecifUnifica!Valor4Ing)
                                            ZStd(5, 5) = IIf(IsNull(rstEspecifUnifica!Valor5Ing), "", rstEspecifUnifica!Valor5Ing)
                                            ZStd(6, 5) = IIf(IsNull(rstEspecifUnifica!Valor6Ing), "", rstEspecifUnifica!Valor6Ing)
                                            ZStd(7, 5) = IIf(IsNull(rstEspecifUnifica!Valor7Ing), "", rstEspecifUnifica!Valor7Ing)
                                            ZStd(8, 5) = IIf(IsNull(rstEspecifUnifica!Valor8Ing), "", rstEspecifUnifica!Valor8Ing)
                                            ZStd(9, 5) = IIf(IsNull(rstEspecifUnifica!Valor9Ing), "", rstEspecifUnifica!Valor9Ing)
                                            ZStd(10, 5) = IIf(IsNull(rstEspecifUnifica!Valor10Ing), "", rstEspecifUnifica!Valor10Ing)
                                                                
                                            ZStd(1, 6) = IIf(IsNull(rstEspecifUnifica!Valor11Ing), "", rstEspecifUnifica!Valor11Ing)
                                            ZStd(2, 6) = IIf(IsNull(rstEspecifUnifica!Valor22Ing), "", rstEspecifUnifica!Valor22Ing)
                                            ZStd(3, 6) = IIf(IsNull(rstEspecifUnifica!Valor33Ing), "", rstEspecifUnifica!Valor33Ing)
                                            ZStd(4, 6) = IIf(IsNull(rstEspecifUnifica!Valor44Ing), "", rstEspecifUnifica!Valor44Ing)
                                            ZStd(5, 6) = IIf(IsNull(rstEspecifUnifica!Valor55Ing), "", rstEspecifUnifica!Valor55Ing)
                                            ZStd(6, 6) = IIf(IsNull(rstEspecifUnifica!Valor66Ing), "", rstEspecifUnifica!Valor66Ing)
                                            ZStd(7, 6) = IIf(IsNull(rstEspecifUnifica!Valor77Ing), "", rstEspecifUnifica!Valor77Ing)
                                            ZStd(8, 6) = IIf(IsNull(rstEspecifUnifica!Valor88Ing), "", rstEspecifUnifica!Valor88Ing)
                                            ZStd(9, 6) = IIf(IsNull(rstEspecifUnifica!Valor99Ing), "", rstEspecifUnifica!Valor99Ing)
                                            ZStd(10, 6) = IIf(IsNull(rstEspecifUnifica!Valor1010Ing), "", rstEspecifUnifica!Valor1010Ing)
                                                                
                                            ZVersion = rstEspecifUnifica!Version
                                            rstEspecifUnifica.Close
                                            LlamaImprime = "S"
                                        End If
                                    
                                    End If
                                    
                                    If LlamaImprime = "S" Then
                                        
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(1) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(1) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(2) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(2) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(3) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(3) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(4) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(4) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(5) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(5) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(6) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(6) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(7) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(7) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(8) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(8) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(9) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(9) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(10) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(10) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                                                
                                        Call Conecta_Empresa
                                        
                                        XEmpresa = Wempresa
                                        Select Case Val(XEmpresa)
                                            Case 1, 3, 5, 6, 7, 10, 11
                                                Wempresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case 2, 4, 8, 9
                                                Wempresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                        End Select
                                        
                                        ZImpreVto = 0
                                        ZRazon = ""
                                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstCliente.RecordCount > 0 Then
                                            ZRazon = Left$(rstCliente!Razon, 50)
                                            ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                                            rstCliente.Close
                                        End If
                                        
                                        ZZImpreVtoTermi = 0
                                        spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
                                            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
                                            rstTerminado.Close
                                        End If
                                        
                                        If ZZImpreVtoTermi = 0 Then
                                            If ZImpreVto <> 1 Then
                                                Rem WFechaElaboracion = ""
                                            End If
                                        End If
                                            
                                        ZCliente = UCase(ZCliente)
                                        ZArticulo = UCase(ZArticulo)
                                        ZClave = ZCliente + ZArticulo
                        
                                        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
                                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstPrecios.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
                                            rstPrecios.Close
                                        End If
                                        
                                        Call Conecta_Empresa
                                                
                                        ZSql = "DELETE Certificado"
                                        spCertificado = ZSql
                                        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                            
                                        LugarMetodo = 0
                                                
                                        For CiclaMetodo = 1 To 10
                                                
                                            If ZOpcion(CiclaMetodo) = 1 Then
                                                
                                                LugarMetodo = LugarMetodo + 1
                                                    
                                                ZOrden = ""
                                                ZClave1 = ZLote
                                                Call Ceros(ZClave1, 6)
                                                ZClave2 = Str$(LugarMetodo)
                                                Call Ceros(ZClave2, 2)
                                                ZClave = ZClave1 + ZClave2
                                                ZMetodo = ZEnsayo(CiclaMetodo)
                                                
                                                If Val(ZStd(CiclaMetodo, 3)) <> 0 Or Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                                    ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                                    If WIdioma = 1 Then
                                                        ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                                    End If
                                                    ZValorNormalII = ""
                                                        Else
                                                    If WIdioma = 0 Then
                                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                            Else
                                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 5), 50)
                                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 6), 50)
                                                    End If
                                                End If
                                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 50)
                                                If WIdioma = 1 Then
                                                    If UCase(Trim(ZValorPartidaI)) = "CUMPLE" Then
                                                        ZValorPartidaI = "OK"
                                                    End If
                                                End If
                                                
                                                ZValorNormalI = Trim(ZValorNormalI)
                                                ZCanti = 50 - Len(ZValorNormalI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalI = Space$(ZCanti) + ZValorNormalI
                                                
                                                ZValorNormalII = Trim(ZValorNormalII)
                                                ZCanti = 50 - Len(ZValorNormalII)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalII = Space$(ZCanti) + ZValorNormalII
                                                
                                                ZValorPartidaI = Trim(ZValorPartidaI)
                                                ZCanti = 50 - Len(ZValorPartidaI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorPartidaI = Space$(ZCanti) + ZValorPartidaI
                                                
                                                ZValorPartidaII = ""
                                                ZObservacionesI = ""
                                                ZObservacionesII = ""
                                                ZObservacionesIII = "Version " + ZVersion
                                                ZObservacionesIV = ""
                                                ZObservacionesV = ""
                                                ZObservacionesVI = ""
                                                If Val(Wempresa) = 1 Then
                                                    ZEmpresa = "Surfactan S.A."
                                                        Else
                                                    ZEmpresa = "Pellital S.A."
                                                End If
                                                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                                ZFechaII = WFechaElaboracion
                                                
                                                ZExamen = Trim(ZDescri(CiclaMetodo))
                                                ZExamenII = ""
                                                ZHasta = Len(Trim(ZExamen))
                                                If ZHasta > 25 Then
                                                    For Cicla = ZHasta To 1 Step -1
                                                        If Mid(ZExamen, Cicla, 1) = Space(1) Then
                                                            ZExamenII = Mid(ZExamen, Cicla - Desde + 1, 25)
                                                            ZExamen = Mid(ZExamen, 1, Cicla - Desde)
                                                            Exit For
                                                        End If
                                                    Next Cicla
                                                End If
                                                        
                                                ZSql = ""
                                                ZSql = ZSql + "INSERT INTO Certificado ("
                                                ZSql = ZSql + "Clave ,"
                                                ZSql = ZSql + "Partida ,"
                                                ZSql = ZSql + "Renglon ,"
                                                ZSql = ZSql + "Razon ,"
                                                ZSql = ZSql + "Orden ,"
                                                ZSql = ZSql + "Terminado ,"
                                                ZSql = ZSql + "Descripcion ,"
                                                ZSql = ZSql + "Fecha ,"
                                                ZSql = ZSql + "FechaII ,"
                                                ZSql = ZSql + "Cantidad ,"
                                                ZSql = ZSql + "Examen ,"
                                                ZSql = ZSql + "ExamenII ,"
                                                ZSql = ZSql + "ValorPartidaI ,"
                                                ZSql = ZSql + "ValorPartidaII ,"
                                                ZSql = ZSql + "ValorNormalI ,"
                                                ZSql = ZSql + "ValorNormalII ,"
                                                ZSql = ZSql + "Observaciones1 ,"
                                                ZSql = ZSql + "Observaciones2 ,"
                                                ZSql = ZSql + "Observaciones3 ,"
                                                ZSql = ZSql + "Observaciones4 ,"
                                                ZSql = ZSql + "Observaciones5 ,"
                                                ZSql = ZSql + "Observaciones6 ,"
                                                ZSql = ZSql + "Metodo ,"
                                                ZSql = ZSql + "Empresa )"
                                                ZSql = ZSql + "Values ("
                                                ZSql = ZSql + "'" + ZClave + "',"
                                                ZSql = ZSql + "'" + ZLote + "',"
                                                ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                                ZSql = ZSql + "'" + ZRazon + "',"
                                                ZSql = ZSql + "'" + ZOrden + "',"
                                                ZSql = ZSql + "'" + ZArticulo + "',"
                                                ZSql = ZSql + "'" + ZDesArticulo + "',"
                                                ZSql = ZSql + "'" + ZFecha + "',"
                                                ZSql = ZSql + "'" + ZFechaII + "',"
                                                ZSql = ZSql + "'" + ZCantidad + "',"
                                                ZSql = ZSql + "'" + ZExamen + "',"
                                                ZSql = ZSql + "'" + ZExamenII + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                                ZSql = ZSql + "'" + ZValorNormalI + "',"
                                                ZSql = ZSql + "'" + ZValorNormalII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesI + "',"
                                                ZSql = ZSql + "'" + ZObservacionesII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                                ZSql = ZSql + "'" + ZMetodo + "',"
                                                ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                                spCertificado = ZSql
                                                Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                        
                                            End If
                                                            
                                        Next CiclaMetodo
                                            
                                        Do
                                            
                                            If LugarMetodo = 10 Then
                                                Exit Do
                                            End If
                                                
                                            LugarMetodo = LugarMetodo + 1
                                                    
                                            ZOrden = ""
                                            ZClave1 = ZLote
                                            Call Ceros(ZClave1, 6)
                                            ZClave2 = Str$(LugarMetodo)
                                            Call Ceros(ZClave2, 2)
                                            ZClave = ZClave1 + ZClave2
                                            ZMetodo = ""
                                            ZExamen = ""
                                            ZValorNormalI = ""
                                            ZValorNormalII = ""
                                            ZValorPartidaI = ""
                                            ZValorPartidaII = ""
                                            ZObservacionesI = ""
                                            ZObservacionesII = ""
                                            ZObservacionesIII = "Version " + ZVersion
                                            ZObservacionesIV = ""
                                            ZObservacionesV = ""
                                            ZObservacionesVI = ""
                                            If Val(Wempresa) = 1 Then
                                                ZEmpresa = "Surfactan S.A."
                                                    Else
                                                ZEmpresa = "Pellital S.A."
                                            End If
                                            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                            ZFechaII = WFechaElaboracion
                                            ZExamenII = ""
                                                        
                                            ZSql = ""
                                            ZSql = ZSql + "INSERT INTO Certificado ("
                                            ZSql = ZSql + "Clave ,"
                                            ZSql = ZSql + "Partida ,"
                                            ZSql = ZSql + "Renglon ,"
                                            ZSql = ZSql + "Razon ,"
                                            ZSql = ZSql + "Orden ,"
                                            ZSql = ZSql + "Terminado ,"
                                            ZSql = ZSql + "Descripcion ,"
                                            ZSql = ZSql + "Fecha ,"
                                            ZSql = ZSql + "Cantidad ,"
                                            ZSql = ZSql + "Examen ,"
                                            ZSql = ZSql + "ValorPartidaI ,"
                                            ZSql = ZSql + "ValorPartidaII ,"
                                            ZSql = ZSql + "ValorNormalI ,"
                                            ZSql = ZSql + "ValorNormalII ,"
                                            ZSql = ZSql + "Observaciones1 ,"
                                            ZSql = ZSql + "Observaciones2 ,"
                                            ZSql = ZSql + "Observaciones3 ,"
                                            ZSql = ZSql + "Observaciones4 ,"
                                            ZSql = ZSql + "Observaciones5 ,"
                                            ZSql = ZSql + "Observaciones6 ,"
                                            ZSql = ZSql + "Metodo ,"
                                            ZSql = ZSql + "Empresa )"
                                            ZSql = ZSql + "Values ("
                                            ZSql = ZSql + "'" + ZClave + "',"
                                            ZSql = ZSql + "'" + ZLote + "',"
                                            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                            ZSql = ZSql + "'" + ZRazon + "',"
                                            ZSql = ZSql + "'" + ZOrden + "',"
                                            ZSql = ZSql + "'" + ZArticulo + "',"
                                            ZSql = ZSql + "'" + ZDesArticulo + "',"
                                            ZSql = ZSql + "'" + ZFecha + "',"
                                            ZSql = ZSql + "'" + ZCantidad + "',"
                                            ZSql = ZSql + "'" + ZExamen + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                            ZSql = ZSql + "'" + ZValorNormalI + "',"
                                            ZSql = ZSql + "'" + ZValorNormalII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesI + "',"
                                            ZSql = ZSql + "'" + ZObservacionesII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                            ZSql = ZSql + "'" + ZMetodo + "',"
                                            ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                            spCertificado = ZSql
                                            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                
                                        Loop
                                                
                                        Listado.WindowTitle = "Certificado de Analisis"
                                        Listado.WindowTop = 0
                                        Listado.WindowLeft = 0
                                        Listado.WindowWidth = Screen.Width
                                        Listado.WindowHeight = Screen.Height
                        
                                        Listado.Destination = 1
                                        Rem Listado.Destination = 0
                                                
                                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                                            If WIdioma = 0 Then
                                                Listado.ReportFileName = "CertificadoNuevo.rpt"
                                                    Else
                                                Listado.ReportFileName = "CertificadoNuevoIngles.rpt"
                                            End If
                                                Else
                                            Listado.ReportFileName = "CertificadoPelli.rpt"
                                        End If
                                                    
                                        DbConnect = db.Connect
                                        DSQ = getDatabase(DbConnect)
                        
                                        Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                                        + "From " _
                                                        + DSQ + ".dbo.Certificado Certificado " _
                                                        + "Where " _
                                                        + "Certificado.Partida >= 0 AND " _
                                                        + "Certificado.Partida <= 999999"
                                                        
                                        Listado.Destination = 1
                                        Rem Listado.Destination = 0
                                        Listado.CopiesToPrinter = 1
                        
                                        Listado.Connect = Connect()
                                        Listado.Action = 1
                                                
                                    End If
                                          
                                End If
                                    
                            Next ZCiclo
                            
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    Wempresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                            
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                
                        Else
                        
                    If Left$(Articulo, 2) = "DY" Then
                        
                        Auxi = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                        ZZCambia = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM Laudo"
                        ZSql = ZSql & " Where Laudo.Laudo = " + "'" + AuxiliarII(DA, ZZLugar) + "'"
                        ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZPartiOri = Trim(rstLaudo!PartiOri)
                            rstLaudo.Close
                            ZZCambia = "S"
                            ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                            ZZEstado = Dir(ZZRuta)
                            ZZEstado = Trim(ZZEstado)
                            If ZZEstado <> "" Then
                            
                                ZZNombreArchi = 1
                                Do
                                    Auxi = Str$(ZZNombreArchi)
                                    Call Ceros(Auxi, 8)
                                    ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                                    
                                    ZZRutaII = ZZNombreArchiII
                                    ZZEstadoII = Dir(ZZRutaII)
                                    ZZEstadoII = Trim(ZZEstadoII)
                                    If ZZEstadoII = "" Then
                                        Exit Do
                                    End If
                                    ZZNombreArchi = ZZNombreArchi + 1
                                Loop
                                
                                FileCopy ZZRuta, ZZRutaII
                                Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                                Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                                ZZImprePdf = "S"
                                Rem TiempoPausa = 2 ' Asigna hora de inicio.
                                Rem Inicio = Timer  ' Establece la hora de inicio.
                                Rem Do While Timer < Inicio + TiempoPausa
                                Rem     DoEvents    ' Cambia a otros procesos.
                                Rem Loop
                            
                                Rem Select Case ZZVersion
                                Rem     Case 1
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                Rem     Case 2
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 3
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 4
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 5
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case Else
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem         Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                                Rem End Select
                                
                                    Else
                                m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                            End If
                        End If
                        
                        If ZZCambia = "N" Then
                                        
                            XEmpresa = Wempresa
                                    
                            Wempresa = "0006"
                            txtOdbc = "Empresa06"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM Laudo"
                            ZSql = ZSql & " Where Laudo.Laudo = " + "'" + AuxiliarII(DA, ZZLugar) + "'"
                            ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                ZZPartiOri = Trim(rstLaudo!PartiOri)
                                rstLaudo.Close
                                ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                                ZZEstado = Dir(ZZRuta)
                                ZZEstado = Trim(ZZEstado)
                                If ZZEstado <> "" Then
                                
                                    ZZNombreArchi = 1
                                    Do
                                        Auxi = Str$(ZZNombreArchi)
                                        Call Ceros(Auxi, 8)
                                        ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                                        
                                        ZZRutaII = ZZNombreArchiII
                                        ZZEstadoII = Dir(ZZRutaII)
                                        ZZEstadoII = Trim(ZZEstadoII)
                                        If ZZEstadoII = "" Then
                                            Exit Do
                                        End If
                                        ZZNombreArchi = ZZNombreArchi + 1
                                    Loop
                                    
                                    FileCopy ZZRuta, ZZRutaII
                                    Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                                    Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                                    ZZImprePdf = "S"
                                    Rem TiempoPausa = 2 ' Asigna hora de inicio.
                                    Rem Inicio = Timer  ' Establece la hora de inicio.
                                    Rem Do While Timer < Inicio + TiempoPausa
                                    Rem     DoEvents    ' Cambia a otros procesos.
                                    Rem Loop
                                
                                    Rem Select Case ZZVersion
                                    Rem     Case 1
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                    Rem     Case 2
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case 3
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem      Case 4
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case 5
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case Else
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem         Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                                    Rem End Select
                                        Else
                                    m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                    a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                End If
                            End If
                            
                            Call Conecta_Empresa
                        
                        End If
                        
                    End If
                    
                End If
            End If
                
        Next ZZCiclo
        
    Next DA
    
    For ZZCicloFicha = 1 To ZLugarFicha
    
        ZZIntervencion = ZImpreFicha(ZZCicloFicha)
        
        Call Ceros(ZZIntervencion, 3)
        
        ZZRuta = "w:\FICHASIS\GUIADEEMERGENCIA" + ZZIntervencion + ".PDF"
        ZZEstado = Dir(ZZRuta)
        ZZEstado = Trim(ZZEstado)
        If ZZEstado <> "" Then
        
            ZZNombreArchi = 1
            Do
                Auxi = Str$(ZZNombreArchi)
                Call Ceros(Auxi, 8)
                ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                
                ZZRutaII = ZZNombreArchiII
                ZZEstadoII = Dir(ZZRutaII)
                ZZEstadoII = Trim(ZZEstadoII)
                If ZZEstadoII = "" Then
                    Exit Do
                End If
                ZZNombreArchi = ZZNombreArchi + 1
            Loop
            
            FileCopy ZZRuta, ZZRutaII
            Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
            Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
            ZZImprePdf = "S"
            Rem TiempoPausa = 2 ' Asigna hora de inicio.
            Rem Inicio = Timer  ' Establece la hora de inicio.
            Rem Do While Timer < Inicio + TiempoPausa
            Rem     DoEvents    ' Cambia a otros procesos.
            Rem Loop
        
        
        
            Rem Select Case ZZVersion
            Rem     Case 1
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem     Case 2
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem     Case 3
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem     Case 4
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem     Case 5
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem     Case Else
            Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            Rem         Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
            Rem End Select
                Else
            m$ = "El articulo " + Articulo + " posee la Ficha de Emergencia Nro " + ZZIntervencion + " y no se ha encontrado"
            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
        End If
    
    Next ZZCicloFicha
    
    If ZZImprePdf = "S" Then
        Call Verifica_Impresora
        Call Verifica_ImpresoraII
        RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docpdf" + Chr$(34) + " C:\pdfprint\*.pdf", 6)
    End If
    
    PrgFactuexpo.Show
    Numero.SetFocus
    
End Sub

Private Sub Impresion_Certificado_Pantalla()

    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        WIdioma = IIf(IsNull(rstClientes!Idioma), "0", rstClientes!Idioma)
        rstClientes.Close
    End If
    
    Call Conecta_Empresa

    For DA = 1 To 99
    
        Articulo = AuxiliarII(DA, 1)
        Cantidad = AuxiliarII(DA, 2)
        Precio = AuxiliarII(DA, 3)
        WRenglon = AuxiliarII(DA, 4)
        WLote1 = AuxiliarII(DA, 5)
        WCanti1 = AuxiliarII(DA, 6)
        WLote2 = AuxiliarII(DA, 7)
        WCanti2 = AuxiliarII(DA, 8)
        Wlote3 = AuxiliarII(DA, 9)
        WCanti3 = AuxiliarII(DA, 10)
        WLote4 = AuxiliarII(DA, 11)
        WCanti4 = AuxiliarII(DA, 12)
        WLote5 = AuxiliarII(DA, 13)
        WCanti5 = AuxiliarII(DA, 14)
        WLote6 = AuxiliarII(DA, 15)
        WCanti6 = AuxiliarII(DA, 16)
        WLote7 = AuxiliarII(DA, 17)
        WCanti7 = AuxiliarII(DA, 18)
        WLote8 = AuxiliarII(DA, 19)
        WCanti8 = AuxiliarII(DA, 20)
        WLote9 = AuxiliarII(DA, 21)
        WCanti9 = AuxiliarII(DA, 22)
        WLote10 = AuxiliarII(DA, 23)
        WCanti10 = AuxiliarII(DA, 24)
        WLote11 = AuxiliarII(DA, 25)
        WCanti11 = AuxiliarII(DA, 26)
        WLote12 = AuxiliarII(DA, 27)
        WCanti12 = AuxiliarII(DA, 28)
        ZZDescriArticulo = Trim(AuxiliarII(DA, 30))
        
        If Articulo = "" Then Exit For
    
        Rem
        Rem certificado de analisis
        Rem
    
        For ZZCiclo = 1 To 12
            
            Select Case ZZCiclo
                Case 1
                    ZZLugar = 5
                Case 2
                    ZZLugar = 7
                Case 3
                    ZZLugar = 9
                Case 4
                    ZZLugar = 11
                Case 5
                    ZZLugar = 13
                Case 6
                    ZZLugar = 15
                Case 7
                    ZZLugar = 17
                Case 8
                    ZZLugar = 19
                Case 9
                    ZZLugar = 21
                Case 10
                    ZZLugar = 23
                Case 11
                    ZZLugar = 25
                Case Else
                    ZZLugar = 27
            End Select
            
            If Val(AuxiliarII(DA, ZZLugar)) <> 0 Then
        
                ZZEntra = "N"
        
                If Left$(UCase(Articulo), 2) = "PT" Then
                
                    XCodigo = Val(Mid$(Articulo, 4, 5))
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    If XCodigo >= 40000 And XCodigo <= 41000 Then
                                        XTipoPro = "TA"
                                            Else
                                        XTipoPro = "PT"
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    If Left$(Articulo, 2) = "YQ" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YH" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YP" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YF" Then
                        XTipoPro = "FA"
                    End If
                    If Left$(Articulo, 2) = "PE" Then
                        XTipoPro = "PT"
                    End If
            
                    ZLinea = 0
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
            
                    Select Case ZLinea
                        Case 8
                            XTipoPro = "PG"
                        Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                            XTipoPro = "FA"
                        Case Else
                    End Select
            
                    If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                    
                        XEmpresa = Wempresa
                        
                        Select Case Val(Wempresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = AuxiliarII(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = AuxiliarII(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = AuxiliarII(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = AuxiliarII(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                            
                        Erase ZOpcion
                        Erase ZValor
                        Erase ZEnsayo
                        Erase ZStd
                        Erase ZDescri
                        Erase ZDescriII
                            
                        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        ZVersion = 0
                        
                        ZZEntra = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM AltaCertificado"
                        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                        ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                        spAltaCertificado = ZSql
                        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstAltaCertificado.RecordCount > 0 Then
                            ZOpcion(1) = rstAltaCertificado!Opcion1
                            ZOpcion(2) = rstAltaCertificado!Opcion2
                            ZOpcion(3) = rstAltaCertificado!Opcion3
                            ZOpcion(4) = rstAltaCertificado!Opcion4
                            ZOpcion(5) = rstAltaCertificado!Opcion5
                            ZOpcion(6) = rstAltaCertificado!Opcion6
                            ZOpcion(7) = rstAltaCertificado!Opcion7
                            ZOpcion(8) = rstAltaCertificado!Opcion8
                            ZOpcion(9) = rstAltaCertificado!Opcion9
                            ZOpcion(10) = rstAltaCertificado!Opcion10
                            rstAltaCertificado.Close
                            ZZEntra = "S"
                        End If
                        
                        If ZZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                        End If
                        
                        If ZZEntra = "S" Then
                            If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                ZZEntra = "N"
                            End If
                        End If
                                
                        If ZZEntra = "N" Then
                            m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                        End If
                                        
                        If ZZEntra = "S" Then
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    CargaEmpresa(1, 1) = "0001"
                                    CargaEmpresa(1, 2) = "Empresa01"
                                    CargaEmpresa(2, 1) = "0003"
                                    CargaEmpresa(2, 2) = "Empresa03"
                                    CargaEmpresa(3, 1) = "0005"
                                    CargaEmpresa(3, 2) = "Empresa05"
                                    CargaEmpresa(4, 1) = "0006"
                                    CargaEmpresa(4, 2) = "Empresa06"
                                    CargaEmpresa(5, 1) = "0007"
                                    CargaEmpresa(5, 2) = "Empresa07"
                                    CargaEmpresa(6, 1) = "0010"
                                    CargaEmpresa(6, 2) = "Empresa10"
                                    CargaEmpresa(7, 1) = "0011"
                                    CargaEmpresa(7, 2) = "Empresa11"
                                    ZHasta1 = 7
                                Case Else
                                    CargaEmpresa(1, 1) = "0002"
                                    CargaEmpresa(1, 2) = "Empresa02"
                                    CargaEmpresa(2, 1) = "0004"
                                    CargaEmpresa(2, 2) = "Empresa04"
                                    CargaEmpresa(3, 1) = "0008"
                                    CargaEmpresa(3, 2) = "Empresa08"
                                    CargaEmpresa(4, 1) = "0009"
                                    CargaEmpresa(4, 2) = "Empresa09"
                                    ZHasta1 = 4
                            End Select
                                
                            For ZCiclo = 1 To ZHasta1
                                
                                Wempresa = CargaEmpresa(ZCiclo, 1)
                                txtOdbc = CargaEmpresa(ZCiclo, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        
                                If Val(ZLote) > 99999 Then
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 6)
                                        Else
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 5)
                                End If
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Prueter"
                                ZSql = ZSql + " Where Prueter.Lote = " + "'" + ZLote + "'"
                                spPrueter = ZSql
                                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPrueter.RecordCount > 0 Then
                                
                                    If Left$(rstPrueter!prueba, 1) = "2" Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                    If rstPrueter!Producto <> ZProducto Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                        
                                    WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
                                            
                                    ZValor(1) = rstPrueter!Valor1
                                    ZValor(2) = rstPrueter!valor2
                                    ZValor(3) = rstPrueter!Valor3
                                    ZValor(4) = rstPrueter!valor4
                                    ZValor(5) = rstPrueter!valor5
                                    ZValor(6) = rstPrueter!valor6
                                    ZValor(7) = rstPrueter!valor7
                                    ZValor(8) = rstPrueter!valor8
                                    ZValor(9) = rstPrueter!valor9
                                    ZValor(10) = rstPrueter!valor10
                                        
                                    rstPrueter.Close
                                    
                                    WFechaElaboracion = ""
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Hoja"
                                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZLote + "'"
                                    spHoja = ZSql
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                                        ZZHoja = rstHoja!Hoja
                                        ZZProducto = rstHoja!Producto
                                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                        ZZFecha = rstHoja!Fecha
                                        ZZMeses = ""
                                        rstHoja.Close
                                        
                                        spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                                            rstTerminado.Close
                                        End If
                                        
                                        If Val(ZZMeses) <> 0 Then
                                        
                                            If Val(ZZRevalida) <> 0 Then
                                                WVida = Val(ZZMesesRevalida)
                                                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                                WAno = Val(Right$(ZZFechaRevalida, 4))
                                                    Else
                                                WVida = Val(ZZMeses)
                                                WMes = Val(Mid$(ZZFecha, 4, 2))
                                                WAno = Val(Right$(ZZFecha, 4))
                                            End If
                                            
                                            For Ciclo = 1 To WVida
                                                WMes = WMes + 1
                                                If WMes > 12 Then
                                                    WAno = WAno + 1
                                                    WMes = 1
                                                End If
                                            Next Ciclo
                                            ZMes = Str$(WMes)
                                            ZAno = Str$(WAno)
                                            Call Ceros(ZMes, 2)
                                            Call Ceros(ZAno, 4)
                                            WFechaElaboracion = ZMes + "/" + ZAno
                                            
                                        End If
                                        
                                    End If
                                        
                                    If Left$(ZArticulo, 2) = "SE" Then
                                        WProducto = "SE" + Mid$(ZArticulo, 3, 10)
                                            Else
                                        WProducto = "PT" + Mid$(ZArticulo, 3, 10)
                                    End If
                                        
                                    Select Case Val(Wempresa)
                                        Case 1, 3, 5, 6, 7, 10, 11
                                            Wempresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            Wempresa = "0004"
                                            txtOdbc = "Empresa04"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                        
                                    LlamaImprime = "N"
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM EspecifUnificaVersion"
                                    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
                                    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                                    spEspecifUnificaVersion = ZSql
                                    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEspecifUnificaVersion.RecordCount > 0 Then
                                        With rstEspecifUnificaVersion
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    
                                                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                                                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                                                            
                                                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                                
                                                    End If
                                                        
                                                    If WDesde > WFechaord And LlamaImprime = "N" Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                    End If
                                                    
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstEspecifUnificaVersion.Close
                                    End If
                                    
                                    If LlamaImprime = "N" Then
                                    
                                        Sql1 = "Select *"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                
                                            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
                                            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
                                            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
                                            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
                                            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
                                            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
                                            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
                                            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
                                            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
                                            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
                                                                
                                            ZStd(1, 1) = rstEspecifUnifica!Valor1
                                            ZStd(2, 1) = rstEspecifUnifica!valor2
                                            ZStd(3, 1) = rstEspecifUnifica!Valor3
                                            ZStd(4, 1) = rstEspecifUnifica!valor4
                                            ZStd(5, 1) = rstEspecifUnifica!valor5
                                            ZStd(6, 1) = rstEspecifUnifica!valor6
                                            ZStd(7, 1) = rstEspecifUnifica!valor7
                                            ZStd(8, 1) = rstEspecifUnifica!valor8
                                            ZStd(9, 1) = rstEspecifUnifica!valor9
                                            ZStd(10, 1) = rstEspecifUnifica!valor10
                                                                
                                            ZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                                            ZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                                            ZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                                            ZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                                            ZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                                            ZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                                            ZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                                            ZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                                            ZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                                            ZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                                                                
                                            ZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                                            ZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                                            ZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                                            ZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                                            ZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                                            ZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                                            ZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                                            ZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                                            ZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                                            ZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                                    
                                            ZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                                            ZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!HAsta2), "", rstEspecifUnifica!HAsta2)
                                            ZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                                            ZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                                            ZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                                            ZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                                            ZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                                            ZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                                            ZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                                            ZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                                
                                            ZStd(1, 5) = IIf(IsNull(rstEspecifUnifica!Valor1Ing), "", rstEspecifUnifica!Valor1Ing)
                                            ZStd(2, 5) = IIf(IsNull(rstEspecifUnifica!Valor2Ing), "", rstEspecifUnifica!Valor2Ing)
                                            ZStd(3, 5) = IIf(IsNull(rstEspecifUnifica!Valor3Ing), "", rstEspecifUnifica!Valor3Ing)
                                            ZStd(4, 5) = IIf(IsNull(rstEspecifUnifica!Valor4Ing), "", rstEspecifUnifica!Valor4Ing)
                                            ZStd(5, 5) = IIf(IsNull(rstEspecifUnifica!Valor5Ing), "", rstEspecifUnifica!Valor5Ing)
                                            ZStd(6, 5) = IIf(IsNull(rstEspecifUnifica!Valor6Ing), "", rstEspecifUnifica!Valor6Ing)
                                            ZStd(7, 5) = IIf(IsNull(rstEspecifUnifica!Valor7Ing), "", rstEspecifUnifica!Valor7Ing)
                                            ZStd(8, 5) = IIf(IsNull(rstEspecifUnifica!Valor8Ing), "", rstEspecifUnifica!Valor8Ing)
                                            ZStd(9, 5) = IIf(IsNull(rstEspecifUnifica!Valor9Ing), "", rstEspecifUnifica!Valor9Ing)
                                            ZStd(10, 5) = IIf(IsNull(rstEspecifUnifica!Valor10Ing), "", rstEspecifUnifica!Valor10Ing)
                                                                
                                            ZStd(1, 6) = IIf(IsNull(rstEspecifUnifica!Valor11Ing), "", rstEspecifUnifica!Valor11Ing)
                                            ZStd(2, 6) = IIf(IsNull(rstEspecifUnifica!Valor22Ing), "", rstEspecifUnifica!Valor22Ing)
                                            ZStd(3, 6) = IIf(IsNull(rstEspecifUnifica!Valor33Ing), "", rstEspecifUnifica!Valor33Ing)
                                            ZStd(4, 6) = IIf(IsNull(rstEspecifUnifica!Valor44Ing), "", rstEspecifUnifica!Valor44Ing)
                                            ZStd(5, 6) = IIf(IsNull(rstEspecifUnifica!Valor55Ing), "", rstEspecifUnifica!Valor55Ing)
                                            ZStd(6, 6) = IIf(IsNull(rstEspecifUnifica!Valor66Ing), "", rstEspecifUnifica!Valor66Ing)
                                            ZStd(7, 6) = IIf(IsNull(rstEspecifUnifica!Valor77Ing), "", rstEspecifUnifica!Valor77Ing)
                                            ZStd(8, 6) = IIf(IsNull(rstEspecifUnifica!Valor88Ing), "", rstEspecifUnifica!Valor88Ing)
                                            ZStd(9, 6) = IIf(IsNull(rstEspecifUnifica!Valor99Ing), "", rstEspecifUnifica!Valor99Ing)
                                            ZStd(10, 6) = IIf(IsNull(rstEspecifUnifica!Valor1010Ing), "", rstEspecifUnifica!Valor1010Ing)
                                                                
                                            ZVersion = rstEspecifUnifica!Version
                                            rstEspecifUnifica.Close
                                            LlamaImprime = "S"
                                        End If
                                    
                                    End If
                                    
                                    If LlamaImprime = "S" Then
                                        
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(1) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(1) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(2) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(2) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(3) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(3) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(4) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(4) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(5) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(5) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(6) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(6) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(7) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(7) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(8) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(8) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(9) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(9) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            If WIdioma = 0 Then
                                                ZDescri(10) = rstEnsayo!Descripcion
                                                    Else
                                                ZDescri(10) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                                            End If
                                            ZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                                                
                                        Call Conecta_Empresa
                                        
                                        XEmpresa = Wempresa
                                        Select Case Val(XEmpresa)
                                            Case 1, 3, 5, 6, 7, 10, 11
                                                Wempresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case 2, 4, 8, 9
                                                Wempresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                        End Select
                                        
                                        ZImpreVto = 0
                                        ZRazon = ""
                                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstCliente.RecordCount > 0 Then
                                            ZRazon = Left$(rstCliente!Razon, 50)
                                            ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                                            rstCliente.Close
                                        End If
                                        
                                        ZZImpreVtoTermi = 0
                                        spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
                                            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
                                            rstTerminado.Close
                                        End If
                                        
                                        If ZZImpreVtoTermi = 0 Then
                                            If ZImpreVto <> 1 Then
                                                Rem WFechaElaboracion = ""
                                            End If
                                        End If
                                            
                                        ZCliente = UCase(ZCliente)
                                        ZArticulo = UCase(ZArticulo)
                                        ZClave = ZCliente + ZArticulo
                        
                                        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
                                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstPrecios.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
                                            rstPrecios.Close
                                        End If
                                        
                                        Call Conecta_Empresa
                                                
                                        ZSql = "DELETE Certificado"
                                        spCertificado = ZSql
                                        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                            
                                        LugarMetodo = 0
                                                
                                        For CiclaMetodo = 1 To 10
                                                
                                            If ZOpcion(CiclaMetodo) = 1 Then
                                                
                                                LugarMetodo = LugarMetodo + 1
                                                    
                                                ZOrden = ""
                                                ZClave1 = ZLote
                                                Call Ceros(ZClave1, 6)
                                                ZClave2 = Str$(LugarMetodo)
                                                Call Ceros(ZClave2, 2)
                                                ZClave = ZClave1 + ZClave2
                                                ZMetodo = ZEnsayo(CiclaMetodo)
                                                
                                                If Val(ZStd(CiclaMetodo, 3)) <> 0 Or Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                                    ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                                    If WIdioma = 1 Then
                                                        ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                                    End If
                                                    ZValorNormalII = ""
                                                        Else
                                                    If WIdioma = 0 Then
                                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                            Else
                                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 5), 50)
                                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 6), 50)
                                                    End If
                                                End If
                                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 80)
                                                If WIdioma = 1 Then
                                                    If UCase(Trim(ZValorPartidaI)) = "CUMPLE" Then
                                                        ZValorPartidaI = "OK"
                                                    End If
                                                End If
                                                
                                                ZValorNormalI = Trim(ZValorNormalI)
                                                ZCanti = 80 - Len(ZValorNormalI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalI = Space$(ZCanti) + ZValorNormalI
                                                
                                                ZValorNormalII = Trim(ZValorNormalII)
                                                ZCanti = 80 - Len(ZValorNormalII)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalII = Space$(ZCanti) + ZValorNormalII
                                                
                                                ZValorPartidaI = Trim(ZValorPartidaI)
                                                ZCanti = 80 - Len(ZValorPartidaI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorPartidaI = Space$(ZCanti) + ZValorPartidaI
                                                
                                                ZValorPartidaII = ""
                                                ZObservacionesI = ""
                                                ZObservacionesII = ""
                                                ZObservacionesIII = "Version " + ZVersion
                                                ZObservacionesIV = ""
                                                ZObservacionesV = ""
                                                ZObservacionesVI = ""
                                                If Val(Wempresa) = 1 Then
                                                    ZEmpresa = "Surfactan S.A."
                                                        Else
                                                    ZEmpresa = "Pellital S.A."
                                                End If
                                                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                                ZFechaII = WFechaElaboracion
                                                
                                                ZExamen = Trim(ZDescri(CiclaMetodo))
                                                ZExamenII = ""
                                                ZHasta = Len(Trim(ZExamen))
                                                If ZHasta > 25 Then
                                                    For Cicla = ZHasta To 1 Step -1
                                                        If Mid(ZExamen, Cicla, 1) = Space(1) Then
                                                            ZExamenII = Mid(ZExamen, Cicla - Desde + 1, 25)
                                                            ZExamen = Mid(ZExamen, 1, Cicla - Desde)
                                                            Exit For
                                                        End If
                                                    Next Cicla
                                                End If
                                                        
                                                ZSql = ""
                                                ZSql = ZSql + "INSERT INTO Certificado ("
                                                ZSql = ZSql + "Clave ,"
                                                ZSql = ZSql + "Partida ,"
                                                ZSql = ZSql + "Renglon ,"
                                                ZSql = ZSql + "Razon ,"
                                                ZSql = ZSql + "Orden ,"
                                                ZSql = ZSql + "Terminado ,"
                                                ZSql = ZSql + "Descripcion ,"
                                                ZSql = ZSql + "Fecha ,"
                                                ZSql = ZSql + "FechaII ,"
                                                ZSql = ZSql + "Cantidad ,"
                                                ZSql = ZSql + "Examen ,"
                                                ZSql = ZSql + "ExamenII ,"
                                                ZSql = ZSql + "ValorPartidaI ,"
                                                ZSql = ZSql + "ValorPartidaII ,"
                                                ZSql = ZSql + "ValorNormalI ,"
                                                ZSql = ZSql + "ValorNormalII ,"
                                                ZSql = ZSql + "Observaciones1 ,"
                                                ZSql = ZSql + "Observaciones2 ,"
                                                ZSql = ZSql + "Observaciones3 ,"
                                                ZSql = ZSql + "Observaciones4 ,"
                                                ZSql = ZSql + "Observaciones5 ,"
                                                ZSql = ZSql + "Observaciones6 ,"
                                                ZSql = ZSql + "Metodo ,"
                                                ZSql = ZSql + "Empresa )"
                                                ZSql = ZSql + "Values ("
                                                ZSql = ZSql + "'" + ZClave + "',"
                                                ZSql = ZSql + "'" + ZLote + "',"
                                                ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                                ZSql = ZSql + "'" + ZRazon + "',"
                                                ZSql = ZSql + "'" + ZOrden + "',"
                                                ZSql = ZSql + "'" + ZArticulo + "',"
                                                ZSql = ZSql + "'" + ZDesArticulo + "',"
                                                ZSql = ZSql + "'" + ZFecha + "',"
                                                ZSql = ZSql + "'" + ZFechaII + "',"
                                                ZSql = ZSql + "'" + ZCantidad + "',"
                                                ZSql = ZSql + "'" + ZExamen + "',"
                                                ZSql = ZSql + "'" + ZExamenII + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                                ZSql = ZSql + "'" + ZValorNormalI + "',"
                                                ZSql = ZSql + "'" + ZValorNormalII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesI + "',"
                                                ZSql = ZSql + "'" + ZObservacionesII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                                ZSql = ZSql + "'" + ZMetodo + "',"
                                                ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                                spCertificado = ZSql
                                                Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                        
                                            End If
                                                            
                                        Next CiclaMetodo
                                            
                                        Do
                                            
                                            If LugarMetodo = 10 Then
                                                Exit Do
                                            End If
                                                
                                            LugarMetodo = LugarMetodo + 1
                                                    
                                            ZOrden = ""
                                            ZClave1 = ZLote
                                            Call Ceros(ZClave1, 6)
                                            ZClave2 = Str$(LugarMetodo)
                                            Call Ceros(ZClave2, 2)
                                            ZClave = ZClave1 + ZClave2
                                            ZMetodo = ""
                                            ZExamen = ""
                                            ZValorNormalI = ""
                                            ZValorNormalII = ""
                                            ZValorPartidaI = ""
                                            ZValorPartidaII = ""
                                            ZObservacionesI = ""
                                            ZObservacionesII = ""
                                            ZObservacionesIII = "Version " + ZVersion
                                            ZObservacionesIV = ""
                                            ZObservacionesV = ""
                                            ZObservacionesVI = ""
                                            If Val(Wempresa) = 1 Then
                                                ZEmpresa = "Surfactan S.A."
                                                    Else
                                                ZEmpresa = "Pellital S.A."
                                            End If
                                            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                            ZFechaII = WFechaElaboracion
                                            ZExamenII = ""
                                                        
                                            ZSql = ""
                                            ZSql = ZSql + "INSERT INTO Certificado ("
                                            ZSql = ZSql + "Clave ,"
                                            ZSql = ZSql + "Partida ,"
                                            ZSql = ZSql + "Renglon ,"
                                            ZSql = ZSql + "Razon ,"
                                            ZSql = ZSql + "Orden ,"
                                            ZSql = ZSql + "Terminado ,"
                                            ZSql = ZSql + "Descripcion ,"
                                            ZSql = ZSql + "Fecha ,"
                                            ZSql = ZSql + "Cantidad ,"
                                            ZSql = ZSql + "Examen ,"
                                            ZSql = ZSql + "ValorPartidaI ,"
                                            ZSql = ZSql + "ValorPartidaII ,"
                                            ZSql = ZSql + "ValorNormalI ,"
                                            ZSql = ZSql + "ValorNormalII ,"
                                            ZSql = ZSql + "Observaciones1 ,"
                                            ZSql = ZSql + "Observaciones2 ,"
                                            ZSql = ZSql + "Observaciones3 ,"
                                            ZSql = ZSql + "Observaciones4 ,"
                                            ZSql = ZSql + "Observaciones5 ,"
                                            ZSql = ZSql + "Observaciones6 ,"
                                            ZSql = ZSql + "Metodo ,"
                                            ZSql = ZSql + "Empresa )"
                                            ZSql = ZSql + "Values ("
                                            ZSql = ZSql + "'" + ZClave + "',"
                                            ZSql = ZSql + "'" + ZLote + "',"
                                            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                            ZSql = ZSql + "'" + ZRazon + "',"
                                            ZSql = ZSql + "'" + ZOrden + "',"
                                            ZSql = ZSql + "'" + ZArticulo + "',"
                                            ZSql = ZSql + "'" + ZDesArticulo + "',"
                                            ZSql = ZSql + "'" + ZFecha + "',"
                                            ZSql = ZSql + "'" + ZCantidad + "',"
                                            ZSql = ZSql + "'" + ZExamen + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                            ZSql = ZSql + "'" + ZValorNormalI + "',"
                                            ZSql = ZSql + "'" + ZValorNormalII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesI + "',"
                                            ZSql = ZSql + "'" + ZObservacionesII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                            ZSql = ZSql + "'" + ZMetodo + "',"
                                            ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                            spCertificado = ZSql
                                            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                
                                        Loop
                                                
                                        Listado.WindowTitle = "Certificado de Analisis"
                                        Listado.WindowTop = 0
                                        Listado.WindowLeft = 0
                                        Listado.WindowWidth = Screen.Width
                                        Listado.WindowHeight = Screen.Height
                        
                                        Listado.Destination = 1
                                        Listado.Destination = 0
                                                
                                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                                            If WIdioma = 0 Then
                                                Listado.ReportFileName = "CertificadoNuevo.rpt"
                                                    Else
                                                Listado.ReportFileName = "CertificadoNuevoIngles.rpt"
                                            End If
                                                Else
                                            Listado.ReportFileName = "CertificadoPelli.rpt"
                                        End If
                                                    
                                        DbConnect = db.Connect
                                        DSQ = getDatabase(DbConnect)
                        
                                        Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                                        + "From " _
                                                        + DSQ + ".dbo.Certificado Certificado " _
                                                        + "Where " _
                                                        + "Certificado.Partida >= 0 AND " _
                                                        + "Certificado.Partida <= 999999"
                                                        
                                        Listado.Destination = 1
                                        Listado.Destination = 0
                                        Listado.CopiesToPrinter = 1
                        
                                        Listado.Connect = Connect()
                                        Listado.Action = 1
                                                
                                    End If
                                          
                                End If
                                    
                            Next ZCiclo
                            
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    Wempresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                            
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                End If
            End If
                
        Next ZZCiclo
        
    Next DA
    
End Sub




Private Sub Verifica_Impresora()
    Dim JobInfo As JOB_INFO_1
    Dim hPrinter As Long, res As Boolean, L As Long
    Dim BytesNecesarios As Long, InfoDevueltos As Long
    Dim buffer() As Byte, PDefault As PRINTER_DEFAULTS, pos As Long, aux As String

    'abro la impresora por defecto
    res = OpenPrinter(Printer.DeviceName, hPrinter, PDefault)
    If Not res Then Exit Sub
    Do
        'miro cuantos bytes necesito
        ReDim buffer(1)
        L = EnumJobs(hPrinter, 0, 9999, 1, buffer(0), 1, BytesNecesarios, InfoDevueltos)
        If BytesNecesarios = 0 Then Exit Sub
    Loop
    'cierro la impresora
    res = ClosePrinter(hPrinter)
    
End Sub

Private Sub Verifica_ImpresoraII()
    Dim JobInfo As JOB_INFO_1
    Dim hPrinter As Long, res As Boolean, L As Long
    Dim BytesNecesarios As Long, InfoDevueltos As Long
    Dim buffer() As Byte, PDefault As PRINTER_DEFAULTS, pos As Long, aux As String

    'abro la impresora por defecto
    res = OpenPrinter("DocPdf", hPrinter, PDefault)
    If Not res Then Exit Sub
    Do
        'miro cuantos bytes necesito
        ReDim buffer(1)
        L = EnumJobs(hPrinter, 0, 9999, 1, buffer(0), 1, BytesNecesarios, InfoDevueltos)
        If BytesNecesarios = 0 Then Exit Sub
    Loop
    'cierro la impresora
    res = ClosePrinter(hPrinter)
    
End Sub


