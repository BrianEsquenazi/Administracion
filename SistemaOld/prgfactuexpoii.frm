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
   Begin VB.CommandButton DatosAdicionales 
      Caption         =   "Datos Adicionales"
      Height          =   450
      Left            =   9120
      TabIndex        =   67
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Cae 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   64
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame CargaAdicional 
      Height          =   4815
      Left            =   480
      TabIndex        =   45
      Top             =   1800
      Visible         =   0   'False
      Width           =   6495
      Begin VB.ComboBox Idioma 
         Height          =   315
         Left            =   1080
         TabIndex        =   69
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton AceptaAdicional 
         Caption         =   "Confirma Datos"
         Height          =   570
         Left            =   3000
         TabIndex        =   68
         Top             =   3960
         Width           =   1215
      End
      Begin VB.ComboBox CipLista 
         Height          =   315
         Left            =   5040
         TabIndex        =   66
         Text            =   "Combo1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Envio1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   55
         Text            =   " "
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox Envio2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   54
         Text            =   " "
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox Pago1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox Pago2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   52
         Text            =   " "
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox NroOrden 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   51
         Text            =   " "
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Consignatario 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   49
         Text            =   " "
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Marca 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   48
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox Dolar2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   3240
         Width           =   5055
      End
      Begin VB.TextBox Dolar1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   2880
         Width           =   5055
      End
      Begin MSMask.MaskEdBox fecorden 
         Height          =   255
         Left            =   4200
         TabIndex        =   50
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
         TabIndex        =   70
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label rrr 
         Caption         =   "Marca"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Envio"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Nro orden"
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Orden"
         Height          =   375
         Left            =   3000
         TabIndex        =   59
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Consignatario"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Incoterms"
         Height          =   375
         Left            =   4080
         TabIndex        =   57
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Dolar"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.CommandButton ReImpresion 
      Caption         =   "ReImpresion"
      Height          =   495
      Left            =   10200
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Graba1 
      Caption         =   "Fc. Exportacion"
      Height          =   495
      Left            =   10200
      TabIndex        =   37
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   32
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      Height          =   495
      Left            =   9120
      TabIndex        =   30
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   8160
      TabIndex        =   21
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox Flete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   44
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Seguro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   43
         Text            =   " "
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Flete"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Seguro"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Interes"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Dto."
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1200
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
      TabIndex        =   42
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label23 
      Caption         =   "Cae"
      Height          =   375
      Left            =   3360
      TabIndex        =   65
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   5640
      TabIndex        =   31
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
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.
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
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private dada As String
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
Private WTexto1 As String
Private WTexto2 As String
Private Auxiliar(100, 50) As String

Dim ZZControlLote(100, 60) As String
Dim ControlLote(12, 2) As String
Dim ControlEnvase(12, 2) As String

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

Dim ZZVector(100, 10) As String
Dim ZZImpre(100, 10) As String
Dim ZZCampo1 As String
Dim ZZCampo2 As String


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

Private Sub AceptaAdicional_Click()
    CargaAdicional.Visible = False
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For A = 0 To 5
        
        Suma = A * 10
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
            
    Next A
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0

    Rem If Val(Paridad.Text) <> 0 Then
    Rem     WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    If WTasa <> 0 Then
        WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
        Call Redondeo(WImpoInteres)
        WNeto = WNeto + WImpoInteres
    End If
    
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
        Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2 + Val(Seguro.Text) + Val(Flete.Text)
    
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

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

Private Sub ConsultaPedido_Click()
    ZZProcesoFactura = 3
    PrgSeleccionaPedido.Show
End Sub

Private Sub DatosAdicionales_Click()
    CargaAdicional.Visible = True
    Marca.SetFocus
End Sub

Private Sub Form_Activate()
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
    End Select
    
    If Idioma.ListIndex < 1 Then
        m$ = "Codigo idioma incorrexto"
        G% = MsgBox(m$, 0, "Factura de Exportacion")
        Exit Sub
    End Select
    
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
    
    If Val(Flete.Text) = 0 Then
        T$ = "Factura de Exportacion"
        Estilo = Estilo = vbYesNo + vbCritical + vbDefaultButton2
        m$ = "No se informo importe de flete, desea continuar "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% <> 6 Then
            Exit Sub
        End If
    End If


    Erase ZZVector
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    ZZImpreTotal = 0
    ZZImpreTotalBruto = 0
    ZZImpreTotalNeto = 0
    
    ZZDesde = 1
    
    For A = 0 To 5
    
        Suma = A * 10
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
                            ZPeso = rstEnvases!Peso
                            ZTipo = rstEnvases!Tipo
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
            
            If Cantidad <> SumaLote Then
                m$ = Articulo + " Verifique la discriminacion de lotes"
                G% = MsgBox(m$, 0, "Emision de facturas")
                Exit Sub
            End If
            
        Next iRow
    Next A

    Call Calcula_Click
    
    Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
    Rem     WImporte = WNeto
    Rem     WNeto = WNeto / 1.21
    Rem     Call Redondeo(WNeto)
    Rem     WIva1 = WImporte - WNeto
    Rem     WIva2 = 0
    Rem End If
    
    Rem Call Graba_FE
    
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
    
    
    
    ZZImpreNumero = Auxi
    ZZCae = Cae.Text
    ZZFechaCae = FechaCae.Text
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
                + WEmpresa + "','" _
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
    ZSql = ZSql & "ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql & "Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql & "FechaCae = " + "'" + FechaCae.Text + "',"
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
    
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    For A = 0 To 5
    
        Suma = A * 10
        DBGrid1.FirstRow = Suma
    
        For iRow = 0 To 9
        
            WRenglon = WRenglon + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            
            DBGrid1.Col = 1
            ZZImpreTerminado = DBGrid1.Text
            
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
            
                If Left$(Articulo, 2) = "PT" Then
                
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
                            If Left$(Articulo, 2) = "DW" Then
                                WLinea = 17
                                    Else
                                If Left$(Articulo, 2) = "DQ" Then
                                    WLinea = 22
                                        Else
                                    WLinea = 5
                                End If
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
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
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
                If WTipoProDy <> "PT" Then
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
                
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                    Select Case WTipoPedido
                        Case "FA", "PT", "BI"
                            XEmpresa = WEmpresa
                            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                                Select Case WTipoPedido
                                    Case "PG", "CO"
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "FA"
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                        WEmpresa = "0007"
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

            End If
                                
        Next iRow
        
    Next A
    
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
        If WTipoProDy <> "PT" Then
            XTipoproDy = "M"
            XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                Else
            XTipoproDy = "T"
            XArticuloDy = "  -   -   "
        End If
        
        If XTipoproDy = "M" Then
        
            XEmpresa = WEmpresa
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "FA"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0007"
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
                
            XEmpresa = WEmpresa
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "FA"
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0007"
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
                    WFecha1 = rstPreciosMp!fecha2
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
                    WFecha1 = rstPrecios!fecha2
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
    
    For Ciclo = 1 To 100
    
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
    
    Call Impresion_Remito
    
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

Private Sub Limpia_Click()

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    CipLista.ListIndex = 0
    Idioma.ListIndex = 0
    
    For A = 0 To 5
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    Seguro.Text = ""
    Flete.Text = ""
    
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
    CipLista.AddItem "DDU"
    CipLista.AddItem "DDP"
    
    CipLista.ListIndex = 0
    
    
    
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
ReDim UserData(0 To 4, 0 To 80)

mTotalRows& = 80

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
For i = 0 To 4
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
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    Seguro.Text = ""
    Flete.Text = ""
    
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

    For A = 0 To 5
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    WNeto = 0
    WSeguro = 0
    WFlete = 0
    
    Erase Auxiliar
    Erase ZZControlLote
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Canti = !Cantidad1
                
                    Select Case Mid$(!Terminado, 1, 2)
                        Case "Z2"
                            WSeguro = WSeguro + (!Precio * !Cantidad)
                                                            
                        Case "Z1"
                            WFlete = WFlete + (!Precio * !Cantidad)
                    
                        Case Else
                            Renglon = Renglon + 1
            
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = !Terminado
                            Auxi1 = !Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                
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
                            ZZControlLote(Renglon, 11) = Str$(!lote6)
                            ZZControlLote(Renglon, 12) = Str$(!CantiLote6)
                            ZZControlLote(Renglon, 13) = Str$(!lote7)
                            ZZControlLote(Renglon, 14) = Str$(!CantiLote7)
                            ZZControlLote(Renglon, 15) = Str$(!lote8)
                            ZZControlLote(Renglon, 16) = Str$(!CantiLote8)
                            ZZControlLote(Renglon, 17) = Str$(!lote9)
                            ZZControlLote(Renglon, 18) = Str$(!CantiLote9)
                            ZZControlLote(Renglon, 19) = Str$(!lote10)
                            ZZControlLote(Renglon, 20) = Str$(!CantiLote10)
                            ZZControlLote(Renglon, 21) = Str$(!lote11)
                            ZZControlLote(Renglon, 22) = Str$(!CantiLote11)
                            ZZControlLote(Renglon, 23) = Str$(!lote12)
                            ZZControlLote(Renglon, 24) = Str$(!CantiLote12)
                               
                            ZZControlLote(Renglon, 31) = Str$(!Env1)
                            ZZControlLote(Renglon, 32) = Str$(!CantiEnv1)
                            ZZControlLote(Renglon, 33) = Str$(!Env2)
                            ZZControlLote(Renglon, 34) = Str$(!CantiLote2)
                            ZZControlLote(Renglon, 35) = Str$(!lote3)
                            ZZControlLote(Renglon, 36) = Str$(!CantiLote3)
                            ZZControlLote(Renglon, 37) = Str$(!lote4)
                            ZZControlLote(Renglon, 38) = Str$(!CantiLote4)
                            ZZControlLote(Renglon, 39) = Str$(!lote5)
                            ZZControlLote(Renglon, 40) = Str$(!CantiLote5)
                            ZZControlLote(Renglon, 41) = Str$(!lote6)
                            ZZControlLote(Renglon, 42) = Str$(!CantiLote6)
                            ZZControlLote(Renglon, 43) = Str$(!lote7)
                            ZZControlLote(Renglon, 44) = Str$(!CantiLote7)
                            ZZControlLote(Renglon, 45) = Str$(!lote8)
                            ZZControlLote(Renglon, 46) = Str$(!CantiLote8)
                            ZZControlLote(Renglon, 47) = Str$(!lote9)
                            ZZControlLote(Renglon, 48) = Str$(!CantiLote9)
                            ZZControlLote(Renglon, 49) = Str$(!lote10)
                            ZZControlLote(Renglon, 50) = Str$(!CantiLote10)
                            ZZControlLote(Renglon, 51) = Str$(!lote11)
                            ZZControlLote(Renglon, 52) = Str$(!CantiLote11)
                            ZZControlLote(Renglon, 53) = Str$(!lote12)
                            ZZControlLote(Renglon, 54) = Str$(!CantiLote12)

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
    
    WRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
        If Left$(Auxi1, 2) <> "PT" Then
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

    For A = 0 To 5
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
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
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!PrecioUs)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("###,###.##", dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!PrecioUs)
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    XRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
        
        If Left$(Auxi1, 2) <> "PT" Then
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
                    DBGrid1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
        End Select
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
            Paridad.Text = Pusing("###,###.##", Paridad.Text)
            
            WSeguro = IIf(IsNull(rstCtacte!Seguro), "0", rstCtacte!Seguro)
            WFlete = IIf(IsNull(rstCtacte!Flete), "0", rstCtacte!Flete)
            Seguro.Text = Str$(WSeguro)
            Seguro.Text = Pusing("###,###.##", Seguro.Text)
            Flete.Text = Str$(WFlete)
            Flete.Text = Pusing("###,###.##", Flete.Text)
            
            Cae.Text = IIf(IsNull(rstCtacte!Cae), "", rstCtacte!Cae)
            Marca.Text = IIf(IsNull(rstCtacte!Marca), "", rstCtacte!Marca)
            Envio1.Text = IIf(IsNull(rstCtacte!Envio1), "", rstCtacte!Envio1)
            Envio2.Text = IIf(IsNull(rstCtacte!Envio2), "", rstCtacte!Envio2)
            Pago1.Text = IIf(IsNull(rstCtacte!Pago1), "", rstCtacte!Pago1)
            Pago2.Text = IIf(IsNull(rstCtacte!Pago2), "", rstCtacte!Pago2)
            NroOrden.Text = IIf(IsNull(rstCtacte!NroOrden), "", rstCtacte!NroOrden)
            fecorden.Text = IIf(IsNull(rstCtacte!fecorden), "", rstCtacte!fecorden)
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
                WVendedor = rstCliente!Vendedor
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
                WVendedor = rstCliente!Vendedor
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
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
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
                A% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            A% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub ReImpresion_Click()
    Call Impresion_Remito
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

        If Val(WEmpresa) = 1 Then
            Rem Open "LPT1" For Output As #99
            Open "dada.txt" For Output As #99
                Else
            If Val(WEmpresa) = 10 Then
                Open "dada.txt" For Output As #99
                    Else
                Open "dada.txt" For Output As #99
                Rem Open "LPT1" For Output As #99
            End If
        End If

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
        
        For A = 0 To 5
        
            Suma = A * 10
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
            
        Next A
        
        XPasa = 0
        
        
        For A = 1 To 99
        
                Producto = WImpresion(A, 1)
                Descri = WImpresion(A, 2)
                Precio = Val(Alinea("##,###.##", WImpresion(A, 4)))
                Cantidad = Val(WImpresion(A, 5))
                Cantidad1 = Val(WImpresion(A, 6))
                WTipo = WImpresion(A, 7)
                WNumero = WImpresion(A, 8)
                Bruto = Val(WImpresion(A, 9))
                    
                If Cantidad <> 0 Then
                
                        If WNumero = WImpresion(A + 1, 8) And XPasa = 0 Then
                        
                            Print #99, Tab(2); Alinea("###", Str$(1));
                            Print #99, Tab(8); WTipo;
                            Print #99, Tab(12); WNumero;
                            Print #99, Tab(22); "Palet conteniendo : ";
                            Print #99, Tab(60); Alinea("#####.##", Str$(24));
                            Suma2 = Suma2 + 24
                            
                            XPasa = 1
                            
                            XCanti = XEnvase(A, 2)
                            Call Ceros(XCanti, 2)
                            spEnvase = "ConsultaEnvases " + "'" + XEnvase(A, 1) + "'"
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
                                    
                            If WNumero <> WImpresion(A + 1, 8) Then
                                XPasa = 0
                            End If
                                
                            If WNumero = WImpresion(A - 1, 8) Then
                            
                                XCanti = XEnvase(A, 2)
                                Call Ceros(XCanti, 2)
                                spEnvase = "ConsultaEnvases " + "'" + XEnvase(A, 1) + "'"
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
            
        Next A
        
        
        
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

    Impre = 0
    ZSql = "DELETE ImpreRemito"
    spImpreRemito = ZSql
    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
    
    For A = 0 To 5
    
        Suma = A * 10
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
        
    Next A
    
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
        Cip.SetFocus
    End If
End Sub

Private Sub Cip_Keypress(KeyAscii As Integer)
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






Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub






