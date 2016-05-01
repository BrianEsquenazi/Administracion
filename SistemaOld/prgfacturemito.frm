VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactuRemito 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Remitos a Facturar"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.TextBox Remito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   40
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Envase1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   34
      Text            =   " "
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Envase2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   33
      Text            =   " "
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Envase3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   32
      Text            =   " "
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Envase4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   31
      Text            =   " "
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Envase5 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   30
      Text            =   " "
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Canti1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   29
      Text            =   " "
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Canti2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   28
      Text            =   " "
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Canti3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   27
      Text            =   " "
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Canti4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   26
      Text            =   " "
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Canti5 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   25
      Text            =   " "
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11040
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11280
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
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
      Left            =   6000
      TabIndex        =   21
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame PantaMotivo 
      Height          =   1815
      Left            =   480
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   10335
      Begin VB.ComboBox ConceptoAtraso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox DescriMotivo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   19
         Top             =   720
         Width           =   9855
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "MOTIVO DE RETRASO DE CUMPLIMIENTO DEL PEDIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   9735
      End
   End
   Begin VB.CommandButton ReImpre 
      Caption         =   "ReImpresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   17
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   14
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
      Left            =   3360
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Cliente 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   7
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9120
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10200
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10560
      TabIndex        =   1
      Top             =   1320
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
      Height          =   1980
      ItemData        =   "prgfacturemito.frx":0000
      Left            =   4200
      List            =   "prgfacturemito.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4215
      Left            =   240
      OleObjectBlob   =   "prgfacturemito.frx":0015
      TabIndex        =   2
      Top             =   1680
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Label Label9 
      Caption         =   "Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3240
      TabIndex        =   41
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   39
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Descri2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   38
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   36
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   35
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Movimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgFactuRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.


Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim MSubject As String
Dim MBody As String
Dim MAttach As String
Dim MAttachI As String
Dim MAttachII As String
Dim MAttachIII As String
Dim MAttachIV As String
Dim MAttachV As String
Dim AllPath As String
Dim WNombreEmail As String
Dim ZZZSuma As Integer
Dim ZZZSumaII As Integer

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
Private WTotal As Double
Private WImpoDto As Double
Private XImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WAdicional As Double
Private ZAdicional As String
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
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private WAceptada As String
Private Stk(19, 4) As String
Private Envase(5, 2) As String
Private parcial As String
Private ZImpreStk(20, 4) As String

Private Auxiliar(100, 30) As String
Private RestaPedido(100, 3) As String
Private ClavePedido(100)

Private BajaLote(12, 2) As String
Private XLote(100, 80) As String
Dim CargaEmpresa(12, 2) As String

Dim rstNumero As Recordset
Dim spNumero As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstEstadisticaLote As Recordset
Dim spEstadisticaLote As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String
Dim rstCertificado As Recordset
Dim spCertificado As String

Dim XParam As String
Dim ZZImpreNumero As String

Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WSaldo4 As Double
Dim WSaldo5 As Double
Dim WSaldo6 As Double
Dim WSaldo7 As Double
Dim WSaldo8 As Double
Dim WSaldo9 As Double
Dim WSaldo10 As Double
Dim WSaldo11 As Double
Dim WSaldo12 As Double

Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim XSaldo4 As String
Dim XSaldo5 As String
Dim XSaldo6 As String
Dim XSaldo7 As String
Dim XSaldo8 As String
Dim XSaldo9 As String
Dim XSaldo10 As String
Dim XSaldo11 As String
Dim XSaldo12 As String

Dim ZZCampo1 As String
Dim ZZCampo2 As String

Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim Compara As Double
Dim ZZIntervencion As String
Dim ZLugarFicha As Integer

Private WTipoPedido As String

Dim ZLote1 As String
Dim ZCantidad1 As String
Dim ZLote2 As String
Dim ZCantidad2 As String
Dim ZLote3 As String
Dim ZCantidad3 As String
Dim ZLote4 As String
Dim ZCantidad4 As String
Dim ZLote5 As String
Dim ZCantidad5 As String
Dim ZLote6 As String
Dim ZCantidad6 As String
Dim ZLote7 As String
Dim ZCantidad7 As String
Dim ZLote8 As String
Dim ZCantidad8 As String
Dim ZLote9 As String
Dim ZCantidad9 As String
Dim ZLote10 As String
Dim ZCantidad10 As String
Dim ZLote11 As String
Dim ZCantidad11 As String
Dim ZLote12 As String
Dim ZCantidad12 As String

Dim ZEnv1 As String
Dim ZCantiEnv1 As String
Dim ZEnv2 As String
Dim ZCantiEnv2 As String
Dim ZEnv3 As String
Dim ZCantiEnv3 As String
Dim ZEnv4 As String
Dim ZCantiEnv4 As String
Dim ZEnv5 As String
Dim ZCantiEnv5 As String
Dim ZEnv6 As String
Dim ZCantiEnv6 As String
Dim ZEnv7 As String
Dim ZCantiEnv7 As String
Dim ZEnv8 As String
Dim ZCantiEnv8 As String
Dim ZEnv9 As String
Dim ZCantiEnv9 As String
Dim ZEnv10 As String
Dim ZCantiEnv10 As String
Dim ZEnv11 As String
Dim ZCantiEnv11 As String
Dim ZEnv12 As String
Dim ZCantiEnv12 As String

Dim ControlLote(12, 2) As String

Dim WSal As Double
Dim WVector(10000, 4) As String
Dim ZClave  As String
Dim ZTipo As String
Dim ZNumero As String
Dim ZRenglon As String
Dim Renglon As Integer
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim ZZValor1 As Double
Dim ZZValor2 As Double
Dim ZZImpreDespa(100, 5) As String
Dim ZZImpreDespaII(100, 5) As String
Dim ZZVector(100, 5) As String
    
Dim ZZBusca(10000) As String
Dim ZZLugarBusca As Integer
 
 
Dim DiaFeriado(100) As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer

Dim ZZLote As String

Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String
Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 4) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String
Dim ZImpreFicha(100) As String

Dim ZZEnvase(10) As String
Dim ZZCanti(10) As String

Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZVersionPedido As Integer
Dim ZVersionAtraso As Integer
Dim ZSedronar As Integer
Dim ZNroSedronar As String

Dim ZZPasaImpre As Integer
Dim FF As Integer
Dim ZZGrabaFactura As String
Dim ZZImpreBarraI As String
Dim ZZImpreBarraII As String














Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double

Private ZEmailFactura As String

Private ZZEnviaPdf(100, 5) As String
Private ZZEnviaPdfII(100, 5) As String
Private ZZLugarEnvia As Integer
Private ZZLugarEnviaII As Integer

Private WCodIb As Integer
Private WCodIbTucu As Integer
Private WCodIbCiudad As Integer

Private WImpoIb As Double
Private WImpoIbTucu As Double
Private WImpoIbCiudad As Double
Private WPorceCm05Tucu As Double

Private WImpoPorceIb As Double
Private WImpoPorceIbTucu As Double
Private WImpoPorceIbCiudad As Double

Private ZZPorceIbCaba As Double

Private WPorceIb As Double

Dim ZZFecha As String
Dim ZZDias As Integer
Dim ZZVto As String
Dim ZDolarEspecial As Integer

Dim VectorCosto(100, 3) As String

Dim ZZRemito As String

Dim ZZRuta As String
Dim ZZRutaII As String
Dim ZZEstado As String
Dim ZZEstadoII As String
Dim ZZNombreArchi As Integer
Dim ZZNombreArchiII As String

Dim ZZZZPrecio As Double



Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 3
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = "S"
    
    For ZZCiclo = 1 To 80
        XLote(WRow, ZZCiclo) = ""
    Next ZZCiclo
    
End Sub



Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    RetVal = Shell("cmd.exe /c Taskkill /f /IM AcroRd32.exe", 6)
    
    PrgFactuRemito.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub ConsultaPedido_Click()
    ZZProcesoFactura = 4
    PrgSeleccionaPedido.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    If Val(Wempresa) = 1 Then
        OPEN_FILE_Ctacte8
        OPEN_FILE_Numero8
        OPEN_FILE_Esta8
    End If
    
    If ZZProcesoFactura = 99 And Val(Pedido.Text) <> 0 Then
        Call Pedido_KeyPress(13)
        Call Fecha_Keypress(13)
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
    End If
    
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    If Val(Remito.Text) = 0 Then
        m$ = "Se debe informar numero de remito"
        CA% = MsgBox(m$, 0, "Emision de Facturas")
        Exit Sub
    End If
    
    If Val(Wempresa) = 1 And Cliente.Text = "P00005" Then
        Rem Call Verifica_Lote
        WEstado = "S"
            Else
        Call Verifica_Lote
    End If
    
    If WEstado = "N" Then
        Call Limpia_Click
        Exit Sub
    End If
    
    If Val(Remito.Text) = 0 Then
        m$ = "Se debe informar numero de remito"
        CA% = MsgBox(m$, 0, "Emision de Facturas")
        Exit Sub
    End If
    
    Call Verifica_Certificado
    
    If WEstado = "N" Then
        Call Limpia_Click
        Exit Sub
    End If
    
    WTipo = "10"
    WNumero = Numero.Text
    WRenglon = "01"
    WCliente = Cliente.Text
    WFecha = Fecha.Text
    Wvencimiento = Fecha.Text
    WVencimiento1 = Fecha.Text
    WEstado = "0"
        
    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WImpre = "FR"
    XSeguro = ""
    XFlete = ""
    WPedido = Pedido.Text
    WRemito = Remito.Text
    WOrden = ""
    WParidad = ""
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
    WClave = "10" + Auxi + "01"
    XEmpresa = "1"
    WDate = Date$
    
    XTotal = ""
    XTotalUs = ""
    XSaldo = ""
    XSaldoUs = ""
    XNet = ""
    XIva1 = ""
    XIva2 = ""
    XSeguro = ""
    XFlete = ""
    XImpoIb = ""
    
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
                + XImpoIb + "'"
                    
    spCtacte = "AltaCtacte " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZCliente = Cliente.Text
    
    Erase Auxiliar
    Erase RestaPedido
    Auxi = 0
    
    Suma = 0
    Renglon = 0
    Renglon1 = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
        
            Suma = Suma + 1
            WRenglon = WRenglon + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            WTipoProDy = Left$(Articulo, 2)
            Rem WBase = Val(Right$(Articulo, 3))
            Rem If WBase <= 5 Then
            Rem     Articulo = Left$(Articulo, 7) + "100"
            Rem End If
                
            DBGrid1.Col = 1
            ZZDescriArticulo = DBGrid1.Text
                
            Precio = 0
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            DBGrid1.Col = 5
            RestaCantidad = Val(DBGrid1.Text)
                
            If Cantidad <> 0 Then
            
                If WTipoProDy = "PT" Then
                
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
                        
                Auxi1 = Str$(Val(Numero.Text) + 900000)
                Call Ceros(Auxi1, 8)
                WTipo = "01"
                WNumero = Str$(Val(Numero.Text) + 900000)
                XRenglon = Str$(Renglon)
                WArticulo = Articulo
                XXCantidad = Str$(Cantidad)
                XPrecioUs = ""
                XPrecio = ""
                XImporteUs = ""
                XImporte = ""
                WCliente = Cliente.Text
                WParidad = ""
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
                
                WLote1 = XLote(Suma, 1)
                WLote2 = XLote(Suma, 3)
                Wlote3 = XLote(Suma, 5)
                WLote4 = XLote(Suma, 7)
                WLote5 = XLote(Suma, 9)
                WLote6 = XLote(Suma, 11)
                WLote7 = XLote(Suma, 13)
                WLote8 = XLote(Suma, 15)
                WLote9 = XLote(Suma, 17)
                WLote10 = XLote(Suma, 19)
                WLote11 = XLote(Suma, 21)
                WLote12 = XLote(Suma, 23)
                
                WImpo = Val(XLote(Suma, 2))
                WCanti1 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 4))
                WCanti2 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 6))
                WCanti3 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 8))
                WCanti4 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 10))
                WCanti5 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 12))
                WCanti6 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 14))
                WCanti7 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 16))
                WCanti8 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 18))
                WCanti9 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 20))
                WCanti10 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 22))
                WCanti11 = Str$(WImpo)
                WImpo = Val(XLote(Suma, 24))
                WCanti12 = Str$(WImpo)
                
                WLoteAdicional = ""
                For ZZCiclo = 11 To 23 Step 2
                    ZZCampo1 = XLote(Suma, ZZCiclo)
                    ZZCampo2 = XLote(Suma, ZZCiclo + 1)
                    Call Ceros(ZZCampo1, 8)
                    Call Ceros(ZZCampo2, 6)
                    WLoteAdicional = WLoteAdicional + ZZCampo1 + ZZCampo2
                Next ZZCiclo
                
                If Left$(WArticulo, 2) = "PT-5" Then
                    If Val(Wempresa) = 1 And Cliente.Text = "P00005" Then
                        WLote1 = ""
                    End If
                End If
                
                If Left$(WArticulo, 2) = "PT-5" Then
                    XMarca = "X"
                End If
                
                
                XEnv1 = XLote(Suma, 31)
                XCantiEnv1 = XLote(Suma, 32)
                XEnv2 = XLote(Suma, 33)
                XCantiEnv2 = XLote(Suma, 34)
                XEnv3 = XLote(Suma, 35)
                XCantiEnv3 = XLote(Suma, 36)
                XEnv4 = XLote(Suma, 37)
                XCantiEnv4 = XLote(Suma, 38)
                XEnv5 = XLote(Suma, 39)
                XCantiEnv5 = XLote(Suma, 40)
                XEnv6 = XLote(Suma, 41)
                XCantiEnv6 = XLote(Suma, 42)
                XEnv7 = XLote(Suma, 43)
                XCantiEnv7 = XLote(Suma, 44)
                XEnv8 = XLote(Suma, 45)
                XCantiEnv8 = XLote(Suma, 46)
                XEnv9 = XLote(Suma, 47)
                XCantiEnv9 = XLote(Suma, 48)
                XEnv10 = XLote(Suma, 49)
                XCantiEnv10 = XLote(Suma, 50)
                XEnv11 = XLote(Suma, 51)
                XCantiEnv11 = XLote(Suma, 52)
                XEnv12 = XLote(Suma, 53)
                XCantiEnv12 = XLote(Suma, 54)
                
                WEnvAdicional = ""
                For ZZCiclo = 41 To 53 Step 2
                    ZZCampo1 = XLote(Suma, ZZCiclo)
                    ZZCampo2 = XLote(Suma, ZZCiclo + 1)
                    Call Ceros(ZZCampo1, 4)
                    Call Ceros(ZZCampo2, 4)
                    WEnvAdicional = WEnvAdicional + ZZCampo1 + ZZCampo2
                Next ZZCiclo
                
                If WCliente = "G00007" And Left$(WArticulo, 8) = "PT-07581" Then
                    XLinea = "18"
                End If
                If WCliente = "G00065" And Left$(WArticulo, 8) = "PT-07581" Then
                    XLinea = "18"
                End If
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
                             + XXCantidad + "','" + XPrecio + "','" _
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
                
                If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                    If Cliente.Text <> "P00005" Then

                        Select Case WTipoPedido
                            Case "FA", "PT", "BI", "TA"
                        
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
                                        Case "TA"
                                            Wempresa = "0003"
                                            txtOdbc = "Empresa03"
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
                                    + XXCantidad + "','" + XPrecio + "','" _
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
                        
                            Else
                        
                        If Left$(WArticulo, 4) <> "PT-5" Then
                        
                            Select Case WTipoPedido
                                Case "FA", "PT", "BI", "TA"
                        
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
                                            Case "TA"
                                                Wempresa = "0003"
                                                txtOdbc = "Empresa03"
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
                                        + XXCantidad + "','" + XPrecio + "','" _
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
                        
                    End If
                End If
                
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
                Auxiliar(Renglon, 29) = RestaCantidad
                Auxiliar(Renglon, 30) = ZZDescriArticulo
                    
            End If
            
            If RestaCantidad <> 0 Then
                Renglon1 = Renglon1 + 1
                RestaPedido(Renglon1, 1) = Articulo
                RestaPedido(Renglon1, 2) = RestaCantidad
                RestaPedido(Renglon1, 3) = ClavePedido(WRenglon)
            End If
                                    
        Next iRow
        
    Next a

    For DA = 1 To Renglon
    
        Articulo = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        Precio = Auxiliar(DA, 3)
        WRenglon = Auxiliar(DA, 4)
        WLote1 = Auxiliar(DA, 5)
        WCanti1 = Auxiliar(DA, 6)
        WLote2 = Auxiliar(DA, 7)
        WCanti2 = Auxiliar(DA, 8)
        Wlote3 = Auxiliar(DA, 9)
        WCanti3 = Auxiliar(DA, 10)
        WLote4 = Auxiliar(DA, 11)
        WCanti4 = Auxiliar(DA, 12)
        WLote5 = Auxiliar(DA, 13)
        WCanti5 = Auxiliar(DA, 14)
        WLote6 = Auxiliar(DA, 15)
        WCanti6 = Auxiliar(DA, 16)
        WLote7 = Auxiliar(DA, 17)
        WCanti7 = Auxiliar(DA, 18)
        WLote8 = Auxiliar(DA, 19)
        WCanti8 = Auxiliar(DA, 20)
        WLote9 = Auxiliar(DA, 21)
        WCanti9 = Auxiliar(DA, 22)
        WLote10 = Auxiliar(DA, 23)
        WCanti10 = Auxiliar(DA, 24)
        WLote11 = Auxiliar(DA, 25)
        WCanti11 = Auxiliar(DA, 26)
        WLote12 = Auxiliar(DA, 27)
        WCanti12 = Auxiliar(DA, 28)
        
        RestaCantidad = Auxiliar(DA, 29)
        
        WTipoProDy = Left$(Articulo, 2)
        If WTipoProDy <> "PT" Then
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
                    Case "TA"
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
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
                WPedido = Str$(rstArticulo!Venta - RestaCantidad)
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
            
            BajaLote(1, 1) = WLote1
            BajaLote(1, 2) = WCanti1
            BajaLote(2, 1) = WLote2
            BajaLote(2, 2) = WCanti2
            BajaLote(3, 1) = Wlote3
            BajaLote(3, 2) = WCanti3
            BajaLote(4, 1) = WLote4
            BajaLote(4, 2) = WCanti4
            BajaLote(5, 1) = WLote5
            BajaLote(5, 2) = WCanti5
            BajaLote(6, 1) = WLote6
            BajaLote(6, 2) = WCanti6
            BajaLote(7, 1) = WLote7
            BajaLote(7, 2) = WCanti7
            BajaLote(8, 1) = WLote8
            BajaLote(8, 2) = WCanti8
            BajaLote(9, 1) = WLote9
            BajaLote(9, 2) = WCanti9
            BajaLote(10, 1) = WLote10
            BajaLote(10, 2) = WCanti10
            BajaLote(11, 1) = WLote11
            BajaLote(11, 2) = WCanti11
            BajaLote(12, 1) = WLote12
            BajaLote(12, 2) = WCanti12
            
            For XDa = 1 To 12
            
                lote1 = BajaLote(XDa, 1)
                Cantidad1 = BajaLote(XDa, 2)
                
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
            Next XDa
            
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
                    Case "TA"
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            End If
                
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = Articulo
                WPedido = Str$(rstTerminado!Pedido - Val(RestaCantidad))
                WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
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
            
            BajaLote(1, 1) = WLote1
            BajaLote(1, 2) = WCanti1
            BajaLote(2, 1) = WLote2
            BajaLote(2, 2) = WCanti2
            BajaLote(3, 1) = Wlote3
            BajaLote(3, 2) = WCanti3
            BajaLote(4, 1) = WLote4
            BajaLote(4, 2) = WCanti4
            BajaLote(5, 1) = WLote5
            BajaLote(5, 2) = WCanti5
            BajaLote(6, 1) = WLote6
            BajaLote(6, 2) = WCanti6
            BajaLote(7, 1) = WLote7
            BajaLote(7, 2) = WCanti7
            BajaLote(8, 1) = WLote8
            BajaLote(8, 2) = WCanti8
            BajaLote(9, 1) = WLote9
            BajaLote(9, 2) = WCanti9
            BajaLote(10, 1) = WLote10
            BajaLote(10, 2) = WCanti10
            BajaLote(11, 1) = WLote11
            BajaLote(11, 2) = WCanti11
            BajaLote(12, 1) = WLote12
            BajaLote(12, 2) = WCanti12
            
            For XDa = 1 To 12
        
                lote1 = BajaLote(XDa, 1)
                Cantidad1 = BajaLote(XDa, 2)
    
                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
                
                lote1 = BajaLote(XDa, 1)
                Cantidad1 = BajaLote(XDa, 2)
                
                If WControla = 0 And Val(lote1) <> 0 Then
                    XParam = "'" + lote1 + "','" _
                                 + Articulo + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                    
                        WClave = rstHoja!Clave
                        WSaldo = Str$(rstHoja!Saldo - Val(Cantidad1))
                        WDate = Date$
                        rstHoja.Close
                        
                        XParam = "'" + WClave + "','" _
                                     + WDate + "','" _
                                     + WSaldo + "'"
                        spHoja = "ModificaHojaSaldo " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                        
                        XParam = "'" + Articulo + "','" _
                                     + lote1 + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
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
        
            Next XDa
            
            Call Conecta_Empresa
            
        End If
        
    Next DA
    
    For DA = 1 To Renglon1
    
        Articulo = RestaPedido(DA, 1)
        Auxi1 = RestaPedido(DA, 2)
        Auxi1 = Pusing("###,###.##", Auxi1)
        Cantidad = Auxi1
        WClavePedido = RestaPedido(DA, 3)
        
        XParam = "'" + Left$(WClavePedido, 6) + "','" _
                     + Right$(WClavePedido, 2) + "'"
        spPedido = "ConsultaPedido2 " + XParam
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            WFacturado = Str$(rstPedido!Facturado + Val(Cantidad))
            If Val(WFacturado) > rstPedido!Cantidad Then
                WFacturado = Str$(rstPedido!Cantidad)
            End If
            WClavePedido = rstPedido!Clave
            rstPedido.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " UltimoLote1 = Lote1" + ","
            ZSql = ZSql + " UltimoCantiLote1 = CantiLote1" + ","
            ZSql = ZSql + " UltimoLote2 = Lote2" + ","
            ZSql = ZSql + " UltimoCantiLote2 = CantiLote2" + ","
            ZSql = ZSql + " UltimoLote3 = Lote3" + ","
            ZSql = ZSql + " UltimoCantiLote3 = CantiLote3" + ","
            ZSql = ZSql + " UltimoLote4 = Lote4" + ","
            ZSql = ZSql + " UltimoCantiLote4 = CantiLote4" + ","
            ZSql = ZSql + " UltimoLote5 = Lote5" + ","
            ZSql = ZSql + " UltimoCantiLote5 = CantiLote5" + ""
            ZSql = ZSql + " Where Clave = " + "'" + WClavePedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            XParam = "'" + WClavePedido + "','" _
                         + WFacturado + "'"
            spPedido = "ModificaPedidoFacturas " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + " Remito = " + "'" + Remito.Text + "',"
        ZSql = ZSql + " CantidadFac = " + "'" + Cantidad + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClavePedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    Next DA
    
    BajaImpre = "N"
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then

        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WTerminado = !Terminado
                    XCodigo = Val(Mid$(WTerminado, 4, 5))
                    Canti = !Cantidad - !Facturado
                    
                    If Canti > 0 Then
                        BajaImpre = "S"
                    End If
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    If BajaImpre = "S" Then
    
        spPedido = "ModificaPedidoVersion " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        If Left$(WTerminado, 2) <> "PT" Then
            Select Case Left$(WTerminado, 2)
                Case "DY", "DS"
                    XTipoPro = "CO"
                Case Else
                    XTipoPro = "PT"
            End Select
                Else
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
                            XTipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If
                
        WLinea = 0
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
        
        Select Case WLinea
            Case 8
                XTipoPro = "PG"
            Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                XTipoPro = "FA"
            Case Else
        End Select
        
        Select Case XTipoPro
            Case "CO"
                XParam = "'" + Pedido.Text + "','" _
                            + "1" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            Case "FA"
                XParam = "'" + Pedido.Text + "','" _
                            + "4" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            Case "BI"
                XParam = "'" + Pedido.Text + "','" _
                            + "3" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            Case "PT"
                XParam = "'" + Pedido.Text + "','" _
                            + "2" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            Case "PG"
                XParam = "'" + Pedido.Text + "','" _
                            + "5" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                WMarca = "X"
                XParam = "'" + Pedido.Text + "','" _
                        + WMarca + "'"
                spPedido = "ModificaPedidoPigmentos " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            Case Else
                XParam = "'" + Pedido.Text + "','" _
                            + "0" + "'"
                spPedido = "ModificaPedidoTipoPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End Select
        
    End If
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Pedido SET "
    ZSql = ZSql & "MarcaAutorizacion = " + "'" + "0" + "',"
    ZSql = ZSql & "MarcaFactura = " + "'" + "0" + "'"
    ZSql = ZSql & " Where Pedido = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    spNumero = "ConsultaNumero " + "'" + "10" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        WCodigo = "10"
        WNumero = Numero.Text
        rstNumero.Close
        XParam = "'" + WCodigo + "','" _
                     + WNumero + "'"
        spNumero = "ModificaNumero " + XParam
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
    Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
    Rem Listado.Destination = 1
    Rem Listado.Action = 1
    
    Call Calcula_Saldo
    
    Erase Envase
    Envase(1, 1) = Envase1.Text
    Envase(2, 1) = Envase2.Text
    Envase(3, 1) = Envase3.Text
    Envase(4, 1) = Envase4.Text
    Envase(5, 1) = Envase5.Text
    
    Envase(1, 2) = Canti1.Text
    Envase(2, 2) = Canti2.Text
    Envase(3, 2) = Canti3.Text
    Envase(4, 2) = Canti4.Text
    Envase(5, 2) = Canti5.Text
    
    For XDa = 1 To 5
        For DA = 1 To 9
            If Val(Envase(XDa, 1)) = Val(Stk(DA, 1)) Then
                Stk(DA, 3) = Val(Envase(XDa, 2))
            End If
        Next DA
    Next XDa
    
    For DA = 1 To 9
        Stk(DA, 4) = Str$(Val(Stk(DA, 2)) + Val(Stk(DA, 3)))
    Next DA
    
    Renglon = 0
    
    For DA = 1 To 5
    
        If Val(Envase(DA, 2)) <> 0 Then
        
            Renglon = Renglon + 1
                
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                    
            Auxi1 = Str$(Val(Remito.Text))
            Call Ceros(Auxi1, 6)
                
            WTipo = "1"
            WCodigo = Str$(Val(Remito.Text) + 200000)
            WRenglon = Str$(Renglon)
            WCliente = Cliente.Text
            WFecha = Fecha.Text
            WEnvase = Envase(DA, 1)
            WCantidad = Envase(DA, 2)
            WMovimiento = "S"
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WClave = Auxi1 + Auxi
            
            XParam = "'" + WClave + "','" _
                    + WTipo + "','" _
                    + WCodigo + "','" _
                    + WRenglon + "','" _
                    + WFecha + "','" _
                    + WFechaord + "','" _
                    + WCliente + "','" _
                    + WEnvase + "','" _
                    + WMovimiento + "','" _
                    + WCantidad + "'"
                
            spMovenv = "AltaMovenv " + XParam
            Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next DA
    
    Call Impresion_Remito
    
    If Val(Wempresa) = 1 Then
        Call Impresion_Varios
        Call Envio_Email
    End If
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    PrgFactuRemito.Show
    Numero.SetFocus
        
    Exit Sub

WError:
    MsgBox Err.Description
    Resume Next
        
End Sub


Private Sub DescriMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(DescriMotivo.Text)) >= 10 And ConceptoAtraso.ListIndex > 0 Then
        
            ZZAtraso = "1"
        
            Sql1 = "Select Max(Numero) as [NumeroMayor]"
            Sql2 = " FROM Atraso"
            spAtraso = Sql1 + Sql2
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                ZZAtraso = Str$(rstAtraso!Numeromayor + 1)
                rstAtraso.Close
            End If
    
            ZFecha = Fecha.Text
            ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZFechaEntrega = Fecha.Text
            ZFechaEntregaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZTerminado = "  -     -   "
            ZArticulo = "  -   -   "
            ZDesTerminado = ""
            ZDesArticulo = ""
            ZConcepto = Str$(ConceptoAtraso.ListIndex + 4)
            ZSolicitud = ""
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Atraso ("
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Problema ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "FechaEntrega ,"
            ZSql = ZSql + "OrdFechaEntrega ,"
            ZSql = ZSql + "DesCliente ,"
            ZSql = ZSql + "DesTerminado ,"
            ZSql = ZSql + "DesArticulo ,"
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Solicitud ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "VersionPedido)"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZAtraso + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZFechaOrd + "',"
            ZSql = ZSql + "'" + Pedido.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + DescriMotivo.Text + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZFechaEntrega + "',"
            ZSql = ZSql + "'" + ZFechaEntregaOrd + "',"
            ZSql = ZSql + "'" + Left$(DesCliente.Caption, 50) + "',"
            ZSql = ZSql + "'" + ZDesTerminado + "',"
            ZSql = ZSql + "'" + ZDesArticulo + "',"
            ZSql = ZSql + "'" + ZConcepto + "',"
            ZSql = ZSql + "'" + ZSolicitud + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + "" + "')"
           
            spAtraso = ZSql
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        
            PantaMotivo.Visible = False
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Click()

    Erase XLote

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "10" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    Canti4.Text = ""
    Canti5.Text = ""
    
    Rem Numero.SetFocus
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 6
                If DBGrid1.Row < 40 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    WRow = DBGrid1.Row
                    DBGrid1.Col = 4
                    KeyCode = 0
                End If
            
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
    Iva(3) = "Consumidor Final"
    Iva(4) = "Exento"
    Iva(5) = "Monotributo"
    Iva(6) = "No Catalogado"
    
    ConceptoAtraso.Clear
    
    ConceptoAtraso.AddItem ""
    ConceptoAtraso.AddItem "Error del Sistema"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Problemas Vehiculos"
    ConceptoAtraso.AddItem "Problemas Logistica"
    ConceptoAtraso.AddItem "Problemas Recepcion Cliente"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Corte de Luz"
    ConceptoAtraso.AddItem "Pedido por el Cliente"
    ConceptoAtraso.AddItem "Falta de Pago"
    ConceptoAtraso.AddItem "Confirmacion Pedido Parcial"
    ConceptoAtraso.AddItem "Envase"
    
    ConceptoAtraso.ListIndex = 0
    Graba.Enabled = True
    Borra.Enabled = True
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 6, 0 To 40)

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
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cant. Entregar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Cant. Descontar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    DBGrid1.Font.Bold = True

    Erase XLote

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "10" + "'"
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
    Numero.SetFocus
     
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    XEntra = "S"
    
    Erase Auxiliar
    Erase ClavePedido
    Erase XLote
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad - !Facturado
                    
                    If Canti <> 0 Then
                    
                        XCanti1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                        XCanti2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                        
                        If XCanti1 <> 0 Or XCanti2 <> 0 Then
                
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
                            DBGrid1.Text = ""
                
                            XCantidad1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###.##", Str$(XCantidad1))
                    
                            xCantidad2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                            DBGrid1.Col = 5
                            DBGrid1.Text = Pusing("###,###.##", Str$(xCantidad2))
                    
                            DBGrid1.Col = 6
                            DBGrid1.Text = "S"
                    
                            
                            Rem If XEntra = "S" Then
                            Rem     Envase1.Text = IIf(IsNull(!Env1), "", !Env1)
                            Rem     Envase2.Text = IIf(IsNull(!Env2), "", !Env2)
                            Rem     Envase3.Text = IIf(IsNull(!Env3), "", !Env3)
                            Rem     Envase4.Text = IIf(IsNull(!Env4), "", !Env4)
                            Rem     Envase5.Text = IIf(IsNull(!Env5), "", !Env5)
                            Rem     Canti1.Text = IIf(IsNull(!CantiEnv1), "", !CantiEnv1)
                            Rem     Canti2.Text = IIf(IsNull(!CantiEnv2), "", !CantiEnv2)
                            Rem     Canti3.Text = IIf(IsNull(!CantiEnv3), "", !CantiEnv3)
                            Rem     Canti4.Text = IIf(IsNull(!CantiEnv4), "", !CantiEnv4)
                            Rem     Canti5.Text = IIf(IsNull(!CantiEnv5), "", !CantiEnv5)
                            Rem     XEntra = ""
                            Rem End If
                            
                            XLote(Renglon, 1) = IIf(IsNull(!lote1), "", !lote1)
                            XLote(Renglon, 2) = IIf(IsNull(!CantiLote1), "", Str$(!CantiLote1))
                            XLote(Renglon, 3) = IIf(IsNull(!lote2), "", !lote2)
                            XLote(Renglon, 4) = IIf(IsNull(!CantiLote2), "", Str$(!CantiLote2))
                            XLote(Renglon, 5) = IIf(IsNull(!lote3), "", !lote3)
                            XLote(Renglon, 6) = IIf(IsNull(!CantiLote3), "", Str$(!CantiLote3))
                            XLote(Renglon, 7) = IIf(IsNull(!lote4), "", !lote4)
                            XLote(Renglon, 8) = IIf(IsNull(!CantiLote4), "", Str$(!CantiLote4))
                            XLote(Renglon, 9) = IIf(IsNull(!lote5), "", !lote5)
                            XLote(Renglon, 10) = IIf(IsNull(!CantiLote5), "", Str$(!CantiLote5))
                            XLote(Renglon, 11) = IIf(IsNull(!lote6), "", !lote6)
                            XLote(Renglon, 12) = IIf(IsNull(!CantiLote6), "", !CantiLote6)
                            XLote(Renglon, 13) = IIf(IsNull(!lote7), "", !lote7)
                            XLote(Renglon, 14) = IIf(IsNull(!CantiLote7), "", !CantiLote7)
                            XLote(Renglon, 15) = IIf(IsNull(!lote8), "", !lote8)
                            XLote(Renglon, 16) = IIf(IsNull(!CantiLote8), "", !CantiLote8)
                            XLote(Renglon, 17) = IIf(IsNull(!lote9), "", !lote9)
                            XLote(Renglon, 18) = IIf(IsNull(!CantiLote9), "", !CantiLote9)
                            XLote(Renglon, 19) = IIf(IsNull(!lote10), "", !lote10)
                            XLote(Renglon, 20) = IIf(IsNull(!CantiLote10), "", !CantiLote10)
                            XLote(Renglon, 21) = IIf(IsNull(!lote11), "", !lote11)
                            XLote(Renglon, 22) = IIf(IsNull(!CantiLote11), "", !CantiLote11)
                            XLote(Renglon, 23) = IIf(IsNull(!lote12), "", !lote12)
                            XLote(Renglon, 24) = IIf(IsNull(!CantiLote12), "", !CantiLote12)
                            
                            XLote(Renglon, 31) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                            XLote(Renglon, 32) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                            XLote(Renglon, 33) = IIf(IsNull(rstPedido!Env2), "0", rstPedido!Env2)
                            XLote(Renglon, 34) = IIf(IsNull(rstPedido!CantiEnv2), "0", rstPedido!CantiEnv2)
                            XLote(Renglon, 35) = IIf(IsNull(rstPedido!Env3), "0", rstPedido!Env3)
                            XLote(Renglon, 36) = IIf(IsNull(rstPedido!CantiEnv3), "0", rstPedido!CantiEnv3)
                            XLote(Renglon, 37) = IIf(IsNull(rstPedido!Env4), "0", rstPedido!Env4)
                            XLote(Renglon, 38) = IIf(IsNull(rstPedido!CantiEnv4), "0", rstPedido!CantiEnv4)
                            XLote(Renglon, 39) = IIf(IsNull(rstPedido!Env5), "0", rstPedido!Env5)
                            XLote(Renglon, 40) = IIf(IsNull(rstPedido!CantiEnv5), "0", rstPedido!CantiEnv5)
                            XLote(Renglon, 41) = IIf(IsNull(rstPedido!Env6), "0", rstPedido!Env6)
                            XLote(Renglon, 42) = IIf(IsNull(rstPedido!CantiEnv6), "0", rstPedido!CantiEnv6)
                            XLote(Renglon, 43) = IIf(IsNull(rstPedido!Env7), "0", rstPedido!Env7)
                            XLote(Renglon, 44) = IIf(IsNull(rstPedido!CantiEnv7), "0", rstPedido!CantiEnv7)
                            XLote(Renglon, 45) = IIf(IsNull(rstPedido!Env8), "0", rstPedido!Env8)
                            XLote(Renglon, 46) = IIf(IsNull(rstPedido!CantiEnv8), "0", rstPedido!CantiEnv8)
                            XLote(Renglon, 47) = IIf(IsNull(rstPedido!Env9), "0", rstPedido!Env9)
                            XLote(Renglon, 48) = IIf(IsNull(rstPedido!CantiEnv9), "0", rstPedido!CantiEnv9)
                            XLote(Renglon, 49) = IIf(IsNull(rstPedido!Env10), "0", rstPedido!Env10)
                            XLote(Renglon, 50) = IIf(IsNull(rstPedido!CantiEnv10), "0", rstPedido!CantiEnv10)
                            XLote(Renglon, 51) = IIf(IsNull(rstPedido!Env11), "0", rstPedido!Env11)
                            XLote(Renglon, 52) = IIf(IsNull(rstPedido!CantiEnv11), "0", rstPedido!CantiEnv11)
                            XLote(Renglon, 53) = IIf(IsNull(rstPedido!Env12), "0", rstPedido!Env12)
                            XLote(Renglon, 54) = IIf(IsNull(rstPedido!CantiEnv12), "0", rstPedido!CantiEnv12)
                    
                            Auxiliar(Renglon, 1) = Auxi1
                            Auxiliar(Renglon, 2) = Canti
                            
                            ClavePedido(Renglon) = !Clave
                            
                        End If
                        
                    End If
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    
    
    
    
    
    
    Erase ZZEnvase
    Erase ZZCanti
    
    Canti1.Text = ""
    Envase1.Text = ""
    Canti2.Text = ""
    Envase2.Text = ""
    Canti3.Text = ""
    Envase3.Text = ""
    Canti4.Text = ""
    Envase4.Text = ""
    Canti5.Text = ""
    Envase5.Text = ""
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    ZZLugar = 0
    
    For CicloEnvase = 1 To 100
    
        For ZZZCicloEnvase = 32 To 54 Step 2
        
            If Val(XLote(CicloEnvase, ZZZCicloEnvase)) <> 0 Then
            
                Entra = "S"
                For CicloEnvaseII = 1 To 5
                    If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, ZZZCicloEnvase - 1) Then
                        ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, ZZZCicloEnvase)))
                        Entra = "N"
                        Exit For
                    End If
                Next CicloEnvaseII
                
                If Entra = "S" Then
                    ZZLugar = ZZLugar + 1
                    ZZCanti(ZZLugar) = XLote(CicloEnvase, ZZZCicloEnvase)
                    ZZEnvase(ZZLugar) = XLote(CicloEnvase, ZZZCicloEnvase - 1)
                End If
                
            End If
            
        Next ZZZCicloEnvase
        
    Next CicloEnvase
    
    Envase1.Text = ZZEnvase(1)
    Envase2.Text = ZZEnvase(2)
    Envase3.Text = ZZEnvase(3)
    Envase4.Text = ZZEnvase(4)
    Envase5.Text = ZZEnvase(5)
    
    Canti1.Text = ZZCanti(1)
    Canti2.Text = ZZCanti(2)
    Canti3.Text = ZZCanti(3)
    Canti4.Text = ZZCanti(4)
    Canti5.Text = ZZCanti(5)
    
    spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri1.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri1.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri2.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri3.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri4.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri5.Caption = ""
    End If
    
    
    
    
    
    
    WRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
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
            
            Case "T"
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
            
                    rstPrecios.Close
                End If

        End Select
        
    Next DA
    
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
    Borra.Enabled = True

End Sub

Private Sub Proceso1_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    Erase XLote
    
    Rem XParam = "'" + "01" + "','" _
    rem             + Numero.Text + "'"
    
    ZZNumero = 900000 + Val(Numero.Text)
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where EStadistica.Tipo = " + "'" + "01" + "'"
    ZSql = ZSql + " and Estadistica.Numero = " + "'" + Str$(ZZNumero) + "'"
    spEstadistica = ZSql
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
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = ""
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Str$(!Cantidad)
                    Auxiliar(Renglon, 3) = Str$(!Precio)
                    Auxiliar(Renglon, 5) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    Auxiliar(Renglon, 6) = IIf(IsNull(rstEstadistica!Canti1), "", rstEstadistica!Canti1)
                    
                    Auxiliar(Renglon, 7) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    Auxiliar(Renglon, 8) = IIf(IsNull(rstEstadistica!Canti2), "", rstEstadistica!Canti2)
                    Auxiliar(Renglon, 9) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    Auxiliar(Renglon, 10) = IIf(IsNull(rstEstadistica!Canti3), "", rstEstadistica!Canti3)
                    Auxiliar(Renglon, 11) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    Auxiliar(Renglon, 12) = IIf(IsNull(rstEstadistica!Canti4), "", rstEstadistica!Canti4)
                    Auxiliar(Renglon, 13) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    Auxiliar(Renglon, 14) = IIf(IsNull(rstEstadistica!Canti5), "", rstEstadistica!Canti5)
                    
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                    
                    If Len(Trim(WLoteAdicional)) = 98 Then
                        Auxiliar(Renglon, 15) = Mid$(WLoteAdicional, 1, 8)
                        Auxiliar(Renglon, 16) = Mid$(WLoteAdicional, 9, 6)
                        Auxiliar(Renglon, 17) = Mid$(WLoteAdicional, 15, 8)
                        Auxiliar(Renglon, 18) = Mid$(WLoteAdicional, 23, 6)
                        Auxiliar(Renglon, 19) = Mid$(WLoteAdicional, 29, 8)
                        Auxiliar(Renglon, 20) = Mid$(WLoteAdicional, 37, 6)
                        Auxiliar(Renglon, 21) = Mid$(WLoteAdicional, 43, 8)
                        Auxiliar(Renglon, 22) = Mid$(WLoteAdicional, 51, 6)
                        Auxiliar(Renglon, 23) = Mid$(WLoteAdicional, 57, 8)
                        Auxiliar(Renglon, 24) = Mid$(WLoteAdicional, 65, 6)
                        Auxiliar(Renglon, 25) = Mid$(WLoteAdicional, 71, 8)
                        Auxiliar(Renglon, 26) = Mid$(WLoteAdicional, 79, 6)
                        Auxiliar(Renglon, 27) = Mid$(WLoteAdicional, 85, 8)
                        Auxiliar(Renglon, 28) = Mid$(WLoteAdicional, 93, 6)
                    End If
                    
                    XLote(Renglon, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    XLote(Renglon, 3) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    XLote(Renglon, 5) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    XLote(Renglon, 7) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    XLote(Renglon, 9) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    XLote(Renglon, 11) = ""
                    XLote(Renglon, 13) = ""
                    XLote(Renglon, 15) = ""
                    XLote(Renglon, 17) = ""
                    XLote(Renglon, 19) = ""
                    XLote(Renglon, 21) = ""
                    XLote(Renglon, 23) = ""
                    
                    XLote(Renglon, 2) = IIf(IsNull(rstEstadistica!Canti1), "", rstEstadistica!Canti1)
                    XLote(Renglon, 4) = IIf(IsNull(rstEstadistica!Canti2), "", rstEstadistica!Canti2)
                    XLote(Renglon, 6) = IIf(IsNull(rstEstadistica!Canti3), "", rstEstadistica!Canti3)
                    XLote(Renglon, 8) = IIf(IsNull(rstEstadistica!Canti4), "", rstEstadistica!Canti4)
                    XLote(Renglon, 10) = IIf(IsNull(rstEstadistica!Canti5), "", rstEstadistica!Canti5)
                    XLote(Renglon, 12) = ""
                    XLote(Renglon, 14) = ""
                    XLote(Renglon, 16) = ""
                    XLote(Renglon, 18) = ""
                    XLote(Renglon, 20) = ""
                    XLote(Renglon, 22) = ""
                    XLote(Renglon, 24) = ""
                    
                    XLote(Renglon, 31) = IIf(IsNull(rstEstadistica!Env1), "", rstEstadistica!Env1)
                    XLote(Renglon, 33) = IIf(IsNull(rstEstadistica!Env2), "", rstEstadistica!Env2)
                    XLote(Renglon, 35) = IIf(IsNull(rstEstadistica!Env3), "", rstEstadistica!Env3)
                    XLote(Renglon, 37) = IIf(IsNull(rstEstadistica!Env4), "", rstEstadistica!Env4)
                    XLote(Renglon, 39) = IIf(IsNull(rstEstadistica!Env5), "", rstEstadistica!Env5)
                    XLote(Renglon, 41) = ""
                    XLote(Renglon, 43) = ""
                    XLote(Renglon, 45) = ""
                    XLote(Renglon, 47) = ""
                    XLote(Renglon, 49) = ""
                    XLote(Renglon, 51) = ""
                    XLote(Renglon, 53) = ""
                    
                    XLote(Renglon, 32) = IIf(IsNull(rstEstadistica!CantiEnv1), "", rstEstadistica!CantiEnv1)
                    XLote(Renglon, 34) = IIf(IsNull(rstEstadistica!CantiEnv2), "", rstEstadistica!CantiEnv2)
                    XLote(Renglon, 36) = IIf(IsNull(rstEstadistica!CantiEnv3), "", rstEstadistica!CantiEnv3)
                    XLote(Renglon, 38) = IIf(IsNull(rstEstadistica!CantiEnv4), "", rstEstadistica!CantiEnv4)
                    XLote(Renglon, 40) = IIf(IsNull(rstEstadistica!CantiEnv5), "", rstEstadistica!CantiEnv5)
                    XLote(Renglon, 42) = ""
                    XLote(Renglon, 44) = ""
                    XLote(Renglon, 46) = ""
                    XLote(Renglon, 48) = ""
                    XLote(Renglon, 50) = ""
                    XLote(Renglon, 52) = ""
                    XLote(Renglon, 54) = ""
                    
                    WLoteAdicional = IIf(IsNull(!LoteAdicional), "", !LoteAdicional)
                    WEnvAdicional = IIf(IsNull(!EnvAdicional), "", !EnvAdicional)
                    
                    If Len(Trim(WLoteAdicional)) = 98 Then
                        XLote(Renglon, 11) = Mid$(WLoteAdicional, 1, 8)
                        XLote(Renglon, 12) = Mid$(WLoteAdicional, 9, 6)
                        XLote(Renglon, 13) = Mid$(WLoteAdicional, 15, 8)
                        XLote(Renglon, 14) = Mid$(WLoteAdicional, 23, 6)
                        XLote(Renglon, 15) = Mid$(WLoteAdicional, 29, 8)
                        XLote(Renglon, 16) = Mid$(WLoteAdicional, 37, 6)
                        XLote(Renglon, 17) = Mid$(WLoteAdicional, 43, 8)
                        XLote(Renglon, 18) = Mid$(WLoteAdicional, 51, 6)
                        XLote(Renglon, 19) = Mid$(WLoteAdicional, 57, 8)
                        XLote(Renglon, 20) = Mid$(WLoteAdicional, 65, 6)
                        XLote(Renglon, 21) = Mid$(WLoteAdicional, 71, 8)
                        XLote(Renglon, 22) = Mid$(WLoteAdicional, 79, 6)
                        XLote(Renglon, 23) = Mid$(WLoteAdicional, 85, 8)
                        XLote(Renglon, 24) = Mid$(WLoteAdicional, 93, 6)
                    End If
                    
                    If Len(Trim(WEnvAdicional)) = 56 Then
                        XLote(Renglon, 41) = Mid$(WEnvAdicional, 1, 4)
                        XLote(Renglon, 42) = Mid$(WEnvAdicional, 5, 4)
                        XLote(Renglon, 43) = Mid$(WEnvAdicional, 9, 4)
                        XLote(Renglon, 44) = Mid$(WEnvAdicional, 13, 4)
                        XLote(Renglon, 45) = Mid$(WEnvAdicional, 17, 4)
                        XLote(Renglon, 46) = Mid$(WEnvAdicional, 21, 4)
                        XLote(Renglon, 47) = Mid$(WEnvAdicional, 25, 4)
                        XLote(Renglon, 48) = Mid$(WEnvAdicional, 29, 4)
                        XLote(Renglon, 49) = Mid$(WEnvAdicional, 33, 4)
                        XLote(Renglon, 50) = Mid$(WEnvAdicional, 37, 4)
                        XLote(Renglon, 51) = Mid$(WEnvAdicional, 41, 4)
                        XLote(Renglon, 52) = Mid$(WEnvAdicional, 45, 4)
                        XLote(Renglon, 53) = Mid$(WEnvAdicional, 49, 4)
                        XLote(Renglon, 54) = Mid$(WEnvAdicional, 53, 4)
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    
    
    
    
    
    
    Erase ZZEnvase
    Erase ZZCanti
    
    Canti1.Text = ""
    Envase1.Text = ""
    Canti2.Text = ""
    Envase2.Text = ""
    Canti3.Text = ""
    Envase3.Text = ""
    Canti4.Text = ""
    Envase4.Text = ""
    Canti5.Text = ""
    Envase5.Text = ""
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    ZZLugar = 0
    
    For CicloEnvase = 1 To 100
    
        For ZZZCiclo = 32 To 54 Step 2
    
            If Val(XLote(CicloEnvase, ZZZCiclo)) <> 0 Then
                Entra = "S"
                For CicloEnvaseII = 1 To 5
                    If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, ZZZCiclo - 1) Then
                        ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, ZZZCiclo)))
                        Entra = "N"
                        Exit For
                    End If
                Next CicloEnvaseII
                
                If Entra = "S" Then
                    ZZLugar = ZZLugar + 1
                    ZZCanti(ZZLugar) = XLote(CicloEnvase, ZZZCiclo)
                    ZZEnvase(ZZLugar) = XLote(CicloEnvase, ZZZCiclo - 1)
                End If
            End If
            
        Next ZZZCiclo
        
    Next CicloEnvase
    
    Envase1.Text = ZZEnvase(1)
    Envase2.Text = ZZEnvase(2)
    Envase3.Text = ZZEnvase(3)
    Envase4.Text = ZZEnvase(4)
    Envase5.Text = ZZEnvase(5)
    
    Canti1.Text = ZZCanti(1)
    Canti2.Text = ZZCanti(2)
    Canti3.Text = ZZCanti(3)
    Canti4.Text = ZZCanti(4)
    Canti5.Text = ZZCanti(5)
    
    spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri1.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri1.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri2.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri3.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri4.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri5.Caption = ""
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
                WClaveMp = Cliente.Text + WArti
                
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
                    
                    Auxiliar(DA, 30) = rstArticulo!Descripcion
                    
                    rstArticulo.Close
                
                    spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPreciosMp.RecordCount > 0 Then
                        DBGrid1.Col = 3
                        DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                        rstPreciosMp.Close
                    End If
                
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
                    
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    
                    Auxiliar(DA, 30) = rstPrecios!Descripcion
                    
                    rstPrecios.Close
                End If
                
        End Select
        
    Next DA
    
    Renglon = Renglon + 1
            
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
                
    DBGrid1.Col = 0
    DBGrid1.Text = ""
                            
    DBGrid1.Col = 1
    DBGrid1.Text = ""
                
    DBGrid1.Col = 2
    DBGrid1.Text = ""
                
    DBGrid1.Col = 3
    DBGrid1.Text = ""
                
    DBGrid1.Col = 4
    DBGrid1.Text = ""
                            
    DBGrid1.Col = 5
    DBGrid1.Text = ""

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
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0
    
    DBGrid1.SetFocus
    
    Graba.Enabled = False
    Borra.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "10" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Pedido.Text = Trim(rstCtacte!Pedido)
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Remito.Text = Trim(rstCtacte!Remito)
            
            rstCtacte.Close
            
            spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                rstPedido.Close
            End If
                
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
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                Rem WDirentrega = rstCliente!DirEntrega
                WDirentrega = ""
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                WDirentrega = ZDirEntrega(ZLugarDirEntrega)
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
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                Pedido.SetFocus
                    Else
                    
                If rstPedido!Tipoped = 5 Then
                    rstPedido.Close
                    m$ = "EL PEDIDO ES DE MUESTRAS"
                    a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                    Pedido.SetFocus
                    Exit Sub
                End If
                    
                Cliente.Text = rstPedido!Cliente
                ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                
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
                
                If Val(Wempresa) = 1 And Cliente.Text = "P00005" Then
                    If Left$(rstPedido!Terminado, 4) = "PT-5" Or rstPedido!Terminado = "PT-03000-001" Then
                        WTipoPedido = "PG"
                    End If
                End If
                
                If Left$(rstPedido!Terminado, 4) = "PT-4" Then
                    WTipoPedido = "TA"
                End If
                
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
                    WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                    WRazon = rstCliente!Razon
                    WDireccion = rstCliente!Direccion
                    WLocalidad = rstCliente!Localidad
                    WProv = rstCliente!Provincia
                    WPostal = rstCliente!Postal
                    WCuit = rstCliente!Cuit
                    Rem WDirentrega = rstCliente!DirEntrega
                    WDirentrega = ""
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                Call Proceso_Click
                Remito.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        m$ = "Coloque Remito de Pta I"
                        a% = MsgBox(m$, 0, "Emision de facturas")
                    Case "FA"
                        m$ = "Coloque Remito de Pta VII"
                        a% = MsgBox(m$, 0, "Emision de facturas")
                    Case "TA"
                        m$ = "Coloque Remito de Pta II"
                        a% = MsgBox(m$, 0, "Emision de facturas")
                    Case Else
                        m$ = "Coloque Remito de Pta V"
                        a% = MsgBox(m$, 0, "Emision de facturas")
                End Select
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub reImpre_Click()

    Rem Call Impresion_Remito_Calculo
    
    Rem If ZZPasaImpre > 16 Then
    Rem     If ZZPasaImpre > 32 Then
    Rem         m$ = "Atencion : Se utilizaran 3 remitos para la impresion de la totalidad de los productos"
    Rem         a% = MsgBox(m$, 0, "Emision de facturas")
    Rem             Else
    Rem         m$ = "Atencion : Se utilizaran 2 remitos para la impresion de la totalidad de los productos"
    Rem         a% = MsgBox(m$, 0, "Emision de facturas")
    Rem     End If
    Rem End If
        
    T$ = "Impresion"
    m$ = "Desea imprimir el remito"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion_RemitoPrueba
    End If
    
    If Val(Wempresa) = 1 Then
        T$ = "Impresion"
        m$ = "Desea imprimir hoja de seguridad, certificados de analisis y guia de emergencia"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion_Varios
       End If
    End If
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
    
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion_Remito()

    Rem m$ = "Coloque el remito en bandeja 1"
    Rem a% = MsgBox(m$, 0, "Impresion de Remitos")

    If Val(Wempresa) = 1 Then
        Open "lpt1" For Output As #1
        Rem Open "DADA1.TXT" For Output As #1
            Else
        Open "lpt1" For Output As #1
        Rem Open "DADA1.TXT" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    End If
    
    Print #1, Chr$(27) + Chr$(40) + "19U"
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72)
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    Print #1, Chr$(27) & "&l4H"
    
    For FF = 1 To 2
        
        Call Impresion_Remito_Cabecera
        
        Impre = 0
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                
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
                
                    ZClase = ""
                    ZIntervencion = ""
                    ZNaciones = ""
                    spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                        ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                        ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                        ZDescriOnu = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
                        ZEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                        ZClase = Trim(ZClase)
                        ZIntervencion = Trim(ZIntervencion)
                        ZNaciones = Trim(ZNaciones)
                        rstTerminado.Close
                    End If
                    
                    If Trim(ZClase) <> "" Then
                    
                        ZImpre = ""
                        ZImpre = "Clase:" + ZClase + " N.ONU:" + ZNaciones + " GRUPO DE EMBALAJE:" + ZEmbalaje
            
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
                        Print #1, Tab(10); Left$(Descri, 40);
                        Print #1, Tab(85); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg."
                        
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72);
                        Print #1, Tab(15); ZDescriOnu
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72);
                        Print #1, Tab(15); ZImpre
                        
                        Impre = Impre + 3
                            
                            Else
                                        
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
                        Print #1, Tab(10); Left$(Descri, 40);
                        Print #1, Tab(85); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg."
                        Impre = Impre + 1
                        
                    End If
                        
                End If
                
                If FF = 2 Then
                
                    ZLote1 = XLote(Suma, 1)
                    ZCantidad1 = XLote(Suma, 2)
                    ZLote2 = XLote(Suma, 3)
                    ZCantidad2 = XLote(Suma, 4)
                    ZLote3 = XLote(Suma, 5)
                    ZCantidad3 = XLote(Suma, 6)
                    ZLote4 = XLote(Suma, 7)
                    ZCantidad4 = XLote(Suma, 8)
                    ZLote5 = XLote(Suma, 9)
                    ZCantidad5 = XLote(Suma, 10)
                    ZLote6 = XLote(Suma, 11)
                    ZCantidad6 = XLote(Suma, 12)
                    ZLote7 = XLote(Suma, 13)
                    ZCantidad7 = XLote(Suma, 14)
                    ZLote8 = XLote(Suma, 15)
                    ZCantidad8 = XLote(Suma, 16)
                    ZLote9 = XLote(Suma, 17)
                    ZCantidad9 = XLote(Suma, 18)
                    ZLote10 = XLote(Suma, 19)
                    ZCantidad10 = XLote(Suma, 20)
                    ZLote11 = XLote(Suma, 21)
                    ZCantidad11 = XLote(Suma, 22)
                    ZLote12 = XLote(Suma, 23)
                    ZCantidad12 = XLote(Suma, 24)
                    
                    If Trim(Producto) <> "" Then
                    
                        If Left$(Producto, 2) = "DY" Then
                        
                            For ZZZCiclo = 1 To 12
                            
                                Select Case ZZZCiclo
                                    Case 1
                                        ZZZLote = ZLote1
                                    Case 2
                                        ZZZLote = ZLote2
                                    Case 3
                                        ZZZLote = ZLote3
                                    Case 4
                                        ZZZLote = ZLote4
                                    Case 5
                                        ZZZLote = ZLote5
                                    Case 6
                                        ZZZLote = ZLote6
                                    Case 7
                                        ZZZLote = ZLote7
                                    Case 8
                                        ZZZLote = ZLote8
                                    Case 9
                                        ZZZLote = ZLote9
                                    Case 10
                                        ZZZLote = ZLote10
                                    Case 11
                                        ZZZLote = ZLote11
                                    Case Else
                                        ZZZLote = ZLote12
                                End Select
                        
                                ZZZArti = Left$(Producto, 3) + Right$(Producto, 7)
                                XParam = "'" + ZZZLote + "','" _
                                             + ZZZArti + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    ZZZPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                                    rstLaudo.Close
                                    
                                    If Trim(ZZZPartiOri) <> "" Then
                                    
                                        Select Case ZZZCiclo
                                            Case 1
                                                ZLote1 = ZZZPartiOri
                                            Case 2
                                                ZLote2 = ZZZPartiOri
                                            Case 3
                                                ZLote3 = ZZZPartiOri
                                            Case 4
                                                ZLote4 = ZZZPartiOri
                                            Case 5
                                                ZLote5 = ZZZPartiOri
                                            Case 6
                                                ZLote6 = ZZZPartiOri
                                            Case 7
                                                ZLote7 = ZZZPartiOri
                                            Case 8
                                                ZLote8 = ZZZPartiOri
                                            Case 9
                                                ZLote9 = ZZZPartiOri
                                            Case 10
                                                ZLote10 = ZZZPartiOri
                                            Case 11
                                                ZLote11 = ZZZPartiOri
                                            Case Else
                                                ZLote12 = ZZZPartiOri
                                        End Select
                                        
                                    End If
                                End If
                                
                            Next ZZZCiclo
                            
                                Else
                                
                            For ZZZCiclo = 1 To 12
                            
                                Select Case ZZZCiclo
                                    Case 1
                                        ZZZLote = ZLote1
                                    Case 2
                                        ZZZLote = ZLote2
                                    Case 3
                                        ZZZLote = ZLote3
                                    Case 4
                                        ZZZLote = ZLote4
                                    Case 5
                                        ZZZLote = ZLote5
                                    Case 6
                                        ZZZLote = ZLote6
                                    Case 7
                                        ZZZLote = ZLote7
                                    Case 8
                                        ZZZLote = ZLote8
                                    Case 9
                                        ZZZLote = ZLote9
                                    Case 10
                                        ZZZLote = ZLote10
                                    Case 11
                                        ZZZLote = ZLote11
                                    Case Else
                                        ZZZLote = ZLote12
                                End Select
                                
                                XParam = "'" + ZZZLote + "','" _
                                             + Producto + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                
                                    ZZZPartiOri = IIf(IsNull(rstHoja!LoteColorante), "", rstHoja!LoteColorante)
                                    
                                    rstHoja.Close
                                    
                                    If Trim(ZZZPartiOri) <> "" Then
                                    
                                        Select Case ZZZCiclo
                                            Case 1
                                                ZLote1 = ZZZPartiOri
                                            Case 2
                                                ZLote2 = ZZZPartiOri
                                            Case 3
                                                ZLote3 = ZZZPartiOri
                                            Case 4
                                                ZLote4 = ZZZPartiOri
                                            Case 5
                                                ZLote5 = ZZZPartiOri
                                            Case 6
                                                ZLote6 = ZZZPartiOri
                                            Case 7
                                                ZLote7 = ZZZPartiOri
                                            Case 8
                                                ZLote8 = ZZZPartiOri
                                            Case 9
                                                ZLote9 = ZZZPartiOri
                                            Case 10
                                                ZLote10 = ZZZPartiOri
                                            Case 11
                                                ZLote11 = ZZZPartiOri
                                            Case Else
                                                ZLote12 = ZZZPartiOri
                                        End Select
                                        
                                    End If
                                End If
                                
                            Next ZZZCiclo
                            
                        End If
                    
                    End If
                    
                    
                    ZEnv1 = XLote(Suma, 31)
                    ZCantiEnv1 = XLote(Suma, 32)
                    ZEnv2 = XLote(Suma, 33)
                    ZCantiEnv2 = XLote(Suma, 34)
                    ZEnv3 = XLote(Suma, 35)
                    ZCantiEnv3 = XLote(Suma, 36)
                    ZEnv4 = XLote(Suma, 37)
                    ZCantiEnv4 = XLote(Suma, 38)
                    ZEnv5 = XLote(Suma, 39)
                    ZCantiEnv5 = XLote(Suma, 40)
                    ZEnv6 = XLote(Suma, 41)
                    ZCantiEnv6 = XLote(Suma, 42)
                    ZEnv7 = XLote(Suma, 43)
                    ZCantiEnv7 = XLote(Suma, 44)
                    ZEnv8 = XLote(Suma, 45)
                    ZCantiEnv8 = XLote(Suma, 46)
                    ZEnv9 = XLote(Suma, 47)
                    ZCantiEnv9 = XLote(Suma, 48)
                    ZEnv10 = XLote(Suma, 49)
                    ZCantiEnv10 = XLote(Suma, 50)
                    ZEnv11 = XLote(Suma, 51)
                    ZCantiEnv11 = XLote(Suma, 52)
                    ZEnv12 = XLote(Suma, 53)
                    ZCantiEnv12 = XLote(Suma, 54)
                    
                    ZDescri1 = ""
                    ZDescri2 = ""
                    ZDescri3 = ""
                    ZDescri4 = ""
                    ZDescri5 = ""
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv1 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri1 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv2 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri2 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv3 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri3 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv4 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri4 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv5 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri5 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv6 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri6 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv7 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri7 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv8 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri8 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv9 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri9 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv10 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri10 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv11 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri11 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv12 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri12 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    If Val(ZCantidad1) <> 0 Then
                    
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72);
                        
                        Print #1, Tab(10); Alinea("#####", ZCantidad1); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote1, 8);
                        If Val(ZCantiEnv1) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv1);
                            Print #1, " x "; ZDescri1;
                        End If
                    
                        If Val(ZCantidad2) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad2); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote2, 8);
                            If Val(ZCantiEnv2) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv2);
                                Print #1, " x "; ZDescri2;
                            End If
                        End If
                    
                        If Val(ZCantidad3) <> 0 Then
                            Print #1, Tab(86); "|";
                            Print #1, Tab(88); Alinea("#####", ZCantidad3); " Kg";
                            Print #1, Tab(97); "Lote:"; Left$(ZLote3, 8);
                            If Val(ZCantiEnv3) <> 0 Then
                                Print #1, Tab(109); Alinea("##", ZCantiEnv3);
                                Print #1, " x "; ZDescri3;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                    If Val(ZCantidad4) <> 0 Then
                    
                        Print #1, Tab(10); Alinea("#####", ZCantidad4); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote4, 8);
                        If Val(ZCantiEnv4) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv4);
                            Print #1, " x "; ZDescri4;
                        End If
                    
                        If Val(ZCantidad5) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad5); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote5, 8);
                            If Val(ZCantiEnv5) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv5);
                                Print #1, " x "; ZDescri5;
                            End If
                        End If
                    
                        If Val(ZCantidad6) <> 0 Then
                            Print #1, Tab(86); "|";
                            Print #1, Tab(88); Alinea("#####", ZCantidad6); " Kg";
                            Print #1, Tab(97); "Lote:"; Left$(ZLote6, 8);
                            If Val(ZCantiEnv6) <> 0 Then
                                Print #1, Tab(109); Alinea("##", ZCantiEnv6);
                                Print #1, " x "; ZDescri6;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                    If Val(ZCantidad7) <> 0 Then
                    
                        Print #1, Tab(10); Alinea("#####", ZCantidad7); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote7, 8);
                        If Val(ZCantiEnv7) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv7);
                            Print #1, " x "; ZDescri7;
                        End If
                    
                        If Val(ZCantidad8) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad8); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote8, 8);
                            If Val(ZCantiEnv8) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv8);
                                Print #1, " x "; ZDescri8;
                            End If
                        End If
                    
                        If Val(ZCantidad9) <> 0 Then
                            Print #1, Tab(86); "|";
                            Print #1, Tab(88); Alinea("#####", ZCantidad9); " Kg";
                            Print #1, Tab(97); "Lote:"; Left$(ZLote9, 8);
                            If Val(ZCantiEnv9) <> 0 Then
                                Print #1, Tab(109); Alinea("##", ZCantiEnv9);
                                Print #1, " x "; ZDescri9;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                    If Val(ZCantidad10) <> 0 Then
                    
                        Print #1, Tab(10); Alinea("#####", ZCantidad10); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote10, 8);
                        If Val(ZCantiEnv10) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv10);
                            Print #1, " x "; ZDescri10;
                        End If
                    
                        If Val(ZCantidad11) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad11); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote11, 8);
                            If Val(ZCantiEnv11) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv11);
                                Print #1, " x "; ZDescri11;
                            End If
                        End If
                    
                        If Val(ZCantidad12) <> 0 Then
                            Print #1, Tab(86); "|";
                            Print #1, Tab(88); Alinea("#####", ZCantidad12); " Kg";
                            Print #1, Tab(97); "Lote:"; Left$(ZLote12, 8);
                            If Val(ZCantiEnv12) <> 0 Then
                                Print #1, Tab(109); Alinea("##", ZCantiEnv12);
                                Print #1, " x "; ZDescri12;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                End If
                
                If Impre > 16 Then
                    Print #1, Chr$(12)
                    Call Impresion_Remito_Cabecera
                    Impre = 0
                End If
                
            Next iRow
            
        Next a
        
        If FF = 1 Then
        
            If Val(Wempresa) = 4 Or Val(Wempresa) = 8 Then
                For aa = Impre To 16
                    Impre = Impre + 1
                    Print #1, ""
                Next aa
                    Else
                For aa = Impre To 16
                    Impre = Impre + 1
                    Print #1, ""
                Next aa
            End If
            
            If Val(Wempresa) = 1 Then
            
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "15" + Chr$(72);
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Impre = Impre + 7
                
                    Else
                    
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "15" + Chr$(72);
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Impre = Impre + 7
                
            End If
        
        End If
        
        Select Case Val(Wempresa)
            Case 4, 8
                If FF = 2 Then
                    For aa = Impre To 16
                        Print #1, ""
                    Next aa
        
                    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                    Print #1, Tab(4); "Pellital S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                    Print #1, Tab(4); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                    Print #1, Tab(4); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                    Print #1, Tab(4); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                    Print #1, Tab(4); "vigentes del ámbito municipal, provincial y nacional"
                    Print #1, Tab(4); "Declaramos que los productos estan adecuadamente acondicionados para soportar los riesgos nosmales de la carga, "
                    Print #1, Tab(4); "transporte, transbordo, descarga y estiba, cumpliendo la reglamentacion en vigor"
                    
                End If
        
                For XDa = 1 To 1
                        For DA = 1 To 9
                                If Val(Stk(DA, 4)) <> 0 Then
                                        
                                        Select Case DA
                                                Case 1
                                                        Lugar = 25
                                                Case 2
                                                        Lugar = 36
                                                Case 3
                                                        Lugar = 47
                                                Case 4
                                                        Lugar = 58
                                                Case 5
                                                        Lugar = 69
                                                Case 6
                                                        Lugar = 80
                                                Case 7
                                                        Lugar = 92
                                                Case 8
                                                        Lugar = 104
                                                Case 9
                                                        Lugar = 116
                                                Case Else
                                        End Select
                                                            
                                        If DA = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(DA, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(DA, XDa);
                                        End If
                                    End If
        
                        Next DA
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For DA = 1 To 9
            
                                If Val(Stk(DA, 4)) <> 0 Then
        
                                        Select Case DA
                                            Case 1
                                                Lugar = 16
                                            Case 2
                                                Lugar = 23
                                            Case 3
                                                Lugar = 31
                                            Case 4
                                                Lugar = 38
                                            Case 5
                                                Lugar = 45
                                            Case 6
                                                Lugar = 52
                                            Case 7
                                                Lugar = 59
                                            Case 8
                                                Lugar = 66
                                            Case 9
                                                Lugar = 73
                                            Case Else
                                    End Select
        
                                    If Val(Stk(DA, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(DA, XDa))));
                                    End If
        
                            End If
                    Next DA
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, ""
                Print #1, ""
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
            
            Case Else
                If FF = 2 Then
                    For aa = Impre To 16
                        Print #1, ""
                    Next aa
        
                    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                    Print #1, Tab(3); "Surfactan S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                    Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                    Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                    Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                    Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional."
                    Print #1, Tab(3); "Declaramos que los productos estan adecuadamente acondicionados para soportar los riesgos nosmales de la carga, "
                    Print #1, Tab(3); "transporte, transbordo, descarga y estiba, cumpliendo la reglamentacion en vigor"
                    
                End If
        
                For XDa = 1 To 1
                        For DA = 1 To 9
                                If Val(Stk(DA, 4)) <> 0 Then
                                        
                                        Select Case DA
                                                Case 1
                                                        Lugar = 22
                                                Case 2
                                                        Lugar = 33
                                                Case 3
                                                        Lugar = 44
                                                Case 4
                                                        Lugar = 55
                                                Case 5
                                                        Lugar = 66
                                                Case 6
                                                        Lugar = 77
                                                Case 7
                                                        Lugar = 89
                                                Case 8
                                                        Lugar = 101
                                                Case 9
                                                        Lugar = 113
                                                Case Else
                                        End Select
                                                            
                                        If DA = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(DA, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(DA, XDa);
                                        End If
                                    End If
        
                        Next DA
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For DA = 1 To 9
            
                                If Val(Stk(DA, 4)) <> 0 Then
        
                                        Select Case DA
                                            Case 1
                                                Lugar = 14
                                            Case 2
                                                Lugar = 21
                                            Case 3
                                                Lugar = 29
                                            Case 4
                                                Lugar = 36
                                            Case 5
                                                Lugar = 43
                                            Case 6
                                                Lugar = 50
                                            Case 7
                                                Lugar = 57
                                            Case 8
                                                Lugar = 64
                                            Case 9
                                                Lugar = 71
                                            Case Else
                                    End Select
        
                                    If Val(Stk(DA, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(DA, XDa))));
                                    End If
        
                            End If
                    Next DA
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
                
        End Select
    
    Next FF
    
    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    Rem Print #1, Chr$(27) & "&l1H"
    
    Close #1

    Rem Open "lpt1:" For Output As #1
    Rem Print #1, Chr$(27) & "&l1H"
    Rem Print #1, "bandeja 2"
    Rem Close #1


End Sub


Sub Impresion_Remito_Cabecera()

    Print #1, ""
    If FF = 2 Then
        Print #1, ""
    End If
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    If Val(Wempresa) = 1 Then
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
            Else
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
    End If
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Tab(7); WRazon
    Print #1, Tab(7); Left$(WDireccion, 33)
    Print #1, Tab(7); Left$(WLocalidad, 33);
    Print #1, Tab(44); Pedido.Text;
    Print #1, Tab(57); Cliente.Text;
    Print #1, Tab(68); ""
    Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
    Print #1, ""
    Print #1, Tab(7); Iva(Val(WCodIva));
    Print #1, Tab(48); WCuit
    Print #1, ""
    Print #1, Tab(30); WDirentrega;
    Print #1, ""
    If FF = 2 Then
        Print #1, Tab(60); "ORIGINAL"
            Else
        Print #1, Tab(60); "DUPLICADO"
    End If
    Print #1, ""
    
End Sub

Sub Impresion_Remito_Calculo()

    Open "Verifica.TXT" For Output As #1
    
    Impre = 0
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
        
            Suma = Suma + 1
            
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
            
                ZClase = ""
                ZIntervencion = ""
                ZNaciones = ""
                spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                    ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                    ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                    ZDesciOnu = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
                    ZClase = Trim(ZClase)
                    ZIntervencion = Trim(ZIntervencion)
                    ZNaciones = Trim(ZNaciones)
                    rstTerminado.Close
                End If
                
                ZImpre = ""
                If Trim(ZClase) <> "" Then
                    ZImpre = "Guia:" + ZIntervencion + " N.ONU:" + ZNaciones + " Clase:" + ZClase
                    ZImpre = Left$(ZImpre, 32)
                End If
        
                Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
                Print #1, Tab(10); Left$(Descri, 40);
                Print #1, Tab(51); ZImpre;
                Print #1, Tab(85); Alinea("#####.##", Str$(Cantidad));
                Print #1, " Kg."
                Impre = Impre + 1
                    
            End If
            
            ZLote1 = XLote(Suma, 1)
            ZCantidad1 = XLote(Suma, 2)
            ZLote2 = XLote(Suma, 3)
            ZCantidad2 = XLote(Suma, 4)
            ZLote3 = XLote(Suma, 5)
            ZCantidad3 = XLote(Suma, 6)
            ZLote4 = XLote(Suma, 7)
            ZCantidad4 = XLote(Suma, 8)
            ZLote5 = XLote(Suma, 9)
            ZCantidad5 = XLote(Suma, 10)
            ZLote6 = XLote(Suma, 11)
            ZCantidad6 = XLote(Suma, 12)
            ZLote7 = XLote(Suma, 13)
            ZCantidad7 = XLote(Suma, 14)
            ZLote8 = XLote(Suma, 15)
            ZCantidad8 = XLote(Suma, 16)
            ZLote9 = XLote(Suma, 17)
            ZCantidad9 = XLote(Suma, 18)
            ZLote10 = XLote(Suma, 19)
            ZCantidad10 = XLote(Suma, 20)
            ZLote11 = XLote(Suma, 21)
            ZCantidad11 = XLote(Suma, 22)
            ZLote12 = XLote(Suma, 23)
            ZCantidad12 = XLote(Suma, 24)
            
            If Trim(Producto) <> "" Then
            
                If Left$(Producto, 2) = "DY" Then
                
                    For ZZZCiclo = 1 To 12
                    
                        Select Case ZZZCiclo
                            Case 1
                                ZZZLote = ZLote1
                            Case 2
                                ZZZLote = ZLote2
                            Case 3
                                ZZZLote = ZLote3
                            Case 4
                                ZZZLote = ZLote4
                            Case 5
                                ZZZLote = ZLote5
                            Case 6
                                ZZZLote = ZLote6
                            Case 7
                                ZZZLote = ZLote7
                            Case 8
                                ZZZLote = ZLote8
                            Case 9
                                ZZZLote = ZLote9
                            Case 10
                                ZZZLote = ZLote10
                            Case 11
                                ZZZLote = ZLote11
                            Case Else
                                ZZZLote = ZLote12
                        End Select
                
                        ZZZArti = Left$(Producto, 3) + Right$(Producto, 7)
                        XParam = "'" + ZZZLote + "','" _
                                     + ZZZArti + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZZPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            rstLaudo.Close
                            
                            If Trim(ZZZPartiOri) <> "" Then
                            
                                Select Case ZZZCiclo
                                    Case 1
                                        ZLote1 = ZZZPartiOri
                                    Case 2
                                        ZLote2 = ZZZPartiOri
                                    Case 3
                                        ZLote3 = ZZZPartiOri
                                    Case 4
                                        ZLote4 = ZZZPartiOri
                                    Case 5
                                        ZLote5 = ZZZPartiOri
                                    Case 6
                                        ZLote6 = ZZZPartiOri
                                    Case 7
                                        ZLote7 = ZZZPartiOri
                                    Case 8
                                        ZLote8 = ZZZPartiOri
                                    Case 9
                                        ZLote9 = ZZZPartiOri
                                    Case 10
                                        ZLote10 = ZZZPartiOri
                                    Case 11
                                        ZLote11 = ZZZPartiOri
                                    Case Else
                                        ZLote12 = ZZZPartiOri
                                End Select
                                
                            End If
                        End If
                        
                    Next ZZZCiclo
                    
                        Else
                        
                    For ZZZCiclo = 1 To 12
                    
                        Select Case ZZZCiclo
                            Case 1
                                ZZZLote = ZLote1
                            Case 2
                                ZZZLote = ZLote2
                            Case 3
                                ZZZLote = ZLote3
                            Case 4
                                ZZZLote = ZLote4
                            Case 5
                                ZZZLote = ZLote5
                            Case 6
                                ZZZLote = ZLote6
                            Case 7
                                ZZZLote = ZLote7
                            Case 8
                                ZZZLote = ZLote8
                            Case 9
                                ZZZLote = ZLote9
                            Case 10
                                ZZZLote = ZLote10
                            Case 11
                                ZZZLote = ZLote11
                            Case Else
                                ZZZLote = ZLote12
                        End Select
                        
                        XParam = "'" + ZZZLote + "','" _
                                     + Producto + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            ZZZPartiOri = IIf(IsNull(rstHoja!LoteColorante), "", rstHoja!LoteColorante)
                            
                            rstHoja.Close
                            
                            If Trim(ZZZPartiOri) <> "" Then
                            
                                Select Case ZZZCiclo
                                    Case 1
                                        ZLote1 = ZZZPartiOri
                                    Case 2
                                        ZLote2 = ZZZPartiOri
                                    Case 3
                                        ZLote3 = ZZZPartiOri
                                    Case 4
                                        ZLote4 = ZZZPartiOri
                                    Case 5
                                        ZLote5 = ZZZPartiOri
                                    Case 6
                                        ZLote6 = ZZZPartiOri
                                    Case 7
                                        ZLote7 = ZZZPartiOri
                                    Case 8
                                        ZLote8 = ZZZPartiOri
                                    Case 9
                                        ZLote9 = ZZZPartiOri
                                    Case 10
                                        ZLote10 = ZZZPartiOri
                                    Case 11
                                        ZLote11 = ZZZPartiOri
                                    Case Else
                                        ZLote12 = ZZZPartiOri
                                End Select
                                
                            End If
                        End If
                        
                    Next ZZZCiclo
                    
                End If
            
            End If
            
            
            ZEnv1 = XLote(Suma, 31)
            ZCantiEnv1 = XLote(Suma, 32)
            ZEnv2 = XLote(Suma, 33)
            ZCantiEnv2 = XLote(Suma, 34)
            ZEnv3 = XLote(Suma, 35)
            ZCantiEnv3 = XLote(Suma, 36)
            ZEnv4 = XLote(Suma, 37)
            ZCantiEnv4 = XLote(Suma, 38)
            ZEnv5 = XLote(Suma, 39)
            ZCantiEnv5 = XLote(Suma, 40)
            ZEnv6 = XLote(Suma, 41)
            ZCantiEnv6 = XLote(Suma, 42)
            ZEnv7 = XLote(Suma, 43)
            ZCantiEnv7 = XLote(Suma, 44)
            ZEnv8 = XLote(Suma, 45)
            ZCantiEnv8 = XLote(Suma, 46)
            ZEnv9 = XLote(Suma, 47)
            ZCantiEnv9 = XLote(Suma, 48)
            ZEnv10 = XLote(Suma, 49)
            ZCantiEnv10 = XLote(Suma, 50)
            ZEnv11 = XLote(Suma, 51)
            ZCantiEnv11 = XLote(Suma, 52)
            ZEnv12 = XLote(Suma, 53)
            ZCantiEnv12 = XLote(Suma, 54)
            
            ZDescri1 = ""
            ZDescri2 = ""
            ZDescri3 = ""
            ZDescri4 = ""
            ZDescri5 = ""
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv1 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri1 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv2 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri2 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv3 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri3 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv4 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri4 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv5 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri5 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv6 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri6 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv7 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri7 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv8 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri8 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv9 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri9 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv10 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri10 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv11 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri11 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv12 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri12 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            If Val(ZCantidad1) <> 0 Then
            
                Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72);
                
                Print #1, Tab(10); Alinea("#####", ZCantidad1); " Kg";
                Print #1, Tab(19); "Lote:"; Left$(ZLote1, 8);
                If Val(ZCantiEnv1) <> 0 Then
                    Print #1, Tab(33); Alinea("##", ZCantiEnv1);
                    Print #1, " x "; ZDescri1;
                End If
            
                If Val(ZCantidad2) <> 0 Then
                    Print #1, Tab(48); "|";
                    Print #1, Tab(50); Alinea("#####", ZCantidad2); " Kg";
                    Print #1, Tab(59); "Lote:"; Left$(ZLote2, 8);
                    If Val(ZCantiEnv2) <> 0 Then
                        Print #1, Tab(71); Alinea("##", ZCantiEnv2);
                        Print #1, " x "; ZDescri2;
                    End If
                End If
            
                If Val(ZCantidad3) <> 0 Then
                    Print #1, Tab(86); "|";
                    Print #1, Tab(88); Alinea("#####", ZCantidad3); " Kg";
                    Print #1, Tab(97); "Lote:"; Left$(ZLote3, 8);
                    If Val(ZCantiEnv3) <> 0 Then
                        Print #1, Tab(109); Alinea("##", ZCantiEnv3);
                        Print #1, " x "; ZDescri3;
                    End If
                End If
                    
                Print #1, ""
                Impre = Impre + 1
                
            End If
            
            If Val(ZCantidad4) <> 0 Then
            
                Print #1, Tab(10); Alinea("#####", ZCantidad4); " Kg";
                Print #1, Tab(19); "Lote:"; Left$(ZLote4, 8);
                If Val(ZCantiEnv4) <> 0 Then
                    Print #1, Tab(33); Alinea("##", ZCantiEnv4);
                    Print #1, " x "; ZDescri4;
                End If
            
                If Val(ZCantidad5) <> 0 Then
                    Print #1, Tab(48); "|";
                    Print #1, Tab(50); Alinea("#####", ZCantidad5); " Kg";
                    Print #1, Tab(59); "Lote:"; Left$(ZLote5, 8);
                    If Val(ZCantiEnv5) <> 0 Then
                        Print #1, Tab(71); Alinea("##", ZCantiEnv5);
                        Print #1, " x "; ZDescri5;
                    End If
                End If
            
                If Val(ZCantidad6) <> 0 Then
                    Print #1, Tab(86); "|";
                    Print #1, Tab(88); Alinea("#####", ZCantidad6); " Kg";
                    Print #1, Tab(97); "Lote:"; Left$(ZLote6, 8);
                    If Val(ZCantiEnv6) <> 0 Then
                        Print #1, Tab(109); Alinea("##", ZCantiEnv6);
                        Print #1, " x "; ZDescri6;
                    End If
                End If
                    
                Print #1, ""
                Impre = Impre + 1
                
            End If
            
            If Val(ZCantidad7) <> 0 Then
            
                Print #1, Tab(10); Alinea("#####", ZCantidad7); " Kg";
                Print #1, Tab(19); "Lote:"; Left$(ZLote7, 8);
                If Val(ZCantiEnv7) <> 0 Then
                    Print #1, Tab(33); Alinea("##", ZCantiEnv7);
                    Print #1, " x "; ZDescri7;
                End If
            
                If Val(ZCantidad8) <> 0 Then
                    Print #1, Tab(48); "|";
                    Print #1, Tab(50); Alinea("#####", ZCantidad8); " Kg";
                    Print #1, Tab(59); "Lote:"; Left$(ZLote8, 8);
                    If Val(ZCantiEnv8) <> 0 Then
                        Print #1, Tab(71); Alinea("##", ZCantiEnv8);
                        Print #1, " x "; ZDescri8;
                    End If
                End If
            
                If Val(ZCantidad9) <> 0 Then
                    Print #1, Tab(86); "|";
                    Print #1, Tab(88); Alinea("#####", ZCantidad9); " Kg";
                    Print #1, Tab(97); "Lote:"; Left$(ZLote9, 8);
                    If Val(ZCantiEnv9) <> 0 Then
                        Print #1, Tab(109); Alinea("##", ZCantiEnv9);
                        Print #1, " x "; ZDescri9;
                    End If
                End If
                    
                Print #1, ""
                Impre = Impre + 1
                
            End If
            
            If Val(ZCantidad10) <> 0 Then
            
                Print #1, Tab(10); Alinea("#####", ZCantidad10); " Kg";
                Print #1, Tab(19); "Lote:"; Left$(ZLote10, 8);
                If Val(ZCantiEnv10) <> 0 Then
                    Print #1, Tab(33); Alinea("##", ZCantiEnv10);
                    Print #1, " x "; ZDescri10;
                End If
            
                If Val(ZCantidad11) <> 0 Then
                    Print #1, Tab(48); "|";
                    Print #1, Tab(50); Alinea("#####", ZCantidad11); " Kg";
                    Print #1, Tab(59); "Lote:"; Left$(ZLote11, 8);
                    If Val(ZCantiEnv11) <> 0 Then
                        Print #1, Tab(71); Alinea("##", ZCantiEnv11);
                        Print #1, " x "; ZDescri11;
                    End If
                End If
            
                If Val(ZCantidad12) <> 0 Then
                    Print #1, Tab(86); "|";
                    Print #1, Tab(88); Alinea("#####", ZCantidad12); " Kg";
                    Print #1, Tab(97); "Lote:"; Left$(ZLote12, 8);
                    If Val(ZCantiEnv12) <> 0 Then
                        Print #1, Tab(109); Alinea("##", ZCantiEnv12);
                        Print #1, " x "; ZDescri12;
                    End If
                End If
                    
                Print #1, ""
                Impre = Impre + 1
                
            End If
            
        Next iRow
        
    Next a
    
    ZZPasaImpre = Impre
    
    Close #1

End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo

    Erase Stk
    
    If Val(Wempresa) = 8 Then
        Stk(1, 1) = "005"
        Stk(2, 1) = "011"
        Stk(3, 1) = "021"
        Stk(4, 1) = "027"
        Stk(5, 1) = "004"
        Stk(6, 1) = "012"
        Stk(7, 1) = "002"
        Stk(8, 1) = "000"
        Stk(9, 1) = "000"
            Else
        Stk(1, 1) = "020"
        Stk(2, 1) = "021"
        Stk(3, 1) = "022"
        Stk(4, 1) = "023"
        Stk(5, 1) = "031"
        Stk(6, 1) = "025"
        Stk(7, 1) = "026"
        Stk(8, 1) = "030"
        Stk(9, 1) = "028"
    End If

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For DA = 1 To 9
                        If Val(Stk(DA, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) + !Cantidad)
                                    Else
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next DA
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If
    
    Rem pone en cero los saldos negativos de stock
    Rem de envase en comodato en clientes
    For DA = 1 To 9
        If Val(Stk(DA, 2)) < 0 Then
            Stk(DA, 2) = "0"
        End If
    Next DA

End Sub

Private Sub Verifica_Lote()

    Renglon = 0
    Renglon1 = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    ZVeriSedronar = "N"
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
            
            Suma = Suma + 1
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            WTipoProDy = Left$(Articulo, 2)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
                    
            If Cantidad <> 0 Then
            
                ZSedronar = 0
                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZSedronar = IIf(IsNull(rstTerminado!Sedronar), "0", rstTerminado!Sedronar)
                    rstTerminado.Close
                End If
                
                If ZSedronar = 1 Then
                    ZVeriSedronar = "S"
                End If
            
                WEstado = "N"
                SumaCant = 0
                
                For ZZCiclo = 1 To 23 Step 2
    
                    WLote = XLote(Suma, ZZCiclo)
                    WImpo = Val(XLote(Suma, ZZCiclo + 1))
                    WCanti = Str$(WImpo)
    
                    If Val(WLote) <> 0 Then
                        SumaCant = SumaCant + Val(WCanti)
                    End If
                    
                Next ZZCiclo
    
                ZZValor1 = SumaCant
                ZZValor2 = Cantidad
                Call Redondeo(ZZValor1)
                Call Redondeo(ZZValor2)
    
                If ZZValor1 = ZZValor2 Then
                    WEstado = "S"
                        Else
                    WEstado = "N"
                    m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                    Exit Sub
                End If
    
                If WEstado = "S" Then
    
                    Erase ControlLote
                    ZZRenglon = 0
                    For ZZCiclo = 1 To 23 Step 2
                        ZZRenglon = ZZRenglon + 1
                        ControlLote(ZZRenglon, 1) = XLote(Suma, ZZCiclo)
                        ControlLote(ZZRenglon, 2) = XLote(Suma, ZZCiclo + 1)
                    Next ZZCiclo
    
                    For Ciclo1 = 1 To 12
                        If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                            For Ciclo2 = 1 To 12
                                If Ciclo1 <> Ciclo2 Then
                                    If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                                        m$ = "A asignado una misma partida 2 veces"
                                        a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        Rem WEstado = "N"
                                        Exit Sub
                                    End If
                                End If
                            Next Ciclo2
                        End If
                        If WEstado = "N" Then
                            Exit For
                        End If
                    Next Ciclo1
        
                End If

                If WEstado = "S" Then
    
                    Erase ControlLote
                    ZZRenglon = 0
                    For ZZCiclo = 1 To 23 Step 2
                        ZZRenglon = ZZRenglon + 1
                        ControlLote(ZZRenglon, 1) = XLote(Suma, ZZCiclo)
                        ControlLote(ZZRenglon, 2) = XLote(Suma, ZZCiclo + 1)
                    Next ZZCiclo
    
                    For Ciclo1 = 1 To 12
    
                        WLote = ControlLote(Ciclo1, 1)
                        WCanti = Val(ControlLote(Ciclo1, 2))
            
                        If Val(WLote) <> 0 Or Val(WCanti) <> 0 Then
            
                        If Left$(Articulo, 2) <> "PT" Then
                            WTipopro = "M"
                                Else
                            WTipopro = "T"
                        End If
            
                        Select Case WTipopro
                            Case "M"
                            
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
                                        Case "TA"
                                            Wempresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            Wempresa = "0007"
                                            txtOdbc = "Empresa07"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                End If
                            
                                WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                                WEntra = "N"
                                XParam = "'" + WLote + "','" _
                                             + WArti + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WSal = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Call Redondeo(WSal)
                                    WEntra = "S"
                                    rstLaudo.Close
                                    If WSal < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        WEstado = "N"
                                        Call Conecta_Empresa
                                        Exit Sub
                                    End If
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArti + "','" _
                                                 + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        rstMovguia.Close
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Call Conecta_Empresa
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                
                                Call Conecta_Empresa
                                
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                            Case Else
                                WEntra = "N"
                                WControla = 0
                                
                                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                    rstTerminado.Close
                                End If
                
                                If WControla = 0 Then
                                
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
                                            Case "TA"
                                                Wempresa = "0003"
                                                txtOdbc = "Empresa03"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                                Wempresa = "0007"
                                                txtOdbc = "Empresa07"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        End Select
                                    End If
                                
                                    XParam = "'" + WLote + "','" _
                                            + Articulo + "'"
                                    spHoja = "ListaHojaProducto " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        rstHoja.Close
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Call Conecta_Empresa
                                            Exit Sub
                                        End If
                                    End If
                
                                    If WEntra = "N" Then
                                        XParam = "'" + Articulo + "','" _
                                                    + WLote + "'"
                                        spMovguia = "ListaMovguiaLote1 " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                            WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                            Call Redondeo(WSal)
                                            WEntra = "S"
                                            rstMovguia.Close
                                            If WSal < WCanti Then
                                                m$ = "La cantidad informada supera al saldo disponible"
                                                a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                                WEstado = "N"
                                                Call Conecta_Empresa
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                    
                                    Call Conecta_Empresa
                
                                        Else
                                        
                                    WEntra = "S"
                                    
                                End If
                                
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                        End Select
            
                        End If
            
                    Next Ciclo1

                End If
                
            End If
                                        
        Next iRow
            
    Next a
    
    If ZVeriSedronar = "S" Then
        ZNroSedronar = ""
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZNroSedronar = Trim(IIf(IsNull(rstCliente!NroSedronar), "", rstCliente!NroSedronar))
            rstCliente.Close
        End If
        If Trim(ZNroSedronar) = "" Then
            m$ = "Atencion: El cliente debe estar inscripto en el Sedronar para adquirir estos productos"
            Aaa% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
            WEstado = "N"
            Exit Sub
        End If
    End If
    
    
    
    
End Sub



Private Sub Impresion_Varios()

    Rem toto

    Erase ZZEnviaPdf
    ZZLugarEnvia = 0
    Erase ZZEnviaPdfII
    ZZLugarEnviaII = 0


    ZZVersion = 0
    FileCopy "w:\base.pdf", "c:\pdfprint\base.pdf"
    FileCopy "w:\base.doc", "c:\pdfprint\base.doc"
    Kill "c:\pdfprint\*.pdf"
    Kill "c:\pdfprint\*.doc"
    FileCopy "w:\base.pdf", "c:\pdfprintii\base.pdf"
    FileCopy "w:\base.doc", "c:\pdfprintii\base.doc"
    Kill "c:\pdfprintii\*.pdf"
    Kill "c:\pdfprintii\*.doc"
    
    ZZImprePdf = "N"
    
    ZEmailFactura = ""
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZEmailFactura = IIf(IsNull(rstCliente!EmailFactura), "", rstCliente!EmailFactura)
        rstCliente.Close
    End If
    

    Rem BY NAN
    Rem  ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
    Rem  ZZEstado = Dir(ZZRuta)
    Rem   ZZEstado = Trim(ZZEstado)
    Rem If ZZEstado <> "" Then
    Rem     ZZVersion = 1
    Rem         Else
    Rem     ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe"
    Rem     ZZEstado = Dir(ZZRuta)
    Rem     ZZEstado = Trim(ZZEstado)
    Rem     If ZZEstado <> "" Then
    Rem         ZZVersion = 2
    Rem             Else
    Rem         ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe"
    Rem         ZZEstado = Dir(ZZRuta)
    Rem         ZZEstado = Trim(ZZEstado)
    Rem         If ZZEstado <> "" Then
    Rem             ZZVersion = 3
    Rem                 Else
    Rem             ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe"
    Rem             ZZEstado = Dir(ZZRuta)
    Rem             ZZEstado = Trim(ZZEstado)
    Rem             If ZZEstado <> "" Then
    Rem                 ZZVersion = 4
    Rem                     Else
    Rem                 ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe"
    Rem                 ZZEstado = Dir(ZZRuta)
    Rem                 ZZEstado = Trim(ZZEstado)
    Rem                 If ZZEstado <> "" Then
    Rem                     ZZVersion = 5
    Rem                       Rem by nan
    Rem                       Else
    Rem                  ZZRuta = "C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe"
    Rem                  ZZEstado = Dir(ZZRuta)
    Rem                  ZZEstado = Trim(ZZEstado)
    Rem                      If ZZEstado <> "" Then
    Rem                        ZZVersion = 6
    Rem                     End If
    Rem                     Rem fin by nan
    Rem
    Rem                 End If
    Rem             End If
    Rem         End If
    Rem     End If
    Rem End If
    
    Erase ZImpreFicha
    ZLugarFicha = 0
    
    For DA = 1 To 99
    
        Articulo = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        Precio = Auxiliar(DA, 3)
        WRenglon = Auxiliar(DA, 4)
        WLote1 = Auxiliar(DA, 5)
        WCanti1 = Auxiliar(DA, 6)
        WLote2 = Auxiliar(DA, 7)
        WCanti2 = Auxiliar(DA, 8)
        Wlote3 = Auxiliar(DA, 9)
        WCanti3 = Auxiliar(DA, 10)
        WLote4 = Auxiliar(DA, 11)
        WCanti4 = Auxiliar(DA, 12)
        WLote5 = Auxiliar(DA, 13)
        WCanti5 = Auxiliar(DA, 14)
        WLote6 = Auxiliar(DA, 15)
        WCanti6 = Auxiliar(DA, 16)
        WLote7 = Auxiliar(DA, 17)
        WCanti7 = Auxiliar(DA, 18)
        WLote8 = Auxiliar(DA, 19)
        WCanti8 = Auxiliar(DA, 20)
        WLote9 = Auxiliar(DA, 21)
        WCanti9 = Auxiliar(DA, 22)
        WLote10 = Auxiliar(DA, 23)
        WCanti10 = Auxiliar(DA, 24)
        WLote11 = Auxiliar(DA, 25)
        WCanti11 = Auxiliar(DA, 26)
        WLote12 = Auxiliar(DA, 27)
        WCanti12 = Auxiliar(DA, 28)
        
        ZZDescriArticulo = Trim(Auxiliar(DA, 30))
        ZZDescriArticuloPDF = ""
        ZZZHasta = Len(ZZDescriArticulo)
        For ZZZCiclo = 1 To ZZZHasta
            If Mid$(ZZDescriArticulo, ZZZCiclo, 1) <> Space(1) Then
                ZZDescriArticuloPDF = ZZDescriArticuloPDF + Mid$(ZZDescriArticulo, ZZZCiclo, 1)
            End If
        Next ZZZCiclo
        
        If Articulo = "" Then Exit For
    
        Rem
        Rem Hoja de Seguridad
        Rem
        
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
            
            If Left$(UCase(Articulo), 2) = "PT" Then

                Es = ZZDescriArticulo
                x = ""
                For XX = 1 To Len(Es)
                    Y = Mid$(Es, XX, 1)
                    If Y <> " " And Y <> "/" Then
                        x = x + Y
                    End If
                Next
                ZZCodArt = x + Mid$(Articulo, 4, 5) + Right$(Articulo, 3)
                
                    Else
                    
                ZZCodArt = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                
            End If
            
            ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArt + ".PDF"
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
                    
                If ZZAviso = "S" Then
                    m$ = "El MSDS  (" + ZZCodArt + ")  no se ha encontrado"
                    a% = MsgBox(m$, 0, "Impresion de comprobantes varios")
                End If
                
            End If
            
        End If
    
                    
        Rem
        Rem ficha de intevencion
        Rem
        If Left$(UCase(Articulo), 2) = "PT" Then
                
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
                    
                    
                    
                    
        Rem *******************certificados*****************
    
        Rem
        Rem certificado de analisis
        Rem
        Rem For ZZCiclo = 1 To 12
        Rem sacar dos sentenvcias de abajo
        Rem y restaurar el ciclo ok
    
        Rem ZZRequiereCertificado = 0
        
        If ZZRequiereCertificado = 1 Then
            ZZZZhasta = 12
                Else
            ZZZZhasta = 0
        End If
        
        For ZZCiclo = 1 To ZZZZhasta
            
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
            
            If Val(Auxiliar(DA, ZZLugar)) <> 0 Then
        
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
            
                    Rem If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                    If XTipoPro = "CO" Then
                    
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
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
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
                                        
                                        If ZZFechaRevalida <> "  /  /    " And ZZFechaRevalida <> "00/00/0000" Then
                                            WFecha = ZZFechaRevalida
                                            WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        End If
                                        
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
                                    
                                        Sql1 = "Select Ensayo1,ensayo2,ensayo3,ensayo4,ensayo5,ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
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
                                            
                                           Rem by nan
                                           rstEspecifUnifica.Close
                                          
                                          
                                        End If
                                          
                                        Rem by nan 21-5-2014
                                        Sql1 = "Select desde1,desde2,desde3,desde4,desde5,desde6,desde7,desde8,desde9,desde10,hasta1,hasta2,hasta3,hasta4,hasta5,hasta6,hasta7,hasta8,hasta9,hasta10,version"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                  
                                                                      
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
                                                                
                                            ZVersion = rstEspecifUnifica!Version
                                            rstEspecifUnifica.Close
                                            LlamaImprime = "S"
                                        End If
                                
                                    End If
                                    
                                    If LlamaImprime = "S" Then
                                        
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(1) = rstEnsayo!Descripcion
                                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(2) = rstEnsayo!Descripcion
                                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(3) = rstEnsayo!Descripcion
                                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(4) = rstEnsayo!Descripcion
                                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(5) = rstEnsayo!Descripcion
                                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(6) = rstEnsayo!Descripcion
                                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(7) = rstEnsayo!Descripcion
                                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(8) = rstEnsayo!Descripcion
                                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(9) = rstEnsayo!Descripcion
                                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(10) = rstEnsayo!Descripcion
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
                                        
                                        Rem If ZZImpreVtoTermi = 0 Then
                                        Rem     If ZImpreVto <> 1 Then
                                        Rem         Rem WFechaElaboracion = ""
                                        Rem     End If
                                        Rem End If
                        
                                        Rem
                                        Rem SI ES COLORANTE NO IMPRIME
                                        Rem LA FECHA DE VENCIMIENTO
                                        Rem
                                        XCodigo = Val(Mid$(ZProducto, 4, 5))
                                        XTipoPro = ""
                                        If Val(Wempresa) = 1 Then
                                            If XCodigo >= 0 And XCodigo <= 999 Then
                                                WFechaElaboracion = ""
                                                XTipoPro = "CO"
                                                    Else
                                                If XCodigo >= 11000 And XCodigo <= 12999 Then
                                                    WFechaElaboracion = ""
                                                    XTipoPro = "CO"
                                                        Else
                                                    XTipoPro = ""
                                                End If
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
                                                
                                                If Val(ZStd(CiclaMetodo, 3)) <> 0 And Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                                    ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo)) + " " + Left$(ZStd(CiclaMetodo, 1), 50)
                                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                        Else
                                                    ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                End If
                                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 80)
                                                
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
                                                            ZExamenII = Mid(ZExamen, Cicla + 1, 25)
                                                            ZExamen = Mid(ZExamen, 1, Cicla)
                                                            Exit For
                                                        End If
                                                    Next Cicla
                                                End If
                                                        
                                                        
                                                        
                                                aa = Auxiliar(1, 1)
                                                aa = Auxiliar(1, 2)
                                                aa = Auxiliar(1, 3)
                                                aa = Auxiliar(1, 4)
                                                aa = Auxiliar(1, 5)
                                                aa = Auxiliar(1, 6)
                                                aa = Auxiliar(1, 7)
                                                aa = Auxiliar(1, 8)
                                                aa = Auxiliar(1, 9)
                                                aa = Auxiliar(1, 10)
                                                aa = Auxiliar(1, 11)
                                                aa = Auxiliar(1, 12)
                                                aa = Auxiliar(1, 13)
                                                aa = Auxiliar(1, 14)
                                                aa = Auxiliar(1, 15)
                                                aa = Auxiliar(1, 16)
                                                aa = Auxiliar(1, 17)
                                                aa = Auxiliar(1, 18)
                                                aa = Auxiliar(1, 19)
                                                aa = Auxiliar(1, 20)
                                                        
                                                        
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
                                            Listado.ReportFileName = "CertificadoNuevo.rpt"
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
                                        Rem BY NAN 29-4-2015
                                        ZZDescriArticuloPDF = Left(ZZDescriArticuloPDF, 12)
                                        
                                        If Trim(ZEmailFactura) <> "" Then
                                            Listado.ReportFileName = "Certificadopdf.rpt"
                                            Listado.Destination = crptToFile
                                            Listado.PrintFileType = crptWinWord
                                            Listado.PrintFileName = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".doc"
                                            ZZDesdedoc = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".doc"
                                            ZZDesdePdf = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".pdf"
                                        End If
                        
                                        Listado.Connect = Connect()
                                        Listado.Action = 1
                                   
                                        If Trim(ZEmailFactura) <> "" Then
                                            ZZLugarEnviaII = ZZLugarEnviaII + 1
                                            ZZEnviaPdfII(ZZLugarEnviaII, 1) = Articulo
                                            ZZEnviaPdfII(ZZLugarEnviaII, 2) = ZZDesdePdf
                                            ZZEnviaPdfII(ZZLugarEnviaII, 3) = ZZDesdedoc
                                            ZZEnviaPdfII(ZZLugarEnviaII, 4) = ZZDescriArticulo
                                            ZZEnviaPdfII(ZZLugarEnviaII, 5) = ZLote
                                        End If
                                                
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
                        ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
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
                                
                                If Trim(ZEmailFactura) <> "" Then
                                    ZZRutaII = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".pdf"
                                    ZZRutadoc = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".doc"
                                End If
                                
                                FileCopy ZZRuta, ZZRutaII
                                ZZLugarEnvia = ZZLugarEnvia + 1
                                ZZEnviaPdf(ZZLugarEnvia, 1) = Articulo
                                ZZEnviaPdf(ZZLugarEnvia, 2) = ZZRutaII
                                ZZEnviaPdf(ZZLugarEnvia, 3) = ZZRutadoc
                                ZZEnviaPdf(ZZLugarEnvia, 4) = ZZDescriArticulo
                                ZZEnviaPdf(ZZLugarEnvia, 5) = ZZPartidaOri
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
                            ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
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
                                    
                                    If Trim(ZEmailFactura) <> "" Then
                                        ZZRutaII = "C:\pdfprintii\CertificadodeSeguridad" + ZZDescriArticulo + ZZPartiOri + ".pdf"
                                        ZZRutadoc = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".doc"
                                    End If
                                    
                                    FileCopy ZZRuta, ZZRutaII
                                    ZZLugarEnvia = ZZLugarEnvia + 1
                                    ZZEnviaPdf(ZZLugarEnvia, 1) = Articulo
                                    ZZEnviaPdf(ZZLugarEnvia, 2) = ZZRutaII
                                    ZZEnviaPdf(ZZLugarEnvia, 3) = ZZRutadoc
                                    ZZEnviaPdf(ZZLugarEnvia, 4) = ZZDescriArticulo
                                    ZZEnviaPdf(ZZLugarEnvia, 5) = ZZPartidaOri
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
        
        TiempoPausa = 3 ' Asigna hora de inicio.
        Inicio = Timer  ' Establece la hora de inicio.
        Do While Timer < Inicio + TiempoPausa
            DoEvents    ' Cambia a otros procesos.
        Loop
        
    End If
    
    
    If Trim(ZEmailFactura) <> "" Then
    
        ZZManda = "N"
        For ZZZCiclo = 1 To 100
            If Trim(ZZEnviaPdfII(ZZZCiclo, 2)) <> "" Then
                ZZManda = "S"
                Exit For
            End If
        Next ZZZCiclo
        
        If ZZManda = "S" Then
    
            ZZImpreAnterior = Printer.DeviceName
            Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "aPDF Writer" + Chr$(34)
            
            MiRuta = CurDir + "\"
            MiRutaII = Left$(CurDir, 1)
            
            Open "C:\pdfprintii\ImprePdf.bat" For Output As #1
            
            ZZZSuma = 0
            ZZZSumaII = 1
            Print #1, "c:"
            Print #1, "cd\archivos de programa\A-PDF Office to PDF"
            
            For ZZZCiclo = 1 To 100
                If Trim(ZZEnviaPdfII(ZZZCiclo, 2)) <> "" Then
                    Print #1, "OfficeToPDF.exe " + Trim(ZZEnviaPdfII(ZZZCiclo, 3)) + " " + Trim(ZZEnviaPdfII(ZZZCiclo, 2))
                    ZZLugarEnvia = ZZLugarEnvia + 1
                    ZZEnviaPdf(ZZLugarEnvia, 1) = ZZEnviaPdfII(ZZZCiclo, 1)
                    ZZEnviaPdf(ZZLugarEnvia, 2) = ZZEnviaPdfII(ZZZCiclo, 2)
                    ZZEnviaPdf(ZZLugarEnvia, 3) = ZZEnviaPdfII(ZZZCiclo, 3)
                    ZZEnviaPdf(ZZLugarEnvia, 4) = ZZEnviaPdfII(ZZZCiclo, 4)
                    ZZEnviaPdf(ZZLugarEnvia, 5) = ZZEnviaPdfII(ZZZCiclo, 5)
                End If
            Next ZZZCiclo
            
            Close #1
            
            Shell "C:\pdfprintii\Imprepdf.bat"
            
            For ZZZCiclo = 1 To 100
                If Trim(ZZEnviaPdfII(ZZZCiclo, 2)) <> "" Then
                    Do
                        ZZEstadoII = Dir(Trim(ZZDesdePdf))
                        If ZZEstadoII <> "" Then
                            Exit Do
                                Else
                            TiempoPausa = 1 ' Asigna hora de inicio.
                            Inicio = Timer  ' Establece la hora de inicio.
                            Do While Timer < Inicio + TiempoPausa
                                DoEvents    ' Cambia a otros procesos.
                            Loop
                        End If
                    Loop
                End If
            Next ZZZCiclo
            
            ChDrive MiRutaII
            ChDir MiRuta
            
            Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + ZZImpreAnterior + Chr$(34)
        
        End If
        
    End If
    
    PrgFactuRemito.Show
    Numero.SetFocus
    
End Sub


Private Sub Envio_Email()
    
    If Trim(ZEmailFactura) <> "" Then
    
        ZZManda = "N"
        For ZZZCiclo = 1 To 100
            If Trim(ZZEnviaPdf(ZZZCiclo, 2)) <> "" Then
                ZZManda = "S"
                Exit For
            End If
        Next ZZZCiclo
        
        If ZZManda = "S" Then
        
        
            ZZZSuma = 0
            ZZZSumaII = 1
            For ZZZCiclo = 1 To 100
                If Trim(ZZEnviaPdf(ZZZCiclo, 2)) <> "" Then
                    ZZZSuma = ZZZSuma + 1
                    If ZZZSuma = 4 Then
                        ZZZSuma = 1
                        ZZZSumaII = ZZZSumaII + 1
                    End If
                End If
            Next ZZZCiclo
        
            ZZZNumero = Right$(Trim(Numero), 5)
        
            ZTexto1 = "Se adjuntan los certificados de analisis correspondientes a los"
            ZTexto6 = "productos detallados en la factura numero " + ZZZNumero
            ZTexto7 = ""
            ZTexto8 = "Surfactan S.A."
            ZTexto9 = "011-4714-4097"
        
        
            sTo = ZEmailFactura
            Rem sTo = "d_esquenazi@yahoo.com"
            sCC = ""
            sBCC = ""
            sSubject = "Certificados de analisis de los Items de la factura nro " + ZZZNumero
            sBody = ZTexto1
            If Trim(ZTexto6) <> "" Then
                sBody = sBody + Chr$(13) + ZTexto6
            End If
            If Trim(ZTexto7) <> "" Then
                sBody = sBody + Chr$(13) + ZTexto7
            End If
            If Trim(ZTexto8) <> "" Then
                sBody = sBody + Chr$(13) + ZTexto8
            End If
            If Trim(ZTexto9) <> "" Then
                sBody = sBody + Chr$(13) + ZTexto9
            End If
            SFile = ZZArchivoEnvio
            
            
            EmailAddress = sTo
            CopiaAddress = sCC
            MSubject = sSubject
            MBody = sBody
            
            ZZSuma = 0
            For ZZZCiclo = 1 To ZZZSumaII
            
                ZZSuma = ZZSuma + 1
                MAttach = ZZEnviaPdf(ZZSuma, 2)
                ZZSuma = ZZSuma + 1
                MAttachI = ZZEnviaPdf(ZZSuma, 2)
                ZZSuma = ZZSuma + 1
                MAttachII = ZZEnviaPdf(ZZSuma, 2)
                MAttachIII = ""
                MAttachIV = ""
                MAttachVI = ""
                MAttachVII = ""
                MAttachVIII = ""
            
                SendEmail
            
            Next ZZZCiclo
            
            m$ = "Se enviaron los certicados de analisis correspondientes por email a " + sTo
            a% = MsgBox(m$, 0, "Certificados de Analisis")
            
            
            Rem For ZZZCiclo = 1 To 100
            Rem     If Trim(ZZEnviaPdf(ZZZCiclo, 2)) <> "" Then
            Rem
            Rem         AA1 = ZZEnviaPdf(ZZZCiclo, 1)
            Rem         aa2 = ZZEnviaPdf(ZZZCiclo, 2)
            Rem         aa3 = ZZEnviaPdf(ZZZCiclo, 3)
            Rem         aa4 = ZZEnviaPdf(ZZZCiclo, 4)
            Rem         aa5 = ZZEnviaPdf(ZZZCiclo, 5)
            Rem
            Rem         ZTexto1 = "Se adjuntan el Certificado de Analisis correspondientes a " + Trim(ZZEnviaPdf(ZZZCiclo, 4))
            Rem         ZTexto6 = "correspondiente a la partida " + Trim(ZZEnviaPdf(ZZZCiclo, 5)) + " de la factura numero " + Trim(ZZZNumero)
            Rem         ZTexto7 = ""
            Rem         ZTexto8 = "Surfactan S.A."
            Rem         ZTexto9 = "011-4714-4097"
            Rem
            Rem
            Rem         sTo = ZEmailFactura
            Rem         sTo = "d_esquenazi@yahoo.com"
            Rem         sCC = ""
            Rem         sBCC = ""
            Rem         sSubject = "Certificados de analisis de los Items de la factura nro " + ZZZNumero
            Rem         sBody = ZTexto1
            Rem         If Trim(ZTexto6) <> "" Then
            Rem             sBody = sBody + Chr$(13) + ZTexto6
            Rem         End If
            Rem         If Trim(ZTexto7) <> "" Then
            Rem             sBody = sBody + Chr$(13) + ZTexto7
            Rem         End If
            Rem         If Trim(ZTexto8) <> "" Then
            Rem             sBody = sBody + Chr$(13) + ZTexto8
            Rem         End If
            Rem         If Trim(ZTexto9) <> "" Then
            Rem             sBody = sBody + Chr$(13) + ZTexto9
            Rem         End If
            Rem
            Rem         EmailAddress = sTo
            Rem         CopiaAddress = sCC
            Rem         MSubject = sSubject
            Rem         MBody = sBody
            Rem
            Rem
            Rem         SFile = ZZEnviaPdf(ZZZCiclo, 2)
            Rem         MAttach = ZZEnviaPdf(ZZZCiclo, 2)
            Rem         MAttachI = ""
            Rem         MAttachII = ""
            Rem         MAttachIII = ""
            Rem         MAttachIV = ""
            Rem         MAttachVI = ""
            Rem         MAttachVII = ""
            Rem         MAttachVIII = ""
            Rem
            Rem
            Rem         m$ = "aca es justo antes de evniar el email"
            Rem         a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
            Rem
            Rem
            Rem         Rem dada
            Rem         Rem dada
            Rem         Rem dada
            Rem         Rem dada
            Rem         Rem dada
            Rem
            Rem         SendEmail
            Rem
            Rem
            Rem         m$ = "aca ya mano el email"
            Rem         a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
            Rem
            Rem
            Rem     End If
            Rem
            Rem Next ZZZCiclo
            
            
                    
            
            
            
            
            
            
        End If
        
    End If

End Sub


Public Sub SendEmail()

    Dim objOutlook As Object
    Dim objMailItem

    Dim NumOfPath As Integer, i As Integer
    Dim AtachPath As String

    On Error GoTo 10

    NumOfPath = 0
    AllPath = ""
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
        .To = EmailAddress
        .cc = CopiaAddress
        .Subject = MSubject
        .Body = MBody
        .Attachments.Add MAttach
        If MAttachI <> "" Then
            .Attachments.Add MAttachI
        End If
        If MAttachII <> "" Then
            .Attachments.Add MAttachII
        End If
        If MAttachIII > "" Then
            .Attachments.Add MAttachIII
        End If
        If MAttachIV <> "" Then
            .Attachments.Add MAttachIV
        End If
        If MAttachV <> "" Then
            .Attachments.Add MAttachV
        End If
        .Send
    End With

    Set objMailItem = Nothing
    Set objOutlook = Nothing
            
    Exit Sub

exit10:
    Exit Sub

10:
    If Err.Number = 429 Then
        MsgBox "Error on connecting with Outlook"
            Else
        MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    End If
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    AllPath = ""

    Resume exit10

End Sub
    
    
    
    
    









Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub





Sub Impresion_RemitoPrueba()

    Rem toto
    Rem toto
    Rem toto
    Rem toto

    Call Verifica_Impresora
    
    Rem Cae.Text = "30214584758415"
    Rem VtoCae.Text = "21/03/2011"
    
    Auxi1 = Str$(Val(Numero.Text) - 100000)
    Call Ceros(Auxi1, 8)
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpreRemito"
    Rem ZSql = ZSql + " Where Numero = " + "'" + Auxi1 + "'"
    spImpreRemito = ZSql
    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)


    Erase ZZVector
    ZZLugarII = 0
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
        
            Suma = Suma + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 0
            Producto = DBGrid1.Text
            
            DBGrid1.Col = 1
            Descri = DBGrid1.Text
            
            DBGrid1.Col = 3
            XPrecio = Val(Alinea("##,###.##", DBGrid1.Text))
        
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            ZZLugarII = ZZLugarII + 1
            
            ZZVector(ZZLugarII, 1) = Str$(Cantidad)
            ZZVector(ZZLugarII, 2) = Trim(Descri)
            ZZVector(ZZLugarII, 3) = Str$(XPrecio)
            ZZVector(ZZLugarII, 4) = Str$(XPrecio * Cantidad)
            ZZVector(ZZLugarII, 5) = Producto
                
        Next iRow
        
    Next a
        
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WPago1 = rstCliente!Pago1
        WPago2 = rstCliente!Pago2
        WVendedor = rstCliente!vendedor
        WProv = rstCliente!Provincia
        WRubro = rstCliente!Rubro
        WCodIva = rstCliente!Iva
        WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
        WCodIb = rstCliente!Ib
        WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
        WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
        WRazon = Trim(rstCliente!Razon)
        WDireccion = Trim(rstCliente!Direccion)
        WLocalidad = Trim(rstCliente!Localidad)
        WPostal = Trim(rstCliente!Postal)
        WCuit = Trim(rstCliente!Cuit)
        WDirentrega = ""
        ZDirEntrega(1) = rstCliente!DirEntrega
        ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        WDirentrega = ZDirEntrega(ZLugarDirEntrega)
        rstCliente.Close
    End If
        
    If Val(Numero.Text) > 100000 Then
        Auxi1 = Str$(Val(Numero.Text) - 100000)
            Else
        Auxi1 = Numero.Text
    End If
    Call Ceros(Auxi1, 8)
    ZZRenglon = 0
                                
    WWNumero = Auxi1
    WWFecha = Fecha.Text
    WWNombre = WRazon
    WWDireccion = WDireccion
    WWLocalidad = WLocalidad
    WWPedido = Pedido.Text
    WWCliente = Cliente.Text
    WWOrden = ""
    WWRemito = Remito.Text
    WWProvincia = Provincia(Val(WProv)) + " (" + WPostal + ")"
    WWCuit = WCuit
    WWDirEntrega = WDirentrega
    WWImpreIva = Iva(Val(WCodIva))
        
    For aaaa = 1 To 16
                
        ZZCantidad = Val(ZZVector(aaaa, 1))
        ZZDescripcion = ZZVector(aaaa, 2)
        ZZPrecio = Val(ZZVector(aaaa, 3))
        ZZParcial = Val(ZZVector(aaaa, 4))
        ZZProducto = ZZVector(aaaa, 5)
        
        If ZZCantidad <> 0 Then
        
            ZClase = ""
            ZIntervencion = ""
            ZNaciones = ""
            spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                ZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                ZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                ZDescriOnu = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
                ZEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                ZClase = Trim(ZClase)
                ZIntervencion = Trim(ZIntervencion)
                ZNaciones = Trim(ZNaciones)
                rstTerminado.Close
            End If
        
            If Trim(ZClase) <> "" Then
            
                ZImpre = ""
                ZImpre = "Clase:" + ZClase + " N.ONU:" + ZNaciones + " GRUPO DE EMBALAJE:" + ZEmbalaje
    
                ZZRenglon = ZZRenglon + 1
                Auxi2 = Str$(ZZRenglon)
                Call Ceros(Auxi2, 2)

                WWClave = Auxi1 + Auxi2
                WWRenglon = Str(ZZRenglon)
                WWDescripcion = Left$(ZZDescripcion, 50)
                WWDescripcionII = ""
                WWCantidad = Str$(ZZCantidad)
                
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
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito ,"
                ZSql = ZSql + "Provincia ,"
                ZSql = ZSql + "Cuit ,"
                ZSql = ZSql + "DirEntrega ,"
                ZSql = ZSql + "ImpreIva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWNumero + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWDireccion + "',"
                ZSql = ZSql + "'" + WWLocalidad + "',"
                ZSql = ZSql + "'" + WWPedido + "',"
                ZSql = ZSql + "'" + WWCliente + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWDescripcion + "',"
                ZSql = ZSql + "'" + WWDescripcionII + "',"
                ZSql = ZSql + "'" + WWCantidad + "',"
                ZSql = ZSql + "'" + WWRemito + "',"
                ZSql = ZSql + "'" + WWProvincia + "',"
                ZSql = ZSql + "'" + WWCuit + "',"
                ZSql = ZSql + "'" + WWDirEntrega + "',"
                ZSql = ZSql + "'" + WWImpreIva + "')"
                    
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
                
                Impre = Impre + 1
                
                If Trim(ZDescriOnu) <> "" Then
                
                    ZZRenglon = ZZRenglon + 1
                    Auxi2 = Str$(ZZRenglon)
                    Call Ceros(Auxi2, 2)
    
                    WWClave = Auxi1 + Auxi2
                    WWRenglon = Str(ZZRenglon)
                    WWDescripcion = ""
                    WWDescripcionII = Trim(ZDescriOnu)
                    WWCantidad = ""
                    
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
                    ZSql = ZSql + "DescriII ,"
                    ZSql = ZSql + "Cantidad ,"
                    ZSql = ZSql + "Remito ,"
                    ZSql = ZSql + "Provincia ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "DirEntrega ,"
                    ZSql = ZSql + "ImpreIva )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWNumero + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWDireccion + "',"
                    ZSql = ZSql + "'" + WWLocalidad + "',"
                    ZSql = ZSql + "'" + WWPedido + "',"
                    ZSql = ZSql + "'" + WWCliente + "',"
                    ZSql = ZSql + "'" + WWOrden + "',"
                    ZSql = ZSql + "'" + WWDescripcion + "',"
                    ZSql = ZSql + "'" + WWDescripcionII + "',"
                    ZSql = ZSql + "'" + WWCantidad + "',"
                    ZSql = ZSql + "'" + WWRemito + "',"
                    ZSql = ZSql + "'" + WWProvincia + "',"
                    ZSql = ZSql + "'" + WWCuit + "',"
                    ZSql = ZSql + "'" + WWDirEntrega + "',"
                    ZSql = ZSql + "'" + WWImpreIva + "')"
                        
                    spImpreRemito = ZSql
                    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
                        
                    Impre = Impre + 1
                    
                End If
                                    
                If ZImpre <> "" Then
                        
                    ZZRenglon = ZZRenglon + 1
                    Auxi2 = Str$(ZZRenglon)
                    Call Ceros(Auxi2, 2)
    
                    WWClave = Auxi1 + Auxi2
                    WWRenglon = Str(ZZRenglon)
                    WWDescripcion = ""
                    WWDescripcionII = ZImpre
                    WWCantidad = ""
                    
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
                    ZSql = ZSql + "DescriII ,"
                    ZSql = ZSql + "Cantidad ,"
                    ZSql = ZSql + "Remito ,"
                    ZSql = ZSql + "Provincia ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "DirEntrega ,"
                    ZSql = ZSql + "ImpreIva )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWNumero + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWDireccion + "',"
                    ZSql = ZSql + "'" + WWLocalidad + "',"
                    ZSql = ZSql + "'" + WWPedido + "',"
                    ZSql = ZSql + "'" + WWCliente + "',"
                    ZSql = ZSql + "'" + WWOrden + "',"
                    ZSql = ZSql + "'" + WWDescripcion + "',"
                    ZSql = ZSql + "'" + WWDescripcionII + "',"
                    ZSql = ZSql + "'" + WWCantidad + "',"
                    ZSql = ZSql + "'" + WWRemito + "',"
                    ZSql = ZSql + "'" + WWProvincia + "',"
                    ZSql = ZSql + "'" + WWCuit + "',"
                    ZSql = ZSql + "'" + WWDirEntrega + "',"
                    ZSql = ZSql + "'" + WWImpreIva + "')"
                        
                    spImpreRemito = ZSql
                    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Impre = Impre + 1
                    
                End If
                    
                    Else
    
                ZZRenglon = ZZRenglon + 1
                Auxi2 = Str$(ZZRenglon)
                Call Ceros(Auxi2, 2)

                WWClave = Auxi1 + Auxi2
                WWRenglon = Str(ZZRenglon)
                WWDescripcion = Left$(ZZDescripcion, 50)
                WWDescripcionII = ""
                WWCantidad = Str$(ZZCantidad)
                
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
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito ,"
                ZSql = ZSql + "Provincia ,"
                ZSql = ZSql + "Cuit ,"
                ZSql = ZSql + "DirEntrega ,"
                ZSql = ZSql + "ImpreIva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWNumero + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWDireccion + "',"
                ZSql = ZSql + "'" + WWLocalidad + "',"
                ZSql = ZSql + "'" + WWPedido + "',"
                ZSql = ZSql + "'" + WWCliente + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWDescripcion + "',"
                ZSql = ZSql + "'" + WWDescripcionII + "',"
                ZSql = ZSql + "'" + WWCantidad + "',"
                ZSql = ZSql + "'" + WWRemito + "',"
                ZSql = ZSql + "'" + WWProvincia + "',"
                ZSql = ZSql + "'" + WWCuit + "',"
                ZSql = ZSql + "'" + WWDirEntrega + "',"
                ZSql = ZSql + "'" + WWImpreIva + "')"
                    
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
            
                Impre = Impre + 1
                
            End If
            
            
            ZLote1 = XLote(aaaa, 1)
            ZCantidad1 = XLote(aaaa, 2)
            ZLote2 = XLote(aaaa, 3)
            ZCantidad2 = XLote(aaaa, 4)
            ZLote3 = XLote(aaaa, 5)
            ZCantidad3 = XLote(aaaa, 6)
            ZLote4 = XLote(aaaa, 7)
            ZCantidad4 = XLote(aaaa, 8)
            ZLote5 = XLote(aaaa, 9)
            ZCantidad5 = XLote(aaaa, 10)
            ZLote6 = XLote(aaaa, 11)
            ZCantidad6 = XLote(aaaa, 12)
            ZLote7 = XLote(aaaa, 13)
            ZCantidad7 = XLote(aaaa, 14)
            ZLote8 = XLote(aaaa, 15)
            ZCantidad8 = XLote(aaaa, 16)
            ZLote9 = XLote(aaaa, 17)
            ZCantidad9 = XLote(aaaa, 18)
            ZLote10 = XLote(aaaa, 19)
            ZCantidad10 = XLote(aaaa, 20)
            ZLote11 = XLote(aaaa, 21)
            ZCantidad11 = XLote(aaaa, 22)
            ZLote12 = XLote(aaaa, 23)
            ZCantidad12 = XLote(aaaa, 24)
                    
            If Trim(ZZProducto) <> "" Then
            
                If Left$(ZZProducto, 2) = "DY" Then
                
                    For ZZZCiclo = 1 To 12
                    
                        Select Case ZZZCiclo
                            Case 1
                                ZZZLote = ZLote1
                            Case 2
                                ZZZLote = ZLote2
                            Case 3
                                ZZZLote = ZLote3
                            Case 4
                                ZZZLote = ZLote4
                            Case 5
                                ZZZLote = ZLote5
                            Case 6
                                ZZZLote = ZLote6
                            Case 7
                                ZZZLote = ZLote7
                            Case 8
                                ZZZLote = ZLote8
                            Case 9
                                ZZZLote = ZLote9
                            Case 10
                                ZZZLote = ZLote10
                            Case 11
                                ZZZLote = ZLote11
                            Case Else
                                ZZZLote = ZLote12
                        End Select
                
                        ZZZArti = Left$(ZZProducto, 3) + Right$(ZZProducto, 7)
                        XParam = "'" + ZZZLote + "','" _
                                     + ZZZArti + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZZPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            rstLaudo.Close
                            
                            If Trim(ZZZPartiOri) <> "" Then
                            
                                Select Case ZZZCiclo
                                    Case 1
                                        ZLote1 = ZZZPartiOri
                                    Case 2
                                        ZLote2 = ZZZPartiOri
                                    Case 3
                                        ZLote3 = ZZZPartiOri
                                    Case 4
                                        ZLote4 = ZZZPartiOri
                                    Case 5
                                        ZLote5 = ZZZPartiOri
                                    Case 6
                                        ZLote6 = ZZZPartiOri
                                    Case 7
                                        ZLote7 = ZZZPartiOri
                                    Case 8
                                        ZLote8 = ZZZPartiOri
                                    Case 9
                                        ZLote9 = ZZZPartiOri
                                    Case 10
                                        ZLote10 = ZZZPartiOri
                                    Case 11
                                        ZLote11 = ZZZPartiOri
                                    Case Else
                                        ZLote12 = ZZZPartiOri
                                End Select
                                
                            End If
                        End If
                        
                    Next ZZZCiclo
                    
                        Else
                        
                    For ZZZCiclo = 1 To 12
                    
                        Select Case ZZZCiclo
                            Case 1
                                ZZZLote = ZLote1
                            Case 2
                                ZZZLote = ZLote2
                            Case 3
                                ZZZLote = ZLote3
                            Case 4
                                ZZZLote = ZLote4
                            Case 5
                                ZZZLote = ZLote5
                            Case 6
                                ZZZLote = ZLote6
                            Case 7
                                ZZZLote = ZLote7
                            Case 8
                                ZZZLote = ZLote8
                            Case 9
                                ZZZLote = ZLote9
                            Case 10
                                ZZZLote = ZLote10
                            Case 11
                                ZZZLote = ZLote11
                            Case Else
                                ZZZLote = ZLote12
                        End Select
                        
                        If Val(ZZZLote) <> 0 Then
                            
                            XParam = "'" + ZZZLote + "','" _
                                         + ZZProducto + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                            
                                ZZZPartiOri = IIf(IsNull(rstHoja!LoteColorante), "", rstHoja!LoteColorante)
                                
                                rstHoja.Close
                                
                                If Trim(ZZZPartiOri) <> "" Then
                                
                                    Select Case ZZZCiclo
                                        Case 1
                                            ZLote1 = ZZZPartiOri
                                        Case 2
                                            ZLote2 = ZZZPartiOri
                                        Case 3
                                            ZLote3 = ZZZPartiOri
                                        Case 4
                                            ZLote4 = ZZZPartiOri
                                        Case 5
                                            ZLote5 = ZZZPartiOri
                                        Case 6
                                            ZLote6 = ZZZPartiOri
                                        Case 7
                                            ZLote7 = ZZZPartiOri
                                        Case 8
                                            ZLote8 = ZZZPartiOri
                                        Case 9
                                            ZLote9 = ZZZPartiOri
                                        Case 10
                                            ZLote10 = ZZZPartiOri
                                        Case 11
                                            ZLote11 = ZZZPartiOri
                                        Case Else
                                            ZLote12 = ZZZPartiOri
                                    End Select
                                    
                                End If
                            End If
                            
                        End If
                        
                    Next ZZZCiclo
                    
                End If
            
            End If
            
            
            ZEnv1 = XLote(aaaa, 31)
            ZCantiEnv1 = XLote(aaaa, 32)
            ZEnv2 = XLote(aaaa, 33)
            ZCantiEnv2 = XLote(aaaa, 34)
            ZEnv3 = XLote(aaaa, 35)
            ZCantiEnv3 = XLote(aaaa, 36)
            ZEnv4 = XLote(aaaa, 37)
            ZCantiEnv4 = XLote(aaaa, 38)
            ZEnv5 = XLote(aaaa, 39)
            ZCantiEnv5 = XLote(aaaa, 40)
            ZEnv6 = XLote(aaaa, 41)
            ZCantiEnv6 = XLote(aaaa, 42)
            ZEnv7 = XLote(aaaa, 43)
            ZCantiEnv7 = XLote(aaaa, 44)
            ZEnv8 = XLote(aaaa, 45)
            ZCantiEnv8 = XLote(aaaa, 46)
            ZEnv9 = XLote(aaaa, 47)
            ZCantiEnv9 = XLote(aaaa, 48)
            ZEnv10 = XLote(aaaa, 49)
            ZCantiEnv10 = XLote(aaaa, 50)
            ZEnv11 = XLote(aaaa, 51)
            ZCantiEnv11 = XLote(aaaa, 52)
            ZEnv12 = XLote(aaaa, 53)
            ZCantiEnv12 = XLote(aaaa, 54)
            
            ZDescri1 = ""
            ZDescri2 = ""
            ZDescri3 = ""
            ZDescri4 = ""
            ZDescri5 = ""
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv1 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri1 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv2 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri2 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv3 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri3 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv4 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri4 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv5 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri5 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv6 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri6 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv7 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri7 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv8 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri8 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv9 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri9 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv10 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri10 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv11 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri11 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            spEnvases = "ConsultaEnvases " + "'" + ZEnv12 + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri12 = Left$(rstEnvases!Abreviatura, 8)
                rstEnvases.Close
            End If
            
            ZZImpreRenglon1 = ""
            ZZImpreRenglon2 = ""
            ZZImpreRenglon3 = ""
            
            
            
            If Val(ZCantidad1) <> 0 Then
                ZZImpreRenglon1 = ZZImpreRenglon1 + Alinea("#####", ZCantidad1) + " Kg Lote:" + Left$(ZLote1, 8) + " "
                If Val(ZCantiEnv1) <> 0 Then
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Alinea("##", ZCantiEnv1) + " x " + ZDescri1
                        Else
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad2) <> 0 Then
                ZZImpreRenglon1 = ZZImpreRenglon1 + " | " + Alinea("#####", ZCantidad2) + " Kg Lote:" + Left$(ZLote2, 8) + " "
                If Val(ZCantiEnv2) <> 0 Then
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Alinea("##", ZCantiEnv2) + " x " + ZDescri2
                        Else
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad3) <> 0 Then
                ZZImpreRenglon1 = ZZImpreRenglon1 + " | " + Alinea("#####", ZCantidad3) + " Kg Lote:" + Left$(ZLote3, 8) + " "
                If Val(ZCantiEnv3) <> 0 Then
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Alinea("##", ZCantiEnv3) + " x " + ZDescri3
                        Else
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad4) <> 0 Then
                ZZImpreRenglon1 = ZZImpreRenglon1 + " | " + Alinea("#####", ZCantidad4) + " Kg Lote:" + Left$(ZLote4, 8) + " "
                If Val(ZCantiEnv4) <> 0 Then
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Alinea("##", ZCantiEnv4) + " x " + ZDescri4
                        Else
                    ZZImpreRenglon1 = ZZImpreRenglon1 + Space$(13)
                End If
            End If
            
            
            
            
            
            
            
            
            
            If Val(ZCantidad5) <> 0 Then
                ZZImpreRenglon2 = ZZImpreRenglon2 + Alinea("#####", ZCantidad5) + " Kg Lote:" + Left$(ZLote1, 8) + " "
                If Val(ZCantiEnv5) <> 0 Then
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Alinea("##", ZCantiEnv5) + " x " + ZDescri1
                        Else
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad6) <> 0 Then
                ZZImpreRenglon2 = ZZImpreRenglon2 + " | " + Alinea("#####", ZCantidad6) + " Kg Lote:" + Left$(ZLote6, 8) + " "
                If Val(ZCantiEnv6) <> 0 Then
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Alinea("##", ZCantiEnv6) + " x " + ZDescri6
                        Else
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad7) <> 0 Then
                ZZImpreRenglon2 = ZZImpreRenglon2 + " | " + Alinea("#####", ZCantidad7) + " Kg Lote:" + Left$(ZLote7, 8) + " "
                If Val(ZCantiEnv7) <> 0 Then
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Alinea("##", ZCantiEnv7) + " x " + ZDescri7
                        Else
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad8) <> 0 Then
                ZZImpreRenglon2 = ZZImpreRenglon2 + " | " + Alinea("#####", ZCantidad8) + " Kg Lote:" + Left$(ZLote8, 8) + " "
                If Val(ZCantiEnv8) <> 0 Then
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Alinea("##", ZCantiEnv8) + " x " + ZDescri8
                        Else
                    ZZImpreRenglon2 = ZZImpreRenglon2 + Space$(13)
                End If
            End If
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            If Val(ZCantidad9) <> 0 Then
                ZZImpreRenglon3 = ZZImpreRenglon3 + Alinea("#####", ZCantidad9) + " Kg Lote:" + Left$(ZLote9, 8) + " "
                If Val(ZCantiEnv9) <> 0 Then
                    ZZImpreRenglon3 = ZZImpreRenglon3 + Alinea("##", ZCantiEnv9) + " x " + ZDescri9
                        Else
                    ZZImpreRenglon3 = ZZImpreRenglon3 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad10) <> 0 Then
                ZZImpreRenglon3 = ZZImpreRenglon3 + " | " + Alinea("#####", ZCantidad10) + " Kg Lote:" + Left$(ZLote10, 8) + " "
                If Val(ZCantiEnv10) <> 0 Then
                    ZZImpreRenglon3 = ZZImpreRenglon3 + Alinea("##", ZCantiEnv10) + " x " + ZDescri10
                        Else
                    ZZImpreRenglon3 = ZZImpreRenglon3 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad11) <> 0 Then
                ZZImpreRenglon3 = ZZImpreRenglon3 + " | " + Alinea("#####", ZCantidad11) + " Kg Lote:" + Left$(ZLote11, 8) + " "
                If Val(ZCantiEnv11) <> 0 Then
                    ZZImpreRenglon4 = ZZImpreRenglon3 + Alinea("##", ZCantiEnv11) + " x " + ZDescri11
                        Else
                    ZZImpreRenglon4 = ZZImpreRenglon3 + Space$(13)
                End If
            End If
            
            If Val(ZCantidad12) <> 0 Then
                ZZImpreRenglon3 = ZZImpreRenglon3 + " | " + Alinea("#####", ZCantidad12) + " Kg Lote:" + Left$(ZLote12, 8) + " "
                If Val(ZCantiEnv12) <> 0 Then
                    ZZImpreRenglon5 = ZZImpreRenglon3 + Alinea("##", ZCantiEnv12) + " x " + ZDescri12
                        Else
                    ZZImpreRenglon5 = ZZImpreRenglon3 + Space$(13)
                End If
            End If
            
            
            
            
            
            
            
            
            
            
            
            If Trim(ZZImpreRenglon1) <> "" Then
            
                ZZRenglon = ZZRenglon + 1
                Auxi2 = Str$(ZZRenglon)
                Call Ceros(Auxi2, 2)
    
                WWClave = Auxi1 + Auxi2
                WWRenglon = Str(ZZRenglon)
                WWDescripcion = ""
                WWDescripcionII = ZZImpreRenglon1
                WWCantidad = ""
                
                aa = Len(WWDescripcionII)
                
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
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito ,"
                ZSql = ZSql + "Provincia ,"
                ZSql = ZSql + "Cuit ,"
                ZSql = ZSql + "DirEntrega ,"
                ZSql = ZSql + "ImpreIva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWNumero + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWDireccion + "',"
                ZSql = ZSql + "'" + WWLocalidad + "',"
                ZSql = ZSql + "'" + WWPedido + "',"
                ZSql = ZSql + "'" + WWCliente + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWDescripcion + "',"
                ZSql = ZSql + "'" + WWDescripcionII + "',"
                ZSql = ZSql + "'" + WWCantidad + "',"
                ZSql = ZSql + "'" + WWRemito + "',"
                ZSql = ZSql + "'" + WWProvincia + "',"
                ZSql = ZSql + "'" + WWCuit + "',"
                ZSql = ZSql + "'" + WWDirEntrega + "',"
                ZSql = ZSql + "'" + WWImpreIva + "')"
                    
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            
            
            
            
            
            
            
            
            If Trim(ZZImpreRenglon2) <> "" Then
            
                ZZRenglon = ZZRenglon + 1
                Auxi2 = Str$(ZZRenglon)
                Call Ceros(Auxi2, 2)
    
                WWClave = Auxi1 + Auxi2
                WWRenglon = Str(ZZRenglon)
                WWDescripcion = ""
                WWDescripcionII = ZZImpreRenglon2
                WWCantidad = ""
                
                aa = Len(WWDescripcionII)
                
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
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito ,"
                ZSql = ZSql + "Provincia ,"
                ZSql = ZSql + "Cuit ,"
                ZSql = ZSql + "DirEntrega ,"
                ZSql = ZSql + "ImpreIva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWNumero + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWDireccion + "',"
                ZSql = ZSql + "'" + WWLocalidad + "',"
                ZSql = ZSql + "'" + WWPedido + "',"
                ZSql = ZSql + "'" + WWCliente + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWDescripcion + "',"
                ZSql = ZSql + "'" + WWDescripcionII + "',"
                ZSql = ZSql + "'" + WWCantidad + "',"
                ZSql = ZSql + "'" + WWRemito + "',"
                ZSql = ZSql + "'" + WWProvincia + "',"
                ZSql = ZSql + "'" + WWCuit + "',"
                ZSql = ZSql + "'" + WWDirEntrega + "',"
                ZSql = ZSql + "'" + WWImpreIva + "')"
                    
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            
            
            
            
            
            If Trim(ZZImpreRenglon3) <> "" Then
            
                ZZRenglon = ZZRenglon + 1
                Auxi2 = Str$(ZZRenglon)
                Call Ceros(Auxi2, 2)
    
                WWClave = Auxi1 + Auxi2
                WWRenglon = Str(ZZRenglon)
                WWDescripcion = ""
                WWDescripcionII = ZZImpreRenglon3
                WWCantidad = ""
                
                aa = Len(WWDescripcionII)
                
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
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Remito ,"
                ZSql = ZSql + "Provincia ,"
                ZSql = ZSql + "Cuit ,"
                ZSql = ZSql + "DirEntrega ,"
                ZSql = ZSql + "ImpreIva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWNumero + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWDireccion + "',"
                ZSql = ZSql + "'" + WWLocalidad + "',"
                ZSql = ZSql + "'" + WWPedido + "',"
                ZSql = ZSql + "'" + WWCliente + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWDescripcion + "',"
                ZSql = ZSql + "'" + WWDescripcionII + "',"
                ZSql = ZSql + "'" + WWCantidad + "',"
                ZSql = ZSql + "'" + WWRemito + "',"
                ZSql = ZSql + "'" + WWProvincia + "',"
                ZSql = ZSql + "'" + WWCuit + "',"
                ZSql = ZSql + "'" + WWDirEntrega + "',"
                ZSql = ZSql + "'" + WWImpreIva + "')"
                    
                spImpreRemito = ZSql
                Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            
            
        End If
            
    Next aaaa
    
    For aaaa = ZZRenglon + 1 To 16
    
        Auxi2 = Str$(aaaa)
        Call Ceros(Auxi2, 2)

        WWClave = Auxi1 + Auxi2
        WWRenglon = Str(aaaa)
        WWDescripcion = ""
        WWDescripcionII = ""
        WWCantidad = ""
        
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
        ZSql = ZSql + "DescriII ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "DirEntrega ,"
        ZSql = ZSql + "ImpreIva )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWNumero + "',"
        ZSql = ZSql + "'" + WWRenglon + "',"
        ZSql = ZSql + "'" + WWFecha + "',"
        ZSql = ZSql + "'" + WWNombre + "',"
        ZSql = ZSql + "'" + WWDireccion + "',"
        ZSql = ZSql + "'" + WWLocalidad + "',"
        ZSql = ZSql + "'" + WWPedido + "',"
        ZSql = ZSql + "'" + WWCliente + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWDescripcion + "',"
        ZSql = ZSql + "'" + WWDescripcionII + "',"
        ZSql = ZSql + "'" + WWCantidad + "',"
        ZSql = ZSql + "'" + WWRemito + "',"
        ZSql = ZSql + "'" + WWProvincia + "',"
        ZSql = ZSql + "'" + WWCuit + "',"
        ZSql = ZSql + "'" + WWDirEntrega + "',"
        ZSql = ZSql + "'" + WWImpreIva + "')"
            
        spImpreRemito = ZSql
        Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
    
    Next aaaa
    
    
    
    
    
    
    
    
    
    
    Erase ZImpreStk
    
    For XDa = 1 To 1
        For DA = 1 To 9
        
            If Val(Stk(DA, 4)) <> 0 Then
                ZImpreStk(DA, XDa) = Stk(DA, XDa)
                spEnvases = "ConsultaEnvases " + "'" + Stk(DA, XDa) + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvases.RecordCount > 0 Then
                    ZImpreStk(DA, XDa) = Left$(rstEnvases!Abreviatura, 10)
                    rstEnvases.Close
                End If
            End If

        Next DA

    Next XDa
    
    
    For XDa = 2 To 4
        For DA = 1 To 9
            If Val(Stk(DA, 4)) <> 0 Then
                ZImpreStk(DA, XDa) = Stk(DA, XDa)
            End If
        Next DA
    Next XDa
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImpreRemito SET "
    ZSql = ZSql + " Impre1 = " + "'" + ZImpreStk(1, 1) + "',"
    ZSql = ZSql + " Impre2 = " + "'" + ZImpreStk(2, 1) + "',"
    ZSql = ZSql + " Impre3 = " + "'" + ZImpreStk(3, 1) + "',"
    ZSql = ZSql + " Impre4 = " + "'" + ZImpreStk(4, 1) + "',"
    ZSql = ZSql + " Impre5 = " + "'" + ZImpreStk(5, 1) + "',"
    ZSql = ZSql + " Impre6 = " + "'" + ZImpreStk(6, 1) + "',"
    ZSql = ZSql + " Impre7 = " + "'" + ZImpreStk(7, 1) + "',"
    ZSql = ZSql + " Impre8 = " + "'" + ZImpreStk(8, 1) + "',"
    ZSql = ZSql + " Impre9 = " + "'" + ZImpreStk(9, 1) + "',"
    ZSql = ZSql + " Impre11 = " + "'" + ZImpreStk(1, 2) + "',"
    ZSql = ZSql + " Impre12 = " + "'" + ZImpreStk(2, 2) + "',"
    ZSql = ZSql + " Impre13 = " + "'" + ZImpreStk(3, 2) + "',"
    ZSql = ZSql + " Impre14 = " + "'" + ZImpreStk(4, 2) + "',"
    ZSql = ZSql + " Impre15 = " + "'" + ZImpreStk(5, 2) + "',"
    ZSql = ZSql + " Impre16 = " + "'" + ZImpreStk(6, 2) + "',"
    ZSql = ZSql + " Impre17 = " + "'" + ZImpreStk(7, 2) + "',"
    ZSql = ZSql + " Impre18 = " + "'" + ZImpreStk(8, 2) + "',"
    ZSql = ZSql + " Impre19 = " + "'" + ZImpreStk(9, 2) + "',"
    ZSql = ZSql + " Impre21 = " + "'" + ZImpreStk(1, 3) + "',"
    ZSql = ZSql + " Impre22 = " + "'" + ZImpreStk(2, 3) + "',"
    ZSql = ZSql + " Impre23 = " + "'" + ZImpreStk(3, 3) + "',"
    ZSql = ZSql + " Impre24 = " + "'" + ZImpreStk(4, 3) + "',"
    ZSql = ZSql + " Impre25 = " + "'" + ZImpreStk(5, 3) + "',"
    ZSql = ZSql + " Impre26 = " + "'" + ZImpreStk(6, 3) + "',"
    ZSql = ZSql + " Impre27 = " + "'" + ZImpreStk(7, 3) + "',"
    ZSql = ZSql + " Impre28 = " + "'" + ZImpreStk(8, 3) + "',"
    ZSql = ZSql + " Impre29 = " + "'" + ZImpreStk(9, 3) + "',"
    ZSql = ZSql + " Impre31 = " + "'" + ZImpreStk(1, 4) + "',"
    ZSql = ZSql + " Impre32 = " + "'" + ZImpreStk(2, 4) + "',"
    ZSql = ZSql + " Impre33 = " + "'" + ZImpreStk(3, 4) + "',"
    ZSql = ZSql + " Impre34 = " + "'" + ZImpreStk(4, 4) + "',"
    ZSql = ZSql + " Impre35 = " + "'" + ZImpreStk(5, 4) + "',"
    ZSql = ZSql + " Impre36 = " + "'" + ZImpreStk(6, 4) + "',"
    ZSql = ZSql + " Impre37 = " + "'" + ZImpreStk(7, 4) + "',"
    ZSql = ZSql + " Impre38 = " + "'" + ZImpreStk(8, 4) + "',"
    ZSql = ZSql + " Impre39 = " + "'" + ZImpreStk(9, 4) + "'"
                 
    spImpreRemito = ZSql
    Set rstImpreRemito = db.OpenRecordset(spImpreRemito, dbOpenSnapshot, dbSQLPassThrough)
    
            
    Listado.WindowTitle = "Remito"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.CopiesToPrinter = 2
    
    If Wempresa = "0008" Then
        Rem Listado.ReportFileName = "ImpreRemitoI.rpt"
        Listado.ReportFileName = "ImpreRemitopelli.rpt"
            Else
        Listado.ReportFileName = "ImpreRemitoI.rpt"
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = ""
                       
    Listado.Destination = 1
    Rem Listado.Destination = 0
      
    Listado.Connect = Connect()
    Listado.Action = 1
    
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






Private Sub Verifica_Certificado()

    ZAprueba = "S"
    
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
    
    Rem ZZRequiereCertificado = 1
    
    If ZZRequiereCertificado = 1 Then
    
        WRenglonRemito = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                Suma = Suma + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text

                DBGrid1.Col = 1
                ZZDescriArticulo = DBGrid1.Text
                ZZDescriArticuloPDF = ""
                ZZZHasta = Len(ZZDescriArticulo)
                For ZZZCiclo = 1 To ZZZHasta
                    If Mid$(ZZDescriArticulo, ZZZCiclo, 1) <> Space(1) Then
                        ZZDescriArticuloPDF = ZZDescriArticuloPDF + Mid$(ZZDescriArticulo, ZZZCiclo, 1)
                    End If
                Next ZZZCiclo
                
                
                WLote1 = XLote(Suma, 1)
                WLote2 = XLote(Suma, 3)
                Wlote3 = XLote(Suma, 5)
                WLote4 = XLote(Suma, 7)
                WLote5 = XLote(Suma, 9)
                WLote6 = XLote(Suma, 11)
                WLote7 = XLote(Suma, 13)
                WLote8 = XLote(Suma, 15)
                WLote9 = XLote(Suma, 17)
                WLote10 = XLote(Suma, 19)
                WLote11 = XLote(Suma, 21)
                WLote12 = XLote(Suma, 23)
                
                                
                
                If Trim(Articulo) <> "" Then
                        
                    Rem
                    Rem certificado de analisis
                    Rem
        
                    For ZZCiclo = 1 To 12
                        
                        Select Case ZZCiclo
                            Case 1
                                WWLote = WLote1
                            Case 2
                                WWLote = WLote2
                            Case 3
                                WWLote = Wlote3
                            Case 4
                                WWLote = WLote4
                            Case 5
                                WWLote = WLote5
                            Case 6
                                WWLote = WLote6
                            Case 7
                                WWLote = WLote7
                            Case 8
                                WWLote = WLote8
                            Case 9
                                WWLote = WLote9
                            Case 10
                                WWLote = WLote10
                            Case 11
                                WWLote = WLote11
                            Case Else
                                WWLote = WLote12
                        End Select
                        
                        If Val(WWLote) <> 0 Then
                
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
                        
                                Rem If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                                If XTipoPro = "CO" Then
                                
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
                                    ZLote = WWLote
                                    ZCliente = Cliente.Text
                                        
                                    ZZEntra = "N"
                                    Erase ZOpcion
                                    
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
                                    
                                    Call Conecta_Empresa
                                        
                                    If ZZEntra = "S" Then
                                        If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                            ZZEntra = "N"
                                        End If
                                    End If
                                    
                                    If ZZEntra = "N" Then
                                        m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                                        Aaa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                        ZAprueba = "N"
                                        WEstado = "N"
                                    End If
                                                    
                                End If
                            
                                    Else
                                    
                                If Left$(Articulo, 2) = "DY" Then
                                
                                    ZZPartiOri = ""
                                    ZProductoDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Laudo"
                                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZProductoDy + "'"
                                    ZSql = ZSql + " and Laudo.Lote = " + "'" + WWLote + "'"
                                    spLaudo = ZSql
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstLaudo.RecordCount > 0 Then
                                        ZZPartiOri = Trim(rstLaudo!PartiOri)
                                        ZZCambia = "S"
                                        rstLaudo.Close
                                    End If
                                        
                                    ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                                    ZZEstado = Dir(ZZRuta)
                                    ZZEstado = Trim(ZZEstado)
                                    If ZZEstado = "" Then
                                        m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                        ssa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                        ZAprueba = "N"
                                        WEstado = "N"
                                    End If
                                    
                                End If
                                
                            End If
                        End If
                            
                    Next ZZCiclo
                End If
        
            Next iRow
        Next a
    End If

End Sub

