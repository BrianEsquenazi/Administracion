VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDevol21 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Devolucion de Mercaderia"
   ClientHeight    =   8145
   ClientLeft      =   345
   ClientTop       =   555
   ClientWidth     =   11400
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8145
   ScaleWidth      =   11400
   Visible         =   0   'False
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   41
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   36
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   35
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   600
      TabIndex        =   29
      Top             =   5160
      Width           =   9735
      Begin VB.TextBox WEntrada 
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
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WTipopro 
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
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox WLote 
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
         Left            =   8520
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
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
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   30
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WDescripcion 
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
         Left            =   2040
         TabIndex        =   34
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label WPrecio 
         Alignment       =   1  'Right Justify
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
         Left            =   5640
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Paridad 
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
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   26
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   6360
      TabIndex        =   17
      Top             =   5880
      Width           =   2655
      Begin VB.Label Label13 
         Caption         =   "Ing.Brutos"
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
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.Label ImpoIb 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   45
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Interes"
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
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Descuento"
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
         TabIndex        =   39
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
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
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
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
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
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
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6480
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   615
      Left            =   10080
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   1800
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.TextBox Cliente 
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
      Left            =   2160
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
      Left            =   2160
      MaxLength       =   8
      TabIndex        =   0
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
      Height          =   570
      Left            =   10080
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
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
      Height          =   570
      Left            =   10080
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
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
      Height          =   570
      Left            =   10080
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "prgdevol21.frx":0000
      TabIndex        =   3
      Top             =   1200
      Width           =   9855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "prgdevol21.frx":09DE
      Left            =   120
      List            =   "prgdevol21.frx":09E5
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
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
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
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
      TabIndex        =   13
      Top             =   840
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
      Left            =   3480
      TabIndex        =   12
      Top             =   480
      Width           =   4455
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
      Caption         =   "Numero de Devolucion"
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
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "PrgDevol21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
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
Private WCodIva As String
Private parcial As Double
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WImporte As Double
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
Private Auxiliar(100, 10) As String
Dim IngreVector(1000, 3) As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
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
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPreciosMp  As Recordset
Dim spPreciosMp As String
Dim XParam As String
Dim Compara As Double
Private WCodIb As Integer
Private WImpoIb As Double

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WPlazo1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPlazo2 = rstPago!Plazo
    End If
    
    Call Calcula_vencimiento(WFecha, WPlazo2, WVencimiento1)

End Sub


Private Sub Borra_Click()

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
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLinea.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    
    WArticulo.SetFocus

End Sub

Private Sub Command1_Click()
    Call Impresion
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
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
            
        Case 1
            Erase IngreVector
            EntraVector = 0
    
            spPreciosMp = "ListaPreciosClienteMp " + "'" + Cliente.Text + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
    
                With rstPreciosMp
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Cliente.Text = rstPreciosMp!Cliente Then
                                WArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                                EntraVector = EntraVector + 1
                                IngreVector(EntraVector, 1) = WArticulo
                                IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                                IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPreciosMp.Close
            End If
    
            For CicloVector = 1 To EntraVector
        
                WTerminado = IngreVector(CicloVector, 1)
                WCliente = IngreVector(CicloVector, 2)
                WArti = IngreVector(CicloVector, 3)
                WDescripcion = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
                IngresaItem = WTerminado + "  " + WDescripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = WCliente + WArti
                WIndice.AddItem IngresaItem
            
            Next CicloVector
        
            spPrecios = "ListaPrecios"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                With rstPrecios
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Cliente.Text = rstPrecios!Cliente Then
                                If rstPrecios!Precio <> "" Then
                                    IngresaItem = rstPrecios!Terminado + "   " + rstPrecios!Descripcion
                                        Else
                                    IngresaItem = rstPrecios!Terminado + "   " + rstPrecios!Descripcion
                                End If
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPrecios.Close
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
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = DBGrid1.Text
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WPrecio.Caption = DBGrid1.Text
    
    DBGrid1.Col = 4
    If Val(DBGrid1.Text) <> 0 Then
        WEntrada.Text = DBGrid1.Text
            Else
        WEntrada.Text = ""
    End If
    
    DBGrid1.Col = 5
    WTipopro.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WLote.Text = DBGrid1.Text
    
    WArticulo.SetFocus
    
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
                    
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            
            DBGrid1.Col = 2
            Cantidad = Val(DBGrid1.Text)
                    
            If Cantidad <> 0 Then
                WNeto = WNeto + (Cantidad * Precio)
            End If
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    Rem If Val(Paridad.Text) <> 0 Then
    Rem    WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    WImpoDto = 0
    WImpoInteres = 0
    
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
    
    Select Case WCodIb
        Case 0, 1
            Select Case Val(WCodIva)
                Case 1
                    WImpoIb = WNeto * 0.02
                Case 2, 4, 5, 6
                    WImpoIb = WNeto * 0.025
                Case Else
                    WImpoIb = 0
            End Select
            Call Redondeo(WImpoIb)
        Case Else
            WImpoIb = 0
    End Select
    Compara = WNeto * Val(Paridad.Text)
    Call Redondeo(Compara)
    If Compara < 100 Then
        WImpoIb = 0
    End If
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.19
            WIva2 = WNeto * 0.095
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Case Else
            WIva1 = WNeto * 0.19
            Call Redondeo(WIva1)
    End Select
            
    If WImpoIb <> 0 Then
        Call Convierte1_datos(Str$(WImpoIb), Auxi)
        ImpoIb.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIb.Caption = "0.00"
    End If
    
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
    
    WTotal = WNeto + WIva1 + WIva2 + WImpoIb
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
    PrgDevol21.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

        Cliente.Text = UCase(Cliente.Text)

        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        DBGrid1.Col = 0
        DBGrid1.Text = ""

        Call Calcula_Click
        
        Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
        Rem     WImporte = WNeto
        Rem    WNeto = WNeto / 1.19
        Rem    Call Redondeo(WNeto)
        Rem    WIva1 = WImporte - WNeto
        Rem    WIva2 = 0
        Rem End If
        
        WTipo = "02"
        WNumero = Numero.Text
        WRenglon = "01"
        WCliente = Cliente.Text
        WFecha = Fecha.Text
        WEstado = "0"
        Rem Wvencimiento = Wvencimiento
        Rem WVencimiento1 = WVencimiento1
        Call Convierte_datos(Str$(Total), Auxi)
        XTotal = Str$(WTotal * -1 * Val(Paridad.Text))
        XTotalUs = Str$(WTotal * -1)
        XSaldo = Str$(WTotal * -1 * Val(Paridad.Text))
        XSaldoUs = Str$(WTotal * -1)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WImpre = "DV"
        XNet = Str$(WNeto * -1 * Val(Paridad.Text))
        XIva1 = Str$(WIva1 * -1 * Val(Paridad.Text))
        XIva2 = Str$(WIva2 * -1 * Val(Paridad.Text))
        XImpoIb = Str$(WImpoIb * -1 * Val(Paridad.Text))
        XSeguro = ""
        XFlete = ""
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
        WClave = "02" + Auxi + "01"
        XEmpresa = "1"
        WDate = Date$
        
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
                Articulo = UCase(DBGrid1.Text)
                WTipoProDy = Left$(Articulo, 2)
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 2
                Entrada = Val(DBGrid1.Text)
                
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 5
                Tipopro = DBGrid1.Text
                
                DBGrid1.Col = 6
                Lote = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    WArti = Tipopro + Mid$(Articulo, 3, 10)
                    If WTipoProDy <> "DY" Then
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                            Else
                        WLinea = 16
                    End If
                 
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                    
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                    WTipo = "02"
                    WNumero = Numero.Text
                    XRenglon = Str$(Renglon)
                    WArticulo = Tipopro + Mid$(Articulo, 3, 10)
                    XCantidad = Str$(Cantidad)
                    XPrecio = Str$(Precio * Val(Paridad.Text))
                    XPrecioUs = Str$(Precio)
                    XImporte = Str$(Precio * Cantidad * -1 * Val(Paridad.Text))
                    XImporteUs = Str$(Precio * Cantidad * -1)
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = WCosto1
                    XCosto1 = WCosto
                    WCoeficiente = ""
                    WPedido = ""
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = Tipopro + Mid$(Articulo, 3, 6)
                    WRemito = ""
                    WClave = "02" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    XMarca = ""
                    WLote1 = Str$(Lote)
                    WCanti1 = Str$(Cantidad)
                    WLote2 = "0"
                    WCanti2 = "0"
                    Wlote3 = "0"
                    WCanti3 = "0"
                    WLote4 = "0"
                    WCanti4 = "0"
                    WLote5 = "0"
                    WCanti5 = "0"
                    WEntrada = Str$(Entrada)
                    WTipopro = Tipopro
                    WHoja = ""
                    If WTipoProDy = "DY" Then
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
                        + WEntrada + "','" _
                        + WTipopro + "','" + WHoja + "','" _
                        + XTipoproDy + "','" + XArticuloDy + "'"
                    
                    spEstadistica = "AltaEstadisticaDev " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Auxiliar(Renglon, 1) = Articulo
                    Auxiliar(Renglon, 2) = Cantidad
                    Auxiliar(Renglon, 3) = Precio
                    Auxiliar(Renglon, 4) = WRenglon
                    Auxiliar(Renglon, 5) = Lote
                    Auxiliar(Renglon, 6) = Tipopro
                    Auxiliar(Renglon, 7) = XTipoproDy
                    Auxiliar(Renglon, 8) = XArticuloDy
                    
                    Dife = Entrada - Cantidad
                    
                    If Dife > 0 Then
                    
                        WArti = "PT-99999-999"
                        spTerminado = "ConsultaTerminado " + "'" + WArti + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                                Else
                            WLinea = 50
                        End If
                    
                        Renglon = Renglon + 1
                        Auxi = Str$(Renglon)
                        Call Ceros(Auxi, 2)
                    
                        Auxi1 = Str$(Numero.Text)
                        Call Ceros(Auxi1, 8)
                        WTipo = "02"
                        WNumero = Numero.Text
                        XRenglon = Str$(Renglon)
                        WArticulo = WArti
                        XCantidad = Str$(Dife)
                        XPrecio = Str$(Precio * Val(Paridad.Text))
                        XPrecioUs = Str$(Precio)
                        XImporte = Str$(Precio * Dife * -1 * Val(Paridad.Text))
                        XImporteUs = Str$(Precio * Dife * -1)
                        WCliente = Cliente.Text
                        WParidad = Paridad.Text
                        XVendedor = Str$(WVendedor)
                        XRubro = Str$(WRubro)
                        XLinea = Str$(WLinea)
                        XCosto2 = ""
                        XCosto1 = ""
                        WCoeficiente = ""
                        WPedido = ""
                        WFecha = Fecha.Text
                        WImporte1 = ""
                        WImporte2 = ""
                        WImporte3 = ""
                        WImporte4 = ""
                        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        XArticulo = "PT-99999"
                        WRemito = ""
                        WClave = "02" + Auxi1 + Auxi
                        WDate = Date$
                        XCanti = ""
                        XImpo = ""
                        XImpoUs = ""
                        XMarca = ""
                        WLote1 = ""
                        WCanti1 = ""
                        WLote2 = "0"
                        WCanti2 = "0"
                        Wlote3 = "0"
                        WCanti3 = "0"
                        WLote4 = "0"
                        WCanti4 = "0"
                        WLote5 = "0"
                        WCanti5 = "0"
                        WEntrada = ""
                        WTipopro = ""
                        WHoja = ""
                        XTipoproDy = "T"
                        XArticuloDy = "  -   -   "
                    
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
                            + WEntrada + "','" _
                            + WTipopro + "','" + WHoja + "','" _
                            + XTipoproDy + "','" + XArticuloDy + "'"
                        
                        spEstadistica = "AltaEstadisticaDev " + XParam
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                        
                End If
                                        
            Next iRow
            
        Next a
        
        For Da = 1 To Renglon
        
            If Val(Auxiliar(Da, 2)) <> 0 Then
        
            Articulo = Auxiliar(Da, 1)
            Cantidad = Auxiliar(Da, 2)
            Precio = Auxiliar(Da, 3)
            WRenglon = Auxiliar(Da, 4)
            Lote = Auxiliar(Da, 5)
            Tipopro = Auxiliar(Da, 6)
            XTipoproDy = Auxiliar(Da, 7)
            XArticuloDy = Auxiliar(Da, 8)
            
            Select Case Tipopro
                Case "NK", "RE"
                
                Case "DY"
                    WControla = 0
                    spArticulo = "ConsultaArticulo " + "'" + XArticuloDy + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
        
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        WCodigo = XArticuloDy
                        WEntradas = Str$(rstArticulo!Entradas + Cantidad)
                        WSalidas = Str$(rstArticulo!Salidas)
                        WDate = Date$
                        rstArticulo.Close
                
                        XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                        spArticulo = "ModificaArticuloMovimientos " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        If WControla = 0 And Val(Lote) <> 0 Then
                            XParam = "'" + Lote + "','" _
                                    + XArticuloDy + "'"
                            spLaudo = "ListaLaudoArticulo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WClave = rstLaudo!Clave
                                WSaldo = Str$(rstLaudo!Saldo + Cantidad)
                                WDate = Date$
                                rstLaudo.Close
                                
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spLaudo = "ModificaLaudoSaldo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                
                                    Else
                                    
                                XParam = "'" + XArticuloDy + "','" _
                                        + Lote + "'"
                                spMovguia = "ListaMovguiaLote " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + Cantidad)
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
                
                Case Else
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        WCodigo = Articulo
                        WEntradas = Str$(rstTerminado!Entradas + Cantidad)
                        WLinea = rstTerminado!Linea
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                                    + WEntradas + "','" _
                                + WDate + "'"
                                                    
                        spTerminado = "ModificaTerminadoEntradas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        If WControla = 0 Then
                            XParam = "'" + Lote + "','" _
                                        + Articulo + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WClave = rstHoja!Clave
                                WSaldo = Str$(rstHoja!Saldo + Cantidad)
                                WDate = Date$
                                WMarca = ""
                                rstHoja.Close
                                
                                XParam = "'" + WClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "','" _
                                            + WMarca + "'"
                                spHoja = "ModificaHojaSaldo2 " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                            Else
                                        
                                XParam = "'" + Articulo + "','" _
                                                + Lote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + Cantidad)
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
                    
                    spTerminado = "ConsultaTerminado " + "'" + "NK" + Mid$(Articulo, 3, 10) + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = "NK" + Mid$(Articulo, 3, 10)
                        WEntradas = Str$(rstTerminado!Entradas - Cantidad)
                        WLinea = rstTerminado!Linea
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                                    + WEntradas + "','" _
                                    + WDate + "'"
                                                                           
                        spTerminado = "ModificaTerminadoEntradas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                        XParam = "'" + Cliente.Text + "','" _
                            + Lote + "','" _
                            + "NK" + Mid$(Articulo, 3, 10) + "'"
                        spEntdev = "ConsultaEntdev2 " + XParam
                        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEntdev.RecordCount > 0 Then
                            XLaboratorio = Str$(rstEntdev!Laboratorio + Cantidad)
                            XSaldo = Str$(rstEntdev!Cantidad - Val(XLaboratorio))
                            rstEntdev.Close
                            XParam = "'" + Cliente.Text + "','" _
                                + Lote + "','" _
                                + "NK" + Mid$(Articulo, 3, 10) + "','" _
                                + XSaldo + "','" _
                                + XLaboratorio + "'"
                            spEntdev = "ModificaEntdev2 " + XParam
                            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                
                    End If
                    
            End Select
            
            End If
            
        Next Da
        
        spNumero = "ConsultaNumero " + "'" + "01" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            WCodigo = "01"
            WNumero = Numero.Text
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call Impresion

        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    WPrecio.Caption = ""
    
    WArticulo.SetFocus
    
End Sub


Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
  
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
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
            Else
        Numero.Text = ""
    End If
    
    Graba.Enabled = True
    Borra.Enabled = True
    Ingresa.Enabled = True
    
    Numero.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        WCliente = Cliente.Text
        WTerminado = WArticulo.Text
        WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
        WClave = Cliente.Text + WArticulo.Text
        WClaveMp = Cliente.Text + WArti
        
        If Left$(WArticulo.Text, 2) = "DY" Then
            XTipoPro = "M"
                Else
            XTipoPro = "T"
        End If
        
        Select Case XTipoPro
            Case "M"
                spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    WEntra = "S"
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    rstPreciosMp.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
                XArticulo = Left$(WArticulo.Text, 3) + Right$(WArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + XArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WEntra = "S"
                    WDescripcion.Caption = rstPrecios!Descripcion
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    rstPrecios.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WEntrada.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEntrada_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WEntrada.Text = Pusing("###,###.##", WEntrada.Text)
        If Left$(WArticulo.Text, 2) = "DY" Then
            WTipopro.Text = "DY"
            WLote.SetFocus
                Else
            WTipopro.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipopro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipopro.Text = "PT" Or WTipopro.Text = "NK" Or WTipopro.Text = "RE" Or WTipopro.Text = "DY" Then
            WLote.SetFocus
        End If
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WLote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Select Case WTipopro.Text
            Case "PT"
                WEntra = "N"
            
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
            
                If WControla = 0 Then
                    XParam = "'" + WLote.Text + "','" _
                                 + WArticulo.Text + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WEntra = "S"
                        rstHoja.Close
                    End If
                    
                    If WEntra = "N" Then
                        XParam = "'" + WArticulo.Text + "','" _
                                + WLote.Text + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                        Else
                        
                    WEntra = "S"
                    
                End If
        
                If WEntra = "N" Then
                    m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
                
            Case "DY"
                WEntra = "N"
        
                WControla = 0
                WArticulo.Text = UCase(WArticulo.Text)
                WArtiDy = Left$(WArticulo.Text, 3) + Right$(WArticulo.Text, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArtiDy + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    rstArticulo.Close
                End If
                
                If WControla = 0 Then
                    WCanti = 0
                    XParam = "'" + WLote.Text + "','" _
                                + WArtiDy + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WEntra = "S"
                        rstLaudo.Close
                    End If
                    
                    If WEntra = "N" Then
                        XParam = "'" + WArtiDy + "','" _
                                + WLote.Text + "'"
                        spMovguia = "ListaMovguiaLote " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                        Else
                    
                    WEntra = "S"
                End If
                
                If WEntra = "N" Then
                    m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
            
            Case Else
                
                Rem WArti = WTipopro.Text + Mid$(WArticulo.Text, 3, 10)
                Rem WEntra = "N"
                
                Rem WControla = 0
                Rem spTerminado = "ConsultaTerminado " + "'" + WArti + "'"
                Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                Rem If rstTerminado.RecordCount > 0 Then
                Rem     WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                Rem     rstTerminado.Close
                Rem End If
                
                Rem If WControla = 0 Then
                Rem     XParam = "'" + WLote.Text + "','" _
                Rem         + WArti + "'"
                Rem     spHoja = "ListaHojaProducto " + XParam
                Rem     Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                Rem    If rstHoja.RecordCount > 0 Then
                Rem         WEntra = "S"
                Rem        rstHoja.Close
                Rem     End If
                Rem
                Rem     If WEntra = "N" Then
                Rem       XParam = "'" + WArti + "','" _
                Rem                 + WLote.Text + "'"
                Rem         spMovguia = "ListaMovguiaLote1 " + XParam
                Rem         Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                Rem         If rstMovguia.RecordCount > 0 Then
                Rem             WEntra = "S"
                Rem             rstMovguia.Close
                Rem         End If
                Rem     End If
                Rem
                Rem         Else
                Rem
                Rem     WEntra = "S"
                Rem
                Rem End If
            
                WEntra = "S"
                If WEntra = "N" Then
                    m$ = WArti + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
        End Select
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
                WPago1 = rstClientes!Pago1
                WPago2 = rstClientes!pago2
                WVendedor = rstClientes!Vendedor
                WProvincia = rstClientes!Provincia
                WRubro = rstClientes!Rubro
                WCodIva = rstClientes!Iva
                WCodIb = IIf(IsNull(rstClientes!Ib), "0", rstClientes!Ib)
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
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            If Mid$(Claveven$, 7, 2) = "DY" Then
            
                spPreciosMp = "ConsultaPreciosMp " + "'" + Claveven$ + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    DBGrid1.Col = 0
                    DBGrid1.Text = Mid$(Claveven$, 7, 3) + "00" + Right$(Claveven$, 7)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    WArticulo.Text = Mid$(Claveven$, 7, 3) + "00" + Right$(Claveven$, 7)
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    rstPreciosMp.Close
                End If
                
                XArticulo = Left$(WArticulo, 3) + Right$(WArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + XArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    WDescripcion.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
            
                    Else
            
                spPrecios = "ConsultaPrecios " + "'" + Claveven$ + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstPrecios!Terminado
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                
                    WArticulo.Text = rstPrecios!Terminado
                    WDescripcion.Caption = rstPrecios!Descripcion
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    
                    rstPrecios.Close
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WCantidad.SetFocus
                    
                End If
                
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6, 7
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
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
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
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
             DBGrid1.Columns(newcnt).Caption = "Entrada"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 700
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i
 
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
            Else
        Numero.Text = ""
    End If
 
    Rem DBGrid1.FirstRow = 0
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
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
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WEntrada.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WTipopro.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = Pusing("###,###", WLote.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
            
            WAnterior = DBGrid1.Row
                
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WEntrada.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WTipopro.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = Pusing("###,###", WLote.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

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
    Erase Auxiliar
    
    XParam = "'" + "02" + "','" _
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
                    DBGrid1.Text = !Articulo
                    Auxi1 = !Articulo
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Abs(!Entrada)
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Precio))
                    
                    DBGrid1.Col = 4
                    DBGrid1.Text = Abs(!Cantidad)
                    
                    DBGrid1.Col = 5
                    DBGrid1.Text = !Tipopro
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = Pusing("######", Str$(!lote1))
                
                    Rem Paridad.Text = !Lote
                    
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
    
    For Da = 1 To XRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        
        If Left$(Auxi1, 2) = "DY" Then
        
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
        
                Else
                
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
            End If
            
        End If
                
    Next Da

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = False
    Borra.Enabled = False
    Ingresa.Enabled = False
    
    Call Calcula_Click

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "02" + Auxi + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Paridad.Text = rstCtacte!Paridad
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!pago2
                WVendedor = rstCliente!Vendedor
                WProvincia = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WCodIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
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
            Rem     A% = MsgBox(m$, 0, "Ingreso de Devoluciones")
            Rem     Numero.SetFocus
            Rem         Else
            Rem     Graba.Enabled = True
            Rem     Borra.Enabled = True
            Rem     Ingresa.Enabled = True
            Rem     WNumero = Numero.Text
            Rem     Numero.Text = WNumero
            Rem     Cliente.SetFocus
            Rem End If
            Graba.Enabled = True
            Borra.Enabled = True
            Ingresa.Enabled = True
            WNumero = Numero.Text
            Numero.Text = WNumero
            Cliente.SetFocus
                
        End If
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Cliente.Text = rstCliente!Cliente
            DesCliente.Caption = rstCliente!Razon
            WPago1 = rstCliente!Pago1
            WPago2 = rstCliente!pago2
            WVendedor = rstCliente!Vendedor
            WProvincia = rstCliente!Provincia
            WRubro = rstCliente!Rubro
            WCodIva = rstCliente!Iva
            WCodIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            WDirentrega = rstCliente!DirEntrega
            rstCliente.Close
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            Fecha.SetFocus
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
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            Paridad.SetFocus
                Else
            m$ = "Formato de fecha invalida"
            a% = MsgBox(m$, 0, "Emision de Devolucion")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Paridad.Text) = 0 Then
            m$ = "No exsite paridad cargada para esta fecha"
            a% = MsgBox(m$, 0, "Emision de Devolucion")
            Paridad.SetFocus
                Else
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
        End If
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

Sub Impresion()

    If Val(WEmpresa) = 1 Then
        Rem Open "dada.txt" For Output As #1
        Open "lpt1" For Output As #1
            Else
        If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
            Rem Open "dada.txt" For Output As #1
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
                Else
            Open "dada.txt" For Output As #1
            Rem Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
    End If
    
    Rem Width #1, 255


    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    ImporteIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
    ImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
    Impotot = Val(Total.Caption) * Val(Paridad.Text)

    For XX% = 1 To 2
    
        If Val(WEmpresa) = 1 Then
            Rem Print #1, ""
        End If

        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(55); "NOTA DE CREDITO"
        Print #1, ""
        Print #1, Tab(59); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(8); WRazon
        Print #1, Tab(8); WDireccion
        Print #1, Tab(8); Left$(WLocalidad, 33);
        Print #1, Tab(55); Cliente.Text
        Print #1, Tab(8); Provincia(Val(WProvincia)); " ("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(8); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, ""
        Print #1, Tab(5); WPago
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Impre = 0

        For a = 0 To 3
        
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
            
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                    Print #1, Tab(1); Alinea("#####.##", Str$(Cantidad));
                    Print #1, " Kg";
                    Print #1, Tab(15); Left$(Descri, 38);
                    parcial = Precio * Cantidad
                    Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                    Rem    Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem    Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem            Else
                    Rem    Precio = Str$(Val(Precio) * 1.19)
                    Rem    Parcial = Str$(Val(Parcial) * 1.19)
                    Rem    Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem    Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem End If
                    Print #1, Tab(56); "U$S"; Alinea("####.##", Str$(Precio));
                    Print #1, Tab(68); Alinea("###,###.##", Str$(parcial))
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a

        For aa = Impre To 21
                Print #1, ""
        Next aa

        Rem M# = Total# / 100
        Rem GoSub 4630
        
        Print #1, Tab(1); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
        Print #1, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
        Print #1, Tab(1); "TIPO DE CAMBIO:";
        Print #1, Alinea("##.##", Paridad.Text);
        Print #1, " I.V.A.:";
        Print #1, Alinea("#,###.##", Str$(ImpoIva));
        Print #1, " TOTAL:";
        Print #1, Alinea("###,###.##", Str$(Impotot))
        Print #1, Tab(1); "IMPORTE NETO:";
        Print #1, Alinea("###,###.##", Str$(ImpoNeto));
        If ImporteIb <> 0 Then
            Print #1, " PERCEPCION DE INGRESOS BRUTOS:";
            Print #1, Alinea("#,###.##", Str$(ImporteIb))
                Else
            Print #1, ""
        End If
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(WEmpresa) <> 1 Then
            Print #1, ""
        End If
        
        Print #1, Tab(65); "U$S"; Alinea("###,###.##", Str$(XNeto))

        If Val(Interes.Caption) <> 0 Then
                Print #1, Tab(56); "Interes";
                Print #1, Tab(65); "U$S"; Alinea("###,###.##", Interes.Caption)
                                                  Else
                Print #1, ""
        End If

        If Val(Dto.Caption) <> 0 Then
                Print #1, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                Print #1, Tab(65); "U$S"; Alinea("###,###.##", Dto.Caption)
                        Else
                Print #1, ""
        End If

        Print #1, Tab(3); M1;
        Print #1, Tab(65); "U$S"; Alinea("###,###.##", Neto.Caption)
        Print #1, Tab(3); M2;
        If Val(Iva1.Caption) <> 0 Then
                Print #1, Tab(61); "19";
                Print #1, Tab(65); "U$S"; Alinea("###,###.##", Iva1.Caption)
                        Else
                Print #1, ""
        End If

        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select
        
        If Val(ImpoIb.Caption) <> 0 Then
                Print #1, Tab(56); "Ret.I.B.";
                Print #1, Tab(65); "U$S"; Alinea("###,###.##", ImpoIb.Caption)
                        Else
                If Val(Iva2.Caption) <> 0 Then
                    Print #1, Tab(60); "10.5";
                    Print #1, Tab(65); "U$S"; Alinea("###,###.##", Iva2.Caption)
                        Else
                    Print #1, ""
                End If
        End If

        Print #1, Tab(65); "U$S"; Alinea("###,###.##", Total.Caption);
        Print #1, Chr$(12)

        Next XX%

        Close #1
        Da = 0

End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)
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
            
                    Da = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To Da
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


