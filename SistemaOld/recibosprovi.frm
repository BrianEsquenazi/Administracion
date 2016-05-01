VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRecibosProvi 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Recibos Provisorios"
   ClientHeight    =   8220
   ClientLeft      =   690
   ClientTop       =   420
   ClientWidth     =   10575
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   10575
   Begin VB.CommandButton DiasII 
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
      Height          =   300
      Left            =   9600
      TabIndex        =   56
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame PantaDias 
      Caption         =   "Informe la tasa mensual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   54
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox DiasTasa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         MaxLength       =   15
         TabIndex        =   55
         Text            =   " "
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox TotalRecibo 
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
      Height          =   315
      Left            =   9360
      MaxLength       =   15
      TabIndex        =   52
      Text            =   " "
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Toto 
      Height          =   285
      Left            =   5160
      TabIndex        =   51
      Top             =   0
      Width           =   150
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
      Height          =   1425
      ItemData        =   "recibosprovi.frx":0000
      Left            =   5520
      List            =   "recibosprovi.frx":0007
      TabIndex        =   50
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
   End
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
      Left            =   5520
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame IngresaCuit 
      Caption         =   "Informe Cuit del Firmante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   47
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Cuit 
         Height          =   285
         Left            =   480
         MaxLength       =   15
         TabIndex        =   48
         Text            =   " "
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox Lectora 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Frame EntraComproSuss 
      Caption         =   "Nro de Comprobante Suss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   44
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ComproSuss 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   45
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame EntraComproIb 
      Caption         =   "Nro de Comprobante I.B."
      Height          =   1095
      Left            =   4680
      TabIndex        =   42
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ComproIB 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   43
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame EntraComproIva 
      Caption         =   "Nro de Comprobante Iva"
      Height          =   1095
      Left            =   4320
      TabIndex        =   40
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ComproIva 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   41
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame EntraComproGanan 
      Caption         =   "Nro de Comprobante Ganancias"
      Height          =   1095
      Left            =   3960
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ComproGanan 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   39
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox RetSuss 
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
      Left            =   6480
      MaxLength       =   15
      TabIndex        =   36
      Text            =   " "
      Top             =   2040
      Width           =   975
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
      Left            =   8280
      MaxLength       =   15
      TabIndex        =   35
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Ingrecuenta 
      Caption         =   "Ingreso de Cuenta Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Cuenta1 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   33
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   31
      Text            =   " "
      Top             =   1560
      Width           =   1215
   End
   Begin Crystal.CrystalReport listado 
      Left            =   10080
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Imprerec.rpt"
      CopiesToPrinter =   2
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox RetOtra 
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
      MaxLength       =   15
      TabIndex        =   5
      Text            =   " "
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox RetIva 
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
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   7
      Text            =   " "
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Retganancias 
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   4
      Text            =   " "
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Recibos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   5295
      Begin VB.OptionButton Tipo3 
         Caption         =   "Varios"
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
         Left            =   3720
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Cobro de Cta.Cte."
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.TextBox Clientes 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "recibosprovi.frx":0015
      TabIndex        =   6
      Top             =   2400
      Width           =   9975
   End
   Begin VB.TextBox Recibo 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9600
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9600
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9600
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   9840
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9600
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Total Recibo"
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
      Left            =   9480
      TabIndex        =   53
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Re.Suss"
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
      Left            =   5760
      TabIndex        =   37
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   7560
      TabIndex        =   34
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Cuenta Contable"
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
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
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
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
      Left            =   8280
      TabIndex        =   27
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
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
      Left            =   2880
      TabIndex        =   26
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Ch.   3) Doc.  4) Varios"
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
      Left            =   4320
      TabIndex        =   25
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Ret. I.B."
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
      Left            =   3960
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Ret.Iva"
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
      Left            =   2160
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Rte.Ganan."
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
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label DesClientes 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   18
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Cilente"
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "NRo de Recibo"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgRecibosProvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private WSaldoUs As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double
Dim rstRecibosProvi As Recordset
Dim spRecibosProvi As String
Dim rstCuit As Recordset
Dim spCuit As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstCambio As Recordset
Dim spCambio As String
Dim rstInterres As Recordset
Dim spInteres As String
Dim XParam As String
Dim XParidad As String
Dim WParidad As String
Dim Pari As Double
Dim WEntra(100, 120) As String
Dim ZPasa As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim ZBancos(1000) As String
Dim XTipo1 As String
Dim XNumero1 As String
Dim ZClaveCheque(100, 10) As String
Dim ZZSuma As Double
Dim ZZInteres As Double

Dim ZZDa As Integer

Private Sub Suma_Datos()

    Rem If Val(WProv) = 24 Then
    Rem     Paridad.Text = "1"
    Rem End If

    Debitos.Caption = ""
    Creditos.Caption = ""
    ZPasa = "S"
    
    Creditos.Caption = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text) + Val(RetSuss.Text))
    For iRow = 0 To 19
    
        DbGrid1.Col = 9
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DbGrid1.Text))
        End If
        
        DbGrid1.Col = 5
        DbGrid1.Row = iRow
        ZTipo = Val(DbGrid1.Text)
        
        DbGrid1.Col = 7
        DbGrid1.Row = iRow
        ZFecha = DbGrid1.Text
        
        WDias = 0
        WFechaDesde = ZFecha
        WFechaHasta = Fecha.Text
        
        WOrdFechaDesde = Right$(WFechaDesde, 4) + Mid$(WFechaDesde, 4, 2) + Left$(WFechaDesde, 2)
        WOrdFechaHasta = Right$(WFechaHasta, 4) + Mid$(WFechaHasta, 4, 2) + Left$(WFechaHasta, 2)
        
        If ZTipo = 2 And WOrdFechaDesde < WOrdFechaHasta Then
        
            XFec1 = ZFecha
            Call Valida_fecha1(XFec1, Auxi)
            If Auxi <> "S" Then
                ZPasa = "N"
                Exit Sub
            End If
        
            Do
                WDias = WDias + 1
                XFec1 = WFechaDesde
                SumaDia = 2
                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                WFechaDesde = XFec2
                If WFechaDesde = WFechaHasta Then
                    Exit Do
                End If
            Loop
            
            If WDias > 30 Then
                ZPasa = "N"
                Exit Sub
            End If
            
        End If
        
        If ZTipo > 4 Then
            ZPasa = "N"
            Exit Sub
        End If
        
    Next iRow
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
End Sub

Private Sub Lee_Datos()

    For iRow = 0 To 19
        For iCol = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow
    cmdAdd.Enabled = True
    
    Auxi1 = Recibo.Text
    Call Ceros(Auxi1, 8)
    
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        ClaveRecibo = Recibo.Text + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
        
            Select Case Val(rstRecibosProvi!Tiporeg)
                Case 1
                    Debito = Debito + 1
                    DbGrid1.Row = Debito - 1
                    DbGrid1.Col = 0
                    DbGrid1.Text = rstRecibosProvi!Tipo1
                    DbGrid1.Col = 1
                    DbGrid1.Text = rstRecibosProvi!Letra1
                    DbGrid1.Col = 2
                    DbGrid1.Text = rstRecibosProvi!Punto1
                    DbGrid1.Col = 3
                    DbGrid1.Text = rstRecibosProvi!Numero1
                    DbGrid1.Col = 4
                    DbGrid1.Text = rstRecibosProvi!Importe1
                    DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                Case 2
                    Credito = Credito + 1
                    DbGrid1.Row = Credito - 1
                    DbGrid1.Col = 5
                    DbGrid1.Text = rstRecibosProvi!Tipo2
                    DbGrid1.Col = 6
                    DbGrid1.Text = rstRecibosProvi!Numero2
                    DbGrid1.Col = 7
                    DbGrid1.Text = rstRecibosProvi!Fecha2
                    DbGrid1.Col = 8
                    DbGrid1.Text = rstRecibosProvi!Banco2
                    DbGrid1.Col = 9
                    DbGrid1.Text = rstRecibosProvi!Importe2
                    DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    WCuenta(DbGrid1.Row) = rstRecibosProvi!Cuenta
                    
                    ZClaveCheque(DbGrid1.Row + 1, 1) = IIf(IsNull(rstRecibosProvi!ClaveCheque), "", rstRecibosProvi!ClaveCheque)
                    ZClaveCheque(DbGrid1.Row + 1, 2) = IIf(IsNull(rstRecibosProvi!BancoCheque), "", rstRecibosProvi!BancoCheque)
                    ZClaveCheque(DbGrid1.Row + 1, 3) = IIf(IsNull(rstRecibosProvi!SucursalCheque), "", rstRecibosProvi!SucursalCheque)
                    ZClaveCheque(DbGrid1.Row + 1, 4) = IIf(IsNull(rstRecibosProvi!ChequeCheque), "", rstRecibosProvi!ChequeCheque)
                    ZClaveCheque(DbGrid1.Row + 1, 5) = IIf(IsNull(rstRecibosProvi!CuentaCheque), "", rstRecibosProvi!CuentaCheque)
                    ZClaveCheque(DbGrid1.Row + 1, 6) = IIf(IsNull(rstRecibosProvi!Cuit), "", rstRecibosProvi!Cuit)
                    
                    ZReciboDefinitivo = IIf(IsNull(rstRecibosProvi!ReciboDefinitivo), "0", rstRecibosProvi!ReciboDefinitivo)
                    If Val(rstRecibosProvi!Tipo2) = "02" And rstRecibosProvi!Estado2 = "X" Then
                         cmdAdd.Enabled = False
                    End If
                    If ZReciboDefinitivo <> 0 Then
                         cmdAdd.Enabled = False
                    End If
                    
                Case Else
            End Select
            rstRecibosProvi.Close
            
                Else
            Exit Do
        End If
    Loop
End Sub

Sub Verifica_datos()
    If Val(Retganancias.Text) = 0 Then
        Retganancias.Text = "0"
    End If
    If Val(RetIva.Text) = 0 Then
        RetIva.Text = "0"
    End If
    If Val(RetOtra.Text) = 0 Then
        RetOtra.Text = "0"
    End If
    If Val(RetSuss.Text) = 0 Then
        RetSuss.Text = "0"
    End If
End Sub

Sub Format_datos()
    Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
    RetIva.Text = Alinea("###,###.##", RetIva.Text)
    RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
    RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
End Sub

Sub Imprime_Datos()
    spClientes = "ConsultaClientes " + "'" + Clientes.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        Clientes.Text = rstClientes!Cliente
        DesClientes.Caption = rstClientes!Razon
        WRazon = rstClientes!Razon
        WDireccion = rstClientes!Direccion
        WLocalidad = rstClientes!Localidad
        WPostal = rstClientes!Postal
        WProvincia = Provincia(rstClientes!Provincia)
        WProv = rstClientes!Provincia
        rstClientes.Close
        Call Format_datos
    End If
End Sub

Private Sub cmdAdd_Click()

    If Recibo.Text <> "" Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        Existe = "N"
    
        ClaveRecibo = Recibo.Text + "01"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            Existe = "S"
            rstRecibosProvi.Close
        End If
        
        If Existe = "S" Then
        
            Existe = "N"
            
            Do
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                ClaveRecibo = Recibo.Text + Auxi1
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM RecibosProvi"
                ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibosProvi.RecordCount > 0 Then
                    If Val(rstRecibosProvi!Tiporeg) = 2 Then
                        If Val(rstRecibosProvi!Tipo2) = 2 Then
                            If rstRecibosProvi!Estado2 <> "P" Or rstRecibosProvi!ReciboProvisorio <> 0 Then
                                Existe = "S"
                            End If
                        End If
                    End If
                    rstRecibosProvi.Close
                        Else
                    Exit Do
                End If
            Loop
            
        End If
        
    End If
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        If ZPasa = "N" Then
            m1$ = "Error en la carga de fecha de cheques"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            Exit Sub
        End If
        
        If Val(Creditos.Caption) <> Val(TotalRecibo.Text) Then
            m1$ = "El total de los valores ingresados no coinciden con el total del recibo informado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            Exit Sub
        End If
        
        ZSql = ""
        ZSql = ZSql + "DELETE RecibosProvi"
        ZSql = ZSql + " Where Recibo = " + "'" + Recibo.Text + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        
        Renglon = 0
            
        For iRow = 0 To 19
        
            WRow = iRow
            DbGrid1.Col = 4
            DbGrid1.Row = iRow
            If Val(DbGrid1.Text) <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                        
                XRecibo = Recibo.Text
                XRenglon = Auxi1
                XClientes = Clientes.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    XTipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    XTipoRec = "2"
                End If
                If Tipo3.Value = True Then
                    XTipoRec = "3"
                End If
                XRetganancias = Str$(Val(Retganancias.Text))
                XRetIva = Str$(Val(RetIva.Text))
                XRetotra = Str$(Val(RetOtra.Text))
                XRetencion = ""
                XTiporeg = "1"
                DbGrid1.Col = 0
                XTipo1 = DbGrid1.Text
                DbGrid1.Col = 1
                XLetra1 = DbGrid1.Text
                DbGrid1.Col = 2
                XPunto1 = DbGrid1.Text
                DbGrid1.Col = 3
                XNumero1 = DbGrid1.Text
                DbGrid1.Col = 4
                XImporte1 = DbGrid1.Text
                XImporteBaja = DbGrid1.Text
                XTipo2 = ""
                XNumero2 = ""
                XFecha2 = ""
                XFechaOrd2 = ""
                XBanco2 = ""
                XImporte2 = ""
                XEstado2 = ""
                XObservaciones = Observaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Credito)
                XCuenta = ""
                XDestino = ""
                XImpolist = ""
                XImpo1list = ""
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""
                XReciboDefinitivo = "0"
                
                XClaveCheque = ""
                XBancoCheque = ""
                XSucursalCheque = ""
                XChequeCheque = ""
                XCuentaCheque = ""
                XCuit = ""
                
                ZSql = "INSERT INTO RecibosProvi ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Recibo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Fechaord ,"
                ZSql = ZSql + "TipoRec ,"
                ZSql = ZSql + "RetGanancias ,"
                ZSql = ZSql + "RetIva ,"
                ZSql = ZSql + "RetOtra ,"
                ZSql = ZSql + "Retencion ,"
                ZSql = ZSql + "TipoReg ,"
                ZSql = ZSql + "Tipo1 ,"
                ZSql = ZSql + "Letra1 ,"
                ZSql = ZSql + "Punto1 ,"
                ZSql = ZSql + "Numero1 ,"
                ZSql = ZSql + "Importe1 ,"
                ZSql = ZSql + "Tipo2 ,"
                ZSql = ZSql + "Numero2 ,"
                ZSql = ZSql + "Fecha2 ,"
                ZSql = ZSql + "banco2 ,"
                ZSql = ZSql + "Importe2 ,"
                ZSql = ZSql + "Estado2 ,"
                ZSql = ZSql + "Empresa ,"
                ZSql = ZSql + "FechaOrd2 ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Impolist ,"
                ZSql = ZSql + "Impo1list ,"
                ZSql = ZSql + "Destino ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Marca ,"
                ZSql = ZSql + "FechaDepo ,"
                ZSql = ZSql + "FechaDepoOrd ,"
                ZSql = ZSql + "ClaveCheque ,"
                ZSql = ZSql + "BancoCheque ,"
                ZSql = ZSql + "SucursalCheque ,"
                ZSql = ZSql + "ChequeCheque ,"
                ZSql = ZSql + "CuentaCheque ,"
                ZSql = ZSql + "ReciboDefinitivo ,"
                ZSql = ZSql + "Cuit)"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + XClave + "',"
                ZSql = ZSql + "'" + XRecibo + "',"
                ZSql = ZSql + "'" + XRenglon + "',"
                ZSql = ZSql + "'" + XClientes + "',"
                ZSql = ZSql + "'" + XFecha + "',"
                ZSql = ZSql + "'" + XFechaOrd + "',"
                ZSql = ZSql + "'" + XTipoRec + "',"
                ZSql = ZSql + "'" + XRetganancias + "',"
                ZSql = ZSql + "'" + XRetIva + "',"
                ZSql = ZSql + "'" + XRetotra + "',"
                ZSql = ZSql + "'" + XRetencion + "',"
                ZSql = ZSql + "'" + XTiporeg + "',"
                ZSql = ZSql + "'" + XTipo1 + "',"
                ZSql = ZSql + "'" + XLetra1 + "',"
                ZSql = ZSql + "'" + XPunto1 + "',"
                ZSql = ZSql + "'" + XNumero1 + "',"
                ZSql = ZSql + "'" + XImporte1 + "',"
                ZSql = ZSql + "'" + XTipo2 + "',"
                ZSql = ZSql + "'" + XNumero2 + "',"
                ZSql = ZSql + "'" + XFecha2 + "',"
                ZSql = ZSql + "'" + XBanco2 + "',"
                ZSql = ZSql + "'" + XImporte2 + "',"
                ZSql = ZSql + "'" + XEstado2 + "',"
                ZSql = ZSql + "'" + XEmpresa + "',"
                ZSql = ZSql + "'" + XFechaOrd2 + "',"
                ZSql = ZSql + "'" + XImporte + "',"
                ZSql = ZSql + "'" + XObservaciones + "',"
                ZSql = ZSql + "'" + XImpolist + "',"
                ZSql = ZSql + "'" + XImpo1list + "',"
                ZSql = ZSql + "'" + XDestino + "',"
                ZSql = ZSql + "'" + XCuenta + "',"
                ZSql = ZSql + "'" + XMarca + "',"
                ZSql = ZSql + "'" + XFechaDepo + "',"
                ZSql = ZSql + "'" + XFechaDepoOrd + "',"
                ZSql = ZSql + "'" + XClaveCheque + "',"
                ZSql = ZSql + "'" + XBancoCheque + "',"
                ZSql = ZSql + "'" + XSucursalCheque + "',"
                ZSql = ZSql + "'" + XChequeCheque + "',"
                ZSql = ZSql + "'" + XCuentaCheque + "',"
                ZSql = ZSql + "'" + XReciboDefinitivo + "',"
                ZSql = ZSql + "'" + XCuit + "')"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
                
        Next iRow
            
            
        For iRow = 0 To 19
        
            DbGrid1.Col = 9
            DbGrid1.Row = iRow
            If Val(DbGrid1.Text) <> 0 Then
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                    
                XRecibo = Recibo.Text
                XRenglon = Auxi1
                XClientes = Clientes.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    XTipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    XTipoRec = "2"
                End If
                If Tipo3.Value = True Then
                    XTipoRec = "3"
                End If
                XRetganancias = Str$(Val(Retganancias.Text))
                XRetIva = Str$(Val(RetIva.Text))
                XRetotra = Str$(Val(RetOtra.Text))
                XRetencion = ""
                XTiporeg = "2"
                XTipo1 = ""
                XLetra1 = ""
                XPunto1 = ""
                XNumero1 = ""
                XImporte1 = ""
                DbGrid1.Col = 5
                XTipo2 = DbGrid1.Text
                DbGrid1.Col = 6
                XNumero2 = DbGrid1.Text
                DbGrid1.Col = 7
                XFecha2 = DbGrid1.Text
                XFechaOrd2 = Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                DbGrid1.Col = 8
                XBanco2 = DbGrid1.Text
                DbGrid1.Col = 9
                XImporte2 = DbGrid1.Text
                XEstado2 = "P"
                XObservaciones = Observaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Credito)
                If XTipo2 = 4 Then
                    XCuenta = WCuenta(iRow)
                        Else
                    XCuenta = ""
                End If
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""
                If Val(XTipo2) = 1 Or Val(XTipo2) = 4 Then
                    XEstado2 = "X"
                End If
                
                XClaveCheque = ZClaveCheque(DbGrid1.Row + 1, 1)
                XBancoCheque = ZClaveCheque(DbGrid1.Row + 1, 2)
                XSucursalCheque = ZClaveCheque(DbGrid1.Row + 1, 3)
                XChequeCheque = ZClaveCheque(DbGrid1.Row + 1, 4)
                XCuentaCheque = ZClaveCheque(DbGrid1.Row + 1, 5)
                XCuit = ZClaveCheque(DbGrid1.Row + 1, 6)
                XReciboDefinitivo = "0"
                    
                ZSql = "INSERT INTO RecibosProvi ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Recibo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Fechaord ,"
                ZSql = ZSql + "TipoRec ,"
                ZSql = ZSql + "RetGanancias ,"
                ZSql = ZSql + "RetIva ,"
                ZSql = ZSql + "RetOtra ,"
                ZSql = ZSql + "Retencion ,"
                ZSql = ZSql + "TipoReg ,"
                ZSql = ZSql + "Tipo1 ,"
                ZSql = ZSql + "Letra1 ,"
                ZSql = ZSql + "Punto1 ,"
                ZSql = ZSql + "Numero1 ,"
                ZSql = ZSql + "Importe1 ,"
                ZSql = ZSql + "Tipo2 ,"
                ZSql = ZSql + "Numero2 ,"
                ZSql = ZSql + "Fecha2 ,"
                ZSql = ZSql + "banco2 ,"
                ZSql = ZSql + "Importe2 ,"
                ZSql = ZSql + "Estado2 ,"
                ZSql = ZSql + "Empresa ,"
                ZSql = ZSql + "FechaOrd2 ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Impolist ,"
                ZSql = ZSql + "Impo1list ,"
                ZSql = ZSql + "Destino ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Marca ,"
                ZSql = ZSql + "FechaDepo ,"
                ZSql = ZSql + "FechaDepoOrd ,"
                ZSql = ZSql + "ClaveCheque ,"
                ZSql = ZSql + "BancoCheque ,"
                ZSql = ZSql + "SucursalCheque ,"
                ZSql = ZSql + "ChequeCheque ,"
                ZSql = ZSql + "CuentaCheque ,"
                ZSql = ZSql + "ReciboDefinitivo ,"
                ZSql = ZSql + "Cuit)"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + XClave + "',"
                ZSql = ZSql + "'" + XRecibo + "',"
                ZSql = ZSql + "'" + XRenglon + "',"
                ZSql = ZSql + "'" + XClientes + "',"
                ZSql = ZSql + "'" + XFecha + "',"
                ZSql = ZSql + "'" + XFechaOrd + "',"
                ZSql = ZSql + "'" + XTipoRec + "',"
                ZSql = ZSql + "'" + XRetganancias + "',"
                ZSql = ZSql + "'" + XRetIva + "',"
                ZSql = ZSql + "'" + XRetotra + "',"
                ZSql = ZSql + "'" + XRetencion + "',"
                ZSql = ZSql + "'" + XTiporeg + "',"
                ZSql = ZSql + "'" + XTipo1 + "',"
                ZSql = ZSql + "'" + XLetra1 + "',"
                ZSql = ZSql + "'" + XPunto1 + "',"
                ZSql = ZSql + "'" + XNumero1 + "',"
                ZSql = ZSql + "'" + XImporte1 + "',"
                ZSql = ZSql + "'" + XTipo2 + "',"
                ZSql = ZSql + "'" + XNumero2 + "',"
                ZSql = ZSql + "'" + XFecha2 + "',"
                ZSql = ZSql + "'" + XBanco2 + "',"
                ZSql = ZSql + "'" + XImporte2 + "',"
                ZSql = ZSql + "'" + XEstado2 + "',"
                ZSql = ZSql + "'" + XEmpresa + "',"
                ZSql = ZSql + "'" + XFechaOrd2 + "',"
                ZSql = ZSql + "'" + XImporte + "',"
                ZSql = ZSql + "'" + XObservaciones + "',"
                ZSql = ZSql + "'" + XImpolist + "',"
                ZSql = ZSql + "'" + XImpo1list + "',"
                ZSql = ZSql + "'" + XDestino + "',"
                ZSql = ZSql + "'" + XCuenta + "',"
                ZSql = ZSql + "'" + XMarca + "',"
                ZSql = ZSql + "'" + XFechaDepo + "',"
                ZSql = ZSql + "'" + XFechaDepoOrd + "',"
                ZSql = ZSql + "'" + XClaveCheque + "',"
                ZSql = ZSql + "'" + XBancoCheque + "',"
                ZSql = ZSql + "'" + XSucursalCheque + "',"
                ZSql = ZSql + "'" + XChequeCheque + "',"
                ZSql = ZSql + "'" + XCuentaCheque + "',"
                ZSql = ZSql + "'" + XReciboDefinitivo + "',"
                ZSql = ZSql + "'" + XCuit + "')"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                
                If Trim(XCuit) <> "" Then
                
                    XClaveCuit = XBancoCheque + XSucursalCheque + XCuentaCheque
            
                    ZSql = "Select *"
                    ZSql = ZSql + " FROM Cuit"
                    ZSql = ZSql + " Where Cuit.Clave = " + "'" + XClaveCuit + "'"
                    spCuit = ZSql
                    Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuit.RecordCount > 0 Then
                        rstCuit.Close
                            Else
                        ZSql = "INSERT INTO Cuit ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Banco ,"
                        ZSql = ZSql + "Sucursal ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Cuit)"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + XClaveCuit + "',"
                        ZSql = ZSql + "'" + XBancoCheque + "',"
                        ZSql = ZSql + "'" + XSucursalCheque + "',"
                        ZSql = ZSql + "'" + XCuentaCheque + "',"
                        ZSql = ZSql + "'" + XCuit + "')"
                        spCuit = ZSql
                        Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                End If
                    
            End If
               
        Next iRow
            
        ZSql = ""
        ZSql = ZSql + "UPDATE RecibosProvi SET "
        ZSql = ZSql + " RetSuss = " + "'" + RetSuss.Text + "',"
        ZSql = ZSql + " ComproGanan = " + "'" + ComproGanan.Text + "',"
        ZSql = ZSql + " ComproIva = " + "'" + ComproIva.Text + "',"
        ZSql = ZSql + " ComproIb = " + "'" + ComproIB.Text + "',"
        ZSql = ZSql + " ComproSuss = " + "'" + ComproSuss.Text + "'"
        ZSql = ZSql + " Where Recibo = " + "'" + Recibo.Text + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            

        Call CmdLimpiar_Click
        Recibo.SetFocus
        
    End If
End Sub


Private Sub cmdDelete_Click()
    If Recibo.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Recibo.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
    End If
    Clientes.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iRow = 0 To 19
        For iCol = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow
    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    ComproGanan.Text = ""
    ComproIva.Text = ""
    ComproIB.Text = ""
    ComproSuss.Text = ""
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Cuenta.Text = ""
    Paridad.Text = ""
    TotalRecibo.Text = ""
    
    cmdAdd.Enabled = True
    
    Erase ZClaveCheque
    
    Ingrecuenta.Visible = False
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
    
    Recibo.Text = ""
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM RecibosProvi"
    Rem spRecibosProvi = ZSql
    Rem Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibosProvi.RecordCount > 0 Then
    Rem     rstRecibosProvi.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibosProvi!ReciboMayor), "0", rstRecibosProvi!ReciboMayor)
    Rem     Recibo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
    Rem     rstRecibosProvi.Close
    Rem         Else
    Rem     Recibo.Text = "1"
    Rem End If
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstImpreRec
        .Close
    End With
    Recibo.SetFocus
    PrgRecibosProvi.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub


Private Sub DiasII_Click()
    PantaDias.Visible = True
    DiasTasa.Text = ""
    DiasTasa.SetFocus
End Sub

Private Sub DiasTasa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        PantaDias.Visible = False
        
        ZSql = ""
        ZSql = ZSql + "DELETE Interes"
        spInteres = ZSql
        Set rstInteres = db.OpenRecordset(spInteres, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZZSuma = 0
        ZZCodigo = 0
        
        FechaBase = Fecha.Text
        
        For iRow = 0 To 19
        
            DbGrid1.Row = iRow
        
            
            DbGrid1.Col = 5
            XTipo2 = DbGrid1.Text
            DbGrid1.Col = 6
            Numero2 = DbGrid1.Text
            DbGrid1.Col = 7
            XFecha2 = DbGrid1.Text
            DbGrid1.Col = 8
            Banco2 = DbGrid1.Text
            DbGrid1.Col = 9
            XImporte2 = DbGrid1.Text
            
            If Val(XImporte2) <> 0 Then
                If XFecha2 = "" Then
                    XFecha2 = Fecha.Text
                End If
                WFechaCheque = Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                WFechaRecibo = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If WFechaCheque < WFechaRecibo Then
                    XFecha2 = Fecha.Text
                End If
                XDias2 = DateDiff("d", FechaBase, XFecha2)
                ZZInteres = ((Val(XImporte2) * XDias2 * (Val(DiasTasa.Text) / 100)) / 365)
                Call Redondeo(ZZInteres)
                ZZSuma = ZZSuma + ZZInteres
                
                ZZCodigo = ZZCodigo + 1
                
                ZSql = "INSERT INTO Interes ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Razon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "FechaII ,"
                ZSql = ZSql + "Banco ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Dias ,"
                ZSql = ZSql + "Tasa ,"
                ZSql = ZSql + "Interes)"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + Str$(ZZCodigo) + "',"
                ZSql = ZSql + "'" + Clientes.Text + "',"
                ZSql = ZSql + "'" + DesClientes.Caption + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + Numero2 + "',"
                ZSql = ZSql + "'" + XFecha2 + "',"
                ZSql = ZSql + "'" + Banco2 + "',"
                ZSql = ZSql + "'" + XImporte2 + "',"
                ZSql = ZSql + "'" + Str$(XDias2) + "',"
                ZSql = ZSql + "'" + DiasTasa.Text + "',"
                ZSql = ZSql + "'" + Str$(ZZInteres) + "')"
                spInteres = ZSql
                Set rstInteres = db.OpenRecordset(spInteres, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next iRow
        
        Call Redondeo(ZZSuma)
        
        
        Listado.WindowTitle = ""
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        Listado.Destination = 1
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT Interes.Codigo, Interes.Cliente, Interes.Razon, Interes.Fecha, Interes.Numero, Interes.FechaII, Interes.Banco, Interes.Importe, Interes.Dias, Interes.Tasa, Interes.Interes " _
                + "From " _
                + DSQ + ".dbo.Interes Interes " _
                + "Where " _
                + "Interes.Codigo >= 0 AND " _
                + "Interes.Codigo <= 999999"
                
        Listado.Connect = Connect()
        
        Listado.GroupSelectionFormula = "{Interes.Codigo} in 0 to 999999"
        Listado.SelectionFormula = "{Interes.Codigo} in 0 to 999999"
        
        Listado.ReportFileName = "Interes.rpt"
        
        Listado.Action = 1
        
        
        f$ = "El interes a pagar es de  " + Str$(ZZSuma)
        A% = MsgBox(f$, 0, "Emision de Recibos")
    
    
    
    
    
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImpreRec
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(Recibo.Text) <> 0 Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Recibo = " + "'" + Recibo.Text + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            Existe = "S"
            Clientes.Text = rstRecibosProvi!Cliente
            Observaciones.Text = rstRecibosProvi!Observaciones
            Fecha.Text = rstRecibosProvi!Fecha
            Retganancias.Text = rstRecibosProvi!Retganancias
            RetIva.Text = rstRecibosProvi!RetIva
            RetOtra.Text = rstRecibosProvi!RetOtra
            RetSuss.Text = IIf(IsNull(rstRecibosProvi!RetSuss), "", rstRecibosProvi!RetSuss)
            ComproGanan.Text = IIf(IsNull(rstRecibosProvi!ComproGanan), "", rstRecibosProvi!ComproGanan)
            ComproIva.Text = IIf(IsNull(rstRecibosProvi!ComproIva), "", rstRecibosProvi!ComproIva)
            ComproIB.Text = IIf(IsNull(rstRecibosProvi!ComproIB), "", rstRecibosProvi!ComproIB)
            ComproSuss.Text = IIf(IsNull(rstRecibosProvi!ComproSuss), "", rstRecibosProvi!ComproSuss)
            Tipo1.Value = True
            Tipo2.Value = False
            Select Case Val(rstRecibosProvi!TipoRec)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case Else
            End Select
            rstRecibosProvi.Close
        End If
        
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
                Else
            Fecha.SetFocus
        End If
        
        End If
        
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Clientes.SetFocus
                Else
            G$ = "Formato de fecha invalido"
            A% = MsgBox(G$, 0, "Emision de Recibos")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Clientes.Text <> "" Then
            With rstClientes
                spClientes = "ConsultaClientes " + "'" + Clientes.Text + "'"
                Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                If rstClientes.RecordCount > 0 Then
                    Clientes.Text = rstClientes!Cliente
                    DesClientes.Caption = rstClientes!Razon
                    WRazon = rstClientes!Razon
                    WDireccion = rstClientes!Direccion
                    WLocalidad = rstClientes!Localidad
                    WPostal = rstClientes!Postal
                    WProvincia = Provincia(rstClientes!Provincia)
                    WProv = rstClientes!Provincia
                    Rem Call Imprime_Datos
                    Observaciones.SetFocus
                    rstClientes.Close
                        Else
                    Clientes.SetFocus
                End If
            End With
        End If
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.SetFocus
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
        Call Suma_Datos
        If Val(Retganancias.Text) <> 0 Then
            EntraComproGanan.Visible = True
            ComproGanan.SetFocus
                Else
            RetIva.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproGanan_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproGanan.Visible = False
        RetIva.SetFocus
    End If
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.Text = Alinea("###,###.##", RetIva.Text)
        Call Suma_Datos
        If Val(RetIva.Text) <> 0 Then
            EntraComproIva.Visible = True
            ComproIva.SetFocus
                Else
            RetOtra.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproIva.Visible = False
        RetOtra.SetFocus
    End If
End Sub

Private Sub RetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
        Call Suma_Datos
        If Val(RetOtra.Text) <> 0 Then
            EntraComproIb.Visible = True
            ComproIB.SetFocus
                Else
            RetSuss.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproIb.Visible = False
        RetSuss.SetFocus
    End If
End Sub

Private Sub RetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
        Call Suma_Datos
        If Val(RetSuss.Text) <> 0 Then
            EntraComproSuss.Visible = True
            ComproSuss.SetFocus
                Else
            TotalRecibo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub totalrecibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TotalRecibo.Text = Alinea("###,###.##", TotalRecibo.Text)
        If Val(Wempresa) = 1 Then
            DbGrid1.Col = 5
                Else
            DbGrid1.Col = 0
        End If
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub ComproSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproSuss.Visible = False
        If Val(Wempresa) = 1 Then
            DbGrid1.Col = 5
                Else
            DbGrid1.Col = 0
        End If
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IngresaCuit.Visible = False
        ZClaveCheque(DbGrid1.Row, 6) = Cuit.Text
        DbGrid1.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    XRow = DbGrid1.Row
    XCol = DbGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Cuenta Corrientes"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
    Ayuda.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spCliente = "ListaCliente"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCliente.Close
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 1
        
            XParam = "'" + Clientes.Text + "','" _
                        + Clientes.Text + "'"
            spCtacte = "ListaCtacteDesdeHasta" + XParam
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
            
            With rstCtacte
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstCtacte!Saldo <> 0 Then
                            Auxi = Str$(rstCtacte!Saldo)
                            Auxi = Mascara("###,###.##", Auxi$)
                            Auxi1 = Str$(rstCtacte!Numero)
                            Call Ceros(Auxi1, 6)
                            IngresaItem = rstCtacte!Impre + " " + Auxi1 + " " + rstCtacte!Fecha + " " + Auxi
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCtacte!Clave
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtacte.Close
            
            End If
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        If Ayuda.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Razon"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
        End If
    
    End If

End Sub


Private Sub Pantalla_Click()
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Clientes.Text = WCliente
                DesClientes.Caption = rstCliente!Razon
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WProvincia = Provincia(rstCliente!Provincia)
                WProv = rstCliente!Provincia
                                Else
                Clientes.Text = ""
            End If
            
            Pantalla.Visible = False
            Clientes.SetFocus
            
        Case 1
        
            If Tipo1.Value = True Then
        
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 0 To 19
                DbGrid1.Row = iRow
                DbGrid1.Col = 0
                Compara2 = DbGrid1.Text
                DbGrid1.Col = 3
                Compara2 = Compara2 + DbGrid1.Text + "01"
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
            For iRow = 0 To 19
                DbGrid1.Row = iRow
                DbGrid1.Col = 0
                If DbGrid1.Text = "" Then
                    XRow = DbGrid1.Row
                    Exit For
                End If
            Next iRow
            
            Indice = Pantalla.ListIndex
            ClaveCtacte = WIndice.List(Indice)
            spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    Auxi = rstCtacte!Tipo
                    Call Ceros(Auxi, 2)
                    DbGrid1.Text = Auxi
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 1
                    DbGrid1.Text = ""
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.Text = ""
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 3
                    Auxi = rstCtacte!Numero
                    Call Ceros(Auxi, 8)
                    DbGrid1.Text = Auxi
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    DbGrid1.Text = rstCtacte!Saldo
                    DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                    rstCtacte.Close
                    
            End If
            
            End If
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.SetFocus
            
            End If
                
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
    
            Case 0
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        DbGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
            Case 1
                Rem If KeyCode = 13 Then
                Rem     DBGrid1.Text = Left$(DBGrid1.Text, 1)
                Rem     If DBGrid1.Text = "A" Or DBGrid1.Text = "C" Then
                Rem         DBGrid1.Col = 2
                Rem         KeyCode = 0
                Rem         Rem no hago anda
                Rem             Else
                Rem         DBGrid1.Col = 1
                Rem         KeyCode = 0
                Rem     End If
                Rem End If
                
            Case 2
                Rem If KeyCode = 13 Then
                Rem     Auxi$ = Str$(Val(DBGrid1.Text))
                Rem     Call Ceros(Auxi$, 4)
                Rem     DBGrid1.Text = Auxi$
                Rem     DBGrid1.Col = 3
                Rem     KeyCode = 0
                Rem End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                    
                    With rstCtacte
                    
                    
                        DbGrid1.Col = 0
                        XTipo = DbGrid1.Text
                        
                        DbGrid1.Col = 3
                        XNumero = DbGrid1.Text
                        
                        ClaveCtacte = XTipo + XNumero + "01"
                        
                        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtacte.RecordCount > 0 Then
                        
                            DbGrid1.Col = 4
                            XRow = DbGrid1.Row
                            If Val(DbGrid1.Text) = 0 Then
                                DbGrid1.Text = !Saldo
                                Call Suma_Datos
                                DbGrid1.Col = 4
                                DbGrid1.Row = XRow
                            End If
                            DbGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                            
                            rstCtacte.Close
                            
                        End If
                    End With
                End If
                
            Case 4
                If KeyCode = 13 Then
                    With rstCtacte
                        DbGrid1.Col = 0
                        XTipo = DbGrid1.Text
                        DbGrid1.Col = 3
                        XNumero = DbGrid1.Text
                        
                        ClaveCtacte = XTipo + XNumero + "01"
                        
                        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtacte.RecordCount > 0 Then
                            Saldo = Alinea("###,###.##", Str$(rstCtacte!Saldo))
                            rstCtacte.Close
                                Else
                            Saldo = 0
                        End If
                    
                    End With
                
                    DbGrid1.Col = 4
                    If Abs(Val(DbGrid1.Text)) > Abs(Val(Saldo)) Then
                        DbGrid1.Text = ""
                        DbGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        If DbGrid1.Row < 10 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 0
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Val(TotalRecibo.Text) <> 0 Then
                    
                        If Len(DbGrid1.Text) = 31 Then
                            Lectora.Text = DbGrid1.Text
                            Call Lectora_Keypress(13)
                                Else
                            If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 4 Or Val(DbGrid1.Text) = 99 Then
                                Auxi$ = Str$(Val(DbGrid1.Text))
                                Call Ceros(Auxi$, 2)
                                DbGrid1.Text = Auxi$
                                Select Case Val(DbGrid1.Text)
                                    Case 1, 4
                                        DbGrid1.Col = 6
                                        DbGrid1.Text = ""
                                        DbGrid1.Col = 7
                                        DbGrid1.Text = ""
                                        DbGrid1.Col = 8
                                        DbGrid1.Text = ""
                                        DbGrid1.Col = 9
                                        KeyCode = 0
                                    Case Else
                                        DbGrid1.Col = 6
                                        KeyCode = 0
                                End Select
                                    Else
                                DbGrid1.Col = 5
                                KeyCode = 0
                            End If
                        End If
                        
                            Else
                            
                        m1$ = "Se debe informar el total del recibo antes de informar los valores"
                        A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                        DbGrid1.Col = 5
                        KeyCode = 0
                        
                    End If
                    
                End If
                
            Case 6
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                    DbGrid1.Col = 7
                    KeyCode = 0
                End If
                
            Case 7
                If KeyCode = 13 Then
                
                    If Len(DbGrid1.Text) = 5 Then
                        If Val(Right$(DbGrid1.Text, 2)) > 6 Then
                            DbGrid1.Text = DbGrid1.Text + "/2016"
                                Else
                            DbGrid1.Text = DbGrid1.Text + "/2016"
                        End If
                    End If
                    
                    DbGrid1.Col = 7
                    Call Valida_fecha1(DbGrid1.Text, Auxi)
                    Rem Call Valida_fecha(DbGrid1.Text, Auxi)
                    
                    If Auxi <> "S" Then
                    
                        DbGrid1.Col = 7
                        KeyCode = 0
                        
                                Else
                                
                        ZPasa = ""
                        ZFecha = DbGrid1.Text
                        DbGrid1.Col = 5
                        ZTipo = Val(DbGrid1.Text)
        
                        WDias = 0
                        WFechaDesde = ZFecha
                        WFechaHasta = Fecha.Text
        
                        WOrdFechaDesde = Right$(WFechaDesde, 4) + Mid$(WFechaDesde, 4, 2) + Left$(WFechaDesde, 2)
                        WOrdFechaHasta = Right$(WFechaHasta, 4) + Mid$(WFechaHasta, 4, 2) + Left$(WFechaHasta, 2)
        
                        If ZTipo = 2 And WOrdFechaDesde < WOrdFechaHasta Then
        
                            Do
                                WDias = WDias + 1
                                XFec1 = WFechaDesde
                                SumaDia = 2
                                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                                WFechaDesde = XFec2
                                If WFechaDesde = WFechaHasta Then
                                    Exit Do
                                End If
                            Loop
            
                            If WDias > 30 Then
                                ZPasa = "N"
                            End If
            
                        End If
                        
                        If ZPasa = "N" Then
                            m1$ = "Error en la carga de fecha de cheque"
                            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                            DbGrid1.Col = 7
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 8
                            If Trim(DbGrid1.Text) <> "" Then
                                DbGrid1.Col = 9
                            End If
                            KeyCode = 0
                        End If
                    
                    End If
                End If
                
            Case 8
                If KeyCode = 13 Then
                    ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
                    ZAgrega = Left$(Clientes.Text, 1) + ZSuma
                    ZLong = Len(ZAgrega)
                    If Right$(DbGrid1.Text, ZLong) <> ZAgrega Then
                        DbGrid1.Text = DbGrid1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
                    End If
                    DbGrid1.Col = 9
                    KeyCode = 0
                End If
                
            Case 9
                If KeyCode = 13 Then
                    If Val(TotalRecibo.Text) <> 0 Then
                    
                        iRow = DbGrid1.Row
                        DbGrid1.Col = 5
                        XTipo = DbGrid1.Text
                        DbGrid1.Col = 9
                        DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        DbGrid1.Row = iRow
                        
                        If Val(XTipo) = 4 Then
                            Cuenta1.Text = WCuenta(DbGrid1.Row)
                            Ingrecuenta.Visible = True
                            Cuenta1.SetFocus
                        End If
                        
                        ZZCuit = ZClaveCheque(DbGrid1.Row + 1, 6)
                        If Val(XTipo) = 2 And ZZCuit = "" Then
                            Cuit.Text = ""
                            IngresaCuit.Visible = True
                            Cuit.SetFocus
                        End If
                        
                        If DbGrid1.Row < 20 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 5
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 5
                            KeyCode = 0
                        End If
                        
                            Else
                            
                        m1$ = "Se debe informar el total del recibo antes de informar los valores"
                        A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                        DbGrid1.Col = 9
                        KeyCode = 0
                        
                    End If
                        
                        
                End If

            Case Else
                
    End Select
                        
    ZZDa = Len(DbGrid1.Text)
    If Len(DbGrid1.Text) = 30 And UCase(Left$(DbGrid1.Text, 1)) = "C" Then
    
        Lectora.Text = "c" + Mid$(DbGrid1.Text, 2, 29) + "e"
        
        ZEntra = "S"
    
        Sql1 = "Select *"
        Sql2 = " FROM Recibos"
        Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
        spRecibos = Sql1 + Sql2 + Sql3
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            m1$ = "Cheque ya cargado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            ZEntra = "N"
            rstRecibos.Close
        End If
    
        Sql1 = "Select *"
        Sql2 = " FROM RecibosProvi"
        Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
        spRecibosProvi = Sql1 + Sql2 + Sql3
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            m1$ = "Cheque ya cargado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            ZEntra = "N"
            rstRecibosProvi.Close
        End If
        
        If ZEntra = "S" Then
            For ZZCiclo = 1 To 100
                If ZClaveCheque(ZZCiclo, 1) = Lectora.Text Then
                    m1$ = "Cheque ya cargado"
                    A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                    ZEntra = "N"
                End If
            Next ZZCiclo
        End If
        
        If ZEntra = "S" Then
    
            ZNombreBanco = ZBancos(Val(Mid$(Lectora.Text, 2, 3)))
            ZNroCuenta = Mid$(Lectora.Text, 12, 8)
        
            ZZBanco = Mid$(Lectora.Text, 2, 3)
            ZZSucursal = Mid$(Lectora.Text, 5, 3)
            ZZNroCheque = Mid$(Lectora.Text, 12, 8)
            ZZNroCuenta = Mid$(Lectora.Text, 20, 11)

            DbGrid1.Col = 5
            DbGrid1.Text = "02"
            DbGrid1.Col = 8
            DbGrid1.Text = ZNombreBanco
            ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
            DbGrid1.Text = DbGrid1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
            DbGrid1.Col = 6
            DbGrid1.Text = ZNroCuenta
            DbGrid1.Col = 7
            DbGrid1.Text = ""
            Rem DbGrid1.Col = 6
            KeyCode = 0
        
            ZClaveCheque(DbGrid1.Row + 1, 1) = Lectora.Text
            ZClaveCheque(DbGrid1.Row + 1, 2) = ZZBanco
            ZClaveCheque(DbGrid1.Row + 1, 3) = ZZSucursal
            ZClaveCheque(DbGrid1.Row + 1, 4) = ZZNroCheque
            ZClaveCheque(DbGrid1.Row + 1, 5) = ZZNroCuenta
            ZClaveCheque(DbGrid1.Row + 1, 6) = ""
        
            ZZClave = ZZBanco + ZZSucursal + ZZNroCuenta
            ZZCuit = ""
        
            ZSql = "Select *"
            ZSql = ZSql + " FROM Cuit"
            ZSql = ZSql + " Where Cuit.Clave = " + "'" + ZZClave + "'"
            spCuit = ZSql
            Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuit.RecordCount > 0 Then
                ZZCuit = Trim(rstCuit!Cuit)
                rstCuit.Close
            End If
            
            ZClaveCheque(DbGrid1.Row + 1, 6) = ZZCuit
            Toto.Text = ""
            Toto.SetFocus
            
                Else
                
            DbGrid1.Col = 5
            DbGrid1.Text = ""
            DbGrid1.Col = 4
            Lectora.Visible = False
            DbGrid1.SetFocus
            
        End If
        
    End If
    
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
        UserData(iCol, mTotalRows - 1) = DbGrid1.Columns(iCol).DefaultValue
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

ReDim UserData(0 To 9, 0 To 19)

mTotalRows& = 20

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DbGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DbGrid1.Columns.Count - 1 To 0 Step -1
     DbGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 9
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Letra"
             DbGrid1.Columns(newcnt).Width = 450
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Punto"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Numero/Cta"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1300
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 8
             DbGrid1.Columns(newcnt).Caption = "Banco"
             DbGrid1.Columns(newcnt).Width = 1500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 9
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 
    Provincia$(0) = "Capital Federal"
    Provincia$(1) = "Buenos Aires"
    Provincia$(2) = "Catamarca"
    Provincia$(3) = "Cordoba"
    Provincia$(4) = "Corrientes"
    Provincia$(5) = "Chaco"
    Provincia$(6) = "Chubut"
    Provincia$(7) = "Entre Rios"
    Provincia$(8) = "Formosa"
    Provincia$(9) = "Jujuy"
    Provincia$(10) = "La Pampa"
    Provincia$(11) = "La Rioja"
    Provincia$(12) = "Mendoza"
    Provincia$(13) = "Misiones"
    Provincia$(14) = "Neuquen"
    Provincia$(15) = "Rio Negro"
    Provincia$(16) = "Salta"
    Provincia$(17) = "San Juan"
    Provincia$(18) = "San Luis"
    Provincia$(19) = "Santa Cruz"
    Provincia$(20) = "Santa Fe"
    Provincia$(21) = "Santiago del Estero"
    Provincia$(22) = "Tucuman"
    Provincia$(23) = "Tierra del Fuego"
    Provincia$(24) = "Exterior"
    Provincia$(25) = ""
    
    ZBancos(3) = "BEAL"
    ZBancos(5) = "AMRO BANK"
    ZBancos(7) = "GALICIA"
    ZBancos(10) = "LLOYDS BANK"
    ZBancos(11) = "NACION"
    ZBancos(14) = "PROVINCIA"
    ZBancos(15) = "BANKBOSTON"
    ZBancos(16) = "CITIBANK"
    ZBancos(17) = "FRANCES"
    ZBancos(18) = "TOKYO"
    ZBancos(20) = "CORDOBA"
    ZBancos(27) = "SUPERVIELLE"
    ZBancos(29) = "CIUDAD"
    ZBancos(30) = "CENTRAL"
    ZBancos(34) = "PATAGONIA"
    ZBancos(44) = "HIPOTECARIO"
    ZBancos(45) = "SAN JUAN"
    ZBancos(46) = "BRASIL"
    ZBancos(60) = "TUCUMAN"
    ZBancos(65) = "ROSARIO"
    ZBancos(72) = "RIO"
    ZBancos(79) = "CUYO"
    ZBancos(83) = "CHUBUT"
    ZBancos(86) = "SANTA CRUZ"
    ZBancos(93) = "LA PAMPA"
    ZBancos(94) = "CORRIENTES "
    ZBancos(97) = "NEUQUEN"
    ZBancos(137) = "EMP.TUCUMAN"
    ZBancos(147) = "B.I.CRED."
    ZBancos(148) = "LA PLATA"
    ZBancos(150) = "HSBC"
    ZBancos(165) = "JPMORGAN"
    ZBancos(191) = "CREDICOOP"
    ZBancos(198) = "VALORES"
    ZBancos(247) = "ROELA"
    ZBancos(254) = "MARIVA"
    ZBancos(259) = "ITAU"
    ZBancos(265) = "HSBC"
    ZBancos(262) = "OF AMERICA"
    ZBancos(266) = "BNP PARIBAS"
    ZBancos(268) = "T.FUEGO"
    ZBancos(269) = "URUGUAY"
    ZBancos(277) = "SAENZ"
    ZBancos(281) = "MERIDIAN"
    ZBancos(285) = "MACRO"
    ZBancos(293) = "MERCURIO"
    ZBancos(294) = "ING.BANK"
    ZBancos(295) = "AMERICAN"
    ZBancos(297) = "BANEX"
    ZBancos(299) = "COMAFI"
    ZBancos(300) = "INVERSION"
    ZBancos(301) = "PIANO"
    ZBancos(303) = "FINANSUR"
    ZBancos(305) = "JULIO"
    ZBancos(306) = "P.INVERSIONES"
    ZBancos(309) = "LA RIOJA"
    ZBancos(310) = "DEL SOL"
    ZBancos(311) = "CHACO"
    ZBancos(312) = "DE INVERSIONES"
    ZBancos(315) = "FORMOSA"
    ZBancos(319) = "CMF"
    ZBancos(320) = "BANEX"
    ZBancos(321) = "S.ESTERO"
    ZBancos(322) = "IND.AZUL"
    ZBancos(325) = "DEUTSCHE BANK"
    ZBancos(330) = "SANTA FE"
    ZBancos(331) = "CETELEM"
    ZBancos(332) = "SERV.FINAN."
    ZBancos(335) = "COFIDIS"
    ZBancos(336) = "BRADESCO"
    ZBancos(338) = "SERV.Y TRANS."
    ZBancos(339) = "RCI BANQUE"
    ZBancos(340) = "DE CREDITO"
    ZBancos(386) = "ENTRE RIOS"
    ZBancos(387) = "SUQUIA"
    ZBancos(388) = "BISEL"
    ZBancos(389) = "COLUMBIA"
     
    ImpreTipo$(1) = "FC"
     
    Tipo1.Value = True
    Tipo2.Value = False
    
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    
    ComproGanan.Text = ""
    ComproIva.Text = ""
    ComproIB.Text = ""
    ComproSuss.Text = ""

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    Paridad.Text = ""
    TotalRecibo.Text = ""
    
    cmdAdd.Enabled = True
    
    Erase ZClaveCheque
    
    Recibo.Text = ""
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM RecibosProvi"
    Rem spRecibosProvi = ZSql
    Rem Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibosProvi.RecordCount > 0 Then
    Rem     rstRecibosProvi.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibosProvi!ReciboMayor), "0", rstRecibosProvi!ReciboMayor)
    Rem     Recibo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
    Rem     rstRecibosProvi.Close
    Rem         Else
    Rem     Recibo.Text = "1"
    Rem End If
    
End Sub

Private Sub Cuenta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta1.Text <> "" Then
            spCuenta = "ConsultaCuentas " + "'" + Cuenta1.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                    WCuenta(DbGrid1.Row - 1) = Cuenta1.Text
                    Ingrecuenta.Visible = False
                    Rem DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 5
                    KeyCode = 0
                    DbGrid1.SetFocus
                    rstCuenta.Close
                        Else
                    Cuenta.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Lectora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Lectora.Text) = 31 Then
        
            ZEntra = "S"
        
            Sql1 = "Select *"
            Sql2 = " FROM Recibos"
            Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
            spRecibos = Sql1 + Sql2 + Sql3
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                m1$ = "Cheque ya cargado"
                A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                ZEntra = "N"
                rstRecibos.Close
            End If
        
            Sql1 = "Select *"
            Sql2 = " FROM RecibosProvi"
            Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
            spRecibosProvi = Sql1 + Sql2 + Sql3
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibosProvi.RecordCount > 0 Then
                m1$ = "Cheque ya cargado"
                A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                ZEntra = "N"
                rstRecibosProvi.Close
            End If
            
            If ZEntra = "S" Then
                For ZZCiclo = 1 To 100
                    If ZClaveCheque(ZZCiclo, 1) = Lectora.Text Then
                        m1$ = "Cheque ya cargado"
                        A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                        ZEntra = "N"
                    End If
                Next ZZCiclo
            End If
            
            If ZEntra = "S" Then
                If Trim(ZBancos(Val(Mid$(Lectora, 2, 3)))) = "" Then
                    m1$ = "error en la lectura de los datos del codigo de banco del cheque"
                    A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                    ZEntra = "N"
                End If
            End If
            
            
            If ZEntra = "S" Then
        
                ZNombreBanco = ZBancos(Val(Mid$(Lectora, 2, 3)))
                ZNroCuenta = Mid$(Lectora, 12, 8)
            
                ZZBanco = Mid$(Lectora, 2, 3)
                ZZSucursal = Mid$(Lectora, 5, 3)
                ZZNroCheque = Mid$(Lectora, 12, 8)
                ZZNroCuenta = Mid$(Lectora, 20, 11)

                DbGrid1.Col = 5
                DbGrid1.Text = "02"
                DbGrid1.Col = 8
                DbGrid1.Text = ZNombreBanco
                ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
                DbGrid1.Text = DbGrid1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
                DbGrid1.Col = 6
                DbGrid1.Text = ZNroCuenta
                DbGrid1.Col = 7
                DbGrid1.Text = ""
                Rem DbGrid1.Col = 6
                KeyCode = 0
            
                ZClaveCheque(DbGrid1.Row + 1, 1) = Lectora.Text
                ZClaveCheque(DbGrid1.Row + 1, 2) = ZZBanco
                ZClaveCheque(DbGrid1.Row + 1, 3) = ZZSucursal
                ZClaveCheque(DbGrid1.Row + 1, 4) = ZZNroCheque
                ZClaveCheque(DbGrid1.Row + 1, 5) = ZZNroCuenta
                ZClaveCheque(DbGrid1.Row + 1, 6) = ""
            
                ZZClave = ZZBanco + ZZSucursal + ZZNroCuenta
                ZZCuit = ""
            
                ZSql = "Select *"
                ZSql = ZSql + " FROM Cuit"
                ZSql = ZSql + " Where Cuit.Clave = " + "'" + ZZClave + "'"
                spCuit = ZSql
                Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuit.RecordCount > 0 Then
                    ZZCuit = Trim(rstCuit!Cuit)
                    rstCuit.Close
                End If
                
                ZClaveCheque(DbGrid1.Row + 1, 6) = ZZCuit
                Lectora.Visible = False
                DbGrid1.SetFocus
                
                    Else
                    
                DbGrid1.Col = 5
                DbGrid1.Text = ""
                DbGrid1.Col = 4
                Lectora.Visible = False
                DbGrid1.SetFocus
                
            End If
                    
            
                Else
                
            DbGrid1.Col = 5
            DbGrid1.Text = ""
            KeyCode = 0
            
            DbGrid1.SetFocus
            Lectora.Visible = False
            
        End If
    End If
End Sub

Private Sub Toto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 69 Or KeyAscii = 101 Then
        DbGrid1.SetFocus
    End If
End Sub

