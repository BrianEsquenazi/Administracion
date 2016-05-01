VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Ordenes de Compra"
   ClientHeight    =   7560
   ClientLeft      =   690
   ClientTop       =   1125
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.TextBox Dada 
      Height          =   285
      Left            =   10080
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Tarjeta 
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
      Left            =   9840
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame DatosImpo 
      Height          =   1815
      Left            =   0
      TabIndex        =   15
      Top             =   4800
      Width           =   11655
      Begin VB.TextBox DJai 
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
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Flete 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox TipoPago 
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
         Left            =   8640
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox PedidoImpo 
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
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Origen 
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
         MaxLength       =   50
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Leyenda 
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
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox TipoImpo 
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
         Left            =   6600
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin MSMask.MaskEdBox FechaImpo 
         Height          =   285
         Left            =   5160
         TabIndex        =   20
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox FechaDJai 
         Height          =   285
         Left            =   2760
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label5 
         Caption         =   "FechaDJai"
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
         Left            =   2760
         TabIndex        =   37
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "DJai"
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
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Flete U$S"
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
         Left            =   10440
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Tipo Pago"
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
         Left            =   8640
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Condicion"
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
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Via"
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
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label27 
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
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label26 
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
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Origen"
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
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Carpeta 
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox TipoOrden 
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
      Left            =   9840
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Moneda 
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
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   8
      Text            =   " "
      Top             =   480
      Width           =   1455
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
      Height          =   500
      Left            =   4080
      TabIndex        =   7
      Top             =   6840
      Width           =   4575
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9480
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid WVector 
      Height          =   3735
      Left            =   0
      TabIndex        =   33
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      _Version        =   327680
      Rows            =   4000
      Cols            =   6
   End
   Begin VB.Label Label34 
      Caption         =   "Forma Pago"
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
      Left            =   8760
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label24 
      Caption         =   "Carpeta"
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
      Left            =   5880
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Tipo Orden"
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
      Left            =   8760
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Moneda"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label DesProveedor 
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
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      TabIndex        =   5
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
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Compra"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgOrdenConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private XMoneda As Integer
Private WMonedaOrden As String
Private WTipoOrden As String
Private WTipoPago As String
Private Cantidad As Single
Private WOrdenprecio As String
Private XPrecio As String
Private XCantidad As String

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim rstNroCarpeta As Recordset
Dim spNroCarpeta As String

Dim XParam As String
Dim ZEmail As String
Dim Vector(100, 10) As String
Dim ZVector(100, 10) As String
Dim XPorceDerechos(100) As String
Dim XSolicitud(100, 3) As String
Dim ZZSolicitud(5000, 7) As String
Dim CargaEmpresa(12, 2) As String
Dim Empe(12, 10) As String

Dim ZZArtiCerti(100) As String
Dim ZEnsayo(100) As String
Dim ZDescriII(100) As String

Private TipoConsulta As String
Private XVector(10, 5) As String
Private Auxi As String
Private WAuxi As String
Private WSaldo As Double
Private Desdelugar As Integer
Dim Tabla(10000) As String
Private WEntre As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim Paridad As Double
Dim ParidadII As Double
Dim WDerechos As String
Dim XDerechos As Double
Dim ZDespacho As Double
Dim ZEntra As String
Dim ZSolicitud As String
Dim ZZArticulo As String

Dim ZZFechaLlegada As String
Dim ZZPagoDespacho As String
Dim ZZImpoDespacho As String
Dim ZZVtoDespacho As String
Dim ZZPagoLetra As String
Dim ZZImpoLetra As String
Dim ZZVtoLetra As String

Dim ZZCosto As Double
Dim ZParidad As Double
Dim ZParidadII As Double
Dim p As Object
Dim ZCoeParidad As Double


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

Dim ZZPasaTo As String
Dim ZZPasaCC As String
Dim ZZPasaBody As String
Dim ZZPasaFile As String



Private Sub cmdClose_Click()

    XEmpresa = WEmpresa
    Select Case WPasaEmpresa
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
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select




    Auxi = Orden.Text
    Call Ceros(Auxi, 6)
    WClave = Auxi + "01"
        
    spOrden = "ConsultaOrden " + "'" + WClave + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        ZZOrigen = IIf(IsNull(rstOrden!Origen), "", rstOrden!Origen)
        ZZLeyenda = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
        ZZPedidoImpo = IIf(IsNull(rstOrden!PedidoImpo), "", rstOrden!PedidoImpo)
        ZZFechaImpo = IIf(IsNull(rstOrden!FechaImpo), "  /  /    ", rstOrden!FechaImpo)
        ZZTipoImpo = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
        ZZTipoPago = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
        ZZFlete = IIf(IsNull(rstOrden!Flete), "", rstOrden!Flete)
        ZZDJai = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
        ZZDJai = Trim(ZZDJai)
        ZZFechaDJai = IIf(IsNull(rstOrden!FechaDJai), "  /  /    ", rstOrden!FechaDJai)
        rstOrden.Close
    End If
    
    
    
    
    ZOrigen = Origen.Text
    ZLeyenda = Leyenda.ListIndex
    ZPedidoImpo = PedidoImpo.Text
    ZFechaImpo = FechaImpo.Text
    ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
    ZTipoImpo = TipoImpo.ListIndex
    ZTipoPago = TipoPago.ListIndex
    ZFlete = Flete.Text
    ZDJai = DJai.Text
    ZFechaDJai = FechaDJai.Text
    
    
    ZZGraba = "N"
    
    If Trim(ZZOrigen) <> Trim(ZOrigen) Then
        ZZGraba = "S"
    End If
    If ZZLeyenda <> ZLeyenda Then
        ZZGraba = "S"
    End If
    If Trim(ZZPedidoImpo) <> Trim(ZPedidoImpo) Then
        ZZGraba = "S"
    End If
    If ZZFechaImpo <> ZFechaImpo Then
        ZZGraba = "S"
    End If
    If ZZTipoImpo <> ZTipoImpo Then
        ZZGraba = "S"
    End If
    If ZZTipoPago <> ZTipoPago Then
        ZZGraba = "S"
    End If
    If ZZFlete <> Val(ZFlete) Then
        ZZGraba = "S"
    End If
    If ZZDJai <> ZDJai Then
        ZZGraba = "S"
    End If
    If ZZFechaDJai <> ZFechaDJai Then
        ZZGraba = "S"
    End If
    
    
    If ZZGraba = "S" Then
        T$ = "Actualizacion de Orden de Compra"
        m$ = "Desea Guardar los cambios a la Orden de Compra "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Orden SET "
            ZSql = ZSql + " DJai = " + "'" + ZDJai + "',"
            ZSql = ZSql + " FechaDJai = " + "'" + ZFechaDJai + "',"
            ZSql = ZSql + " Flete = " + "'" + ZFlete + "',"
            ZSql = ZSql + " Origen = " + "'" + ZOrigen + "',"
            ZSql = ZSql + " Leyenda = " + "'" + Str$(ZLeyenda) + "',"
            ZSql = ZSql + " PedidoImpo = " + "'" + ZPedidoImpo + "',"
            ZSql = ZSql + " FechaImpo = " + "'" + ZFechaImpo + "',"
            ZSql = ZSql + " OrdFechaImpo = " + "'" + ZOrdFechaImpo + "',"
            ZSql = ZSql + " TipoImpo = " + "'" + Str$(ZTipoImpo) + "',"
            ZSql = ZSql + " TipoPago = " + "'" + Str$(ZTipoPago) + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
               
        End If
        
    End If
    
    Call Conecta_Empresa
    
    PrgOrdenConsulta.Hide
End Sub


Private Sub Form_Activate()

    Orden.Text = WPasaOrden
    Call Orden_KeyPress(13)
    
End Sub


Private Sub Form_Load()

    Call Limpia_Vector

    TipoImpo.Clear
    
    TipoImpo.AddItem ""
    TipoImpo.AddItem "Maritimo"
    TipoImpo.AddItem "Terrestre"
    TipoImpo.AddItem "Areo"
    
    TipoImpo.ListIndex = 0
 
    Tarjeta.Clear
    
    Tarjeta.AddItem ""
    Tarjeta.AddItem "PymeNacion"
    
    Tarjeta.ListIndex = 0
 
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    ZEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    Flete.Text = ""
    
    TipoImpo.ListIndex = 0
    
    Leyenda.Clear
    
    Leyenda.AddItem ""
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CIF"
    Leyenda.AddItem "CFR"
    Leyenda.AddItem "CPT"
    Leyenda.AddItem "EXW"
    Leyenda.AddItem "FCA"
    
    Leyenda.ListIndex = 0
 
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Euros"
    Moneda.AddItem ""

    Moneda.ListIndex = 0
    
    TipoOrden.Clear
    
    TipoOrden.AddItem "Local"
    TipoOrden.AddItem "Importacion"
    TipoOrden.AddItem "Prestamo"
    TipoOrden.AddItem "Envases y Filtros"
    TipoOrden.AddItem "Drogas Lab."

    TipoOrden.ListIndex = 0
    
    TipoPago.Clear
    
    TipoPago.AddItem ""
    TipoPago.AddItem "Pago Anticipado"
    TipoPago.AddItem "A la vista"
    TipoPago.AddItem "Cuenta Corriente"

    TipoPago.ListIndex = 0
 
    Orden.Text = WPasaOrden
    Call Orden_KeyPress(13)

    WVector.Col = 1
    WVector.Row = 1
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    Renglon = 0
    Erase Vector
    Erase ZZArtiCerti
    
    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
            
                WVector.Row = Renglon
                
                WVector.Col = 1
                WVector.Text = rstOrden!Articulo
                Auxi1 = rstOrden!Articulo
                
                WVector.Col = 3
                WVector.Text = Pusing("###,###.##", rstOrden!Cantidad)
                
                WVector.Col = 4
                WVector.Text = Pusing("###,###.##", rstOrden!Precio)
                
                WLugar = WVector.Row
                        
                XPorceDerechos(WLugar) = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                
                XSolicitud(WLugar, 1) = IIf(IsNull(rstOrden!Solicitud1), "0", rstOrden!Solicitud1)
                XSolicitud(WLugar, 2) = IIf(IsNull(rstOrden!Solicitud2), "0", rstOrden!Solicitud2)
                XSolicitud(WLugar, 3) = IIf(IsNull(rstOrden!Solicitud3), "0", rstOrden!Solicitud3)
                
                ZZArtiCerti(WLugar) = rstOrden!Articulo
                
                If rstOrden!Recibida > 0 Then
                 Rem   Graba.Enabled = False
                End If
                
                Rem ZMarca = IIf(IsNull(rstOrden!Marca), "", rstOrden!Marca)
                Rem If ZMarca = "X" Then
                Rem     Graba.Enabled = False
                Rem End If
                
                Vector(Renglon, 1) = Auxi1
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    
    WRenglon = Renglon
    Renglon = 0
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
        Auxi1 = Vector(Renglon, 1)
    
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector.TextMatrix(Renglon, 2) = rstArticulo!Descripcion
            WDerechos = IIf(IsNull(rstArticulo!Derechos), "0", rstArticulo!Derechos)
            WVector.TextMatrix(Renglon, 5) = IIf(IsNull(rstArticulo!Posarance), "0", rstArticulo!Posarance)
            WVector.TextMatrix(Renglon, 6) = Str$(WDerechos)
            WVector.TextMatrix(Renglon, 6) = Pusing("###,###.##", WVector.TextMatrix(Renglon, 6))
            rstArticulo.Close
        End If
    
    Next DA
    
    Call Conecta_Empresa
    
    WVector.Col = 1
    WVector.Row = 1
    
End Sub


    
    

Private Sub Limpia_Vector()

    Rem wvector.Height = 4095
    Rem wvector.Left = 120
    Rem wvector.Top = 1200
    Rem wvector.Width = 10000

    WVector.Clear
    WVector.Font.Bold = True
    
    WVector.FixedCols = 1
    WVector.Cols = 7
    WVector.FixedRows = 1
    WVector.Rows = 100
    
    WVector.ColWidth(0) = 200
    WVector.Row = 0
    For Ciclo = 1 To WVector.Cols - 1
        WVector.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector.Text = "Producto"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector.Text = "Descripcion"
                WVector.ColWidth(Ciclo) = 3500
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector.Text = "Cantidad"
                WVector.ColWidth(Ciclo) = 1000
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector.Text = "Precio"
                WVector.ColWidth(Ciclo) = 1000
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector.Text = "Posicion Arancelaria"
                WVector.ColWidth(Ciclo) = 2000
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                WVector.Text = "% Derechos"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector.Cols - 1
        WAncho = WAncho + WVector.ColWidth(Ciclo)
    Next Ciclo
    WVector.Width = WAncho

    ' Size the columns.
    Font.Name = WVector.Font.Name
    Font.Size = WVector.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector.AllowUserResizing = flexResizeBoth
    
    WVector.Visible = True
    
    WVector.Col = 1
    WVector.Row = 1
    Rem Call WVector_Click
    
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case WPasaEmpresa
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
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
    
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            Origen.Text = IIf(IsNull(rstOrden!Origen), "", rstOrden!Origen)
            Carpeta.Text = IIf(IsNull(rstOrden!Carpeta), "", rstOrden!Carpeta)
            Moneda.ListIndex = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            Tarjeta.ListIndex = IIf(IsNull(rstOrden!Tarjeta), "0", rstOrden!Tarjeta)
            Rem by nan quito rem a tipoOrden 5-3-2014
            TipoOrden.ListIndex = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            TipoPago.ListIndex = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
            Leyenda.ListIndex = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
            PedidoImpo.Text = IIf(IsNull(rstOrden!PedidoImpo), "", rstOrden!PedidoImpo)
            FechaImpo.Text = IIf(IsNull(rstOrden!FechaImpo), "  /  /    ", rstOrden!FechaImpo)
            TipoImpo.ListIndex = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
            Flete.Text = IIf(IsNull(rstOrden!Flete), "", rstOrden!Flete)
            Fecha.Text = rstOrden!Fecha
            Proveedor.Text = rstOrden!Proveedor
            DJai.Text = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
            DJai.Text = Trim(DJai.Text)
            FechaDJai.Text = IIf(IsNull(rstOrden!FechaDJai), "  /  /    ", rstOrden!FechaDJai)
            rstOrden.Close
            
            Call Proceso_Click
            
            Call Conecta_Empresa
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Proveedor.Text = rstProveedor!Proveedor
                DesProveedor.Caption = rstProveedor!Nombre
                ZEmail = rstProveedor!email
                rstProveedor.Close
            End If
            
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

