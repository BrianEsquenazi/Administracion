VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEti3 
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   9045
   ClientLeft      =   735
   ClientTop       =   645
   ClientWidth     =   10320
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9045
   ScaleWidth      =   10320
   Begin VB.CheckBox Check2 
      Caption         =   "Rangos"
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
      Left            =   3000
      TabIndex        =   41
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame PantaDirEntrega 
      Caption         =   "Seleccion de Lugar de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   26
      Top             =   4800
      Visible         =   0   'False
      Width           =   9375
      Begin VB.ListBox ListaDirEntrega 
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
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   9015
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9840
      Top             =   8640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "eti1.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8760
      TabIndex        =   5
      Top             =   8760
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
      Height          =   2460
      ItemData        =   "eti3.frx":0000
      Left            =   720
      List            =   "eti3.frx":0007
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7920
      TabIndex        =   3
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   9240
      TabIndex        =   2
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   9975
      Begin VB.CommandButton ImpreCaratula 
         Caption         =   "Caratula"
         Height          =   375
         Left            =   8280
         TabIndex        =   44
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox HastaNumero 
         Height          =   285
         Left            =   7200
         TabIndex        =   42
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox DesdeNumero 
         Height          =   285
         Left            =   5040
         TabIndex        =   40
         Top             =   3720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Etiqueta SGS"
         Height          =   375
         Left            =   8280
         TabIndex        =   36
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox LoteMP 
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   34
         Text            =   " "
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox tipofarma 
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
         Left            =   4080
         TabIndex        =   33
         Text            =   "Tipo"
         Top             =   2880
         Width           =   3975
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
         Left            =   6960
         MaxLength       =   6
         TabIndex        =   31
         Text            =   " "
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox TipoProceso 
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
         Left            =   3720
         TabIndex        =   30
         Text            =   " "
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox Idioma 
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
         Left            =   4680
         TabIndex        =   29
         Text            =   " "
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox DescripcionFarma 
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
         Left            =   240
         LinkTimeout     =   100
         MaxLength       =   30
         TabIndex        =   28
         Text            =   " "
         Top             =   3360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton Limpia 
         Caption         =   "  Limpia Pantalla"
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
         Left            =   8640
         TabIndex        =   25
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Baja 
         Caption         =   "  Limpia Etiquetas"
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
         Left            =   8640
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Tara 
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
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox Tipo 
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
         Left            =   4320
         TabIndex        =   21
         Text            =   "Tipo"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox Descripcion 
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
         Left            =   2280
         LinkTimeout     =   100
         MaxLength       =   100
         TabIndex        =   19
         Text            =   " "
         Top             =   1920
         Width           =   5775
      End
      Begin VB.TextBox Etiquetas 
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
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Cantidad 
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
         MaxLength       =   5
         TabIndex        =   15
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "  "
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   327680
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   8640
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   8640
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Hasta"
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
         Left            =   6120
         TabIndex        =   43
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Desde"
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
         Left            =   4200
         TabIndex        =   39
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Imprimir numero de etiqueta"
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
         Left            =   720
         TabIndex        =   38
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Lote MP"
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
         Left            =   4080
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label PedidoII 
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
         Height          =   495
         Left            =   6120
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Tara "
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
         Left            =   5040
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label DesProducto 
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
         Left            =   3000
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
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
         Left            =   3720
         TabIndex        =   17
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de Etiquetas"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Producto Terminado"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "PrgEti3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLote As String
Private WCantidad As String
Private WImpreadi As String
Private WClase As String
Private WRiesgo As String
Private WIntervencion As String
Private WNaciones As String
Private WEmbalaje As String
Private WTipoeti As String
Private WObservaciones As String
Private empece As String

Dim rstHoja As Recordset
Dim spHoja As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCertificado As Recordset
Dim spCertificado As String

Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 6) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String
Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String

Dim XParam As String
Dim Da As Integer
Dim WTerminado(100) As String
Dim LugarTerminado As Integer
Dim WConservacion As String
Dim WElaboracion As String
Dim WVencimento As String
Dim WVida As Single
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Dim XMes As String
Dim XAno As String
Dim WMes As Single
Dim WAno As Single
Dim WDirentrega As String
Dim WImpreVto As Integer
Dim ZImpreVto As Integer
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim CargaEmpresa(12, 2) As String
Dim ZLote As String
Dim ZZLote As String
Dim Empe(12, 10) As String
Dim TipoPro As String

Dim ZFechaVto As String
Dim ZVto As String
Dim ZZImpreVtoTermi As Integer
Dim ZZVidaUtil As Integer
Dim ZZZZVencimiento As String
Dim ZZZZElaboracion As String
Dim ZZZZFechaVto As String
Dim ZZZZFechaElaboracion As String
Dim ZZZZLoteOriginal As String



Dim ZEtiI As Integer
Dim ZEtiII As Integer
Dim ZDescriDirEntrega As String
Dim ZZOrdenCpa As String

Dim ZImpre(1000) As String
Dim ZImpreI(1000) As String
Dim ZImpreII(1000) As String
Dim ZImpreIII(1000) As String
Dim Pasa As String
Dim ZLugarImpre As Integer
Dim ZLugarImpreI As Integer
Dim ZLugarImpreII As Integer
Dim ZLugarImpreIII As Integer
Dim ZZLogo(100) As Integer
Dim ZZImpreFrase(100) As String



Private Sub busca_mono()

    Pasa = "N"
    
    XEmpresa = Wempresa
    
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    ZTerminado = Terminado.Text
            
    spTerminado = "Consultamono " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Pasa = "S"
        rstTerminado.Close
    End If
    
    Call Conecta_Empresa

End Sub

Private Sub Acepta_Click()

    Dim ListaSga(1000) As String
    
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
        If ListaSga(Ciclo) = UCase(Terminado.Text) Then
            m$ = "Se debe imprimir etiqueta SGS"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Exit Sub
                Else
            If ListaSga(Ciclo) = "" Then
                Exit For
            End If
        End If
    Next Ciclo



        
        Rem para imprimir nº de etiqueta
        If Check1.Value = 1 Then
            If Check2.Value = 1 Then
                ZZDesdeNumero = Val(DesdeNumero.Text)
                ZZHastaNumero = Val(HastaNumero.Text)
                    Else
                ZZDesdeNumero = 1
                ZZHastaNumero = Val(Etiquetas.Text)
            End If
        End If


    If Val(Pedido.Text) >= 376194 Then

        Rem by nan 6-10-2015
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
        
        
        
  
        
        
        
        TipoPro = "PT"
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If Left$(Terminado.Text, 2) <> "PT" Then
            Select Case Left$(Terminado.Text, 2)
                Case "DY", "DS"
                    TipoPro = "CO"
                Case "QC"
                    TipoPro = "FA"
                Case Else
                    TipoPro = "PT"
            End Select
                Else
            If XCodigo >= 0 And XCodigo <= 999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    TipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        TipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            TipoPro = "BI"
                                Else
                            TipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If
        
        If Left$(Terminado.Text, 2) = "YQ" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YH" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YP" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YF" Then
            TipoPro = "FA"
        End If
        
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                TipoPro = "FA"
            End If
        End If
        
        If TipoPro <> "CO" Then
          Rem by nan 6-10-2015 linea 6 ylinea 16 va nueva
           If WLinea = 6 Or WLinea = 16 Or WLinea = 10 Then
               m$ = "Se debe imprimir etiqueta SGS"
               G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            
            Exit Sub
           End If
        
        End If
        
    End If



    Pasa = 0
    Rem Etiquetas = Etiquetas.Text
    Rem On Error GoTo WError
    
    Rem by nan para farma
    If Wempresa = "0005" Then
    
        If tipofarma.ListIndex = 0 Then
             Tipo.ListIndex = 0
        End If
        If tipofarma.ListIndex = 1 Then
            Tipo.ListIndex = 6
        End If
        If tipofarma.ListIndex = 2 Then
            Tipo.ListIndex = 5
        End If
        If tipofarma.ListIndex = 3 Then
             Tipo.ListIndex = 3
        End If
        If tipofarma.ListIndex = 4 Then
            Tipo.ListIndex = 4
        End If
       
    End If
    Rem fin by nan
  
    
    
    
    
    
    
    
    
    
    
    Rem **********************
    
    If Tipo.ListIndex = 3 Or Tipo.ListIndex = 4 Then
        Call Imprime_Certificado
        Exit Sub
    End If
    
    If Tipo.ListIndex = 5 Then
        Call Imprime_EnProceso
        Exit Sub
    End If
    
    
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
        ZEtiI = Trim(IIf(IsNull(rstClientes!EtiI), "0", rstClientes!EtiI))
        ZEtiII = Trim(IIf(IsNull(rstClientes!EtiII), "0", rstClientes!EtiII))
        rstClientes.Close
    End If
    
    If Trim(Cliente.Text) <> "" And Val(Wempresa) = 1 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " and Pedido.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and Pedido.Terminado = " + "'" + Terminado.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            rstPedido.Close
            Pasa = 1
                
                Else
                                        
            Rem by nan para 8
            Rem   If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            Rem BY NAN BUSCO EN 8
            OPEN_FILE_Pedido2
            With rstPedido2
               Auxi = Pedido.Text
               Call Ceros(Auxi, 6)
               .Index = "Clave"
               .Seek ">=", Auxi
                If .NoMatch = False Then
                    ZZCliente = rstPedido2!Cliente
                    Pasa = 1
                        Else
                    Pasa = 0
                End If
            End With
            rstPedido2.Close
            Rem BY NAN FIN
            
        End If
            
        If Pasa = 0 Then
            Rem by nan
            Call Conecta_Empresa
            m$ = "Pedido incorrecto"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Exit Sub
        End If
        Rem   End If
            
        
        
        
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZZCliente = rstPedido!Cliente
            ZZOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
            ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            ZDescriDirEntrega = ""
            
            rstPedido.Close
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                
                    Else
                    
                Call Conecta_Empresa
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
                
            End If
            
                Else
                
            Rem by nan busco datos de entrega para 8
            OPEN_FILE_Pedido2
            With rstPedido2
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                .Index = "Clave"
                .Seek ">=", Auxi
                If .NoMatch = False Then
                    ZZCliente = rstPedido2!Cliente
                    Rem  Pasa = 1
                        Else
                    Rem    Pasa = 0
                End If
            End With
            rstPedido2.Close
                       
            Rem ***************
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
     
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                 
                        Else
            
                Call Conecta_Empresa
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
            
            
            End If
             
            Rem by nan fin busco 8
            Rem m$ = "Pedido Inexistente"
            Rem  G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Rem  Exit Sub
        
        End If
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    Call Conecta_Empresa
    
    If WIdioma = 1 Then
        Idioma.ListIndex = 1
    End If
    
    
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        ZReal = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
        ZRealAnt = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
        ZSumaReal = ZReal + ZRealAnt
        If ZSumaReal = 0 And Cliente.Text <> "" And rstHoja!Producto = Terminado.Text Then
            If Val(XEmpresa) = 1 Then
                rstHoja.Close
                m$ = "El lote informado no esta aprobado por laboratorio"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
               Exit Sub
            End If
        End If
        rstHoja.Close
    End If
    
    If Trim(Cliente.Text) <> "" Then
        If Left$(Terminado.Text, 2) <> "PT" Then
            m$ = "Solo se puede emitir etiqietas a productos PT"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Terminado.Text = "  -     -   "
            Lote.SetFocus
            Exit Sub
        End If
    End If
    
    
    
    Rem recalcular aca la marca
    Rem de vencida
    
    WMarcaVencida = ""
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
        rstTerminado.Close
    End If
    
    If Val(ZZMeses) <> 0 Then

        XEmpresa = Wempresa
        ZZFechaActual = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
        
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
        
        For CiclaEmpresa = 1 To 11

            Wempresa = Empe(CiclaEmpresa, 1)
            txtOdbc = Empe(CiclaEmpresa, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZZRenglon = 0
            ZZCantidadLote = 0
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZRenglon = ZZRenglon + 1
                            ZZCantidadLote = rstHoja!Canti1
                            ZZCantidad = rstHoja!Cantidad
                            ZZTipo = rstHoja!Tipo
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstHoja.Close
            End If
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Lote.Text + "'"
            ZSql = ZSql + " and Hoja.Producto = " + "'" + Terminado.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                ZZFecha = rstHoja!Fecha
                ZZArticulo = rstHoja!Articulo
                ZZLoteArti = rstHoja!lote1
                rstHoja.Close
                Exit For
            End If
            
        Next CiclaEmpresa
        
        Call Conecta_Empresa
        
        Rem VERIFICA EL 75%
        
        Rem dada
        Rem dada
        Rem dada
        Rem dada
        
        If Val(ZZRevalida) <> 0 Then
        
            WVida = Int(Val(ZZMesesRevalida) * 0.75)
            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
            WAno = Val(Right$(ZZFechaRevalida, 4))
            
                Else
                
            WVida = Int(Val(ZZMeses) * 0.75)
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
        ZZOrdVto = ZAno + ZMes + "01"
        
        If ZZOrdVto < ZZFechaActual Then
            WMarcaVencida = "S"
        End If
        
        Rem VERIFICA EL 100%
        
        If Val(ZZRevalida) <> 0 Then
        
            WVida = Int(Val(ZZMesesRevalida))
            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
            WAno = Val(Right$(ZZFechaRevalida, 4))
            
                Else
                
            WVida = Int(Val(ZZMeses))
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
        ZZOrdVto = ZAno + ZMes + "01"
        
        If ZZOrdVto < ZZFechaActual Then
            WMarcaVencida = "V"
        End If

    End If
    
    
    
    
    
    
    
    Rem WMarcaVencida = "S"
    Rem WEntra = "N"
    Rem
    Rem XParam = "'" + Lote.Text + "','" _
    rem         + Terminado.Text + "'"
    Rem spHoja = "ListaHojaProducto " + XParam
    Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstHoja.RecordCount > 0 Then
    Rem     WEntra = "S"
    Rem     WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
    Rem     rstHoja.Close
    Rem End If
    Rem
    Rem If WEntra = "N" Then
    Rem     XParam = "'" + Terminado.Text + "','" _
    rem             + Lote.Text + "'"
    Rem     spMovguia = "ListaMovguiaLote1 " + XParam
    Rem     Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If rstMovguia.RecordCount > 0 Then
    Rem         WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
    Rem         rstMovguia.Close
    Rem     End If
    Rem End If
    
    Rem If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
    Rem     If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
    Rem             Else
    Rem         m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
    rem              "Por favor comuniquese con el laboratorio para su revalida"
    Rem         G% = MsgBox(m$, 0, "Impresion de Etiquetas")
    Rem         Terminado.Text = "  -     -   "
    Rem         Lote.SetFocus
    Rem         Exit Sub
    Rem     End If
    Rem End If
    
    If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
        m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
             "Por favor comuniquese con el laboratorio para su revalida"
        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Terminado.Text = "  -     -   "
        Lote.SetFocus
        Exit Sub
    End If
    
    
    Rem Listado.DataFiles(0) = WEmpresa + "coti.mdb"
    Rem Listado.DataFiles(1) = ""
    Rem Listado.DataFiles(0) = WEmpresa + "VENT.mdb"
    Rem Listado.DataFiles(2) = WEmpresa + "admi.mdb"
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    
    Rem ****** by nan 19-9-2014 se agrega la linea 41000 es un producto peligroso *********
    
    If (XCodigo >= 40000 And XCodigo <= 49999) And (XCodigo <> 41000) Then
    Rem If (XCodigo >= 4400 And XCodigo <= 4499) Then
        Call Imprime_EtiquetaVerde
        Exit Sub
    End If
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
            
        XEmpresa = Wempresa
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
        
        TipoPro = "PT"
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If Left$(Terminado.Text, 2) <> "PT" Then
            Select Case Left$(Terminado.Text, 2)
                Case "DY", "DS"
                    TipoPro = "CO"
                Case "QC"
                    TipoPro = "FA"
                Case Else
                    TipoPro = "PT"
            End Select
                Else
            If XCodigo >= 0 And XCodigo <= 999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    TipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        TipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            TipoPro = "BI"
                                Else
                            TipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If
        
        If Left$(Terminado.Text, 2) = "YQ" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YH" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YP" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YF" Then
            TipoPro = "FA"
        End If
        
        
        
        
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                TipoPro = "FA"
            End If
        End If
            
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
        
        WVida = 0
        WLinea = 0
        
        DescripcionFarma.Text = ""
        
        DesProducto.Caption = ""
        ZZImpreVtoTermi = 0
                    
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
        
            WLinea = rstTerminado!Linea
            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
            
            If Idioma.ListIndex = 0 Then
            
                DesProducto.Caption = Trim(rstTerminado!Descripcion)
                Descripcion.Text = ""
                If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                    Descripcion.Text = Trim(rstTerminado!Descripcion)
                End If
                DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
                
                WConservacion = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
                WConservacion = RTrim(WConservacion)
                WConservacionII = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
                WConservacionII = RTrim(WConservacionII)
                
                    Else
                    
                DesProducto.Caption = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
                Descripcion.Text = ""
                If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                    Descripcion.Text = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
                End If
                DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiquetaIngles), "", rstTerminado!DescriEtiquetaIngles)
                
                WConservacion = IIf(IsNull(rstTerminado!ConservacionIngles), "", rstTerminado!ConservacionIngles)
                WConservacion = RTrim(WConservacion)
                WConservacionII = IIf(IsNull(rstTerminado!ConservacionIIIngles), "", rstTerminado!ConservacionIIIngles)
                WConservacionII = RTrim(WConservacionII)
                
                DesProducto.Caption = Trim(DesProducto.Caption)
                Descripcion.Text = Trim(Descripcion.Text)
                DescripcionFarma.Text = Trim(DescripcionFarma.Text)
                
            End If
            
            Rem WConservacion = "asdfsd fsd fsd fsd fsd"
            Rem WConservacionII = " sdfsdfsdf dsf sd fsdf sdfsd"
            
            WImpreadi = ""
            WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
            
            WClase = ""
            WRiesgo = ""
            WIntervencion = ""
            WNaciones = ""
            WEmbalaje = ""
            wdescriOnu = ""
            
            WClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
            WRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
            WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
            
            WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
            WObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
            WObservaciones = ""
            WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
            
            rstTerminado.Close
            
        End If
        
        If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
            WVida = 0
        End If
            
        spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            DescripcionFarma.Text = IIf(IsNull(rstPrecios!DescripcionFarma), "", rstPrecios!DescripcionFarma)
            Descripcion.Text = Trim(Left$(rstPrecios!Descripcion, 50))
            rstPrecios.Close
        End If
                    
        spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            DesCliente.Caption = rstClientes!Razon
            rstClientes.Close
        End If
                    
        Call Conecta_Empresa
        
        Wvencimiento = ""
        XEmpresa = Wempresa
        
        Select Case Val(Wempresa)
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
        
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
            
                WMes = Val(Mid$(rstHoja!Fecha, 4, 2))
                WAno = Val(Right$(rstHoja!Fecha, 4))
            
                ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                
                If Val(ZZRevalida) <> 0 Then
                    WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                    WAno = Val(Right$(ZZFechaRevalida, 4))
                    WVida = Val(ZZMesesRevalida)
                End If
            
                For Ciclo = 1 To WVida
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                Next Ciclo
                WElaboracion = rstHoja!Fecha
                Rem XFec1 = WElaboracion
                Rem SumaDia = WVida + 1
                Rem Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                If WVida <> 0 Then
                    XMes = Str$(WMes)
                    XAno = Str$(WAno)
                    Call Ceros(XMes, 2)
                    Call Ceros(XAno, 4)
                    Wvencimiento = "01/" + XMes + "/" + XAno
                End If
                rstHoja.Close
                
                ZZRenglon = 0
                ZZTipo = ""
                ZZTerminado = ""
                ZZArticulo = ""
                ZZCantidad = 0
                ZZCantidadLote = 0
                ZZLote = ""
                spHoja = "ListaHoja " + "'" + Lote.Text + "'"
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    With rstHoja
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                ZZRenglon = ZZRenglon + 1
                                ZZTipo = rstHoja!Tipo
                                ZZTerminado = rstHoja!Terminado
                                ZZArticulo = rstHoja!Articulo
                                ZZCantidad = rstHoja!Cantidad
                                ZZCantidadLote = rstHoja!Canti1
                                ZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstHoja.Close
                End If
                
            End If
            
        Next ZCiclo
        
        Call Conecta_Empresa
        
        
        ZZZZLoteOriginal = ""
        If TipoPro = "FA" Then
            If Terminado.Text = "PT-25062-777" Or Terminado.Text = "PT-25062-778" Or Terminado.Text = "PT-25106-777" Or Terminado.Text = "PT-25106-778" Or Terminado.Text = "PT-25046-777" Or Terminado.Text = "PT-25049-777" Or Terminado.Text = "PT-25136-777" Or Terminado.Text = "PT-25135-777" Then
                Call Calcula_Mono_Otro
                If ZZZZElaboracion <> "" Then
                    WElaboracion = ZZZZElaboracion
                End If
                If ZZZZVencimiento <> "" Then
                    Wvencimiento = ZZZZVencimiento
                End If
            End If
        End If
        Rem DEJE ACA
        Rem DEJE ACA
        Rem DEJE ACA
        Rem DEJE ACA
        
        Rem If Tipopro <> "FA" Then
        Rem
        Rem     Rem veo si es mono
        Rem     If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
        Rem
        Rem         ZVto = ""
        Rem         ZLaudo = ZZLote
        Rem         ZArticulo = ZZArticulo
        Rem         ZFecha = ""
        Rem         ZFechaVto = ""
        Rem
        Rem         For ZCiclo = 1 To ZHasta1
        Rem
        Rem             WEmpresa = CargaEmpresa(ZCiclo, 1)
        Rem             txtOdbc = CargaEmpresa(ZCiclo, 2)
        Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem
        Rem             ZSql = ""
        Rem             ZSql = ZSql + "Select *"
        Rem             ZSql = ZSql + " FROM Laudo"
        Rem             ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
        Rem             ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
        Rem             spLaudo = ZSql
        Rem             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstLaudo.RecordCount > 0 Then
        Rem                 ZFecha = rstLaudo!Fecha
        Rem                 ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
        Rem                 rstLaudo.Close
        Rem                 Exit For
        Rem             End If
        Rem
        Rem         Next ZCiclo
        Rem
        Rem         Call Conecta_Empresa
        Rem
        Rem         ZVto = ""
        Rem         ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        Rem         If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
        Rem             Call Valida_fecha(ZFechaVto, Auxi)
        Rem             If Auxi = "S" Then
        Rem                 ZVto = ZFechaVto
        Rem             End If
        Rem         End If
        Rem
        Rem         If ZVto = "" Then
        Rem
        Rem             ZMeses = 0
        Rem             ZSql = ""
        Rem             ZSql = ZSql + "Select *"
        Rem             ZSql = ZSql + " FROM Articulo"
        Rem             ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        Rem             spArticulo = ZSql
        Rem             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstArticulo.RecordCount > 0 Then
        Rem                 ZMeses = rstArticulo!Meses
        Rem                 rstArticulo.Close
        Rem             End If
        Rem
        Rem             WMes = Val(Mid$(ZFecha, 4, 2))
        Rem             WAno = Val(Right$(ZFecha, 4))
        Rem             For ZCiclo = 1 To ZMeses
        Rem                 WMes = WMes + 1
        Rem                 If WMes > 12 Then
        Rem                     WAno = WAno + 1
        Rem                     WMes = 1
        Rem                 End If
        Rem             Next ZCiclo
        Rem
        Rem             XMes = Str$(WMes)
        Rem             XAno = Str$(WAno)
        Rem             Call Ceros(XMes, 2)
        Rem             Call Ceros(XAno, 4)
        Rem             If Val(Left$(ZFecha, 2)) <= 30 Then
        Rem                 If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
        Rem                     ZVto = "28/" + XMes + "/" + XAno
        Rem                         Else
        Rem                     ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
        Rem                 End If
        Rem                     Else
        Rem                 If Val(XMes) = 2 Then
        Rem                     ZVto = "28/" + XMes + "/" + XAno
        Rem                         Else
        Rem                     ZVto = "30/" + XMes + "/" + XAno
        Rem                 End If
        Rem             End If
        Rem
        Rem         End If
        Rem
        Rem         Rem
        Rem         Rem
        Rem         Rem verifica venciminiento
        Rem         Rem
        Rem         Rem
        Rem         Rem
        Rem
        Rem         ZZVidaUtil = 0
        Rem         spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Rem         Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstTerminado.RecordCount > 0 Then
        Rem             ZZVidaUtil = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
        Rem             ZZVidaUtil = Int(ZZVidaUtil * 0.25)
        Rem             rstTerminado.Close
        Rem         End If
        Rem
        Rem         WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Rem         WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
        Rem
        Rem         WFechaVencimiento = "01" + Mid$(ZVto, 3, 10)
        Rem         WFechaVencimientoOrd = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        Rem
        Rem         Pasa = "S"
        Rem         If Left$(WFechaActualOrd, 6) >= Left$(WFechaVencimientoOrd, 6) Then
        Rem
        Rem             Pasa = "N"
        Rem
        Rem                 Else
        Rem
        Rem             Meses = 0
        Rem             WMes = Val(Mid$(WFechaActual, 4, 2))
        Rem             WAno = Val(Right$(WFechaActual, 4))
        Rem             Do
        Rem                 Meses = Meses + 1
        Rem                 WMes = WMes + 1
        Rem                 If WMes > 12 Then
        Rem                     WAno = WAno + 1
        Rem                     WMes = 1
        Rem                 End If
        Rem                 XMes = Str$(WMes)
        Rem                 XAno = Str$(WAno)
        Rem                 Call Ceros(XMes, 2)
        Rem                 Call Ceros(XAno, 4)
        Rem                 WCompara = "01/" + XMes + "/" + XAno
        Rem                 If WCompara = WFechaVencimiento Then
        Rem                     Exit Do
        Rem                 End If
        Rem             Loop
        Rem
        Rem             If ZZVidaUtil >= Meses Then
        Rem                 Pasa = "N"
        Rem             End If
        Rem
        Rem         End If
        Rem
        Rem         If Pasa = "N" Then
        Rem            m$ = "EL Producto tiene menos de 25% de la vida util del PT"
        Rem           G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Rem           Exit Sub
        Rem         End If
        Rem
        Rem         Wvencimiento = ZVto
        Rem
        Rem     End If
        Rem
        Rem End If
        
        
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                         Exit Do
                    End If
                Loop
            End If
        End With
        
        WTara = Val(Tara.Text)
        WNeto = Val(Cantidad.Text)
        
        If WTara = 0 Then
            WBruto = 0
                Else
            WBruto = WTara + WNeto
        End If
        
        WRazon = ""
        Rem WDirEntrega = ""
                
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
                
        WImpreVto = 0
        spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            WRazon = rstClientes!Razon
            WImpreVto = IIf(IsNull(rstClientes!ImpreVto), "0", rstClientes!ImpreVto)
            WProv = rstClientes!Provincia
            Rem WDirEntrega = rstClientes!DirEntrega
            rstClientes.Close
        End If
        

        
        
        
        
        ZVencimiento = Wvencimiento
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            Rem no hago nada
                Else
            Rem vencimiento
            If ZZImpreVtoTermi = 0 Then
                If WImpreVto = 0 Then
                    Rem ZVencimiento = ""
                End If
            End If
        End If
                
        Call Conecta_Empresa
        
        If TipoPro <> "FA" Then
            Da = 0
            If Len(Descripcion.Text) > 16 Then
                For Da = 17 To 1 Step -1
                    If Mid$(Descripcion.Text, Da, 1) = Space$(1) Then
                        ZZNombre = Mid$(Descripcion.Text, 1, Da)
                        ZZNombreII = Mid$(Descripcion.Text, Da + 1, 100)
                        Exit For
                    End If
                Next Da
                    Else
                ZZNombre = Descripcion.Text
                ZZNombreII = ""
            End If
            If TipoProceso.ListIndex > 0 Then
                ZZNombre = Descripcion.Text
                ZZNombreII = ""
            End If
                Else
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
        If Tipo.ListIndex <> 0 Then
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
        If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
        End If
        
        
        If Tipo.ListIndex = 7 Or (Tipo.ListIndex = 6 And TipoPro = "FA") Then
            If Len(Descripcion.Text) > 25 Then
                For Da = 26 To 1 Step -1
                    If Mid$(Descripcion.Text, Da, 1) = Space$(1) Then
                        ZZNombre = Mid$(Descripcion.Text, 1, Da)
                        ZZNombreII = Mid$(Descripcion.Text, Da + 1, 100)
                        Exit For
                    End If
                Next Da
                    Else
                ZZNombre = Descripcion.Text
                ZZNombreII = ""
            End If
        End If
        
        
        
        Da = 0
        ZRazon = ""
        ZRazonII = ""
        If Len(WRazon) > 27 Then
            For Da = 27 To 1 Step -1
                If Mid$(WRazon, Da, 1) = Space$(1) Then
                    ZRazon = Mid$(WRazon, 1, Da)
                    ZRazonII = Mid$(WRazon, Da + 1, 100)
                    Exit For
                End If
            Next Da
                Else
            ZRazon = WRazon
            ZRazonII = ""
        End If
        
        If Val(WProv) = 24 Then
            If Idioma.ListIndex = 0 Then
                WObservaciones = "Hecho en Argentina"
                    Else
                WObservaciones = "Made in Argentina"
            End If
        End If
            
        Rem If Val(WProv) = 1 Then
        Rem     WObservaciones = "Industria Argentina"
        Rem End If
            
            
        ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
            
            
        With rstEtiqueta
            For Da = 1 To Val(Etiquetas)
                .Index = "Codigo"
                .AddNew
                !Codigo = Da
                WLote = Lote.Text
                Call Ceros(WLote, 6)
                WCantidad = Cantidad.Text
                Call Ceros(WCantidad, 4)
                !Terminado = Terminado.Text
                !Lote = Val(Lote.Text)
                !Cliente = Cliente.Text
                !Cantidad = Val(Cantidad.Text)
                !Nombre = Left$(ZZNombre, 30)
                !NombreII = Left$(ZZNombreII, 30)
                Rem by nan
                Rem  If Trim(Cliente.Text) <> "" Then
                !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + " " + WLote
                !Impre2 = ""
                Rem          Else
                Rem      !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3)
                Rem      !Impre2 = WLote
                Rem   End If
                
                If Trim(ZZZZLoteOriginal) <> "" Then
                    !Impre2 = Left$(ZZZZLoteOriginal, 30)
                End If
             
             
             
                If Tipo.ListIndex = 6 Or Tipo.ListIndex = 7 Then
                    !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3)
                    !Impre2 = WLote
                End If
             
             
                !Razon = ZRazon
                !DirEntrega = ZRazonII
                !Clase = WClase
                !Intervencion = WIntervencion
                If TipoProceso.ListIndex = 0 Then
                    Rem by nan
                    If wdescriOnu = "" Then
                        wdescriOnu = "0"
                    End If
                    !Descrionu = wdescriOnu
                End If
                
                !Naciones = WNaciones
                !Embalaje = WEmbalaje
                !Bruto = WBruto
                !Tara = WTara
                !Neto = WNeto
                !Observaciones = Left$(WObservaciones, 20)
                !Elaboracion = Right$(WElaboracion, 7)
                !Vencimiento = Right$(ZVencimiento, 7)
                !Conservacion = Trim(WConservacion)
                !ConservacionII = Trim(WConservacionII)
                If TipoPro = "FA" Then
                    !NombreFarmaI = DesProducto.Caption
                    !NombreFarmaII = DescripcionFarma.Text
                    Rem If Tipo.ListIndex = 6 Then
                    Rem     !NombreFarmaI = ZZNombre
                    Rem     !NombreFarmaII = ZZNombreII
                    Rem End If
                        Else
                    Rem by nan 23-3-2011
                    conserva = !ConservacionII
                    !NombreFarmaI = "MANTENER EN ENVASE ORIGINAL CERRADO"
                    !NombreFarmaI = conserva
                    !NombreFarmaII = ""
                    If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
                        If Cliente.Text = "Z00007" Then
                            !NombreFarmaII = "PRODUCTO PARA LA PRODUCCION DE CUERO"
                        End If
                    End If
                End If
                !TipoPro = ""
                If Trim(Cliente.Text) = "" Then
                    !TipoPro = Left$(Terminado.Text, 2)
                End If
                If Tipo.ListIndex = 7 Then
                    !TipoPro = Left$(Terminado.Text, 2)
                End If
                
                ZFazon = "N"
                Select Case Val(WLinea)
                    Case 3, 4, 5, 7, 8, 9, 11, 12, 14, 19, 22
                        ZFazon = "N"
                    Case 6, 16, 17
                        ZFazon = "N"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        ZFazon = "N"
                    Case Else
                        ZFazon = "S"
                End Select
                If TipoPro = "CO" Then
                    !NombreFarmaI = ""
                    !NombreFarmaII = ""
                End If
                If ZFazon = "S" Then
                    !NombreFarmaI = ""
                    !NombreFarmaII = ""
                End If
                
                !ImpreOc = ""
                !ImpreDirEntrega = ""
                
                If ZEtiI = 1 Then
                    If Trim(ZZOrdenCpa) <> "" Then
                        !ImpreOc = "Orden Cpa.:" + ZZOrdenCpa
                    End If
                End If
                If ZEtiII = 1 Then
                    !ImpreDirEntrega = ZDescriDirEntrega
                End If
                
                Rem by nan 3-2-2016 agregar rutina para numeracion ***************
                If Wempresa = "0005" Then
                    If Tipo.ListIndex = 6 Then
                        If Check1.Value = 0 Then
                            !DirEntrega = ""
                                Else
                            Rem !DirEntrega = Str$(ZZDesdeNumero) + "/" + Str$(ZZHastaNumero)
                            !DirEntrega = Str$(ZZDesdeNumero)
                            ZZDesdeNumero = ZZDesdeNumero + 1
                        End If
                    End If
                End If
                Rem fin by nan
                
                .Update
            Next Da
        End With
    
        Listado.WindowTitle = "Emision de Etiquetas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Da = 0
        If TipoPro = "FA" Then
            Da = Len(Trim(DesProducto.Caption))
                Else
            Da = Len(Trim(Descripcion.Text))
        End If
        
        Select Case Tipo.ListIndex
            Case 0
                Select Case Val(XEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        If TipoProceso.ListIndex = 0 Then
                        
                            If WImpreadi = "S" Then
                                m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                            End If
                            If Idioma.ListIndex = 0 Then
                                Listado.ReportFileName = "WEti1Nuevo.RPT"
                                    Else
                                Listado.ReportFileName = "WEti1NuevoIngles.rpt"
                            End If
                            
                                Else
                                
                            If WImpreadi <> "S" Then
                                If Da > 20 Then
                                    Listado.ReportFileName = "eti10.rpt"
                                        Else
                                    Listado.ReportFileName = "eti1.rpt"
                                End If
                                    Else
                                m$ = " Coloque la etiqueta que en su margen tengo el Numero " + WTipoeti
                                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                                If Da > 20 Then
                                    Listado.ReportFileName = "eti110.rpt"
                                        Else
                                    Listado.ReportFileName = "eti101.rpt"
                                End If
                            End If
                            
                        End If
                                
                    Case Else
                        If WImpreadi <> "S" Then
                            If Da > 20 Then
                                Listado.ReportFileName = "eti10Pellital.rpt"
                                    Else
                                Listado.ReportFileName = "eti1Pellital.rpt"
                            End If
                                Else
                            m$ = " Coloque la etiqueta que en su margen tengo el Numero " + WTipoeti
                            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                            If Da > 20 Then
                                Listado.ReportFileName = "eti110Pellital.rpt"
                                    Else
                                Listado.ReportFileName = "eti101Pellital.rpt"
                            End If
                        End If
                        
                End Select
            
            Case 1
                If TipoProceso.ListIndex = 0 Then
                
                    If WImpreadi = "S" Then
                        m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    End If
                    If Da > 23 Then
                        Listado.ReportFileName = "weti20Nuevo.rpt"
                            Else
                        If Da > 16 Then
                            Listado.ReportFileName = "weti30Nuevo.rpt"
                                Else
                            Listado.ReportFileName = "weti2Nuevo.rpt"
                        End If
                    End If
                    
                        Else
                        
                    If WImpreadi <> "S" Then
                        If Da > 20 Then
                            Listado.ReportFileName = "weti20.rpt"
                                Else
                            Listado.ReportFileName = "weti2.rpt"
                        End If
                            Else
                        Rem m$ = "Producto Peligrosos no se pueden imprimir en etiquetas Chicas"
                        Rem G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        If Da > 20 Then
                            Listado.ReportFileName = "weti20.rpt"
                                Else
                            Listado.ReportFileName = "weti2.rpt"
                        End If
                    End If
                    
                End If
                
            Case Else
                Listado.ReportFileName = "WEtiBlanco.rpt"
            
        End Select
            
        Rem If WVida <> 0 Then
        Rem
        Rem     WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Rem     WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
        Rem
        Rem     WFechaVencimiento = "01" + Mid$(WVencimiento, 3, 10)
        Rem     WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
        Rem
        Rem     Pasa = "S"
        Rem     If WFechaActualOrd >= WFechaVencimientoOrd Then
        Rem         Pasa = "N"
        Rem             Else
        Rem         Meses = 0
        Rem         WMes = Val(Mid$(WFechaActual, 4, 2))
        Rem         WAno = Val(Right$(WFechaActual, 4))
        Rem         Do
        Rem             Meses = Meses + 1
        Rem             WMes = WMes + 1
        Rem             If WMes > 12 Then
        Rem                 WAno = WAno + 1
        Rem                 WMes = 1
        Rem             End If
        Rem             XMes = Str$(WMes)
        Rem             XAno = Str$(WAno)
        Rem             Call Ceros(XMes, 2)
        Rem             Call Ceros(XAno, 4)
        Rem             WCompara = "01/" + XMes + "/" + XAno
        Rem             If WCompara = WFechaVencimiento Then
        Rem                 Exit Do
        Rem             End If
        Rem         Loop
        Rem
        Rem        ZMeses = Int(WVida / 2)
        Rem        If ZMeses > 12 Then
        Rem            ZMeses = 12
        Rem         End If
        Rem
        Rem         If Meses <= ZMeses Then
        Rem             Pasa = "N"
        Rem         End If
        Rem
        Rem     End If
        Rem
        Rem     If Pasa = "N" Then
        Rem         m$ = "EL Producto tiene menos de un año o ya paso mas del 50% de su de vida util"
        Rem         G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Rem         Da = 0
        Rem         With rstEtiqueta
        Rem             .Index = "Codigo"
        Rem             .Seek ">=", Da
        Rem             If .NoMatch = False Then
        Rem                 Do
        Rem                     .Delete
        Rem                     .MoveNext
        Rem                     If .EOF = True Then
        Rem                         Exit Do
        Rem                     End If
        Rem                 Loop
        Rem             End If
        Rem         End With
        Rem         Exit Sub
        Rem     End If
        Rem
        Rem End If
        
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
            
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                If Trim(WClase) = "" Then
                
                    m$ = "Coloque la etiqueta correspondirentes a los productos de Farma"
                    G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    
                    If Da > 30 Then
                        Listado.ReportFileName = "EtiFarmaIII.rpt"
                            Else
                        Listado.ReportFileName = "EtiFarmaIV.rpt"
                    End If
                    
                        Else
                
                    m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                    G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                
                    Listado.ReportFileName = "EtiFarmaPeligroso.rpt"
                    
                    Rem If Left$(WClase, 1) = "8" Then
                    Rem     Listado.ReportFileName = "EtiFarmaIIIPrueba8.rpt"
                    Rem         Else
                    Rem     If Left$(WClase, 1) = "6" Then
                    Rem         Listado.ReportFileName = "EtiFarmaIIIPrueba6.rpt"
                    Rem             Else
                    Rem         Listado.ReportFileName = "EtiFarmaIIIPrueba.rpt"
                    Rem     End If
                    Rem End If
                    
                End If
                
            End If
        End If
        
        If Tipo.ListIndex = 6 Then
        
            XCodigo = Val(Mid$(Terminado.Text, 4, 5))
            If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            
                If Idioma.ListIndex = 0 Then
                
                    Listado.ReportFileName = "EtiquetaInteriorFarmaNuevo.rpt"
                    
                    If Trim(WClase) <> "" Then
                        m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        Listado.ReportFileName = "EtiquetaInteriorFarmaNuevoPeligroso.rpt"
                        Rem If Left$(WClase, 1) = "8" Then
                        Rem     Listado.ReportFileName = "EtiquetaInterior8Farma.rpt"
                        Rem         Else
                        Rem     If Left$(WClase, 1) = "6" Then
                        Rem         Listado.ReportFileName = "EtiquetaInterior6Farma.rpt"
                        Rem             Else
                        Rem         Listado.ReportFileName = "EtiquetaInteriorFarma.rpt"
                        Rem     End If
                        Rem End If
                    End If
                    
                        Else
                        
                    Listado.ReportFileName = "EtiquetaInteriorFarmaInglesNuevo.rpt"
                    If Trim(WClase) <> "" Then
                        m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        Rem If Left$(WClase, 1) = "8" Then
                        Rem     Listado.ReportFileName = "EtiquetaInterior8FarmaIngles.rpt"
                        Rem         Else
                        Rem     If Left$(WClase, 1) = "6" Then
                        Rem         Listado.ReportFileName = "EtiquetaInterior6FarmaIngles.rpt"
                        Rem             Else
                        Rem         Listado.ReportFileName = "EtiquetaInteriorFarmaIngles.rpt"
                        Rem     End If
                        Rem End If
                    End If
                    
                End If
            
                    Else
                    
                If Idioma.ListIndex = 0 Then
                
                    Listado.ReportFileName = "EtiquetaInterior.rpt"
                    If Trim(WClase) <> "" Then
                        If Left$(WClase, 1) = "8" Then
                            Listado.ReportFileName = "EtiquetaInterior8.rpt"
                                Else
                            If Left$(WClase, 1) = "6" Then
                                Listado.ReportFileName = "EtiquetaInterior6.rpt"
                                    Else
                                Listado.ReportFileName = "EtiquetaInterior.rpt"
                            End If
                        End If
                    End If
                    
                        Else
                        
                    Listado.ReportFileName = "EtiquetaInteriorIngles.rpt"
                    If Trim(WClase) <> "" Then
                        If Left$(WClase, 1) = "8" Then
                            Listado.ReportFileName = "EtiquetaInterior8Ingles.rpt"
                                Else
                            If Left$(WClase, 1) = "6" Then
                                Listado.ReportFileName = "EtiquetaInterior6Ingles.rpt"
                                    Else
                                Listado.ReportFileName = "EtiquetaInteriorIngles.rpt"
                            End If
                        End If
                    End If
                    
                End If
                
            End If
            
        End If
        
        If Tipo.ListIndex = 7 Then
        
            If Trim(WClase) <> "" Then
                m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            End If
            Rem BY NAN
            Listado.ReportFileName = "EtiquetaInteriorQuimicos.rpt"
            Rem by nan
         
        End If
        
        If Tipo.ListIndex = 8 Then
            If WImpreadi = "S" Then
                m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            End If
            Listado.ReportFileName = "WEtiPigmento.rpt"
        End If
        
        Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.Connect = Connect()
        
        Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
        Listado.DataFiles(1) = ""
        
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.PrinterCopies = 1
        Listado.Action = 1
        
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Imprime_EnProceso()

    Salida = "N"
    Da = 0
    
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
            
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
    
        WVida = 0
        WLinea = 0
        DescripcionFarma.Text = ""
                
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesProducto.Caption = Trim(rstTerminado!Descripcion)
            WLinea = rstTerminado!Linea
            Descripcion.Text = Trim(rstTerminado!Descripcion)
            DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
            
            WImpreadi = ""
            WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
            
            WClase = ""
            WIntervencion = ""
            WNaciones = ""
            WEmbalaje = ""
            wdescriOnu = ""
            
            WClase = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
            WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
            
            WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
            WObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
            WObservaciones = ""
            WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
            WConservacion = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
            WConservacion = RTrim(WConservacion)
            WConservacionII = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
            WConservacionII = RTrim(WConservacionII)
            rstTerminado.Close
        End If
    
        Call Conecta_Empresa
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
        WTara = Val(Tara.Text)
        WNeto = Val(Cantidad.Text)
    
        If WTara = 0 Then
            WBruto = 0
                Else
            WBruto = WTara + WNeto
        End If
    
        WRazon = ""
        Rem WDirEntrega = ""
            
        spHoja = "ListaHoja " + "'" + Lote.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            WMes = Val(Mid$(rstHoja!Fecha, 4, 2))
            WAno = Val(Right$(rstHoja!Fecha, 4))
        
            ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
            ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
            ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
            
            If Val(ZZRevalida) <> 0 Then
                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                WAno = Val(Right$(ZZFechaRevalida, 4))
                WVida = Val(ZZMesesRevalida)
            End If
        
            For Ciclo = 1 To WVida
                WMes = WMes + 1
                If WMes > 12 Then
                    WAno = WAno + 1
                    WMes = 1
                End If
            Next Ciclo
            WElaboracion = rstHoja!Fecha
            Rem XFec1 = WElaboracion
            Rem SumaDia = WVida + 1
            Rem Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            If WVida <> 0 Then
                XMes = Str$(WMes)
                XAno = Str$(WAno)
                Call Ceros(XMes, 2)
                Call Ceros(XAno, 4)
                Wvencimiento = "01/" + XMes + "/" + XAno
            End If
            rstHoja.Close
            
        End If
            
            
        With rstEtiqueta
            For Da = 1 To Val(Etiquetas)
                .Index = "Codigo"
                .AddNew
                !Codigo = Da
                WLote = Lote.Text
               Rem by nan
                Call Ceros(WLote, 6)
                WCantidad = Cantidad.Text
                Call Ceros(WCantidad, 4)
                !Terminado = Terminado.Text
                !Lote = WLote
                !Cliente = Cliente.Text
                !Cantidad = Val(Cantidad.Text)
                !Nombre = Left$(Descripcion.Text, 30)
                !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + Space$(1) + WLote + Space$(1) + WCantidad
                !Razon = WRazon
                !DirEntrega = WDirentrega
                !Clase = WRiesgo
                !Intervencion = WIntervencion
                If TipoProceso.ListIndex = 0 Then
                Rem by nan
                   !Descrionu = wdescriOnu
                End If
                !Naciones = WNaciones
                !Embalaje = WEmbalaje
                !Bruto = WBruto
                !Tara = WTara
                !Neto = WNeto
                !Observaciones = Left$(WObservaciones, 20)
                !Elaboracion = Right$(WElaboracion, 7)
                !Vencimiento = Right$(Wvencimiento, 7)
                !Conservacion = Trim(WConservacion)
                !ConservacionII = Trim(WConservacionII)
                !NombreFarmaI = DesProducto.Caption
                !NombreFarmaII = DescripcionFarma.Text
                !TipoPro = ""
                If Trim(Cliente.Text) = "" Then
                    !TipoPro = Left$(Terminado.Text, 2)
                End If
                .Update
            Next Da
        End With

        Listado.WindowTitle = "Emision de Etiquetas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height

        Rem Listado.ReportFileName = "EtiFarmaV.rpt"
        Listado.ReportFileName = "Eti211201.rpt"
    
            Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
        Listado.DataFiles(1) = ""
    
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.PrinterCopies = 1
        Listado.Action = 1
        
        FSFS = Listado.ReportFileName
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    
    
     With rstEtiqueta
        .Close
    End With
     With rstEtiquetaII
        .Close
    End With
     With rstEtiquetaIII
        .Close
    End With
     With rstEtiquetaIV
        .Close
    End With
    
    PrgEti3.Hide
    
    
    Unload Me
    Menu.Show
    
    
End Sub

Private Sub Baja_Click()
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With


    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With



    With rstEtiquetaIII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With


    With rstEtiquetaIV
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With




End Sub


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check2.Visible = True
            Else
        Check2.Visible = False
        Check2.Value = False
        Label10.Visible = False
        Label11.Visible = False
        DesdeNumero.Visible = False
        HastaNumero.Visible = False
   End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Label10.Visible = True
        Label11.Visible = True
        DesdeNumero.Visible = True
        HastaNumero.Visible = True
            Else
        Label10.Visible = False
        Label11.Visible = False
        DesdeNumero.Visible = False
        HastaNumero.Visible = False
    End If
End Sub

Private Sub Command1_Click()
    
    Erase ZImpre
    Erase ZImpreI
    Erase ZImpreII
    Erase ZImpreIII
    
    ZLugarImpre = 0
    ZLugarImpreI = 0
    ZLugarImpreII = 0
    ZLugarImpreIII = 0
    
    Rem para imprimir nº de etiqueta
    If Check1.Value = 1 Then
        If Check2.Value = 1 Then
            ZZDesdeNumero = Val(DesdeNumero.Text)
            ZZHastaNumero = Val(HastaNumero.Text)
                Else
            ZZDesdeNumero = 1
            ZZHastaNumero = Val(Etiquetas.Text)
        End If
    End If
    
    
    
    Rem end by nan
    
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
            
            
            
            
            
    ZZClave = "PT" + Mid$(Terminado.Text, 3, 10) + "001"

    Sql1 = "Select *"
    Sql2 = " FROM DatosEtiqueta"
    Sql3 = " Where DatosEtiqueta.Clave = " + "'" + ZZClave + "'"
    spDatosEtiqueta = Sql1 + Sql2 + Sql3
    Set rstDatosEtiqueta = db.OpenRecordset(spDatosEtiqueta, dbOpenSnapshot, dbSQLPassThrough)
    If rstDatosEtiqueta.RecordCount > 0 Then
        Rem estan cargados
        rstDatosEtiqueta.Close
            Else
        Call Conecta_Empresa
        m$ = "No estan cargados los datos adicionales de peligrosidad"
        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Exit Sub
    End If
    
    
    
    Call Conecta_Empresa
    
    
    
    
    
    Salida = "N"
    Da = 0
    
    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    With rstEtiquetaIII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    With rstEtiquetaIV
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    
    
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
        ZEtiI = Trim(IIf(IsNull(rstClientes!EtiI), "0", rstClientes!EtiI))
        ZEtiII = Trim(IIf(IsNull(rstClientes!EtiII), "0", rstClientes!EtiII))
        rstClientes.Close
    End If
    
    
    TipoPro = "PT"
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If Left$(Terminado.Text, 2) <> "PT" Then
        Select Case Left$(Terminado.Text, 2)
            Case "DY", "DS"
                TipoPro = "CO"
            Case "QC"
                TipoPro = "FA"
            Case Else
                TipoPro = "PT"
        End Select
            Else
        If XCodigo >= 0 And XCodigo <= 999 Then
            TipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    TipoPro = "FA"
                        Else
                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                        TipoPro = "BI"
                            Else
                        TipoPro = "PT"
                    End If
                End If
            End If
        End If
    End If
    
    If Left$(Terminado.Text, 2) = "YQ" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YH" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YP" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YF" Then
        TipoPro = "FA"
    End If
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
            TipoPro = "FA"
        End If
    End If
    
    
    
    If Trim(Cliente.Text) <> "" And Val(Wempresa) = 1 And TipoPro <> "CO" Then
    
        Pasa = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " and Pedido.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and Pedido.Terminado = " + "'" + Terminado.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            rstPedido.Close
            Pasa = 1
                
                Else
                                        
            Rem by nan para 8
            Rem   If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            Rem BY NAN BUSCO EN 8
            OPEN_FILE_Pedido2
            With rstPedido2
               Auxi = Pedido.Text
               Call Ceros(Auxi, 6)
               .Index = "Clave"
               .Seek ">=", Auxi
                If .NoMatch = False Then
                    ZZCliente = rstPedido2!Cliente
                    Pasa = 1
                        Else
                    Pasa = 0
                End If
            End With
            rstPedido2.Close
            Rem BY NAN FIN
            
        End If
            
        If Pasa = 0 Then
            Rem by nan
            Call Conecta_Empresa
            m$ = "Pedido incorrecto"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Exit Sub
        End If
        
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZZCliente = rstPedido!Cliente
            ZZOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
            ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            ZDescriDirEntrega = ""
            
            rstPedido.Close
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    Rem ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    Rem ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    Rem ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    Rem ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDirEntrega(2) = ""
                    ZDirEntrega(3) = ""
                    ZDirEntrega(4) = ""
                    ZDirEntrega(5) = ""
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                
                    Else
                    
                Call Conecta_Empresa
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
                
            End If
            
                Else
                
            Rem by nan busco datos de entrega para 8
            OPEN_FILE_Pedido2
            With rstPedido2
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                .Index = "Clave"
                .Seek ">=", Auxi
                If .NoMatch = False Then
                    ZZCliente = rstPedido2!Cliente
                End If
            End With
            rstPedido2.Close
            Rem ***************
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
     
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    Rem ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    Rem ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    Rem ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    Rem ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDirEntrega(2) = ""
                    ZDirEntrega(3) = ""
                    ZDirEntrega(4) = ""
                    ZDirEntrega(5) = ""
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                 
                        Else
            
                Call Conecta_Empresa
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
            
            End If
             
        End If
        
    End If
    
    Call Conecta_Empresa
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    spHoja = "ListaHoja " + "'" + Lote.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        ZReal = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
        ZRealAnt = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
        ZSumaReal = ZReal + ZRealAnt
        If ZSumaReal = 0 And Cliente.Text <> "" And rstHoja!Producto = Terminado.Text Then
            If Val(XEmpresa) = 1 Then
                rstHoja.Close
                m$ = "El lote informado no esta aprobado por laboratorio"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
               Exit Sub
            End If
        End If
        rstHoja.Close
    End If
    
    If Trim(Cliente.Text) <> "" Then
        If Left$(Terminado.Text, 2) <> "PT" Then
            m$ = "Solo se puede emitir etiqietas a productos PT"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Terminado.Text = "  -     -   "
            Lote.SetFocus
            Exit Sub
        End If
    End If
    
    
    
    Rem recalcular aca la marca
    Rem de vencida
    
    WMarcaVencida = ""
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
        rstTerminado.Close
    End If
    
    If Val(ZZMeses) <> 0 Then

        XEmpresa = Wempresa
        ZZFechaActual = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
        
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
        
        For CiclaEmpresa = 1 To 11

            Wempresa = Empe(CiclaEmpresa, 1)
            txtOdbc = Empe(CiclaEmpresa, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZZRenglon = 0
            ZZCantidadLote = 0
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZRenglon = ZZRenglon + 1
                            ZZCantidadLote = rstHoja!Canti1
                            ZZCantidad = rstHoja!Cantidad
                            ZZTipo = rstHoja!Tipo
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstHoja.Close
            End If
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Lote.Text + "'"
            ZSql = ZSql + " and Hoja.Producto = " + "'" + Terminado.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                ZZFecha = rstHoja!Fecha
                ZZArticulo = rstHoja!Articulo
                ZZLoteArti = rstHoja!lote1
                rstHoja.Close
                Exit For
            End If
            
        Next CiclaEmpresa
        
        Call Conecta_Empresa
        
        Rem VERIFICA EL 75%
        
        If Val(ZZRevalida) <> 0 Then
        
            WVida = Int(Val(ZZMesesRevalida) * 0.75)
            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
            WAno = Val(Right$(ZZFechaRevalida, 4))
            
                Else
                
            WVida = Int(Val(ZZMeses) * 0.75)
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
        ZZOrdVto = ZAno + ZMes + "01"
        
        If ZZOrdVto < ZZFechaActual Then
            WMarcaVencida = "S"
        End If
        
        Rem VERIFICA EL 100%
        
        If Val(ZZRevalida) <> 0 Then
        
            WVida = Int(Val(ZZMesesRevalida))
            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
            WAno = Val(Right$(ZZFechaRevalida, 4))
            
                Else
                
            WVida = Int(Val(ZZMeses))
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
        ZZOrdVto = ZAno + ZMes + "01"
        
        If ZZOrdVto < ZZFechaActual Then
            WMarcaVencida = "V"
        End If

    End If
    
    If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
        m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
             "Por favor comuniquese con el laboratorio para su revalida"
        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Terminado.Text = "  -     -   "
        Lote.SetFocus
        Exit Sub
    End If
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    
    Salida = "N"
    Da = 0
    
    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    With rstEtiquetaIII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    With rstEtiquetaIV
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    
    
    
    
    
    If Salida <> "S" Then
            
        XEmpresa = Wempresa
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
        
        TipoPro = "PT"
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If Left$(Terminado.Text, 2) <> "PT" Then
            Select Case Left$(Terminado.Text, 2)
                Case "DY", "DS"
                    TipoPro = "CO"
                Case "QC"
                    TipoPro = "FA"
                Case Else
                    TipoPro = "PT"
            End Select
                Else
            If XCodigo >= 0 And XCodigo <= 999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    TipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        TipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            TipoPro = "BI"
                                Else
                            TipoPro = "PT"
                        End If
                    End If
                End If
            End If
        End If
        
        If Left$(Terminado.Text, 2) = "YQ" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YH" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YP" Then
            TipoPro = "PT"
        End If
        If Left$(Terminado.Text, 2) = "YF" Then
            TipoPro = "FA"
        End If
        
        
        
        
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                TipoPro = "FA"
            End If
        End If
            
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
        
        WVida = 0
        WLinea = 0
        
        DescripcionFarma.Text = ""
        DesProducto.Caption = ""
        ZZImpreVtoTermi = 0
        WObservaciones = ""
                    
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
        
            WLinea = rstTerminado!Linea
            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
            
            If Idioma.ListIndex = 0 Then
            
                DesProducto.Caption = Trim(rstTerminado!Descripcion)
                Descripcion.Text = ""
                If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                    Descripcion.Text = Trim(rstTerminado!Descripcion)
                End If
                DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
                
                WConservacion = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
                WConservacion = RTrim(WConservacion)
                WConservacionII = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
                WConservacionII = RTrim(WConservacionII)
                
                    Else
                    
                DesProducto.Caption = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
                Descripcion.Text = ""
                If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                    Descripcion.Text = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
                End If
                DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiquetaIngles), "", rstTerminado!DescriEtiquetaIngles)
                
                WConservacion = IIf(IsNull(rstTerminado!ConservacionIngles), "", rstTerminado!ConservacionIngles)
                WConservacion = RTrim(WConservacion)
                WConservacionII = IIf(IsNull(rstTerminado!ConservacionIIIngles), "", rstTerminado!ConservacionIIIngles)
                WConservacionII = RTrim(WConservacionII)
                
                DesProducto.Caption = Trim(DesProducto.Caption)
                Descripcion.Text = Trim(Descripcion.Text)
                DescripcionFarma.Text = Trim(DescripcionFarma.Text)
                
            End If
            
            WImpreadi = ""
            WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
            
            WClase = ""
            WRiesgo = ""
            WIntervencion = ""
            WNaciones = ""
            WEmbalaje = ""
            wdescriOnu = ""
            
            WClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
            WRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
            WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
            
            WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
            WObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
            WObservaciones = ""
            WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
            
            rstTerminado.Close
            
        End If
        
        If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
            WVida = 0
        End If
            
        spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            DescripcionFarma.Text = IIf(IsNull(rstPrecios!DescripcionFarma), "", rstPrecios!DescripcionFarma)
            Descripcion.Text = Trim(Left$(rstPrecios!Descripcion, 50))
            rstPrecios.Close
        End If
                    
        spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            DesCliente.Caption = rstClientes!Razon
            rstClientes.Close
        End If
                    
        Call Conecta_Empresa
        
        Wvencimiento = ""
        XEmpresa = Wempresa
        
        Select Case Val(Wempresa)
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
        
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
            
                WMes = Val(Mid$(rstHoja!Fecha, 4, 2))
                WAno = Val(Right$(rstHoja!Fecha, 4))
            
                ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                
                If Val(ZZRevalida) <> 0 Then
                    WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                    WAno = Val(Right$(ZZFechaRevalida, 4))
                    WVida = Val(ZZMesesRevalida)
                End If
            
                For Ciclo = 1 To WVida
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                Next Ciclo
                WElaboracion = rstHoja!Fecha
                If WVida <> 0 Then
                    XMes = Str$(WMes)
                    XAno = Str$(WAno)
                    Call Ceros(XMes, 2)
                    Call Ceros(XAno, 4)
                    Wvencimiento = "01/" + XMes + "/" + XAno
                End If
                rstHoja.Close
                
                ZZRenglon = 0
                ZZTipo = ""
                ZZTerminado = ""
                ZZArticulo = ""
                ZZCantidad = 0
                ZZCantidadLote = 0
                ZZLote = ""
                spHoja = "ListaHoja " + "'" + Lote.Text + "'"
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    With rstHoja
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                ZZRenglon = ZZRenglon + 1
                                ZZTipo = rstHoja!Tipo
                                ZZTerminado = rstHoja!Terminado
                                ZZArticulo = rstHoja!Articulo
                                ZZCantidad = rstHoja!Cantidad
                                ZZCantidadLote = rstHoja!Canti1
                                ZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstHoja.Close
                End If
                
            End If
            
        Next ZCiclo
        
        Call Conecta_Empresa
        
        
        ZZZZLoteOriginal = ""
        If TipoPro = "FA" Then
        
            Rem *******************CARGA DE LOTES MONO CON IMPRESION DE LOTE DE REFERENCIA *************
            Rem  If Terminado.Text = "PT-25062-777" Or Terminado.Text = "PT-25062-778" Or Terminado.Text = "PT-25106-777" Or Terminado.Text = "PT-25106-778" Or Terminado.Text = "PT-25046-777" Or Terminado.Text = "PT-25049-777" Or Terminado.Text = "PT-25136-777" Or Terminado.Text = "PT-25135-777" Or Terminado.Text = "PT-25049-777" Or Terminado.Text = "PT-25046-777" Or Terminado.Text = "PT-25049-778" Or Terminado.Text = "PT-25013-777" Or Terminado.Text = "PT-25057-778" Or Terminado.Text = "PT-25028-102" Or Terminado.Text = "PT-25201-100" Or Terminado.Text = "PT-25046-778" Then
            Rem    If Terminado.Text = "PT-25062-777" Or Terminado.Text = "PT-25062-778" Or Terminado.Text = "PT-25106-777" Or Terminado.Text = "PT-25106-778" Or Terminado.Text = "PT-25046-777" Or Terminado.Text = "PT-25049-777" Or Terminado.Text = "PT-25013-777" Or Terminado.Text = "PT-25135-777" Or Terminado.Text = "PT-25049-777" Or Terminado.Text = "PT-25046-777" Or Terminado.Text = "PT-25049-778" Or Terminado.Text = "PT-25013-777" Or Terminado.Text = "PT-25057-778" Or Terminado.Text = "PT-25028-102" Or Terminado.Text = "PT-25201-100" Or Terminado.Text = "PT-25046-778" Or Terminado.Text = "PT-25057-778" Or Terminado.Text = "PT-25201-100" Or Terminado.Text = "PT-25083-100" Then
                          
                            
            Rem *****************15-01-2016  by nan
            Rem by nan busco en la tabla codigomono si esta cargado el producto
            Rem rutina que busca el producto
            Call busca_mono
            If Pasa = "S" Then
                Call Calcula_Mono_Otro
                If ZZZZElaboracion <> "" Then
                    WElaboracion = ZZZZElaboracion
                End If
                If ZZZZVencimiento <> "" Then
                    Wvencimiento = ZZZZVencimiento
                End If
            End If
            
                Else
                
            WElaboracion = ""
            
       End If
        
         Rem si es este producto no quier que aparezca fecha de
         Rem SE PIDIO 1/2/2016 QUE se vueva agregar fecha y lot de este producto
        
            Rem If Terminado.Text = "PT-25049-777" Then
            Rem  ZZZZLoteOriginal = ""
            Rem  End If
        
        
        
        Rem DEJE ACA
        Rem DEJE ACA
        Rem DEJE ACA
        Rem DEJE ACA
        
        Rem If Tipopro <> "FA" Then
        Rem
        Rem     Rem veo si es mono
        Rem     If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
        Rem
        Rem         ZVto = ""
        Rem         ZLaudo = ZZLote
        Rem         ZArticulo = ZZArticulo
        Rem         ZFecha = ""
        Rem         ZFechaVto = ""
        Rem
        Rem         For ZCiclo = 1 To ZHasta1
        Rem
        Rem             WEmpresa = CargaEmpresa(ZCiclo, 1)
        Rem             txtOdbc = CargaEmpresa(ZCiclo, 2)
        Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem
        Rem             ZSql = ""
        Rem             ZSql = ZSql + "Select *"
        Rem             ZSql = ZSql + " FROM Laudo"
        Rem             ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
        Rem             ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
        Rem             spLaudo = ZSql
        Rem             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstLaudo.RecordCount > 0 Then
        Rem                 ZFecha = rstLaudo!Fecha
        Rem                 ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
        Rem                 rstLaudo.Close
        Rem                 Exit For
        Rem             End If
        Rem
        Rem         Next ZCiclo
        Rem
        Rem         Call Conecta_Empresa
        Rem
        Rem         ZVto = ""
        Rem         ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        Rem         If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
        Rem             Call Valida_fecha(ZFechaVto, Auxi)
        Rem             If Auxi = "S" Then
        Rem                 ZVto = ZFechaVto
        Rem             End If
        Rem         End If
        Rem
        Rem         If ZVto = "" Then
        Rem
        Rem             ZMeses = 0
        Rem             ZSql = ""
        Rem             ZSql = ZSql + "Select *"
        Rem             ZSql = ZSql + " FROM Articulo"
        Rem             ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        Rem             spArticulo = ZSql
        Rem             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstArticulo.RecordCount > 0 Then
        Rem                 ZMeses = rstArticulo!Meses
        Rem                 rstArticulo.Close
        Rem             End If
        Rem
        Rem             WMes = Val(Mid$(ZFecha, 4, 2))
        Rem             WAno = Val(Right$(ZFecha, 4))
        Rem             For ZCiclo = 1 To ZMeses
        Rem                 WMes = WMes + 1
        Rem                 If WMes > 12 Then
        Rem                     WAno = WAno + 1
        Rem                     WMes = 1
        Rem                 End If
        Rem             Next ZCiclo
        Rem
        Rem             XMes = Str$(WMes)
        Rem             XAno = Str$(WAno)
        Rem             Call Ceros(XMes, 2)
        Rem             Call Ceros(XAno, 4)
        Rem             If Val(Left$(ZFecha, 2)) <= 30 Then
        Rem                 If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
        Rem                     ZVto = "28/" + XMes + "/" + XAno
        Rem                         Else
        Rem                     ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
        Rem                 End If
        Rem                     Else
        Rem                 If Val(XMes) = 2 Then
        Rem                     ZVto = "28/" + XMes + "/" + XAno
        Rem                         Else
        Rem                     ZVto = "30/" + XMes + "/" + XAno
        Rem                 End If
        Rem             End If
        Rem
        Rem         End If
        Rem
        Rem         Rem
        Rem         Rem
        Rem         Rem verifica venciminiento
        Rem         Rem
        Rem         Rem
        Rem         Rem
        Rem
        Rem         ZZVidaUtil = 0
        Rem         spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Rem         Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstTerminado.RecordCount > 0 Then
        Rem             ZZVidaUtil = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
        Rem             ZZVidaUtil = Int(ZZVidaUtil * 0.25)
        Rem             rstTerminado.Close
        Rem         End If
        Rem
        Rem         WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Rem         WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
        Rem
        Rem         WFechaVencimiento = "01" + Mid$(ZVto, 3, 10)
        Rem         WFechaVencimientoOrd = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        Rem
        Rem         Pasa = "S"
        Rem         If Left$(WFechaActualOrd, 6) >= Left$(WFechaVencimientoOrd, 6) Then
        Rem
        Rem             Pasa = "N"
        Rem
        Rem                 Else
        Rem
        Rem             Meses = 0
        Rem             WMes = Val(Mid$(WFechaActual, 4, 2))
        Rem             WAno = Val(Right$(WFechaActual, 4))
        Rem             Do
        Rem                 Meses = Meses + 1
        Rem                 WMes = WMes + 1
        Rem                 If WMes > 12 Then
        Rem                     WAno = WAno + 1
        Rem                     WMes = 1
        Rem                 End If
        Rem                 XMes = Str$(WMes)
        Rem                 XAno = Str$(WAno)
        Rem                 Call Ceros(XMes, 2)
        Rem                 Call Ceros(XAno, 4)
        Rem                 WCompara = "01/" + XMes + "/" + XAno
        Rem                 If WCompara = WFechaVencimiento Then
        Rem                     Exit Do
        Rem                 End If
        Rem             Loop
        Rem
        Rem             If ZZVidaUtil >= Meses Then
        Rem                 Pasa = "N"
        Rem             End If
        Rem
        Rem         End If
        Rem
        Rem         If Pasa = "N" Then
        Rem            m$ = "EL Producto tiene menos de 25% de la vida util del PT"
        Rem           G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        Rem           Exit Sub
        Rem         End If
        Rem
        Rem         Wvencimiento = ZVto
        Rem
        Rem     End If
        Rem
        Rem End If
        
        
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                         Exit Do
                    End If
                Loop
            End If
        End With
        
        WTara = Val(Tara.Text)
        WNeto = Val(Cantidad.Text)
        
        If WTara = 0 Then
            WBruto = 0
                Else
            WBruto = WTara + WNeto
        End If
        
        WRazon = ""
        Rem WDirEntrega = ""
                
                
                
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
                
        WImpreVto = 0
        spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            WRazon = "Cliente : " + rstClientes!Razon
            WImpreVto = IIf(IsNull(rstClientes!ImpreVto), "0", rstClientes!ImpreVto)
            WProv = rstClientes!Provincia
            Rem WDirEntrega = rstClientes!DirEntrega
            rstClientes.Close
        End If
        
        ZVencimiento = Wvencimiento
        If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
            Rem no hago nada
                Else
            Rem vencimiento
            If ZZImpreVtoTermi = 0 Then
                If WImpreVto = 0 Then
                    Rem ZVencimiento = ""
                End If
            End If
        End If
                
        Call Conecta_Empresa
        
        Rem If TipoPro = "FA" Then
        Rem     Descripcion.Text = DesProducto.Caption
        Rem End If
        
        
        
        Descripcion.Text = Trim(Descripcion.Text)
        Descripcion.Text = Left$(Descripcion.Text, 50)
        ZZLargo = Len(Descripcion.Text)
        If ZZLargo < 17 Then
            ZZNombre = Descripcion.Text
            ZZNombreII = ""
            ZZNombreIII = ""
                Else
            If ZZLargo <= 30 Then
                ZZNombre = ""
                ZZNombreII = Descripcion.Text
                ZZNombreIII = ""
                    Else
                ZZNombre = ""
                ZZNombreII = ""
                ZZNombreIII = Descripcion.Text
            End If
        End If
        
        
        
        
        WRazon = Trim(WRazon)
        WRazon = Left$(WRazon, 50)
        ZZLargo = Len(WRazon)
        If ZZLargo < 20 Then
            ZRazon = WRazon
            ZRazonII = ""
            ZRazonIII = ""
                Else
            If ZZLargo <= 35 Then
                ZRazon = ""
                ZRazonII = WRazon
                ZRazonIII = ""
                    Else
                ZRazon = ""
                ZRazonII = ""
                ZRazonIII = WRazon
            End If
        End If
        
        
        
        
        Rem atencion
        Rem atencion
        Rem atencion
        Rem ver aca el tema de la razon
        Rem atencion
        Rem atencion
        Rem atencion
        
        
        If Val(WProv) = 24 Then
            If Idioma.ListIndex = 0 Then
                WObservaciones = "Hecho en Argentina"
                    Else
                WObservaciones = "Made in Argentina"
            End If
        End If
            
        Rem If Val(WProv) = 1 Then
        Rem     WObservaciones = "Industria Argentina"
        Rem End If
            
            
        ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
            
            
            
            
            
            
            
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
        
        
        
        
        
        
        
                
                
        For Ciclo = 1 To 999
        
            Auxi = Ciclo
            Call Ceros(Auxi, 3)
            
            ZZClave = "PT" + Mid$(Terminado.Text, 3, 10) + Auxi
        
            Sql1 = "Select *"
            Sql2 = " FROM DatosEtiqueta"
            Sql3 = " Where DatosEtiqueta.Clave = " + "'" + ZZClave + "'"
            spDatosEtiqueta = Sql1 + Sql2 + Sql3
            Set rstDatosEtiqueta = db.OpenRecordset(spDatosEtiqueta, dbOpenSnapshot, dbSQLPassThrough)
            If rstDatosEtiqueta.RecordCount > 0 Then
        
                ZZPalabra = rstDatosEtiqueta!Palabra
                ZZLogo(1) = rstDatosEtiqueta!pictograma1
                ZZLogo(2) = rstDatosEtiqueta!pictograma2
                ZZLogo(3) = rstDatosEtiqueta!pictograma3
                ZZLogo(4) = rstDatosEtiqueta!pictograma4
                ZZLogo(5) = rstDatosEtiqueta!pictograma5
                ZZLogo(6) = rstDatosEtiqueta!pictograma6
                ZZLogo(7) = rstDatosEtiqueta!pictograma7
                ZZLogo(8) = rstDatosEtiqueta!pictograma8
                ZZLogo(9) = rstDatosEtiqueta!pictograma9
                
                Select Case rstDatosEtiqueta!Tipo
                    Case 1
                        If Trim(rstDatosEtiqueta!descripcion1h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiqueta!descripcion1h)
                        End If
                        If Trim(rstDatosEtiqueta!descripcion2h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiqueta!descripcion2h)
                        End If
                        If Trim(rstDatosEtiqueta!descripcion3h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiqueta!descripcion3h)
                        End If
                        
                    Case 2
                        If Trim(rstDatosEtiqueta!descripcion1p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiqueta!descripcion1p)
                        End If
                        If Trim(rstDatosEtiqueta!descripcion2p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiqueta!descripcion2p)
                        End If
                        If Trim(rstDatosEtiqueta!descripcion3p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiqueta!descripcion3p)
                        End If
                        If Trim(rstDatosEtiqueta!Observaciones) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiqueta!Observaciones)
                        End If
                        
                    Case 3
                        If Trim(rstDatosEtiqueta!denominacion) <> "" Then
                            ZLugarImpreIII = ZLugarImpreIII + 1
                            ZImpreIII(ZLugarImpreIII) = Trim(rstDatosEtiqueta!denominacion)
                        End If
                        
                    Case Else
                End Select
                
                    Else
                    
                Exit For
                
            End If
                    
        Next Ciclo
    
        Rem ZLugarH = 0
        Rem For Ciclo = 1 To 99
        Rem     If Trim(ZImpreI(Ciclo)) <> "" Then
        Rem         Sql1 = "Select *"
        Rem         Sql2 = " FROM FraseH"
        Rem         Sql3 = " Where FraseH.Codigo = " + "'" + ZImpreI(Ciclo) + "'"
        Rem         spFraseH = Sql1 + Sql2 + Sql3
        Rem         Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstFraseH.RecordCount > 0 Then
        Rem             If Trim(rstFraseH!Descripcion) <> "" Then
        Rem                 ZLugarH = ZLugarH + 1
        Rem                 ZImpreI(ZLugarH) = rstFraseH!Descripcion
        Rem             End If
        Rem             If Trim(rstFraseH!DescripcionII) <> "" Then
        Rem                 ZLugarH = ZLugarH + 1
        Rem                 ZImpreI(ZLugarH) = rstFraseH!DescripcionII
        Rem             End If
        Rem             If Trim(rstFraseH!DescripcionIII) <> "" Then
        Rem                 ZLugarH = ZLugarH + 1
        Rem                 ZImpreI(ZLugarH) = rstFraseH!DescripcionIII
        Rem             End If
        Rem             rstFraseH.Close
        Rem         End If
        Rem     End If
        Rem Next Ciclo
    
        Call Conecta_Empresa
        
        Erase ZZImpreFrase
        LugarFrase = 1
        
        Rem If TipoNuevo.ListIndex = 0 Then
        Rem     ZZCorte = 170
        Rem         Else
        Rem     ZZCorte = 185
        Rem End If
        
        If TipoPro <> "CO" Then
            ZZCorte = 185
                Else
            ZZCorte = 115
        End If
        
              Rem para pellital
        Select Case Val(Wempresa)
            Case 2, 4, 8, 9
                ZZCorte = 185
            Case Else
        End Select
    
        
        
        
        ZZEntraVarios = "N"
        
        For Ciclo = 1 To 99
        
            If Trim(ZImpreIII(Ciclo)) <> "" Then
            
                ZZEntraVarios = "S"
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreIII(Ciclo)) + " "
                
                If Len(ZZImpreFrase(LugarFrase)) > ZZCorte Then
                    
                Do
                
                    ZZHastaIII = Len(ZZImpreFrase(LugarFrase))
                    
                    ZZHastaII = Len(ZZImpreFrase(LugarFrase))
                    For Da = 1 To ZZHastaIII
                        If Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) >= 65 And Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) <= 90 Then
                            ZZHastaII = ZZHastaII + 0.5
                        End If
                    Next Da
                
                    If ZZHastaII > ZZCorte Then
                    
                        For Da = ZZHastaIII - 1 To 1 Step -1
                            If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "-" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "+" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "," Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "/" Then

                                Auxi = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                ZZHastaIII = Len(Auxi)
                                ZZHastaII = 0
                                For DaIII = 1 To ZZHastaIII
                                    ZZHastaII = ZZHastaII + 1
                                    If Asc(Mid$(Auxi, DaIII, 1)) >= 65 And Asc(Mid$(Auxi, DaIII, 1)) <= 90 Then
                                        ZZHastaII = ZZHastaII + 0.5
                                    End If
                                Next DaIII
                                If ZZHastaII <= ZZCorte Then
                                    Auxi = ZZImpreFrase(LugarFrase)
                                    ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                    LugarFrase = LugarFrase + 1
                                    ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Mid$(Auxi, Da + 1, ZZCorte)
                                    Exit For
                                End If
                                
                            End If
                        Next Da

                            Else
                            
                        Exit Do
                        
                    End If
                Loop
                    
                End If
            End If
        
        Next Ciclo
        
        If ZZEntraVarios = "S" Then
            LugarFrase = LugarFrase + 1
        End If
        
        
        
        
        
        
        
        
        
        ZZEntraH = "N"
        
        For Ciclo = 1 To 99
        
            If Trim(ZImpreI(Ciclo)) <> "" Then
            
                If ZZEntraH = "N" Then
                    ZZImpreFrase(LugarFrase) = "INDICACIONES DE PELIGRO : "
                End If
                ZZEntraH = "S"
                AA1 = Trim(ZImpreI(Ciclo))
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreI(Ciclo)) + " "
                
                Do
                
                    ZZHastaIII = Len(ZZImpreFrase(LugarFrase))
                    
                    ZZHastaII = Len(ZZImpreFrase(LugarFrase))
                    For Da = 1 To ZZHastaIII
                        If Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) >= 65 And Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) <= 90 Then
                            ZZHastaII = ZZHastaII + 0.5
                        End If
                    Next Da
                
                    If ZZHastaII > ZZCorte Then
                    
                        For Da = ZZHastaIII - 1 To 1 Step -1
                            If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "-" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "+" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "," Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "/" Then

                                Auxi = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                ZZHastaIII = Len(Auxi)
                                ZZHastaII = 0
                                For DaIII = 1 To ZZHastaIII
                                    ZZHastaII = ZZHastaII + 1
                                    If Asc(Mid$(Auxi, DaIII, 1)) >= 65 And Asc(Mid$(Auxi, DaIII, 1)) <= 90 Then
                                        ZZHastaII = ZZHastaII + 0.5
                                    End If
                                Next DaIII
                                If ZZHastaII <= ZZCorte Then
                                    Auxi = ZZImpreFrase(LugarFrase)
                                    ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                    LugarFrase = LugarFrase + 1
                                    ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Mid$(Auxi, Da + 1, ZZCorte)
                                    Exit For
                                End If
                                
                            End If
                        Next Da

                            Else
                            
                        Exit Do
                        
                    End If
                Loop
            End If
        
        Next Ciclo
            
        If ZZEntraH = "S" Then
            LugarFrase = LugarFrase + 1
        End If
        
        
        
        
        
        
        
        
        ZZEntraP = "N"
        
        For Ciclo = 1 To 99
        
            If Trim(ZImpreII(Ciclo)) <> "" Then
            
                If ZZEntraP = "N" Then
                    ZZImpreFrase(LugarFrase) = "DECLARACIONES DE PRUDENCIA : "
                End If
                ZZEntraP = "S"
                AA1 = Trim(ZImpreII(Ciclo))
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreII(Ciclo)) + " "
                
                Do
                
                    ZZHastaIII = Len(ZZImpreFrase(LugarFrase))
                    
                    ZZHastaII = Len(ZZImpreFrase(LugarFrase))
                    For Da = 1 To ZZHastaIII
                        If Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) >= 65 And Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) <= 90 Then
                            ZZHastaII = ZZHastaII + 0.5
                        End If
                    Next Da
                
                    If ZZHastaII > ZZCorte Then
                    
                        For Da = ZZHastaIII - 1 To 1 Step -1
                            If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "-" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "+" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "," Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "/" Then
                            
                                Auxi = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                ZZHastaIII = Len(Auxi)
                                ZZHastaII = 0
                                For DaIII = 1 To ZZHastaIII
                                    ZZHastaII = ZZHastaII + 1
                                    If Asc(Mid$(Auxi, DaIII, 1)) >= 65 And Asc(Mid$(Auxi, DaIII, 1)) <= 90 Then
                                        ZZHastaII = ZZHastaII + 0.5
                                    End If
                                Next DaIII
                                If ZZHastaII <= ZZCorte Then
                                    Auxi = ZZImpreFrase(LugarFrase)
                                    ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                    LugarFrase = LugarFrase + 1
                                    ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Mid$(Auxi, Da + 1, ZZCorte)
                                    Exit For
                                End If
                                
                            End If
                        Next Da

                            Else
                            
                        Exit Do
                        
                    End If
                Loop
            End If
        
        Next Ciclo
        
        
            
            
            
            
            
        For Ciclo = 1 To 19
        
            If Trim(ZZImpreFrase(Ciclo)) <> "" Then
            
                For CicloII = 1 To ZZHasta
                
                    If Mid$(ZZImpreFrase(Ciclo), CicloII, 1) = Space$(1) Then
                        ZZImpreFrase(Ciclo) = Mid$(ZZImpreFrase(Ciclo), 1, CicloII) + " " + Mid$(ZZImpreFrase(Ciclo), CicloII + 1, ZZCorte)
                        ZZHasta = Len(Trim(ZZImpreFrase(Ciclo)))
                        CicloII = CicloII + 1
                        If CicloII = ZZCorte Or ZZHasta = ZZCorte Then
                            Exit For
                        End If
                    End If
                    
                Next CicloII
                
                ZZImpreFrase(Ciclo) = Trim(ZZImpreFrase(Ciclo))
                
                
            End If
            
        Next Ciclo
                        
                    
            
        
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        With rstEtiquetaII
            For Da = 1 To Val(Etiquetas)
                .Index = "Codigo"
                .AddNew
                !Codigo = Da
                WLote = Lote.Text
                Call Ceros(WLote, 6)
                WCantidad = Cantidad.Text
                Call Ceros(WCantidad, 4)
                !Terminado = Terminado.Text
                !Lote = Val(Lote.Text)
                !Cliente = Cliente.Text
                !Cantidad = Val(Cantidad.Text)
                !Nombre = ZZNombre
                !NombreII = ZZNombreII
                !NombreIII = ZZNombreIII
                !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + " " + WLote
                !Impre2 = ""
                
                If Trim(ZZZZLoteOriginal) <> "" Then
                    !Impre2 = Left$(ZZZZLoteOriginal, 30)
                End If
             
                If Tipo.ListIndex = 6 Or Tipo.ListIndex = 7 Then
                    !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3)
                    !Impre2 = WLote
                End If
             
             
             
             
             
                !Razon = ZRazon
                !RazonII = ZRazonII
                !RazonIII = ZRazonIII
                Rem !DirEntrega = ZRazonII
                !Clase = WClase
                !Intervencion = WIntervencion
                If TipoProceso.ListIndex = 0 Then
                    If wdescriOnu = "" Then
                        wdescriOnu = "0"
                    End If
                    !Descrionu = wdescriOnu
                End If
                
                !Naciones = WNaciones
                !Embalaje = WEmbalaje
                !Bruto = WBruto
                !Tara = WTara
                !Neto = WNeto
                !Observaciones = Left$(WObservaciones, 20)
                !Elaboracion = Right$(WElaboracion, 7)
                !Vencimiento = Right$(ZVencimiento, 7)
                If Val(Pedido.Text) = 0 Then
                    !Conservacion = Trim(WConservacion)
                        Else
                    !Conservacion = Trim(WConservacionII)
                End If
                !ConservacionII = Trim(WConservacionII)
                
                If TipoPro = "FA" Then
                    !NombreFarmaII = DescripcionFarma.Text
                        Else
                    !NombreFarmaII = ""
                End If
                
                
                Rem If Tipopro = "FA" Then
                Rem     !NombreFarmaI = DesProducto.Caption
                Rem     !NombreFarmaII = DescripcionFarma.Text
                Rem     Rem If Tipo.ListIndex = 6 Then
                Rem     Rem     !NombreFarmaI = ZZNombre
                Rem     Rem     !NombreFarmaII = ZZNombreII
                Rem     Rem End If
                Rem         Else
                Rem     Rem by nan 23-3-2011
                Rem     conserva = !ConservacionII
                Rem     !NombreFarmaI = "MANTENER EN ENVASE ORIGINAL CERRADO"
                Rem     !NombreFarmaI = conserva
                Rem     !NombreFarmaII = ""
                Rem     If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 9 Then
                Rem         If Cliente.Text = "Z00007" Then
                Rem             !NombreFarmaII = "PRODUCTO PARA LA PRODUCCION DE CUERO"
                Rem         End If
                Rem     End If
                Rem End If
                
                !TipoPro = ""
                If Trim(Cliente.Text) = "" Then
                    !TipoPro = Left$(Terminado.Text, 2)
                End If
                If Tipo.ListIndex = 7 Then
                    !TipoPro = Left$(Terminado.Text, 2)
                End If
                
                ZFazon = "N"
                Select Case Val(WLinea)
                    Case 3, 4, 5, 7, 8, 9, 11, 12, 14, 19, 22
                        ZFazon = "N"
                    Case 6, 16, 17
                        ZFazon = "N"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        ZFazon = "N"
                    Case Else
                        ZFazon = "S"
                End Select
                
                Rem If Tipopro = "CO" Then
                Rem     !NombreFarmaI = ""
                Rem     !NombreFarmaII = ""
                Rem End If
                Rem If ZFazon = "S" Then
                Rem     !NombreFarmaI = ""
                Rem     !NombreFarmaII = ""
                Rem End If
                
                !ImpreOc = ""
                Rem !ImpreDirEntrega = ""
                
                If ZEtiI = 1 Then
                    If Trim(ZZOrdenCpa) <> "" Then
                        !ImpreOc = "Orden Cpa.:" + ZZOrdenCpa
                    End If
                End If
                Rem If ZEtiII = 1 Then
                Rem     !ImpreDirEntrega = ZDescriDirEntrega
                Rem End If
                
                
                !foto1 = 0
                !foto2 = 0
                !foto3 = 0
                !foto4 = 0
                !foto5 = 0
                
                For Ciclo = 1 To 9
                    If ZZLogo(Ciclo) <> 0 Then
                        Select Case ZZLogo(Ciclo)
                            Case 1
                                !foto1 = Ciclo
                            Case 2
                                !foto2 = Ciclo
                            Case 3
                                !foto3 = Ciclo
                            Case 4
                                !foto4 = Ciclo
                            Case 5
                                !foto5 = Ciclo
                            Case Else
                        End Select
                    End If
                Next Ciclo
                
                Rem by nan agrego campo para imprimir numero eti
                Rem lo pruebo para farma
                If Wempresa = "0005" Then
                    If Check1.Value = 0 Then
                        !ImpreNumero = ""
                            Else
                        Rem !ImpreNumero = Str$(ZZDesdeNumero) + "/" + Str$(ZZHastaNumero)
                        !ImpreNumero = Str$(ZZDesdeNumero)
                        ZZDesdeNumero = ZZDesdeNumero + 1
                    End If
                End If
                Rem en by nan
                
                .Update
            Next Da
        End With
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
            
            
            
            
            
            
            
            
            
            
        With rstEtiquetaIII
            For Da = 1 To Val(Etiquetas)
                .Index = "Codigo"
                .AddNew
                !Codigo = Da
                
                aa = Len(ZZImpreFrase(1))
                
                !Frase1 = ZZImpreFrase(1)
                !Frase2 = ZZImpreFrase(2)
                !Frase3 = ZZImpreFrase(3)
                !Frase4 = ZZImpreFrase(4)
                !Frase5 = ZZImpreFrase(5)
                !Frase6 = ZZImpreFrase(6)
                !Frase7 = ZZImpreFrase(7)
                !Frase8 = ZZImpreFrase(8)
                !Frase9 = ZZImpreFrase(9)
                !Frase10 = ZZImpreFrase(10)
                
                
                .Update
            Next Da
        End With
        
        
        
        
        
        
        
            
            
            
        With rstEtiquetaIV
            For Da = 1 To Val(Etiquetas)
                .Index = "Codigo"
                .AddNew
                !Codigo = Da
                
                !Frase20 = ""
                If ZZPalabra = 1 Then
                    !Frase20 = "PELIGRO"
                End If
                If ZZPalabra = 2 Then
                    !Frase20 = "ATENCION"
                End If
                
                !Frase11 = ZZImpreFrase(11)
                !Frase12 = ZZImpreFrase(12)
                !Frase13 = ZZImpreFrase(13)
                !Frase14 = ZZImpreFrase(14)
                !Frase15 = ZZImpreFrase(15)
                !Frase16 = ZZImpreFrase(16)
                !Frase17 = ZZImpreFrase(17)
                !Frase18 = ZZImpreFrase(18)
                !Frase19 = ZZImpreFrase(19)
                
                !Frase20 = ""
                If ZZPalabra = 1 Then
                    !Frase20 = "PELIGRO"
                End If
                If ZZPalabra = 2 Then
                    !Frase20 = "ATENCION"
                End If
                
                
                .Update
            Next Da
        End With
        
        
        
        
        
        
        
        
        
        
   
        
    
        Listado.WindowTitle = "Emision de Etiquetas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Rem If TipoNuevo.ListIndex = 0 Then
        Rem     Listado.ReportFileName = "EtiNuevaNorma.RPT"
        Rem         Else
        Rem     Listado.ReportFileName = "EtiNuevaNormaChica.RPT"
        Rem End If
        
        Select Case Terminado.Text
            Case "SE-00300-050", "SE-00200-072", "SE-00161-150", "SE-00165-100", "SE-00463-151", "SE-00162-150"
                Rem paa estos productos fuerzo que tome
                Rem la etiqueta comun
                TipoPro = ""
            Case Else
        End Select
        
        If TipoPro = "CO" Then
            Listado.ReportFileName = "EtiNuevaColorante.RPT"
                Else
            If Left$(Terminado.Text, 5) = "PT-25" Then
                Listado.ReportFileName = "EtiNuevaNormaChica25000.RPT"
                    Else
                Listado.ReportFileName = "EtiNuevaNormaChica.RPT"
            End If
        End If
        
        Rem ************para pellital********************
        Select Case Val(Wempresa)
               Case 2, 4, 8, 9
         Rem   Listado.ReportFileName = "EtiNuevaNormaChica.RPT"
            Listado.ReportFileName = "EtiNuevaNormaChicapellital.RPT"
           
           Case Else
           
        End Select
       
        
       
            
        XCodigo = Val(Mid$(Terminado.Text, 4, 5))
            
        
        If TipoPro <> "CO" Then
            If Trim(WClase) <> "" Then
                m$ = " Coloque la etiqueta de producto Peligroso Clase Nro.: " + WClase
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            End If
        End If
        
        Rem Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"21.
        Rem Listado.DataFiles(1) = ""
        
        
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.PrinterCopies = 1
        Listado.Action = 1
        
        
        Da = 0
        With rstEtiquetaII
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
        Da = 0
        With rstEtiquetaIII
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
        Da = 0
        With rstEtiquetaIV
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
        Call Limpia_Click

        
    
    End If
    
    Exit Sub

WError:

    Resume Next

End Sub

Sub Form_Load()

    XEmpresa = Wempresa

    TipoProceso.Clear
    
    TipoProceso.AddItem "Etiqueta Nueva"
    TipoProceso.AddItem "Etiqueta Anterior"
    
    TipoProceso.ListIndex = 0

    Tipo.Clear
    Rem************
    tipofarma.Clear
   
    Rem by nan agrego cambios para farma
    If Wempresa = "0005" Then
        Tipo.Visible = False
        tipofarma.AddItem "Farma Grande                  ZT-330-100"
        tipofarma.AddItem "Farma Chica Autoadhesiva      ZT-211-100"
        tipofarma.AddItem "Farma Producto en Proceso     ZT-211-201"
        Rem debe ser 5
        tipofarma.AddItem "Certificado"
        Rem debe ser 3
        tipofarma.AddItem "Certificado (Pantalla)"
        Rem debe ser 4
        tipofarma.ListIndex = 0
            Else
        tipofarma.Visible = False
    End If
              
              
              
           
    Rem FIN BY NAN******************
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    Tipo.AddItem "Blanca"
    Tipo.AddItem "Certificado"
    Tipo.AddItem "Certificado (Pantalla)"
    Tipo.AddItem "Producto en Proceso"
    Tipo.AddItem "Etiqueta Interior"
    Tipo.AddItem "Etiqueta Autoadhesivas"
    Tipo.AddItem "Etiqueta Pigmentos"
   
   
    Select Case Val(XEmpresa)
        Case 5
            Tipo.ListIndex = 1
        Case Else
            Tipo.ListIndex = 0
    End Select
    
    Idioma.Clear
    
    Idioma.AddItem "Castellano"
    Idioma.AddItem "Ingles"
    
    Idioma.ListIndex = 0

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    Tara.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Pedido.Text = ""
    LoteMP.Text = ""
    
    Pedido.Visible = False
    PedidoII.Visible = False

    Rem para imprimir numero de etiquetas

    Label10.Visible = False
    Label11.Visible = False
    DesdeNumero.Visible = False
    HastaNumero.Visible = False
    Check1.Value = 0
    Check2.Value = 0
    Check2.Visible = False

End Sub

Private Sub ImpreCaratula_Click()

    If Trim(Cliente.Text) <> "" And Val(Pedido.Text) <> 0 Then

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

        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZZCliente = rstPedido!Cliente
            ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            ZDescriDirEntrega = ""
            
            rstPedido.Close
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    If Trim(ZDescriDirEntrega) = "" Then
                        ZDescriDirEntrega = Trim(rstCliente!Direccion) + " - " + Trim(rstCliente!Localidad)
                    End If
                    rstCliente.Close
                End If
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Cliente SET "
                ZSql = ZSql & "ImpreDireccion = " + "'" + Left$(ZDescriDirEntrega, 100) + "'"
                ZSql = ZSql & " Where Cliente = " + "'" + Cliente.Text + "'"
                spCliente = ZSql
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                            
                Listado.WindowTitle = "Caratula"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
    
                If Val(Etiquetas.Text) = 0 Then
                    Etiquetas.Text = "1"
                End If
    
                Listado.Destination = 1
                Listado.CopiesToPrinter = Val(Etiquetas.Text)
                Rem Listado.Destination = 0
                            
                If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                    Listado.ReportFileName = "ImpreCaratulaSurfa.rpt"
                        Else
                    Listado.ReportFileName = "ImpreCaratulaPelli.rpt"
                End If
                
                Listado.GroupSelectionFormula = "{Cliente.Cliente} in " + Chr$(34) + Cliente.Text + Chr$(34) + " to " + Chr$(34) + Cliente.Text + Chr$(34)
                                
                DbConnect = db.Connect
                DSQ = getDatabase(DbConnect)

                Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.ImpreDireccion " _
                            + "From " _
                            + DSQ + ".dbo.Cliente Cliente " _
                            + "Where " _
                            + "Cliente.Cliente >= '" + Cliente.Text + "' AND " _
                            + "Cliente.Cliente <= '" + Cliente.Text + "'"
                                    
                Listado.Connect = Connect()
                Listado.Action = 1
                
                Listado.CopiesToPrinter = 1
                
                Call Conecta_Empresa
                
                    Else
                    
                Call Conecta_Empresa
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Exit Sub
                
            End If
            
                Else
                
            Call Conecta_Empresa
            m$ = "Pedido inexistente"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            Exit Sub
            
        End If
        
    End If

End Sub

Sub Limpia_Click()

    If Wempresa = "0005" Then
        tipofarma.Visible = True
        Tipo.Visible = False
            Else
        Tipo.Visible = True
        tipofarma.Visible = False
    End If
    
    tipofarma.Clear
    tipofarma.AddItem "Farma Grande                        ZT-330-100"
    tipofarma.AddItem "Farma Chica Autoadhesiva      ZT-211-100"
    tipofarma.AddItem "Farma Producto en Proceso    ZT-211-201"
    Rem debe ser 5
    tipofarma.AddItem "Certificado"
    Rem debe ser 3
    tipofarma.AddItem "Certificado (Pantalla)"
    Rem debe ser 4
    
    
    
   Rem fin by nan
    
   
    Tipo.Clear
    
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    Tipo.AddItem "Blanca"
    Tipo.AddItem "Certificado"
    Tipo.AddItem "Certificado (Pantalla)"
    Tipo.AddItem "Producto en Proceso"
    Tipo.AddItem "Etiqueta Interior"
    Tipo.AddItem "Etiqueta Autoadhesivas"
    Tipo.AddItem "Etiqueta Pigmentos"
    
    
    Select Case Val(XEmpresa)
        Case 5
            Tipo.ListIndex = 1
        Case Else
            Tipo.ListIndex = 0
    End Select
    
    Idioma.Clear
    
    Idioma.AddItem "Castellano"
    Idioma.AddItem "Ingles"
    
    Idioma.ListIndex = 0

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    Tara.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Pedido.Text = ""
    LoteMP.Text = ""
    
    Pedido.Visible = False
    PedidoII.Visible = False
    
    Lote.SetFocus
            
    Rem limpio impresion numero eti
    Check1.Value = 0
    Label10.Visible = False
    Label11.Visible = False
    DesdeNumero.Visible = False
    HastaNumero.Visible = False
    DesdeNumero.Text = ""
    HastaNumero.Text = ""
    
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)

    Rem On Error GoTo WError

    If KeyAscii = 13 Then
    
        LugarTerminado = 0
        
        Terminado.Text = "  -     -   "
        Ingresa = "N"
        
        spHoja = "ListaHoja " + "'" + Lote.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            LugarTerminado = LugarTerminado + 1
            WTerminado(LugarTerminado) = UCase(rstHoja!Producto)
            rstHoja.Close
        End If
            
        spMovguia = "ListaMovguiaLoteSolo " + "'" + Lote.Text + "'"
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            With rstMovguia
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaTerminado = "S"
                        For Ciclo = 1 To LugarTerminado
                            If WTerminado(Ciclo) = rstMovguia!Terminado Then
                                IngresaTerminado = "N"
                                Exit For
                            End If
                        Next Ciclo
                        If IngresaTerminado = "S" And rstMovguia!Tipo = "T" Then
                            LugarTerminado = LugarTerminado + 1
                            WTerminado(LugarTerminado) = rstMovguia!Terminado
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstMovguia.Close
        End If
        
        If LugarTerminado = 1 Then
            Terminado.Text = WTerminado(1)
            Ingresa = "S"
        End If
            
        If LugarTerminado > 1 Then
            Call Elije_Lote
        End If
            
        If Ingresa = "N" Then
            Lote.SetFocus
                Else
            Call Ejecuta_Lote
        End If
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Ejecuta_Lote()

    Rem On Error GoTo WError

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
                
    
    DescripcionFarma.Text = ""
    DesProducto.Caption = ""
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        WLinea = rstTerminado!Linea
        If Idioma.ListIndex = 0 Then
        
            DesProducto.Caption = Trim(rstTerminado!Descripcion)
            Descripcion.Text = ""
            If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                Descripcion.Text = Trim(rstTerminado!Descripcion)
            End If
            DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
            
            If Val(XEmpresa) = 5 Then
                Descripcion.Text = Trim(rstTerminado!Descripcion) + " - " + DescripcionFarma.Text
                DesProducto.Caption = Trim(rstTerminado!Descripcion) + " - " + DescripcionFarma.Text
            End If
            
                Else
                
            DesProducto.Caption = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
            Descripcion.Text = ""
            If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                Descripcion.Text = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
            End If
            DescripcionFarma.Text = IIf(IsNull(rstTerminado!DescriEtiquetaIngles), "", rstTerminado!DescriEtiquetaIngles)
            
            DesProducto.Caption = Trim(DesProducto.Caption)
            Descripcion.Text = Trim(Descripcion.Text)
            DescripcionFarma.Text = Trim(DescripcionFarma.Text)
            
            If Val(XEmpresa) = 5 Then
                Descripcion.Text = Trim(Descripcion.Text) + " - " + DescripcionFarma.Text
                DesProducto.Caption = Trim(Descripcion.Text) + " - " + DescripcionFarma.Text
            End If
            
        End If
        
        WImpreadi = ""
        WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
        
        WClase = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        wdescriOnu = ""
        
        WClase = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
        WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
        WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
        WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
        wdescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
        
        WTipoeti = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
        WObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
        WObservaciones = ""
        rstTerminado.Close
    End If
    
    TipoPro = "PT"
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If Left$(Terminado.Text, 2) <> "PT" Then
        Select Case Left$(Terminado.Text, 2)
            Case "DY", "DS"
                TipoPro = "CO"
            Case "QC"
                TipoPro = "FA"
            Case Else
                TipoPro = "PT"
        End Select
            Else
        If XCodigo >= 0 And XCodigo <= 999 Then
            TipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                TipoPro = "CO"
                    Else
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    TipoPro = "FA"
                        Else
                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                        TipoPro = "BI"
                            Else
                        TipoPro = "PT"
                    End If
                End If
            End If
        End If
    End If
    
    If Left$(Terminado.Text, 2) = "YQ" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YH" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YP" Then
        TipoPro = "PT"
    End If
    If Left$(Terminado.Text, 2) = "YF" Then
        TipoPro = "FA"
    End If
    
    
    
    
    XCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If (XCodigo >= 25000 And XCodigo <= 25999) Or WLinea = 10 Or WLinea = 20 Or WLinea = 22 Or WLinea = 24 Or WLinea = 25 Or WLinea = 26 Or WLinea = 27 Or WLinea = 28 Or WLinea = 29 Or WLinea = 30 Then
        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
            TipoPro = "FA"
        End If
    End If
    
                
    spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Descripcion.Text = Trim(Left$(rstPrecios!Descripcion, 50))
        rstPrecios.Close
    End If
                
    Call Conecta_Empresa
                
    Cliente.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Elije_Lote()

    Pantalla.Clear
    
    For Ciclo = 1 To LugarTerminado
        spTerminado = "ConsultaTerminado " + "'" + WTerminado(Ciclo) + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDesTerminado = Trim(rstTerminado!Descripcion)
            rstTerminado.Close
        End If
    
        Pantalla.AddItem WTerminado(Ciclo) + "   " + WDesTerminado
        
    Next Ciclo
    
    Pantalla.Visible = True
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Terminado.Text = Left$(Pantalla.Text, 12)
    Call Ejecuta_Lote
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
        
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
            
            ZEtiI = 0
            ZEtiII = 0
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
            
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                ZEtiI = Trim(IIf(IsNull(rstClientes!EtiI), "0", rstClientes!EtiI))
                ZEtiII = Trim(IIf(IsNull(rstClientes!EtiII), "0", rstClientes!EtiII))
                
                Rem If ZEtiI = 1 Then
                Rem     Pedido.Visible = True
                Rem     PedidoII.Visible = True
                Rem     Pedido.SetFocus
                Rem         Else
                Rem     Pedido.Visible = False
                Rem     PedidoII.Visible = False
                Rem     Cantidad.SetFocus
                Rem End If
                
                Pedido.Visible = True
                PedidoII.Visible = True
                Pedido.SetFocus
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstClientes!DirEntrega
                Rem ZDirEntrega(2) = Trim(IIf(IsNull(rstClientes!DirEntregaII), "", rstClientes!DirEntregaII))
                Rem ZDirEntrega(3) = Trim(IIf(IsNull(rstClientes!DirEntregaIII), "", rstClientes!DirEntregaIII))
                Rem ZDirEntrega(4) = Trim(IIf(IsNull(rstClientes!DirEntregaIV), "", rstClientes!DirEntregaIV))
                Rem ZDirEntrega(5) = Trim(IIf(IsNull(rstClientes!DirEntregaV), "", rstClientes!DirEntregaV))
                ZDirEntrega(2) = ""
                ZDirEntrega(3) = ""
                ZDirEntrega(4) = ""
                ZDirEntrega(5) = ""
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                        Else
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                End If
                
                rstClientes.Close
                
                spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Descripcion.Text = Trim(Left$(rstPrecios!Descripcion, 50))
                    rstPrecios.Close
                        Else
                    Descripcion.Text = ""
                    If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                        Descripcion.Text = DesProducto.Caption
                    End If
                End If
            End If
            
            Call Conecta_Empresa
            
                Else
                
            Descripcion.Text = ""
            If Val(XEmpresa) = 2 Or Val(XEmpresa) = 4 Or Val(XEmpresa) = 8 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 9 Then
                Descripcion.Text = DesProducto.Caption
            End If
            
            If TipoPro = "CO" Then
                Descripcion.Text = DesProducto.Caption
            End If
            
            ZEtiI = 0
            ZEtiII = 0
            
            Pedido.Visible = False
            PedidoII.Visible = False
            
            Rem If Val(XEmpresa) = 5 Then
            Rem     Descripcion.Text = DesProducto.Caption + " - " + DescripcionFarma.Text
            Rem     DesProducto.Caption = DesProducto.Caption + " - " + DescripcionFarma.Text
            Rem End If
            
            Cantidad.SetFocus
            
        End If
        
    End If
    
End Sub

Private Sub ListaDirEntrega_Click()
    ZLugarDirEntrega = ListaDirEntrega.ListIndex + 1
    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
    PantaDirEntrega.Visible = False
    Pedido.SetFocus
End Sub


Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lote.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Etiquetas.SetFocus
    End If
End Sub

Private Sub Etiquetas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tara.SetFocus
    End If
End Sub



Private Sub Tara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Pasa = 0
        XEmpresa = Wempresa
        
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select

        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZZCliente = rstPedido!Cliente
            ZZOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
            ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            
            rstPedido.Close
            Pasa = 1
            
                   Else
            
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Rem BY NAN BUSCO EN 8
                OPEN_FILE_Pedido2
                With rstPedido2
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    .Index = "Clave"
                    .Seek ">=", Auxi
                    If .NoMatch = False Then
                        ZZCliente = rstPedido2!Cliente
                        Pasa = 1
                            Else
                        Pasa = 0
                    End If
                End With
                Rem BY NAN FIN
              rstPedido2.Close
            End If
            
        End If
            
        If Pasa = 1 Then
            
            If Trim(UCase(Cliente.Text)) = Trim(UCase(ZZCliente)) Then
                
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    Rem ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    Rem ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    Rem ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    Rem ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                    ZDirEntrega(2) = ""
                    ZDirEntrega(3) = ""
                    ZDirEntrega(4) = ""
                    ZDirEntrega(5) = ""
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                    rstCliente.Close
                End If
                
                Cantidad.SetFocus
                
                    Else
                    
                m$ = "El pedido no corresponde al cliente informado"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Pedido.Text = ""
                
            End If
        End If
            
            
        If Pasa = 0 Then
            
            m$ = "Pedido Inexistente"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            
        End If
    
        Call Conecta_Empresa
    
    End If
    
    If KeyAscii = 27 Then
        Pedido.Text = ""
    End If

End Sub

Private Sub LoteMp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LoteMP.Text = ""
    End If
End Sub


Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem BY NAN
    
    If Wempresa = "0005" Then
        tipofarma.Clear
        tipofarma.Visible = True
        Tipo.Visible = False
        tipofarma.AddItem "Farma Grande                        ZT-330-100"
        tipofarma.AddItem "Farma Chica Autoadhesiva      ZT-211-100"
        tipofarma.AddItem "Farma Producto en Proceso     ZT-211-201"
        Rem debe ser 5
        tipofarma.AddItem "Certificado"
        Rem debe ser 3
        tipofarma.AddItem "Certificado (Pantalla)"
        tipofarma.ListIndex = 0
        Rem debe ser 4
            Else
        tipofarma.Visible = False
        Tipo.Visible = True
    End If
              
    Tipo.Clear
             
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    Tipo.AddItem "Blanca"
    Tipo.AddItem "Certificado"
    Tipo.AddItem "Certificado (Pantalla)"
    Tipo.AddItem "Producto en Proceso"
    Tipo.AddItem "Etiqueta Interior"
    Tipo.AddItem "Etiqueta Autoadhesivas"
    Tipo.AddItem "Etiqueta Pigmentos"
    Tipo.ListIndex = 0

 
    
    Rem FIN BY NAN
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
    OPEN_FILE_EtiquetaII
    OPEN_FILE_EtiquetaIII
    OPEN_FILE_EtiquetaIV

End Sub

Private Sub Imprime_Certificado()

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
    
    ZArticulo = Terminado.Text
    ZProducto = Terminado.Text
    ZLote = Lote.Text
    ZCantidad = Cantidad.Text
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
                    
            If Val(Lote.Text) > 99999 Then
                ZZLote = ZLote
                Call Ceros(ZZLote, 6)
                    Else
                ZZLote = ZLote
                Call Ceros(ZLote, 5)
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
                    
                If rstPrueter!Producto <> Terminado.Text Then
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
                ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Lote.Text + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                    
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    
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
                    
                    
                    Call Calcula_Mono
                    
                    If ZZZZVencimiento <> "" Then
                        WMes = Val(Mid$(ZZZZVencimiento, 4, 2))
                        WAno = Val(Right$(ZZZZVencimiento, 4))
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
                                    m$ = "ATENCION : La partida esta asociada a una version de especificaciones que no es la actual" + Chr$(13) + _
                                         "Version " + Str$(rstEspecifUnificaVersion!Version) + Chr$(13) + _
                                         "Fecha de vigencia del : " + rstEspecifUnificaVersion!FechaInicio + " al " + rstEspecifUnificaVersion!FechaFinal
                                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                            
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
                                    
                                    m$ = "ATENCION : La partida esta asociada a una version de especificaciones que no es la actual" + Chr$(13) + _
                                         "Version " + Str$(rstEspecifUnificaVersion!Version) + Chr$(13) + _
                                         "Fecha de vigencia del : " + rstEspecifUnificaVersion!FechaInicio + " al " + rstEspecifUnificaVersion!FechaFinal
                                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                    
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
                
                    ZSql = ""
                    ZSql = ZSql + "Select EspecifUnifica.Version, EspecifUnifica.Producto, EspecifUnifica.Ensayo1, EspecifUnifica.Ensayo2, EspecifUnifica.Ensayo3, EspecifUnifica.Ensayo4, EspecifUnifica.Ensayo5, EspecifUnifica.Ensayo6, EspecifUnifica.Ensayo7, EspecifUnifica.Ensayo8, EspecifUnifica.Ensayo9, EspecifUnifica.Ensayo10, "
                    ZSql = ZSql + "EspecifUnifica.Valor1, EspecifUnifica.Valor2, EspecifUnifica.Valor3, EspecifUnifica.Valor4, EspecifUnifica.Valor5, EspecifUnifica.Valor6, EspecifUnifica.Valor7, EspecifUnifica.Valor8, EspecifUnifica.Valor9, EspecifUnifica.Valor10, "
                    ZSql = ZSql + "EspecifUnifica.Valor11, EspecifUnifica.Valor22, EspecifUnifica.Valor33, EspecifUnifica.Valor44, EspecifUnifica.Valor55, EspecifUnifica.Valor66, EspecifUnifica.Valor77, EspecifUnifica.Valor88, EspecifUnifica.Valor99, EspecifUnifica.Valor1010, "
                    ZSql = ZSql + "EspecifUnifica.Valor1Ing, EspecifUnifica.Valor2Ing, EspecifUnifica.Valor3Ing, EspecifUnifica.Valor4Ing, EspecifUnifica.Valor5Ing, EspecifUnifica.Valor6Ing, EspecifUnifica.Valor7Ing, EspecifUnifica.Valor8Ing, EspecifUnifica.Valor9Ing, EspecifUnifica.Valor10Ing, "
                    ZSql = ZSql + "EspecifUnifica.Valor11Ing, EspecifUnifica.Valor22Ing, EspecifUnifica.Valor33Ing, EspecifUnifica.Valor44Ing, EspecifUnifica.Valor55Ing, EspecifUnifica.Valor66Ing, EspecifUnifica.Valor77Ing, EspecifUnifica.Valor88Ing, EspecifUnifica.Valor99Ing, EspecifUnifica.Valor1010Ing, "
                    ZSql = ZSql + "EspecifUnifica.Desde1, EspecifUnifica.Desde2, EspecifUnifica.Desde3, EspecifUnifica.Desde4, EspecifUnifica.Desde5, EspecifUnifica.Desde6, EspecifUnifica.Desde7, EspecifUnifica.Desde8, EspecifUnifica.Desde9, EspecifUnifica.Desde10, "
                    ZSql = ZSql + "EspecifUnifica.Hasta1, EspecifUnifica.Hasta2, EspecifUnifica.Hasta3, EspecifUnifica.Hasta4, EspecifUnifica.Hasta5, EspecifUnifica.Hasta6, EspecifUnifica.Hasta7, EspecifUnifica.Hasta8, EspecifUnifica.Hasta9, EspecifUnifica.Hasta10 "
                    ZSql = ZSql + " FROM EspecifUnifica"
                    ZSql = ZSql + " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                    spEspecifUnifica = ZSql
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        If Idioma.ListIndex = 0 Then
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
                        WIdioma = IIf(IsNull(rstClientes!Idioma), "0", rstClientes!Idioma)
                        ZEtiI = Trim(IIf(IsNull(rstClientes!EtiI), "0", rstClientes!EtiI))
                        ZEtiII = Trim(IIf(IsNull(rstClientes!EtiII), "0", rstClientes!EtiII))
                        ZRazon = Left$(rstCliente!Razon, 50)
                        ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                        rstCliente.Close
                    End If
                    
                    ZOrdenCpa = ""
                    If Val(Pedido.Text) <> 0 Then
                        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPedido.RecordCount > 0 Then
                            ZZOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                            rstPedido.Close
                        End If
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
                                
                            ZOrden = ZOrdenCpa
                            ZClave1 = ZLote
                            Call Ceros(ZClave1, 6)
                            ZClave2 = Str$(LugarMetodo)
                            Call Ceros(ZClave2, 2)
                            ZClave = ZClave1 + ZClave2
                            ZMetodo = ZEnsayo(CiclaMetodo)
                            
                            If Val(ZStd(CiclaMetodo, 3)) <> 0 Or Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                If Idioma.ListIndex = 1 Then
                                    ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                End If
                                ZValorNormalII = ""
                                    Else
                                If Idioma.ListIndex = 0 Then
                                    ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                        Else
                                    ZValorNormalI = Left$(ZStd(CiclaMetodo, 5), 50)
                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 6), 50)
                                End If
                            End If
                            ZValorPartidaI = Left$(ZValor(CiclaMetodo), 50)
                            If Idioma.ListIndex = 1 Then
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
                                        ZExamenII = Mid(ZExamen, Cicla + 1, 25)
                                        ZExamen = Mid(ZExamen, 1, Cicla)
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
                        If Idioma.ListIndex = 0 Then
                            Listado.ReportFileName = "CertificadoNuevo.rpt"
                                Else
                            Listado.ReportFileName = "CertificadoNuevoIngles.rpt"
                        End If
                            Else
                        If Cliente.Text = "Z00007" Then
                            Listado.ReportFileName = "CertificadoPelliZ00007.rpt"
                                Else
                            Listado.ReportFileName = "CertificadoPelli.rpt"
                        End If
                    End If
                                
                    DbConnect = db.Connect
                    DSQ = getDatabase(DbConnect)
    
                    Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                    + "From " _
                                    + DSQ + ".dbo.Certificado Certificado " _
                                    + "Where " _
                                    + "Certificado.Partida >= 0 AND " _
                                    + "Certificado.Partida <= 999999"
                                    
                    If Tipo.ListIndex = 3 Then
                        Listado.Destination = 1
                            Else
                        Listado.Destination = 0
                    End If
    
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

End Sub

Private Sub Imprime_EtiquetaVerde()

    Rem On Error GoTo WError
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Imprsion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
                If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        WVida = 0
        Wvencimiento = ""
        Descripcion.Text = ""
                
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Descripcion.Text = Trim(rstTerminado!Descripcion)
            WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
            rstTerminado.Close
        End If
        
        If WVida > 0 Then
        
            For ZCiclo = 1 To 7
            
                Select Case ZCiclo
                    Case 1
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 3
                        Wempresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 4
                        Wempresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 5
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 6
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            
                spHoja = "ListaHoja " + "'" + Lote.Text + "'"
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WMes = Val(Mid$(rstHoja!Fecha, 4, 2))
                    WAno = Val(Right$(rstHoja!Fecha, 4))
                    For Ciclo = 1 To WVida
                        WMes = WMes + 1
                        If WMes > 12 Then
                            WAno = WAno + 1
                            WMes = 1
                        End If
                    Next Ciclo
                    If WVida <> 0 Then
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        Wvencimiento = "01/" + XMes + "/" + XAno
                    End If
                    rstHoja.Close
                End If
                
            Next ZCiclo
            
            Call Conecta_Empresa
    
        End If
        
        ZCantidad = Int(Val(Etiquetas.Text) / 2)
        If ZCantidad * 2 <> Val(Etiquetas.Text) Then
            ZCantidad = ZCantidad + 1
        End If
        
        Rem Descripcion.Text = "TENDIA AMARILLO"
        Rem Descripcion.Text = "TENDIA AMARILLO SBRSX 102"
        Rem Descripcion.Text = "TENDIA AMARILLO PATITO CON AUTOLIMPIANTE"
        
        With rstEtiqueta
            For Da = 1 To ZCantidad
                .Index = "Codigo"
                .AddNew
        
                ZDa = Int((Da - 1) / 2)
         
                !Codigo = Da
                !Clase = "Codigo : " + Terminado.Text
                !Razon = Descripcion.Text
                !DirEntrega = "Partida Nro. : " + Lote.Text
                !Conservacion = "F.Vto. : " + Wvencimiento
                !Nombre = "Neto : " + Cantidad.Text + " Kg."
                !Neto = ZDa
                
                .Update
            Next Da
        End With

        Listado.WindowTitle = "Emision de Etiquetas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height

        If Len(Descripcion.Text) > 15 Then
            If Len(Descripcion.Text) > 25 Then
                Listado.ReportFileName = "WEtiVerdePtChicaii.rpt"
                    Else
                Listado.ReportFileName = "WEtiVerdePtChica.rpt"
            End If
                Else
            Listado.ReportFileName = "WEtiVerdePt.rpt"
        End If
        Listado.DataFiles(0) = Wempresa + "Auxi.mdb"

        Listado.Destination = 1
        Listado.PrinterCopies = 1
        Listado.Action = 1
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub





Private Sub Calcula_Mono()

    
    ZZZZVencimiento = ""
    XEmpresa = Wempresa
    
    Select Case Val(Wempresa)
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
    
        spHoja = "ListaHoja " + "'" + Lote.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            rstHoja.Close
            
            ZZZZRenglon = 0
            ZZZZCantidad = 0
            ZZZZCantidadLote = 0
            ZZZZTipo = ""
            ZZZZLote = ""
            ZZZZTerminado = ""
            ZZZZArticulo = ""
            
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZZZRenglon = ZZZZRenglon + 1
                            ZZZZCantidad = rstHoja!Cantidad
                            ZZZZCantidadLote = rstHoja!Canti1
                            ZZZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                            ZZZZTipo = rstHoja!Tipo
                            ZZZZTerminado = rstHoja!Terminado
                            ZZZZArticulo = rstHoja!Articulo
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstHoja.Close
            End If
            
        End If
        
    Next ZCiclo
    
    Call Conecta_Empresa

    Rem If ZZZZRenglon = 1 And ZZZZCantidad = ZZZZCantidadLote And ZZZZTipo = "M" Then
    Rem
    Rem      ZZZZVto = ""
    Rem      ZZZZLaudo = ZZZZLote
    Rem      ZZZZFecha = ""
    Rem      ZZZZFechaVto = ""
    Rem
    Rem      For ZCiclo = 1 To ZHasta1
    Rem
    Rem          WEmpresa = CargaEmpresa(ZCiclo, 1)
    Rem          txtOdbc = CargaEmpresa(ZCiclo, 2)
    Rem          strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Rem          Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Rem
    Rem         ZSql = ""
    Rem          ZSql = ZSql + "Select *"
    Rem          ZSql = ZSql + " FROM Laudo"
    Rem          ZSql = ZSql + " Where Laudo = " + "'" + Str$(ZZZZLaudo) + "'"
    Rem          ZSql = ZSql + " and Articulo = " + "'" + ZZZZArticulo + "'"
    Rem          spLaudo = ZSql
    Rem          Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    Rem          If rstLaudo.RecordCount > 0 Then
    Rem              ZZZZFecha = rstLaudo!Fecha
    Rem              ZZZZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
    Rem              rstLaudo.Close
    Rem              Exit For
    Rem          End If
    Rem
    Rem      Next ZCiclo
    Rem
    Rem      Call Conecta_Empresa
    Rem
    Rem      ZZZZVto = ""
    Rem      ZZZZOrdFecha = Right$(ZZZZFecha, 4) + Mid$(ZZZZFecha, 4, 2) + Left$(ZZZZFecha, 2)
    Rem      If ZZZZFechaVto <> "" And ZZZZFechaVto <> "  /  /    " And ZZZZFechaVto <> "00/00/0000" Then
    Rem          Call Valida_fecha(ZZZZFechaVto, Auxi)
    Rem          If Auxi = "S" Then
    Rem              ZZZZVto = ZZZZFechaVto
    Rem          End If
    Rem     End If
    Rem
    Rem     If ZZZZVto = "" Then
    Rem
    Rem          ZZZZMeses = 0
    Rem          ZSql = ""
    Rem          ZSql = ZSql + "Select *"
    Rem          ZSql = ZSql + " FROM Articulo"
    Rem          ZSql = ZSql + " Where Codigo = " + "'" + ZZZZArticulo + "'"
    Rem          spArticulo = ZSql
    Rem          Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem          If rstArticulo.RecordCount > 0 Then
    Rem              ZZZZMeses = rstArticulo!Meses
    Rem              rstArticulo.Close
    Rem          End If
    Rem
    Rem          WMes = Val(Mid$(ZZZZFecha, 4, 2))
    Rem          WAno = Val(Right$(ZZZZFecha, 4))
    Rem          For ZCiclo = 1 To ZZZZMeses
    Rem              WMes = WMes + 1
    Rem              If WMes > 12 Then
    Rem                  WAno = WAno + 1
    Rem                  WMes = 1
    Rem              End If
    Rem          Next ZCiclo
    Rem
    Rem          XMes = Str$(WMes)
    Rem          XAno = Str$(WAno)
    Rem          Call Ceros(XMes, 2)
    Rem          Call Ceros(XAno, 4)
    Rem          If Val(Left$(ZZZZFecha, 2)) <= 30 Then
    Rem              If Val(XMes) = 2 And Val(Left$(ZZZZFecha, 2)) > 28 Then
    Rem                  ZZZZVto = "28/" + XMes + "/" + XAno
    Rem                      Else
    Rem                  ZZZZVto = Left$(ZZZZFecha, 3) + XMes + "/" + XAno
    Rem              End If
    Rem                  Else
    Rem              If Val(XMes) = 2 Then
     Rem                 ZZZZVto = "28/" + XMes + "/" + XAno
    Rem                      Else
    Rem                  ZZZZVto = "30/" + XMes + "/" + XAno
    Rem              End If
    Rem          End If
    Rem
    Rem     End If
    Rem
    Rem     ZZZZVencimiento = ZZZZVto
    Rem
    Rem End If


End Sub



Private Sub Calcula_Mono_Otro()

    
    ZZZZLoteOriginal = ""
    ZZZZVencimiento = ""
    XEmpresa = Wempresa
    
    Select Case Val(Wempresa)
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
    
        spHoja = "ListaHoja " + "'" + Lote.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            rstHoja.Close
            
            ZZZZRenglon = 0
            ZZZZCantidad = 0
            ZZZZCantidadLote = 0
            ZZZZTipo = ""
            ZZZZLote = ""
            ZZZZTerminado = ""
            ZZZZArticulo = ""
            
            spHoja = "ListaHoja " + "'" + Lote.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZZZRenglon = ZZZZRenglon + 1
                            ZZZZCantidad = rstHoja!Cantidad
                            ZZZZCantidadLote = rstHoja!Canti1
                            ZZZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                            ZZZZTipo = rstHoja!Tipo
                            ZZZZTerminado = rstHoja!Terminado
                            ZZZZArticulo = rstHoja!Articulo
                            
                            If ZZZZLote = 0 Then
                                ZZZZLote = Val(LoteMP.Text)
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstHoja.Close
            End If
            
        End If
        
    Next ZCiclo
    
    Call Conecta_Empresa
    
        
    If ZZZZRenglon = 1 And ZZZZTipo = "M" Then
    
         ZZZZVto = ""
         ZZZZLaudo = ZZZZLote
         ZZZZFecha = ""
         ZZZZFechaVto = ""
         ZZZZFechaElaboracion = ""
         ZZZZLoteOriginal = ""
    
         For ZCiclo = 1 To ZHasta1
    
             Wempresa = CargaEmpresa(ZCiclo, 1)
             txtOdbc = CargaEmpresa(ZCiclo, 2)
             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
             ZSql = ""
             ZSql = ZSql + "Select *"
             ZSql = ZSql + " FROM Laudo"
             ZSql = ZSql + " Where Laudo = " + "'" + Str$(ZZZZLaudo) + "'"
             ZSql = ZSql + " and Articulo = " + "'" + ZZZZArticulo + "'"
             spLaudo = ZSql
             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
             If rstLaudo.RecordCount > 0 Then
                  ZZZZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                  ZZZZFechaElaboracion = IIf(IsNull(rstLaudo!FechaElaboracion), "", rstLaudo!FechaElaboracion)
                  ZZZZLoteOriginal = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                  rstLaudo.Close
                  Exit For
             End If
    
         Next ZCiclo
    
         Call Conecta_Empresa
    
         ZZZZVto = ""
         If ZZZZFechaVto <> "" And ZZZZFechaVto <> "  /  /    " And ZZZZFechaVto <> "00/00/0000" Then
             Call Valida_fecha(ZZZZFechaVto, Auxi)
             If Auxi = "S" Then
                 ZZZZVto = ZZZZFechaVto
             End If
        End If
    
         ZZZZElabora = ""
         If ZZZZFechaElaboracion <> "" And ZZZZFechaElaboracion <> "  /  /    " And ZZZZFechaElaboracion <> "00/00/0000" Then
             Call Valida_fecha(ZZZZFechaElaboracion, Auxi)
             If Auxi = "S" Then
                 ZZZZElabora = ZZZZFechaElaboracion
             End If
        End If
    
        ZZZZVencimiento = ZZZZVto
        ZZZZElaboracion = ZZZZElabora
    
    End If

End Sub


