VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgLaudo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Laudo de Liberacion"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.CommandButton Historial 
      Caption         =   "Historial"
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
      Left            =   3360
      MaskColor       =   &H00C00000&
      TabIndex        =   36
      Top             =   6240
      Width           =   975
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   35
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox PartiOri 
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
      MaxLength       =   20
      TabIndex        =   34
      Top             =   480
      Width           =   3255
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
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
      Left            =   2280
      MaskColor       =   &H00C00000&
      TabIndex        =   17
      Top             =   6240
      Width           =   975
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
      Height          =   1425
      Left            =   6840
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   7560
      TabIndex        =   15
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
   Begin VB.TextBox Laudo 
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
      TabIndex        =   13
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
      Height          =   500
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   11
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
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
      Left            =   1200
      MaskColor       =   &H00C00000&
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
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
      Height          =   500
      Left            =   2280
      MaskColor       =   &H00C00000&
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   11655
      Begin VB.TextBox WNuevo 
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
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   23
         Text            =   " "
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox WRechazo 
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
         Height          =   300
         Left            =   9240
         MaxLength       =   6
         TabIndex        =   22
         Text            =   " "
         Top             =   720
         Width           =   1095
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
         Height          =   300
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   21
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox WDevuelta 
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
         Height          =   300
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox WLiberada 
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
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   19
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox WOrden 
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
         Height          =   300
         Left            =   360
         MaxLength       =   6
         TabIndex        =   18
         Text            =   " "
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nueva OC"
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
         Left            =   10320
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Rech."
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
         Left            =   9240
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Lote"
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
         Left            =   8160
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Devuelta"
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
         Left            =   7080
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Liberada"
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
         Left            =   6000
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden"
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
         Left            =   360
         TabIndex        =   24
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
         Height          =   300
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
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
      Height          =   500
      Left            =   120
      MaskColor       =   &H00C00000&
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "laudo.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   11655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   0
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
      Height          =   1815
      ItemData        =   "laudo.frx":09E6
      Left            =   4680
      List            =   "laudo.frx":09ED
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6255
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
      Height          =   500
      Left            =   1200
      MaskColor       =   &H00C00000&
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Origen Mercaderia"
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
      Left            =   5760
      TabIndex        =   33
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Partida del Proveedor"
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
      TabIndex        =   32
      Top             =   480
      Width           =   2175
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
      Left            =   5760
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Laudo"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Prglaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 8 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private WInforme As String
Private Pasa As String
Private Cantidad As String
Private Orden As String
Private Articulo As String
Private Verifica(100, 2) As String
Private Entra As String
Private Auxiliar(100, 6) As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim WCosto1 As String
Dim WCosto3 As String
Dim WPrecio As Double
Dim XStock As Double
Dim XCosto As Double
Dim XCostoTotal As Double
Dim XStock1 As Double
Dim XCosto1 As Double
Dim XCostoTotal1 As Double
Dim XStock2 As Double
Dim XCosto2 As Double
Dim XCostoTotal2 As Double
Dim XCosto3 As Double
Dim WTipoOrden As Single
Dim WMoneda As Single
Dim WSaldo As Double

Dim ZParidad As Double
Dim ZParidadII As Double
Dim ZCoeParidad As Double
Dim ZFechaInforme As String

Private XLote(100, 7) As String

Private Sub Borra_Click()

    Lugar = (DBGrid1.FirstRow * 10) + DBGrid1.Row + 1
            
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
    
    DBGrid1.Col = 7
    DBGrid1.Text = ""
    
    Verifica(Lugar, 1) = ""
    Verifica(Lugar, 2) = ""
    
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLiberada.Text = ""
    WDevuelta.Text = ""
    WLote.Text = ""
    WRechazo.Text = ""
    WNuevo.Text = ""
        
    WLinea.Text = ""
    
    WOrden.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstFichaMat
        .Close
    End With
    Rem With rstOrden
    Rem     .Close
    Rem End With
    Rem With rstInforme
    Rem     .Close
    Rem End With
    Rem With rstLaudo
    Rem     .Close
    Rem End With
    Rem With rstProveedor
    Rem     .Close
    Rem End With
    
    Rem DbsVentas.Close
    Rem DbsAdminis.Close
    Rem DbsCotiza.Close
    
    Prglaudo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Orden de Compra"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem OPEN_FILE_LAUDO
    Rem OPEN_FILE_Informe
    Rem OPEN_FILE_Orden
    Rem OPEN_FILE_Proveedor
    OPEN_FILE_FichaMat
End Sub

Private Sub Historial_Click()

    Rem dada

    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
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
    
    WArticulo = WArticulo.Text
    WLote = Laudo.Text
    WFecha = Fecha.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WCodigo = Laudo.Text
    WCantidad = Val(WLiberada.Text)
    WComprobante = WOrden.Text
    WDescri = "Laudo"
    WPartiOri = ""
    
    nrolote = Laudo.Text
            
    spOrden = "ListaOrden" + "'" + WComprobante + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        WProveedor = rstOrden!Proveedor
        rstOrden.Close
    End If
    
    WObservaciones = ""
        
    spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WObservaciones = RstProveedor!Nombre
        RstProveedor.Close
    End If
        
    WDesArticulo = ""
        
    spArticulo = "ConsultaArticulo " + " '" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WDesArticulo = rstArticulo!Descripcion
        WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
        rstArticulo.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.Laudo = " + "'" + Laudo.Text + "'"
    ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArticulo + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
        rstLaudo.Close
    End If
    
    If WLiberadaAnt <> 0 Then
        WCantidad = WLiberadaAnt
            Else
        WCantidad = WLiberada
    End If
            
    With rstFichaMat
        .AddNew
        !Articulo = WArticulo
        !Fecha = WFecha
        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        !Tipo = 0
        !Numero = WCodigo
        !Inicial = 0
        !Entrada = WCantidad
        !Salida = 0
        !Descripcion = WDesArticulo
        !Observaciones = WObservaciones
        !Lista1 = WDescri
        !Lista2 = ""
        !Lote = WLote
        !Saldo = 0
        !Empresa = ""
        !PartiOri = WPartiOri
        .Update
    End With
    
    
            
            
            
            
            
            
            
            
            
            
            
            
            
            
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Tipo = "M" Then
                
                    XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    
                    For Da = 1 To 3
                        If Val(XLote(Da, 1)) = Val(nrolote) Then
                
                            WArticulo = rstHoja!Articulo
                            WCantidad = XLote(Da, 2)
                            WFecha = rstHoja!Fecha
                            WHoja = rstHoja!Hoja
                            WSaldo = 0
                
                            With rstFichaMat
                
                                .AddNew
                                !Articulo = WArticulo
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WHoja
                                !Inicial = 0
                                !Entrada = 0
                                !Salida = WCantidad
                                !Observaciones = ""
                                !Descripcion = WDesArticulo
                                !Lista1 = "Hoja"
                                !Lista2 = ""
                                !Lote = Val(nrolote)
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                            End With
                        End If
                    Next Da
                        
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstHoja.Close
    End If
    
    
    
    
    
    
    
    
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Tipo = "M" Then
                
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    If Val(WLote) = Val(nrolote) Then
                
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        WTipomov = Val(rstMovvar!Tipomov)
                        WObservaciones = rstMovvar!Observaciones
                        WSaldo = 0
                    
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            If WTipomov = 0 Or WTipomov = 1 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Tipo = "M" Then
            
                    WArticulo = rstMovguia!Articulo
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    Rem WObservaciones = rstMovvar!Observaciones
                        
                    If WMovi = "E" Then
                        WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                            Else
                        WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                        ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        WSaldo = 0
                    End If

                        
                    If WMovi = "S" Then
                        Select Case WDestino
                            Case 1
                                WObservaciones = "Envio a Surfactan"
                            Case 2
                                WObservaciones = "Envio a Pellital"
                            Case 3
                                WObservaciones = "Envio a Surfactan II"
                            Case 4
                                WObservaciones = "Envio a Pellital II"
                            Case 5
                                WObservaciones = "Envio a Surfactan III"
                            Case 6
                                WObservaciones = "Envio a Surfactan IV"
                            Case 7
                                WObservaciones = "Envio a Surfactan V"
                            Case 8
                                WObservaciones = "Envio a Pellital V"
                            Case 9
                                WObservaciones = "Envio a Pellital IV"
                            Case 10
                                WObservaciones = "Envio a Surfactan VI"
                            Case 11
                                WObservaciones = "Envio a Surfactan VII"
                            Case Else
                        End Select
                            
                                Else
                                
                        Select Case WTipomov
                            Case 1
                                WObservaciones = "Recepcion de Surfactan"
                            Case 2
                                WObservaciones = "Recepcion de Pellital"
                            Case 3
                                WObservaciones = "Recepcion de Surfactan II"
                            Case 4
                                WObservaciones = "Recepcion de Pellital II"
                            Case 5
                                WObservaciones = "Recepcion de Surfactan III"
                            Case 6
                                WObservaciones = "Recepcion de Surfactan IV"
                            Case 7
                                WObservaciones = "Recepcion de Surfactan V"
                            Case 8
                                WObservaciones = "Recepcion de Pellital V"
                            Case 9
                                WObservaciones = "Recepcion de Pellital IV"
                            Case 10
                                WObservaciones = "Recepcion de Surfactan VI"
                            Case 11
                                WObservaciones = "Recepcion de Surfactan VII"
                            Case Else
                        End Select
                            
                    End If
                    
                    If WLote = Val(nrolote) Then
                        With rstFichaMat
                
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            If !Numero > 900000 Then
                                !Lista1 = "Prestamo"
                                !Numero = !Numero - 900000
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                            
                        End With
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovguia.Close
    End If
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Tipo = "M" Then
                
                    WArticulo = rstMovlab!Articulo
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    Rem WSaldo = IIf(IsNull(rstMovlab!Saldo), "0", rstMovlab!Saldo)
                    
                    If Val(WLote) = Val(nrolote) Then
                        
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            !Lista1 = "Mov.Lab"
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstMovlab.Close
    End If
    
    
    Rem PROCESA LOS las devoluciones de mercaderia
    
    WAuxiliar = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
    
    XParam = "'" + WAuxiliar + "','" _
                 + WAuxiliar + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                    WCantidad = rstEntdev!Cantidad
                    WFecha = rstEntdev!Fecha
                    WCodigo = rstEntdev!Codigo
                    WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                    WPartiOri = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                    WSaldo = rstEntdev!Saldo
                
                    If Val(nrolote) = WLote Then
                        With rstFichaMat
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            !Entrada = WCantidad
                            !Salida = 0
                            !Observaciones = ""
                            !Lista1 = "Ent.Dev."
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !PartiOri = WPartiOri
                            .Update
                        End With
                    End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
    
    
    
    
    
    
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WObservaciones = !Observaciones
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                If !Lista1 = "Devol." Or !Lista1 = "Factura" Then
                    spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                End If
                !Descripcion = WDescripcion
                !Observaciones = WObservaciones
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.ReportFileName = "WLotemat.rpt"

    Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.Destination = 0
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1

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
            spOrden = "ListaOrdenTotal"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstOrden!Orden) + " " + rstOrden!Articulo
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstOrden!Clave
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstOrden.Close
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    WOrden.Text = DBGrid1.Text

    DBGrid1.Col = 1
    If Len(DBGrid1.Text) = 10 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 2
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 3
    If Val(DBGrid1.Text) <> 0 Then
        WLiberada.Text = DBGrid1.Text
            Else
        WLiberada.Text = ""
    End If
        
    DBGrid1.Col = 4
    If Val(DBGrid1.Text) <> 0 Then
        WDevuelta.Text = DBGrid1.Text
            Else
        WDevuelta.Text = ""
    End If
    
    DBGrid1.Col = 5
    WLote.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WRechazo.Text = DBGrid1.Text
    
    DBGrid1.Col = 7
    WNuevo.Text = DBGrid1.Text
    
    WOrden.SetFocus

End Sub

Private Sub Graba_Click()

    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha del laudo de liberacion es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Renglon = 0
    Erase Auxiliar
    
    spLaudo = "ListaLaudo " + "'" + Laudo.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)

    If rstLaudo.RecordCount > 0 Then
    With rstLaudo
        .MoveFirst
        Do
            If .EOF = False Then
                Renglon = Renglon + 1
                Auxiliar(Renglon, 1) = rstLaudo!Orden
                Auxiliar(Renglon, 2) = rstLaudo!Laudo
                Auxiliar(Renglon, 3) = rstLaudo!Articulo
                Auxiliar(Renglon, 4) = rstLaudo!Liberada
                Auxiliar(Renglon, 5) = rstLaudo!devuelta
                Auxiliar(Renglon, 6) = rstLaudo!Actualiza
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstLaudo.Close
    End If
    
    For Da = 1 To Renglon

        Orden = Auxiliar(Da, 1)
        Laudo = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Liberada = Auxiliar(Da, 4)
        devuelta = Auxiliar(Da, 5)
        Actualiza = Auxiliar(Da, 6)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            WCodigo = Articulo
            If Nueva = "S" Then
                WLaboratorio = Str$(rstArticulo!Laboratorio + Val(Liberada) + Val(devuelta))
                    Else
                WLaboratorio = Str$(rstArticulo!Laboratorio + Val(Liberada))
            End If
            WEntradas = Str$(rstArticulo!Entradas - Val(Liberada))
            WCosto1 = Str$(rstArticulo!Costo1)
            WCosto3 = Str$(IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3))
            WDate = Date$
            rstArticulo.Close
                
            XParam = "'" + WCodigo + "','" _
                    + WLaboratorio + "','" _
                    + WEntradas + "','" _
                    + WDate + "','" _
                    + WCosto1 + "','" _
                    + WCosto3 + "'"
            rstArticulo.Close
                                           
            spArticulo = "ModificaArticuloLaudo " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        XParam = "'" + Orden + "','" _
                     + Articulo + "'"
        spOrden = "ListaOrdenArticulo " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
        If rstOrden.RecordCount > 0 Then
        
            WClave = rstOrden!Clave
            WLiberada = Str$(rstOrden!Liberada - Val(Liberada))
            WDevuelta = Str$(rstOrden!devuelta - Val(devuelta))
            WDate = Date$
            WFechaEntrega = Fecha.Text
            rstOrden.Close
            
            XParam = "'" + WClave + "','" _
                    + WLiberada + "','" _
                    + WDevuelta + "','" _
                    + WFechaEntrega + "','" _
                    + WDate + "'"
                                                                  
            spOrden = "ModificaOrdenLaudo " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Da
                
    spLaudo = "BorrarLaudo " + "'" + Laudo.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenDynaset, dbSQLPassThrough)
    
    
    Rem da de alta el laudo
    
    
    
    
    Renglon = 0
    Erase Auxiliar
    ZFechaInforme = "  /  /    "
        
    DBGrid1.Refresh
    
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Orden = DBGrid1.Text
                    
            DBGrid1.Col = 1
            Articulo = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 3
            Liberada = DBGrid1.Text
                    
            DBGrid1.Col = 4
            devuelta = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Lote = DBGrid1.Text
                                
            DBGrid1.Col = 6
            Rechazo = DBGrid1.Text
                    
            DBGrid1.Col = 7
            Nuevo = DBGrid1
    
            If Articulo <> "" Then
            
                XParam = "'" + Orden + "','" _
                        + Articulo + "'"
                spInforme = "ListaInformeOrdenArticulo " + XParam
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    Informe = rstInforme!Informe
                    ZFechaInforme = rstInforme!Fecha
                    rstInforme.Close
                End If
                        
                Call Busca_Informe(Orden, WInforme, Articulo, Pasa, Cantidad)
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Laudo.Text)
                Call Ceros(Auxi1, 6)
                
                WClave = Auxi1 + Auxi
                WLaudo = Laudo.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WOrden = Orden
                WArticulo = Articulo
                WLiberada = Val(Liberada)
                WDevuelta = Val(devuelta)
                WLote = Val(Lote)
                WRechazo = Val(Rechazo)
                WActualiza = Nuevo
                WMarca = ""
                WInforme = WInforme
                WDate = Date$
                WSaldo = WLiberada
        
                XParam = "'" + WClave + "','" _
                         + WLaudo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WArticulo + "','" _
                         + WLiberada + "','" _
                         + WDevuelta + "','" _
                         + WOrden + "','" _
                         + WMarca + "','" _
                         + WLote + "','" _
                         + WRechazo + "','" _
                         + WInforme + "','" _
                         + WActualiza + "','" _
                         + WDate + "','" _
                         + WSaldo + "'"
                         
                spLaudo = "AltaLaudo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                
                Auxiliar(Renglon, 1) = WOrden
                Auxiliar(Renglon, 2) = WLaudo
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WLiberada
                Auxiliar(Renglon, 5) = WDevuelta
                Auxiliar(Renglon, 6) = WActualiza
                
            End If
                                        
        Next iRow
            
    Next a
    
    ZParidad = 0
    ZParidadII = 0
    ZCoeParidad = 1
    
    XEmpresa = WEmpresa
    
    Select Case Val(XEmpresa)
        Case 2, 4, 8, 9
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    spCambios = "ConsultaCambio  " + "'" + ZFechaInforme + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        ZParidad = rstCambios!Cambio
        ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
        If ZParidadII <> 0 And ZParidad <> 0 Then
            ZCoeParidad = ZParidadII / ZParidad
                Else
            ZCoeParidad = 1
        End If
        rstCambios.Close
    End If
    
    Call Conecta_Empresa
        
    For Da = 1 To Renglon

        Orden = Auxiliar(Da, 1)
        Laudo = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Liberada = Auxiliar(Da, 4)
        devuelta = Auxiliar(Da, 5)
        Actualiza = Auxiliar(Da, 6)
        
        WPrecio = 0
        
        XParam = "'" + Orden + "','" _
                     + Articulo + "'"
        spOrden = "ListaOrdenArticulo " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WMoneda = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            WClave = rstOrden!Clave
            WLiberada = rstOrden!Liberada + Val(Liberada)
            WDevuelta = rstOrden!devuelta + Val(devuelta)
            WFechaEntrega = Fecha.Text
            WPrecio = rstOrden!Precio
            WDate = Date$
            rstOrden.Close
                
            XParam = "'" + WClave + "','" _
                    + WLiberada + "','" _
                    + WDevuelta + "','" _
                    + WFechaEntrega + "','" _
                    + WDate + "'"
                                           
            spOrden = "ModificaOrdenLaudo " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            WCodigo = Articulo
            WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Liberada) - Val(devuelta))
            
            If WTipoOrden <> 1 Then
            
                XStock1 = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                End If
                XCostoTotal1 = XStock1 * XCosto1
            
                XStock2 = Val(Liberada)
                If WMoneda = 2 Then
                    XCosto2 = WPrecio * ZCoeParidad
                        Else
                    XCosto2 = WPrecio
                End If
                XCostoTotal2 = XStock2 * XCosto2
            
                XCosto = 0
                XStock = XStock1 + XStock2
                XCostoTotal = XCostoTotal1 + XCostoTotal2
                If XStock <> 0 Then
                    XCosto = XCostoTotal / XStock
                End If
            
                Call Redondeo(XCosto)
                    
                If WMoneda = 2 Then
                    WCosto1 = Str$(WPrecio * ZCoeParidad)
                        Else
                    WCosto1 = Str$(WPrecio)
                End If
                WCosto3 = Str$(XCosto)
                
                    Else
                    
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo1), "0", rstArticulo!Costo1)
                    XCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                    XCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                End If
                
            End If
                
            WEntradas = Str$(rstArticulo!Entradas + Val(Liberada))
            WDate = Date$
            rstArticulo.Close
            
            If WMoneda = 0 Or WMoneda = 2 Then
                XParam = "'" + WProducto + "','" _
                    + WLaboratorio + "','" _
                    + WEntradas + "','" _
                    + WDate + "','" _
                    + WCosto1 + "','" _
                    + WCosto3 + "'"
                spArticulo = "ModificaArticuloLaudoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WProducto + "','" _
                    + WLaboratorio + "','" _
                    + WEntradas + "','" _
                    + WDate + "','" _
                    + WCosto1 + "','" _
                    + WCosto3 + "'"
                spArticulo = "ModificaArticuloLaudoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
            Rem Actualiza costos de todas las empresas
            
            XEmpresa = WEmpresa
            
            XParam = "'" + WCodigo + "','" _
                         + WCosto1 + "'"
            
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            If WMoneda = 0 Or WMoneda = 2 Then
                spArticulo = "ModificaArticuloCostoDolares " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "ModificaArticuloCostoPesos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
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
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
        End If
        
    Next Da
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Laudo.SetFocus
    
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLiberada.Text = ""
    WDevuelta.Text = ""
    WLote.Text = ""
    WRechazo.Text = ""
    WNuevo.Text = ""
    
    WOrden.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Graba.Enabled = True
    
    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WLiberada.Text = ""
    WDevuelta.Text = ""
    WLote.Text = ""
    WRechazo.Text = ""
    WNuevo.Text = ""

    Laudo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Origen.Text = ""
    PartiOri.Text = ""

    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 7
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Rem With rstLaudo
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Laudo.Text = !Laudo + 1
    Rem             Else
    Rem         Laudo.Text = ""
    Rem     End If
    Rem End With

    spLaudo = "ListaLaudoNumero"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveLast
            Laudo.Text = rstLaudo!Laudo + 1
        End With
        rstLaudo.Close
            Else
        Laudo.Text = "1"
    End If

    Erase Verifica
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    Laudo.SetFocus

End Sub

Private Sub WOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WArticulo.SetFocus
            rstOrden.Close
                Else
            WOrden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        Pasa = "N"
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        If WArticulo.Text = rstOrden!Articulo Then
                            Pasa = "S"
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        End If
        
        If Pasa = "S" Then
            Call Busca_Informe(WOrden.Text, WInforme, WArticulo.Text, Pasa, Cantidad)
            If Pasa = "S" Then
                spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                        WDescripcion.Caption = rstArticulo!Descripcion
                        WLiberada.SetFocus
                End If
                    Else
                WArticulo.SetFocus
            End If
                        Else
            WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WLiberada_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Informe(WOrden.Text, WInforme, WArticulo.Text, Pasa, Cantidad)
        Canti = Val(WLiberada.Text) + Val(WDevuelta.Text)
        If Canti > Val(Cantidad) Then
            m$ = "La cantidad a laudar supera la informada en el informe de recepcion"
            a% = MsgBox(m$, 0, "Ingreso de laudo de liberacion")
            WLiberada.Text = ""
            WDevuelta.Text = ""
            WLiberada.SetFocus
                Else
            WLiberada.Text = Pusing("###,###.##", WLiberada.Text)
            WDevuelta.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WDevuelta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Informe(WOrden.Text, WInforme, WArticulo.Text, Pasa, Cantidad)
        Canti = Val(WLiberada.Text) + Val(WDevuelta.Text)
        If Canti > Val(Cantidad) Then
            m$ = "La cantidad a laudar supera la informada en el informe de recepcion"
            a% = MsgBox(m$, 0, "Ingreso de laudo de liberacion")
            WLiberada.Text = ""
            WDevuelta.Text = ""
            WLiberada.SetFocus
                Else
            WLote.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WDevuelta.Text) <> 0 Then
            WRechazo.SetFocus
                    Else
            WRechazo.Text = ""
            WNuevo.Text = "N"
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WRechazo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WNuevo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WNuevo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WNuevo.Text = "S" Or WNuevo.Text = "N" Then
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WNuevo.SetFocus
        End If
    End If
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spOrden = "ConsultaOrden " + "'" + Claveven$ + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WOrden.Text = rstOrden!Orden
                WArticulo.Text = rstOrden!Articulo
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
            End If
                    
            Call Alta_Vector
                    
            If Entra = "S" Then
                    
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                Rem Call DBGrid1.SetFocus
                        
                WLiberada.SetFocus
                        
                    Else
                            
                Call Ingresa_Click
                        
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 7, 0 To 40)

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
For i = 0 To 7
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Orden"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3400
             DBGrid1.Columns(newcnt).AllowSizing = False
             Rem DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Liberada"
             DBGrid1.Columns(newcnt).Width = 1050
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Devuelta"
             DBGrid1.Columns(newcnt).Width = 1050
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1050
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Rechazo"
             DBGrid1.Columns(newcnt).Width = 1050
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Nueva O/C"
             DBGrid1.Columns(newcnt).Width = 1050
             DBGrid1.Columns(newcnt).AllowSizing = False
             Rem DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Laudo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    PartiOri.Text = ""
    Origen.Text = ""
    Graba.Enabled = True

    Rem With rstLaudo
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Laudo.Text = !Laudo + 1
    Rem             Else
    Rem        Laudo.Text = ""
    Rem     End If
    Rem End With
    
    spLaudo = "ListaLaudoNumero"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveLast
            Laudo.Text = rstLaudo!Laudo + 1
        End With
        rstLaudo.Close
            Else
        Laudo.Text = "1"
    End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            Prglaudo.Caption = "Ingreso de Laudo de Liberacion :  " + !Nombre
        End If
    End With
    
    
    Laudo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    WGraba = "S"

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 7
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    
    spLaudo = "ListaLaudo " + "'" + Laudo.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstLaudo!Orden
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstLaudo!Articulo
                    Auxi1 = rstLaudo!Articulo
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Val(rstLaudo!Liberada))
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Val(rstLaudo!devuelta))
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstLaudo!Lote
                
                    DBGrid1.Col = 6
                    DBGrid1.Text = rstLaudo!Rechazo
                
                    DBGrid1.Col = 7
                    DBGrid1.Text = rstLaudo!Actualiza
                    
                    WCanti = rstLaudo!Liberada
                    WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    
                    If WCanti <> WSaldo Then
                        WGraba = "N"
                    End If
                    
                    Auxiliar(Renglon, 1) = rstLaudo!Articulo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        spArticulo = "ConsultaArticulo " + "'" + Auxiliar(Renglon, 1) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 2
            DBGrid1.Text = rstArticulo!Descripcion
            WOrden.SetFocus
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
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    If WGraba = "N" Then
        Graba.Enabled = False
    End If
    
    WOrden.SetFocus
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub Alta_Vector()
    Entra = "S"
    
    If Val(WLinea.Text) = 0 Then
            For Da = 1 To 100
                If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text Then
                    Entra = "N"
                    Exit For
                End If
            Next Da
                Else
            Lugar = (DBGrid1.FirstRow * 10) + DBGrid1.Row + 1
            For Da = 1 To 100
                If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text And Da <> Lugar Then
                    Entra = "N"
                    Exit For
                End If
            Next Da
    End If
    
    If Entra = "N" Then
            m$ = "El articulo ya se encuentra dado de alta en el laudo de liberacion"
            a% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
    End If
                
    If Entra = "S" Then

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WOrden.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WLiberada.Text)
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WDevuelta.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WRechazo.Text
            
            DBGrid1.Col = 7
            DBGrid1.Text = WNuevo.Text
                
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
            Verifica(Renglon, 1) = WArticulo.Text
            Verifica(Renglon, 2) = WOrden.Text
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WOrden.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WDescripcion.Caption
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WLiberada.Text)
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WDevuelta.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WRechazo.Text
            
            DBGrid1.Col = 7
            DBGrid1.Text = WNuevo.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
            Lugar = (DBGrid1.FirstRow * 10) + DBGrid1.Row + 1
            Verifica(Lugar, 1) = WArticulo.Text
            Verifica(Lugar, 2) = WOrden.Text
            
    End If
    
    End If

End Sub

Private Sub Laudo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spLaudo = "ListaLaudo " + "'" + Laudo.Text + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            Fecha.Text = rstLaudo!Fecha
            WPartiorianterior = IIf(IsNull(rstLaudo!PartiOriAnterior), "", rstLaudo!PartiOriAnterior)
            If WPartiorianterior <> "" Then
                PartiOri.Text = WPartiorianterior
                    Else
                PartiOri.Text = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
            End If
            Origen.Text = IIf(IsNull(rstLaudo!Origen), "", rstLaudo!Origen)
            Call Proceso_Click
                Else
            WLaudo = Laudo.Text
            Call Limpia_Click
            Laudo.Text = WLaudo
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            PartiOri.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub PartiOri_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Origen.SetFocus
    End If
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WOrden.SetFocus
    End If
End Sub

Sub Busca_Informe(Orden As String, Informe As String, Articulo As String, Pasa As String, WCantidad As String)

    Informe = ""
    Pasa = "N"
    
    XParam = "'" + Orden + "','" _
                 + Articulo + "'"
    spInforme = "ListaInformeOrdenArticulo " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstInforme.RecordCount > 0 Then
        Informe = rstInforme!Informe
        WCantidad = rstInforme!Cantidad
        Pasa = "S"
        rstInforme.Close
    End If
    
End Sub




