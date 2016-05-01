VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsFicMatAnt 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Ficha de Stock Materia Prima Historico"
   ClientHeight    =   7695
   ClientLeft      =   210
   ClientTop       =   840
   ClientWidth     =   11490
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   11490
   Begin VB.Frame PantaHistoria 
      Caption         =   "Historial"
      Height          =   3135
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox HistoriaOrden 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox HistoriaInforme 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox HistoriaRemito 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox HistoriaFactura 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton PantaHistoriaCierra 
         Caption         =   "Cierra"
         Height          =   495
         Left            =   3360
         TabIndex        =   19
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox HistoriaCarpeta 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin MSMask.MaskEdBox HistoriaFechaOrden 
         Height          =   285
         Left            =   5280
         TabIndex        =   24
         Top             =   360
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
      Begin MSMask.MaskEdBox HistoriaFechaInforme 
         Height          =   285
         Left            =   5280
         TabIndex        =   25
         Top             =   720
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
      Begin MSMask.MaskEdBox HistoriaFechaFactura 
         Height          =   285
         Left            =   5280
         TabIndex        =   26
         Top             =   1440
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
      Begin VB.Label Label7 
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
         Left            =   840
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Recepcion"
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
         Left            =   840
         TabIndex        =   30
         Top             =   720
         Width           =   2055
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
         Left            =   840
         TabIndex        =   29
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Factura"
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
         Left            =   840
         TabIndex        =   28
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label11 
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
         Left            =   840
         TabIndex        =   27
         Top             =   1800
         Width           =   1695
      End
   End
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
      Height          =   300
      Left            =   3480
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   5
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   6
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3840
      Width           =   375
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WLotemat.rpt"
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   1680
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
      Height          =   1425
      ItemData        =   "consficmatAnt.frx":0000
      Left            =   120
      List            =   "consficmatAnt.frx":0007
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   3480
      TabIndex        =   2
      Top             =   120
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5295
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9340
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label DesArticulo 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Articulo"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsFicMatAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private Vector(10000, 10) As String
Private XLote(100, 7) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim XParam As String
Private WOrden As String
Private WXInicial As Double
Private WXSalidas As Double
Private WXEntradas As Double
Private WXStock As Double
Private WCanti As Double
Private WSaldo As Double

Private Sub cmdClose_Click()
    Articulo.SetFocus
    PrgConsFicMatAnt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    
    Pantalla.Clear
    WIndice.Clear

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstArticulo!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaMat
End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Articulo.Text = rstArticulo!Codigo
        DesArticulo.Caption = rstArticulo!Descripcion
        rstArticulo.Close
        Call Proceso_Click
        WVector1.SetFocus
            Else
        Articulo.Text = WArticulo
    End If
    Articulo.SetFocus
    
End Sub


Private Sub Form_Load()

    Call Limpia_Vector

    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgConsFicMatAnt.Caption = "Consulta de Ficha de Stock de Materias Primas Historico :  " + !Nombre
        End If
    End With
    
    Rem Articulo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Articulo.Text = UCase(Articulo.Text)

    WXInicial = 0
    WXEntradas = 0
    WXSalidas = 0
    WXStock = 0
    
    Call Limpia_Vector
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1
    
    Renglon = 0
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
                
        WArticulo = rstArticulo!Codigo
        WInicial = rstArticulo!Inicial
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                                        
        Renglon = Renglon + 1
                                
        WVector1.Row = Renglon
                   
        WVector1.Col = 1
        WVector1.Text = "14/12/2000"
                        
        WVector1.Col = 2
        WVector1.Text = ""
                        
        WVector1.Col = 3
        WVector1.Text = ""
                        
        WVector1.Col = 4
        WVector1.Text = "Saldo Inicial"
                        
        WVector1.Col = 5
        WVector1.Text = Pusing("###,###.##", Str$(rstArticulo!Inicial))
                
        WVector1.Col = 6
        WVector1.Text = ""
                
        WXInicial = rstArticulo!Inicial
        
        rstArticulo.Close
                
    End If
                
    Rem PROCESA LOS LAUDOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                If XFec < WOrdFechaCierre Then
                
                WMarcaAnt = IIf(IsNull(rstLaudo!MarcaAnt), "0", rstLaudo!MarcaAnt)
                WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                
                If rstLaudo!MarcaAnt = "X" And rstLaudo!Liberadaant = 0 Then
                
                        Else
                    
                    If rstLaudo!Articulo = Articulo.Text Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        WFecha = rstLaudo!Fecha
                        WLaudo = rstLaudo!Laudo
                        WOrden = rstLaudo!Orden
                        WDevuelta = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        WRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        WSaldo = IIf(IsNull(rstLaudo!Saldoant), "0", rstLaudo!Saldoant)
                        WLiberada = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        Call Redondeo(WSaldo)
                        
                        If WLiberada <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Laudo"
                            Vector(Lugar, 3) = WLaudo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(WLiberada))
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WOrden
                            Vector(Lugar, 9) = WLaudo
                            Vector(Lugar, 10) = WSaldo
                    
                            WXEntradas = WXEntradas + WLiberada
                            
                        End If
                        
                        If WDevuelta <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Rechazo"
                            Vector(Lugar, 3) = WRechazo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = "(" + Pusing("###,###.##", Str$(rstLaudo!devueltaant)) + ")"
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WOrden
                            Vector(Lugar, 9) = WRechazo
                            Vector(Lugar, 10) = "0"
                            
                        End If
                
                    End If
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        WOrden = Vector(Ciclo, 8)
        
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WDEsProveedor = rstProveedor!Nombre
            rstProveedor.Close
        End If
    
        Vector(Ciclo, 4) = WDEsProveedor
        
    Next Ciclo
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
        
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
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
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If XFec < WOrdFechaCierre Then
                
                If rstHoja!MarcaAnt = "X" Or XFec < "20001218" Then
                
                        Else
                        
                    fr = rstHoja!Clave
                        
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = Articulo.Text Then
                    
                
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) = "" Then
                                XLote(Da, 2) = "0"
                            End If
                        
                            WCanti = XLote(Da, 2)
                            If WCanti <> 0 Then
                
                                WArticulo = rstHoja!Articulo
                                WCanti = XLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = XLote(Da, 1)
                        
                                Lugar = Lugar + 1
                        
                                Vector(Lugar, 1) = !Fecha
                                Vector(Lugar, 2) = "Hoja"
                                Vector(Lugar, 3) = WHoja
                                Vector(Lugar, 4) = ""
                                Vector(Lugar, 5) = ""
                                Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCanti * 1))
                                Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                Vector(Lugar, 9) = WLote
                                Vector(Lugar, 10) = ""
                        
                                WXSalidas = WXSalidas + WCanti
                                
                            End If
                        Next Da

                    End If
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Articulo > Articulo.Text Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
        
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
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
                
                If rstMovvar!FechaOrd < "20001218" Or rstMovvar!FechaOrd > WOrdFechaCierre Then
                
                        Else
                        
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = Articulo.Text Then
                    
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovvar!Fecha
                        If rstMovvar!Tipomov = 0 Or rstMovvar!Tipomov = 1 Then
                            Vector(Lugar, 2) = "Mov.Var"
                                Else
                            Vector(Lugar, 2) = "Guia In"
                        End If
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovvar!Observaciones
                        If rstMovvar!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovvar!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovvar!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(rstMovvar!Cantidad))
                            WXSalidas = WXSalidas + rstMovvar!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        Vector(Lugar, 10) = ""
                        
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
        
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
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
                
                If rstMovguia!FechaOrd < WOrdFechaCierre Then

                If rstMovguia!MarcaAnt = "X" And rstMovguia!Cantidadant = 0 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = Articulo.Text Then
                    
                    
                        Dada = rstMovguia!Clave
                    
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidadant
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovguia!Fecha
                        If Val(WCodigo) > 900000 Then
                            Vector(Lugar, 2) = "Prestamo"
                            Vector(Lugar, 3) = WCodigo - 900000
                                Else
                            Vector(Lugar, 2) = "Guia In"
                            Vector(Lugar, 3) = WCodigo
                        End If
                        Rem Vector(Lugar, 4) = rstMovguia!Observaciones
                                
                        If rstMovguia!Movi = "E" Then
                            Select Case WTipomov
                                Case 1
                                    Vector(Lugar, 4) = "Recepcion de Surfactan"
                                Case 2
                                    Vector(Lugar, 4) = "Recepcion de Pellital"
                                Case 3
                                    Vector(Lugar, 4) = "Recepcion de Surfactan II"
                                Case 4
                                    Vector(Lugar, 4) = "Recepcion de Pellital II"
                                Case 5
                                    Vector(Lugar, 4) = "Recepcion de Surfactan III"
                                Case 6
                                    Vector(Lugar, 4) = "Recepcion de Surfactan IV"
                                Case 7
                                    Vector(Lugar, 4) = "Recepcion de Surfactan V"
                                Case 8
                                    Vector(Lugar, 4) = "Recepcion de Pellital V"
                                Case 9
                                    Vector(Lugar, 4) = "Recepcion de Pellital IV"
                                Case 10
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovguia!Cantidadant))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovguia!Cantidadant
                            Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            Vector(Lugar, 10) = IIf(IsNull(rstMovguia!Saldoant), "0", rstMovguia!Saldoant)
                                Else
                            Select Case WDestino
                                Case 1
                                    Vector(Lugar, 4) = "Envio a Surfactan"
                                Case 2
                                    Vector(Lugar, 4) = "Envio a Pellital"
                                Case 3
                                    Vector(Lugar, 4) = "Envio a Surfactan II"
                                Case 4
                                    Vector(Lugar, 4) = "Envio a Pellital II"
                                Case 5
                                    Vector(Lugar, 4) = "Envio a Surfactan III"
                                Case 6
                                    Vector(Lugar, 4) = "Envio a Surfactan IV"
                                Case 7
                                    Vector(Lugar, 4) = "Envio a Surfactan V"
                                Case 8
                                    Vector(Lugar, 4) = "Envio a Pellital V"
                                Case 9
                                    Vector(Lugar, 4) = "Envio a Pellital IV"
                                Case 10
                                    Vector(Lugar, 4) = "Envio a Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = ""
                            ZCantidadAnt = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(ZCantidadAnt))
                            
                            WXSalidas = WXSalidas + ZCantidadAnt
                            Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            Vector(Lugar, 10) = ""
                        End If
                        Vector(Lugar, 7) = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                        
                        
                    End If
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
    
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    Rem PROCESA LAS movimietnos varios de laboratorio
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
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
                
                If rstMovlab!FechaOrd < "20001218" Or rstMovlab!FechaOrd > WOrdFechaCierre Then
                
                        Else
                
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = Articulo.Text Then
                
                        WArticulo = rstMovlab!Articulo
                        WCantidad = rstMovlab!Cantidad
                        WFecha = rstMovlab!Fecha
                        WCodigo = rstMovlab!Codigo
                        WMovi = rstMovlab!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovlab!Fecha
                        Vector(Lugar, 2) = "Mov.Lab"
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovlab!Observaciones
                        If rstMovlab!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovlab!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovlab!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(rstMovlab!Cantidad))
                            WXSalidas = WXSalidas + rstMovlab!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                        Vector(Lugar, 10) = ""
                        
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
        
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    Rem PROCESA LAS VENTAS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WFec = IIf(IsNull(rstEstadistica!Fecha), "", rstEstadistica!Fecha)
                
                If WFec <> "" Then
                
                XFec = Right$(rstEstadistica!Fecha, 4) + Mid$(rstEstadistica!Fecha, 4, 2) + Left$(rstEstadistica!Fecha, 2)
                If XFec < "20001218" Or XFec > WOrdFechaCierre Then
                
                        Else
                
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = Articulo.Text Then
                
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        
                        XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                        XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                        XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                        XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                        XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                        XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                        XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                        XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                    
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                            XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                            XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                            XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                            XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                            XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                            XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                            XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                            XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                            XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                            XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                            XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                            XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                            XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                Else
                            XLote(6, 1) = "0"
                            XLote(6, 2) = "0"
                            XLote(7, 1) = "0"
                            XLote(7, 2) = "0"
                            XLote(8, 1) = "0"
                            XLote(8, 2) = "0"
                            XLote(9, 1) = "0"
                            XLote(9, 2) = "0"
                            XLote(10, 1) = "0"
                            XLote(10, 2) = "0"
                            XLote(11, 1) = "0"
                            XLote(11, 2) = "0"
                            XLote(12, 1) = "0"
                            XLote(12, 2) = "0"
                        End If
                
                        For Da = 1 To 12
                
                            WLote = XLote(Da, 1)
                            WCanti = Val(XLote(Da, 2))
                        
                            If WCanti <> 0 Then
                                WCantidad = WCanti
                                Lugar = Lugar + 1
                                Vector(Lugar, 1) = WFecha
                                If rstEstadistica!Tipo = 1 Then
                                    Vector(Lugar, 2) = "Factura"
                                        Else
                                    Vector(Lugar, 2) = "Devol"
                                End If
                                Vector(Lugar, 3) = WCodigo
                                Vector(Lugar, 4) = rstEstadistica!Cliente
                                If rstEstadistica!Tipo = 2 Then
                                    Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                                    Vector(Lugar, 6) = ""
                                    WXEntradas = WXEntradas + WCantidad
                                        Else
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                                    WXSalidas = WXSalidas + WCantidad
                                End If
                                Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                Vector(Lugar, 9) = WLote
                                Vector(Lugar, 10) = ""
                            End If
                        
                        Next Da
                        
                    End If
                
                End If
                
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        rstEstadistica.Close
    End If
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(Dada, 7) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
                
        WVector1.Col = 1
        WVector1.Text = Vector(Cicla, 1)
                        
        WVector1.Col = 2
        WVector1.Text = Vector(Cicla, 2)
                                               
        WVector1.Col = 3
        WVector1.Text = Vector(Cicla, 3)
                        
        WVector1.Col = 4
        WVector1.Text = Vector(Cicla, 4)
                        
        WVector1.Col = 5
        WVector1.Text = Vector(Cicla, 5)
                
        WVector1.Col = 6
        WVector1.Text = Vector(Cicla, 6)
        
        WVector1.Col = 7
        WVector1.Text = Vector(Cicla, 9)
        
        Rem WVector1.Col = 8
        Rem WVector1.Text = Vector(Cicla, 10)
    
    Next Cicla
    
    WXStock = WXInicial + WXEntradas - WXSalidas
    
    Rem XInicial.Text = Pusing("###,###.##", Str$(WXInicial))
    Rem XEntradas.Text = Pusing("###,###.##", Str$(WXEntradas))
    Rem XSalidas.Text = Pusing("###,###.##", Str$(WXSalidas))
    Rem XStock.Text = Pusing("###,###.##", Str$(WXStock))
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1
    
    WVector1.SetFocus

End Sub

Private Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        WArticulo = Articulo.Text
        Articulo.Text = WArticulo
        
        spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Call Proceso_Click
            WVector1.SetFocus
                Else
            Articulo.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 10001
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WVector1.Text = "Entradas"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Salidas"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Historial_Click()

    WVector1.Col = 2
    WMovimiento = WVector1.Text
    
    If WMovimiento = "Laudo" Then
        
        WVector1.Col = 3
        WLaudo = WVector1.Text
        
        HistoriaOrden.Text = ""
        HistoriaInforme.Text = ""
        HistoriaRemito.Text = ""
        HistoriaFactura.Text = ""
        HistoriaCarpeta.Text = ""
        
        HistoriaFechaOrden.Text = "  /  /    "
        HistoriaFechaInforme.Text = "  /  /    "
        HistoriaFechaFactura.Text = "  /  /    "
        
        WOrden = ""
        WInforme = ""
        WRemito = ""
        WFactura = ""
        WProveedor = ""
        WCarpeta = ""
        
        WFechaOrden = "  /  /    "
        WFechaInforme = "  /  /    "
        WFechaFactura = "  /  /    "
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + WLaudo + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + Articulo.Text + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WInforme = Str$(rstLaudo!Informe)
            WOrden = rstLaudo!Orden
            rstLaudo.Close
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Informe.Informe = " + "'" + WInforme + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                WRemito = Str$(rstInforme!Remito)
                WFechaInforme = rstInforme!Fecha
                rstInforme.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Orden.Orden = " + "'" + WOrden + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WFechaOrden = rstOrden!Fecha
                WProveedor = rstOrden!Proveedor
                WCarpeta = rstOrden!Carpeta
                rstOrden.Close
            End If
            
            XEmpresa = Wempresa
            
            
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM IvaComp"
            ZSql = ZSql + " Where IvaComp.Remito LIKE " + "'" + "%" + Trim(WRemito) + "%" + "'"
            ZSql = ZSql + " and IvaComp.Proveedor = " + "'" + WProveedor + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                WFactura = Str$(Val(rstIvaComp!Numero))
                WFechaFactura = rstIvaComp!Fecha
                rstIvaComp.Close
            End If
            
            Call Conecta_Empresa
        
        End If
        
        HistoriaOrden.Text = WOrden
        HistoriaInforme.Text = WInforme
        HistoriaRemito.Text = WRemito
        HistoriaFactura.Text = WFactura
        HistoriaCarpeta.Text = WCarpeta
        
        HistoriaFechaOrden.Text = WFechaOrden
        HistoriaFechaInforme.Text = WFechaInforme
        HistoriaFechaFactura.Text = WFechaFactura
        
        PantaHistoria.Visible = True
        
    End If
    
End Sub

Private Sub PantaHistoriaCierra_Click()
    PantaHistoria.Visible = False
End Sub
