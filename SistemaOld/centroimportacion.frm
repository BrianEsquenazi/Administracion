VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCentroImportacion 
   AutoRedraw      =   -1  'True
   Caption         =   "Central de control de importaciones"
   ClientHeight    =   10830
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   15135
   LinkTopic       =   "Form2"
   ScaleHeight     =   10830
   ScaleWidth      =   15135
   Begin VB.ComboBox Activas 
      Height          =   315
      Left            =   12960
      TabIndex        =   50
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox SumaArticulo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame PantaOrden 
      Height          =   975
      Left            =   2880
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox FiltroOrden 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   17
         Text            =   " "
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox SumaDespacho 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox SumaLetra 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox FiltroIII 
      Height          =   315
      Left            =   11520
      TabIndex        =   39
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox FiltroII 
      Height          =   315
      Left            =   9120
      TabIndex        =   38
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame PantaVto 
      Height          =   975
      Left            =   1440
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   8055
      Begin MSMask.MaskEdBox FiltroVtoI 
         Height          =   300
         Left            =   4080
         TabIndex        =   31
         Top             =   360
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FiltroVtoII 
         Height          =   300
         Left            =   5760
         TabIndex        =   32
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Desde - Hasta Vencimiento de la Letra"
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton Actualiza 
      Caption         =   "Actualiza"
      Height          =   375
      Left            =   13680
      TabIndex        =   37
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame PantaArticulo 
      Height          =   4695
      Left            =   1800
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ListBox WIndice 
         Height          =   1035
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox AyudaII 
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
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ListBox PantallaII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         ItemData        =   "centroimportacion.frx":0000
         Left            =   1680
         List            =   "centroimportacion.frx":0007
         TabIndex        =   46
         Top             =   1200
         Visible         =   0   'False
         Width           =   4695
      End
      Begin MSMask.MaskEdBox FiltroArticulo 
         Height          =   300
         Left            =   4080
         TabIndex        =   36
         Top             =   360
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
      Begin VB.Label Label9 
         Caption         =   "Materia Prima "
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
         Left            =   2160
         TabIndex        =   35
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame PantaLlegada 
      Height          =   975
      Left            =   1440
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   8055
      Begin MSMask.MaskEdBox FiltroLLegadaI 
         Height          =   300
         Left            =   4080
         TabIndex        =   27
         Top             =   360
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FiltroLLegadaII 
         Height          =   300
         Left            =   5760
         TabIndex        =   28
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         Caption         =   "Desde - Hasta Fecha de Prevista de llegada"
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
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame PantaCarpeta 
      Height          =   975
      Left            =   3720
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox FiltroCarpeta 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   24
         Text            =   " "
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame PantaFecha 
      Height          =   975
      Left            =   1440
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   8055
      Begin MSMask.MaskEdBox FiltroFechaI 
         Height          =   300
         Left            =   4080
         TabIndex        =   21
         Top             =   360
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FiltroFechaII 
         Height          =   300
         Left            =   5760
         TabIndex        =   22
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Desde - Hasta Fecha de Orden de Compra"
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.ComboBox FiltroI 
      Height          =   315
      Left            =   6600
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox OrdenaI 
      Height          =   315
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame PantaExporta 
      Height          =   4695
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   2760
         TabIndex        =   9
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox NombreExporta 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton CancelaExporta 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton ConfirmaExporta 
         Caption         =   "Confirma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Exportaii 
      Caption         =   "Exportacion (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin Crystal.CrystalReport ListaGRilla 
      Left            =   14400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ControlImpoImpre.rpt"
   End
   Begin VB.ListBox Lista 
      Height          =   645
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      ItemData        =   "centroimportacion.frx":0015
      Left            =   1920
      List            =   "centroimportacion.frx":001C
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   7335
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12726
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.Label ImpreArticuloII 
      Caption         =   "Kgs."
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
      Left            =   12120
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label ImpreArticulo 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9120
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "U$S Letra"
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
      TabIndex        =   43
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "$ Despacho"
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
      Left            =   2280
      TabIndex        =   42
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Filtro"
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
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Ordenamiento"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgCentroImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ver As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Dim Empe(20, 10) As String
Dim XParam As String
Rem Dim Auxiliar(20000)
Rem Dim WPasa(20000) As String
Dim XEmpresa As String

Dim ColumnaOpcion As Integer
Dim ColumnaOpcionII As Integer
Dim ColumnaOpcionIII As Integer

Dim Seleccion As String
Dim SeleccionII As String

Dim SeleccionIII As String
Dim SeleccionIV As String

Dim SeleccionV As String
Dim SeleccionVI As String

Dim ZZOrdena As Integer

Dim ZZOrdenaFiltro As Integer
Dim WDesdeFecha As String
Dim WHastaFecha As String

Dim ZZFiltro(2000, 2) As String
Dim ZZLugarFiltro As Integer
Dim ZZTipoFiltro As Integer

Dim ZZFiltroOrdenI As String
Dim ZZFiltroOrdenII As String
Dim ZZFiltroOrdenIII As String


Dim ZZColumnaI As Integer
Dim ZZColumnaII As Integer
Dim ZZColumnaIII As Integer

Dim ZZSumaDespacho As String
Dim ZZSumaLetra As String

Dim ZCarpeta As String
Dim ZSaldo As Double
Dim ZProcesa  As String
Dim ZPasa(5000, 30) As String

Dim ZFechaDJai As String
Dim WFecha As String
Dim WDias1 As Integer
Dim Wvencimiento As String



Private Sub Actualiza_Click()
    Call Proceso_Click
End Sub

Private Sub CancelaExporta_Click()
    PantaExporta.Visible = False
    Rem Dir1.Path = "\\193.168.0.2\g$\vb"
    Rem ChDir "\\193.168.0.2\g$\vb"
End Sub

Private Sub Command1_Click()

    For Ciclo = 1 To 1000
    
        ZOrden = Muestra.TextMatrix(Ciclo, 1)
        
        If Val(ZOrden) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Orden = " + "'" + ZOrden + "'"
            ZSql = ZSql + " Order by Orden.Clave"
            
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
            
                ZZVtoLetra = IIf(IsNull(rstOrden!VtoLetra), "  /  /    ", rstOrden!VtoLetra)
                ZZVtoLetraII = IIf(IsNull(rstOrden!VtoLetraII), "  /  /    ", rstOrden!VtoLetraII)
                rstOrden.Close
                
                If Trim(ZZVtoLetra) = "" Or ZZVtoLetra = "  /  /    " Then
                
                    ZZVtoLetra = ZZVtoLetraII
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Orden SET "
                    ZSql = ZSql + " VtoLetra = " + "'" + ZZVtoLetra + "'"
                    ZSql = ZSql + " Where Orden = " + "'" + ZOrden + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
            End If
        
        End If
    
    Next Ciclo

End Sub

Private Sub ConfirmaExporta_Click()

    Rem If NombreExporta.Text = "" Then
    Rem     m$ = "Se debe informar un nombre de archivo"
    Rem     A% = MsgBox(m$, 0, "Exportacion de Muestras")
    Rem     Exit Sub
    Rem End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With

    ZSql = ""
    ZSql = ZSql + "DELETE ControlImpoImpre"
    spControlImpoImpre = ZSql
    Set rstControlImpoImpre = db.OpenRecordset(spControlImpoImpre, dbOpenSnapshot, dbSQLPassThrough)
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
    
        ZOrden = Muestra.TextMatrix(Ciclo, 1)
        ZPta = Muestra.TextMatrix(Ciclo, 2)
        ZFecha = Muestra.TextMatrix(Ciclo, 3)
        ZProveedor = Muestra.TextMatrix(Ciclo, 4)
        ZMoneda = Muestra.TextMatrix(Ciclo, 5)
        ZCarpeta = Muestra.TextMatrix(Ciclo, 6)
        ZDJai = Muestra.TextMatrix(Ciclo, 7)
        ZOrigen = Muestra.TextMatrix(Ciclo, 8)
        ZIncoterms = Muestra.TextMatrix(Ciclo, 9)
        ZTransporte = Muestra.TextMatrix(Ciclo, 10)
        ZFLLegada = Muestra.TextMatrix(Ciclo, 11)
        ZTPago = Muestra.TextMatrix(Ciclo, 12)
        ZDespacho = Muestra.TextMatrix(Ciclo, 13)
        ZPagoDespacho = Muestra.TextMatrix(Ciclo, 14)
        ZLetra = Muestra.TextMatrix(Ciclo, 15)
        ZPagoLetra = Muestra.TextMatrix(Ciclo, 16)
        ZVtoLetra = Muestra.TextMatrix(Ciclo, 17)
        ZPagoParcial = Muestra.TextMatrix(Ciclo, 18)
        ZFEmbarque = Muestra.TextMatrix(Ciclo, 19)
        
        If ZPagoDespacho = "Pendiente" Then
            ZSumaI = ZDespacho
                Else
            ZSumaI = "0"
        End If
        
        If ZPagoLetra = "Pendiente" Then
            ZSumaII = Str(Val(ZLetra) - Val(ZPagoParcial))
                Else
            ZSumaII = "0"
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ControlImpoImpre ("
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Pta ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Moneda ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Djai ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Incoterms ,"
        ZSql = ZSql + "Transporte,"
        ZSql = ZSql + "FLLegada  ,"
        ZSql = ZSql + "TPago ,"
        ZSql = ZSql + "SumaI ,"
        ZSql = ZSql + "SumaII ,"
        ZSql = ZSql + "Despacho ,"
        ZSql = ZSql + "PagoDespacho ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "PagoLetra ,"
        ZSql = ZSql + "VtoLetra ,"
        ZSql = ZSql + "PagoParcial ,"
        ZSql = ZSql + "FEmbarque) "
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZOrden + "',"
        ZSql = ZSql + "'" + ZPta + "',"
        ZSql = ZSql + "'" + ZFecha + "',"
        ZSql = ZSql + "'" + ZProveedor + "',"
        ZSql = ZSql + "'" + ZMoneda + "',"
        ZSql = ZSql + "'" + ZCarpeta + "',"
        ZSql = ZSql + "'" + ZDJai + "',"
        ZSql = ZSql + "'" + ZOrigen + "',"
        ZSql = ZSql + "'" + ZIncoterms + "',"
        ZSql = ZSql + "'" + ZTransporte + "',"
        ZSql = ZSql + "'" + ZFLLegada + "',"
        ZSql = ZSql + "'" + ZTPago + "',"
        ZSql = ZSql + "'" + ZSumaI + "',"
        ZSql = ZSql + "'" + ZSumaII + "',"
        ZSql = ZSql + "'" + ZDespacho + "',"
        ZSql = ZSql + "'" + ZPagoDespacho + "',"
        ZSql = ZSql + "'" + ZLetra + "',"
        ZSql = ZSql + "'" + ZPagoLetra + "',"
        ZSql = ZSql + "'" + ZVtoLetra + "',"
        ZSql = ZSql + "'" + ZPagoParcial + "',"
        ZSql = ZSql + "'" + ZFEmbarque + "')"
       
        spControlImpoImpre = ZSql
        Set rstControlImpoImpre = db.OpenRecordset(spControlImpoImpre, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    DoEvents

    ListaGRilla.WindowTitle = ""
    ListaGRilla.WindowTop = 0
    ListaGRilla.WindowLeft = 0
    ListaGRilla.WindowWidth = Screen.Width
    ListaGRilla.WindowHeight = Screen.Height


    ListaGRilla.Destination = 0
    Rem ListaGRilla.PrintFileType = crptExcel50
    Rem ListaGRilla.PrintFileName = Dir1.Path + "\" + NombreExporta.Text + ".xls"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT ControlImpoImpre.Orden, ControlImpoImpre.Pta, ControlImpoImpre.Fecha, ControlImpoImpre.Proveedor, ControlImpoImpre.Moneda, ControlImpoImpre.Carpeta, ControlImpoImpre.Djai, ControlImpoImpre.Origen, ControlImpoImpre.Incoterms, ControlImpoImpre.Transporte, ControlImpoImpre.FLLegada, ControlImpoImpre.TPago, ControlImpoImpre.Despacho, ControlImpoImpre.PagoDespacho, ControlImpoImpre.Letra, ControlImpoImpre.PagoLetra, ControlImpoImpre.VtoLetra " _
        + "From " _
        + DSQ + ".dbo.ControlImpoImpre ControlImpoImpre " _
        + "Where " _
        + "ControlImpoImpre.Orden >= 0 AND " _
        + "ControlImpoImpre.Orden <= 999999"
    
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
    PantaExporta.Visible = False
    
    Rem by nan
    Rem Dir1.Path = "\\193.168.0.2\g$\vb"
    Rem by nan
    
End Sub

Private Sub Exportaii_Click()

    NombreExporta.Text = ""
     
    Drive1.Drive = "C:"
    
    
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
  
    Rem PantaExporta.Visible = True
    Call ConfirmaExporta_Click

End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub cmdClose_Click()
    PrgAju.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Impresion_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With

    spMuestraImpre = "BorrarMuestraImpre "
    Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
        ZNumero = Str$(Ciclo)
        ZPedido = Left$(Muestra.TextMatrix(Ciclo, 1), 6)
        ZFecha = Left$(Muestra.TextMatrix(Ciclo, 2), 10)
        ZCodigo = Left$(Muestra.TextMatrix(Ciclo, 3), 15)
        ZDescripcion = Left$(Muestra.TextMatrix(Ciclo, 4), 50)
        ZCantidad = Left$(Muestra.TextMatrix(Ciclo, 5), 10)
        ZDescriCliente = Left$(Muestra.TextMatrix(Ciclo, 6), 50)
        ZCliente = Left$(Muestra.TextMatrix(Ciclo, 7), 50)
        ZObservaciones = Left$(Muestra.TextMatrix(Ciclo, 8), 50)
        ZFecha2 = Left$(Muestra.TextMatrix(Ciclo, 9), 10)
        ZRemito = Left$(Muestra.TextMatrix(Ciclo, 10), 10)
        ZHojaRuta = Left$(Muestra.TextMatrix(Ciclo, 11), 10)
        ZCodigo2 = Left$(Muestra.TextMatrix(Ciclo, 12), 15)
        ZDescripcion2 = Left$(Muestra.TextMatrix(Ciclo, 13), 50)
        ZLote = Left$(Muestra.TextMatrix(Ciclo, 14), 10)
        ZObservaciones2 = Left$(Muestra.TextMatrix(Ciclo, 15), 50)
        ZCantidad2 = Left$(Muestra.TextMatrix(Ciclo, 16), 10)
        ZActualiza = Left$(Muestra.TextMatrix(Ciclo, 17), 1)
        ZFechaEmbarque = Left$(Muestra.TextMatrix(Ciclo, 19), 10)
        
        Sql1 = "INSERT INTO MuestraImpre ("
        Sql2 = "Numero ,"
        Sql3 = "Fecha ,"
        Sql4 = "Codigo ,"
        Sql5 = "Descripcion ,"
        Sql6 = "Cantidad ,"
        Sql7 = "DescriCliente ,"
        Sql8 = "Cliente ,"
        Sql9 = "Observaciones ,"
        Sql10 = "Fecha2 ,"
        Sql11 = "Codigo2 ,"
        Sql12 = "Descripcion2 ,"
        Sql13 = "Lote ,"
        Sql14 = "Observaciones2 ,"
        Sql15 = "Cantidad2 ,"
        Sql16 = "Actualiza ,"
        Sql17 = "DesEmpresa) "
        Sql18 = "Values ("
        Sql19 = "'" + ZNumero + "',"
        Sql20 = "'" + ZFecha + "',"
        Sql21 = "'" + ZCodigo + "',"
        Sql22 = "'" + ZDescripcion + "',"
        Sql23 = "'" + ZCantidad + "',"
        Sql24 = "'" + ZDescriCliente + "',"
        Sql25 = "'" + ZCliente + "',"
        Sql26 = "'" + ZObservaciones + "',"
        Sql27 = "'" + ZFecha2 + "',"
        Sql28 = "'" + ZCodigo2 + "',"
        Sql29 = "'" + ZDescripcion2 + "',"
        Sql30 = "'" + ZLote + "',"
        Sql31 = "'" + ZObservaciones2 + "',"
        Sql32 = "'" + ZCantidad2 + "',"
        Sql33 = "'" + ZActualiza + "',"
        Sql34 = "'" + WDesEmpresa + "')"
       
        spMuestraImpre = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                     Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                     Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                     Sql31 + Sql32 + Sql33 + Sql34
        Set rstMuestraImpre = db.OpenRecordset(spMuestraImpre, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo

    ListaGRilla.Destination = 1
    Rem ListaGRilla.Destination = 0
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    ListaGRilla.SQLQuery = "SELECT MuestraImpre.Numero, MuestraImpre.Fecha, MuestraImpre.Codigo, MuestraImpre.Descripcion, MuestraImpre.Cantidad, MuestraImpre.DescriCLiente, MuestraImpre.Cliente, MuestraImpre.Observaciones, MuestraImpre.Fecha2, MuestraImpre.Codigo2, MuestraImpre.Descripcion2, MuestraImpre.Lote, MuestraImpre.Observaciones2, MuestraImpre.Cantidad2 " _
                    + "From " _
                    + DSQ + ".dbo.MuestraImpre MuestraImpre " _
                    + "Where " _
                    + "MuestraImpre.Numero >= 0 AND " _
                    + "MuestraImpre.Numero <= 999999 " _
                    + "Order By MuestraImpre.Numero ASC"
    ListaGRilla.Connect = Connect()
    ListaGRilla.Action = 1
    
End Sub

Private Sub Proceso_Click()

    If ZProcesa = "N" Then
        Exit Sub
    End If

    Pantalla.Visible = False
    Call Limpia_Vector
    WLugar = 0
    
    If Seleccion = "" Then
        ColumnaOpcion = 0
    End If
    
    If SeleccionIII = "" Then
        ColumnaOpcionII = 0
    End If
    
    If SeleccionV = "" Then
        ColumnaOpcionIII = 0
    End If
    
    
    XEmpresa = Wempresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0007"
        Empe(4, 2) = "Empresa07"
        Empe(5, 1) = "0010"
        Empe(5, 2) = "Empresa10"
        Empe(6, 1) = "0011"
        Empe(6, 2) = "Empresa11"
        XHasta = 6
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    
    If ColumnaOpcion = 2 Then
        ZZDesde = Seleccion
        ZZHasta = Seleccion
            Else
        ZZDesde = 1
        ZZHasta = XHasta
    End If
    
    ZZSumaLetra = 0
    ZZSumaDespacho = 0
    ZZSumaArticulo = 0
    
    ImpreArticulo.Visible = False
    ImpreArticuloII.Visible = False
    SumaArticulo.Visible = False
        
    For CiclaEmpresa = ZZDesde To ZZHasta
    
        Wempresa = Empe(CiclaEmpresa, 1)
        txtOdbc = Empe(CiclaEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Select Case ColumnaOpcion
            Case 0, 1, 2
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
                
                
            Case 3
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.FechaOrd >= " + "'" + Seleccion + "'"
                ZSql = ZSql + " and Orden.FechaOrd <= " + "'" + SeleccionII + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 4
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.Proveedor = " + "'" + Seleccion + "'"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 6
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.Djai = " + "'" + Seleccion + "'"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 7
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Origen = " + "'" + Seleccion + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 8
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.Leyenda = " + "'" + Seleccion + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 9
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.TipoImpo = " + "'" + Seleccion + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 10
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                WDesdeFecha = Seleccion
                WHastaFecha = SeleccionII
                
            Case 11
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.TipoPago = " + "'" + Seleccion + "'"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 12
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.fechaord, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " and Orden.PagoDespacho = " + "'" + Seleccion + "'"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 13
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fechaord, Orden.fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                Rem ZSql = ZSql + " and Orden.PagoLetra = " + "'" + Seleccion + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case 14
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fechaord, Orden.fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Renglon = 1"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                WDesdeFecha = Right$(FiltroVtoI.Text, 4) + Mid$(FiltroVtoI.Text, 4, 2) + Left$(FiltroVtoI.Text, 2)
                WHastaFecha = Right$(FiltroVtoII.Text, 4) + Mid$(FiltroVtoII.Text, 4, 2) + Left$(FiltroVtoII.Text, 2)
                
            Case 15
                ZSql = ""
                ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fechaord, Orden.fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
                ZSql = ZSql + " FROM Orden, Proveedor"
                ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
                ZSql = ZSql + " and Orden.Tipo = 1"
                ZSql = ZSql + " and fechaord >=20140101"
                ZSql = ZSql + " and Orden.Cantidad <> 0"
                ZSql = ZSql + " and Orden.Articulo = " + "'" + Seleccion + "'"
                ZSql = ZSql + " Order by Orden.Clave"
                spOrden = ZSql
                
            Case Else
        End Select
        
                
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
        
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                    
                        Rem by nan
                        If Activas.ListIndex <> 1 Then
                            If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                                ZEntra = "N"
                                    Else
                                ZEntra = "S"
                            End If
                                Else
                            If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                                aa = rstOrden!Orden
                                ZEntra = "S"
                                    Else
                                ZEntra = "N"
                            End If
                        End If
                    
                        If ColumnaOpcion = 10 Then
                            ZZLlegada = rstOrden!FechaLlegada
                            ZZOrdLLegada = Right$(ZZLlegada, 4) + Mid$(ZZLlegada, 4, 2) + Left$(ZZLlegada, 2)
                            If WDesdeFecha > ZZOrdLLegada Or WHastaFecha < ZZOrdLLegada Then
                                ZEntra = "N"
                            End If
                        End If
                        
                        If ColumnaOpcion = 14 Then
                            ZZVtoLetra = rstOrden!VtoLetra
                            ZZOrdVtoLetra = Right$(ZZVtoLetra, 4) + Mid$(ZZVtoLetra, 4, 2) + Left$(ZZVtoLetra, 2)
                            If WDesdeFecha > ZZOrdVtoLetra Or WHastaFecha < ZZOrdVtoLetra Then
                                ZEntra = "N"
                            End If
                        End If
                        
                        If ZEntra = "S" And ColumnaOpcionII <> 0 Then
                        
                            Select Case ColumnaOpcionII
                                Case 1
                                    If CiclaEmpresa <> SeleccionIII Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 2
                                    If rstOrden!FechaOrd < SeleccionIII Or rstOrden!FechaOrd > SeleccionIV Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 3
                                    If Trim(rstOrden!Proveedor) <> Trim(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 4
                                    If Trim(rstOrden!Origen) <> Trim(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 5
                                    If Trim(rstOrden!Leyenda) <> Trim(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 6
                                    If Trim(rstOrden!TipoImpo) <> Trim(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 7
                                    ZZLlegada = rstOrden!FechaLlegada
                                    ZZOrdLLegada = Right$(ZZLlegada, 4) + Mid$(ZZLlegada, 4, 2) + Left$(ZZLlegada, 2)
                                    WDesdeFecha = SeleccionIII
                                    WHastaFecha = SeleccionIV
                                    If WDesdeFecha > ZZOrdLLegada Or WHastaFecha < ZZOrdLLegada Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 8
                                    If rstOrden!TipoPago <> Val(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 9
                                    If rstOrden!PagoDespacho <> Val(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 10
                                    Rem If rstOrden!PagoLetra <> Val(SeleccionIII) Then
                                    Rem     ZEntra = "N"
                                    Rem End If
                                    
                                Case 11
                                    ZZVtoLetra = rstOrden!VtoLetra
                                    ZZOrdVtoLetra = Right$(ZZVtoLetra, 4) + Mid$(ZZVtoLetra, 4, 2) + Left$(ZZVtoLetra, 2)
                                    WDesdeFecha = SeleccionIII
                                    WHastaFecha = SeleccionIV
                                    If WDesdeFecha > ZZOrdVtoLetra Or WHastaFecha < ZZOrdVtoLetra Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 12
                                    If UCase(rstOrden!Articulo) <> UCase(SeleccionIII) Then
                                        ZEntra = "N"
                                    End If
                                    
                                    
                                Case Else
                                
                            End Select
                        
                        End If
                        
                        
                        If ZEntra = "S" And ColumnaOpcionIII <> 0 Then
                        
                            Select Case ColumnaOpcionIII
                                Case 1
                                    If CiclaEmpresa <> SeleccionV Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 2
                                    If rstOrden!FechaOrd < SeleccionV Or rstOrden!FechaOrd > SeleccionVI Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 3
                                    If Trim(rstOrden!Proveedor) <> Trim(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 4
                                    If Trim(rstOrden!Origen) <> Trim(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 5
                                    If Trim(rstOrden!Leyenda) <> Trim(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 6
                                    If Trim(rstOrden!TipoImpo) <> Trim(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 7
                                    ZZLlegada = rstOrden!FechaLlegada
                                    ZZOrdLLegada = Right$(ZZLlegada, 4) + Mid$(ZZLlegada, 4, 2) + Left$(ZZLlegada, 2)
                                    WDesdeFecha = SeleccionV
                                    WHastaFecha = SeleccionVI
                                    If WDesdeFecha > ZZOrdLLegada Or WHastaFecha < ZZOrdLLegada Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 8
                                    If rstOrden!TipoPago <> Val(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 9
                                    If rstOrden!PagoDespacho <> Val(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 10
                                    Rem If rstOrden!PagoLetra <> Val(SeleccionV) Then
                                    Rem     ZEntra = "N"
                                    Rem End If
                                    
                                Case 11
                                    ZZVtoLetra = rstOrden!VtoLetra
                                    ZZOrdVtoLetra = Right$(ZZVtoLetra, 4) + Mid$(ZZVtoLetra, 4, 2) + Left$(ZZVtoLetra, 2)
                                    WDesdeFecha = SeleccionV
                                    WHastaFecha = SeleccionVI
                                    If WDesdeFecha > ZZOrdVtoLetra Or WHastaFecha < ZZOrdVtoLetra Then
                                        ZEntra = "N"
                                    End If
                                    
                                Case 12
                                    If UCase(rstOrden!Articulo) <> UCase(SeleccionV) Then
                                        ZEntra = "N"
                                    End If
                                    
                                    
                                Case Else
                                
                            End Select
                        
                        End If
                        
                        
                        
                        
                        If ZEntra = "S" Then
                        
                            WLugar = WLugar + 1
                                                
                            ZDJai = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
                            
                            Muestra.TextMatrix(WLugar, 1) = rstOrden!Orden
                            Select Case CiclaEmpresa
                                Case 1
                                    Muestra.TextMatrix(WLugar, 2) = "I"
                                Case 2
                                    Muestra.TextMatrix(WLugar, 2) = "II"
                                Case 3
                                    Muestra.TextMatrix(WLugar, 2) = "III"
                                Case 4
                                    Muestra.TextMatrix(WLugar, 2) = "V"
                                Case 5
                                    Muestra.TextMatrix(WLugar, 2) = "VI"
                                Case 5
                                    Muestra.TextMatrix(WLugar, 2) = "VII"
                                Case Else
                            End Select
                            Muestra.TextMatrix(WLugar, 3) = Left$(rstOrden!Fecha, 5) + "/" + Mid$(rstOrden!Fecha, 9, 2)
                            Muestra.TextMatrix(WLugar, 4) = rstOrden!WProveedor
                            
                            Select Case rstOrden!Moneda
                                Case 0
                                    Muestra.TextMatrix(WLugar, 5) = "U$S"
                                Case 1
                                    Muestra.TextMatrix(WLugar, 5) = "$"
                                Case 2
                                    Muestra.TextMatrix(WLugar, 5) = "Eur"
                            End Select
                            
                            Muestra.TextMatrix(WLugar, 6) = rstOrden!Carpeta
                            
                            Muestra.TextMatrix(WLugar, 7) = ZDJai
                            
                            
                            
                            
                            Muestra.TextMatrix(WLugar, 8) = rstOrden!Origen
                            
                            Select Case rstOrden!Leyenda
                                Case 1
                                    Muestra.TextMatrix(WLugar, 9) = "FOB"
                                Case 2
                                    Muestra.TextMatrix(WLugar, 9) = "CIF"
                                Case 3
                                    Muestra.TextMatrix(WLugar, 9) = "CFR"
                                Case 4
                                    Muestra.TextMatrix(WLugar, 9) = "CPT"
                                Case 5
                                    Muestra.TextMatrix(WLugar, 9) = "EXW"
                                Case 6
                                    Muestra.TextMatrix(WLugar, 9) = "FCA"
                                Case Else
                                    Muestra.TextMatrix(WLugar, 9) = ""
                            End Select
                            
                            
                            Select Case rstOrden!TipoImpo
                                Case 1
                                    Muestra.TextMatrix(WLugar, 10) = "Maritimo"
                                Case 2
                                    Muestra.TextMatrix(WLugar, 10) = "Terrestre"
                                Case 3
                                    Muestra.TextMatrix(WLugar, 10) = "Aereo"
                                Case Else
                                    Muestra.TextMatrix(WLugar, 10) = ""
                            End Select
                            
                            Muestra.TextMatrix(WLugar, 11) = IIf(IsNull(rstOrden!FechaLlegada), "", rstOrden!FechaLlegada)
                            
                            Select Case rstOrden!TipoPago
                                Case 1
                                    Muestra.TextMatrix(WLugar, 12) = "Pago Anti."
                                Case 2
                                    Muestra.TextMatrix(WLugar, 12) = "A la vista"
                                Case 3
                                    Muestra.TextMatrix(WLugar, 12) = "Cta.Cte."
                                Case Else
                                    Muestra.TextMatrix(WLugar, 12) = ""
                            End Select
                            
                            Muestra.TextMatrix(WLugar, 13) = IIf(IsNull(rstOrden!ImpoDespacho), "0", rstOrden!ImpoDespacho)
                            Muestra.TextMatrix(WLugar, 13) = Pusing("###,###", Muestra.TextMatrix(WLugar, 13))
                            Select Case rstOrden!PagoDespacho
                                Case 0
                                    ZZSumaDespacho = ZZSumaDespacho + Val(Muestra.TextMatrix(WLugar, 13))
                                    Muestra.TextMatrix(WLugar, 14) = "Pendiente"
                                Case Else
                                    Muestra.TextMatrix(WLugar, 14) = "Pagado"
                            End Select
                                    
                            
                            
                            saaa = rstOrden!Carpeta
                            
                            
                            Muestra.TextMatrix(WLugar, 15) = IIf(IsNull(rstOrden!ImpoLetra), "0", rstOrden!ImpoLetra)
                            Muestra.TextMatrix(WLugar, 15) = Pusing("###,###", Muestra.TextMatrix(WLugar, 15))
                            Select Case rstOrden!PagoLetra
                                Case 0
                                    Muestra.TextMatrix(WLugar, 16) = "Pendiente"
                                Case Else
                                    Muestra.TextMatrix(WLugar, 16) = "Pagado"
                            End Select
                            Muestra.TextMatrix(WLugar, 17) = IIf(IsNull(rstOrden!VtoLetra), "0", rstOrden!VtoLetra)
                            
                            If ColumnaOpcion = 15 Then
                                    ZZSumaArticulo = ZZSumaArticulo + rstOrden!Cantidad
                            End If
                            
                            Muestra.TextMatrix(WLugar, 19) = IIf(IsNull(rstOrden!FechaEmbarque), "", rstOrden!FechaEmbarque)
                            Muestra.TextMatrix(WLugar, 20) = IIf(IsNull(rstOrden!FechaDJai), "", rstOrden!FechaDJai)
                            Muestra.TextMatrix(WLugar, 21) = rstOrden!Proveedor
                            
                            
                            
                            
                            Rem Muestra.TextMatrix(WLugar, 5) = rstOrden!Articulo
                            Rem Muestra.TextMatrix(WLugar, 6) = rstOrden!Cantidad
                            Rem Muestra.TextMatrix(WLugar, 7) = rstOrden!Precio
                            Rem Muestra.TextMatrix(WLugar, 12) = rstOrden!Derechos
                        
                        End If
                        
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                    
                    Loop
                End If
            
            End With
            rstOrden.Close
        
        
        
        End If
        
    Next CiclaEmpresa
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To WLugar
        
        Rem If Val(Muestra.TextMatrix(Ciclo, 6)) = 3100 Then Stop
        
        ZProveedor = Muestra.TextMatrix(Ciclo, 21)
        ZPagoLetra = Muestra.TextMatrix(Ciclo, 16)
        ZCarpeta = Muestra.TextMatrix(Ciclo, 6)
        ZTipo = Muestra.TextMatrix(Ciclo, 6)
            

        If ZPagoLetra = "Pagado" Then
        
            ZCarpeta = Muestra.TextMatrix(Ciclo, 6)
            Call Ceros(ZCarpeta, 4)
            ZProveedor = Muestra.TextMatrix(Ciclo, 21)
            ZPagoLetra = Muestra.TextMatrix(Ciclo, 16)
        
            ZNroInterno = 0
            ZTotal = 0
            ZSaldo = 0
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM IvaComp"
            ZSql = ZSql + " Where IvaComp.Proveedor = " + "'" + ZProveedor + "'"
            ZSql = ZSql + " and IvaComp.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and IvaComp.Punto = " + "'" + ZCarpeta + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                ZNroInterno = rstIvaComp!nrointerno
                rstIvaComp.Close
            End If
            
            If ZNroInterno <> 0 Then
                
                Rem ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.NroInterno = " + "'" + Str$(ZNroInterno) + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCtePrv.RecordCount > 0 Then
                    ZSaldo = rstCtaCtePrv!Saldo
                    rstCtaCtePrv.Close
                End If
            
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    Muestra.TextMatrix(Ciclo, 16) = "Pendiente"
                End If
            
                    Else
        
                Muestra.TextMatrix(Ciclo, 16) = "Pendiente"
        
            End If
        
        End If

    
        If Muestra.TextMatrix(Ciclo, 16) = "Pendiente" Then
    
            ZPagado = 0
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.Carpeta = " + "'" + ZCarpeta + "'"
            spPago = ZSql
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            
                            If ZProveedor = rstPago!Proveedor Then
                                If rstPago!Tiporeg = 1 Then
                                    
                                    ZImporte = rstPago!Importe1 / rstPago!Paridad
                                    ZPagado = ZPagado + ZImporte
                                End If
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPago.Close
            End If
                    
            If ZPagado <> 0 Then
                Muestra.TextMatrix(Ciclo, 18) = Str$(ZPagado)
                Muestra.TextMatrix(Ciclo, 18) = Pusing("###,###", Muestra.TextMatrix(Ciclo, 18))
            
                ZPagado = ZPagado * 1.02
                If ZPagado > Val(Muestra.TextMatrix(Ciclo, 15)) Then
                    Muestra.TextMatrix(Ciclo, 16) = "Pagado"
                    Muestra.TextMatrix(Ciclo, 18) = ""
                End If
            
            End If
                    
        End If
        
        ZPagoLetra = Muestra.TextMatrix(Ciclo, 16)
        Select Case ZPagoLetra
            Case "Pendiente"
                ZTipoPagoLetra = "0"
            Case Else
                ZTipoPagoLetra = "1"
        End Select
    
        If ColumnaOpcion = 13 Then
            If ZTipoPagoLetra <> Val(Seleccion) Then
                For CicloBaja = 1 To 21
                    Muestra.TextMatrix(Ciclo, CicloBaja) = ""
                Next CicloBaja
            End If
        End If
    
        If ColumnaOpcionII = 10 Then
            If ZTipoPagoLetra <> Val(SeleccionIII) Then
                For CicloBaja = 1 To 21
                    Muestra.TextMatrix(Ciclo, CicloBaja) = ""
                Next CicloBaja
            End If
        End If
    
        If ColumnaOpcionIII = 10 Then
            If ZTipoPagoLetra <> Val(SeleccionV) Then
                For CicloBaja = 1 To 21
                    Muestra.TextMatrix(Ciclo, CicloBaja) = ""
                Next CicloBaja
            End If
        End If
    
        Select Case ZTipoPagoLetra
            Case 0
                ZZSumaLetra = ZZSumaLetra + Val(Muestra.TextMatrix(WLugar, 15)) - Val(Muestra.TextMatrix(WLugar, 18))
            Case Else
        End Select
    
    
        ZDJai = Muestra.TextMatrix(Ciclo, 7)
        ZFechaLlegada = Muestra.TextMatrix(Ciclo, 11)
        ZPagoDespacho = Muestra.TextMatrix(Ciclo, 14)
        ZPagoLetra = Muestra.TextMatrix(Ciclo, 16)
        ZFechaLetra = Muestra.TextMatrix(Ciclo, 17)
        ZFechaDJai = Muestra.TextMatrix(Ciclo, 20)
        
        If Trim(ZDJai) <> "" Then
            If Trim(ZFechaDJai) <> "" Then
            
                Call Valida_fecha(ZFechaDJai, Auxi)
                If Auxi = "S" Then
                    WDias1 = 180
                    WFecha = ZFechaDJai
                    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
        
                    ZZOrdFecha = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
                    ZZOrdFechaDjai = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        
                    If ZZOrdFechaDjai < ZZOrdFecha Then
                        Muestra.Row = Ciclo
                        Muestra.Col = 7
                        Muestra.CellBackColor = &H8080FF
                    End If
                End If
            End If
        End If
        
        If ZPagoDespacho = "Pendiente" Then
            ZZOrdFechaLlegada = Right$(ZFechaLlegada, 4) + Mid$(ZFechaLlegada, 4, 2) + Left$(ZFechaLlegada, 2)
            ZZOrdFecha = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
            If ZZOrdFechaLlegada <= ZZOrdFecha Then
                Muestra.Row = Ciclo
                Muestra.Col = 13
                Muestra.CellBackColor = &H8080FF
            End If
        End If
        
        If ZPagoLetra = "Pendiente" Then
            If ZFechaLetra <> "" And ZFechaLetra <> "  /  /    " Then
                ZZOrdFechaLetra = Right$(ZFechaLetra, 4) + Mid$(ZFechaLetra, 4, 2) + Left$(ZFechaLetra, 2)
                ZZOrdFecha = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
                If ZZOrdFechaLetra <= ZZOrdFecha Then
                    Muestra.Row = Ciclo
                    Muestra.Col = 15
                    Muestra.CellBackColor = &H8080FF
                End If
            End If
        End If
        
    Next Ciclo
    
    
    If ColumnaOpcion = 13 Or ColumnaOpcionII = 10 Or ColumnaOpcionIII = 10 Then
    
        Erase ZPasa
        
        For Ciclo = 1 To 4999
            For CicloII = 1 To 21
                ZPasa(Ciclo, CicloII) = Muestra.TextMatrix(Ciclo, CicloII)
            Next CicloII
        Next Ciclo
        
        Call Limpia_Vector
        ZRenglon = 0
        
        For Ciclo = 1 To 4999
            If Val(ZPasa(Ciclo, 1)) <> 0 Then
                ZRenglon = ZRenglon + 1
                For CicloII = 1 To 21
                    Muestra.TextMatrix(ZRenglon, CicloII) = ZPasa(Ciclo, CicloII)
                Next CicloII
            End If
        Next Ciclo
        
    End If
    
    SumaDespacho.Text = Str$(ZZSumaDespacho)
    SumaDespacho.Text = Pusing("###,###,###", SumaDespacho.Text)
    
    SumaLetra.Text = Str$(ZZSumaLetra)
    SumaLetra.Text = Pusing("###,###,###", SumaLetra.Text)
    
    If ColumnaOpcion = 15 Then
        SumaArticulo.Text = Str$(ZZSumaArticulo)
        SumaArticulo.Text = Pusing("###,###,###", SumaArticulo.Text)
        ImpreArticulo.Caption = UCase(Seleccion)
        ImpreArticulo.Visible = True
        ImpreArticuloII.Visible = True
        SumaArticulo.Visible = True
    End If
    
    Muestra.Visible = True
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 1
    Muestra.Row = 1
    Muestra.TopRow = 1

    Rem Muestra.SetFocus
    
End Sub


Private Sub Form_Load()

    Rem **********se modifica tamano para gerencia
    If WOperador <> "17" Then
        Muestra.Height = 9375
        Muestra.Left = 120
        Muestra.Top = 960
        Muestra.Width = 15015
    End If
    
    
    OrdenaI.Clear
    
    OrdenaI.AddItem "Orden"
    OrdenaI.AddItem "Planta"
    OrdenaI.AddItem "Fecha"
    OrdenaI.AddItem "Proveedor"
    OrdenaI.AddItem "Moneda"
    OrdenaI.AddItem "Carpeta"
    OrdenaI.AddItem "DJai"
    OrdenaI.AddItem "Origen"
    OrdenaI.AddItem "Incoterms"
    OrdenaI.AddItem "Tipo"
    OrdenaI.AddItem "Fecha LLegada"
    OrdenaI.AddItem "Tipo Pago"
    OrdenaI.AddItem "Despacho"
    OrdenaI.AddItem "Pago Des"
    OrdenaI.AddItem "Letra"
    OrdenaI.AddItem "Pago Letra"
    OrdenaI.AddItem "Vto.Letra"
    
    OrdenaI.ListIndex = 5
    

    FiltroI.Clear
    
    FiltroI.AddItem ""
    FiltroI.AddItem "Orden"
    FiltroI.AddItem "Planta"
    FiltroI.AddItem "Fecha"
    FiltroI.AddItem "Proveedor"
    FiltroI.AddItem "Carpeta"
    FiltroI.AddItem "DJai"
    FiltroI.AddItem "Origen"
    FiltroI.AddItem "Incoterms"
    FiltroI.AddItem "Transporte"
    FiltroI.AddItem "F.LLegada"
    FiltroI.AddItem "T.Pago"
    FiltroI.AddItem "Pago Despacho"
    FiltroI.AddItem "Pago Letra"
    FiltroI.AddItem "Vto. Letra"
    FiltroI.AddItem "M.Prima"
    
    FiltroI.ListIndex = 0

    ZProcesa = "N"

    FiltroII.Clear
    
    FiltroII.AddItem ""
    FiltroII.AddItem "Planta"
    FiltroII.AddItem "Fecha"
    FiltroII.AddItem "Proveedor"
    FiltroII.AddItem "Origen"
    FiltroII.AddItem "Incoterms"
    FiltroII.AddItem "Transporte"
    FiltroII.AddItem "F.LLegada"
    FiltroII.AddItem "T.Pago"
    FiltroII.AddItem "Pago Despacho"
    FiltroII.AddItem "Pago Letra"
    FiltroII.AddItem "Vto. Letra"
    
    FiltroII.ListIndex = 0


    FiltroIII.Clear
    
    FiltroIII.AddItem ""
    FiltroIII.AddItem "Planta"
    FiltroIII.AddItem "Fecha"
    FiltroIII.AddItem "Proveedor"
    FiltroIII.AddItem "Origen"
    FiltroIII.AddItem "Incoterms"
    FiltroIII.AddItem "Transporte"
    FiltroIII.AddItem "F.LLegada"
    FiltroIII.AddItem "T.Pago"
    FiltroIII.AddItem "Pago Despacho"
    FiltroIII.AddItem "Pago Letra"
    FiltroIII.AddItem "Vto. Letra"
    
    FiltroIII.ListIndex = 0
    
    Activas.Clear
    
    Activas.AddItem "Activas"
    Activas.AddItem "Cerradas"
    
    Activas.ListIndex = 0

    ZProcesa = ""

    Rem Call Proceso_Click
    
End Sub



Private Sub Muestra_DblClick()

    If Val(Muestra.TextMatrix(Muestra.Row, 1)) <> 0 Then

        WPasaOrden = Muestra.TextMatrix(Muestra.Row, 1)
        
        If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
            Select Case Muestra.TextMatrix(Muestra.Row, 2)
                Case "I"
                    WPasaEmpresa = 1
                Case "II"
                    WPasaEmpresa = 3
                Case "III"
                    WPasaEmpresa = 5
                Case "V"
                    WPasaEmpresa = 7
                Case "VI"
                    WPasaEmpresa = 10
                Case "VII"
                    WPasaEmpresa = 11
                Case Else
            End Select
                Else
            Select Case Muestra.TextMatrix(Muestra.Row, 2)
                Case "I"
                    WPasaEmpresa = 2
                Case "II"
                    WPasaEmpresa = 4
                Case "III"
                    WPasaEmpresa = 8
                Case "V"
                    WPasaEmpresa = 9
                Case Else
            End Select
        
        End If
        
        WPasaCarpeta = Muestra.TextMatrix(Muestra.Row, 6)
        
        PrgOrdenArchivos.Show
        PrgOrdenComplementoConsulta.Show
        PrgOrdenConsulta.Show
        
    End If
    
End Sub

Private Sub FiltroI_Click()
    
    ColumnaOpcion = FiltroI.ListIndex
    ZZFiltroOrdenI = FiltroI.ListIndex
    ZZTipoFiltro = 1
    
    Pantalla.Visible = False
    Ayuda.Visible = False
    PantaOrden.Visible = False
    PantaFecha.Visible = False
    PantaCarpeta.Visible = False
    
    Select Case ColumnaOpcion
        Case 0
            Call Proceso_Click
            
        Case 1
            FiltroOrden.Text = ""
            PantaOrden.Visible = True
            FiltroOrden.SetFocus
            
        Case 2
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Planta I"
            Pantalla.AddItem "Planta II"
            Pantalla.AddItem "Planta III"
            Pantalla.AddItem "Planta V"
            Pantalla.AddItem "Planta VI"
            Pantalla.AddItem "Planta VII"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 3
            FiltroFechaI.Text = "  /  /    "
            FiltroFechaII.Text = "  /  /    "
            PantaFecha.Visible = True
            FiltroFechaI.SetFocus
            
        Case 4, 6, 7
            ZZLugarFiltro = 0
            Erase ZZFiltro
            
            XEmpresa = Wempresa
            
            For CiclaEmpresa = 1 To 6
            
                Select Case CiclaEmpresa
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
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 5
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 6
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
            
            
                Pasa = 0
                Corte = ""
                
                Select Case ColumnaOpcion
                    Case 4
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Proveedor"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Proveedor"
                    Case 6
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Djai"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Djai"
                    Case 7
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Origen"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Origen"
                
                    Case Else
                End Select
                
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                Select Case ColumnaOpcion
                                    Case 4
                                        ZZCompara = rstOrden!Proveedor
                                    Case 6
                                        ZDJai = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
                                        ZZCompara = ZDJai
                                    Case 7
                                        ZZCompara = rstOrden!Origen
                                End Select
                                
                                If Trim(ZZCompara) <> "" Then
                            
                                    If Pasa = 0 Then
                                        Pasa = 1
                                        Corte = ZZCompara
                                    End If
                                    If Corte <> ZZCompara Then
                                    
                                        ZZEntra = "S"
                                        For CicloFiltro = 1 To ZZLugarFiltro
                                            If Corte = ZZFiltro(CicloFiltro, 1) Then
                                                ZZEntra = "N"
                                                Exit For
                                            End If
                                        Next CicloFiltro
                                        
                                        If ZZEntra = "S" Then
                                            ZZLugarFiltro = ZZLugarFiltro + 1
                                            ZZFiltro(ZZLugarFiltro, 1) = Corte
                                            ZZFiltro(ZZLugarFiltro, 2) = Corte
                                        End If
                                        Corte = ZZCompara
                                    End If
                                    
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    ZZLugarFiltro = ZZLugarFiltro + 1
                    ZZFiltro(ZZLugarFiltro, 1) = Corte
                    ZZFiltro(ZZLugarFiltro, 2) = Corte
                    rstOrden.Close
                End If
                
            Next CiclaEmpresa
            
            Call Conecta_Empresa
            
            If ColumnaOpcion = 4 Then
                For CicloFiltro = 1 To ZZLugarFiltro
                    ZZProveedor = ZZFiltro(CicloFiltro, 1)
                    ZSql = ""
                    ZSql = ZSql + "Select Proveedor.Proveedor, Proveedor.Nombre"
                    ZSql = ZSql + " FROM Proveedor"
                    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If rstProveedor.RecordCount > 0 Then
                        ZZFiltro(CicloFiltro, 2) = rstProveedor!Nombre
                        rstProveedor.Close
                    End If
                Next CicloFiltro
            End If
            
            ZZOrdenaFiltro = 2
            Call Ordena_Filtro
            
            Pantalla.Clear
            Pantalla.AddItem ""
            For CicloFiltro = 1 To ZZLugarFiltro
                Pantalla.AddItem ZZFiltro(CicloFiltro, 2)
            Next CicloFiltro
            
            Pantalla.Visible = True
            
        Case 5
            FiltroCarpeta.Text = ""
            PantaCarpeta.Visible = True
            FiltroCarpeta.SetFocus
            
        Case 8
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "FOB"
            Pantalla.AddItem "CIF"
            Pantalla.AddItem "CFR"
            Pantalla.AddItem "CPT"
            Pantalla.AddItem "EXW"
            Pantalla.AddItem "FCA"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 9
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Maritmo"
            Pantalla.AddItem "Terrestre"
            Pantalla.AddItem "Aereo"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 10
            FiltroLLegadaI.Text = "  /  /    "
            FiltroLLegadaII.Text = "  /  /    "
            PantaLlegada.Visible = True
            FiltroLLegadaI.SetFocus
            
        Case 11
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pago Anti."
            Pantalla.AddItem "A la vista"
            Pantalla.AddItem "Cta.Cte."
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 12
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 13
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 14
            FiltroVtoI.Text = "  /  /    "
            FiltroVtoII.Text = "  /  /    "
            PantaVto.Visible = True
            FiltroVtoI.SetFocus
            
        Case 15
            FiltroArticulo.Text = "  -   -   "
            PantaArticulo.Visible = True
            FiltroArticulo.SetFocus
            
            Dim IngresaItem As String
            PantallaII.Clear
            WIndice.Clear

            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        PantallaII.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
            AyudaII.Visible = True
            PantallaII.Visible = True
            AyudaII.Text = ""
            
        Case Else
        
    End Select
    
End Sub

Private Sub FiltroII_Click()
    
    ColumnaOpcionII = FiltroII.ListIndex
    ZZFiltroOrdenII = FiltroII.ListIndex
    ZZTipoFiltro = 2
    
    Pantalla.Visible = False
    Ayuda.Visible = False
    PantaOrden.Visible = False
    PantaFecha.Visible = False
    PantaCarpeta.Visible = False
    
    Select Case ColumnaOpcionII
        Case 0
            Call Proceso_Click
            
        Case 1
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Planta I"
            Pantalla.AddItem "Planta II"
            Pantalla.AddItem "Planta III"
            Pantalla.AddItem "Planta V"
            Pantalla.AddItem "Planta VI"
            Pantalla.AddItem "Planta VII"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 2
            FiltroFechaI.Text = "  /  /    "
            FiltroFechaII.Text = "  /  /    "
            PantaFecha.Visible = True
            FiltroFechaI.SetFocus
            
        Case 3, 4
            ZZLugarFiltro = 0
            Erase ZZFiltro
            
            XEmpresa = Wempresa
            
            For CiclaEmpresa = 1 To 6
            
                Select Case CiclaEmpresa
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
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 5
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 6
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
            
            
                Pasa = 0
                Corte = ""
                
                Select Case ColumnaOpcionII
                    Case 3
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Proveedor"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Proveedor"
                    Case 4
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Origen"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Origen"
                
                    Case Else
                End Select
                
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                Select Case ColumnaOpcionII
                                    Case 3
                                        ZZCompara = rstOrden!Proveedor
                                    Case 4
                                        ZZCompara = rstOrden!Origen
                                End Select
                                
                                If Trim(ZZCompara) <> "" Then
                            
                                    If Pasa = 0 Then
                                        Pasa = 1
                                        Corte = ZZCompara
                                    End If
                                    If Corte <> ZZCompara Then
                                    
                                        ZZEntra = "S"
                                        For CicloFiltro = 1 To ZZLugarFiltro
                                            If Corte = ZZFiltro(CicloFiltro, 1) Then
                                                ZZEntra = "N"
                                                Exit For
                                            End If
                                        Next CicloFiltro
                                        
                                        If ZZEntra = "S" Then
                                            ZZLugarFiltro = ZZLugarFiltro + 1
                                            ZZFiltro(ZZLugarFiltro, 1) = Corte
                                            ZZFiltro(ZZLugarFiltro, 2) = Corte
                                        End If
                                        Corte = ZZCompara
                                    End If
                                    
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    ZZLugarFiltro = ZZLugarFiltro + 1
                    ZZFiltro(ZZLugarFiltro, 1) = Corte
                    ZZFiltro(ZZLugarFiltro, 2) = Corte
                    rstOrden.Close
                End If
                
            Next CiclaEmpresa
            
            Call Conecta_Empresa
            
            If ColumnaOpcionII = 3 Then
                For CicloFiltro = 1 To ZZLugarFiltro
                    ZZProveedor = ZZFiltro(CicloFiltro, 1)
                    ZSql = ""
                    ZSql = ZSql + "Select Proveedor.Proveedor, Proveedor.Nombre"
                    ZSql = ZSql + " FROM Proveedor"
                    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If rstProveedor.RecordCount > 0 Then
                        ZZFiltro(CicloFiltro, 2) = rstProveedor!Nombre
                        rstProveedor.Close
                    End If
                Next CicloFiltro
            End If
            
            ZZOrdenaFiltro = 2
            Call Ordena_Filtro
            
            Pantalla.Clear
            Pantalla.AddItem ""
            For CicloFiltro = 1 To ZZLugarFiltro
                Pantalla.AddItem ZZFiltro(CicloFiltro, 2)
            Next CicloFiltro
            
            Pantalla.Visible = True
            
        Case 5
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "FOB"
            Pantalla.AddItem "CIF"
            Pantalla.AddItem "CFR"
            Pantalla.AddItem "CPT"
            Pantalla.AddItem "EXW"
            Pantalla.AddItem "FCA"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 6
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Maritmo"
            Pantalla.AddItem "Terrestre"
            Pantalla.AddItem "Aereo"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 7
            FiltroLLegadaI.Text = "  /  /    "
            FiltroLLegadaII.Text = "  /  /    "
            PantaLlegada.Visible = True
            FiltroLLegadaI.SetFocus
            
        Case 8
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pago Anti."
            Pantalla.AddItem "A la vista"
            Pantalla.AddItem "Cta.Cte."
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 9
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 10
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 11
            FiltroVtoI.Text = "  /  /    "
            FiltroVtoII.Text = "  /  /    "
            PantaVto.Visible = True
            FiltroVtoI.SetFocus
            
        Case 12
            FiltroArticulo.Text = "  -   -   "
            PantaArticulo.Visible = True
            FiltroArticulo.SetFocus
            
        Case Else
        
    End Select
    
End Sub















Private Sub FiltroIII_Click()
    
    ColumnaOpcionIII = FiltroIII.ListIndex
    ZZFiltroOrdenIII = FiltroIII.ListIndex
    ZZTipoFiltro = 3
    
    Pantalla.Visible = False
    Ayuda.Visible = False
    PantaOrden.Visible = False
    PantaFecha.Visible = False
    PantaCarpeta.Visible = False
    
    Select Case ColumnaOpcionIII
        Case 0
            Call Proceso_Click
            
        Case 1
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Planta I"
            Pantalla.AddItem "Planta II"
            Pantalla.AddItem "Planta III"
            Pantalla.AddItem "Planta V"
            Pantalla.AddItem "Planta VI"
            Pantalla.AddItem "Planta VII"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 2
            FiltroFechaI.Text = "  /  /    "
            FiltroFechaII.Text = "  /  /    "
            PantaFecha.Visible = True
            FiltroFechaI.SetFocus
            
        Case 3, 4
            ZZLugarFiltro = 0
            Erase ZZFiltro
            
            XEmpresa = Wempresa
            
            For CiclaEmpresa = 1 To 6
            
                Select Case CiclaEmpresa
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
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 5
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case 6
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    Case Else
                End Select
            
            
                Pasa = 0
                Corte = ""
                
                Select Case ColumnaOpcionIII
                    Case 3
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Proveedor"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Proveedor"
                    Case 4
                        ZSql = ""
                        ZSql = ZSql + "Select Orden.Origen"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Tipo = 1"
                        ZSql = ZSql + " and Orden.Recibida = 0"
                        ZSql = ZSql + " and Orden.Cantidad <> 0"
                        ZSql = ZSql + " Order by Orden.Origen"
                
                    Case Else
                End Select
                
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                Select Case ColumnaOpcionIII
                                    Case 3
                                        ZZCompara = rstOrden!Proveedor
                                    Case 4
                                        ZZCompara = rstOrden!Origen
                                End Select
                                
                                If Trim(ZZCompara) <> "" Then
                            
                                    If Pasa = 0 Then
                                        Pasa = 1
                                        Corte = ZZCompara
                                    End If
                                    If Corte <> ZZCompara Then
                                    
                                        ZZEntra = "S"
                                        For CicloFiltro = 1 To ZZLugarFiltro
                                            If Corte = ZZFiltro(CicloFiltro, 1) Then
                                                ZZEntra = "N"
                                                Exit For
                                            End If
                                        Next CicloFiltro
                                        
                                        If ZZEntra = "S" Then
                                            ZZLugarFiltro = ZZLugarFiltro + 1
                                            ZZFiltro(ZZLugarFiltro, 1) = Corte
                                            ZZFiltro(ZZLugarFiltro, 2) = Corte
                                        End If
                                        Corte = ZZCompara
                                    End If
                                    
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    ZZLugarFiltro = ZZLugarFiltro + 1
                    ZZFiltro(ZZLugarFiltro, 1) = Corte
                    ZZFiltro(ZZLugarFiltro, 2) = Corte
                    rstOrden.Close
                End If
                
            Next CiclaEmpresa
            
            Call Conecta_Empresa
            
            If ColumnaOpcionIII = 3 Then
                For CicloFiltro = 1 To ZZLugarFiltro
                    ZZProveedor = ZZFiltro(CicloFiltro, 1)
                    ZSql = ""
                    ZSql = ZSql + "Select Proveedor.Proveedor, Proveedor.Nombre"
                    ZSql = ZSql + " FROM Proveedor"
                    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If rstProveedor.RecordCount > 0 Then
                        ZZFiltro(CicloFiltro, 2) = rstProveedor!Nombre
                        rstProveedor.Close
                    End If
                Next CicloFiltro
            End If
            
            ZZOrdenaFiltro = 2
            Call Ordena_Filtro
            
            Pantalla.Clear
            Pantalla.AddItem ""
            For CicloFiltro = 1 To ZZLugarFiltro
                Pantalla.AddItem ZZFiltro(CicloFiltro, 2)
            Next CicloFiltro
            
            Pantalla.Visible = True
            
        Case 5
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "FOB"
            Pantalla.AddItem "CIF"
            Pantalla.AddItem "CFR"
            Pantalla.AddItem "CPT"
            Pantalla.AddItem "EXW"
            Pantalla.AddItem "FCA"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            ZZFiltro(4, 1) = "4"
            ZZFiltro(5, 1) = "5"
            ZZFiltro(6, 1) = "6"
            
            Pantalla.Visible = True
            
        Case 6
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Maritmo"
            Pantalla.AddItem "Terrestre"
            Pantalla.AddItem "Aereo"
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 7
            FiltroLLegadaI.Text = "  /  /    "
            FiltroLLegadaII.Text = "  /  /    "
            PantaLlegada.Visible = True
            FiltroLLegadaI.SetFocus
            
        Case 8
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pago Anti."
            Pantalla.AddItem "A la vista"
            Pantalla.AddItem "Cta.Cte."
            
            ZZFiltro(1, 1) = "1"
            ZZFiltro(2, 1) = "2"
            ZZFiltro(3, 1) = "3"
            
            Pantalla.Visible = True
            
        Case 9
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 10
            Pantalla.Clear
            Pantalla.AddItem ""
            Pantalla.AddItem "Pendiente"
            Pantalla.AddItem "Pagado"
            
            ZZFiltro(1, 1) = "0"
            ZZFiltro(2, 1) = "1"
            
            Pantalla.Visible = True
            
        Case 11
            FiltroVtoI.Text = "  /  /    "
            FiltroVtoII.Text = "  /  /    "
            PantaVto.Visible = True
            FiltroVtoI.SetFocus
            
        Case 12
            FiltroArticulo.Text = "  -   -   "
            PantaArticulo.Visible = True
            FiltroArticulo.SetFocus
            
        Case Else
        
    End Select
    
End Sub

















Private Sub FiltroOrden_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        ZZFiltroOrdenI = FiltroOrden.Text
    
        XEmpresa = Wempresa
        WLugar = 0
        Call Limpia_Vector
        
        For CiclaEmpresa = 1 To 6
        
            Select Case CiclaEmpresa
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
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case 5
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case 6
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case Else
            End Select
        
        
            ZSql = ""
            ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
            ZSql = ZSql + " FROM Orden, Proveedor"
            ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
            ZSql = ZSql + " and Orden.Orden = " + "'" + FiltroOrden.Text + "'"
            ZSql = ZSql + " and fechaord >=20140101"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                             
                If Activas.ListIndex <> 1 Then
                    If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                        ZEntra = "N"
                            Else
                        ZEntra = "S"
                    End If
                        Else
                    If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                        ZEntra = "S"
                            Else
                        ZEntra = "N"
                    End If
                End If
           
                If ZEntra = "S" Then
               
                    WLugar = WLugar + 1
                    ZDJai = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
                    
                    Muestra.TextMatrix(WLugar, 1) = rstOrden!Orden
                    Select Case CiclaEmpresa
                        Case 1
                            Muestra.TextMatrix(WLugar, 2) = "I"
                        Case 2
                            Muestra.TextMatrix(WLugar, 2) = "II"
                        Case 3
                            Muestra.TextMatrix(WLugar, 2) = "III"
                        Case 4
                            Muestra.TextMatrix(WLugar, 2) = "V"
                        Case 5
                            Muestra.TextMatrix(WLugar, 2) = "VI"
                        Case 6
                            Muestra.TextMatrix(WLugar, 2) = "VII"
                        Case Else
                    End Select
                    Muestra.TextMatrix(WLugar, 3) = Left$(rstOrden!Fecha, 5) + "/" + Mid$(rstOrden!Fecha, 9, 2)
                    Muestra.TextMatrix(WLugar, 4) = rstOrden!WProveedor
                    
                    Select Case rstOrden!Moneda
                        Case 0
                            Muestra.TextMatrix(WLugar, 5) = "U$S"
                        Case 1
                            Muestra.TextMatrix(WLugar, 5) = "$"
                        Case 2
                            Muestra.TextMatrix(WLugar, 5) = "Eur"
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 6) = rstOrden!Carpeta
                    Muestra.TextMatrix(WLugar, 7) = ZDJai
                    Muestra.TextMatrix(WLugar, 8) = rstOrden!Origen
                    
                    Select Case rstOrden!Leyenda
                        Case 1
                            Muestra.TextMatrix(WLugar, 9) = "FOB"
                        Case 2
                            Muestra.TextMatrix(WLugar, 9) = "CIF"
                        Case 3
                            Muestra.TextMatrix(WLugar, 9) = "CFR"
                        Case 4
                            Muestra.TextMatrix(WLugar, 9) = "CPT"
                        Case 5
                            Muestra.TextMatrix(WLugar, 9) = "EXW"
                        Case 6
                            Muestra.TextMatrix(WLugar, 9) = "FCA"
                        Case Else
                            Muestra.TextMatrix(WLugar, 9) = ""
                    End Select
                    
                    
                    Select Case rstOrden!TipoImpo
                        Case 1
                            Muestra.TextMatrix(WLugar, 10) = "Maritimo"
                        Case 2
                            Muestra.TextMatrix(WLugar, 10) = "Terrestre"
                        Case 3
                            Muestra.TextMatrix(WLugar, 10) = "Aereo"
                        Case Else
                            Muestra.TextMatrix(WLugar, 10) = ""
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 11) = rstOrden!FechaLlegada
                    
                    Select Case rstOrden!TipoPago
                        Case 1
                            Muestra.TextMatrix(WLugar, 12) = "Pago Anti."
                        Case 2
                            Muestra.TextMatrix(WLugar, 12) = "A la vista"
                        Case 3
                            Muestra.TextMatrix(WLugar, 12) = "Cta.Cte."
                        Case Else
                            Muestra.TextMatrix(WLugar, 12) = ""
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 13) = rstOrden!ImpoDespacho
                    Muestra.TextMatrix(WLugar, 13) = Pusing("###,###", Muestra.TextMatrix(WLugar, 13))
                    Select Case rstOrden!PagoDespacho
                        Case 0
                            Muestra.TextMatrix(WLugar, 14) = "Pendiente"
                        Case Else
                            Muestra.TextMatrix(WLugar, 14) = "Pagado"
                    End Select
                            
                    
                    
                    Muestra.TextMatrix(WLugar, 15) = rstOrden!ImpoLetra
                    Muestra.TextMatrix(WLugar, 15) = Pusing("###,###", Muestra.TextMatrix(WLugar, 15))
                    Select Case rstOrden!PagoLetra
                        Case 0
                            Muestra.TextMatrix(WLugar, 16) = "Pendiente"
                        Case Else
                            Muestra.TextMatrix(WLugar, 16) = "Pagado"
                    End Select
                    Muestra.TextMatrix(WLugar, 17) = rstOrden!VtoLetra
                    Muestra.TextMatrix(WLugar, 19) = IIf(IsNull(rstOrden!FechaEmbarque), "", rstOrden!FechaEmbarque)
                    Muestra.TextMatrix(WLugar, 20) = IIf(IsNull(rstOrden!FechaDJai), "", rstOrden!FechaDJai)
                    Muestra.TextMatrix(WLugar, 21) = rstOrden!Proveedor
                
                    rstOrden.Close
                    Exit For
                    
                        Else
                        
                    rstOrden.Close
                    
                End If
                
            End If
            
        Next CiclaEmpresa
        
        PantaOrden.Visible = False
        Call Conecta_Empresa
    
        Muestra.Col = 1
        Muestra.Row = 1
        Muestra.TopRow = 1
        
    End If
    If KeyAscii = 27 Then
        PantaOrden.Visible = False
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub FiltroFechaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FiltroFechaII.SetFocus
    End If
    If KeyAscii = 27 Then
        PantaFecha.Visible = False
    End If
End Sub


Private Sub FiltroFechaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WDesdeFecha = Right$(FiltroFechaI.Text, 4) + Mid$(FiltroFechaI.Text, 4, 2) + Left$(FiltroFechaI.Text, 2)
        WHastaFecha = Right$(FiltroFechaII.Text, 4) + Mid$(FiltroFechaII.Text, 4, 2) + Left$(FiltroFechaII.Text, 2)
        
        Select Case ZZTipoFiltro
            Case 1
                Seleccion = Right$(FiltroFechaI.Text, 4) + Mid$(FiltroFechaI.Text, 4, 2) + Left$(FiltroFechaI.Text, 2)
                SeleccionII = Right$(FiltroFechaII.Text, 4) + Mid$(FiltroFechaII.Text, 4, 2) + Left$(FiltroFechaII.Text, 2)
            Case 2
                SeleccionIII = Right$(FiltroFechaI.Text, 4) + Mid$(FiltroFechaI.Text, 4, 2) + Left$(FiltroFechaI.Text, 2)
                SeleccionIV = Right$(FiltroFechaII.Text, 4) + Mid$(FiltroFechaII.Text, 4, 2) + Left$(FiltroFechaII.Text, 2)
            Case 3
                SeleccionV = Right$(FiltroFechaI.Text, 4) + Mid$(FiltroFechaI.Text, 4, 2) + Left$(FiltroFechaI.Text, 2)
                SeleccionVI = Right$(FiltroFechaII.Text, 4) + Mid$(FiltroFechaII.Text, 4, 2) + Left$(FiltroFechaII.Text, 2)
            Case Else
        End Select
        
        PantaFecha.Visible = False
        Call Proceso_Click
        
    End If
    If KeyAscii = 27 Then
        PantaFecha.Visible = False
    End If
End Sub

Private Sub FiltroCarpeta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        XEmpresa = Wempresa
        WLugar = 0
        Call Limpia_Vector
        
        For CiclaEmpresa = 1 To 6
        
            Select Case CiclaEmpresa
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
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case 5
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case 6
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                Case Else
            End Select
        
        
            ZSql = ""
            ZSql = ZSql + "Select Orden.Tipo, Orden.Recibida, Orden.Cantidad, Orden.Clave, Orden.Orden, Orden.fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Condicion, Orden.Moneda, Orden.Carpeta, Orden.Djai, Orden.derechos, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.PagoDespacho, Orden.Fechallegada, Orden.FechaEmbarque, Orden.impodespacho, Orden.tipopago, Orden.vtodespacho, Orden.impoletra, Orden.vtoletra, Orden.pagoletra, Orden.fechadjai, Proveedor.Nombre as [WProveedor]"
            ZSql = ZSql + " FROM Orden, Proveedor"
            ZSql = ZSql + " Where Orden.Proveedor = Proveedor.Proveedor"
            ZSql = ZSql + " and Orden.Carpeta = " + "'" + FiltroCarpeta.Text + "'"
            ZSql = ZSql + " and fechaord >=20140101"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
            
                Rem by nan
                Rem   If rstOrden!Tipo = 1 And rstOrden!Recibida = 0 And rstOrden!Cantidad <> 0 Then
                If Activas.ListIndex <> 1 Then
                    If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                        ZEntra = "N"
                            Else
                        ZEntra = "S"
                    End If
                        Else
                    If rstOrden!Recibida <> 0 And rstOrden!PagoLetra = 1 Then
                        ZEntra = "S"
                            Else
                        ZEntra = "N"
                    End If
                End If
                
                    
                If ZEntra = "S" Then
            
                    WLugar = WLugar + 1
                    ZDJai = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
                    
                    Muestra.TextMatrix(WLugar, 1) = rstOrden!Orden
                    Select Case CiclaEmpresa
                        Case 1
                            Muestra.TextMatrix(WLugar, 2) = "I"
                        Case 2
                            Muestra.TextMatrix(WLugar, 2) = "II"
                        Case 3
                            Muestra.TextMatrix(WLugar, 2) = "III"
                        Case 4
                            Muestra.TextMatrix(WLugar, 2) = "V"
                        Case 5
                            Muestra.TextMatrix(WLugar, 2) = "VI"
                        Case 5
                            Muestra.TextMatrix(WLugar, 2) = "VII"
                        Case Else
                    End Select
                    Muestra.TextMatrix(WLugar, 3) = Left$(rstOrden!Fecha, 5) + "/" + Mid$(rstOrden!Fecha, 9, 2)
                    Muestra.TextMatrix(WLugar, 4) = rstOrden!WProveedor
                    
                    Select Case rstOrden!Moneda
                        Case 0
                            Muestra.TextMatrix(WLugar, 5) = "U$S"
                        Case 1
                            Muestra.TextMatrix(WLugar, 5) = "$"
                        Case 2
                            Muestra.TextMatrix(WLugar, 5) = "Eur"
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 6) = rstOrden!Carpeta
                    Muestra.TextMatrix(WLugar, 7) = ZDJai
                    Muestra.TextMatrix(WLugar, 8) = rstOrden!Origen
                    
                    Select Case rstOrden!Leyenda
                        Case 1
                            Muestra.TextMatrix(WLugar, 9) = "FOB"
                        Case 2
                            Muestra.TextMatrix(WLugar, 9) = "CIF"
                        Case 3
                            Muestra.TextMatrix(WLugar, 9) = "CFR"
                        Case 4
                            Muestra.TextMatrix(WLugar, 9) = "CPT"
                        Case 5
                            Muestra.TextMatrix(WLugar, 9) = "EXW"
                        Case 6
                            Muestra.TextMatrix(WLugar, 9) = "FCA"
                        Case Else
                            Muestra.TextMatrix(WLugar, 9) = ""
                    End Select
                    
                    
                    Select Case rstOrden!TipoImpo
                        Case 1
                            Muestra.TextMatrix(WLugar, 10) = "Maritimo"
                        Case 2
                            Muestra.TextMatrix(WLugar, 10) = "Terrestre"
                        Case 3
                            Muestra.TextMatrix(WLugar, 10) = "Aereo"
                        Case Else
                            Muestra.TextMatrix(WLugar, 10) = ""
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 11) = rstOrden!FechaLlegada
                    
                    Select Case rstOrden!TipoPago
                        Case 1
                            Muestra.TextMatrix(WLugar, 12) = "Pago Anti."
                        Case 2
                            Muestra.TextMatrix(WLugar, 12) = "A la vista"
                        Case 3
                            Muestra.TextMatrix(WLugar, 12) = "Cta.Cte."
                        Case Else
                            Muestra.TextMatrix(WLugar, 12) = ""
                    End Select
                    
                    Muestra.TextMatrix(WLugar, 13) = rstOrden!ImpoDespacho
                    Muestra.TextMatrix(WLugar, 13) = Pusing("###,###", Muestra.TextMatrix(WLugar, 13))
                    Select Case rstOrden!PagoDespacho
                        Case 0
                            Muestra.TextMatrix(WLugar, 14) = "Pendiente"
                        Case Else
                            Muestra.TextMatrix(WLugar, 14) = "Pagado"
                    End Select
                            
                    
                    
                    Muestra.TextMatrix(WLugar, 15) = rstOrden!ImpoLetra
                    Muestra.TextMatrix(WLugar, 15) = Pusing("###,###", Muestra.TextMatrix(WLugar, 15))
                    Select Case rstOrden!PagoLetra
                        Case 0
                            Muestra.TextMatrix(WLugar, 16) = "Pendiente"
                        Case Else
                            Muestra.TextMatrix(WLugar, 16) = "Pagado"
                    End Select
                    Muestra.TextMatrix(WLugar, 17) = rstOrden!VtoLetra
                    Muestra.TextMatrix(WLugar, 19) = IIf(IsNull(rstOrden!FechaEmbarque), "", rstOrden!FechaEmbarque)
                    Muestra.TextMatrix(WLugar, 20) = IIf(IsNull(rstOrden!FechaDJai), "", rstOrden!FechaDJai)
                    Muestra.TextMatrix(WLugar, 21) = rstOrden!Proveedor

                    rstOrden.Close
                    Exit For
                    
                        Else
                        
                    rstOrden.Close
                    
                End If
                
            End If
            
        Next CiclaEmpresa
    
        Muestra.Col = 1
        Muestra.Row = 1
        Muestra.TopRow = 1
        
        PantaCarpeta.Visible = False
        Call Conecta_Empresa
        
    End If
    If KeyAscii = 27 Then
        PantaCarpeta.Visible = False
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub FiltroLLegadaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FiltroLLegadaII.SetFocus
    End If
    If KeyAscii = 27 Then
        PantaLlegada.Visible = False
    End If
End Sub


Private Sub FiltroLLegadaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WDesdeFecha = Right$(FiltroLLegadaI.Text, 4) + Mid$(FiltroLLegadaI.Text, 4, 2) + Left$(FiltroLLegadaI.Text, 2)
        WHastaFecha = Right$(FiltroLLegadaII.Text, 4) + Mid$(FiltroLLegadaII.Text, 4, 2) + Left$(FiltroLLegadaII.Text, 2)
        
        Select Case ZZTipoFiltro
            Case 1
                Seleccion = Right$(FiltroLLegadaI.Text, 4) + Mid$(FiltroLLegadaI.Text, 4, 2) + Left$(FiltroLLegadaI.Text, 2)
                SeleccionII = Right$(FiltroLLegadaII.Text, 4) + Mid$(FiltroLLegadaII.Text, 4, 2) + Left$(FiltroLLegadaII.Text, 2)
            Case 2
                SeleccionIII = Right$(FiltroLLegadaI.Text, 4) + Mid$(FiltroLLegadaI.Text, 4, 2) + Left$(FiltroLLegadaI.Text, 2)
                SeleccionIV = Right$(FiltroLLegadaII.Text, 4) + Mid$(FiltroLLegadaII.Text, 4, 2) + Left$(FiltroLLegadaII.Text, 2)
            Case 3
                SeleccionV = Right$(FiltroLLegadaI.Text, 4) + Mid$(FiltroLLegadaI.Text, 4, 2) + Left$(FiltroLLegadaI.Text, 2)
                SeleccionVI = Right$(FiltroLLegadaII.Text, 4) + Mid$(FiltroLLegadaII.Text, 4, 2) + Left$(FiltroLLegadaII.Text, 2)
            Case Else
        End Select
        
        PantaLlegada.Visible = False
        Call Proceso_Click
        
    End If
    If KeyAscii = 27 Then
        PantaLlegada.Visible = False
    End If
End Sub





Private Sub FiltroVtoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FiltroVtoII.SetFocus
    End If
    If KeyAscii = 27 Then
        PantaVto.Visible = False
    End If
End Sub


Private Sub FiltroVtoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WDesdeFecha = Right$(FiltroVtoI.Text, 4) + Mid$(FiltroVtoI.Text, 4, 2) + Left$(FiltroVtoI.Text, 2)
        WHastaFecha = Right$(FiltroVtoII.Text, 4) + Mid$(FiltroVtoII.Text, 4, 2) + Left$(FiltroVtoII.Text, 2)
        
        Select Case ZZTipoFiltro
            Case 1
                Seleccion = Right$(FiltroVtoI.Text, 4) + Mid$(FiltroVtoI.Text, 4, 2) + Left$(FiltroVtoI.Text, 2)
                SeleccionII = Right$(FiltroVtoII.Text, 4) + Mid$(FiltroVtoII.Text, 4, 2) + Left$(FiltroVtoII.Text, 2)
            Case 2
                SeleccionIII = Right$(FiltroVtoI.Text, 4) + Mid$(FiltroVtoI.Text, 4, 2) + Left$(FiltroVtoI.Text, 2)
                SeleccionIV = Right$(FiltroVtoII.Text, 4) + Mid$(FiltroVtoII.Text, 4, 2) + Left$(FiltroVtoII.Text, 2)
            Case 3
                SeleccionV = Right$(FiltroVtoI.Text, 4) + Mid$(FiltroVtoI.Text, 4, 2) + Left$(FiltroVtoI.Text, 2)
                SeleccionVI = Right$(FiltroVtoII.Text, 4) + Mid$(FiltroVtoII.Text, 4, 2) + Left$(FiltroVtoII.Text, 2)
            Case Else
        End Select
        
        PantaVto.Visible = False
        Call Proceso_Click
        
    End If
    If KeyAscii = 27 Then
        PantaVto.Visible = False
    End If
End Sub


Private Sub FiltroArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        
        Select Case ZZTipoFiltro
            Case 1
                Seleccion = FiltroArticulo.Text
                SeleccionII = ""
            Case 2
                SeleccionIII = FiltroArticulo.Text
                SeleccionIV = ""
            Case 3
                SeleccionV = FiltroArticulo.Text
                SeleccionVI = ""
            Case Else
        End Select
        
        PantaArticulo.Visible = False
        Call Proceso_Click
        
    End If
    If KeyAscii = 27 Then
        PantaArticulo.Visible = False
    End If
End Sub















Private Sub Form_Activate()
    Rem **********se modifica tamano para gerencia
    If WOperador <> "17" Then
        Muestra.Height = 9375
        Muestra.Left = 120
        Muestra.Top = 960
        Muestra.Width = 15015
    End If
    
    
    
    ZZOrdena = 0
    Rem Call Proceso_Click
End Sub

Private Sub Muestra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 22
    Muestra.FixedRows = 1
    Muestra.Rows = 5000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Orden"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Pta"
                Muestra.ColWidth(Ciclo) = 500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Proveedor"
                Muestra.ColWidth(Ciclo) = 1800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "Mon"
                Muestra.ColWidth(Ciclo) = 500
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Carpeta"
                Muestra.ColWidth(Ciclo) = 700
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "DJai"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                Muestra.Text = "Origen"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                Muestra.Text = "Incoterms"
                Muestra.ColWidth(Ciclo) = 800
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                Muestra.Text = "Transporte"
                Muestra.ColWidth(Ciclo) = 900
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                Muestra.Text = "F.LLegada"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 12
                Muestra.Text = "T.Pago"
                Muestra.ColWidth(Ciclo) = 900
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 13
                Muestra.Text = "Despacho"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 14
                Muestra.Text = "Pago Des"
                Muestra.ColWidth(Ciclo) = 900
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 15
                Muestra.Text = "Letra"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 16
                Muestra.Text = "Pago Letra"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 17
                Muestra.Text = "Vto Letra"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 18
                Muestra.Text = "Pago Parcial"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 19
                Muestra.Text = "F.Embarque"
                Muestra.ColWidth(Ciclo) = 1100
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 20
                Muestra.Text = ""
                Muestra.ColWidth(Ciclo) = 10
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 21
                Muestra.Text = ""
                Muestra.ColWidth(Ciclo) = 10
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub OrdenaI_click()

    Dim Vector(5000, 20) As String
    Dim AuxiVector(20) As String

    Erase Vector

    For Ciclo = 1 To 5000
        If Val(Muestra.TextMatrix(Ciclo, 1)) = 0 Then
            ZZHasta = Ciclo - 1
            Exit For
                Else
            For ZColu = 1 To 17
                Vector(Ciclo, ZColu) = Muestra.TextMatrix(Ciclo, ZColu)
            Next ZColu
        End If
    Next Ciclo

    ZZTipoOrden = OrdenaI.ListIndex + 1

    For Ciclo = 1 To ZZHasta

        For Dada = Ciclo + 1 To ZZHasta


            ZZOrdenI = Vector(Ciclo, ZZTipoOrden)
            ZZOrdenII = Vector(Dada, ZZTipoOrden)
            
            If ZZTipoOrden = 3 Or ZZTipoOrden = 11 Or ZZTipoOrden = 17 Then
            
                WAno = Right$(ZZOrdenI, 4)
                WMes = Mid$(ZZOrdenI, 4, 2)
                WDia = Left$(ZZOrdenI, 2)
                ZZOrdenI = WAno + WMes + WDia
            
                WAno = Right$(ZZOrdenII, 4)
                WMes = Mid$(ZZOrdenII, 4, 2)
                WDia = Left$(ZZOrdenII, 2)
                ZZOrdenII = WAno + WMes + WDia
            
            End If

            If ZZOrdenI > ZZOrdenII Then

                Erase AuxiVector
                For ZColu = 1 To 17
                    AuxiVector(ZColu) = Vector(Ciclo, ZColu)
                Next ZColu
                
                For ZColu = 1 To 17
                    Vector(Ciclo, ZColu) = Vector(Dada, ZColu)
                Next ZColu
                
                For ZColu = 1 To 17
                    Vector(Dada, ZColu) = AuxiVector(ZColu)
                Next ZColu

            End If

        Next Dada

    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To ZZHasta
        For ZColu = 1 To 17
            Muestra.TextMatrix(Ciclo, ZColu) = Vector(Ciclo, ZColu)
        Next ZColu
    Next Ciclo
    
End Sub




Private Sub Ordena_Filtro()

    Dim Vector(5000, 20) As String
    Dim AuxiVector(20) As String

    Erase Vector

    For Ciclo = 1 To 5000
        If Trim(ZZFiltro(Ciclo, ZZOrdenaFiltro)) = "" Then
            ZZHasta = Ciclo - 1
            Exit For
                Else
            For ZColu = 1 To 2
                Vector(Ciclo, ZColu) = ZZFiltro(Ciclo, ZColu)
            Next ZColu
        End If
    Next Ciclo

    ZZTipoOrden = ZZOrdenaFiltro

    For Ciclo = 1 To ZZHasta

        For Dada = Ciclo + 1 To ZZHasta

            If Vector(Ciclo, ZZTipoOrden) > Vector(Dada, ZZTipoOrden) Then

                Erase AuxiVector
                For ZColu = 1 To 2
                    AuxiVector(ZColu) = Vector(Ciclo, ZColu)
                Next ZColu
                
                For ZColu = 1 To 2
                    Vector(Ciclo, ZColu) = Vector(Dada, ZColu)
                Next ZColu
                
                For ZColu = 1 To 2
                    Vector(Dada, ZColu) = AuxiVector(ZColu)
                Next ZColu

            End If

        Next Dada

    Next Ciclo
    
    Erase ZZFiltro
    
    For Ciclo = 1 To ZZHasta
        For ZColu = 1 To 2
            ZZFiltro(Ciclo, ZColu) = Vector(Ciclo, ZColu)
        Next ZColu
    Next Ciclo
    
End Sub




Private Sub pantalla_Click()
    If Pantalla.ListIndex <> 0 Then
        Rem Seleccion = Pantalla.Text
        
        Select Case ZZTipoFiltro
            Case 1
                Seleccion = ZZFiltro(Pantalla.ListIndex, 1)
            Case 2
                SeleccionIII = ZZFiltro(Pantalla.ListIndex, 1)
            Case Else
                SeleccionV = ZZFiltro(Pantalla.ListIndex, 1)
        End Select
        
            Else
            
        Seleccion = ""
        
        Select Case ZZTipoFiltro
            Case 1
                ColumnaOpcion = 0
            Case 2
                ColumnaOpcionII = 0
            Case Else
                ColumnaOpcionIII = 0
        End Select
        
    End If
    Pantalla.Visible = False
    Ayuda.Visible = False
    Call Proceso_Click
End Sub




Private Sub aYUDAII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        PantallaII.Clear
        WIndice.Clear
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + AyudaII.Text + "%" + "'"
        ZSql = ZSql + " Order by Articulo.Codigo"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then

            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        PantallaII.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        
                        .MoveNext
            
                            Else
                
                        Exit Do
        
                    End If
                Loop
            End With

            rstArticulo.Close
        End If
    
    End If
End Sub

Private Sub pantallaii_Click()

    PantallaII.Visible = False
    
    Indice = PantallaII.ListIndex
    FiltroArticulo = WIndice.List(Indice)
    Call FiltroArticulo_KeyPress(13)
    
End Sub

