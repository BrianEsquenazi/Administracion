VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaSac 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de SAC por Centro"
   ClientHeight    =   8385
   ClientLeft      =   2160
   ClientTop       =   525
   ClientWidth     =   8505
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   8505
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   960
      TabIndex        =   41
      Top             =   2640
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Comentario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   44
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox IngresoCausa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   43
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox IngresoNoCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame PantaOpcion 
      Height          =   1455
      Left            =   3360
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton CerrarOpciones 
         Caption         =   "Confirmar"
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
         Left            =   1200
         TabIndex        =   35
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox Opcion9 
         Caption         =   "Opcion9"
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
         Left            =   720
         TabIndex        =   34
         Top             =   3720
         Width           =   3200
      End
      Begin VB.CheckBox Opcion10 
         Caption         =   "Opcion10"
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
         Left            =   720
         TabIndex        =   33
         Top             =   4080
         Width           =   3200
      End
      Begin VB.CheckBox Opcion7 
         Caption         =   "Opcion7"
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
         Left            =   720
         TabIndex        =   32
         Top             =   3000
         Width           =   3200
      End
      Begin VB.CheckBox Opcion8 
         Caption         =   "Opcion8"
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
         Left            =   720
         TabIndex        =   31
         Top             =   3360
         Width           =   3200
      End
      Begin VB.CheckBox Opcion5 
         Caption         =   "Opcion5"
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
         Left            =   720
         TabIndex        =   30
         Top             =   2280
         Width           =   3200
      End
      Begin VB.CheckBox Opcion6 
         Caption         =   "Opcion6"
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
         Left            =   720
         TabIndex        =   29
         Top             =   2640
         Width           =   3200
      End
      Begin VB.CheckBox Opcion3 
         Caption         =   "Opcion3"
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
         Left            =   720
         TabIndex        =   28
         Top             =   1560
         Width           =   3200
      End
      Begin VB.CheckBox Opcion4 
         Caption         =   "Opcion4"
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
         Left            =   720
         TabIndex        =   27
         Top             =   1920
         Width           =   3200
      End
      Begin VB.CheckBox Opcion1 
         Caption         =   "Opcion1"
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
         Left            =   720
         TabIndex        =   26
         Top             =   840
         Width           =   3200
      End
      Begin VB.CheckBox Opcion2 
         Caption         =   "Opcion2"
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
         Left            =   720
         TabIndex        =   25
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label TituloOpcion 
         Alignment       =   2  'Center
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton listadito 
      Caption         =   "Listado basico"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   7095
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6120
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6000
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Verorigen 
         Caption         =   "Parametros"
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
         Left            =   1680
         TabIndex        =   40
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Verestado 
         Caption         =   "Parametros"
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
         Left            =   1680
         TabIndex        =   39
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Vertipo 
         Caption         =   "Parametros"
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
         Left            =   1680
         TabIndex        =   38
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox AnoII 
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
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   37
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox TipoListado 
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
         TabIndex        =   21
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox Responsable 
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
         TabIndex        =   16
         Text            =   " "
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox HastaCentro 
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
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   9
         Text            =   " "
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Ano 
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
         TabIndex        =   8
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox DesdeCentro 
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
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   3360
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
         Height          =   735
         Left            =   5520
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
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
         Height          =   735
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label DesResponsable 
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
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Centro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\listadosac.rpt"
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   600
      TabIndex        =   14
      Top             =   4800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin Crystal.CrystalReport listado2 
      Left            =   7920
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\accionesnan.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgListaSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim XParam As String
Dim ZZLugar As Integer

Dim ZZOpcion As Integer
Dim ZZClave(100) As Integer
Dim ZZClaveI(100) As Integer
Dim ZZClaveII(100) As Integer
Dim ZZClaveIII(100) As Integer

Dim ZZPasaTipo As String
Dim ZZPasaAno As String
Dim ZZPasaNumero As String

Dim ZZCentro As String
Dim ZZFecha As String
Dim ZZOrigen As String
Dim ZZEstado As String
Dim ZZResponsableEmisor As String
Dim ZZResponsableDestino As String
Dim ZZReferencia As String
Dim ZZTitulo As String

Dim ZZAccion11 As String
Dim ZZAccion12 As String
Dim ZZAccion21 As String
Dim ZZAccion22 As String
Dim ZZAccion31 As String
Dim ZZAccion32 As String
Dim ZZAccion41 As String
Dim ZZAccion42 As String
Dim ZZAccion51 As String
Dim ZZAccion52 As String
Dim ZZAccion61 As String
Dim ZZAccion62 As String

Dim ZZResponsable1 As String
Dim ZZResponsable2 As String
Dim ZZResponsable3 As String
Dim ZZResponsable4 As String
Dim ZZResponsable5 As String
Dim ZZResponsable6 As String

Dim ZZPlazo1 As String
Dim ZZPlazo2 As String
Dim ZZPlazo3 As String
Dim ZZPlazo4 As String
Dim ZZPlazo5 As String
Dim ZZPlazo6 As String

Dim ZZResponsable11 As String
Dim ZZResponsable12 As String
Dim ZZResponsable13 As String
Dim ZZResponsable14 As String
Dim ZZResponsable15 As String
Dim ZZResponsable16 As String

Dim ZZResponsable21 As String
Dim ZZResponsable22 As String
Dim ZZResponsable23 As String
Dim ZZResponsable24 As String
Dim ZZResponsable25 As String
Dim ZZResponsable26 As String

Dim ZZResponsable31 As String
Dim ZZResponsable32 As String
Dim ZZResponsable33 As String
Dim ZZResponsable34 As String
Dim ZZResponsable35 As String
Dim ZZResponsable36 As String
            
Dim ZZFecha1 As String
Dim ZZFecha2 As String
Dim ZZFecha3 As String
Dim ZZFecha4 As String
Dim ZZFecha5 As String
Dim ZZFecha6 As String
            
Dim ZZFecha21 As String
Dim ZZFecha22 As String
Dim ZZFecha23 As String
Dim ZZFecha24 As String
Dim ZZFecha25 As String
Dim ZZFecha26 As String
            
Dim ZZFecha31 As String
Dim ZZFecha32 As String
Dim ZZFecha33 As String
Dim ZZFecha34 As String
Dim ZZFecha35 As String
Dim ZZFecha36 As String
            
Dim ZZComentario11 As String
Dim ZZComentario12 As String
Dim ZZComentario21 As String
Dim ZZComentario22 As String
Dim ZZComentario31 As String
Dim ZZComentario32 As String
Dim ZZComentario41 As String
Dim ZZComentario42 As String
Dim ZZComentario51 As String
Dim ZZComentario52 As String
Dim ZZComentario61 As String
Dim ZZComentario62 As String
            
Dim ZZComentario211 As String
Dim ZZComentario212 As String
Dim ZZComentario221 As String
Dim ZZComentario222 As String
Dim ZZComentario231 As String
Dim ZZComentario232 As String
Dim ZZComentario241 As String
Dim ZZComentario242 As String
Dim ZZComentario251 As String
Dim ZZComentario252 As String
Dim ZZComentario261 As String
Dim ZZComentario262 As String



Dim ZZEstado11 As String
Dim ZZEstado12 As String
Dim ZZEstado13 As String
Dim ZZEstado14 As String
Dim ZZEstado15 As String
Dim ZZEstado16 As String

Dim ZZEstado1 As String
Dim ZZEstado2 As String
Dim ZZEstado3 As String
Dim ZZEstado4 As String
Dim ZZEstado5 As String
Dim ZZEstado6 As String

Dim ZZEstado31 As String
Dim ZZEstado32 As String
Dim ZZEstado33 As String
Dim ZZEstado34 As String
Dim ZZEstado35 As String
Dim ZZEstado36 As String

Dim ZZDesTipo As String
Dim ZZDesCentro As String
Dim ZZDesResponsableEmisor As String
Dim ZZDesResponsableDestino As String
        
Dim ZZDesResponsable1 As String
Dim ZZDesResponsable2 As String
Dim ZZDesResponsable3 As String
Dim ZZDesResponsable4 As String
Dim ZZDesResponsable5 As String
Dim ZZDesResponsable6 As String
        
Dim ZZDesResponsable11 As String
Dim ZZDesResponsable12 As String
Dim ZZDesResponsable13 As String
Dim ZZDesResponsable14 As String
Dim ZZDesResponsable15 As String
Dim ZZDesResponsable16 As String

Dim ZZDesResponsable21 As String
Dim ZZDesResponsable22 As String
Dim ZZDesResponsable23 As String
Dim ZZDesResponsable24 As String
Dim ZZDesResponsable25 As String
Dim ZZDesResponsable26 As String

Dim ZZDesResponsable31 As String
Dim ZZDesResponsable32 As String
Dim ZZDesResponsable33 As String
Dim ZZDesResponsable34 As String
Dim ZZDesResponsable35 As String
Dim ZZDesResponsable36 As String

Dim ZVector(1000, 7) As String
Dim ZLugar As Integer

Private Sub DesdeCentro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCentro.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeCentro.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaCentro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaCentro.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) <> 0 Then
            AnoII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub AnoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) <> 0 Then
            Responsable.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        AnoII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub listadito_Click()

    If Ano.Text = "" Then
        m$ = "Se debe ingresar el año "
        G% = MsgBox(m$, 0, "AÑO")
        Exit Sub
    End If

    salidita = 1
    
    Call Acepta_Click

    listado2.WindowTitle = "Listado de SAC"
    listado2.WindowTop = 0
    listado2.WindowLeft = 0
    listado2.WindowWidth = Screen.Width
    listado2.WindowHeight = Screen.Height
    
    listado2.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
                    
    listado2.Connect = Connect()
    listado2.ReportFileName = "accionestotales.rpt"
    listado2.Action = 1


    DesdeCentro.Text = ""
    HastaCentro.Text = ""
             
    Ano.Text = ""
    AnoII.Text = ""
    
    salidita = ""
    
End Sub

Private Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                DesdeCentro.SetFocus
            End If
                Else
            Responsable.Text = ""
            DesResponsable.Caption = ""
            DesdeCentro.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
        DesResponsable.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acepta_Click()

    If Val(AnoII.Text) = 0 Then
        AnoII.Text = Ano.Text
    End If
    
    WDesde = Ano.Text + "0101"
    WHasta = AnoII.Text + "1231"
    
    Rem ZCantidad = 0
    Rem ZCantidadAcciones = 0
    Rem ZCantidadImple = 0
    Rem ZCantidadCerradas = 0
    
    Rem Sql1 = "Select Clave, Ordfecha, Plazo1, Plazo2, Plazo3, Plazo4, Plazo5, Plazo6, Fecha1, Fecha2, Fecha3, Fecha4, Fecha5, Fecha6, Fecha21, Fecha22, Fecha23, Fecha24, Fecha25, Fecha26, Responsable1, Responsable2, Responsable3, Responsable4, Responsable5, Responsable6"
    Rem Sql2 = " FROM CargaSac"
    Rem Sql3 = " Order by CargaSac.Clave"
    Rem spCargaSac = Sql1 + Sql2 + Sql3
    Rem Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstCargaSac.RecordCount > 0 Then
    Rem     With rstCargaSac
    Rem         .MoveFirst
    Rem         Do
    Rem             If .EOF = False Then
    Rem                 If WDesde <= rstCargaSac!ordfecha And WHasta >= rstCargaSac!ordfecha Then
    Rem
    Rem                     ZCantidad = ZCantidad + 1
    Rem
    Rem                     If rstCargaSac!Plazo1 <> "  /  /    " Then
    Rem
    Rem                         ZCantidadAcciones = ZCantidadAcciones + 1
    Rem
    Rem                         If rstCargaSac!Responsable1 <> 0 Then
    Rem
    Rem                             ZCantidadImple = ZCantidadImple + 1
    Rem
    Rem                             If rstCargaSac!Fecha21 <> "  /  /    " Then
    Rem                                 ZCantidadCerradas = ZCantidadCerradas + 1
    Rem                             End If
    Rem
    Rem                         End If
    Rem
    Rem                     End If
    Rem
    Rem                 End If
    Rem                 .MoveNext
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem     rstCargaSac.Close
    Rem End If
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE CargaSac SET "
    Rem ZSql = ZSql + " Cantidad = " + "'" + Str$(ZCantidad) + "',"
    Rem ZSql = ZSql + " CantidadAcciones = " + "'" + Str$(ZCantidadAcciones) + "',"
    Rem ZSql = ZSql + " CantidadImplementadas = " + "'" + Str$(ZCantidadImple) + "',"
    Rem ZSql = ZSql + " CantidadCerradas = " + "'" + Str$(ZCantidadCerradas) + "',"
    Rem ZSql = ZSql + " Porce1 = " + "'" + "0" + "',"
    Rem ZSql = ZSql + " Porce2 = " + "'" + "0" + "'"
    Rem ZSql = ZSql + " Where Centro = " + "'" + Centro.Text + "'"
    Rem spCargaSac = ZSql
    Rem Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    Rem Select Case Tipo.ListIndex
    Rem     Case 0
    Rem         DesdeEstado = "0"
    Rem         HastaEstado = "6"
    Rem     Case 1
    Rem         DesdeEstado = "1"
    Rem         HastaEstado = "1"
    Rem     Case 2
    Rem         DesdeEstado = "2"
    Rem         HastaEstado = "2"
    Rem     Case 3
    Rem         DesdeEstado = "3"
    Rem         HastaEstado = "3"
    Rem     Case 4
    Rem         DesdeEstado = "4"
    Rem         HastaEstado = "4"
    Rem     Case 5
    Rem         DesdeEstado = "5"
    Rem         HastaEstado = "5"
    Rem     Case 6
    Rem         DesdeEstado = "6"
    Rem         HastaEstado = "6"
    Rem     Case 7
    Rem         DesdeEstado = "0"
    Rem         HastaEstado = "5"
    Rem     Case Else
    Rem End Select
    
    
    
    
    
    Rem Select Case Origen.ListIndex
    Rem     Case 0
    Rem         DesdeOrigen = "0"
    Rem         HastaOrigen = "5"
    Rem     Case 1
    Rem         DesdeOrigen = "1"
    Rem         HastaOrigen = "1"
    Rem     Case 2
    Rem         DesdeOrigen = "2"
    Rem         HastaOrigen = "2"
    Rem     Case 3
    Rem         DesdeOrigen = "3"
    Rem         HastaOrigen = "3"
    Rem     Case 4
    Rem         DesdeOrigen = "4"
    Rem         HastaOrigen = "4"
    Rem     Case 5
    Rem         DesdeOrigen = "5"
    Rem         HastaOrigen = "5"
    Rem     Case Else
    Rem End Select
    
    
    
    
    Rem by nan para la salida del listadito completo
    If salidita <> "" Then
          
        DesdeCentro.Text = "1"
        HastaCentro.Text = "99"
          
        WDesde = Ano.Text + "0101"
        WHasta = AnoII.Text + "1231"
        Rem WHasta = Right$(Date, 4) + "1231"
        TipoListado.ListIndex = 1
    
    End If
            
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaSac SET "
    ZSql = ZSql + " Marca = " + "'" + "N" + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase ZVector
    ZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Centro >= " + "'" + DesdeCentro.Text + "'"
    ZSql = ZSql + " and CargaSac.Centro <= " + "'" + HastaCentro.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano >= " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano <= " + "'" + AnoII.Text + "'"
    ZSql = ZSql + " Order by CargaSac.Clave"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
        .MoveFirst
        Do
            If .EOF = False Then
            
                ZEntraI = "N"
                ZEntraII = "N"
                ZEntraIII = "N"
                
                For Ciclo = 1 To 100
                    If ZZClaveI(Ciclo) = rstCargaSac!Tipo Then
                        ZEntraI = "S"
                    End If
                Next Ciclo
                
                For Ciclo = 1 To 100
                    If ZZClaveII(Ciclo) = rstCargaSac!Estado Then
                        ZEntraII = "S"
                    End If
                Next Ciclo
                
                For Ciclo = 1 To 100
                    If ZZClaveIII(Ciclo) = rstCargaSac!Origen Then
                        ZEntraIII = "S"
                    End If
                Next Ciclo
            
                If ZEntraI = "S" And ZEntraII = "S" And ZEntraIII = "S" Then
                
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = rstCargaSac!Clave
                    ZVector(ZLugar, 2) = rstCargaSac!Tipo
                    ZVector(ZLugar, 3) = rstCargaSac!Ano
                    ZVector(ZLugar, 4) = rstCargaSac!Numero
                    ZVector(ZLugar, 5) = rstCargaSac!Centro
                    ZVector(ZLugar, 6) = rstCargaSac!ResponsableEmisor
                    ZVector(ZLugar, 7) = rstCargaSac!ResponsableDestino
                    
                End If
                
                .MoveNext
                
                        Else
                        
                    Exit Do
                End If
            Loop
        End With
        rstCargaSac.Close
    End If

    For Ciclo = 1 To ZLugar

        ZZZClave = ZVector(Ciclo, 1)
        ZZTipo = ZVector(Ciclo, 2)
        ZZAno = ZVector(Ciclo, 3)
        ZZNumero = ZVector(Ciclo, 4)
        ZZCentro = ZVector(Ciclo, 5)
        ZZResponsableEmisor = Val(ZVector(Ciclo, 6))
        ZZResponsableDestino = Val(ZVector(Ciclo, 7))
        ZZResponsablecentro = 0
        ZZResponsableAccion1 = 0
        ZZResponsableAccion2 = 0
        ZZResponsableAccion3 = 0
        ZZResponsableAccion4 = 0
        ZZResponsableAccion5 = 0
        ZZResponsableAccion6 = 0
        ZZEstado1 = 0
        ZZEstado2 = 0
        ZZEstado3 = 0
        ZZEstado4 = 0
        ZZEstado5 = 0
        ZZEstado6 = 0
        ZZFecha1 = ""
        ZZFecha2 = ""
        ZZFecha3 = ""
        ZZFecha4 = ""
        ZZFecha5 = ""
        ZZFecha6 = ""
        ZZObserva1 = ""
        ZZObserva2 = ""
        ZZObserva3 = ""
        ZZObserva4 = ""
        ZZObserva5 = ""
        ZZObserva6 = ""

        Sql1 = "Select *"
        Sql2 = " FROM CentroSac"
        Sql3 = " Where CentroSac.Codigo = " + "'" + ZZCentro + "'"
        spCentroSac = Sql1 + Sql2 + Sql3
        Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCentroSac.RecordCount > 0 Then
            ZZResponsablecentro = rstCentroSac!Responsable
            rstCentroSac.Close
        End If

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacII"
        ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + ZZTipo + "'"
        ZSql = ZSql + " and CargaSacII.Ano = " + "'" + ZZAno + "'"
        ZSql = ZSql + " and CargaSacII.Numero = " + "'" + ZZNumero + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacII.RecordCount > 0 Then
            ZZResponsableAccion1 = rstCargaSacII!Responsable1
            ZZResponsableAccion2 = rstCargaSacII!Responsable2
            ZZResponsableAccion3 = rstCargaSacII!Responsable3
            ZZResponsableAccion4 = rstCargaSacII!Responsable4
            ZZResponsableAccion5 = rstCargaSacII!Responsable5
            ZZResponsableAccion6 = rstCargaSacII!Responsable6
            rstCargaSacII.Close
            
                Else
                
            ZZAccion11 = ""
            ZZAccion12 = ""
            ZZAccion21 = ""
            ZZAccion22 = ""
            ZZAccion31 = ""
            ZZAccion32 = ""
            ZZAccion41 = ""
            ZZAccion42 = ""
            ZZAccion51 = ""
            ZZAccion52 = ""
            ZZAccion61 = ""
            ZZAccion62 = ""
            
            ZZResponsable1 = ""
            ZZResponsable2 = ""
            ZZResponsable3 = ""
            ZZResponsable4 = ""
            ZZResponsable5 = ""
            ZZResponsable6 = ""
            
            ZZPlazo1 = "  /  /    "
            ZZPlazo2 = "  /  /    "
            ZZPlazo3 = "  /  /    "
            ZZPlazo4 = "  /  /    "
            ZZPlazo5 = "  /  /    "
            ZZPlazo6 = "  /  /    "
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSacII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Accion11 ,"
            ZSql = ZSql + "Accion12 ,"
            ZSql = ZSql + "Accion21 ,"
            ZSql = ZSql + "Accion22 ,"
            ZSql = ZSql + "Accion31 ,"
            ZSql = ZSql + "Accion32 ,"
            ZSql = ZSql + "Accion41 ,"
            ZSql = ZSql + "Accion42 ,"
            ZSql = ZSql + "Accion51 ,"
            ZSql = ZSql + "Accion52 ,"
            ZSql = ZSql + "Accion61 ,"
            ZSql = ZSql + "Accion62 ,"
            ZSql = ZSql + "Responsable1 ,"
            ZSql = ZSql + "Responsable2 ,"
            ZSql = ZSql + "Responsable3 ,"
            ZSql = ZSql + "Responsable4 ,"
            ZSql = ZSql + "Responsable5 ,"
            ZSql = ZSql + "Responsable6 ,"
            ZSql = ZSql + "Plazo1 ,"
            ZSql = ZSql + "Plazo2 ,"
            ZSql = ZSql + "Plazo3 ,"
            ZSql = ZSql + "Plazo4 ,"
            ZSql = ZSql + "Plazo5 ,"
            ZSql = ZSql + "Plazo6 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZAno + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZAccion11 + "',"
            ZSql = ZSql + "'" + ZZAccion12 + "',"
            ZSql = ZSql + "'" + ZZAccion21 + "',"
            ZSql = ZSql + "'" + ZZAccion22 + "',"
            ZSql = ZSql + "'" + ZZAccion31 + "',"
            ZSql = ZSql + "'" + ZZAccion32 + "',"
            ZSql = ZSql + "'" + ZZAccion41 + "',"
            ZSql = ZSql + "'" + ZZAccion42 + "',"
            ZSql = ZSql + "'" + ZZAccion51 + "',"
            ZSql = ZSql + "'" + ZZAccion52 + "',"
            ZSql = ZSql + "'" + ZZAccion61 + "',"
            ZSql = ZSql + "'" + ZZAccion62 + "',"
            ZSql = ZSql + "'" + ZZResponsable1 + "',"
            ZSql = ZSql + "'" + ZZResponsable2 + "',"
            ZSql = ZSql + "'" + ZZResponsable3 + "',"
            ZSql = ZSql + "'" + ZZResponsable4 + "',"
            ZSql = ZSql + "'" + ZZResponsable5 + "',"
            ZSql = ZSql + "'" + ZZResponsable6 + "',"
            ZSql = ZSql + "'" + ZZPlazo1 + "',"
            ZSql = ZSql + "'" + ZZPlazo2 + "',"
            ZSql = ZSql + "'" + ZZPlazo3 + "',"
            ZSql = ZSql + "'" + ZZPlazo4 + "',"
            ZSql = ZSql + "'" + ZZPlazo5 + "',"
            ZSql = ZSql + "'" + ZZPlazo6 + "')"
            
            spCargaSacII = ZSql
            Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIII"
        ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + ZZTipo + "'"
        ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + ZZAno + "'"
        ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + ZZNumero + "'"
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIII.RecordCount > 0 Then
            ZZEstado1 = rstCargaSacIII!Estado1
            ZZEstado2 = rstCargaSacIII!Estado2
            ZZEstado3 = rstCargaSacIII!Estado3
            ZZEstado4 = rstCargaSacIII!Estado4
            ZZEstado5 = rstCargaSacIII!Estado5
            ZZEstado6 = rstCargaSacIII!Estado6
            ZZFecha1 = rstCargaSacIII!Fecha1
            ZZFecha2 = rstCargaSacIII!Fecha2
            ZZFecha3 = rstCargaSacIII!Fecha3
            ZZFecha4 = rstCargaSacIII!Fecha4
            ZZFecha5 = rstCargaSacIII!Fecha5
            ZZFecha6 = rstCargaSacIII!Fecha6
            ZZObserva1 = rstCargaSacIII!Comentario11 + " " + rstCargaSacIII!Comentario12
            ZZObserva2 = rstCargaSacIII!Comentario21 + " " + rstCargaSacIII!Comentario22
            ZZObserva3 = rstCargaSacIII!Comentario31 + " " + rstCargaSacIII!Comentario32
            ZZObserva4 = rstCargaSacIII!Comentario41 + " " + rstCargaSacIII!Comentario42
            ZZObserva5 = rstCargaSacIII!Comentario51 + " " + rstCargaSacIII!Comentario52
            ZZObserva6 = rstCargaSacIII!Comentario61 + " " + rstCargaSacIII!Comentario62
            rstCargaSacIII.Close
        End If
    
        Entra = "N"
    
        If Val(Responsable.Text) = ZZResponsableEmisor Then
            Entra = "S"
        End If
        
        If Val(Responsable.Text) = ZZResponsableDestino Then
            Entra = "S"
        End If
    
        If Val(Responsable.Text) = ZZResponsablecentro Then
            Entra = "S"
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion1 Then
            If ZZEstado1 = 0 Then
                Entra = "S"
            End If
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion2 Then
            If ZZEstado2 = 0 Then
                Entra = "S"
            End If
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion3 Then
            If ZZEstado3 = 0 Then
                Entra = "S"
            End If
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion4 Then
            If ZZEstado4 = 0 Then
                Entra = "S"
            End If
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion5 Then
            If ZZEstado5 = 0 Then
                Entra = "S"
            End If
        End If
    
        If Val(Responsable.Text) = ZZResponsableAccion6 Then
            If ZZEstado6 = 0 Then
                Entra = "S"
            End If
        End If
        
        If Val(Responsable.Text) = 0 Then
            Entra = "S"
        End If
        
        If Entra = "S" Then
        
            ZZDesResponsableAccion1 = ""
            ZZDesResponsableAccion2 = ""
            ZZDesResponsableAccion3 = ""
            ZZDesResponsableAccion4 = ""
            ZZDesResponsableAccion5 = ""
            ZZDesResponsableAccion6 = ""
        
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion1) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion1 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion2) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion2 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion3) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion3 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion4) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion4 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion5) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion5 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsableAccion6) + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                ZZDesResponsableAccion6 = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
            End If
            
            If ZZEstado1 = 1 Then
                ZZImpreEstado1 = "Implementado"
                    Else
                ZZImpreEstado1 = ""
            End If
            
            If ZZEstado2 = 1 Then
                ZZImpreEstado2 = "Implementado"
                    Else
                ZZImpreEstado2 = ""
            End If
            
            If ZZEstado3 = 1 Then
                ZZImpreEstado3 = "Implementado"
                    Else
                ZZImpreEstado3 = ""
            End If
            
            If ZZEstado4 = 1 Then
                ZZImpreEstado4 = "Implementado"
                    Else
                ZZImpreEstado4 = ""
            End If
            
            If ZZEstado5 = 1 Then
                ZZImpreEstado5 = "Implementado"
                    Else
                ZZImpreEstado5 = ""
            End If
            
            If ZZEstado6 = 1 Then
                ZZImpreEstado6 = "Implementado"
                    Else
                ZZImpreEstado6 = ""
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSac SET "
            ZSql = ZSql + " Marca = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZZClave + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSacII SET "
            ZSql = ZSql + " Estado1 = " + "'" + Left$(ZZImpreEstado1, 10) + "',"
            ZSql = ZSql + " Estado2 = " + "'" + Left$(ZZImpreEstado2, 10) + "',"
            ZSql = ZSql + " Estado3 = " + "'" + Left$(ZZImpreEstado3, 10) + "',"
            ZSql = ZSql + " Estado4 = " + "'" + Left$(ZZImpreEstado4, 10) + "',"
            ZSql = ZSql + " Estado5 = " + "'" + Left$(ZZImpreEstado5, 10) + "',"
            ZSql = ZSql + " Estado6 = " + "'" + Left$(ZZImpreEstado6, 10) + "',"
            ZSql = ZSql + " FImplementa1 = " + "'" + ZZFecha1 + "',"
            ZSql = ZSql + " FImplementa2 = " + "'" + ZZFecha2 + "',"
            ZSql = ZSql + " FImplementa3 = " + "'" + ZZFecha3 + "',"
            ZSql = ZSql + " FImplementa4 = " + "'" + ZZFecha4 + "',"
            ZSql = ZSql + " FImplementa5 = " + "'" + ZZFecha5 + "',"
            ZSql = ZSql + " FImplementa6 = " + "'" + ZZFecha6 + "',"
            ZSql = ZSql + " OImplementa1 = " + "'" + Left$(ZZObserva1, 100) + "',"
            ZSql = ZSql + " OImplementa2 = " + "'" + Left$(ZZObserva2, 100) + "',"
            ZSql = ZSql + " OImplementa3 = " + "'" + Left$(ZZObserva3, 100) + "',"
            ZSql = ZSql + " OImplementa4 = " + "'" + Left$(ZZObserva4, 100) + "',"
            ZSql = ZSql + " OImplementa5 = " + "'" + Left$(ZZObserva5, 100) + "',"
            ZSql = ZSql + " OImplementa6 = " + "'" + Left$(ZZObserva6, 100) + "',"
            ZSql = ZSql + " DesResponsable1 = " + "'" + ZZDesResponsableAccion1 + "',"
            ZSql = ZSql + " DesResponsable2 = " + "'" + ZZDesResponsableAccion2 + "',"
            ZSql = ZSql + " DesResponsable3 = " + "'" + ZZDesResponsableAccion3 + "',"
            ZSql = ZSql + " DesResponsable4 = " + "'" + ZZDesResponsableAccion4 + "',"
            ZSql = ZSql + " DesResponsable5 = " + "'" + ZZDesResponsableAccion5 + "',"
            ZSql = ZSql + " DesResponsable6 = " + "'" + ZZDesResponsableAccion6 + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZZClave + "'"
            spCargaSacII = ZSql
            Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
        
Rem by nan
        
           
    If TipoListado.ListIndex = 2 Then
    
        Dim ZZImpreVector(10000, 3) As String
        
        Erase ZZImpreVector
        ZZLugar = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Centro >= " + "'" + DesdeCentro.Text + "'"
        ZSql = ZSql + " and CargaSac.Centro <= " + "'" + HastaCentro.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano >= " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano <= " + "'" + AnoII.Text + "'"
        ZSql = ZSql + " and CargaSac.Marca = " + "'" + "S" + "'"
        ZSql = ZSql + " Order by CargaSac.Clave"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
            With rstCargaSac
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZLugar = ZZLugar + 1
                    
                    ZZImpreVector(ZZLugar, 1) = rstCargaSac!Tipo
                    ZZImpreVector(ZZLugar, 2) = rstCargaSac!Ano
                    ZZImpreVector(ZZLugar, 3) = rstCargaSac!Numero
                    
                    .MoveNext
                    
                            Else
                            
                        Exit Do
                    End If
                Loop
            End With
            rstCargaSac.Close
        End If
        
        
        For CicloII = 1 To ZZLugar
        
            ZZLugar = ZZLugar + 1
            
            ZZPasaTipo = ZZImpreVector(CicloII, 1)
            ZZPasaAno = ZZImpreVector(CicloII, 2)
            ZZPasaNumero = ZZImpreVector(CicloII, 3)
            
            Call Impresion_Ficha
        
        Next CicloII
        

            Else
    
            
        Listado.WindowTitle = "Listado de SAC"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        Uno = "{CargaSAc.Centro} in " + DesdeCentro.Text + " to " + HastaCentro.Text
        Dos = " and {CargaSAC.Ano} in " + Ano.Text + " to " + AnoII.Text
        TYres = " and {CargaSAC.Marca} = " + Chr$(34) + "S" + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno + Dos + Tres
        Listado.SelectionFormula = Uno + Dos + Tres
       
        If Impresora.Value = True Then
            Listado.Destination = 1
                Else
            Listado.Destination = 0
        End If
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        If TipoListado.ListIndex = 0 Then
        
            Listado.SQLQuery = "SELECT CargaSAC.Tipo, CargaSAC.Ano, CargaSAC.Numero, CargaSAC.Centro, CargaSAC.Fecha, CargaSAC.OrdFecha, CargaSAC.Origen, CargaSAC.Estado, CargaSAC.IngresoNoCon, CargaSAC.IngresoCausa, CargaSAC.Titulo, CargaSAC.Referencia, CargaSAC.Marca, " _
                        + "CentroSac.Descripcion, " _
                        + "CargaSacII.Accion11, CargaSacII.Accion12, CargaSacII.Accion21, CargaSacII.Accion22, CargaSacII.Accion31, CargaSacII.Accion32, CargaSacII.Accion41, CargaSacII.Accion42, CargaSacII.Accion51, CargaSacII.Accion61, CargaSacII.Accion62, CargaSacII.Responsable1, CargaSacII.Responsable2, CargaSacII.Responsable3, CargaSacII.Responsable4, CargaSacII.Responsable5, CargaSacII.Responsable6, CargaSacII.Plazo1, CargaSacII.Plazo2, CargaSacII.Plazo3, CargaSacII.Plazo4, CargaSacII.Plazo5, CargaSacII.Plazo6, CargaSacII.DesResponsable1, CargaSacII.DesResponsable2, CargaSacII.DesResponsable3, CargaSacII.DesResponsable4, CargaSacII.DesResponsable5, CargaSacII.DesResponsable6, CargaSacII.Estado1, CargaSacII.Estado2, CargaSacII.Estado3, CargaSacII.Estado4, CargaSacII.Estado5, CargaSacII.Estado6, " _
                        + "CargaSacII.FImplementa1, CargaSacII.FImplementa2, CargaSacII.FImplementa3, CargaSacII.FImplementa4, CargaSacII.FImplementa5, CargaSacII.FImplementa6, " _
                        + "CargaSacII.OImplementa1, CargaSacII.OImplementa2, CargaSacII.OImplementa3, CargaSacII.OImplementa4, CargaSacII.OImplementa5, CargaSacII.OImplementa6 " _
                        + "From " _
                        + DSQ + ".dbo.CargaSAC CargaSAC, " _
                        + DSQ + ".dbo.CentroSac CentroSac, " _
                        + DSQ + ".dbo.CargaSacII CargaSacII " _
                        + "Where " _
                        + "CargaSAC.Centro = CentroSac.Codigo AND " _
                        + "CargaSAC.Clave = CargaSacII.Clave AND " _
                        + "CargaSAC.Ano <= " + AnoII.Text + " AND " _
                        + "CargaSAC.Centro >= " + DesdeCentro.Text + " AND " + "CargaSAC.Centro <= " + HastaCentro.Text + " AND " _
                        + "CargaSAC.Marca = 'S'"
                        
            
            
            
            Listado.Connect = Connect()
            Listado.ReportFileName = "ListadoSac.rpt"
            
                Else
        
            Listado.SQLQuery = "SELECT CargaSAC.Tipo, CargaSAC.Ano, CargaSAC.Numero, CargaSAC.Centro, CargaSAC.Fecha, CargaSAC.OrdFecha, CargaSAC.Origen, CargaSAC.Estado, CargaSAC.IngresoNoCon, CargaSAC.IngresoCausa, CargaSAC.Titulo, CargaSAC.Referencia, CargaSAC.Marca, " _
                        + "CentroSac.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.CargaSAC CargaSAC, " _
                        + DSQ + ".dbo.CentroSac CentroSac " _
                        + "Where " _
                        + "CargaSAC.Centro = CentroSac.Codigo AND " _
                        + "CargaSAC.Centro >= " + DesdeCentro.Text + " AND " _
                        + "CargaSAC.Centro <= " + HastaCentro.Text + " AND " _
                        + "CargaSAC.Ano >= " + Ano.Text + " AND " _
                        + "CargaSAC.Ano <= " + AnoII.Text + " AND " _
                        + "CargaSAC.Marca = 'S'"
                        
            Listado.Connect = Connect()
            Listado.ReportFileName = "ListadoSacResumido.rpt"
            
        End If
        
        If salidita = "" Then
            Listado.Action = 1
        End If
    
        salidita = ""
        
    End If
    
End Sub

Private Sub Cancela_click()
    PrgListaSac.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    
   
    
    Rem Tipo.Clear

    Rem Tipo.AddItem "Total"
    Rem Tipo.AddItem "Iniciada"
    Rem Tipo.AddItem "Investigacion"
    Rem Tipo.AddItem "Implementacion"
    Rem Tipo.AddItem "Imple. a Verificar"
    Rem Tipo.AddItem "Imple. Verificada"
    Rem Tipo.AddItem "Cerrada"
    Rem Tipo.AddItem "Total Abiertas"
    
    Rem Tipo.ListIndex = 0
    
    Rem Origen.Clear
    
    Rem Origen.AddItem "Total"
    Rem Origen.AddItem "Auditoria"
    Rem Origen.AddItem "Reclamo"
    Rem Origen.AddItem "I. No Conformidad"
    Rem Origen.AddItem "Proceso/Sist"
    Rem Origen.AddItem "Otro"
    
    Rem Origen.ListIndex = 0
    
    TipoListado.Clear
    
    TipoListado.AddItem "Completo"
    TipoListado.AddItem "Resumido"
    TipoListado.AddItem "Ficha"
    
    TipoListado.ListIndex = 0
    
    DesdeCentro.Text = ""
    HastaCentro.Text = ""
    Ano.Text = ""
    AnoII.Text = ""
    Responsable.Text = ""
    DesResponsable.Caption = ""
    
    ZZLugar = 1
    Call Opcion
    
    Panta.Value = False
    Impresora.Value = True
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

        Call Limpia_Ayuda
        LugarAyuda = 0
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        Select Case ZZLugar
            Case 1
                Sql1 = "Select *"
                Sql2 = " FROM CentroSac"
                Sql3 = " Order by CentroSac.Codigo"
                spCentroSac = Sql1 + Sql2 + Sql3
                Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCentroSac.RecordCount > 0 Then
                    With rstCentroSac
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                da = Len(rstCentroSac!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstCentroSac!Descripcion, aa, WEspacios) Then
                                        LugarAyuda = LugarAyuda + 1
                                        Pantalla.Row = LugarAyuda
                                        Pantalla.Col = 1
                                        Pantalla.Text = rstCentroSac!Codigo
                                        Pantalla.Col = 2
                                        Pantalla.Text = rstCentroSac!Descripcion
                                        IngresaItem = rstCentroSac!Codigo
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
                    rstCentroSac.Close
                End If
    
            Case Else
                Sql1 = "Select *"
                Sql2 = " FROM ResponsableSac"
                Sql3 = " Order by ResponsableSac.Codigo"
                spResponsableSac = Sql1 + Sql2 + Sql3
                Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstResponsableSac.RecordCount > 0 Then
                    With rstResponsableSac
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                da = Len(rstResponsableSac!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                        LugarAyuda = LugarAyuda + 1
                                        Pantalla.Row = LugarAyuda
                                        Pantalla.Col = 1
                                        Pantalla.Text = rstResponsableSac!Codigo
                                        Pantalla.Col = 2
                                        Pantalla.Text = rstResponsableSac!Descripcion
                                        IngresaItem = rstResponsableSac!Codigo
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
                    rstResponsableSac.Close
                End If
        End Select
    End If
    
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    Pantalla.FixedCols = 1
    Pantalla.Cols = 3
    Pantalla.FixedRows = 1
    Pantalla.Rows = 10001
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Select Case Ciclo
            Case 1
                Pantalla.Text = "Codigo"
                Pantalla.ColWidth(Ciclo) = 1000
                Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                Pantalla.Text = "Nombre"
                Pantalla.ColWidth(Ciclo) = 6000
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = Pantalla.Text
        WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        WTitulo(Ciclo).Width = Pantalla.CellWidth
        WTitulo(Ciclo).Height = Pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub

Private Sub pantalla_Click()
    Indice = Pantalla.Row - 1
    Select Case ZZLugar
        Case 1
            DesdeCentro.Text = WIndice.List(Indice)
            HastaCentro.Text = WIndice.List(Indice)
        Case Else
            Responsable.Text = WIndice.List(Indice)
            Call Responsable_Keypress(13)
    End Select
End Sub

Private Sub Opcion()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    Select Case ZZLugar
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstCentroSac!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstCentroSac!Descripcion
                            IngresaItem = rstCentroSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If
        Case Else
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstResponsableSac!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstResponsableSac!Descripcion
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Responsable_DblClick()
    ZZLugar = 2
    Call Opcion
End Sub

Private Sub DesdeCentro_DblClick()
    ZZLugar = 1
    Call Opcion
End Sub

Private Sub HastaCentro_DblClick()
    ZZLugar = 1
    Call Opcion
End Sub

Private Sub WClave22_KeyPress(KeyAscii As Integer)

Rem If KeyAscii = 13 Then
Rem Ano.Text WClave22.Text

Rem End If

End Sub
Private Sub Vertipo_Click()

    ZZOpcion = 1
    
    TituloOpcion.Caption = "TIPOS"
    
    ZZLugar = 0
    Opcion1.Visible = False
    Opcion2.Visible = False
    Opcion3.Visible = False
    Opcion4.Visible = False
    Opcion5.Visible = False
    Opcion6.Visible = False
    Opcion7.Visible = False
    Opcion8.Visible = False
    Opcion9.Visible = False
    Opcion10.Visible = False
    
    If ZZClaveI(1) <> 0 Then
        Opcion1.Value = 1
            Else
        Opcion1.Value = 0
    End If
    If ZZClaveI(2) <> 0 Then
        Opcion2.Value = 1
            Else
        Opcion2.Value = 0
    End If
    If ZZClaveI(3) <> 0 Then
        Opcion3.Value = 1
            Else
        Opcion3.Value = 0
    End If
    If ZZClaveI(4) <> 0 Then
        Opcion4.Value = 1
            Else
        Opcion4.Value = 0
    End If
    If ZZClaveI(5) <> 0 Then
        Opcion5.Value = 1
            Else
        Opcion5.Value = 0
    End If
    If ZZClaveI(6) <> 0 Then
        Opcion6.Value = 1
            Else
        Opcion6.Value = 0
    End If
    If ZZClaveI(7) <> 0 Then
        Opcion7.Value = 1
            Else
        Opcion7.Value = 0
    End If
    If ZZClaveI(8) <> 0 Then
        Opcion8.Value = 1
            Else
        Opcion8.Value = 0
    End If
    If ZZClaveI(9) <> 0 Then
        Opcion9.Value = 1
            Else
        Opcion9.Value = 0
    End If
    If ZZClaveI(10) <> 0 Then
        Opcion10.Value = 1
            Else
        Opcion10.Value = 0
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Order by TipoSac.Codigo"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        With rstTipoSac
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZLugar = ZZLugar + 1
                    Select Case ZZLugar
                        Case 1
                            Opcion1.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(1) = rstTipoSac!Codigo
                            Opcion1.Visible = True
                        Case 2
                            Opcion2.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(2) = rstTipoSac!Codigo
                            Opcion2.Visible = True
                        Case 3
                            Opcion3.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(3) = rstTipoSac!Codigo
                            Opcion3.Visible = True
                        Case 4
                            Opcion4.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(4) = rstTipoSac!Codigo
                            Opcion4.Visible = True
                        Case 5
                            Opcion5.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(5) = rstTipoSac!Codigo
                            Opcion5.Visible = True
                        Case 6
                            Opcion6.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(6) = rstTipoSac!Codigo
                            Opcion6.Visible = True
                        Case 7
                            Opcion7.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(7) = rstTipoSac!Codigo
                            Opcion7.Visible = True
                        Case 8
                            Opcion8.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(8) = rstTipoSac!Codigo
                            Opcion8.Visible = True
                        Case 9
                            Opcion9.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(9) = rstTipoSac!Codigo
                            Opcion9.Visible = True
                        Case 10
                            Opcion10.Caption = Trim(rstTipoSac!Descripcion)
                            ZZClave(10) = rstTipoSac!Codigo
                            Opcion10.Visible = True
                        Case Else
                    End Select
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTipoSac.Close
    End If
    
    PantaOpcion.Visible = True
    
    PantaOpcion.Height = 5655
    PantaOpcion.Left = 2160
    PantaOpcion.Top = 840
    PantaOpcion.Width = 4205
    
End Sub

Private Sub Verestado_Click()

    ZZOpcion = 2
    
    TituloOpcion.Caption = "ESTADO"
    
    ZZLugar = 0
    Opcion1.Visible = False
    Opcion2.Visible = False
    Opcion3.Visible = False
    Opcion4.Visible = False
    Opcion5.Visible = False
    Opcion6.Visible = False
    Opcion7.Visible = False
    Opcion8.Visible = False
    Opcion9.Visible = False
    Opcion10.Visible = False
    
    If ZZClaveII(1) <> 0 Then
        Opcion1.Value = 1
            Else
        Opcion1.Value = 0
    End If
    If ZZClaveII(2) <> 0 Then
        Opcion2.Value = 1
            Else
        Opcion2.Value = 0
    End If
    If ZZClaveII(3) <> 0 Then
        Opcion3.Value = 1
            Else
        Opcion3.Value = 0
    End If
    If ZZClaveII(4) <> 0 Then
        Opcion4.Value = 1
            Else
        Opcion4.Value = 0
    End If
    If ZZClaveII(5) <> 0 Then
        Opcion5.Value = 1
            Else
        Opcion5.Value = 0
    End If
    If ZZClaveII(6) <> 0 Then
        Opcion6.Value = 1
            Else
        Opcion6.Value = 0
    End If
    If ZZClaveII(7) <> 0 Then
        Opcion7.Value = 1
            Else
        Opcion7.Value = 0
    End If
    If ZZClaveII(8) <> 0 Then
        Opcion8.Value = 1
            Else
        Opcion8.Value = 0
    End If
    If ZZClaveII(9) <> 0 Then
        Opcion9.Value = 1
            Else
        Opcion9.Value = 0
    End If
    If ZZClaveII(10) <> 0 Then
        Opcion10.Value = 1
            Else
        Opcion10.Value = 0
    End If
    
    Opcion1.Caption = "Iniciada"
    ZZClave(1) = 1
    Opcion1.Visible = True
    
    Opcion2.Caption = "Investigacion"
    ZZClave(2) = 2
    Opcion2.Visible = True
        
    Opcion3.Caption = "Implementacion"
    ZZClave(3) = 3
    Opcion3.Visible = True
    
    Opcion4.Caption = "Imp.a Verificar"
    ZZClave(4) = 4
    Opcion4.Visible = True
    
    Opcion5.Caption = "Imp. Varificada"
    ZZClave(5) = 5
    Opcion5.Visible = True
    
    Opcion6.Caption = "Cerrada"
    ZZClave(6) = 6
    Opcion6.Visible = True
    
    PantaOpcion.Visible = True
    
    PantaOpcion.Height = 5655
    PantaOpcion.Left = 2160
    PantaOpcion.Top = 840
    PantaOpcion.Width = 4205
    

End Sub


Private Sub VerOrigen_Click()

    ZZOpcion = 3
    
    TituloOpcion.Caption = "ORIGEN"
    
    ZZLugar = 0
    Opcion1.Visible = False
    Opcion2.Visible = False
    Opcion3.Visible = False
    Opcion4.Visible = False
    Opcion5.Visible = False
    Opcion6.Visible = False
    Opcion7.Visible = False
    Opcion8.Visible = False
    Opcion9.Visible = False
    Opcion10.Visible = False
    
    If ZZClaveIII(1) <> 0 Then
        Opcion1.Value = 1
            Else
        Opcion1.Value = 0
    End If
    If ZZClaveIII(2) <> 0 Then
        Opcion2.Value = 1
            Else
        Opcion2.Value = 0
    End If
    If ZZClaveIII(3) <> 0 Then
        Opcion3.Value = 1
            Else
        Opcion3.Value = 0
    End If
    If ZZClaveIII(4) <> 0 Then
        Opcion4.Value = 1
            Else
        Opcion4.Value = 0
    End If
    If ZZClaveIII(5) <> 0 Then
        Opcion5.Value = 1
            Else
        Opcion5.Value = 0
    End If
    If ZZClaveIII(6) <> 0 Then
        Opcion6.Value = 1
            Else
        Opcion6.Value = 0
    End If
    If ZZClaveIII(7) <> 0 Then
        Opcion7.Value = 1
            Else
        Opcion7.Value = 0
    End If
    If ZZClaveIII(8) <> 0 Then
        Opcion8.Value = 1
            Else
        Opcion8.Value = 0
    End If
    If ZZClaveIII(9) <> 0 Then
        Opcion9.Value = 1
            Else
        Opcion9.Value = 0
    End If
    If ZZClaveIII(10) <> 0 Then
        Opcion10.Value = 1
            Else
        Opcion10.Value = 0
    End If
    
    Opcion1.Caption = "Auditoria"
    ZZClave(1) = 1
    Opcion1.Visible = True
    
    Opcion2.Caption = "Reclamo"
    ZZClave(2) = 2
    Opcion2.Visible = True
        
    Opcion3.Caption = "I. No Conformidad"
    ZZClave(3) = 3
    Opcion3.Visible = True
    
    Opcion4.Caption = "Proceso/Sist."
    ZZClave(4) = 4
    Opcion4.Visible = True
    
    Opcion5.Caption = "Otros"
    ZZClave(5) = 5
    Opcion5.Visible = True
    
    PantaOpcion.Visible = True
    
    PantaOpcion.Height = 5655
    PantaOpcion.Left = 2160
    PantaOpcion.Top = 840
    PantaOpcion.Width = 4205
    
End Sub



Private Sub CerrarOpciones_Click()

    PantaOpcion.Visible = False
    
    Select Case ZZOpcion
        Case 1
            Erase ZZClaveI
        Case 2
            Erase ZZClaveII
        Case Else
            Erase ZZClaveIII
    End Select
    
    If Opcion1.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(1) = ZZClave(1)
            Case 2
                ZZClaveII(1) = ZZClave(1)
            Case Else
                ZZClaveIII(1) = ZZClave(1)
        End Select
    End If
    
    If Opcion2.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(2) = ZZClave(2)
            Case 2
                ZZClaveII(2) = ZZClave(2)
            Case Else
                ZZClaveIII(2) = ZZClave(2)
        End Select
    End If
    
    If Opcion3.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(3) = ZZClave(3)
            Case 2
                ZZClaveII(3) = ZZClave(3)
            Case Else
                ZZClaveIII(3) = ZZClave(3)
        End Select
    End If
    
    If Opcion4.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(4) = ZZClave(4)
            Case 2
                ZZClaveII(4) = ZZClave(4)
            Case Else
                ZZClaveIII(4) = ZZClave(4)
        End Select
    End If
    
    If Opcion5.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(5) = ZZClave(5)
            Case 2
                ZZClaveII(5) = ZZClave(5)
            Case Else
                ZZClaveIII(5) = ZZClave(5)
        End Select
    End If
    
    If Opcion6.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(6) = ZZClave(6)
            Case 2
                ZZClaveII(6) = ZZClave(6)
            Case Else
                ZZClaveIII(6) = ZZClave(6)
        End Select
    End If
    
    If Opcion7.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(7) = ZZClave(7)
            Case 2
                ZZClaveII(7) = ZZClave(7)
            Case Else
                ZZClaveIII(7) = ZZClave(7)
        End Select
    End If
    
    If Opcion8.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(8) = ZZClave(8)
            Case 2
                ZZClaveII(8) = ZZClave(8)
            Case Else
                ZZClaveIII(8) = ZZClave(8)
        End Select
    End If
    
    If Opcion9.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(9) = ZZClave(9)
            Case 2
                ZZClaveII(9) = ZZClave(9)
            Case Else
                ZZClaveIII(9) = ZZClave(9)
        End Select
    End If
    
    If Opcion10.Value = 1 Then
        Select Case ZZOpcion
            Case 1
                ZZClaveI(10) = ZZClave(10)
            Case 2
                ZZClaveII(10) = ZZClave(10)
            Case Else
                ZZClaveIII(10) = ZZClave(10)
        End Select
    End If
    
End Sub



Private Sub Impresion_Ficha()

    Dim ZZImpreEstado(100) As String
    Dim ZZImpreEstadoI(100) As String
    Dim ZZImpreEstadoII(100) As String
    Dim ZZImpreOrigen(100) As String
    
    ZZImpreEstado(1) = "INICIADA"
    ZZImpreEstado(2) = "INVESTIGACION"
    ZZImpreEstado(3) = "IMPLEMENTACION"
    ZZImpreEstado(4) = "IMPLEMENTACION A VERIFICAR"
    ZZImpreEstado(5) = "IMPLEMENTACION VERIFICADA"
    ZZImpreEstado(6) = "CERRADA"
    ZZImpreEstado(7) = "ANULADA"

    ZZImpreOrigen(1) = "Auditoria"
    ZZImpreOrigen(2) = "Reclamo"
    ZZImpreOrigen(3) = "I. No Conformidad"
    ZZImpreOrigen(4) = "Proceso/Sist"
    ZZImpreOrigen(5) = "Otro"
    
    ZZImpreEstadoI(1) = "Imple."
    ZZImpreEstadoI(2) = "Nula"
    ZZImpreEstadoI(3) = "No Imple."
    
    ZZImpreEstadoII(1) = "No Imple."
    ZZImpreEstadoII(2) = "Imple."
    ZZImpreEstadoII(3) = "Nula"
    ZZImpreEstadoII(4) = "Cerrada"

    Auxi3 = ZZPasaTipo
    Auxi1 = ZZPasaAno
    Auxi2 = ZZPasaNumero
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    ZZZClave = Auxi3 + Auxi1 + Auxi2

    ZSql = ""
    ZSql = ZSql + "DELETE ImpreSac"
    Rem ZSql = ZSql + " Where ImpreSac.Tipo = " + "'" + Tipo.Text + "'"
    Rem ZSql = ZSql + " and ImpreSac.Ano = " + "'" + Ano.Text + "'"
    Rem ZSql = ZSql + " and ImpreSac.Numero = " + "'" + Numero.Text + "'"
    rsImpreSac = ZSql
    Set rstImpreSac = db.OpenRecordset(rsImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZZPasaTipo + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZZPasaAno + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZZPasaNumero + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        
        ZZCentro = rstCargaSac!Centro
        ZZFecha = rstCargaSac!Fecha
        ZZOrigen = ZZImpreOrigen(rstCargaSac!Origen)
        ZZEstado = ZZImpreEstado(rstCargaSac!Estado)
        ZZResponsableEmisor = rstCargaSac!ResponsableEmisor
        ZZResponsableDestino = rstCargaSac!ResponsableDestino
        ZZReferencia = Trim(rstCargaSac!Referencia)
        ZZTitulo = rstCargaSac!Titulo
        IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
        IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
        
        rstCargaSac.Close
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacII"
        ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + ZZPasaTipo + "'"
        ZSql = ZSql + " and CargaSacII.Ano = " + "'" + ZZPasaAno + "'"
        ZSql = ZSql + " and CargaSacII.Numero = " + "'" + ZZPasaNumero + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacII.RecordCount > 0 Then
        
            ZZAccion11 = Trim(rstCargaSacII!Accion11)
            ZZAccion12 = Trim(rstCargaSacII!Accion12)
            ZZAccion21 = Trim(rstCargaSacII!Accion21)
            ZZAccion22 = Trim(rstCargaSacII!Accion22)
            ZZAccion31 = Trim(rstCargaSacII!Accion31)
            ZZAccion32 = Trim(rstCargaSacII!Accion32)
            ZZAccion41 = Trim(rstCargaSacII!Accion41)
            ZZAccion42 = Trim(rstCargaSacII!Accion42)
            ZZAccion51 = Trim(rstCargaSacII!Accion51)
            ZZAccion52 = Trim(rstCargaSacII!Accion52)
            ZZAccion61 = Trim(rstCargaSacII!Accion61)
            ZZAccion62 = Trim(rstCargaSacII!Accion62)
            
            ZZResponsable1 = rstCargaSacII!Responsable1
            ZZResponsable2 = rstCargaSacII!Responsable2
            ZZResponsable3 = rstCargaSacII!Responsable3
            ZZResponsable4 = rstCargaSacII!Responsable4
            ZZResponsable5 = rstCargaSacII!Responsable5
            ZZResponsable6 = rstCargaSacII!Responsable6
            
            ZZPlazo1 = rstCargaSacII!Plazo1
            ZZPlazo2 = rstCargaSacII!Plazo2
            ZZPlazo3 = rstCargaSacII!Plazo3
            ZZPlazo4 = rstCargaSacII!Plazo4
            ZZPlazo5 = rstCargaSacII!Plazo5
            ZZPlazo6 = rstCargaSacII!Plazo6
            
            rstCargaSacII.Close
        End If
        
    
    
    
    
        
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIII"
        ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + ZZPasaTipo + "'"
        ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + ZZPasaAno + "'"
        ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + ZZPasaNumero + "'"
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIII.RecordCount > 0 Then
        
            ZZResponsable11 = rstCargaSacIII!Responsable1
            ZZResponsable12 = rstCargaSacIII!Responsable2
            ZZResponsable13 = rstCargaSacIII!Responsable3
            ZZResponsable14 = rstCargaSacIII!Responsable4
            ZZResponsable15 = rstCargaSacIII!Responsable5
            ZZResponsable16 = rstCargaSacIII!Responsable6
            
            ZZFecha1 = rstCargaSacIII!Fecha1
            ZZFecha2 = rstCargaSacIII!Fecha2
            ZZFecha3 = rstCargaSacIII!Fecha3
            ZZFecha4 = rstCargaSacIII!Fecha4
            ZZFecha5 = rstCargaSacIII!Fecha5
            ZZFecha6 = rstCargaSacIII!Fecha6
            
            ZZComentario11 = Trim(rstCargaSacIII!Comentario11)
            ZZComentario12 = Trim(rstCargaSacIII!Comentario12)
            ZZComentario21 = Trim(rstCargaSacIII!Comentario21)
            ZZComentario22 = Trim(rstCargaSacIII!Comentario22)
            ZZComentario31 = Trim(rstCargaSacIII!Comentario31)
            ZZComentario32 = Trim(rstCargaSacIII!Comentario32)
            ZZComentario41 = Trim(rstCargaSacIII!Comentario41)
            ZZComentario42 = Trim(rstCargaSacIII!Comentario42)
            ZZComentario51 = Trim(rstCargaSacIII!Comentario51)
            ZZComentario52 = Trim(rstCargaSacIII!Comentario52)
            ZZComentario61 = Trim(rstCargaSacIII!Comentario61)
            ZZComentario62 = Trim(rstCargaSacIII!Comentario62)
            
            ZZZZEstado11 = rstCargaSacIII!Estado1
            ZZZZEstado12 = rstCargaSacIII!Estado2
            ZZZZEstado13 = rstCargaSacIII!Estado3
            ZZZZEstado14 = rstCargaSacIII!Estado4
            ZZZZEstado15 = rstCargaSacIII!Estado5
            ZZZZEstado16 = rstCargaSacIII!Estado6
            
            If ZZZZEstado11 < 0 Then
                ZZZZEstado11 = 0
            End If
            If ZZZZEstado12 < 0 Then
                ZZZZEstado12 = 0
            End If
            If ZZZZEstado13 < 0 Then
                ZZZZEstado13 = 0
            End If
            If ZZZZEstado14 < 0 Then
                ZZZZEstado14 = 0
            End If
            If ZZZZEstado15 < 0 Then
                ZZZZEstado15 = 0
            End If
            If ZZZZEstado16 < 0 Then
                ZZZZEstado16 = 0
            End If
            
            ZZEstado11 = ZZImpreEstadoI(ZZZZEstado11)
            ZZEstado12 = ZZImpreEstadoI(ZZZZEstado12)
            ZZEstado13 = ZZImpreEstadoI(ZZZZEstado13)
            ZZEstado14 = ZZImpreEstadoI(ZZZZEstado14)
            ZZEstado15 = ZZImpreEstadoI(ZZZZEstado15)
            ZZEstado16 = ZZImpreEstadoI(ZZZZEstado16)
            
            rstCargaSacIII.Close
        End If
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIV"
        ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + ZZPasaTipo + "'"
        ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + ZZPasaAno + "'"
        ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + ZZPasaNumero + "'"
        spCargaSacIV = ZSql
        Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIV.RecordCount > 0 Then
        
            ZZResponsable21 = rstCargaSacIV!Responsable1
            ZZResponsable22 = rstCargaSacIV!Responsable2
            ZZResponsable23 = rstCargaSacIV!Responsable3
            ZZResponsable24 = rstCargaSacIV!Responsable4
            ZZResponsable25 = rstCargaSacIV!Responsable5
            ZZResponsable26 = rstCargaSacIV!Responsable6
            
            ZZResponsable31 = IIf(IsNull(rstCargaSacIV!Responsable11), "", rstCargaSacIV!Responsable11)
            ZZResponsable32 = IIf(IsNull(rstCargaSacIV!Responsable12), "", rstCargaSacIV!Responsable12)
            ZZResponsable33 = IIf(IsNull(rstCargaSacIV!Responsable13), "", rstCargaSacIV!Responsable13)
            ZZResponsable34 = IIf(IsNull(rstCargaSacIV!Responsable14), "", rstCargaSacIV!Responsable14)
            ZZResponsable35 = IIf(IsNull(rstCargaSacIV!Responsable15), "", rstCargaSacIV!Responsable15)
            ZZResponsable36 = IIf(IsNull(rstCargaSacIV!Responsable16), "", rstCargaSacIV!Responsable16)
            
            ZZFecha21 = rstCargaSacIV!Fecha1
            ZZFecha22 = rstCargaSacIV!Fecha2
            ZZFecha23 = rstCargaSacIV!Fecha3
            ZZFecha24 = rstCargaSacIV!Fecha4
            ZZFecha25 = rstCargaSacIV!Fecha5
            ZZFecha26 = rstCargaSacIV!Fecha6
            
            ZZFecha31 = IIf(IsNull(rstCargaSacIV!Fecha11), "  /  /    ", rstCargaSacIV!Fecha11)
            ZZFecha32 = IIf(IsNull(rstCargaSacIV!Fecha12), "  /  /    ", rstCargaSacIV!Fecha12)
            ZZFecha33 = IIf(IsNull(rstCargaSacIV!Fecha13), "  /  /    ", rstCargaSacIV!Fecha13)
            ZZFecha34 = IIf(IsNull(rstCargaSacIV!Fecha14), "  /  /    ", rstCargaSacIV!Fecha14)
            ZZFecha35 = IIf(IsNull(rstCargaSacIV!Fecha15), "  /  /    ", rstCargaSacIV!Fecha15)
            ZZFecha36 = IIf(IsNull(rstCargaSacIV!Fecha16), "  /  /    ", rstCargaSacIV!Fecha16)
            
            ZZComentario211 = Trim(rstCargaSacIV!Comentario11)
            ZZComentario212 = Trim(rstCargaSacIV!Comentario12)
            ZZComentario221 = Trim(rstCargaSacIV!Comentario21)
            ZZComentario222 = Trim(rstCargaSacIV!Comentario22)
            ZZComentario231 = Trim(rstCargaSacIV!Comentario31)
            ZZComentario232 = Trim(rstCargaSacIV!Comentario32)
            ZZComentario241 = Trim(rstCargaSacIV!Comentario41)
            ZZComentario242 = Trim(rstCargaSacIV!Comentario42)
            ZZComentario251 = Trim(rstCargaSacIV!Comentario51)
            ZZComentario252 = Trim(rstCargaSacIV!Comentario52)
            ZZComentario261 = Trim(rstCargaSacIV!Comentario61)
            ZZComentario262 = Trim(rstCargaSacIV!Comentario62)
            
            ZZEstado1 = ZZImpreEstadoII(rstCargaSacIV!Estado1)
            ZZEstado2 = ZZImpreEstadoII(rstCargaSacIV!Estado2)
            ZZEstado3 = ZZImpreEstadoII(rstCargaSacIV!Estado3)
            ZZEstado4 = ZZImpreEstadoII(rstCargaSacIV!Estado4)
            ZZEstado5 = ZZImpreEstadoII(rstCargaSacIV!Estado5)
            ZZEstado6 = ZZImpreEstadoII(rstCargaSacIV!Estado6)
            
            ZZEstado31 = IIf(IsNull(rstCargaSacIV!Estado11), "0", rstCargaSacIV!Estado11)
            ZZEstado32 = IIf(IsNull(rstCargaSacIV!Estado12), "0", rstCargaSacIV!Estado12)
            ZZEstado33 = IIf(IsNull(rstCargaSacIV!Estado13), "0", rstCargaSacIV!Estado13)
            ZZEstado34 = IIf(IsNull(rstCargaSacIV!Estado14), "0", rstCargaSacIV!Estado14)
            ZZEstado35 = IIf(IsNull(rstCargaSacIV!Estado15), "0", rstCargaSacIV!Estado15)
            ZZEstado36 = IIf(IsNull(rstCargaSacIV!Estado16), "0", rstCargaSacIV!Estado16)
            
            ZZEstado31 = ZZImpreEstadoII(ZZEstado31)
            ZZEstado32 = ZZImpreEstadoII(ZZEstado32)
            ZZEstado33 = ZZImpreEstadoII(ZZEstado33)
            ZZEstado34 = ZZImpreEstadoII(ZZEstado34)
            ZZEstado35 = ZZImpreEstadoII(ZZEstado35)
            ZZEstado36 = ZZImpreEstadoII(ZZEstado36)
            
            
            rstCargaSacIV.Close
            
        End If
        
        ZResponsableEmisor = Val(ZZResponsableEmisor)
        ZResponsableDestino = Val(ZZResponsableDestino)
        
        ZResponsable1 = Val(ZZResponsable1)
        ZResponsable2 = Val(ZZResponsable2)
        ZResponsable3 = Val(ZZResponsable3)
        ZResponsable4 = Val(ZZResponsable4)
        ZResponsable5 = Val(ZZResponsable5)
        ZResponsable6 = Val(ZZResponsable6)
        
        ZResponsable11 = Val(ZZResponsable11)
        ZResponsable12 = Val(ZZResponsable12)
        ZResponsable13 = Val(ZZResponsable13)
        ZResponsable14 = Val(ZZResponsable14)
        ZResponsable15 = Val(ZZResponsable15)
        ZResponsable16 = Val(ZZResponsable16)
        
        ZResponsable21 = Val(ZZResponsable21)
        ZResponsable22 = Val(ZZResponsable22)
        ZResponsable23 = Val(ZZResponsable23)
        ZResponsable24 = Val(ZZResponsable24)
        ZResponsable25 = Val(ZZResponsable25)
        ZResponsable26 = Val(ZZResponsable26)
        
        ZResponsable31 = Val(ZZResponsable31)
        ZResponsable32 = Val(ZZResponsable32)
        ZResponsable33 = Val(ZZResponsable33)
        ZResponsable34 = Val(ZZResponsable34)
        ZResponsable35 = Val(ZZResponsable35)
        ZResponsable36 = Val(ZZResponsable36)
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacAdicional"
        ZSql = ZSql + " Where CargaSacAdicional.Tipo = " + "'" + ZZPasaTipo + "'"
        ZSql = ZSql + " and CargaSacAdicional.Ano = " + "'" + ZZPasaAno + "'"
        ZSql = ZSql + " and CargaSacAdicional.Numero = " + "'" + ZZPasaNumero + "'"
        spCargaSacAdicional = ZSql
        Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacAdicional.RecordCount > 0 Then
            Comentario.Text = IIf(IsNull(rstCargaSacAdicional!Dato1), "", rstCargaSacAdicional!Dato1)
            rstCargaSacAdicional.Close
        End If
        
        
        
        
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + ZZPasaTipo + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            ZZDesTipo = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM CentroSac"
        Sql3 = " Where CentroSac.Codigo = " + "'" + ZZCentro + "'"
        spCentroSac = Sql1 + Sql2 + Sql3
        Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCentroSac.RecordCount > 0 Then
            ZZDesCentro = Trim(rstCentroSac!Descripcion)
            rstCentroSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsableEmisor + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsableEmisor = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsableDestino + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsableDestino = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
    
    
    
    
    
    
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable1 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable1 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable2 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable2 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable3 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable3 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable4 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable4 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable5 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable5 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable6 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable6 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable11 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable11 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable12 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable12 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable13 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable13 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable14 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable14 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable15 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable15 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable16 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable16 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        
        
        
        
        
        
        
        
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable21 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable21 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable22 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable22 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable23 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable23 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable24 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable24 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable25 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable25 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable26 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable26 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        
        
        
        
        
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable31 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable31 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable32 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable32 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable33 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable33 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable34 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable34 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable35 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable35 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ZZResponsable36 + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            ZZDesResponsable36 = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
        End If
        
        
        
        
        
        
        
        Sql1 = "Select *"
        Sql2 = " FROM CentroSac"
        Sql3 = " Where CentroSac.Codigo = " + "'" + ZZCentro + "'"
        spCentroSac = Sql1 + Sql2 + Sql3
        Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCentroSac.RecordCount > 0 Then
            ZZDesCentro = Trim(rstCentroSac!Descripcion)
            rstCentroSac.Close
        End If
        
        
            
        
        
        
        
        
        
        
        
        ZZGraba = "S"
        
        If Trim(ZZAccion11) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion11
            ZZImpre12 = ZZAccion12
            ZZImpre13 = ZZDesResponsable1
            ZZImpre14 = ZZPlazo1
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable11
            ZZImpre24 = ZZFecha1
            ZZImpre25 = ZZEstado11
            ZZImpre26 = ZZComentario11
            ZZImpre27 = ZZComentario12
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable21
            ZZImpre32 = ZZEstado1
            ZZImpre33 = ZZFecha21
            ZZImpre34 = ZZDesResponsable31
            ZZImpre35 = ZZEstado31
            ZZImpre36 = ZZFecha31
            ZZImpre37 = ZZComentario211
            ZZImpre38 = ZZComentario212
            ZZImpre39 = ""
            ZZImpre40 = "1"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        If Trim(ZZAccion21) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion21
            ZZImpre12 = ZZAccion22
            ZZImpre13 = ZZDesResponsable2
            ZZImpre14 = ZZPlazo2
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable12
            ZZImpre24 = ZZFecha2
            ZZImpre25 = ZZEstado12
            ZZImpre26 = ZZComentario21
            ZZImpre27 = ZZComentario22
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable22
            ZZImpre32 = ZZEstado2
            ZZImpre33 = ZZFecha22
            ZZImpre34 = ZZDesResponsable32
            ZZImpre35 = ZZEstado32
            ZZImpre36 = ZZFecha32
            ZZImpre37 = ZZComentario221
            ZZImpre38 = ZZComentario222
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + xxDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        
        
        
        If Trim(ZZAccion31) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion31
            ZZImpre12 = ZZAccion32
            ZZImpre13 = ZZDesResponsable3
            ZZImpre14 = ZZPlazo3
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable13
            ZZImpre24 = ZZFecha3
            ZZImpre25 = ZZEstado13
            ZZImpre26 = ZZComentario31
            ZZImpre27 = ZZComentario32
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable23
            ZZImpre32 = ZZEstado3
            ZZImpre33 = ZZFecha23
            ZZImpre34 = ZZDesResponsable33
            ZZImpre35 = ZZEstado33
            ZZImpre36 = ZZFecha33
            ZZImpre37 = ZZComentario231
            ZZImpre38 = ZZComentario232
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        
        
        
        If Trim(ZZAccion41) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion41
            ZZImpre12 = ZZAccion42
            ZZImpre13 = ZZDesResponsable4
            ZZImpre14 = ZZPlazo4
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable14
            ZZImpre24 = ZZFecha4
            ZZImpre25 = ZZEstado14
            ZZImpre26 = ZZComentario41
            ZZImpre27 = ZZComentario42
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable24
            ZZImpre32 = ZZEstado4
            ZZImpre33 = ZZFecha24
            ZZImpre34 = ZZDesResponsable34
            ZZImpre35 = ZZEstado34
            ZZImpre36 = ZZFecha34
            ZZImpre37 = ZZComentario241
            ZZImpre38 = ZZComentario242
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        If Trim(ZZAccion51) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion51
            ZZImpre12 = ZZAccion52
            ZZImpre13 = ZZDesResponsable5
            ZZImpre14 = ZZPlazo5
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable15
            ZZImpre24 = ZZFecha5
            ZZImpre25 = ZZEstado15
            ZZImpre26 = ZZComentario51
            ZZImpre27 = ZZComentario52
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable25
            ZZImpre32 = ZZEstado5
            ZZImpre33 = ZZFecha25
            ZZImpre34 = ZZDesResponsable35
            ZZImpre35 = ZZEstado35
            ZZImpre36 = ZZFecha35
            ZZImpre37 = ZZComentario251
            ZZImpre38 = ZZComentario252
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        If Trim(ZZAccion61) <> "" Then
            
            ZZGraba = "N"
            
            ZZImpre11 = ZZAccion61
            ZZImpre12 = ZZAccion62
            ZZImpre13 = ZZDesResponsable6
            ZZImpre14 = ZZPlazo6
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ZZDesResponsable16
            ZZImpre24 = ZZFecha6
            ZZImpre25 = ZZEstado16
            ZZImpre26 = ZZComentario61
            ZZImpre27 = ZZComentario62
            ZZImpre28 = ""
            
            ZZImpre31 = ZZDesResponsable26
            ZZImpre32 = ZZEstado6
            ZZImpre33 = ZZFecha26
            ZZImpre34 = ZZDesResponsable36
            ZZImpre35 = ZZEstado36
            ZZImpre36 = ZZFecha36
            ZZImpre37 = ZZComentario261
            ZZImpre38 = ZZComentario262
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        If ZZGraba = "S" Then
            
            ZZImpre11 = ""
            ZZImpre12 = ""
            ZZImpre13 = ""
            ZZImpre14 = ""
            ZZImpre15 = ""
            
            ZZImpre21 = ""
            ZZImpre22 = ""
            ZZImpre23 = ""
            ZZImpre24 = ""
            ZZImpre25 = ""
            ZZImpre26 = ""
            ZZImpre27 = ""
            ZZImpre28 = ""
            
            ZZImpre31 = ""
            ZZImpre32 = ""
            ZZImpre33 = ""
            ZZImpre34 = ""
            ZZImpre35 = ""
            ZZImpre36 = ""
            ZZImpre37 = ""
            ZZImpre38 = ""
            ZZImpre39 = ""
            ZZImpre40 = "2"
            
            ZZCorte = "1"
            
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "DesTipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "DesCentro ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "IngresoNoCon ,"
            ZSql = ZSql + "IngresoCausa ,"
            ZSql = ZSql + "DesResponsableEmisor ,"
            ZSql = ZSql + "DesResponsableDestino ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Referencia ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Impre11 ,"
            ZSql = ZSql + "Impre12 ,"
            ZSql = ZSql + "Impre13 ,"
            ZSql = ZSql + "Impre14 ,"
            ZSql = ZSql + "Impre15 ,"
            ZSql = ZSql + "Impre21 ,"
            ZSql = ZSql + "Impre22 ,"
            ZSql = ZSql + "Impre23 ,"
            ZSql = ZSql + "Impre24 ,"
            ZSql = ZSql + "Impre25 ,"
            ZSql = ZSql + "Impre26 ,"
            ZSql = ZSql + "Impre27 ,"
            ZSql = ZSql + "Impre28 ,"
            ZSql = ZSql + "Impre31 ,"
            ZSql = ZSql + "Impre32 ,"
            ZSql = ZSql + "Impre33 ,"
            ZSql = ZSql + "Impre34 ,"
            ZSql = ZSql + "Impre35 ,"
            ZSql = ZSql + "Impre36 ,"
            ZSql = ZSql + "Impre37 ,"
            ZSql = ZSql + "Impre38 ,"
            ZSql = ZSql + "Impre39 ,"
            ZSql = ZSql + "Impre40 ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZClave + "',"
            ZSql = ZSql + "'" + ZZPasaTipo + "',"
            ZSql = ZSql + "'" + ZZDesTipo + "',"
            ZSql = ZSql + "'" + ZZPasaAno + "',"
            ZSql = ZSql + "'" + ZZPasaNumero + "',"
            ZSql = ZSql + "'" + ZZDesCentro + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZOrigen + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
            ZSql = ZSql + "'" + IngresoCausa.Text + "',"
            ZSql = ZSql + "'" + ZZDesResponsableEmisor + "',"
            ZSql = ZSql + "'" + ZZDesResponsableDestino + "',"
            ZSql = ZSql + "'" + ZZTitulo + "',"
            ZSql = ZSql + "'" + ZZReferencia + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZImpre11 + "',"
            ZSql = ZSql + "'" + ZZImpre12 + "',"
            ZSql = ZSql + "'" + ZZImpre13 + "',"
            ZSql = ZSql + "'" + ZZImpre14 + "',"
            ZSql = ZSql + "'" + ZZImpre15 + "',"
            ZSql = ZSql + "'" + ZZImpre21 + "',"
            ZSql = ZSql + "'" + ZZImpre22 + "',"
            ZSql = ZSql + "'" + ZZImpre23 + "',"
            ZSql = ZSql + "'" + ZZImpre24 + "',"
            ZSql = ZSql + "'" + ZZImpre25 + "',"
            ZSql = ZSql + "'" + ZZImpre26 + "',"
            ZSql = ZSql + "'" + ZZImpre27 + "',"
            ZSql = ZSql + "'" + ZZImpre28 + "',"
            ZSql = ZSql + "'" + ZZImpre31 + "',"
            ZSql = ZSql + "'" + ZZImpre32 + "',"
            ZSql = ZSql + "'" + ZZImpre33 + "',"
            ZSql = ZSql + "'" + ZZImpre34 + "',"
            ZSql = ZSql + "'" + ZZImpre35 + "',"
            ZSql = ZSql + "'" + ZZImpre36 + "',"
            ZSql = ZSql + "'" + ZZImpre37 + "',"
            ZSql = ZSql + "'" + ZZImpre38 + "',"
            ZSql = ZSql + "'" + ZZImpre39 + "',"
            ZSql = ZSql + "'" + ZZImpre40 + "',"
            ZSql = ZSql + "'" + Comentario.Text + "')"
            
            spImpreSac = ZSql
            Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        
        
        
        
        
        
        Listado.WindowTitle = "Impresion de Ficha"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT ImpreSac.Tipo, ImpreSac.DesTipo, ImpreSac.Ano, ImpreSac.Numero, ImpreSac.DesCentro, ImpreSac.Fecha, ImpreSac.Origen, ImpreSac.Estado, ImpreSac.IngresoNoCon, ImpreSac.IngresoCausa, ImpreSac.DesResponsableEmisor, ImpreSac.DesResponsableDestino, ImpreSac.Titulo, ImpreSac.Referencia, ImpreSac.Corte, ImpreSac.Impre11, ImpreSac.Impre12, ImpreSac.Impre13, ImpreSac.Impre14, ImpreSac.Impre23, ImpreSac.Impre24, ImpreSac.Impre25, ImpreSac.Impre26, ImpreSac.Impre27, ImpreSac.Impre31, ImpreSac.Impre32, ImpreSac.Impre33, ImpreSac.Impre34, ImpreSac.Impre35, ImpreSac.Impre36, ImpreSac.Impre40, ImpreSac.Comentario " _
                + "From " _
                + DSQ + ".dbo.ImpreSac ImpreSac"
                
        Rem Uno = "{Planifica.ResponsableII} in " + ZDesdeII + " to " + ZHastaII
        Rem Dos = " and {Planifica.OrdVencimiento} in " + Chr$(34) + DesdeFecha + Chr$(34) + " to " + Chr$(34) + HastaFecha + Chr$(34)
        Rem Tres = " and {Planifica.Estado} in " + ZDesdeIII + " to " + ZHastaIII
        Rem Cuatro = " and {Planifica.Responsable} in " + ZDesdeI + " to " + ZHastaI
        
        Listado.GroupSelectionFormula = ""
        Listado.SelectionFormula = ""
        
        Listado.Connect = Connect()
        
        If Impresora.Value = True Then
            Listado.Destination = 1
                Else
            Listado.Destination = 0
        End If
        
        Listado.ReportFileName = "ImpreSac.Rpt"
        Listado.Action = 1
    
    
    End If

End Sub


