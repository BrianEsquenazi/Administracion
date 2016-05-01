VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgChoferes 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Choferes"
   ClientHeight    =   8310
   ClientLeft      =   285
   ClientTop       =   480
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   11430
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
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   36
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox AplicaIII 
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox ComentarioI 
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
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   20
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox ComentarioII 
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
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox ComentarioIII 
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
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   18
      Top             =   3480
      Width           =   2775
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Hasta 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   10
         Text            =   " "
         Top             =   360
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1200
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "choferes.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "choferes.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":0A56
         ToolTipText     =   "Cancela la Impresion"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5280
      TabIndex        =   14
      Top             =   1440
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "choferes.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "choferes.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "choferes.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "choferes.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "choferes.frx":2786
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Width           =   480
      End
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
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Codigo 
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
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Choferes.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Efluentes de Lavado"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
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
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   9015
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
      Height          =   2160
      Left            =   1080
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox FechaVtoI 
      Height          =   300
      Left            =   4800
      TabIndex        =   22
      Top             =   2760
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
   Begin MSMask.MaskEdBox FechaEntregaI 
      Height          =   300
      Left            =   6720
      TabIndex        =   23
      Top             =   2760
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
   Begin MSMask.MaskEdBox FechaVtoII 
      Height          =   300
      Left            =   4800
      TabIndex        =   24
      Top             =   3120
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
   Begin MSMask.MaskEdBox FechaEntregaII 
      Height          =   300
      Left            =   6720
      TabIndex        =   25
      Top             =   3120
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
   Begin MSMask.MaskEdBox FechaVtoIII 
      Height          =   300
      Left            =   4800
      TabIndex        =   26
      Top             =   3480
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
   Begin MSMask.MaskEdBox FechaEntregaIII 
      Height          =   300
      Left            =   6720
      TabIndex        =   27
      Top             =   3480
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
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   1080
      Width           =   1575
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
      Left            =   3840
      TabIndex        =   37
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Licencia de Conducir"
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
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Art"
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
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cert. Habilit. P/Transp de Cargas Peligrosas"
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
      Index           =   5
      Left            =   120
      TabIndex        =   33
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
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
      Index           =   8
      Left            =   120
      TabIndex        =   32
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   4800
      TabIndex        =   31
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Entrega"
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
      Index           =   10
      Left            =   6720
      TabIndex        =   30
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obervaciones"
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
      Index           =   11
      Left            =   8400
      TabIndex        =   29
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aplica"
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
      Index           =   12
      Left            =   3840
      TabIndex        =   28
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "choferes.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "choferes.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "choferes.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "choferes.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "choferes.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "choferes.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "choferes.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre y Apellido"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgChoferes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstChofer As Recordset
Dim spChofer As String

Dim ZAplicaIII As Integer

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()
    Sql1 = "Select *"
    Sql2 = " FROM Chofer"
    Sql3 = " Where Chofer.Codigo = " + "'" + Codigo.Text + "'"
    spChofer = Sql1 + Sql2 + Sql3
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        Descripcion.Text = Trim(rstChofer!Descripcion)
        Proveedor.Text = IIf(IsNull(rstChofer!Proveedor), "", rstChofer!Proveedor)
        
        ZAplicaIII = IIf(IsNull(rstChofer!AplicaIII), "0", rstChofer!AplicaIII)
        AplicaIII.Value = ZAplicaIII
        
        FechaVtoI.Text = IIf(IsNull(rstChofer!FechaVtoI), "  /  /    ", rstChofer!FechaVtoI)
        FechaVtoII.Text = IIf(IsNull(rstChofer!FechaVtoII), "  /  /    ", rstChofer!FechaVtoII)
        FechaVtoIII.Text = IIf(IsNull(rstChofer!FechaVtoIII), "  /  /    ", rstChofer!FechaVtoIII)
        
        FechaEntregaI.Text = IIf(IsNull(rstChofer!FechaEntregaI), "  /  /    ", rstChofer!FechaEntregaI)
        FechaEntregaII.Text = IIf(IsNull(rstChofer!FechaEntregaII), "  /  /    ", rstChofer!FechaEntregaII)
        FechaEntregaIII.Text = IIf(IsNull(rstChofer!FechaEntregaIII), "  /  /    ", rstChofer!FechaEntregaIII)
        
        ComentarioI.Text = IIf(IsNull(rstChofer!ComentarioI), "", rstChofer!ComentarioI)
        ComentarioII.Text = IIf(IsNull(rstChofer!ComentarioII), "", rstChofer!ComentarioII)
        ComentarioIII.Text = IIf(IsNull(rstChofer!ComentarioIII), "", rstChofer!ComentarioIII)
        
        ComentarioI.Text = Trim(ComentarioI.Text)
        ComentarioII.Text = Trim(ComentarioII.Text)
        ComentarioIII.Text = Trim(ComentarioIII.Text)
        
        rstChofer.Close
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Proveedor"
    ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
    End If
    
End Sub

Private Sub Acepta_Click()
    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Chofer.Codigo, Chofer.Descripcion, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Chofer Chofer, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Chofer.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Chofer.Codigo >= " + Desde.Text + " AND " _
                + "Chofer.Codigo <= " + Hasta.Text
    
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Chofer.Codigo} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Chofer.Codigo} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(Codigo.Text) <> 0 Then
    
        ZOrdFechaVtoI = Right$(FechaVtoI.Text, 4) + Mid$(FechaVtoI.Text, 4, 2) + Left$(FechaVtoI.Text, 2)
        ZOrdFechaVtoII = Right$(FechaVtoII.Text, 4) + Mid$(FechaVtoII.Text, 4, 2) + Left$(FechaVtoII.Text, 2)
        ZOrdFechaVtoIII = Right$(FechaVtoIII.Text, 4) + Mid$(FechaVtoIII.Text, 4, 2) + Left$(FechaVtoIII.Text, 2)
    
        ZOrdFechaEntregaI = Right$(FechaEntregaI.Text, 4) + Mid$(FechaEntregaI.Text, 4, 2) + Left$(FechaEntregaI.Text, 2)
        ZOrdFechaEntregaII = Right$(FechaEntregaII.Text, 4) + Mid$(FechaEntregaII.Text, 4, 2) + Left$(FechaEntregaII.Text, 2)
        ZOrdFechaEntregaIII = Right$(FechaEntregaIII.Text, 4) + Mid$(FechaEntregaIII.Text, 4, 2) + Left$(FechaEntregaIII.Text, 2)
        
        XAplicaIII = Str$(AplicaIII.Value)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Chofer"
        ZSql = ZSql + " Where Chofer.Codigo = " + "'" + Codigo.Text + "'"
        spChofer = ZSql
        Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
        If rstChofer.RecordCount > 0 Then
            rstChofer.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Chofer SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " AplicaIII = " + "'" + XAplicaIII + "',"
            ZSql = ZSql + " FechaVtoI = " + "'" + FechaVtoI.Text + "',"
            ZSql = ZSql + " FechaVtoII = " + "'" + FechaVtoII.Text + "',"
            ZSql = ZSql + " FechaVtoIII = " + "'" + FechaVtoIII.Text + "',"
            ZSql = ZSql + " OrdFechaVtoI = " + "'" + ZOrdFechaVtoI + "',"
            ZSql = ZSql + " OrdFechaVtoII = " + "'" + ZOrdFechaVtoII + "',"
            ZSql = ZSql + " OrdFechaVtoIII = " + "'" + ZOrdFechaVtoIII + "',"
            ZSql = ZSql + " FechaEntregaI = " + "'" + FechaEntregaI.Text + "',"
            ZSql = ZSql + " FechaEntregaII = " + "'" + FechaEntregaII.Text + "',"
            ZSql = ZSql + " FechaEntregaIII = " + "'" + FechaEntregaIII.Text + "',"
            ZSql = ZSql + " OrdFechaEntregaI = " + "'" + ZOrdFechaEntregaI + "',"
            ZSql = ZSql + " OrdFechaEntregaII = " + "'" + ZOrdFechaEntregaII + "',"
            ZSql = ZSql + " OrdFechaEntregaIII = " + "'" + ZOrdFechaEntregaIII + "',"
            ZSql = ZSql + " ComentarioI = " + "'" + ComentarioI.Text + "',"
            ZSql = ZSql + " ComentarioII = " + "'" + ComentarioII.Text + "',"
            ZSql = ZSql + " ComentarioIII = " + "'" + ComentarioIII.Text + "',"
            ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spChofer = ZSql
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Chofer ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "AplicaIII ,"
            ZSql = ZSql + "FechaVtoI ,"
            ZSql = ZSql + "FechaVtoII ,"
            ZSql = ZSql + "FechaVtoIII ,"
            ZSql = ZSql + "OrdFechaVtoI ,"
            ZSql = ZSql + "OrdFechaVtoII ,"
            ZSql = ZSql + "OrdFechaVtoIII ,"
            ZSql = ZSql + "FechaEntregaI ,"
            ZSql = ZSql + "FechaEntregaII ,"
            ZSql = ZSql + "FechaEntregaIII ,"
            ZSql = ZSql + "OrdFechaEntregaI ,"
            ZSql = ZSql + "OrdFechaEntregaII ,"
            ZSql = ZSql + "OrdFechaEntregaIII ,"
            ZSql = ZSql + "ComentarioI ,"
            ZSql = ZSql + "ComentarioII ,"
            ZSql = ZSql + "ComentarioIII ,"
            ZSql = ZSql + "CodigoEmpresa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + XAplicaIII + "',"
            ZSql = ZSql + "'" + FechaVtoI.Text + "',"
            ZSql = ZSql + "'" + FechaVtoII.Text + "',"
            ZSql = ZSql + "'" + FechaVtoIII.Text + "',"
            ZSql = ZSql + "'" + ZOrdFechaVtoI + "',"
            ZSql = ZSql + "'" + ZOrdFechaVtoII + "',"
            ZSql = ZSql + "'" + ZOrdFechaVtoIII + "',"
            ZSql = ZSql + "'" + FechaEntregaI.Text + "',"
            ZSql = ZSql + "'" + FechaEntregaII.Text + "',"
            ZSql = ZSql + "'" + FechaEntregaIII.Text + "',"
            ZSql = ZSql + "'" + ZOrdFechaEntregaI + "',"
            ZSql = ZSql + "'" + ZOrdFechaEntregaII + "',"
            ZSql = ZSql + "'" + ZOrdFechaEntregaIII + "',"
            ZSql = ZSql + "'" + ComentarioI.Text + "',"
            ZSql = ZSql + "'" + ComentarioII.Text + "',"
            ZSql = ZSql + "'" + ComentarioIII.Text + "',"
            ZSql = ZSql + "'" + "1" + "')"
            spChofer = ZSql
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM Chofer"
        Sql3 = " Where Chofer.Codigo = " + "'" + Codigo.Text + "'"
        spChofer = Sql1 + Sql2 + Sql3
        Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
        If rstChofer.RecordCount > 0 Then
            rstChofer.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE Chofer"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spChofer = Sql1 + Sql2
                Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    AplicaIII.Value = False
    
    FechaVtoI.Text = "  /  /    "
    FechaVtoII.Text = "  /  /    "
    FechaVtoIII.Text = "  /  /    "
    
    FechaEntregaI.Text = "  /  /    "
    FechaEntregaII.Text = "  /  /    "
    FechaEntregaIII.Text = "  /  /    "
    
    ComentarioI.Text = ""
    ComentarioII.Text = ""
    ComentarioIII.Text = ""

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Chofer"
    spChofer = Sql1 + Sql2
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        rstChofer.MoveLast
        ZCodigo = IIf(IsNull(rstChofer!CodigoMayor), "0", rstChofer!CodigoMayor)
        Codigo.Text = ZCodigo + 1
        rstChofer.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgChoferes.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    Sql1 = "Select *"
    Sql2 = " FROM Chofer"
    Sql3 = " Where Chofer.Codigo < " + "'" + Codigo.Text + "'"
    spChofer = Sql1 + Sql2 + Sql3
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        With rstChofer
            .MoveLast
            Codigo.Text = rstChofer!Codigo
        End With
        rstChofer.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Choferes")
    End If
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Proveedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM Proveedor"
        ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = rstProveedor!Nombre
            rstProveedor.Close
            FechaVtoI.SetFocus
                Else
            DesProveedor.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
End Sub






Private Sub FechaVtoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaVtoI.Text, Auxi)
        If Auxi = "S" Or FechaVtoI.Text = "  /  /    " Then
            FechaEntregaI.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaVtoI.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaEntregaI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntregaI.Text, Auxi)
        If Auxi = "S" Or FechaEntregaI.Text = "  /  /    " Then
            ComentarioI.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEntregaI.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComentarioI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaVtoII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioI.Text = ""
    End If
End Sub






Private Sub FechaVtoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaVtoII.Text, Auxi)
        If Auxi = "S" Or FechaVtoII.Text = "  /  /    " Then
            FechaEntregaII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaVtoII.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaEntregaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntregaII.Text, Auxi)
        If Auxi = "S" Or FechaEntregaII.Text = "  /  /    " Then
            ComentarioII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEntregaII.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComentarioII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaVtoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioII.Text = ""
    End If
End Sub






Private Sub FechaVtoIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaVtoIII.Text, Auxi)
        If Auxi = "S" Or FechaVtoIII.Text = "  /  /    " Then
            FechaEntregaIII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaVtoIII.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaEntregaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntregaIII.Text, Auxi)
        If Auxi = "S" Or FechaEntregaIII.Text = "  /  /    " Then
            ComentarioIII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEntregaIII.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComentarioIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioIII.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Chofer"
            Sql3 = " Where Chofer.Codigo = " + "'" + Codigo.Text + "'"
            spChofer = Sql1 + Sql2 + Sql3
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                rstChofer.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Choferes"
     Opcion.AddItem "Proveedores"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Chofer"
            Sql3 = " Order by Chofer.Codigo"
            spChofer = Sql1 + Sql2 + Sql3
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                With rstChofer
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstChofer!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstChofer!Descripcion
                            IngresaItem = rstChofer!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstChofer.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor.Proveedor"
            spProveedor = Sql1 + Sql2 + Sql3
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstProveedor!Proveedor
                            Pantalla.Col = 2
                            Pantalla.Text = rstProveedor!Nombre
                            IngresaItem = rstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case 1
            Indice = Pantalla.Row - 1
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_Keypress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    Sql1 = "Select Min(Codigo) as [CodigoMenor]"
    Sql2 = " FROM Chofer"
    spChofer = Sql1 + Sql2
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        rstChofer.MoveFirst
        Codigo.Text = rstChofer!CodigoMenor
        rstChofer.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Chofer"
    spChofer = Sql1 + Sql2
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        rstChofer.MoveLast
        Codigo.Text = rstChofer!CodigoMayor
        rstChofer.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Chofer"
    Sql3 = " Where Chofer.Codigo > " + "'" + Codigo.Text + "'"
    spChofer = Sql1 + Sql2 + Sql3
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        With rstChofer
            .MoveFirst
            Codigo.Text = rstChofer!Codigo
        End With
        rstChofer.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Choferes")
    End If

End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    AplicaIII.Value = False
    
    FechaVtoI.Text = "  /  /    "
    FechaVtoII.Text = "  /  /    "
    FechaVtoIII.Text = "  /  /    "
    
    FechaEntregaI.Text = "  /  /    "
    FechaEntregaII.Text = "  /  /    "
    FechaEntregaIII.Text = "  /  /    "
    
    ComentarioI.Text = ""
    ComentarioII.Text = ""
    ComentarioIII.Text = ""
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Chofer"
    spChofer = Sql1 + Sql2
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        rstChofer.MoveLast
        ZCodigo = IIf(IsNull(rstChofer!CodigoMayor), "0", rstChofer!CodigoMayor)
        Codigo.Text = ZCodigo + 1
        rstChofer.Close
    End If
    
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Chofer"
            Sql3 = " Order by Chofer.Codigo"
            spChofer = Sql1 + Sql2 + Sql3
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                With rstChofer
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstChofer!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstChofer!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstChofer!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstChofer!Descripcion
                                    IngresaItem = rstChofer!Codigo
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
                rstChofer.Close
            End If
                
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor.Proveedor"
            spProveedor = Sql1 + Sql2 + Sql3
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstProveedor!Nombre) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstProveedor!Nombre, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstProveedor!Proveedor
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstProveedor!Nombre
                                    IngresaItem = rstProveedor!Proveedor
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
                rstProveedor.Close
            End If
                
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Codigo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Choferes"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Camiones"
    Opcion.AddItem "Proveedor"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub


Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0, 1
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
    End Select
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    Select Case XIndice
        Case 0, 1
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Codigo"
                        Pantalla.ColWidth(Ciclo) = 1500
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 6000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
        Case Else
            
    End Select
    
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





