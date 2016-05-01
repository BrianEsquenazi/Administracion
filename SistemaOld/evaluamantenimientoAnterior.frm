VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEvaluaMantenimientoAnterior 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Evaluacion de Proveedores de Mantenimiento"
   ClientHeight    =   8250
   ClientLeft      =   285
   ClientTop       =   300
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11430
   Begin VB.CommandButton Baja 
      Caption         =   "Inhabilitacion"
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
      Left            =   8520
      TabIndex        =   41
      Top             =   5520
      Width           =   2775
   End
   Begin VB.ComboBox Califica5 
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
      Left            =   9600
      TabIndex        =   39
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox Califica4 
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
      Left            =   9600
      TabIndex        =   38
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox Califica3 
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
      Left            =   9600
      TabIndex        =   37
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox Califica2 
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
      Left            =   9600
      TabIndex        =   36
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox Califica1 
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
      Left            =   9600
      TabIndex        =   35
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Evaluador 
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
      Left            =   5400
      MaxLength       =   50
      TabIndex        =   22
      Top             =   480
      Width           =   5775
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   8520
      TabIndex        =   19
      Top             =   4560
      Width           =   2775
      Begin VB.TextBox Promedio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Calificacion Proveedor"
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
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
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   18
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Mes 
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
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   615
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
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1455
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
      TabIndex        =   6
      Top             =   7560
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
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   4560
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "evaluamantenimientoAnterior.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "evaluamantenimientoAnterior.frx":030A
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "evaluamantenimientoAnterior.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "evaluamantenimientoAnterior.frx":0A56
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "evaluamantenimientoAnterior.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "evaluamantenimientoAnterior.frx":11A2
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "evaluamantenimientoAnterior.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "evaluamantenimientoAnterior.frx":18EE
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   8175
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Camiones.rpt"
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
      Left            =   9960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
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
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2160
      TabIndex        =   24
      Top             =   840
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
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   5400
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
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
   Begin VB.TextBox Observaciones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   40
      Top             =   5520
      Width           =   8295
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calificacion"
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
      Index           =   21
      Left            =   9600
      TabIndex        =   34
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Parametro 
      Caption         =   "El equipamiento intervenido queso en correcto estado de funcionameinto"
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
      Height          =   495
      Index           =   19
      Left            =   4680
      TabIndex        =   33
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Parametro 
      Caption         =   "Trabajo realizado segun las reglas de arte"
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
      Index           =   18
      Left            =   4680
      TabIndex        =   32
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Parametro 
      Caption         =   "Trabajo cumplido en los plazos solicitados por Surfactan"
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
      Index           =   17
      Left            =   4680
      TabIndex        =   31
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Parametro 
      Caption         =   "Presenta certificado de ART cin clausula de no repeticion"
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
      Index           =   14
      Left            =   4680
      TabIndex        =   30
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Parametro 
      Caption         =   "Recuperación de Envases/ Entrega de Certificados /  Remitos conformados/ otros"
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
      Height          =   495
      Index           =   2
      Left            =   4680
      TabIndex        =   29
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Criterio"
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
      Left            =   4680
      TabIndex        =   28
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblLabels 
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
      Index           =   16
      Left            =   3840
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblLabels 
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
      Index           =   15
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Evaluador"
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
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Periodo"
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
      TabIndex        =   17
      Top             =   480
      Width           =   1455
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
      TabIndex        =   15
      Top             =   120
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
      TabIndex        =   14
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Parametro 
      Caption         =   "Cumplimientos de condiciones de seguridad y Medio Ambiente"
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
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Parametro 
      Caption         =   "ART"
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
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "Rapidez de Respuesta"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Parametro 
      Caption         =   "Prolijidad en la ejecucion de las tareas"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Parametro 
      Caption         =   "Trabajo Correctamente Realizado"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3600
      MouseIcon       =   "evaluamantenimientoAnterior.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":203A
      ToolTipText     =   "Impresion "
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1920
      MouseIcon       =   "evaluamantenimientoAnterior.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":2B86
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   240
      MouseIcon       =   "evaluamantenimientoAnterior.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":36D2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1080
      MouseIcon       =   "evaluamantenimientoAnterior.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":421E
      ToolTipText     =   "Elimina el Registro"
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4440
      MouseIcon       =   "evaluamantenimientoAnterior.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":4D6A
      ToolTipText     =   "Salida"
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2760
      MouseIcon       =   "evaluamantenimientoAnterior.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "evaluamantenimientoAnterior.frx":58B6
      ToolTipText     =   "Consulta de Datos"
      Top             =   4680
      Width           =   480
   End
End
Attribute VB_Name = "PrgEvaluaMantenimientoAnterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstEvaluaI As Recordset
Dim spEvaluaI As String

Dim ZProveedor As String
Dim ZMes As String
Dim ZAno As String
Dim ZClave As String

Sub Imprime_Datos()

    ZProveedor = Proveedor.Text
    ZMes = Mes.Text
    ZAno = Ano.Text
    
    Call Ceros(ZProveedor, 11)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZClave = ZProveedor + ZMes + ZAno

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 2"
    
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
    
        Evaluador.Text = Trim(rstEvaluaI!Evaluador)
        Observaciones.Text = Trim(rstEvaluaI!Observaciones)
        
        Fecha.Text = rstEvaluaI!Fecha
        Vencimiento.Text = rstEvaluaI!Vencimiento
        
        Califica1.ListIndex = rstEvaluaI!Califica1
        Califica2.ListIndex = rstEvaluaI!Califica2
        Califica3.ListIndex = rstEvaluaI!Califica3
        Califica4.ListIndex = rstEvaluaI!Califica4
        Califica5.ListIndex = rstEvaluaI!Califica5
        
        Promedio.Text = rstEvaluaI!Promedio
        
        rstEvaluaI.Close
        
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

Private Sub cmdAdd_Click()
    If Val(Proveedor.Text) <> 0 And Val(Mes.Text) <> 0 And Val(Ano.Text) <> 0 Then
    
        If Califica1.ListIndex = 0 Or Califica2.ListIndex = 0 Or Califica3.ListIndex = 0 Or Califica4.ListIndex = 0 Or Califica5.ListIndex = 0 Then
            m$ = "Se debe calificar todos los items"
            A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
            Exit Sub
        End If

        ZProveedor = Proveedor.Text
        ZMes = Mes.Text
        ZAno = Ano.Text
        
        Call Ceros(ZProveedor, 11)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        ZPeriodo = ZAno + ZMes
        ZClave = ZProveedor + ZMes + ZAno
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EvaluaI"
        ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
        ZSql = ZSql + " and EvaluaI.Tipo = 2"
        spEvaluaI = ZSql
        Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
        If rstEvaluaI.RecordCount > 0 Then
        
            rstEvaluaI.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE EvaluaI SET "
            ZSql = ZSql + " Clave = " + "'" + ZClave + "',"
            ZSql = ZSql + " Tipo = " + "'" + "2" + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Mes = " + "'" + Mes.Text + "',"
            ZSql = ZSql + " Ano = " + "'" + Ano.Text + "',"
            ZSql = ZSql + " Periodo = " + "'" + ZPeriodo + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " Vencimiento = " + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + " Evaluador = " + "'" + Evaluador.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Califica1 = " + "'" + Str$(Califica1.ListIndex) + "',"
            ZSql = ZSql + " Califica2 = " + "'" + Str$(Califica2.ListIndex) + "',"
            ZSql = ZSql + " Califica3 = " + "'" + Str$(Califica3.ListIndex) + "',"
            ZSql = ZSql + " Califica4 = " + "'" + Str$(Califica4.ListIndex) + "',"
            ZSql = ZSql + " Califica5 = " + "'" + Str$(Califica5.ListIndex) + "',"
            ZSql = ZSql + " Parametro1 = " + "'" + Left$(Parametro(3).Caption, 50) + "',"
            ZSql = ZSql + " Parametro2 = " + "'" + Left$(Parametro(4).Caption, 50) + "',"
            ZSql = ZSql + " Parametro3 = " + "'" + Left$(Parametro(5).Caption, 50) + "',"
            ZSql = ZSql + " Parametro4 = " + "'" + Left$(Parametro(6).Caption, 50) + "',"
            ZSql = ZSql + " Parametro5 = " + "'" + Left$(Parametro(7).Caption, 50) + "',"
            ZSql = ZSql + " Criterio1 = " + "'" + Left$(Parametro(19).Caption, 50) + "',"
            ZSql = ZSql + " Criterio2 = " + "'" + Left$(Parametro(18).Caption, 50) + "',"
            ZSql = ZSql + " Criterio3 = " + "'" + Left$(Parametro(17).Caption, 50) + "',"
            ZSql = ZSql + " Criterio4 = " + "'" + Left$(Parametro(14).Caption, 50) + "',"
            ZSql = ZSql + " Criterio5 = " + "'" + Left$(Parametro(2).Caption, 50) + "',"
            ZSql = ZSql + " Promedio = " + "'" + Promedio.Text + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            aa = Left$(Parametro(3).Caption, 50)
                
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO EvaluaI ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Mes ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Vencimiento ,"
            ZSql = ZSql + "Evaluador ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Califica1 ,"
            ZSql = ZSql + "Califica2 ,"
            ZSql = ZSql + "Califica3 ,"
            ZSql = ZSql + "Califica4 ,"
            ZSql = ZSql + "Califica5 ,"
            ZSql = ZSql + "Parametro1 ,"
            ZSql = ZSql + "Parametro2 ,"
            ZSql = ZSql + "Parametro3 ,"
            ZSql = ZSql + "Parametro4 ,"
            ZSql = ZSql + "Parametro5 ,"
            ZSql = ZSql + "Criterio1 ,"
            ZSql = ZSql + "Criterio2 ,"
            ZSql = ZSql + "Criterio3 ,"
            ZSql = ZSql + "Criterio4 ,"
            ZSql = ZSql + "Criterio5 ,"
            ZSql = ZSql + "Promedio )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Mes.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + ZPeriodo + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + "'" + Evaluador.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Str$(Califica1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica5.ListIndex) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(3).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(4).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(5).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(6).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(7).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(19).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(18).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(17).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(14).Caption, 50) + "',"
            ZSql = ZSql + "'" + Left$(Parametro(7).Caption, 50) + "',"
            ZSql = ZSql + "'" + Promedio.Text + "')"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        Call CmdLimpiar_Click
        Proveedor.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Val(Proveedor.Text) <> 0 And Val(Mes.Text) <> 0 And Val(Ano.Text) <> 0 Then

        ZProveedor = Proveedor.Text
        ZMes = Mes.Text
        ZAno = Ano.Text
        
        Call Ceros(ZProveedor, 11)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        ZClave = ZProveedor + ZMes + ZAno
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EvaluaI"
        ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
        ZSql = ZSql + " and EvaluaI.Tipo = 2"
        spEvaluaI = ZSql
        Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
        If rstEvaluaI.RecordCount > 0 Then
            rstEvaluaI.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE EvaluaI"
                ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                ZSql = ZSql + " and EvaluaI.Tipo = 2"
                spEvaluaI = ZSql
                Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    
    Proveedor.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Mes.Text = ""
    Ano.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Evaluador.Text = ""
    Observaciones.Text = ""
    Califica1.ListIndex = 0
    Califica2.ListIndex = 0
    Califica3.ListIndex = 0
    Califica4.ListIndex = 0
    Califica5.ListIndex = 0
    Promedio.Text = ""
    
    Proveedor.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgEvaluaMantenimiento.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()

    ZProveedor = Proveedor.Text
    ZMes = Mes.Text
    ZAno = Ano.Text
    
    Call Ceros(ZProveedor, 11)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZClave = ZProveedor + ZMes + ZAno

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave < " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 2"
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
        With rstEvaluaI
            .MoveLast
            Proveedor.Text = rstEvaluaI!Proveedor
            Ano.Text = rstEvaluaI!Ano
            Mes.Text = rstEvaluaI!Mes
        End With
        rstEvaluaI.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
    End If
    
End Sub

Private Sub Dominio1_Click()
End Sub

Private Sub Evaluador_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Evaluador.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
End Sub

Private Sub Punto35_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto35.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Calcula_Promedio()

    Promedio.Text = "10"
    Promedio.Text = Pusing("###,###.##", Promedio.Text)

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
            Mes.SetFocus
                Else
            DesProveedor.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Mes.Text) >= 1 And Val(Mes.Text) <= 12 Then
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Mes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) >= 1990 And Val(Mes.Text) <= 2100 Then

            ZProveedor = Proveedor.Text
            ZMes = Mes.Text
            ZAno = Ano.Text
            
            Call Ceros(ZProveedor, 11)
            Call Ceros(ZMes, 2)
            Call Ceros(ZAno, 4)
            
            ZClave = ZProveedor + ZMes + ZAno
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM EvaluaI"
            ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            If rstEvaluaI.RecordCount > 0 Then
            
                rstEvaluaI.Close
                Call Imprime_Datos
                
                    Else
                    
                ZProveedor = Proveedor.Text
                ZMes = Mes.Text
                ZAno = Ano.Text
                CmdLimpiar_Click
                Proveedor.Text = ZProveedor
                Mes.Text = ZMes
                Ano.Text = ZAno
                
            End If
            
            Evaluador.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

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

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_Keypress(13)
            
        Case Else
    End Select
    
End Sub


Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave > " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 2"
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
        With rstEvaluaI
            .MoveFirst
            Proveedor.Text = rstEvaluaI!Proveedor
            Ano.Text = rstEvaluaI!Ano
            Mes.Text = rstEvaluaI!Mes
        End With
        rstEvaluaI.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
    End If

End Sub

Sub Form_Load()

    Califica1.Clear
    Califica2.Clear
    Califica3.Clear
    Califica4.Clear
    Califica5.Clear
    
    Califica1.AddItem ""
    Califica1.AddItem "Cumple"
    Califica1.AddItem "No Cumple"
    
    Califica2.AddItem ""
    Califica2.AddItem "Cumple"
    Califica2.AddItem "No Cumple"
    
    Califica3.AddItem ""
    Califica3.AddItem "Cumple"
    Califica3.AddItem "Parcial"
    Califica3.AddItem "No Cumple"
    Califica3.AddItem "No Aplica"
    
    Califica4.AddItem ""
    Califica4.AddItem "Cumple"
    Califica4.AddItem "Parcial"
    Califica4.AddItem "No Cumple"
    Califica4.AddItem "No Aplica"
    
    Califica5.AddItem ""
    Califica5.AddItem "Cumple"
    Califica5.AddItem "Parcial"
    Califica5.AddItem "No Cumple"
    Califica5.AddItem "No Aplica"
    
    Califica1.ListIndex = 0
    Califica2.ListIndex = 0
    Califica3.ListIndex = 0
    Califica4.ListIndex = 0
    Califica5.ListIndex = 0

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Mes.Text = ""
    Ano.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Evaluador.Text = ""
    Observaciones.Text = ""
    Promedio.Text = ""
    
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

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedor"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
    End Select
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    Select Case XIndice
        Case 0
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





