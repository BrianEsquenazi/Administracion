VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEvaluaMantenimiento 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Evaluacion de Proveedores de Mantenimiento"
   ClientHeight    =   8250
   ClientLeft      =   285
   ClientTop       =   300
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11430
   Begin VB.TextBox ObservacionesProve 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   63
      Top             =   5520
      Width           =   5535
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
      Height          =   2415
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   62
      Top             =   5520
      Width           =   5535
   End
   Begin VB.TextBox DesPromedio11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   60
      Top             =   3600
      Width           =   1350
   End
   Begin VB.TextBox DesPromedio33 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   9960
      MaxLength       =   20
      TabIndex        =   59
      Top             =   3600
      Width           =   1350
   End
   Begin VB.TextBox DesPromedio22 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   58
      Top             =   3600
      Width           =   1350
   End
   Begin VB.TextBox Promedio22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Promedio33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   9960
      MaxLength       =   20
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Promedio11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   53
      Top             =   3600
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.ComboBox Sector2 
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
      Left            =   8520
      TabIndex        =   51
      Top             =   1320
      Width           =   1350
   End
   Begin VB.ComboBox Sector3 
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
      Left            =   9960
      TabIndex        =   50
      Top             =   1320
      Width           =   1350
   End
   Begin VB.ComboBox Sector1 
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
      Left            =   7080
      TabIndex        =   49
      Top             =   1320
      Width           =   1350
   End
   Begin VB.ComboBox Califica11 
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
      Left            =   7080
      TabIndex        =   45
      Top             =   1720
      Width           =   1350
   End
   Begin VB.ComboBox Califica12 
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
      Left            =   7080
      TabIndex        =   44
      Top             =   2160
      Width           =   1350
   End
   Begin VB.ComboBox Califica13 
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
      Left            =   7080
      TabIndex        =   43
      Top             =   2640
      Width           =   1350
   End
   Begin VB.ComboBox Califica14 
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
      Left            =   7080
      TabIndex        =   42
      Top             =   3120
      Width           =   1350
   End
   Begin VB.ComboBox Califica24 
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
      Left            =   8520
      TabIndex        =   41
      Top             =   3120
      Width           =   1350
   End
   Begin VB.ComboBox Califica23 
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
      Left            =   8520
      TabIndex        =   40
      Top             =   2640
      Width           =   1350
   End
   Begin VB.ComboBox Califica22 
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
      Left            =   8520
      TabIndex        =   39
      Top             =   2160
      Width           =   1350
   End
   Begin VB.ComboBox Califica21 
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
      Left            =   8520
      TabIndex        =   38
      Top             =   1720
      Width           =   1350
   End
   Begin VB.ComboBox Califica34 
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
      Left            =   9960
      TabIndex        =   37
      Top             =   3120
      Width           =   1350
   End
   Begin VB.ComboBox Califica33 
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
      Left            =   9960
      TabIndex        =   36
      Top             =   2640
      Width           =   1350
   End
   Begin VB.ComboBox Califica32 
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
      Left            =   9960
      TabIndex        =   35
      Top             =   2160
      Width           =   1350
   End
   Begin VB.ComboBox Califica31 
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
      Left            =   9960
      TabIndex        =   34
      Top             =   1720
      Width           =   1350
   End
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
      Height          =   495
      Left            =   7080
      TabIndex        =   33
      Top             =   3960
      Width           =   4215
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
      Left            =   11040
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   390
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
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   7080
      TabIndex        =   19
      Top             =   4440
      Width           =   4215
      Begin VB.TextBox DesPromedio 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   20
         TabIndex        =   61
         Top             =   240
         Width           =   1695
      End
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
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   120
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
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   600
         MouseIcon       =   "evaluaMantenimiento.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "evaluaMantenimiento.frx":030A
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1200
         MouseIcon       =   "evaluaMantenimiento.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "evaluaMantenimiento.frx":0A56
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   1800
         MouseIcon       =   "evaluaMantenimiento.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "evaluaMantenimiento.frx":11A2
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   120
         MouseIcon       =   "evaluaMantenimiento.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "evaluaMantenimiento.frx":18EE
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
      Width           =   7935
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
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4048
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2160
      TabIndex        =   24
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
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   5400
      TabIndex        =   56
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
   Begin VB.Label Cartel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor Inhabiloitado"
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
      TabIndex        =   66
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observaciones de la Evaluacion"
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
      Left            =   5760
      TabIndex        =   65
      Top             =   5160
      Width           =   5535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observaciones del Proveedor"
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
      TabIndex        =   64
      Top             =   5160
      Width           =   5535
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
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sector"
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
      Index           =   4
      Left            =   6120
      TabIndex        =   52
      Top             =   1320
      Width           =   855
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
      Left            =   7080
      TabIndex        =   48
      Top             =   960
      Width           =   1350
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
      Index           =   2
      Left            =   8520
      TabIndex        =   47
      Top             =   960
      Width           =   1350
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
      Index           =   3
      Left            =   9960
      TabIndex        =   46
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Parametro 
      Caption         =   "Equipo / Instalacion queda en correcto funcionamiento"
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
      Left            =   3480
      TabIndex        =   31
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "Cumple con el trabajo y entrega de documentacion en los plazos solicitados por surfactan"
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
      Left            =   3480
      TabIndex        =   30
      Top             =   2160
      Width           =   3255
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
      Index           =   17
      Left            =   3480
      TabIndex        =   29
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "El proveedor respeta los lineamientos de Surfactan - ART"
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
      Left            =   3480
      TabIndex        =   28
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Parametro 
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
      Left            =   11160
      TabIndex        =   27
      Top             =   -120
      Width           =   375
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
      Left            =   3480
      TabIndex        =   26
      Top             =   1320
      Width           =   2535
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Width           =   3255
   End
   Begin VB.Label Parametro 
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
      Left            =   11160
      TabIndex        =   11
      Top             =   0
      Width           =   255
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
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "Prolijidad en Ejecucion del trabajo"
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
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "Rapidez de Respuesta (Plazos)"
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
      Width           =   3255
   End
   Begin VB.Label Parametro 
      Caption         =   "Cumplimiento de Requisitos Basicos del Servicio"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3120
      MouseIcon       =   "evaluaMantenimiento.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":203A
      ToolTipText     =   "Impresion "
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1680
      MouseIcon       =   "evaluaMantenimiento.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":2B86
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   240
      MouseIcon       =   "evaluaMantenimiento.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":36D2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   960
      MouseIcon       =   "evaluaMantenimiento.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":421E
      ToolTipText     =   "Elimina el Registro"
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "evaluaMantenimiento.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":4D6A
      ToolTipText     =   "Salida"
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2400
      MouseIcon       =   "evaluaMantenimiento.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "evaluaMantenimiento.frx":58B6
      ToolTipText     =   "Consulta de Datos"
      Top             =   3960
      Width           =   480
   End
End
Attribute VB_Name = "PrgEvaluaMantenimiento"
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

Dim ZCalifica11 As Integer
Dim ZCalifica12  As Integer
Dim ZCalifica13  As Integer
Dim ZCalifica14  As Integer

Dim ZCalifica21   As Integer
Dim ZCalifica22   As Integer
Dim ZCalifica23   As Integer
Dim ZCalifica24   As Integer

Dim ZCalifica31   As Integer
Dim ZCalifica32   As Integer
Dim ZCalifica33   As Integer
Dim ZCalifica34   As Integer


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
    ZSql = ZSql + " and EvaluaI.Tipo = 4"
    
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
    
        Evaluador.Text = Trim(rstEvaluaI!Evaluador)
        Observaciones.Text = IIf(IsNull(rstEvaluaI!Observaciones), "", rstEvaluaI!Observaciones)
        
        Fecha.Text = rstEvaluaI!Fecha
        Vencimiento.Text = rstEvaluaI!Vencimiento
        
        Sector1.ListIndex = rstEvaluaI!Sector1
        Sector2.ListIndex = rstEvaluaI!Sector2
        Sector3.ListIndex = rstEvaluaI!Sector3
        
        Califica11.ListIndex = rstEvaluaI!Califica11
        Califica12.ListIndex = rstEvaluaI!Califica12
        Califica13.ListIndex = rstEvaluaI!Califica13
        Califica14.ListIndex = rstEvaluaI!Califica14
        
        Califica21.ListIndex = rstEvaluaI!Califica21
        Califica22.ListIndex = rstEvaluaI!Califica22
        Califica23.ListIndex = rstEvaluaI!Califica23
        Califica24.ListIndex = rstEvaluaI!Califica24
        
        Califica31.ListIndex = rstEvaluaI!Califica31
        Califica32.ListIndex = rstEvaluaI!Califica32
        Califica33.ListIndex = rstEvaluaI!Califica33
        Califica34.ListIndex = rstEvaluaI!Califica34
        
        Promedio.Text = rstEvaluaI!Promedio
        
        Promedio11.Text = rstEvaluaI!Promedio11
        Promedio22.Text = rstEvaluaI!Promedio22
        Promedio33.Text = rstEvaluaI!Promedio33
        
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
        ObservacionesProve.Text = IIf(IsNull(rstProveedor!ObservacionesII), "", rstProveedor!ObservacionesII)
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
        ObservacionesProve.Text = ""
    End If
    
    Call Calcula_Promedio

End Sub

Private Sub Baja_Click()
    T$ = "Inhabilitacion de Proveedores"
    m$ = "Desea Inhabilitar al proveedor "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        EmpresaReal = WEmpresa
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0009"
        txtOdbc = "Empresa09"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = EmpresaReal
        txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        m$ = "El Proveedor ha sido inhabilitado"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        
        Cartel.Visible = True
        
        
        
    End If
End Sub

Private Sub cmdAdd_Click()
    If Val(Proveedor.Text) <> 0 And Val(Mes.Text) <> 0 And Val(Ano.Text) <> 0 Then

        If Sector1.ListIndex <> 0 Then
            If Califica11.ListIndex = 0 Or Califica12.ListIndex = 0 Or Califica13.ListIndex = 0 Or Califica14.ListIndex = 0 Then
                m$ = "Se debe calificar todos los items"
                A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
                Exit Sub
            End If
        End If

        If Sector2.ListIndex <> 0 Then
            If Califica21.ListIndex = 0 Or Califica22.ListIndex = 0 Or Califica23.ListIndex = 0 Or Califica24.ListIndex = 0 Then
                m$ = "Se debe calificar todos los items"
                A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
                Exit Sub
            End If
        End If

        If Sector3.ListIndex <> 0 Then
            If Califica31.ListIndex = 0 Or Califica32.ListIndex = 0 Or Califica33.ListIndex = 0 Or Califica34.ListIndex = 0 Then
                m$ = "Se debe calificar todos los items"
                A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
                Exit Sub
            End If
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
        ZSql = ZSql + " and EvaluaI.Tipo = 4"
        spEvaluaI = ZSql
        Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
        If rstEvaluaI.RecordCount > 0 Then
        
            rstEvaluaI.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE EvaluaI SET "
            ZSql = ZSql + " Clave = " + "'" + ZClave + "',"
            ZSql = ZSql + " Tipo = " + "'" + "4" + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Mes = " + "'" + Mes.Text + "',"
            ZSql = ZSql + " Ano = " + "'" + Ano.Text + "',"
            ZSql = ZSql + " Periodo = " + "'" + ZPeriodo + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " Vencimiento = " + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + " Evaluador = " + "'" + Evaluador.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Sector1 = " + "'" + Str$(Sector1.ListIndex) + "',"
            ZSql = ZSql + " Sector2 = " + "'" + Str$(Sector2.ListIndex) + "',"
            ZSql = ZSql + " Sector3 = " + "'" + Str$(Sector3.ListIndex) + "',"
            ZSql = ZSql + " DesSector1 = " + "'" + Sector1.Text + "',"
            ZSql = ZSql + " DesSector2 = " + "'" + Sector2.Text + "',"
            ZSql = ZSql + " DesSector3 = " + "'" + Sector3.Text + "',"
            ZSql = ZSql + " Califica11 = " + "'" + Str$(Califica11.ListIndex) + "',"
            ZSql = ZSql + " Califica12 = " + "'" + Str$(Califica12.ListIndex) + "',"
            ZSql = ZSql + " Califica13 = " + "'" + Str$(Califica13.ListIndex) + "',"
            ZSql = ZSql + " Califica14 = " + "'" + Str$(Califica14.ListIndex) + "',"
            ZSql = ZSql + " Califica15 = " + "'" + "0" + "',"
            ZSql = ZSql + " Califica21 = " + "'" + Str$(Califica21.ListIndex) + "',"
            ZSql = ZSql + " Califica22 = " + "'" + Str$(Califica22.ListIndex) + "',"
            ZSql = ZSql + " Califica23 = " + "'" + Str$(Califica23.ListIndex) + "',"
            ZSql = ZSql + " Califica24 = " + "'" + Str$(Califica24.ListIndex) + "',"
            ZSql = ZSql + " Califica25 = " + "'" + "0" + "',"
            ZSql = ZSql + " Califica31 = " + "'" + Str$(Califica31.ListIndex) + "',"
            ZSql = ZSql + " Califica32 = " + "'" + Str$(Califica32.ListIndex) + "',"
            ZSql = ZSql + " Califica33 = " + "'" + Str$(Califica33.ListIndex) + "',"
            ZSql = ZSql + " Califica34 = " + "'" + Str$(Califica34.ListIndex) + "',"
            ZSql = ZSql + " Califica35 = " + "'" + "0" + "',"
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
            ZSql = ZSql + " Promedio11 = " + "'" + Promedio11.Text + "',"
            ZSql = ZSql + " Promedio22 = " + "'" + Promedio22.Text + "',"
            ZSql = ZSql + " Promedio33 = " + "'" + Promedio33.Text + "',"
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
            ZSql = ZSql + "Sector1 ,"
            ZSql = ZSql + "Sector2 ,"
            ZSql = ZSql + "Sector3 ,"
            ZSql = ZSql + "DesSector1 ,"
            ZSql = ZSql + "DesSector2 ,"
            ZSql = ZSql + "DesSector3 ,"
            ZSql = ZSql + "Califica11 ,"
            ZSql = ZSql + "Califica12 ,"
            ZSql = ZSql + "Califica13 ,"
            ZSql = ZSql + "Califica14 ,"
            ZSql = ZSql + "Califica15 ,"
            ZSql = ZSql + "Califica21 ,"
            ZSql = ZSql + "Califica22 ,"
            ZSql = ZSql + "Califica23 ,"
            ZSql = ZSql + "Califica24 ,"
            ZSql = ZSql + "Califica25 ,"
            ZSql = ZSql + "Califica31 ,"
            ZSql = ZSql + "Califica32 ,"
            ZSql = ZSql + "Califica33 ,"
            ZSql = ZSql + "Califica34 ,"
            ZSql = ZSql + "Califica35 ,"
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
            ZSql = ZSql + "Promedio11 ,"
            ZSql = ZSql + "Promedio22 ,"
            ZSql = ZSql + "Promedio33 ,"
            ZSql = ZSql + "Promedio )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + "4" + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Mes.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + ZPeriodo + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + "'" + Evaluador.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Str$(Sector1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Sector2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Sector3.ListIndex) + "',"
            ZSql = ZSql + "'" + Sector1.Text + "',"
            ZSql = ZSql + "'" + Sector2.Text + "',"
            ZSql = ZSql + "'" + Sector3.Text + "',"
            ZSql = ZSql + "'" + Str$(Califica11.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica12.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica13.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica14.ListIndex) + "',"
            ZSql = ZSql + "'" + "0" + "',"
            ZSql = ZSql + "'" + Str$(Califica21.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica22.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica23.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica24.ListIndex) + "',"
            ZSql = ZSql + "'" + "0" + "',"
            ZSql = ZSql + "'" + Str$(Califica31.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica32.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica33.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Califica34.ListIndex) + "',"
            ZSql = ZSql + "'" + "0" + "',"
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
            ZSql = ZSql + "'" + Promedio11.Text + "',"
            ZSql = ZSql + "'" + Promedio22.Text + "',"
            ZSql = ZSql + "'" + Promedio33.Text + "',"
            ZSql = ZSql + "'" + Promedio.Text + "')"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        EmpresaReal = WEmpresa
        
        Select Case Val(Promedio.Text)
            Case 1
                WCalifica = "3"
            Case 5
                WCalifica = "2"
            Case 10
                WCalifica = "1"
            Case Else
                WCalifica = "0"
        End Select
        WFechaCalifica = "31" + "/" + Mes.Text + "/" + Ano.Text
        WOrdFechaCalifica = Ano.Text + Mes.Text + "31"
    
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0009"
        txtOdbc = "Empresa09"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesProve.Text + "',"
        ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
        ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
        ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = EmpresaReal
        txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
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
        ZSql = ZSql + " and EvaluaI.Tipo = 4"
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
                ZSql = ZSql + " and EvaluaI.Tipo = 4"
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
    ObservacionesProve.Text = ""
    Cartel.Visible = False
    
    Sector1.ListIndex = 0
    Sector2.ListIndex = 0
    Sector3.ListIndex = 0
    
    Califica11.ListIndex = 0
    Califica12.ListIndex = 0
    Califica13.ListIndex = 0
    Califica14.ListIndex = 0
    Califica21.ListIndex = 0
    Califica22.ListIndex = 0
    Califica23.ListIndex = 0
    Califica24.ListIndex = 0
    Califica31.ListIndex = 0
    Califica32.ListIndex = 0
    Califica33.ListIndex = 0
    Califica34.ListIndex = 0
    
    Promedio.Text = ""
    Promedio11.Text = ""
    Promedio22.Text = ""
    Promedio33.Text = ""
    
    DesPromedio.Text = ""
    DesPromedio11.Text = ""
    DesPromedio22.Text = ""
    DesPromedio33.Text = ""
    
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
    ZSql = ZSql + " and EvaluaI.Tipo = 4"
    ZSql = ZSql + " Order by Clave"
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

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub









Private Sub Calcula_Promedio()


    Promedio11.Text = "0"
    Promedio22.Text = "0"
    Promedio33.Text = "0"
    
    
    ZCalifica11 = Califica11.ListIndex
    ZCalifica12 = Califica12.ListIndex
    ZCalifica13 = Califica13.ListIndex
    ZCalifica14 = Califica14.ListIndex
    If ZCalifica12 = 4 Then
        ZCalifica12 = 1
    End If
    If ZCalifica13 = 4 Then
        ZCalifica13 = 1
    End If
    If ZCalifica14 = 4 Then
        ZCalifica14 = 1
    End If
    
    
    ZCalifica21 = Califica21.ListIndex
    ZCalifica22 = Califica22.ListIndex
    ZCalifica23 = Califica23.ListIndex
    ZCalifica24 = Califica24.ListIndex
    If ZCalifica22 = 4 Then
        ZCalifica22 = 1
    End If
    If ZCalifica23 = 4 Then
        ZCalifica23 = 1
    End If
    If ZCalifica24 = 4 Then
        ZCalifica24 = 1
    End If
    
    
    ZCalifica31 = Califica31.ListIndex
    ZCalifica32 = Califica32.ListIndex
    ZCalifica33 = Califica33.ListIndex
    ZCalifica34 = Califica34.ListIndex
    If ZCalifica32 = 4 Then
        ZCalifica32 = 1
    End If
    If ZCalifica33 = 4 Then
        ZCalifica33 = 1
    End If
    If ZCalifica34 = 4 Then
        ZCalifica34 = 1
    End If
    

    If Sector1.ListIndex <> 0 Then
        If ZCalifica11 = 2 Then
            Promedio11.Text = "1"
                Else
            If ZCalifica12 = 1 And ZCalifica13 = 1 And ZCalifica14 = 1 Then
                Promedio11.Text = "10"
                    Else
                Promedio11.Text = "5"
            End If
        End If
    End If
            
    If Sector2.ListIndex <> 0 Then
        If ZCalifica21 = 2 Then
            Promedio22.Text = "1"
                Else
            If ZCalifica22 = 1 And ZCalifica23 = 1 And ZCalifica24 = 1 Then
                Promedio22.Text = "10"
                    Else
                Promedio22.Text = "5"
            End If
        End If
    End If
            
    If Sector3.ListIndex <> 0 Then
        If ZCalifica31 = 2 Then
            Promedio33.Text = "1"
                Else
            If ZCalifica32 = 1 And ZCalifica33 = 1 And ZCalifica34 = 1 Then
                Promedio33.Text = "10"
                    Else
                Promedio33.Text = "5"
            End If
        End If
    End If
    
    Select Case Val(Promedio11.Text)
        Case 1
            DesPromedio11.Text = "No Apto"
            DesPromedio11.BackColor = &H8080FF
        Case 5
            DesPromedio11.Text = "Condicional"
            DesPromedio11.BackColor = &HC0FFFF
        Case 10
            DesPromedio11.Text = "Apto"
            DesPromedio11.BackColor = &HC0FFC0
        Case Else
            DesPromedio11.Text = ""
            DesPromedio11.BackColor = &HFFFFFF
    End Select
    
    Select Case Val(Promedio22.Text)
        Case 1
            DesPromedio22.Text = "No Apto"
            DesPromedio22.BackColor = &H8080FF
        Case 5
            DesPromedio22.Text = "Condicional"
            DesPromedio22.BackColor = &HC0FFFF
        Case 10
            DesPromedio22.Text = "Apto"
            DesPromedio22.BackColor = &HC0FFC0
        Case Else
            DesPromedio22.Text = ""
            DesPromedio22.BackColor = &HFFFFFF
    End Select
    
    Select Case Val(Promedio33.Text)
        Case 1
            DesPromedio33.Text = "No Apto"
            DesPromedio33.BackColor = &H8080FF
        Case 5
            DesPromedio33.Text = "Condicional"
            DesPromedio33.BackColor = &HC0FFFF
        Case 10
            DesPromedio33.Text = "Apto"
            DesPromedio33.BackColor = &HC0FFC0
        Case Else
            DesPromedio33.Text = ""
            DesPromedio33.BackColor = &HFFFFFF
    End Select
    
    Promedio.Text = "10"
    
    If Val(Promedio11.Text) = 0 And Val(Promedio22.Text) = 0 And Val(Promedio33.Text) = 0 Then
        Promedio.Text = "0"
    End If
    If Val(Promedio11.Text) = 5 Or Val(Promedio22.Text) = 5 Or Val(Promedio33.Text) = 5 Then
        Promedio.Text = "5"
    End If
    If Val(Promedio11.Text) = 1 Or Val(Promedio22.Text) = 1 Or Val(Promedio33.Text) = 1 Then
        Promedio.Text = "1"
    End If
            
    Select Case Val(Promedio.Text)
        Case 1
            DesPromedio.Text = "No Apto"
            DesPromedio.BackColor = &H8080FF
        Case 5
            DesPromedio.Text = "Condicional"
            DesPromedio.BackColor = &HC0FFFF
        Case 10
            DesPromedio.Text = "Apto"
            DesPromedio.BackColor = &HC0FFC0
        Case Else
            DesPromedio.Text = ""
            DesPromedio.BackColor = &HFFFFFF
    End Select

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
            ObservacionesProve.Text = IIf(IsNull(rstProveedor!ObservacionesII), "", rstProveedor!ObservacionesII)
            WEstado = IIf(IsNull(rstProveedor!Estado), "0", rstProveedor!Estado)
            If WEstado = 2 Then
                Cartel.Visible = True
                    Else
                Cartel.Visible = False
            End If
            rstProveedor.Close
            Mes.SetFocus
                Else
            DesProveedor.Caption = ""
            ObservacionesProve.Text = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
        ObservacionesProve.Text = ""
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
            ZSql = ZSql + " and EvaluaI.Tipo = 4"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            If rstEvaluaI.RecordCount > 0 Then
            
                rstEvaluaI.Close
                Call Imprime_Datos
                
                    Else
                    
                ZDesProveedor = DesProveedor.Caption
                ZProveedor = Proveedor.Text
                ZMes = Mes.Text
                ZAno = Ano.Text
                CmdLimpiar_Click
                Proveedor.Text = ZProveedor
                Mes.Text = ZMes
                Ano.Text = ZAno
                DesProveedor.Caption = ZDesProveedor
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Proveedor"
                ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    DesProveedor.Caption = rstProveedor!Nombre
                    ObservacionesProve.Text = IIf(IsNull(rstProveedor!ObservacionesII), "", rstProveedor!ObservacionesII)
                    WEstado = IIf(IsNull(rstProveedor!Estado), "0", rstProveedor!Estado)
                    If WEstado = 2 Then
                        Cartel.Visible = True
                            Else
                        Cartel.Visible = False
                    End If
                    rstProveedor.Close
                        Else
                    DesProveedor.Caption = ""
                End If
                
            End If
            
            Observaciones.SetFocus
            
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
    ZSql = ZSql + " Where EvaluaI.Clave > " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 4"
    ZSql = ZSql + " Order by Clave"
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

    Califica11.Clear
    Califica12.Clear
    Califica13.Clear
    Califica14.Clear

    Califica21.Clear
    Califica22.Clear
    Califica23.Clear
    Califica24.Clear

    Califica31.Clear
    Califica32.Clear
    Califica33.Clear
    Califica34.Clear
    
    Califica11.AddItem ""
    Califica11.AddItem "Cumple"
    Califica11.AddItem "No Cumple"
    
    Califica12.AddItem ""
    Califica12.AddItem "Cumple"
    Califica12.AddItem "Parcial"
    Califica12.AddItem "No Cumple"
    Califica12.AddItem "No Aplica"
    
    Califica13.AddItem ""
    Califica13.AddItem "Cumple"
    Califica13.AddItem "Parcial"
    Califica13.AddItem "No Cumple"
    Califica13.AddItem "No Aplica"
    
    Califica14.AddItem ""
    Califica14.AddItem "Cumple"
    Califica14.AddItem "Parcial"
    Califica14.AddItem "No Cumple"
    Califica14.AddItem "No Aplica"
    
    Califica21.AddItem ""
    Califica21.AddItem "Cumple"
    Califica21.AddItem "No Cumple"
    
    Califica22.AddItem ""
    Califica22.AddItem "Cumple"
    Califica22.AddItem "Parcial"
    Califica22.AddItem "No Cumple"
    Califica22.AddItem "No Aplica"
    
    Califica23.AddItem ""
    Califica23.AddItem "Cumple"
    Califica23.AddItem "Parcial"
    Califica23.AddItem "No Cumple"
    Califica23.AddItem "No Aplica"
    
    Califica24.AddItem ""
    Califica24.AddItem "Cumple"
    Califica24.AddItem "Parcial"
    Califica24.AddItem "No Cumple"
    Califica24.AddItem "No Aplica"
    
    
    Califica31.AddItem ""
    Califica31.AddItem "Cumple"
    Califica31.AddItem "No Cumple"
    
    Califica32.AddItem ""
    Califica32.AddItem "Cumple"
    Califica32.AddItem "Parcial"
    Califica32.AddItem "No Cumple"
    Califica32.AddItem "No Aplica"
    
    Califica33.AddItem ""
    Califica33.AddItem "Cumple"
    Califica33.AddItem "Parcial"
    Califica33.AddItem "No Cumple"
    Califica33.AddItem "No Aplica"

    Califica34.AddItem ""
    Califica34.AddItem "Cumple"
    Califica34.AddItem "Parcial"
    Califica34.AddItem "No Cumple"
    Califica34.AddItem "No Aplica"
    
    
    Califica11.ListIndex = 0
    Califica12.ListIndex = 0
    Califica13.ListIndex = 0
    Califica14.ListIndex = 0
    
    Califica21.ListIndex = 0
    Califica22.ListIndex = 0
    Califica23.ListIndex = 0
    Califica24.ListIndex = 0
    
    Califica31.ListIndex = 0
    Califica32.ListIndex = 0
    Califica33.ListIndex = 0
    Califica34.ListIndex = 0
    
    Sector1.Clear
    Sector2.Clear
    Sector3.Clear

    Sector1.AddItem ""
    Sector1.AddItem "C.Calidad"
    Sector1.AddItem "Desarrollo"
    Sector1.AddItem "Pigmentos"
    Sector1.AddItem "Textil"
    Sector1.AddItem "Manten."

    Sector2.AddItem ""
    Sector2.AddItem "C.Calidad"
    Sector2.AddItem "Desarrollo"
    Sector2.AddItem "Pigmentos"
    Sector2.AddItem "Textil"
    Sector2.AddItem "Manten."

    Sector3.AddItem ""
    Sector3.AddItem "C.Calidad"
    Sector3.AddItem "Desarrollo"
    Sector3.AddItem "Pigmentos"
    Sector3.AddItem "Textil"
    Sector3.AddItem "Manten."
    
    Sector1.ListIndex = 0
    Sector2.ListIndex = 0
    Sector3.ListIndex = 0

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Mes.Text = ""
    Ano.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Evaluador.Text = ""
    Observaciones.Text = ""
    ObservacionesProve.Text = ""
    
    Promedio.Text = ""
    Promedio11.Text = ""
    Promedio22.Text = ""
    Promedio33.Text = ""
    
    DesPromedio.Text = ""
    DesPromedio11.Text = ""
    DesPromedio22.Text = ""
    DesPromedio33.Text = ""
    
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
                            DA = Len(rstProveedor!Nombre) - WEspacios
                            For aa = 1 To DA + 1
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

Private Sub Califica11_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica12_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica13_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica14_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica21_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica22_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica23_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica24_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica31_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica32_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica33_Click()
    Call Calcula_Promedio
End Sub

Private Sub Califica34_Click()
    Call Calcula_Promedio
End Sub

