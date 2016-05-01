VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProyectos 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Proyectos"
   ClientHeight    =   8475
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8475
   ScaleWidth      =   11790
   Begin VB.Frame xclave2 
      Height          =   1935
      Left            =   3240
      TabIndex        =   57
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cancelagraba2 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   59
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox wClave2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   58
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   60
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "proyectos.frx":0000
      Left            =   8640
      List            =   "proyectos.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   56
      Text            =   "Solicitante"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3240
      TabIndex        =   52
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   54
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   53
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   55
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox Solicitante 
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
      MaxLength       =   20
      TabIndex        =   50
      Text            =   " "
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox Tipo4 
      Caption         =   "Obra - Equipo"
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
      Left            =   8760
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Tipo3 
      Caption         =   "Obra - Mano de Obra"
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
      Left            =   8760
      TabIndex        =   46
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox Tipo2 
      Caption         =   "Obra - Material"
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
      Left            =   8760
      TabIndex        =   45
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Tipo1 
      Caption         =   "Equipo"
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
      Left            =   8760
      TabIndex        =   44
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Planta 
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
      ItemData        =   "proyectos.frx":009A
      Left            =   4680
      List            =   "proyectos.frx":009C
      TabIndex        =   42
      Text            =   "Planta I"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Ano 
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
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   41
      Text            =   " "
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Gasto 
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
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   39
      Text            =   " "
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ComboBox Prioridad 
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
      Left            =   1680
      TabIndex        =   34
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ComboBox Estado 
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
      ItemData        =   "proyectos.frx":009E
      Left            =   1680
      List            =   "proyectos.frx":00A0
      TabIndex        =   32
      Text            =   "Pendiente de aprobacion"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox Presupuesto 
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
      Left            =   6720
      MaxLength       =   15
      TabIndex        =   29
      Text            =   " "
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Centro 
      BackColor       =   &H80000000&
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox ObservaV 
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
      TabIndex        =   27
      Top             =   2640
      Width           =   9015
   End
   Begin VB.TextBox ObservaIV 
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
      TabIndex        =   26
      Top             =   2280
      Width           =   9015
   End
   Begin VB.TextBox ObservaIII 
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
      TabIndex        =   25
      Top             =   1920
      Width           =   9015
   End
   Begin VB.TextBox ObservaII 
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
      TabIndex        =   24
      Top             =   1560
      Width           =   9015
   End
   Begin VB.TextBox ObservaI 
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
      TabIndex        =   22
      Top             =   1200
      Width           =   9015
   End
   Begin VB.TextBox Sector 
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
      TabIndex        =   18
      Text            =   " "
      Top             =   480
      Width           =   1095
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6600
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
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1440
      TabIndex        =   5
      Top             =   6120
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
         MouseIcon       =   "proyectos.frx":00A2
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":03AC
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "proyectos.frx":07EE
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":0AF8
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
      Left            =   5040
      TabIndex        =   14
      Top             =   4560
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "proyectos.frx":0F3A
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":1244
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "proyectos.frx":1686
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":1990
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "proyectos.frx":1DD2
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":20DC
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "proyectos.frx":251E
         MousePointer    =   99  'Custom
         Picture         =   "proyectos.frx":2828
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
      Left            =   120
      TabIndex        =   13
      Top             =   5520
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Sector.rpt"
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
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
      Left            =   1560
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   2415
      Left            =   1080
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox FechaInicio 
      Height          =   285
      Left            =   1680
      TabIndex        =   35
      Top             =   3720
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
   Begin MSMask.MaskEdBox FechaFinal 
      Height          =   285
      Left            =   6360
      TabIndex        =   37
      Top             =   3720
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
   Begin MSMask.MaskEdBox FechaAprobado 
      Height          =   285
      Left            =   1680
      TabIndex        =   48
      Top             =   4440
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
   Begin VB.Image BusquedaCodigo 
      Height          =   480
      Left            =   6600
      MouseIcon       =   "proyectos.frx":2C6A
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":2F74
      ToolTipText     =   "Consulta de Proveedores"
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label13 
      Caption         =   "Solicitante"
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
      Left            =   7200
      TabIndex        =   51
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha Aprobac."
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
      TabIndex        =   49
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Importe Real"
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
      Left            =   4200
      TabIndex        =   40
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha Finalizacion"
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
      Left            =   4200
      TabIndex        =   38
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha Inicio"
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
      TabIndex        =   36
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Prioridad"
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
      TabIndex        =   33
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Planta"
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
      Left            =   3960
      TabIndex        =   31
      Top             =   3345
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Importe Presupuestado en Pesos ($)"
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
      TabIndex        =   30
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblLabels 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label23 
      Caption         =   "Año Asignado"
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
      Left            =   8040
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label DesSector 
      BackColor       =   &H00FFFF00&
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
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3480
      MouseIcon       =   "proyectos.frx":33B6
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":36C0
      ToolTipText     =   "Impresion "
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1800
      MouseIcon       =   "proyectos.frx":3F02
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":420C
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   120
      MouseIcon       =   "proyectos.frx":4A4E
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":4D58
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   960
      MouseIcon       =   "proyectos.frx":559A
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":58A4
      ToolTipText     =   "Elimina el Registro"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "proyectos.frx":60E6
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":63F0
      ToolTipText     =   "Salida"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2640
      MouseIcon       =   "proyectos.frx":6C32
      MousePointer    =   99  'Custom
      Picture         =   "proyectos.frx":6F3C
      ToolTipText     =   "Consulta de Datos"
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label lblLabels 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
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
      Width           =   1095
   End
End
Attribute VB_Name = "PrgProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProyecto As Recordset
Dim spProyecto As String
Dim rstSectorInve As Recordset
Dim spSectorInve As String
Private WGraba As String

Sub Verifica_datos()
    If Val(codigo.Text) = 0 Then
        codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()
    estado.Clear
    
    estado.AddItem ""
    estado.AddItem "Pendiente de Aprobar"
    estado.AddItem "Aprobado"
    estado.AddItem "En Ejecucion"
    estado.AddItem "Finalizado"
    
    
    sql1 = "Select *"
    sql2 = " FROM Proyecto"
    sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
    spProyecto = sql1 + sql2 + sql3
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        sector.Text = Str(rstProyecto!sector)
        Ano.Text = Str$(rstProyecto!Ano)
        descripcion.Text = Trim(rstProyecto!descripcion)
        ObservaI.Text = Trim(rstProyecto!ObservaI)
        ObservaII.Text = Trim(rstProyecto!ObservaII)
        ObservaIII.Text = Trim(rstProyecto!ObservaIII)
        ObservaIV.Text = Trim(rstProyecto!ObservaIV)
        ObservaV.Text = Trim(rstProyecto!ObservaV)
        Centro.Text = Trim(rstProyecto!Centro)
        Presupuesto.Text = Str(rstProyecto!Presupuesto)
        Presupuesto.Text = Pusing("###,###.##", Presupuesto.Text)
        Prioridad.ListIndex = rstProyecto!Prioridad
        estado.ListIndex = rstProyecto!estado
               
        planta.Text = rstProyecto!planta
        Rem Planta.ListIndex = Planta.Text
        
        
        FechaInicio.Text = "00/00/0000"
        FechaInicio.Text = IIf(IsNull(rstProyecto!FechaInicio), "  /  /    ", rstProyecto!FechaInicio)
        FechaFinal.Text = rstProyecto!FechaFinal
        Gasto.Text = Str$(rstProyecto!Gasto)
        Gasto.Text = Pusing("###,###.##", Gasto.Text)
        Tipo1.Value = rstProyecto!Tipo1
        Tipo2.Value = rstProyecto!Tipo2
        Tipo3.Value = rstProyecto!Tipo3
        Tipo4.Value = rstProyecto!Tipo4
        FechaAprobado.Text = IIf(IsNull(rstProyecto!FechaAprobado), "  /  /    ", rstProyecto!FechaAprobado)
        Solicitante.Text = IIf(IsNull(rstProyecto!Solicitante), "", rstProyecto!Solicitante)
        rstProyecto.Close
        Combo1.Text = Solicitante.Text
         Combo1.Enabled = False
         planta.Enabled = False
    
    End If
    
    Rem nan
    If estado.Text = "Aprobado" Or estado.Text = "Finalizado" Then
        Presupuesto.Enabled = False
        FechaAprobado.Enabled = False
        Gasto.Enabled = False
        Centro.Enabled = False
        If estado.Text = "Finalizado" Then
            Prioridad.Enabled = False
        End If
    End If
     
    
    
    sql1 = "Select *"
    sql2 = " FROM SectorInve"
    sql3 = " Where SectorInve.Codigo = " + "'" + sector.Text + "'"
    spSectorInve = sql1 + sql2 + sql3
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        DesSector.Caption = Trim(rstSectorInve!descripcion)
        rstSectorInve.Close
    End If
    
    ZSuma = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Avance"
    ZSql = ZSql + " Where Avance.Proyecto = " + "'" + codigo.Text + "'"
    spAvance = ZSql
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        With rstAvance
            .MoveFirst
            Do
                If .EOF = False Then
                    ZSuma = ZSuma + rstAvance!Importe
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstAvance.Close
    End If
    
    Gasto.Text = Str$(ZSuma)
    Gasto.Text = Pusing("###,###.##", Gasto.Text)
    
End Sub

Private Sub BusquedaCodigo_Click()
    
    ZCodigo = 0
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proyecto"
    ZSql = ZSql + " Where Proyecto.Planta = " + "'" + Str$(planta.ListIndex) + "'"
    ZSql = ZSql + " Order by Proyecto.Codigo"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        With rstProyecto
            .MoveFirst
            Do
                If .EOF = False Then

                    ZCodigo = rstProyecto!codigo
                    .MoveNext
                    
                        Else
                    
                    Exit Do
                
                End If
            
            Loop
        End With
        rstProyecto.Close
    End If
    
    codigo.Text = Str$(ZCodigo + 1)
    
End Sub

Private Sub cancelagraba2_Click()
 xclave2.Visible = False
End Sub

Private Sub cmdAdd_Click()




    If Val(codigo.Text) <> 0 And WGraba = "N" Then
        Zestado = 0
        sql1 = "Select *"
        sql2 = " FROM Proyecto"
        sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
        spProyecto = sql1 + sql2 + sql3
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        If rstProyecto.RecordCount > 0 Then
            Zestado = rstProyecto!estado
            rstProyecto.Close
        End If
        If Zestado <= 1 And estado.ListIndex >= 2 Then
            WGraba = "N"
                Else
            WGraba = "S"
        End If
            Else
        sql1 = "Select *"
        sql2 = " FROM Proyecto"
        sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
        spProyecto = sql1 + sql2 + sql3
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        If rstProyecto.RecordCount > 0 Then
            Zestado = rstProyecto!estado
            rstProyecto.Close
                Else
            Rem nan
            If WGraba = "S" Then
                   Else
                WGraba = "N"
            End If
        
        End If
    End If
      
    If WGraba <> "S" Then
              
        Call Ingresa_clave
                 
            Else
        
        Rem BY NAN
        pepe = estado.Text
        If pepe = "Pendiente de Aprobar" Then
            estado.Text = 1
        End If
        If pepe = "Aprobado" Then
            estado.Text = 2
        End If
        If pepe = "En Ejecucion" Then
            estado.Text = 3
        End If
        If pepe = "Finalizado" Then
            estado.Text = 4
        End If
            
        If Val(codigo.Text) <> 0 Then
                 
            If planta.ListIndex = 0 Then
                m$ = "Se debe informar planta"
                A% = MsgBox(m$, 0, "Archivo de Proyectos")
                Exit Sub
            End If
    
            If Prioridad.ListIndex = 0 Then
                m$ = "Se debe informar Prioridad"
                A% = MsgBox(m$, 0, "Archivo de Proyectos")
                Exit Sub
            End If
    
            
            Rem    m$ = "Se debe informar Estado del Proyecto"
            Rem    A% = MsgBox(m$, 0, "Archivo de Proyectos")
            Rem    Exit Sub
            Rem End If
    
            If estado.ListIndex = 2 Then
                If FechaAprobado.Text = "  /  /    " Then
                    m$ = "Se debe informar Fecha de Aprobacion"
                    A% = MsgBox(m$, 0, "Archivo de Proyectos")
                    Exit Sub
                End If
            End If
    
            If estado.ListIndex = -1 Then
                If Trim(Solicitante.Text) = "" Then
                    m$ = "Se debe informar Solicitante"
                    A% = MsgBox(m$, 0, "Archivo de Proyectos")
                    
                   Combo1.Enabled = True
                   
                    Exit Sub
                End If
            End If
        
   
        
            ZTipo1 = Tipo1.Value
            ZTipo2 = Tipo2.Value
            ZTipo3 = Tipo3.Value
            ZTipo4 = Tipo4.Value
        
            sql1 = "Select *"
            sql2 = " FROM Proyecto"
            sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
            spProyecto = sql1 + sql2 + sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                rstProyecto.Close
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Proyecto SET "
                ZSql = ZSql + " Sector = " + "'" + sector.Text + "',"
                ZSql = ZSql + " Ano = " + "'" + Ano.Text + "',"
                ZSql = ZSql + " Descripcion = " + "'" + descripcion.Text + "',"
                ZSql = ZSql + " ObservaI = " + "'" + ObservaI.Text + "',"
                ZSql = ZSql + " ObservaII = " + "'" + ObservaII.Text + "',"
                ZSql = ZSql + " ObservaIII = " + "'" + ObservaIII.Text + "',"
                ZSql = ZSql + " ObservaIV = " + "'" + ObservaIV.Text + "',"
                ZSql = ZSql + " ObservaV = " + "'" + ObservaV.Text + "',"
                ZSql = ZSql + " Centro = " + "'" + Centro.Text + "',"
                ZSql = ZSql + " Presupuesto = " + "'" + Presupuesto.Text + "',"
                ZSql = ZSql + " Gasto = " + "'" + Gasto.Text + "',"
                ZSql = ZSql + " Prioridad = " + "'" + Str$(Prioridad.ListIndex) + "',"
                ZSql = ZSql + " Estado = " + "'" + Str$(estado.Text) + "',"
                ZSql = ZSql + " Planta = " + "'" + Str$(planta.Text) + "',"
                ZSql = ZSql + " Tipo1 = " + "'" + Str$(ZTipo1) + "',"
                ZSql = ZSql + " Tipo2 = " + "'" + Str$(ZTipo2) + "',"
                ZSql = ZSql + " Tipo3 = " + "'" + Str$(ZTipo3) + "',"
                ZSql = ZSql + " Tipo4 = " + "'" + Str$(ZTipo4) + "',"
                ZSql = ZSql + " Solicitante = " + "'" + Solicitante.Text + "',"
                ZSql = ZSql + " FechaAprobado = " + "'" + FechaAprobado.Text + "',"
                ZSql = ZSql + " FechaInicio = " + "'" + FechaInicio.Text + "',"
                ZSql = ZSql + " FechaFinal = " + "'" + FechaFinal.Text + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + codigo.Text + "'"
                spProyecto = ZSql
                Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            
                    Else
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Proyecto ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Sector ,"
                ZSql = ZSql + "Ano ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "ObservaI ,"
                ZSql = ZSql + "ObservaII ,"
                ZSql = ZSql + "ObservaIII ,"
                ZSql = ZSql + "ObservaIV ,"
                ZSql = ZSql + "ObservaV ,"
                ZSql = ZSql + "Centro ,"
                ZSql = ZSql + "Presupuesto ,"
                ZSql = ZSql + "Gasto ,"
                ZSql = ZSql + "Prioridad ,"
                ZSql = ZSql + "Estado ,"
                ZSql = ZSql + "Planta ,"
                ZSql = ZSql + "Tipo1 ,"
                ZSql = ZSql + "Tipo2 ,"
                ZSql = ZSql + "Tipo3 ,"
                ZSql = ZSql + "Tipo4 ,"
                ZSql = ZSql + "Solicitante ,"
                ZSql = ZSql + "FechaAprobado ,"
                ZSql = ZSql + "FechaInicio ,"
                ZSql = ZSql + "FechaFinal )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + codigo.Text + "',"
                ZSql = ZSql + "'" + sector.Text + "',"
                ZSql = ZSql + "'" + Ano.Text + "',"
                ZSql = ZSql + "'" + descripcion.Text + "',"
                ZSql = ZSql + "'" + ObservaI.Text + "',"
                ZSql = ZSql + "'" + ObservaII.Text + "',"
                ZSql = ZSql + "'" + ObservaIII.Text + "',"
                ZSql = ZSql + "'" + ObservaIV.Text + "',"
                ZSql = ZSql + "'" + ObservaV.Text + "',"
                ZSql = ZSql + "'" + Centro.Text + "',"
                ZSql = ZSql + "'" + Presupuesto.Text + "',"
                ZSql = ZSql + "'" + Gasto.Text + "',"
                ZSql = ZSql + "'" + Str$(Prioridad.ListIndex) + "',"
                ZSql = ZSql + "'" + Str$(estado.Text) + "',"
                ZSql = ZSql + "'" + Str$(planta.ListIndex) + "',"
                ZSql = ZSql + "'" + Str$(ZTipo1) + "',"
                ZSql = ZSql + "'" + Str$(ZTipo2) + "',"
                ZSql = ZSql + "'" + Str$(ZTipo3) + "',"
                ZSql = ZSql + "'" + Str$(ZTipo4) + "',"
                ZSql = ZSql + "'" + Solicitante.Text + "',"
                ZSql = ZSql + "'" + FechaAprobado.Text + "',"
                ZSql = ZSql + "'" + FechaInicio.Text + "',"
                ZSql = ZSql + "'" + FechaFinal.Text + "')"
                spProyecto = ZSql
                Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
                planta.Text = ""
            End If
    
            Call CmdLimpiar_Click
            codigo.SetFocus
            
        End If
        
    End If

End Sub

Private Sub cmdDelete_Click()
   
    If Val(codigo.Text) <> 0 Then
        sql1 = "Select *"
        sql2 = " FROM Proyecto"
        sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
        spProyecto = sql1 + sql2 + sql3
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        If rstProyecto.RecordCount > 0 Then
           If wgraba2 <> "S" Then
              rstProyecto.Close
              T$ = "Borrar Registro"
              m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
             If Respuesta% = 6 Then
                       
               Call Ingresa_clave2
          
               End If
          
            Else
        
     
                sql1 = "DELETE Proyecto"
                sql2 = " Where Codigo = " + "'" + codigo.Text + "'"
                spProyecto = sql1 + sql2
                Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    codigo.SetFocus
  
  
  

End Sub

Private Sub CmdLimpiar_Click()
    
    Gasto.Text = " "
    WGraba = ""
    codigo.Text = ""
    sector.Text = ""
    Ano.Text = ""
    descripcion.Text = ""
    ObservaI.Text = ""
    ObservaII.Text = ""
    ObservaIII.Text = ""
    ObservaIV.Text = ""
    ObservaV.Text = ""
    Centro.Text = ""
    Presupuesto.Text = ""
    FechaInicio.Text = "  /  /    "
    FechaFinal.Text = "  /  /    "
    FechaAprobado.Text = "  /  /    "
    Solicitante.Text = ""
    Combo1.Text = ""
    Combo1.Enabled = True
    
    Prioridad.ListIndex = 0
    estado.ListIndex = 0
    planta.Text = ""
    planta.Enabled = True
        
    Tipo1.Value = 0
    Tipo2.Value = 0
    Tipo3.Value = 0
    Tipo4.Value = 0
  
    DesSector.Caption = ""
    estado.Text = " "
    
    Rem   sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Rem   Sql2 = " FROM Proyecto"
    Rem   spProyecto = sql1 + Sql2
    Rem   Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    Rem   If rstProyecto.RecordCount > 0 Then
    Rem       rstProyecto.MoveLast
    Rem       ZCodigo = IIf(IsNull(rstProyecto!CodigoMayor), "0", rstProyecto!CodigoMayor)
    Rem       Codigo.Text = ZCodigo + 1
    Rem       rstProyecto.Close
    Rem   End If
    Rem   If Val(Codigo.Text) = 0 Then
    Rem       Codigo.Text = "1"
    Rem   End If
    
    codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgProyectos.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()
    sql1 = "Select *"
    sql2 = " FROM Proyecto"
    sql3 = " Where Proyecto.Codigo < " + "'" + codigo.Text + "'"
    sql4 = " Order by Proyecto.Codigo"
    spProyecto = sql1 + sql2 + sql3 + sql4
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        With rstProyecto
            .MoveLast
            codigo.Text = rstProyecto!codigo
        End With
        rstProyecto.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No existe registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Proyectos")
    End If
End Sub

Private Sub Combo1_Change()
    Solicitante.Text = Combo1.Text
End Sub

Private Sub Combo1_Click()
    Solicitante.Text = Combo1.Text
End Sub

Private Sub Estado_DropDown()

    If estado.Text = "Pendiente de Aprobar" Then
        If WOperador = 3 Then
            estado.Clear
            Rem Estado.Text = "Aprobado"
            Rem  Estado.Text = "Pendiente de Aprobar"
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
            ent = 1
                Else
            Rem estado.Clear
            estado.Enabled = False
       End If
    End If
        
    If estado.Text = "Aprobado" Then
                                  
        If WOperador = 3 Then
        
            estado.Clear
          
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
            ent = 1
            
                Else
                
            estado.Clear
                 
            estado.Text = "En Ejecucion"
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
        End If
        
    End If
    
    If estado.Text = "En Ejecucion" Then
                           
        If WOperador = 3 Then
            estado.Clear
            estado.AddItem "Pendiente de Aprobar"
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
                Else
            estado.Clear
            estado.Text = "En Ejecucion"
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
        End If
        
    End If
              
    If estado.Text = "Finalizado" Then
        estado.Enabled = False
    End If
                  
    If estado.Text = "" Then
        If WOperador = 3 Then
            estado.Clear
            estado.AddItem "Pendiente de Aprobar"
            estado.AddItem "Aprobado"
            estado.AddItem "En Ejecucion"
            estado.AddItem "Finalizado"
                Else
            estado.Clear
            estado.Text = "Pendiente de Aprobar"
            estado.AddItem "Pendiente de Aprobar"
        End If
    End If
       
End Sub


Private Sub Planta_Click()

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    sql2 = "  FROM Proyecto"
    
    Rem by nan
    If planta.Text = "Planta I" Then
        sql3 = " Where proyecto.Codigo >1000 and proyecto.codigo< 2000 "
            Else
        If planta.Text = "Planta II" Then
            sql3 = " Where proyecto.Codigo >2000 and proyecto.codigo< 3000 "
                Else
            If planta.Text = "Planta III" Then
                sql3 = " Where proyecto.Codigo >3000 and proyecto.codigo< 4000 "
                  Rem    sql3 = " Where proyecto.Codigo >5000  "
            End If
        End If
          If planta.Text = "Planta V" Then
            sql3 = " Where proyecto.Codigo >5000 and proyecto.codigo <6000 "
        End If
         
         
         If planta.Text = "Planta VI" Then
            sql3 = " Where proyecto.Codigo >6000 and proyecto.codigo< 7000 "
        End If
    
        If planta.Text = "Planta VII" Then
            sql3 = " Where proyecto.Codigo >7000 and proyecto.codigo< 8000 "
        End If
    End If
    
    
    spProyecto = sql1 + sql2 + sql3
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        rstProyecto.MoveLast
        ZCodigo = IIf(IsNull(rstProyecto!CodigoMayor), "0", rstProyecto!CodigoMayor)
        codigo.Text = ZCodigo + 1
        rstProyecto.Close
    End If
End Sub

Private Sub Sector_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sql1 = "Select *"
        sql2 = " FROM SectorInve"
        sql3 = " Where SectorInve.Codigo = " + "'" + sector.Text + "'"
        spSectorInve = sql1 + sql2 + sql3
        Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
        If rstSectorInve.RecordCount > 0 Then
            DesSector.Caption = rstSectorInve!descripcion
            rstSectorInve.Close
            descripcion.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        sector.Text = ""
        DesSector.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI.SetFocus
    End If
    If KeyAscii = 27 Then
        descripcion.Text = ""
    End If
End Sub

Private Sub ObservaI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI.Text = ""
    End If
End Sub

Private Sub ObservaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII.Text = ""
    End If
End Sub

Private Sub ObservaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIII.Text = ""
    End If
End Sub

Private Sub ObservaIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIV.Text = ""
    End If
End Sub

Private Sub ObservaV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Centro.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaV.Text = ""
    End If
End Sub

Private Sub Centro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Presupuesto.SetFocus
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
    End If
End Sub

Private Sub Presupuesto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Presupuesto.Text = Pusing("###,###.##", Presupuesto.Text)
        Prioridad.SetFocus
    End If
    If KeyAscii = 27 Then
        Presupuesto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Prioridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        planta.SetFocus
    End If
End Sub

Private Sub Planta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solicitante.SetFocus
    End If
End Sub

Private Sub Solicitante_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaInicio.SetFocus
    End If
    If KeyAscii = 27 Then
        Solicitante.Text = ""
    End If
End Sub

Private Sub FechaInicio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaFinal.SetFocus
    End If
    If KeyAscii = 27 Then
        FechaInicio.Text = "  /  /    "
    End If
End Sub

Private Sub FechaFinal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        estado.SetFocus
    End If
    If KeyAscii = 27 Then
        FechaFinal.Text = "  /  /    "
    End If
End Sub

Private Sub FechaAprobado_Keypress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        FechaAprobado.Text = "  /  /    "
    End If
End Sub

Private Sub Estado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sector.SetFocus
        Rem nan
         
    End If
End Sub

Private Sub Gasto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Gasto.Text = Pusing("###,###.##", Gasto.Text)
        sector.SetFocus
    End If
    If KeyAscii = 27 Then
        Gasto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
   estado.Enabled = True
   If KeyAscii = 13 Then
        If Val(codigo.Text) <> 0 Then
            sql1 = "Select *"
            sql2 = " FROM Proyecto"
            sql3 = " Where Proyecto.Codigo = " + "'" + codigo.Text + "'"
            spProyecto = sql1 + sql2 + sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                rstProyecto.Close
                Call Imprime_Datos
                    Else
                WCodigo = codigo.Text
                CmdLimpiar_Click
                codigo.Text = WCodigo
            End If
        End If
        sector.SetFocus
    End If
    If KeyAscii = 27 Then
        codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Proyectos"
     Opcion.AddItem "Sector"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            sql2 = " FROM Proyecto"
            sql3 = " Order by Proyecto.Codigo"
            spProyecto = sql1 + sql2 + sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstProyecto!codigo
                            pantalla.Col = 2
                            pantalla.Text = rstProyecto!descripcion
                            IngresaItem = rstProyecto!codigo
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            
        Case 1
            sql1 = "Select *"
            sql2 = " FROM SectorInve"
            sql3 = " Order by SectorInve.Codigo"
            spSectorInve = sql1 + sql2 + sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                With rstSectorInve
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstSectorInve!codigo
                            pantalla.Col = 2
                            pantalla.Text = rstSectorInve!descripcion
                            IngresaItem = rstSectorInve!codigo
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSectorInve.Close
            End If
            
        Case Else
    End Select
            
    pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla_Click()

    pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = pantalla.Row - 1
            codigo.Text = windice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case 1
            Indice = pantalla.Row - 1
            sector.Text = windice.List(Indice)
            Call Sector_Keypress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    sql1 = "Select Min(Codigo) as [CodigoMenor]"
    sql2 = " FROM Proyecto"
    spProyecto = sql1 + sql2
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        rstProyecto.MoveFirst
        codigo.Text = rstProyecto!CodigoMenor
        rstProyecto.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    sql2 = " FROM Proyecto"
    spProyecto = sql1 + sql2
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        rstProyecto.MoveLast
        codigo.Text = rstProyecto!CodigoMayor
        rstProyecto.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    sql1 = "Select *"
    sql2 = " FROM Proyecto"
    sql3 = " Where Proyecto.Codigo > " + "'" + codigo.Text + "'"
    sql4 = " Order by Proyecto.Codigo"
    spProyecto = sql1 + sql2 + sql3 + sql4
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        With rstProyecto
            .MoveFirst
            codigo.Text = rstProyecto!codigo
        End With
        rstProyecto.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Proyectos")
    End If

End Sub

Sub Form_Load()

If WOperador = 36 Then
       Combo1.Text = "HFrias"
  End If
If WOperador = 70 Then
       Combo1.Text = "DRodriguez"
  End If

    Prioridad.Clear
    
    Prioridad.AddItem ""
    Prioridad.AddItem "Baja"
    Prioridad.AddItem "Media"
    Prioridad.AddItem "Alta"
    
    
  If WOperador = 3 Then
    
    estado.Clear
    estado.AddItem ""
    estado.AddItem "Pendiente de Aprobar"
    estado.AddItem "Aprobado"
    estado.AddItem "En Ejecucion"
    estado.AddItem "Finalizado"
       Else
       estado.Clear
       estado.AddItem "Pendiente de Aprobar"
    
     
    End If
    
    planta.Clear
    
    planta.AddItem ""
    planta.AddItem "Planta I"
    planta.AddItem "Planta II"
    planta.AddItem "Planta III"
    planta.AddItem ""
    planta.AddItem "Planta V"
    planta.AddItem "Planta VI"
    planta.AddItem "Planta VII"

    WGraba = ""
    codigo.Text = ""
    sector.Text = ""
    Ano.Text = ""
    descripcion.Text = ""
    ObservaI.Text = ""
    ObservaII.Text = ""
    ObservaIII.Text = ""
    ObservaIV.Text = ""
    ObservaV.Text = ""
    Centro.Text = ""
    Presupuesto.Text = ""
    FechaInicio.Text = "  /  /    "
    FechaFinal.Text = "  /  /    "
    FechaAprobado.Text = "  /  /    "
    Solicitante.Text = ""
    
    Prioridad.ListIndex = 0
    estado.ListIndex = 0
    Rem Planta.ListIndex = 0
    
    Tipo1.Value = 0
    Tipo2.Value = 0
    Tipo3.Value = 0
    Tipo4.Value = 0
    
    DesSector.Caption = ""
    
    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    sql2 = " FROM Proyecto"
    Rem by nan
    
    Rem    If Planta.Text = "Planta I" Then
    Rem     sql3 = "where codigo >1000 and codigo <2000 "
    Rem     End If
    
    
    
    
    Rem    spProyecto = sql1 + Sql2
    Rem    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    Rem    If rstProyecto.RecordCount > 0 Then
    Rem        rstProyecto.MoveLast
    Rem        ZCodigo = IIf(IsNull(rstProyecto!CodigoMayor), "0", rstProyecto!CodigoMayor)
    Rem        Codigo.Text = ZCodigo + 1
    Rem        rstProyecto.Close
    Rem    End If
        
    Rem    If Val(Codigo.Text) = 0 Then
    Rem        Codigo.Text = "1"
    Rem    End If
    Rem    Codigo.Text = " "
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            sql2 = " FROM Proyecto"
            sql3 = " Order by Proyecto.Codigo"
            spProyecto = sql1 + sql2 + sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstProyecto!descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstProyecto!descripcion, aa, WEspacios) Then
                                    Lugarayuda = Lugarayuda + 1
                                    pantalla.Row = Lugarayuda
                                    pantalla.Col = 1
                                    pantalla.Text = rstProyecto!codigo
                                    pantalla.Col = 2
                                    pantalla.Text = rstProyecto!descripcion
                                    IngresaItem = rstProyecto!codigo
                                    windice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
                
        Case 1
            sql1 = "Select *"
            sql2 = " FROM SectorInve"
            sql3 = " Order by SectorInve.Codigo"
            spSectorInve = sql1 + sql2 + sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                With rstSectorInve
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstSectorInve!descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstSectorInve!descripcion, aa, WEspacios) Then
                                    Lugarayuda = Lugarayuda + 1
                                    pantalla.Row = Lugarayuda
                                    pantalla.Col = 1
                                    pantalla.Text = rstSectorInve!codigo
                                    pantalla.Col = 2
                                    pantalla.Text = rstSectorInve!descripcion
                                    IngresaItem = rstSectorInve!codigo
                                    windice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSectorInve.Close
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
    Opcion.AddItem "Proyectos"
    Opcion.AddItem "Sectores"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Sector_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyectos"
    Opcion.AddItem "Sectores"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    pantalla.Clear
    pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            pantalla.FixedCols = 1
            pantalla.Cols = 3
            pantalla.FixedRows = 1
            pantalla.Rows = 10001
        Case 1
            pantalla.FixedCols = 1
            pantalla.Cols = 3
            pantalla.FixedRows = 1
            pantalla.Rows = 10001
    End Select
    
    pantalla.ColWidth(0) = 200
    pantalla.Row = 0
    
    Select Case XIndice
        Case 0
            For Ciclo = 1 To pantalla.Cols - 1
                pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        pantalla.Text = "Proyecto"
                        pantalla.ColWidth(Ciclo) = 1000
                        pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        pantalla.Text = "Nombre"
                        pantalla.ColWidth(Ciclo) = 6000
                        pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
        Case 1
            For Ciclo = 1 To pantalla.Cols - 1
                pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        pantalla.Text = "Centro"
                        pantalla.ColWidth(Ciclo) = 1000
                        pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        pantalla.Text = "Nombre"
                        pantalla.ColWidth(Ciclo) = 6000
                        pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    pantalla.Row = 0
    For Ciclo = 1 To pantalla.Cols - 1
        pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = pantalla.Text
        WTitulo(Ciclo).Left = pantalla.CellLeft + pantalla.Left
        WTitulo(Ciclo).Top = pantalla.CellTop + pantalla.Top
        WTitulo(Ciclo).Width = pantalla.CellWidth
        WTitulo(Ciclo).Height = pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To pantalla.Cols - 1
        WAncho = WAncho + pantalla.ColWidth(Ciclo)
    Next Ciclo
    pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = pantalla.Font.Name
    Font.Size = pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    pantalla.AllowUserResizing = flexResizeBoth
    
    pantalla.Col = 1
    pantalla.Row = 1
    
End Sub


Sub Ingresa_clave2()
wClave2.Text = ""
    xclave2.Visible = True
    wClave2.SetFocus
End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
 Rem       If UCase(Trim(WClave)) = UCase(Trim(WClaveOperador)) Then
            If UCase(Trim(WClave)) <> "" Then
            XClave.Visible = False
            WGraba = "S"
            Call cmdAdd_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClave.SetFocus
        End If
    End If
End Sub





Private Sub wclave2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        wgraba2 = "N"
        If UCase(Trim(wClave2)) = UCase(Trim(WWClaveOperador)) Then
            xclave2.Visible = False
            wgraba2 = "S"
            
         Call cmdDelete_Click
          Rem  Call cmdAdd_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            wClave2.SetFocus
        End If
    End If
End Sub

