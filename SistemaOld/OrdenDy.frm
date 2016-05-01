VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenDy 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Pedido de Importacion  de Dy"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.ComboBox Despachante 
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
      ItemData        =   "OrdenDy.frx":0000
      Left            =   9600
      List            =   "OrdenDy.frx":0002
      TabIndex        =   48
      Top             =   480
      Width           =   2175
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      Top             =   1200
      Width           =   9735
   End
   Begin VB.TextBox CondPago 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   1560
      Width           =   9735
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   840
      Width           =   9735
   End
   Begin VB.Frame EmailDy 
      Height          =   4335
      Left            =   480
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox ZTexto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   40
         Text            =   "OrdenDy.frx":0004
         Top             =   1080
         Width           =   9735
      End
      Begin VB.TextBox ZAsunto 
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   38
         Text            =   " "
         Top             =   720
         Width           =   9735
      End
      Begin VB.TextBox ZEmail 
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   36
         Text            =   " "
         Top             =   360
         Width           =   9735
      End
      Begin VB.Image OkEmail 
         Height          =   480
         Left            =   4800
         MouseIcon       =   "OrdenDy.frx":0006
         MousePointer    =   99  'Custom
         Picture         =   "OrdenDy.frx":0310
         ToolTipText     =   "Ejecyta el Proceso"
         Top             =   3720
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "Texto"
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
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Asunto"
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
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Email"
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
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
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
      Left            =   6240
      TabIndex        =   32
      Top             =   480
      Width           =   1935
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   31
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
      Index           =   4
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3600
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3480
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3480
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
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3600
      Width           =   375
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
      Left            =   6000
      MaxLength       =   11
      TabIndex        =   25
      Text            =   " "
      Top             =   120
      Width           =   1455
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
      Left            =   10080
      TabIndex        =   23
      Top             =   120
      Width           =   1695
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
      Left            =   8520
      TabIndex        =   21
      Top             =   120
      Width           =   1095
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   18
      Text            =   " "
      Top             =   480
      Width           =   1455
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
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3600
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   15
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
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton AgregaRenglon 
      Caption         =   "Agrega Renglon"
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
      Left            =   10560
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox WTexto1 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Ayuda 
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
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11280
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   6480
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
      Height          =   1980
      ItemData        =   "OrdenDy.frx":0752
      Left            =   120
      List            =   "OrdenDy.frx":0759
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label13 
      Caption         =   "Despachante"
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
      Left            =   8280
      TabIndex        =   47
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Cond. Pago"
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
      Left            =   240
      TabIndex        =   46
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Forwarder"
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
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Observ."
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
      Left            =   240
      TabIndex        =   41
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   5280
      TabIndex        =   33
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Left            =   7560
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   9720
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label23 
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
      Left            =   5280
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   20
      Top             =   480
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
      Left            =   3120
      TabIndex        =   19
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "OrdenDy.frx":0767
      MousePointer    =   99  'Custom
      Picture         =   "OrdenDy.frx":0A71
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "OrdenDy.frx":12B3
      MousePointer    =   99  'Custom
      Picture         =   "OrdenDy.frx":15BD
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "OrdenDy.frx":1DFF
      MousePointer    =   99  'Custom
      Picture         =   "OrdenDy.frx":2109
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "OrdenDy.frx":294B
      MousePointer    =   99  'Custom
      Picture         =   "OrdenDy.frx":2C55
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Pedido"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgOrdenDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim rsOrden As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim XPaso As String
Dim Renglon As Integer
Dim ZCodigo As String
Dim WRenglon As String
Dim ZDespachante As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private WGraba As String
Dim ZVector(100, 2) As String


Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "P.Terminados"
     Opcion.Visible = True
     
End Sub


Private Sub OkEmail_Click()
    EmailDy.Visible = False
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor"
            spProveedor = Sql1 + Sql2 + Sql3
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                With RstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = RstProveedor!Proveedor + " " + RstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = RstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                RstProveedor.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
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
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                Sql1 = "Select *"
                Sql2 = " FROM Proveedor"
                Sql3 = " Order by Proveedor"
                spProveedor = Sql1 + Sql2 + Sql3
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    With RstProveedor
                        .MoveFirst
                        Do
                            If .EOF = False Then
                        
                                Da = Len(RstProveedor!Nombre) - WEspacios
                                For Aaa = 1 To Da
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(RstProveedor!Nombre, Aaa, WEspacios) Then
                                        IngresaItem = RstProveedor!Proveedor + " " + RstProveedor!Nombre
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = RstProveedor!Proveedor
                                        WIndice.AddItem IngresaItem
                                    End If
                                Next Aaa
                            
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    RstProveedor.Close
                End If
                
            Case 2
                Sql1 = "Select *"
                Sql2 = " FROM Articulo"
                Sql3 = " Order by Codigo"
                spArticulo = Sql1 + Sql2 + Sql3
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    With rstArticulo
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                Da = Len(rstArticulo!Descripcion) - WEspacios
                                For Aaa = 1 To Da
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstArticulo!Codigo
                                        WIndice.AddItem IngresaItem
                                    End If
                                Next Aaa
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstArticulo.Close
                End If
                
            Case Else
        End Select
    
    End If

End Sub

Private Sub cmdClose1_Click()
Rem    Call Limpia_Click
    PrgOrdenDy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    WGraba = "S"
    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else

        ZCodigo = Str$(Val(Codigo.Text) + 900000)
        
        ZPasa = "S"
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZCodigo + "'"
        ZSql = ZSql + " Order by Orden.Clave"
    
        rsOrden = ZSql
        Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        If rstOrden!Recibida <> 0 Then
                            ZPasa = "N"
                        End If
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
        
        If ZPasa = "N" Then
            m$ = "No se puede modificar el pedido ya que el mismo ya fue utilizado"
            a% = MsgBox(m$, 0, "Pedidos de Importacion")
            Exit Sub
        End If
        
        Erase ZVector
        ZLugar = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZCodigo + "'"
        ZSql = ZSql + " Order by Orden.Clave"
    
        rsOrden = ZSql
        Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        
                        ZVector(ZLugar, 1) = rstOrden!Articulo
                        ZVector(ZLugar, 2) = Str$(rstOrden!Cantidad)
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
        
        For ZCiclo = 1 To ZLugar
        
            ZArticulo = ZVector(ZCiclo, 1)
            ZCantidad = ZVector(ZCiclo, 2)
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Pedido = Pedido - " + "'" + ZCantidad + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + ZArticulo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        Next ZCiclo
        
        
        
        
        
        
        Sql1 = "DELETE Orden"
        Sql2 = " Where Orden = " + "'" + ZCodigo + "'"
        rsOrden = Sql1 + Sql2
        Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        WRenglon = 0
    
        For iRow = 1 To 100
        
            If WVector1.TextMatrix(iRow, 1) <> "" Then
    
                Articulo = WVector1.TextMatrix(iRow, 1)
                Cantidad = WVector1.TextMatrix(iRow, 3)
                Precio = WVector1.TextMatrix(iRow, 4)
                Entrega = WVector1.TextMatrix(iRow, 5)
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
                
                ZCodigo = Str$(Val(Codigo.Text) + 900000)
                Auxi1 = ZCodigo
                Call Ceros(Auxi1, 6)
        
                WClave = Auxi1 + Auxi
                
                ZClave = WClave
                ZOrden = ZCodigo
                ZRenglon = Str$(WRenglon)
                ZFecha = Fecha.Text
                ZProveedor = Proveedor.Text
                ZArticulo = Articulo
                ZCantidad = Cantidad
                ZPrecio = Precio
                ZFecha1 = Entrega
                ZFecha2 = Entrega
                ZCondicion = ""
                ZRecibida = "0"
                ZSaldo = Cantidad
                ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZLiberada = "0"
                Zdevuelta = "0"
                ZFechaEntrega = Entrega
                ZDate = ""
                ZOrdFecha2 = Right$(Entrega, 4) + Mid$(Entrega, 4, 2) + Left$(Entrega, 2)
                ZMoneda = Str$(Moneda.ListIndex)
                ZTipo = ""
                ZDerechos = ""
                ZOrigen = Origen.Text
                ZCarpeta = ""
                ZImpresion = "N"
                ZLeyenda = Str$(Leyenda.ListIndex)
                ZPedidoImpo = Codigo.Text
                ZFechaImpo = Fecha.Text
                ZOrdFechaImpo = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZTipoImpo = Str$(TipoImpo.ListIndex)
                ZDespachante = Str$(Despachante.ListIndex)
                ZObservaciones = Observaciones.Text
                ZObservaII = ObservaII.Text
                ZCondPago = CondPago.Text
                ZPosicion = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Posicion"
                ZSql = ZSql + " Where Posicion.Articulo = " + "'" + Articulo + "'"
                rsPosicion = ZSql
                Set rstPosicion = db.OpenRecordset(rsPosicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstPosicion.RecordCount > 0 Then
                    ZPosicion = rstPosicion!Posicion
                    rstPosicion.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Orden ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Proveedor ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Precio ,"
                ZSql = ZSql + "Fecha1 ,"
                ZSql = ZSql + "Fecha2 ,"
                ZSql = ZSql + "Condicion ,"
                ZSql = ZSql + "Recibida ,"
                ZSql = ZSql + "Saldo ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Liberada ,"
                ZSql = ZSql + "Devuelta ,"
                ZSql = ZSql + "FechaEntrega ,"
                ZSql = ZSql + "WDate ,"
                ZSql = ZSql + "OrdFecha2 ,"
                ZSql = ZSql + "Moneda ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Carpeta ,"
                ZSql = ZSql + "Derechos ,"
                ZSql = ZSql + "Origen ,"
                ZSql = ZSql + "Impresion ,"
                ZSql = ZSql + "Leyenda ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "ObservaII ,"
                ZSql = ZSql + "CondPago ,"
                ZSql = ZSql + "PedidoImpo ,"
                ZSql = ZSql + "FechaImpo ,"
                ZSql = ZSql + "OrdFechaImpo ,"
                ZSql = ZSql + "Despachante ,"
                ZSql = ZSql + "Posicion ,"
                ZSql = ZSql + "TipoImpo )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZClave + "',"
                ZSql = ZSql + "'" + ZOrden + "',"
                ZSql = ZSql + "'" + ZRenglon + "',"
                ZSql = ZSql + "'" + ZFecha + "',"
                ZSql = ZSql + "'" + ZProveedor + "',"
                ZSql = ZSql + "'" + ZArticulo + "',"
                ZSql = ZSql + "'" + ZCantidad + "',"
                ZSql = ZSql + "'" + ZPrecio + "',"
                ZSql = ZSql + "'" + ZFecha1 + "',"
                ZSql = ZSql + "'" + ZFecha2 + "',"
                ZSql = ZSql + "'" + ZCondicion + "',"
                ZSql = ZSql + "'" + ZRecibida + "',"
                ZSql = ZSql + "'" + ZSaldo + "',"
                ZSql = ZSql + "'" + ZFechaOrd + "',"
                ZSql = ZSql + "'" + ZLiberada + "',"
                ZSql = ZSql + "'" + Zdevuelta + "',"
                ZSql = ZSql + "'" + ZFechaEntrega + "',"
                ZSql = ZSql + "'" + ZDate + "',"
                ZSql = ZSql + "'" + ZOrdFecha2 + "',"
                ZSql = ZSql + "'" + ZMoneda + "',"
                ZSql = ZSql + "'" + ZTipo + "',"
                ZSql = ZSql + "'" + ZCarpeta + "',"
                ZSql = ZSql + "'" + ZDerechos + "',"
                ZSql = ZSql + "'" + ZOrigen + "',"
                ZSql = ZSql + "'" + ZImpresion + "',"
                ZSql = ZSql + "'" + ZLeyenda + "',"
                ZSql = ZSql + "'" + ZObservaciones + "',"
                ZSql = ZSql + "'" + ZObservaII + "',"
                ZSql = ZSql + "'" + ZCondPago + "',"
                ZSql = ZSql + "'" + ZPedidoImpo + "',"
                ZSql = ZSql + "'" + ZFechaImpo + "',"
                ZSql = ZSql + "'" + ZOrdFechaImpo + "',"
                ZSql = ZSql + "'" + ZDespachante + "',"
                ZSql = ZSql + "'" + ZPosicion + "',"
                ZSql = ZSql + "'" + ZTipoImpo + "')"
                
                rsOrden = ZSql
                Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Articulo SET "
                ZSql = ZSql & "Pedido = Pedido + " + "'" + ZCantidad + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + ZArticulo + "'"
                
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            End If
            
        Next iRow
        
        T$ = "Orden de Compra"
        m$ = "Desea imprimir la Orden de Compra"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WEmail = RstProveedor!EMail
                RstProveedor.Close
            End If
        
            Listado.WindowTitle = "Emision de Orden de Compra"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height

            ZCodigo = Str$(Val(Codigo.Text) + 900000)
            Listado.GroupSelectionFormula = "{Orden.Orden} in " + ZCodigo + " to " + ZCodigo
            Listado.Destination = 1
            Rem Listado.Destination = 0
    
            Listado.ReportFileName = "ImpreOrdenDy.rpt"
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Moneda, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.Observaciones, Orden.CondPago, Orden.ObservaII, Orden.Despachante, Orden.Posicion, " _
                        + "Proveedor.Nombre, " _
                        + "Articulo.Descripcion, Articulo.CodigoDy " _
                        + "From " _
                        + DSQ + ".dbo.Orden Orden, " _
                        + DSQ + ".dbo.Proveedor Proveedor, " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where " _
                        + "Orden.Proveedor = Proveedor.Proveedor AND " _
                        + "Orden.Articulo = Articulo.Codigo AND " _
                        + "Orden.Orden >= " + ZCodigo + " AND " _
                        + "Orden.Orden <= " + ZCodigo
                            
            Listado.Connect = Connect()
            Listado.Action = 1
            Codigo.SetFocus
            
        End If
        
        T$ = "Orden de Compra"
        m$ = "Desea enviar la O/C via email al proveedor"
        Respuesta% = MsgBox(m$, 256 + 4, T$)
        If Respuesta% = 6 Then
        
            Rem spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            Rem If RstProveedor.RecordCount > 0 Then
            Rem     WEmail = RstProveedor!EMail
            Rem     RstProveedor.Close
            Rem End If
        
            Rem ZEmail.Text = WEmail
            Rem ZAsunto.Text = "PEDIDO NRO.: " + Codigo.Text
            Rem ZTexto.Text = "Se remite por la presente el pedido nro. :  " + Codigo.Text
            
            Rem EmailDy.Visible = True
            Rem ZEmail.SetFocus
            Rem by nan
            Call OkE41_Click
    
        End If
        
        Call Limpia_Click

        WVector1.Col = 1
        WVector1.Row = 1
        
        Codigo.SetFocus
        
    End If
        
End Sub


Private Sub OkE41_Click()

    Rem by nan
  Rem  MiRuta = CurDir + "\"
 Rem   MiRutaII = Left$(CurDir, 1)
    Rem fin by nan
    
    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

            
    ZCodigo = Str$(Val(Codigo.Text) + 900000)
    Listado.GroupSelectionFormula = "{Orden.Orden} in " + ZCodigo + " to " + ZCodigo
    Rem Listado.Destination = 1
    Listado.Destination = 0
    Rem Listado.Destination = 3
    Listado.PrintFileType = crptWinWord
    
    Listado.ReportFileName = "ImpreOrdenDy.rpt"
            
    Rem Listado.EMailToList = ZEmail.Text
    Rem Listado.EMailSubject = ZAsunto.Text
    Rem Listado.EMailMessage = ZTexto.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Orden.OrdenOrden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fechaentrega, Orden.Moneda, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.Observaciones, Orden.CondPago, Orden.ObservaII, Orden.Despachante, Orden.Posicion, " _
                        + "Proveedor.Nombre, " _
                        + "Articulo.Descripcion, Articulo.CodigoDy " _
                        + "From " _
                        + DSQ + ".dbo.Orden Orden, " _
                        + DSQ + ".dbo.Proveedor Proveedor, " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where " _
                        + "Orden.Proveedor = Proveedor.Proveedor AND " _
                        + "Orden.Articulo = Articulo.Codigo AND " _
                        + "Orden.Orden >= " + ZCodigo + " AND " _
                        + "Orden.Orden <= " + ZCodigo
                            
    Listado.Connect = Connect()
    Listado.Action = 1
            
            
    Rem Listado.EMailToList = "coelho.carlos@dystar.com"
    Rem Listado.EMailSubject = ZAsunto.Text
    Rem Listado.EMailMessage = ZTexto.Text
    
    Rem DbConnect = db.Connect
    Rem DSQ = getDatabase(DbConnect)
    Rem Listado.SQLQuery = "SELECT Orden.OrdenOrden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fechaentrega, Orden.Moneda, Orden.Origen, Orden.Leyenda, Orden.TipoImpo, Orden.Observaciones, Orden.CondPago, " _
    REM                     + "Proveedor.Nombre, " _
    REM                     + "Articulo.Descripcion, Articulo.CodigoDy " _
    REM                     + "From " _
    REM                     + DSQ + ".dbo.Orden Orden, " _
    REM                     + DSQ + ".dbo.Proveedor Proveedor, " _
    REM                     + DSQ + ".dbo.Articulo Articulo " _
    REM                     + "Where " _
    REM                     + "Orden.Proveedor = Proveedor.Proveedor AND " _
    REM                     + "Orden.Articulo = Articulo.Codigo AND " _
    REM                     + "Orden.Orden >= " + ZCodigo + " AND " _
    REM                     + "Orden.Orden <= " + ZCodigo
                            
    Rem Listado.Connect = Connect()
    Rem Listado.Action = 1
            
            
    Rem EmailDy.Visible = False
    Rem Call Limpia_Click

    Rem WVector1.Col = 1
    Rem WVector1.Row = 1
        
    Rem Codigo.SetFocus

    Rem by nnan
 Rem   ChDrive MiRutaII
 Rem   ChDir MiRuta
    Rem end nan
    
End Sub


Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Observaciones.Text = ""
    ObservaII.Text = ""
    CondPago.Text = ""
    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Origen.Text = ""
    Rem Carpeta.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Leyenda.ListIndex = 0
    TipoImpo.ListIndex = 0
    Despachante.ListIndex = 0
    Moneda.ListIndex = 0
    
    Renglon = 0
    WGraba = ""
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Sql1 = "Select Max(Orden) as [CodigoMayor]"
    Sql2 = " FROM Orden"
    spOrden = Sql1 + Sql2
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        rstOrden.MoveLast
        WCodigoMayor = IIf(IsNull(rstOrden!CodigoMayor), "0", rstOrden!CodigoMayor)
        XCodigo = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstOrden.Close
            Else
        XCodigo = "1"
    End If
    
    If Val(XCodigo) < 900000 Then
        XCodigo = "1"
            Else
        XCodigo = Mid$(Str$(Val(XCodigo) - 900000), 2, 8)
    End If
    Codigo.Text = XCodigo
    
    Codigo.SetFocus
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
        Case 1
            Indice = Pantalla.ListIndex
            XArticulo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + XArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstArticulo!Codigo)
                WVector1.Col = 2
                WVector1.Text = Trim(rstArticulo!Descripcion)
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Despachante.Clear
    
    Despachante.AddItem "Grupo Operadores"
    Despachante.AddItem "Perera"
    Despachante.AddItem "Bellavich Pablo"
    
    Despachante.ListIndex = 0

    Leyenda.Clear
    
    Leyenda.AddItem ""
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CIF"
    Leyenda.AddItem "CFR"
    Leyenda.AddItem "CPT"
    Leyenda.AddItem "EXW"
    Leyenda.AddItem "FCA"
    
    Leyenda.ListIndex = 0
    
    TipoImpo.Clear
    
    TipoImpo.AddItem ""
    TipoImpo.AddItem "Maritimo"
    TipoImpo.AddItem "Terrestre"
    TipoImpo.AddItem "Aereo"
    
    TipoImpo.ListIndex = 0
    
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"

    Moneda.ListIndex = 0
    
    
    WVector1.Col = 1
    WVector1.Row = 1

    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Origen.Text = ""
    Rem Carpeta.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Leyenda.ListIndex = 0
    TipoImpo.ListIndex = 0
    Despachante.ListIndex = 0
    Moneda.ListIndex = 0
    Observaciones.Text = ""
    ObservaII.Text = ""
    CondPago.Text = ""
    
    Sql1 = "Select Max(Orden) as [CodigoMayor]"
    Sql2 = " FROM Orden"
    spOrden = Sql1 + Sql2
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        rstOrden.MoveLast
        WCodigoMayor = IIf(IsNull(rstOrden!CodigoMayor), "0", rstOrden!CodigoMayor)
        XCodigo = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstOrden.Close
            Else
        XCodigo = "1"
    End If
    
    If Val(XCodigo) < 900000 Then
        XCodigo = "1"
            Else
        XCodigo = Mid$(Str$(Val(XCodigo) - 900000), 2, 8)
    End If
    Codigo.Text = XCodigo

    WGraba = ""
    Renglon = 0
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub




Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZCodigo = Str$(Val(Codigo.Text) + 900000)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden = " + "'" + ZCodigo + "'"
    ZSql = ZSql + " Order by Orden.Clave"
    
    rsOrden = ZSql
    Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstOrden!Fecha
                    Proveedor.Text = rstOrden!Proveedor
                    Origen.Text = rstOrden!Origen
                    Rem Carpeta.Text = rstOrden!Carpeta
                    Leyenda.ListIndex = rstOrden!Leyenda
                    TipoImpo.ListIndex = rstOrden!TipoImpo
                    Moneda.ListIndex = rstOrden!Moneda
                    Observaciones.Text = IIf(IsNull(rstOrden!Observaciones), "", rstOrden!Observaciones)
                    ObservaII.Text = IIf(IsNull(rstOrden!ObservaII), "", rstOrden!ObservaII)
                    CondPago.Text = IIf(IsNull(rstOrden!CondPago), "", rstOrden!CondPago)
                    ZDespachante = IIf(IsNull(rstOrden!Despachante), "0", rstOrden!Despachante)
                    Despachante.ListIndex = Val(ZDespachante)
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = rstOrden!Articulo
                    
                    WVector1.Col = 3
                    WVector1.Text = Str$(rstOrden!Cantidad)
            
                    WVector1.Col = 4
                    WVector1.Text = Str$(rstOrden!Precio)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 5
                    WVector1.Text = rstOrden!Fecha1
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        WArticulo = WVector1.TextMatrix(Ciclo, 1)
        
        Sql1 = "Select *"
        Sql2 = " FROM Articulo"
        Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
        spArticulo = Sql1 + Sql2 + Sql3
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.TextMatrix(Ciclo, 2) = Trim(rstArticulo!Descripcion)
            rstArticulo.Close
        End If
        
    Next Ciclo
    
    Sql1 = "Select *"
    Sql2 = " FROM Proveedor"
    Sql3 = " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = Sql1 + Sql2 + Sql3
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = Trim(RstProveedor!Nombre)
        RstProveedor.Close
    End If
    
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCodigo = Str$(Val(Codigo.Text) + 900000)
        Sql1 = "Select *"
        Sql2 = " FROM Orden"
        Sql3 = " Where Orden.Orden = " + "'" + ZCodigo + "'"
        rsOrden = Sql1 + Sql2 + Sql3
        Set rstOrden = db.OpenRecordset(rsOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            rstOrden.Close
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Origen.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Origen.Text = ""
    End If
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Leyenda.SetFocus
    End If
    If KeyAscii = 27 Then
        Origen.Text = ""
    End If
End Sub

Private Sub Leyenda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            Moneda.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Moneda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoImpo.SetFocus
    End If
End Sub

Private Sub TipoImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub ObservaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CondPago.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII.Text = ""
    End If
End Sub

Private Sub CondPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        CondPago.Text = ""
    End If
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            WArticulo = WVector1.Text
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = Trim(rstArticulo!Descripcion)
                rstArticulo.Close
                
                WPrecio = 0
                WFecha = ""
                
                XEmpresa = WEmpresa
        
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                spCotiza = "ListaCotizaProveedor " + "'" + Proveedor.Text + "'"
                Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
            
                If rstCotiza.RecordCount > 0 Then
                    With rstCotiza
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If WArticulo = rstCotiza!Articulo Then
                                    If rstCotiza!FechaOrd > WFecha Then
                                        WPrecio = rstCotiza!Precio
                                        WFecha = rstCotiza!FechaOrd
                                    End If
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCotiza.Close
                End If
                
                WVector1.Col = 4
                WVector1.Text = Str$(WPrecio)
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                WVector1.Col = 2
    
                Call Conecta_Empresa
                
                    Else
                WControl = "N"
            End If
            
        Case 4
            WArticulo = WVector1.TextMatrix(WVector1.Row, 1)
            WPrecio = 0
                
            XEmpresa = WEmpresa
        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spCotiza = "ListaCotizaProveedor " + "'" + Proveedor.Text + "'"
            Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstCotiza.RecordCount > 0 Then
                With rstCotiza
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If WArticulo = rstCotiza!Articulo Then
                                If rstCotiza!FechaOrd > WFecha Then
                                    WPrecio = rstCotiza!Precio
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCotiza.Close
            End If
            
            Call Conecta_Empresa
            
            Rem If Val(WVector1.Text) <> WPrecio Then
            
                T$ = "Pedido de Importaciones"
                m$ = "Desea registrar la cotizacion nueva"
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
                
                    XEmpresa = WEmpresa
        
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                    WCotiza = 1
                    spCotiza = "ListaCotizaNumero"
                    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCotiza.RecordCount > 0 Then
                        With rstCotiza
                            .MoveLast
                            WCotiza = rstCotiza!Cotiza + 1
                        End With
                        rstCotiza.Close
                    End If

                    XRenglon = 1
    
                    Auxi = Str$(XRenglon)
                    Call Ceros(Auxi, 2)
                    Auxi1 = Str$(WCotiza)
                    Call Ceros(Auxi1, 6)
                        
                    WCot = Str$(WCotiza)
                    WRenglon = Str$(XRenglon)
                    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    WProveedor = Proveedor.Text
                    WArticulo = WArticulo
                    WPrecio = WVector1.Text
                    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WCondicion = ""
                    WObservaciones = ""
                    WClave = Auxi1 + Auxi
                    WDate = Date$
                    WMoneda = "0"
        
                    XParam = "'" + WClave + "','" _
                                 + WCot + "','" _
                                 + WRenglon + "','" _
                                 + WFecha + "','" _
                                 + WProveedor + "','" _
                                 + WArticulo + "','" _
                                 + WPrecio + "','" _
                                 + WCondicion + "','" _
                                 + WObservaciones + "','" _
                                 + WFechaord + "','" _
                                 + WDate + "','" _
                                 + WMoneda + "'"
                    
                    spCotiza = "AltaCotizaII " + XParam
                    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    
                    Call Conecta_Empresa
    
                End If
                
            Rem End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Articulo = WVector1.TextMatrix(iRow, 1)
        If Articulo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 0 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglon_Click()

    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
        WVector1.TextMatrix(iRow, 4) = WVector1.TextMatrix(iRow - 1, 4)
        WVector1.TextMatrix(iRow, 5) = WVector1.TextMatrix(iRow - 1, 5)
    Next iRow

    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = ""
    WVector1.TextMatrix(Hasta, 4) = ""
    WVector1.TextMatrix(Hasta, 5) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 6
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "F.Entrega"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
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

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
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
        ZGrabaIII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGrabaIII = IIf(IsNull(rstOperador!GrabaIII), "", rstOperador!GrabaIII)
            rstOperador.Close
        End If
        
        If ZGrabaIII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Ingreso de Procesos de Fabricacion")
            WClave.SetFocus
        End If
        
    End If
End Sub

