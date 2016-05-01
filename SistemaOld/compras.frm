VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompras 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Comprobantes de Proveedores"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   990
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   15180
   Begin VB.Frame PantaDiscrimina 
      Height          =   1215
      Left            =   120
      TabIndex        =   96
      Top             =   6720
      Visible         =   0   'False
      Width           =   5175
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   1560
         TabIndex        =   100
         Top             =   240
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
      Begin VB.TextBox WTexto12 
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
         Left            =   480
         TabIndex        =   102
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox WTexto22 
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
         Left            =   840
         TabIndex        =   101
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox WCombo12 
         Height          =   315
         Left            =   2400
         TabIndex        =   99
         Top             =   240
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton CierraDiscrimina 
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4680
         TabIndex        =   98
         Top             =   5280
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   4935
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   8705
         _Version        =   327680
         BackColor       =   12648384
      End
   End
   Begin VB.CommandButton Apertura 
      Caption         =   "Apertura"
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
      Left            =   11160
      TabIndex        =   103
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox SoloIva 
      Caption         =   "Solo Iva"
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
      Left            =   9840
      TabIndex        =   95
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Iva105 
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
      Left            =   8520
      MaxLength       =   15
      TabIndex        =   93
      Text            =   " "
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame PantaPyme 
      Height          =   2295
      Left            =   3720
      TabIndex        =   86
      Top             =   4080
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox MesCuota 
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
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   92
         Text            =   " "
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox AnoCuota 
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   91
         Text            =   " "
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton CierraPyme 
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   90
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Cuotas 
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
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   87
         Text            =   " "
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha 1 Cuota"
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
         Left            =   360
         TabIndex        =   89
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Cuotas"
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
         Left            =   360
         TabIndex        =   88
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame BusquedaNro 
      Height          =   1935
      Left            =   1560
      TabIndex        =   70
      Top             =   3720
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton CerrarBusquedaNro 
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
         Height          =   345
         Left            =   3480
         TabIndex        =   83
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox NumeroII 
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
         MaxLength       =   8
         TabIndex        =   80
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox PuntoII 
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
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   79
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox LetraII 
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
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   77
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox TipoCompII 
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
         Left            =   1440
         TabIndex        =   75
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TipoII 
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
         Left            =   1200
         TabIndex        =   74
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox ProveedorII 
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
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   71
         Text            =   " "
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Numero"
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
         Left            =   5640
         TabIndex        =   82
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Punto"
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
         Left            =   3840
         TabIndex        =   81
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Letra"
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
         Left            =   2640
         TabIndex        =   78
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label24 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
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
         Left            =   240
         TabIndex        =   73
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label DesProveedorII 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   72
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton ConsultaII 
      Caption         =   "Consulta Nro factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9600
      TabIndex        =   84
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Remito 
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
      MaxLength       =   30
      TabIndex        =   69
      Text            =   " "
      Top             =   960
      Width           =   2535
   End
   Begin VB.CheckBox Rechazado 
      Caption         =   "Ch.Rech."
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
      Left            =   7320
      TabIndex        =   67
      Top             =   480
      Width           =   1215
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
      Height          =   1620
      Left            =   5280
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
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
      Height          =   2595
      ItemData        =   "compras.frx":0000
      Left            =   4560
      List            =   "compras.frx":0007
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   5520
      Width           =   375
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
      Left            =   1800
      TabIndex        =   62
      Top             =   4440
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   61
      Top             =   5040
      Visible         =   0   'False
      Width           =   390
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
      Left            =   2400
      TabIndex        =   60
      Top             =   4440
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   5160
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5040
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Despacho 
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
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   54
      Text            =   " "
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Cai 
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
      Left            =   9480
      MaxLength       =   14
      TabIndex        =   50
      Text            =   " "
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Paridad 
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
      Left            =   9840
      MaxLength       =   15
      TabIndex        =   49
      Text            =   " "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Pago 
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
      Left            =   6840
      TabIndex        =   47
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Tipo 
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
      Left            =   600
      TabIndex        =   45
      Text            =   " "
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox TipoComp 
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
      Left            =   840
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento1 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
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
   Begin VB.TextBox NroInterno 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
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
      Left            =   8520
      TabIndex        =   41
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Exento 
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
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   15
      Text            =   " "
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Iva27 
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
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   13
      Text            =   " "
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Iva21 
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
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   11
      Text            =   " "
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Ib 
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   44
      Text            =   " "
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Iva5 
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   12
      Text            =   " "
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Punto 
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
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   4
      Text            =   " "
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Numero 
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
      Left            =   6120
      MaxLength       =   8
      TabIndex        =   5
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Letra 
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
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame4 
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
      Height          =   615
      Left            =   3360
      TabIndex        =   28
      Top             =   2880
      Width           =   3495
      Begin VB.OptionButton Contado3 
         Caption         =   "Nacion"
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
         TabIndex        =   85
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Contado2 
         Caption         =   "Cta.Cte."
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
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Contado1 
         Caption         =   "Efectivo"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Periodo 
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   960
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
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
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
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   960
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
   Begin VB.TextBox Neto 
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
      Height          =   315
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   10
      Text            =   " "
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   3600
      MaxLength       =   11
      TabIndex        =   1
      Text            =   " "
      Top             =   0
      Width           =   1335
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
      Height          =   465
      Left            =   9600
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
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
      Height          =   465
      Left            =   8520
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7320
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7320
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   11760
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox VtoCai 
      Height          =   285
      Left            =   9840
      TabIndex        =   52
      Top             =   480
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3975
      Left            =   120
      TabIndex        =   56
      Top             =   3720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7011
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3000
      TabIndex        =   63
      Top             =   4440
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
   Begin VB.Label Label28 
      Caption         =   "Importe Iva 10.5%"
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
      Left            =   6840
      TabIndex        =   94
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label21 
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
      Left            =   6720
      TabIndex        =   68
      Top             =   960
      Width           =   975
   End
   Begin VB.Label TotalDebito 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   6600
      TabIndex        =   66
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label TotalCredito 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   8280
      TabIndex        =   65
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label20 
      Caption         =   "Despacho"
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
      Left            =   6840
      TabIndex        =   55
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Vto CAI"
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
      Left            =   8760
      TabIndex        =   53
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "C.A.I."
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
      Left            =   8760
      TabIndex        =   51
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Paridad"
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
      Left            =   8760
      TabIndex        =   48
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Forma de Pago"
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
      Left            =   5400
      TabIndex        =   46
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Interno"
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
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   1800
      TabIndex        =   42
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Importe Total"
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
      TabIndex        =   40
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Importe No Gravado"
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
      Left            =   3360
      TabIndex        =   39
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Importe Iva 27%"
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
      Left            =   3360
      TabIndex        =   38
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Importe Iva 21%"
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
      Left            =   3360
      TabIndex        =   37
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Importe Perc. I.B."
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
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Iva R.G. 3337"
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
      TabIndex        =   35
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Punto"
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
      Left            =   3240
      TabIndex        =   34
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Numero"
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
      Left            =   5040
      TabIndex        =   33
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Letra"
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
      Left            =   2040
      TabIndex        =   31
      Top             =   480
      Width           =   495
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
      Height          =   285
      Left            =   5040
      TabIndex        =   27
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de Iva"
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
      TabIndex        =   26
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha de vencimiento"
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
      TabIndex        =   25
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Importe Neto"
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
      TabIndex        =   23
      Top             =   1920
      Width           =   2175
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
      Left            =   2640
      TabIndex        =   22
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha de Emision"
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
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "PrgCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Dato As String
Private Auxi As String
Private WImpo As Double
Private WProveedor As String
Private SumaDebito As Double
Private SumaCredito As Double
Private Uno As Double
Private Dos As Double
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstImputac As Recordset
Dim spImputac As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim rstIvaCompAdicional As Recordset
Dim spIvaCompAdicional As String

Dim XParam As String
Dim cParam As String

Dim ZIva105 As Double

Dim ZNroRemito(100) As String
Dim ZNroOrden(100, 2) As String
Dim ZZRemito As String
Dim EmpresaTrabajo As String
Dim ZPyme As String

Dim ZZCuotas As Integer
Dim ZZMesCuota As Integer
Dim ZZAnoCuota As Integer
Dim ZZValorCuota As Double
Dim ZZIva As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Rem para el vector II

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String

Sub Calcula_total()

    WImpo = 0
    Call Format_datos
    
    If SoloIva.Value <> 1 Then
        Dato = Neto.Text
        If Val(Dato) <> 0 Then
            WImpo = WImpo + Val(Dato)
        End If
    End If
    
    Dato = Iva21.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva5.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva27.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva105.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Ib.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Exento.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Total.Caption = WImpo
    Total.Caption = Pusing("#,###,###.##", Total.Caption)
    
End Sub

Sub Alinea_datos()
    Tipo.Text = Str$(TipoComp.ListIndex + 1)
    WTipo = Tipo.Text
    Call Ceros(WTipo, 2)
    Tipo.Text = WTipo
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    Punto.Text = WPunto
    WNumero = Numero.Text
    Call Ceros(WNumero, 8)
    Numero.Text = WNumero
    Letra.Text = Left$(Letra.Text, 1)
End Sub

Sub Imprime_Descripcion()
    With RstProveedor
        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            Desproveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
                Else
            Desproveedor.Caption = ""
        End If
    End With
End Sub

Sub Verifica_datos()
    If Val(Neto.Text) = 0 Then
        Neto.Text = "0"
    End If
    If Val(Iva21.Text) = 0 Then
        Iva21.Text = "0"
    End If
    If Val(Iva5.Text) = 0 Then
        Iva5.Text = "0"
    End If
    If Val(Iva27.Text) = 0 Then
        Iva27.Text = "0"
    End If
    If Val(Iva105.Text) = 0 Then
        Iva105.Text = "0"
    End If
    If Val(Ib.Text) = 0 Then
        Ib.Text = "0"
    End If
    If Val(Exento.Text) = 0 Then
        Exento.Text = "0"
    End If
    If Val(Total.Caption) = 0 Then
        Total.Caption = "0"
    End If
    If Val(Paridad.Text) = 0 Then
        Paridad.Text = "0"
    End If
End Sub

Sub Format_datos()
    If Val(Paridad.Text) <> 0 Then
        Paridad.Text = Pusing("#,###.####", Paridad.Text)
    End If
    If Val(Neto.Text) <> 0 Then
        Neto.Text = Pusing("#,###,###.##", Neto.Text)
    End If
    If Val(Iva21.Text) <> 0 Then
        Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
    End If
    If Val(Iva5.Text) <> 0 Then
        Iva5.Text = Pusing("#,###,###.##", Iva5.Text)
    End If
    If Val(Iva27.Text) <> 0 Then
        Iva27.Text = Pusing("#,###,###.##", Iva27.Text)
    End If
    If Val(Iva105.Text) <> 0 Then
        Iva105.Text = Pusing("#,###,###.##", Iva105.Text)
    End If
    If Val(Ib.Text) <> 0 Then
        Ib.Text = Pusing("#,###,###.##", Ib.Text)
    End If
    If Val(Exento.Text) <> 0 Then
        Exento.Text = Pusing("#,###,###.##", Exento.Text)
    End If
    Total.Caption = Pusing("#,###,###.##", Total.Caption)
End Sub

Sub Imprime_Datos()

    spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
            Proveedor.Text = rstIvaComp!Proveedor
            TipoComp.ListIndex = rstIvaComp!Tipo - 1
            Letra.Text = rstIvaComp!Letra
            Punto.Text = rstIvaComp!Punto
            Numero.Text = rstIvaComp!Numero
            Call Alinea_datos
            Fecha.Text = rstIvaComp!Fecha
            Vencimiento.Text = rstIvaComp!Vencimiento
            Vencimiento1.Text = rstIvaComp!Vencimiento1
            Periodo.Text = rstIvaComp!Periodo
            Neto.Text = Abs(rstIvaComp!Neto)
            Iva21.Text = Abs(rstIvaComp!Iva21)
            Iva5.Text = Abs(rstIvaComp!Iva5)
            Iva27.Text = Abs(rstIvaComp!Iva27)
            
            ZIva105 = IIf(IsNull(rstIvaComp!Iva105), "0", rstIvaComp!Iva105)
            Iva105.Text = Abs(ZIva105)
            
            Ib.Text = Abs(rstIvaComp!Ib)
            Exento.Text = Abs(rstIvaComp!Exento)
            Call Calcula_total
            Contado1.Value = False
            Contado2.Value = False
            Contado3.Value = False
            Select Case Val(rstIvaComp!Contado)
                Case 1
                    Contado1.Value = True
                Case 2
                    Contado2.Value = True
                Case 3
                    Contado3.Value = True
                Case Else
            End Select
            Paridad.Text = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
            Pago.ListIndex = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
            Cai.Text = IIf(IsNull(rstIvaComp!Cai), "", rstIvaComp!Cai)
            Cai.Text = Trim(Cai.Text)
            VtoCai.Text = IIf(IsNull(rstIvaComp!VtoCai), "  /  /    ", rstIvaComp!VtoCai)
            Despacho.Text = IIf(IsNull(rstIvaComp!Despacho), "", rstIvaComp!Despacho)
            Remito.Text = IIf(IsNull(rstIvaComp!Remito), "", rstIvaComp!Remito)
            
            ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
            Rechazado.Value = ZRechazado
            
            ZSoloIva = IIf(IsNull(rstIvaComp!SoloIva), "0", rstIvaComp!SoloIva)
            SoloIva.Value = ZSoloIva
            
            rstIvaComp.Close
            Call Format_datos
            Call Imprime_Descripcion
    End If
    
    Renglon = 0
    Call Limpia_Vector
        
    For A = 1 To 50
        
        WTipoMovi = "2"
        
        Auxi1 = NroInterno.Text
        Call Ceros(Auxi1, 6)
        XNroInterno = Auxi1
            
        Renglon = A
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        WRenglon = Auxi1$
            
        ClaveImputac = WTipoMovi + XNroInterno + WRenglon
            
        spImputac = "Consultaimputac " + "'" + ClaveImputac + "'"
        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
        If rstImputac.RecordCount > 0 Then
            
            WVector1.Row = Val(rstImputac!Renglon)
                
            WVector1.Col = 1
            WVector1.Text = rstImputac!Cuenta
                        
            WVector1.Col = 3
            WVector1.Text = Str$(rstImputac!Debito)
            WVector1.Text = Pusing("#,###,###.##", WVector1.Text)

            WVector1.Col = 4
            WVector1.Text = Str$(rstImputac!Credito)
            WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                
            WCuenta = rstImputac!Cuenta
            rstImputac.Close
                
            With rstCuenta
                spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End With
                
        End If
            
    Next A
    
    
    
    Renglon = 0
    Call Limpia_VectorII
        
    For A = 1 To 50
        
        Auxi = NroInterno.Text
        Call Ceros(Auxi, 8)
            
        Renglon = A
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
            
        ZZClave = Auxi + Auxi1
            
        ZSql = "Select *"
        ZSql = ZSql + " FROM IvaCompAdicional"
        ZSql = ZSql + " Where IvaCompAdicional.Clave = " + "'" + ZZClave + "'"
        spIvaCompAdicional = ZSql
        Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaCompAdicional.RecordCount > 0 Then
        
            WVector2.TextMatrix(A, 1) = rstIvaCompAdicional!Cuit
            WVector2.TextMatrix(A, 2) = rstIvaCompAdicional!Razon
            WVector2.TextMatrix(A, 3) = rstIvaCompAdicional!Tipo
            WVector2.TextMatrix(A, 4) = rstIvaCompAdicional!Letra
            WVector2.TextMatrix(A, 5) = Trim(rstIvaCompAdicional!Punto)
            WVector2.TextMatrix(A, 6) = Trim(rstIvaCompAdicional!Numero)
            WVector2.TextMatrix(A, 7) = rstIvaCompAdicional!Fecha
            WVector2.TextMatrix(A, 8) = Str$(rstIvaCompAdicional!Neto)
            WVector2.TextMatrix(A, 9) = Str$(rstIvaCompAdicional!Iva21)
            WVector2.TextMatrix(A, 10) = Str$(rstIvaCompAdicional!Iva27)
            WVector2.TextMatrix(A, 11) = Str$(rstIvaCompAdicional!Iva105)
            WVector2.TextMatrix(A, 12) = Str$(rstIvaCompAdicional!perceiva)
            WVector2.TextMatrix(A, 13) = Str$(rstIvaCompAdicional!perceib)
            WVector2.TextMatrix(A, 14) = Str$(rstIvaCompAdicional!Exento)
                        
            WVector2.TextMatrix(A, 8) = Pusing("###,###.##", WVector2.TextMatrix(A, 8))
            WVector2.TextMatrix(A, 9) = Pusing("###,###.##", WVector2.TextMatrix(A, 9))
            WVector2.TextMatrix(A, 10) = Pusing("###,###.##", WVector2.TextMatrix(A, 10))
            WVector2.TextMatrix(A, 11) = Pusing("###,###.##", WVector2.TextMatrix(A, 11))
            WVector2.TextMatrix(A, 12) = Pusing("###,###.##", WVector2.TextMatrix(A, 12))
            WVector2.TextMatrix(A, 13) = Pusing("###,###.##", WVector2.TextMatrix(A, 13))
            WVector2.TextMatrix(A, 14) = Pusing("###,###.##", WVector2.TextMatrix(A, 14))

            rstIvaCompAdicional.Close
                
            Rem ZSql = "Select *"
            Rem ZSql = ZSql + " FROM Proveedor"
            Rem ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + XXProveedor + "'"
            Rem spProveedor = ZSql
            Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            Rem If RstProveedor.RecordCount > 0 Then
            Rem     WVector2.TextMatrix(A, 2) = RstProveedor!Nombre
            Rem     RstProveedor.Close
            Rem End If
                
        End If
            
    Next A
    
    
    
    
    Call Calcula_Click
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub


Private Sub Apertura_Click()

    PantaDiscrimina.Height = 6000
    PantaDiscrimina.Left = 0
    PantaDiscrimina.Top = 240
    PantaDiscrimina.Width = 15000
    
    PantaDiscrimina.Visible = True

End Sub

Private Sub CierraDiscrimina_Click()
    
    For iRow = 1 To 50
        WVector1.TextMatrix(iRow, 1) = ""
        WVector1.TextMatrix(iRow, 2) = ""
        WVector1.TextMatrix(iRow, 3) = ""
        WVector1.TextMatrix(iRow, 4) = ""
    Next iRow
    
    ZZLugar = 0
    
    ZZTotal = 0
    ZZIva21 = 0
    ZZIva27 = 0
    ZZIva105 = 0
    ZZIva5 = 0
    ZZib = 0
    
    ZZLugar = 0
    
    
    For iRow = 1 To 50
        
        WWCuit = WVector2.TextMatrix(iRow, 1)
        WWRazon = WVector2.TextMatrix(iRow, 2)
        WWTipo = WVector2.TextMatrix(iRow, 3)
        WWLetra = WVector2.TextMatrix(iRow, 4)
        WWPunto = WVector2.TextMatrix(iRow, 5)
        WWNumero = WVector2.TextMatrix(iRow, 6)
        WWFecha = WVector2.TextMatrix(iRow, 7)
        WWNeto = WVector2.TextMatrix(iRow, 8)
        WWIva21 = WVector2.TextMatrix(iRow, 9)
        WWIva27 = WVector2.TextMatrix(iRow, 10)
        WWIva105 = WVector2.TextMatrix(iRow, 11)
        WWPerceIva = WVector2.TextMatrix(iRow, 12)
        WWPerceIb = WVector2.TextMatrix(iRow, 13)
        WWExento = WVector2.TextMatrix(iRow, 14)
        
        WWTotal = Val(WWNeto) + Val(WWIva21) + Val(WWIva27) + Val(WWIva105) + Val(WWPerceIb) + Val(WWPerceIva) + Val(WWExento)
        
        If WWCuit <> "" Then
            ZZTotal = ZZTotal + WWTotal
            ZZIva21 = ZZIva21 + Val(WWIva21)
            ZZIva27 = ZZIva27 + Val(WWIva27)
            ZZIva105 = ZZIva105 + Val(WWIva105)
            ZZIva5 = ZZIva5 + Val(WWPerceIva)
            ZZib = ZZib + Val(WWPerceIb)
        End If
                                        
    Next iRow
    
    
    If ZZTotal <> 0 Then
        
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "2001"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = ""
        WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Str$(ZZTotal))
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
    
    End If
    
    If ZZIva21 <> 0 Then
        
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "151"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Str$(ZZIva21))
        WVector1.TextMatrix(ZZLugar, 4) = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
        
    End If
    
    If ZZIva27 <> 0 Then
    
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "151"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Str$(ZZIva27))
        WVector1.TextMatrix(ZZLugar, 4) = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
    
    End If
    
    If ZZIva105 <> 0 Then
    
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "151"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Str$(ZZIva105))
        WVector1.TextMatrix(ZZLugar, 4) = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
    
    End If
    
    If ZZIva5 <> 0 Then
        
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "152"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Str$(ZZIva5))
        WVector1.TextMatrix(ZZLugar, 4) = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
    
    End If
    
    If ZZib <> 0 Then
    
        ZZLugar = ZZLugar + 1
        
        WVector1.TextMatrix(ZZLugar, 1) = "164"
        WVector1.TextMatrix(ZZLugar, 2) = ""
        WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Str$(ZZib))
        WVector1.TextMatrix(ZZLugar, 4) = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
            rstCuenta.Close
        End If
        
    End If
        
    Call Calcula_Click
    
    PantaDiscrimina.Visible = False

End Sub

Private Sub cmdAdd_Click()


    Rem dada
    Rem dada
    Rem dada

    ZMes = Mid$(Periodo.Text, 4, 2)
    ZAno = Mid$(Periodo.Text, 7, 4)
    ZEstado = 1

    ZSql = "Select *"
    ZSql = ZSql + " FROM Cierre"
    ZSql = ZSql + " Where Cierre.Mes = " + "'" + ZMes + "'"
    ZSql = ZSql + " and Cierre.Ano = " + "'" + ZAno + "'"
    spCierre = ZSql
    Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
    If rstCierre.RecordCount > 0 Then
        ZEstado = rstCierre!Estado
        rstCierre.Close
    End If
    
    If ZEstado = 1 Then
        m$ = "El mes ya a sido cerrrado, no se puede ingresar ni modificar mas datos"
        A% = MsgBox(m$, 64, "Ingreso de Comprobantes")
        Exit Sub
    End If
    
    ZZPasaRemito = Remito.Text
    ZZPasaProveedor = Proveedor.Text
    Call Verifica_Pyme
    
    Rem ZPyme = "N"
    
    If ZPyme = "S" And Contado3.Value = False Then
        T$ = "Ingreso de comprobantes de Proveedores"
        m$ = "La Orden de Comrpa indica que se paga con Pyme Banco Nacion" + Chr$(13) + _
             "y difiere de la forma de pago informado en la carga del comprobante" + Chr$(13) + _
             "Desea continuar con la grabacion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% <> 6 Then
            Exit Sub
        End If
    End If
        
    If ZPyme = "S" Then
        If Val(Cuotas.Text) = 0 Then
            T$ = "Ingreso de comprobantes de Proveedores"
            m$ = "No se informo la cantidad de cuotas para la financiacion de la compra"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            Exit Sub
        End If
    End If
    
    If ZPyme = "S" Then
        If Val(MesCuota.Text) = 0 Or Val(AnoCuota.Text) = 0 Then
            T$ = "Ingreso de comprobantes de Proveedores"
            m$ = "No se informo la fecha de inicio para la financiacion de la compra"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            Exit Sub
        End If
    End If
        
    If Val(Punto.Text) = 0 Then
        T$ = "Ingreso de comprobantes de Proveedores"
        m$ = "No se informo Punto de Venta"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        Exit Sub
    End If
        
    If UCase(Letra.Text) = "C" Then
        If Val(Iva21.Text) <> 0 Or Val(Iva27.Text) <> 0 Or Val(Iva105.Text) <> 0 Or Val(Iva5.Text) <> 0 Or Val(Ib.Text) <> 0 Or Val(Exento.Text) <> 0 Then
            T$ = "Ingreso de comprobantes de Proveedores"
            m$ = "En facturas C el importe total debe ser informado en el campo Importe Neto"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            Exit Sub
        End If
    End If
        
        
    If Val(NroInterno.Text) = 0 Then
    
        If Val(Wempresa) = 1 Then
        
            spIvaComp = "ListaIvacompNumero"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                With rstIvaComp
                    .MoveLast
                    NroInterno.Text = rstIvaComp!NroInterno + 1
                End With
                rstIvaComp.Close
            End If
            
            Rem  m$ = "El numero interno asignado es " + NroInterno.Text
            Rem A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
                Else
                
            ZHasta = "119000"
            ZSql = ""
            ZSql = ZSql + "Select IvaComp.NroInterno"
            ZSql = ZSql + " FROM Ivacomp"
            ZSql = ZSql + " Where Ivacomp.NroInterno <= " + "'" + ZHasta + "'"
            ZSql = ZSql + " Order by Ivacomp.NroInterno"
            
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                With rstIvaComp
                    .MoveLast
                    NroInterno.Text = rstIvaComp!NroInterno + 1
                End With
                rstIvaComp.Close
            End If
            
            Rem m$ = "El numero interno asignado es " + NroInterno.Text
            Rem A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
        End If
        
    End If

    If Val(NroInterno.Text) <> 0 Then
    
        Tipo.Text = TipoComp.ListIndex + 1
            
        WPasa = "S"
        Call Verifica_datos
        
'        With rstProveedor
'            .Index = "Proveedor"
'            .Seek "=", Proveedor.Text
'            If .NoMatch = True Then
'                WPasa = "N"
'                m$ = "Codigo de Proveedor Incorrecto"
'                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
'            End If
'        End With
        
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de emision, formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de vencimiento (1), formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If

        Call Valida_fecha(Vencimiento1.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de vencimiento (2), formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If

        Call Valida_fecha(Periodo.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de Iva, formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
         
        If Val(Tipo.Text) < 1 Or Val(Tipo.Text) > 3 Then
           WPasa = "N"
           m$ = "Tipo de Comprobante Invalido"
           A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
            
        If Left$(Letra.Text, 1) <> "A" And Left$(Letra.Text, 1) <> "B" And Left$(Letra.Text, 1) <> "C" And Left$(Letra.Text, 1) <> "X" And Left$(Letra.Text, 1) <> "M" And Left$(Letra.Text, 1) <> "I" Then
            WPasa = "N"
            m$ = "Letra del Comprobante Invalido"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        If Pago.ListIndex = 0 Then
            WPasa = "N"
            m$ = "Clausula de Forma de pago no informada"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        If Pago.ListIndex = 2 Then
            If Val(Paridad.Text) = 0 Then
                WPasa = "N"
                m$ = "Paridad no informada"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            End If
        End If
        
        If Contado3.Value = False Then
            Call Alinea_datos
            ClaveCtaprv = Proveedor.Text + Letra.Text + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                If RstCtaPrv!Saldo <> RstCtaPrv!Total Then
                    m$ = "El Comprobante se encuentra total o parcialmente cancelado"
                    A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
                    WPasa = "N"
                End If
            End If
        End If
         Call Calcula_Click
        
        SumaDebito = 0
        SumaCredito = 0
        
        For iRow = 1 To 50
            Debito = WVector1.TextMatrix(iRow, 3)
            SumaDebito = SumaDebito + Val(Debito)
            Credito = WVector1.TextMatrix(iRow, 4)
            SumaCredito = SumaCredito + Val(Credito)
        Next iRow
                    
        Call Redondeo(SumaDebito)
        Call Redondeo(SumaCredito)
        
        If SumaDebito <> SumaCredito Then
        
            WPasa = "N"
            m$ = "Importe total de debito distinto al del credito"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
                Else
                
            Dato = Val(Total.Caption)
            Uno = Val(Total.Caption)
            Dato = SumaDebito
            Dato = Pusing("#,###,###.##", Dato)
            Dos = SumaDebito
            
            Call Redondeo(Uno)
            Call Redondeo(Dos)

            If Uno <> Dos Then
                WPasa = "N"
                m$ = "Importe del comprobante distinto a la imputacion contable"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            End If
            
        End If
        
        If WPasa = "S" Then
            Call Alinea_datos
            spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount = 0 Then
                Rem rstIvaComp.Close
                
                ZSql = "Select *"
                ZSql = ZSql + " FROM Ivacomp"
                ZSql = ZSql + " Where Ivacomp.Proveedor = " + "'" + Proveedor.Text + "'"
                ZSql = ZSql + " and Ivacomp.Tipo = " + "'" + WTipo + "'"
                ZSql = ZSql + " and Ivacomp.Letra = " + "'" + Letra.Text + "'"
                ZSql = ZSql + " and Ivacomp.Punto = " + "'" + WPunto + "'"
                ZSql = ZSql + " and Ivacomp.Numero = " + "'" + WNumero + "'"
                spIvaComp = ZSql
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                
                    ZZAA = rstIvaComp!NroInterno
                
                    rstIvaComp.Close
                    WPasa = "N"
                    m$ = "El comprobante ya se encuentra ingresado"
                    A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
                    
                End If
                
                    Else
                    
                rstIvaComp.Close
                
            End If
        End If
        
        WVector1.Col = 1
        WVector1.Row = 1
        
        If WPasa = "S" Then
    
            Call Alinea_datos

            spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount = 0 Then
            
                Call Verifica_datos
                
                Rem ALTA DE IVA COMPRAS
                
                XNroInterno = NroInterno.Text
                XProveedor = Proveedor.Text
                XTipo = Tipo.Text
                XLetra = Letra.Text
                XPunto = Punto.Text
                XNumero = Numero.Text
                XFecha = Fecha.Text
                Xvencimiento = Vencimiento.Text
                XVencimiento1 = Vencimiento1.Text
                XPeriodo = Periodo.Text
                XNeto = Neto.Text
                XIva21 = Iva21.Text
                XIva5 = Iva5.Text
                XIva27 = Iva27.Text
                XIva105 = Iva105.Text
                XIb = Ib.Text
                XExento = Exento.Text
                Select Case Val(Tipo.Text)
                    Case 1
                        XImpre = "FC"
                    Case 2
                        XImpre = "ND"
                    Case 3
                        XImpre = "NC"
                        XNeto = Str$(Val(Neto.Text) * -1)
                        XIva21 = Str$(Val(Iva21.Text) * -1)
                        XIva5 = Str$(Val(Iva5.Text) * -1)
                        XIva27 = Str$(Val(Iva27.Text) * -1)
                        XIva105 = Str$(Val(Iva105.Text) * -1)
                        XIb = Str$(Val(Ib.Text) * -1)
                        XExento = Str$(Val(Exento.Text) * -1)
                    Case Else
                        XImpre = "  "
                End Select
                XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Contado1.Value = True Then
                    XContado = "1"
                End If
                If Contado2.Value = True Then
                    XContado = "2"
                End If
                If Contado3.Value = True Then
                    XContado = "3"
                End If
                XEmpresa = "1"
                XNetolist = ""
                XExentolist = ""
                XParidad = Paridad.Text
                XPAgo = Str$(Pago.ListIndex)
                
                XParam = "'" + XNroInterno + "','" _
                        + XProveedor + "','" + XTipo + "','" _
                        + XLetra + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + Xvencimiento + "','" _
                        + XVencimiento1 + "','" + XPeriodo + "','" _
                        + XNeto + "','" _
                        + XIva21 + "','" _
                        + XIva5 + "','" + XIva27 + "','" _
                        + XIb + "','" + XExento + "','" _
                        + XContado + "','" _
                        + XImpre + "','" + XOrdFecha + "','" _
                        + XEmpresa + "','" + XNetolist + "','" _
                        + XExentolist + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                
                spIvaComp = "AltaIvaCompras " + XParam
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                WNroInterno = NroInterno.Text
                WProveedor = Proveedor.Text
                WTipo = Tipo.Text
                WLetra = Letra.Text
                WPunto = Punto.Text
                WNumero = Numero.Text
                WContado = XContado
                WFecha = Fecha.Text
                Wvencimiento = Vencimiento.Text
                WVencimiento1 = Vencimiento1.Text
                
                    Else
                    
                rstIvaComp.Close
                    
                Call Verifica_datos
                
                Rem modifica DE IVA COMPRAS
                
                XNroInterno = NroInterno.Text
                XProveedor = Proveedor.Text
                XTipo = Tipo.Text
                XLetra = Letra.Text
                XPunto = Punto.Text
                XNumero = Numero.Text
                XFecha = Fecha.Text
                Xvencimiento = Vencimiento.Text
                XVencimiento1 = Vencimiento1.Text
                XPeriodo = Periodo.Text
                XNeto = Neto.Text
                XIva21 = Iva21.Text
                XIva5 = Iva5.Text
                XIva27 = Iva27.Text
                XIva105 = Iva105.Text
                XIb = Ib.Text
                XExento = Exento.Text
                Select Case Val(Tipo.Text)
                    Case 1
                        XImpre = "FC"
                    Case 2
                        XImpre = "ND"
                    Case 3
                        XImpre = "NC"
                        XNeto = Str$(Val(Neto.Text) * -1)
                        XIva21 = Str$(Val(Iva21.Text) * -1)
                        XIva5 = Str$(Val(Iva5.Text) * -1)
                        XIva27 = Str$(Val(Iva27.Text) * -1)
                        XIva105 = Str$(Val(Iva105.Text) * -1)
                        XIb = Str$(Val(Ib.Text) * -1)
                        XExento = Str$(Val(Exento.Text) * -1)
                    Case Else
                        XImpre = "  "
                End Select
                XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Contado1.Value = True Then
                    XContado = "1"
                End If
                If Contado2.Value = True Then
                    XContado = "2"
                End If
                If Contado3.Value = True Then
                    XContado = "3"
                End If
                XEmpresa = "1"
                XNetolist = ""
                XExentolist = ""
                XParidad = Paridad.Text
                XPAgo = Str$(Pago.ListIndex)
                
                XParam = "'" + XNroInterno + "','" _
                        + XProveedor + "','" + XTipo + "','" _
                        + XLetra + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + Xvencimiento + "','" _
                        + XVencimiento1 + "','" + XPeriodo + "','" _
                        + XNeto + "','" _
                        + XIva21 + "','" _
                        + XIva5 + "','" + XIva27 + "','" _
                        + XIb + "','" + XExento + "','" _
                        + XContado + "','" _
                        + XImpre + "','" + XOrdFecha + "','" _
                        + XEmpresa + "','" + XNetolist + "','" _
                        + XExentolist + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                
                spIvaComp = "ActualizaIvaCompras " + XParam
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                WNroInterno = NroInterno.Text
                WProveedor = Proveedor.Text
                WTipo = Tipo.Text
                WLetra = Letra.Text
                WPunto = Punto.Text
                WNumero = Numero.Text
                WContado = XContado
                WFecha = Fecha.Text
                Wvencimiento = Vencimiento.Text
                WVencimiento1 = Vencimiento1.Text
                
            End If
            
            XParam = "'" + XNroInterno + "','" _
                         + Cai.Text + "','" _
                         + VtoCai.Text + "'"
                
            spIvaComp = "ActualizaIvaComprasCai " + XParam
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
            ZRechazado = "0"
            If Rechazado.Value = 1 Then
                ZRechazado = "1"
            End If
            
            ZSoloIva = "0"
            If SoloIva.Value = 1 Then
                ZSoloIva = "1"
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE IvaComp SET "
            ZSql = ZSql + " Iva105 = " + "'" + XIva105 + "',"
            ZSql = ZSql + " Despacho = " + "'" + Despacho.Text + "',"
            ZSql = ZSql + " Remito = " + "'" + Remito.Text + "',"
            ZSql = ZSql + " Rechazado = " + "'" + ZRechazado + "',"
            ZSql = ZSql + " SoloIva = " + "'" + ZSoloIva + "'"
            ZSql = ZSql + " Where NroInterno = " + "'" + XNroInterno + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
            m$ = "El numero interno asignado es " + NroInterno.Text
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
            
            
            
            
            
            
            
            
            Rem borra LAS IMPUTACIONES CONTABLES
        
            spImputac = "BorrarImputac " + "'" + NroInterno.Text + "'"
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
            
            
            Rem En caso de pyme nacion
            Rem borra los datos de la grabacion anterior
            
            If Val(NroInterno.Text) <> 0 Then
                
                ZSql = ""
                ZSql = ZSql + "DELETE CtaCtePrv"
                ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                
                Rem ZSql = ""
                Rem ZSql = ZSql + "DELETE Imputac"
                Rem ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                Rem spImputac = ZSql
                Rem Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE IvaComp"
                ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                spIvaComp = ZSql
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            
            
        
            Rem GRABA LAS IMPUTACIONES CONTABLES
        
            Renglon = 0
            Auxi1 = WNroInterno
            Call Ceros(Auxi1, 6)
            WNroInterno = Auxi1
            
            For iRow = 1 To 50
                
                WCuenta = WVector1.TextMatrix(iRow, 1)
                Debito = Val(WVector1.TextMatrix(iRow, 3))
                Credito = Val(WVector1.TextMatrix(iRow, 4))
                
                If WCuenta <> "" Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                        
                    XTipomovi = "2"
                    XNroInterno = WNroInterno
                    XProveedor = WProveedor
                    XTipocomp = WTipo
                    XLetracomp = WLetra
                    XPuntocomp = WPunto
                    XNrocomp = WNumero
                    XRenglon = Str$(Renglon + 1)
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    XRenglon = Auxi1$
                    XFecha = Fecha.Text
                    XObservaciones = ""
                    XCuenta = WCuenta
                    If Debito <> "" Then
                        XDebito = Str$(Debito)
                    End If
                    If Credito <> "" Then
                        XCredito = Str$(Credito)
                    End If
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XTitulo = "Compras"
                    XEmpresa = "1"
                    XClave = XTipomovi + XNroInterno + XRenglon
                    XDebitolist = ""
                    XCreditolist = ""
                        
                    XParam = "'" + XClave + "','" _
                                + XTipomovi + "','" + XProveedor + "','" _
                                + XTipocomp + "','" _
                                + XLetracomp + "','" + XPuntocomp + "','" _
                                + XNrocomp + "','" _
                                + XRenglon + "','" _
                                + XFecha + "','" + XObservaciones + "','" _
                                + XCuenta + "','" _
                                + XDebito + "','" _
                                + XCredito + "','" + XFechaOrd + "','" _
                                + XTitulo + "','" + XEmpresa + "','" _
                                + XDebitolist + "','" _
                                + XCreditolist + "','" _
                                + XNroInterno + "'"
                                
                    spImputac = "AltaImputacion " + XParam
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                                        
            Next iRow
        
        
        
        
        
        
        
        
        
        
            
            
            
            
            
            
            
            
            Rem graba la apertura de facturas
        
            
            ZSql = ""
            ZSql = ZSql + "DELETE IvaCompAdicional"
            ZSql = ZSql + " Where NroInterno = " + "'" + NroInterno.Text + "'"
            spIvaCompAdicional = ZSql
            Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
        
            Rem GRABA LAS IMPUTACIONES CONTABLES
        
            Renglon = 0
            Auxi1 = WNroInterno
            Call Ceros(Auxi1, 6)
            WNroInterno = Auxi1
            
            For iRow = 1 To 50
                
                WWCuit = WVector2.TextMatrix(iRow, 1)
                WWRazon = WVector2.TextMatrix(iRow, 2)
                WWTipo = WVector2.TextMatrix(iRow, 3)
                WWLetra = WVector2.TextMatrix(iRow, 4)
                WWPunto = WVector2.TextMatrix(iRow, 5)
                WWNumero = WVector2.TextMatrix(iRow, 6)
                WWFecha = WVector2.TextMatrix(iRow, 7)
                WWNeto = WVector2.TextMatrix(iRow, 8)
                WWIva21 = WVector2.TextMatrix(iRow, 9)
                WWIva27 = WVector2.TextMatrix(iRow, 10)
                WWIva105 = WVector2.TextMatrix(iRow, 11)
                WWPerceIva = WVector2.TextMatrix(iRow, 12)
                WWPerceIb = WVector2.TextMatrix(iRow, 13)
                WWExento = WVector2.TextMatrix(iRow, 14)
                
                If WWCuit <> "" Then
                
                    Auxi = WNroInterno
                    Call Ceros(Auxi, 8)
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    XClave = Auxi + Auxi1
                        
                    XNroInterno = WNroInterno
                    XRenglon = Str$(Renglon)
                    XCuit = WWCuit
                    XRazon = WWRazon
                    XTipo = WWTipo
                    XLetra = WWLetra
                    XPunto = WWPunto
                    XNumero = WWNumero
                    XFecha = WWFecha
                    XFechaOrd = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                    XNeto = WWNeto
                    XIva21 = WWIva21
                    XIva27 = WWIva27
                    XIva105 = WWIva105
                    XPerceIva = WWPerceIva
                    XPerceIb = WWPerceIb
                    XExento = WWExento
                    
                    
                    ZSql = "INSERT INTO IvaCompAdicional ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "NroInterno ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "Razon ,"
                    ZSql = ZSql + "Tipo ,"
                    ZSql = ZSql + "Letra ,"
                    ZSql = ZSql + "Punto ,"
                    ZSql = ZSql + "Numero ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "OrdFecha ,"
                    ZSql = ZSql + "Neto ,"
                    ZSql = ZSql + "Iva21 ,"
                    ZSql = ZSql + "Iva27 ,"
                    ZSql = ZSql + "Iva105 ,"
                    ZSql = ZSql + "PerceIva ,"
                    ZSql = ZSql + "PerceIb ,"
                    ZSql = ZSql + "Exento)"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + XClave + "',"
                    ZSql = ZSql + "'" + XNroInterno + "',"
                    ZSql = ZSql + "'" + XRenglon + "',"
                    ZSql = ZSql + "'" + XCuit + "',"
                    ZSql = ZSql + "'" + XRazon + "',"
                    ZSql = ZSql + "'" + XTipo + "',"
                    ZSql = ZSql + "'" + XLetra + "',"
                    ZSql = ZSql + "'" + XPunto + "',"
                    ZSql = ZSql + "'" + XNumero + "',"
                    ZSql = ZSql + "'" + XFecha + "',"
                    ZSql = ZSql + "'" + XOrdFecha + "',"
                    ZSql = ZSql + "'" + XNeto + "',"
                    ZSql = ZSql + "'" + XIva21 + "',"
                    ZSql = ZSql + "'" + XIva27 + "',"
                    ZSql = ZSql + "'" + XIva105 + "',"
                    ZSql = ZSql + "'" + XPerceIva + "',"
                    ZSql = ZSql + "'" + XPerceIb + "',"
                    ZSql = ZSql + "'" + XExento + "')"
                    spIvaCompAdicional = ZSql
                    Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                                        
            Next iRow
        
        
        
        
        
        
        
        
        
        
        
            Rem graba la cta.cte

            If Val(WContado) = 2 Or Val(WContado) = 3 Then
        
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount = 0 Then
                
                    XProveedor = WProveedor
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = WFecha
                    XEstado = "1"
                    Xvencimiento = Wvencimiento
                    XVencimiento1 = WVencimiento1
                    XNroInterno = WNroInterno
                    XTotal = Total.Caption
                    XSaldo = Total.Caption
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
                    XOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                    Select Case Val(WTipo)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XTotal = Str$(Val(Total.Caption) * -1)
                            XSaldo = Str$(Val(Total.Caption) * -1)
                        Case Else
                            XImpre = ""
                    End Select
                    XEmpresa = "1"
                    XSaldolist = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = Paridad.Text
                    XPAgo = Str$(Pago.ListIndex)
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spConsulta = "AltaCtaPrv " + XParam
                    Set rstConsulta = db.OpenRecordset(spConsulta + cParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    RstCtaPrv.Close
                  
                    XProveedor = WProveedor
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = WFecha
                    XEstado = "1"
                    Xvencimiento = Wvencimiento
                    XVencimiento1 = WVencimiento1
                    XNroInterno = WNroInterno
                    XTotal = Total.Caption
                    XSaldo = Total.Caption
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
                    XOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                    Select Case Val(WTipo)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XTotal = Str$(Val(Total.Caption) * -1)
                            XSaldo = Str$(Val(Total.Caption) * -1)
                        Case Else
                            XImpre = ""
                    End Select
                    XEmpresa = "1"
                    XSaldolist = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = Paridad.Text
                    XPAgo = Str$(Pago.ListIndex)
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spCtaprv = "ModificaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            End If
    
            XObservaciones = ""
            If Contado1.Value = True Then
                XContado = "1"
            End If
            If Contado2.Value = True Then
                XContado = "2"
            End If
            If Contado3.Value = True Then
                XContado = "3"
                XSaldo = "0"
            End If
            
            XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CtaCtePrv SET "
            ZSql = ZSql + " Observaciones = " + "'" + XObservaciones + "',"
            ZSql = ZSql + " Saldo = " + "'" + XSaldo + "',"
            ZSql = ZSql + " Tarjeta = " + "'" + XContado + "'"
            ZSql = ZSql + " Where Clave = " + "'" + XClave + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)

            Sql1 = "UPDATE Proveedor SET "
            Sql2 = " Cai = " + "'" + Cai.Text + "',"
            Sql3 = " VtoCai = " + "'" + VtoCai.Text + "'"
            Sql4 = " Where Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = Sql1 + Sql2 + Sql3 + Sql4
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
           
            
            
            If Val(WContado) = 3 Then
            
                ZZNroInternoAsociado = XNroInterno
                XProveedor = "10077777777"
                
                Auxi = "1"
                Call Ceros(Auxi, 4)
                ZZPunto = Auxi
                
                ZZProveedor = XProveedor
                ZZValorCuota = Val(Total.Caption) / Val(Cuotas.Text)
                Call Redondeo(ZZValorCuota)
                ZZImporte = Str$(ZZValorCuota)
                ZZLetra = "A"
                
                ZZMes = MesCuota.Text
                ZZAno = AnoCuota.Text
                
                For Ciclo = 1 To Val(Cuotas.Text)
                
                    Auxi = Ciclo
                    Call Ceros(Auxi, 2)
                    
                    Auxi1 = Numero.Text
                    Call Ceros(Auxi1, 8)
                    
                    ZZNumero = Right$(Auxi1, 6) + Auxi
                    
                
                    Auxi = ZZMes
                    Call Ceros(Auxi, 2)
                    Auxi1 = ZZAno
                    Call Ceros(Auxi1, 4)
                    ZZFecha = "01/" + Auxi + "/" + Auxi1
                    
                    
                    ZZNroInterno = "0"
                    If Val(Wempresa) = 1 Then
                        spIvaComp = "ListaIvacompNumero"
                        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                        If rstIvaComp.RecordCount > 0 Then
                            With rstIvaComp
                                .MoveLast
                                ZZNroInterno = Str$(rstIvaComp!NroInterno + 1)
                            End With
                            rstIvaComp.Close
                        End If
                            Else
                        ZHasta = "119000"
                        ZSql = ""
                        ZSql = ZSql + "Select IvaComp.NroInterno"
                        ZSql = ZSql + " FROM Ivacomp"
                        ZSql = ZSql + " Where Ivacomp.NroInterno <= " + "'" + ZHasta + "'"
                        ZSql = ZSql + " Order by Ivacomp.NroInterno"
                        spIvaComp = ZSql
                        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                        If rstIvaComp.RecordCount > 0 Then
                            With rstIvaComp
                                .MoveLast
                                ZZNroInterno = Str$(rstIvaComp!NroInterno + 1)
                            End With
                            rstIvaComp.Close
                        End If
                    End If
                    
                    
                    Rem ALTA DE IVA COMPRAS
                    
                    XNroInterno = ZZNroInterno
                    XProveedor = ZZProveedor
                    XTipo = Tipo.Text
                    XLetra = ZZLetra
                    XPunto = ZZPunto
                    XNumero = ZZNumero
                    XFecha = ZZFecha
                    Xvencimiento = ZZFecha
                    XVencimiento1 = ZZFecha
                    XPeriodo = ZZFecha
                    XNeto = ZZImporte
                    XIva21 = "0"
                    XIva5 = "0"
                    XIva27 = "0"
                    XIva105 = "0"
                    XIb = "0"
                    XExento = "0"
                    Select Case Val(Tipo.Text)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XNeto = Str$(Val(ZZImporte) * -1)
                        Case Else
                            XImpre = "  "
                    End Select
                    XOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    
                    XContado = "3"
                    
                    XEmpresa = "1"
                    XNetolist = ""
                    XExentolist = ""
                    XParidad = ""
                    XPAgo = "1"
                    
                    XParam = "'" + XNroInterno + "','" _
                            + XProveedor + "','" + XTipo + "','" _
                            + XLetra + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + Xvencimiento + "','" _
                            + XVencimiento1 + "','" + XPeriodo + "','" _
                            + XNeto + "','" _
                            + XIva21 + "','" _
                            + XIva5 + "','" + XIva27 + "','" _
                            + XIb + "','" + XExento + "','" _
                            + XContado + "','" _
                            + XImpre + "','" + XOrdFecha + "','" _
                            + XEmpresa + "','" + XNetolist + "','" _
                            + XExentolist + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spIvaComp = "AltaIvaCompras " + XParam
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE IvaComp SET "
                    ZSql = ZSql + " Iva105 = " + "'" + XIva105 + "',"
                    ZSql = ZSql + " NroInternoAsociado = " + "'" + ZZNroInternoAsociado + "'"
                    ZSql = ZSql + " Where NroInterno = " + "'" + XNroInterno + "'"
                    spIvaComp = ZSql
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    WNroInterno = ZZNroInterno
                    WProveedor = ZZProveedor
                    WTipo = Tipo.Text
                    WLetra = ZZLetra
                    WPunto = ZZPunto
                    WNumero = ZZNumero
                    WContado = XContado
                    WFecha = ZZFecha
                    Wvencimiento = ZZFecha
                    WVencimiento1 = ZZFecha
                    
                    
                
                    Rem graba la cta.cte
                    
                    XProveedor = ZZProveedor
                    XLetra = ZZLetra
                    XTipo = WTipo
                    XPunto = ZZPunto
                    XNumero = ZZNumero
                    XFecha = ZZFecha
                    XEstado = "1"
                    Xvencimiento = ZZFecha
                    XVencimiento1 = ZZFecha
                    XNroInterno = ZZNroInterno
                    XTotal = ZZImporte
                    XSaldo = ZZImporte
                    XClave = ZZProveedor + ZZLetra + WTipo + ZZPunto + ZZNumero
                    XOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    XOrdVencimiento = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                    Select Case Val(WTipo)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XTotal = Str$(Val(ZZImporte) * -1)
                            XSaldo = Str$(Val(ZZImporte) * -1)
                        Case Else
                            XImpre = ""
                    End Select
                    XEmpresa = "1"
                    XSaldolist = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = "1"
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spConsulta = "AltaCtaPrv " + XParam
                    Set rstConsulta = db.OpenRecordset(spConsulta + cParam, dbOpenSnapshot, dbSQLPassThrough)
            
                    XObservaciones = ""
                    XContado = "3"
                    ZZOrdFechaOriginal = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CtaCtePrv SET "
                    ZSql = ZSql + " Interes = " + "'" + "0" + "',"
                    ZSql = ZSql + " IvaInteres = " + "'" + "0" + "',"
                    ZSql = ZSql + " DesProveOriginal = " + "'" + Desproveedor.Caption + "',"
                    ZSql = ZSql + " FacturaOriginal = " + "'" + Numero.Text + "',"
                    ZSql = ZSql + " Cuota = " + "'" + Str$(Ciclo) + "',"
                    ZSql = ZSql + " ImporteOriginal = " + "'" + Total.Caption + "',"
                    ZSql = ZSql + " FechaOriginal = " + "'" + Fecha.Text + "',"
                    ZSql = ZSql + " OrdFechaOriginal = " + "'" + ZZOrdFechaOriginal + "',"
                    ZSql = ZSql + " NroInternoAsociado = " + "'" + ZZNroInternoAsociado + "',"
                    ZSql = ZSql + " Observaciones = " + "'" + XObservaciones + "',"
                    ZSql = ZSql + " Tarjeta = " + "'" + XContado + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + XClave + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZZMes = Str$(Val(ZZMes) + 1)
                    If Val(ZZMes) > 12 Then
                        ZZMes = "1"
                        ZZAno = Str$(Val(ZZAno) + 1)
                    End If
        
                Next Ciclo
                
            End If
            
            Call CmdLimpiar_Click
        
        End If
        
        NroInterno.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    WPasa = "S"

    If Contado3.Value = False Then
        Call Alinea_datos
        ClaveCtaprv = Proveedor.Text + Letra.Text + WTipo + WPunto + WNumero
        spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
        Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
        If RstCtaPrv.RecordCount > 0 Then
            If RstCtaPrv!Saldo <> RstCtaPrv!Total Then
                m$ = "El Comprobante se encuentra total o parcialmente cancelado"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
                WPasa = "N"
            End If
            RstCtaPrv.Close
        End If
    End If
    
    If WPasa = "S" Then

        spIvaComp = "ConsultaIvacomp " + "'" + NroInterno.Text + "'"
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
        
            rstIvaComp.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                spIvaComp = "BorrarIvacomp " + "'" + NroInterno.Text + "'"
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                If Val(NroInterno.Text) <> 0 Then
                    ZSql = ""
                    ZSql = ZSql + "DELETE CtaCtePrv"
                    ZSql = ZSql + " Where NroInterno = " + "'" + NroInterno.Text + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
                spImputac = "BorrarImputac " + "'" + NroInterno.Text + "'"
                Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
            
                Rem En caso de pyme nacion
                Rem borra los datos de la grabacion anterior
                
                If Val(NroInterno.Text) <> 0 Then
                    
                    ZSql = ""
                    ZSql = ZSql + "DELETE CtaCtePrv"
                    ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Rem ZSql = ""
                    Rem ZSql = ZSql + "DELETE Imputac"
                    Rem ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                    Rem spImputac = ZSql
                    Rem Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZSql = ""
                    ZSql = ZSql + "DELETE IvaComp"
                    ZSql = ZSql + " Where NroInternoAsociado = " + "'" + NroInterno.Text + "'"
                    spIvaComp = ZSql
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                    
                Call CmdLimpiar_Click
                    
            End If
        End If
        
    End If
    
    NroInterno.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII

    NroInterno.Text = ""
    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Vencimiento1.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Iva105.Text = ""
    Ib.Text = ""
    Exento.Text = ""
    Total.Caption = ""
    Paridad.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    Despacho.Text = ""
    Remito.Text = ""
    Contado1.Value = False
    Contado2.Value = True
    Contado3.Value = False
    Desproveedor.Caption = ""
    Remito.Text = ""
    Cuotas.Text = ""
    MesCuota.Text = ""
    AnoCuota.Text = ""
    
    TipoComp.ListIndex = 0
    Pago.ListIndex = 0
    Rechazado.Value = False
    SoloIva.Value = False
    
    NroInterno.Text = ""
    spIvaComp = "ListaIvacompNumero"
    Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstIvacomp.RecordCount > 0 Then
    Rem     With rstIvacomp
    Rem         .MoveLast
    Rem         NroInterno.Text = rstIvacomp!NroInterno + 1
    Rem     End With
    Rem     rstIvacomp.Close
    Rem End If
    
    NroInterno.SetFocus
End Sub

Private Sub cmdClose_Click()

    CmdLimpiar_Click

    NroInterno.SetFocus
    PrgCompras.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub


Private Sub Command1_Click()

    PantaDiscrimina.Height = 6000
    PantaDiscrimina.Left = 0
    PantaDiscrimina.Top = 240
    PantaDiscrimina.Width = 15000
    
    PantaDiscrimina.Visible = True


End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub



Private Sub Proveedor_KeyPress(KeyAscii As Integer)

    WProveedor = Proveedor.Text
    Proveedor.Text = WProveedor

    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Desproveedor.Caption = RstProveedor!Nombre
        If Trim(Cai.Text) = "" Then
            Cai.Text = IIf(IsNull(RstProveedor!Cai), "", RstProveedor!Cai)
            VtoCai.Text = IIf(IsNull(RstProveedor!VtoCai), "  /  /    ", RstProveedor!VtoCai)
        End If
        Letra.SetFocus
        RstProveedor.Close
            Else
        Proveedor.SetFocus
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Letra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Left$(Letra.Text, 1) = "A" Or Left$(Letra.Text, 1) = "B" Or Left$(Letra.Text, 1) = "C" Or Left$(Letra.Text, 1) = "X" Or Left$(Letra.Text, 1) = "M" Or Left$(Letra.Text, 1) = "I" Then
            Punto.SetFocus
                Else
            Letra.SetFocus
        End If
    End If
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
        Punto.Text = WPunto
        Numero.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WNumero = Numero.Text
        Call Ceros(WNumero, 8)
        Numero.Text = WNumero
        Cai.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoCai.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub VtoCai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(VtoCai.Text, Auxi)
        If Auxi = "S" Or VtoCai.Text = "  /  /    " Then
            WNumero = Numero.Text
            Call Ceros(WNumero, 8)
            Numero.Text = WNumero
            Fecha.SetFocus
                Else
            VtoCai.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
        
            If Periodo.Text = "  /  /    " Then
                Periodo.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            End If
            
            XEmpresa = Wempresa
                        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Claveven$ = Proveedor.Text
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                  
                Tipom = RstProveedor!TipoProv
                Dias = RstProveedor!Dias
                Dias = UCase(Dias)
                If Val(Dias) = 0 Then
                  Vencimiento.Text = Fecha.Text
                  Dias = 0
                End If
                Dias = Left$(Dias, 3)
                
                If Dias = "CON" Then
                    Vencimiento.Text = Fecha.Text
                     Else
                  If Tipom = 1 Then
                  Rem   Dias = "60"
                      Else
                       If Dias <> "15" Or Dias <> "30" Then
                    Rem     Dias = "30"
                         End If
                     End If
                End If
                
                
                If Dias <> "CON" Then
                    ZZDias = Trim(Str$(Val(Dias)))
                    Fecha2 = DateValue(Fecha.Text)
                    Vencimiento.Text = DateAdd("d", ZZDias, Fecha2)
                End If
                    
                RstProveedor.Close
                    
            End If
            
            Call Conecta_Empresa
            
            Rem fin by nan
            
            Vencimiento.SetFocus
            Vencimiento1.Text = Vencimiento.Text
            
           Rem ene end by nan
       Rem     Remito.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Remito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Rem Vencimiento.SetFocus
          Pago.SetFocus
          
        If Trim(Remito.Text) <> "" Then
            Call Remito_dblclick
        End If
    End If
End Sub

Private Sub Remito_dblclick()


    ZZPasaRemito = Remito.Text
    ZZPasaProveedor = Proveedor.Text
    ZZPasaProceso = 0
    
    Call Verifica_Pyme
    
    If ZPyme = "S" Then
        If Val(Cuotas.Text) = 0 Then
            Cuotas.Text = Trim(Str$(ZZCuotas))
        End If
        If Val(MesCuota.Text) = 0 Or Val(AnoCuota.Text) = 0 Then
            MesCuota.Text = Trim(Str$(ZZMesCuota))
            AnoCuota.Text = Trim(Str$(ZZAnoCuota))
        End If
        Contado3.Value = True
    End If
    
    PrgConsultaInforme.Show

End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Wvencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
            If Wvencimiento >= WFecha Then
                Vencimiento1.SetFocus
                    Else
                Vencimiento.SetFocus
            End If
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Vencimiento1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento1.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Wvencimiento = Right$(Vencimiento1.Text, 4) + Mid$(Vencimiento1.Text, 4, 2) + Left$(Vencimiento1.Text, 2)
            If Wvencimiento >= WFecha Then
                Periodo.SetFocus
                    Else
                Vencimiento1.SetFocus
            End If
                Else
            Vencimiento1.SetFocus
        End If
    End If
End Sub

Private Sub Periodo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Periodo.Text, Auxi)
        If Auxi = "S" Then
          Remito.SetFocus
          Rem  Pago.SetFocus
                Else
            Periodo.SetFocus
        End If
    End If
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Pago.ListIndex = 2 Then
            Paridad.SetFocus
                Else
            Paridad.Text = ""
            Neto.SetFocus
        End If
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = Pusing("#,###.####", Paridad.Text)
        Neto.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Neto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Neto.Text = Pusing("#,###,###.##", Neto.Text)
        If Val(Iva21.Text) = 0 Then
            If Letra.Text = "A" Or Letra.Text = "M" Then
                ZZIva = Val(Neto.Text) * 0.21
                Call Redondeo(ZZIva)
                Iva21.Text = Str$(ZZIva)
                Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
            End If
        End If
        Call Calcula_total
        Iva21.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
        Call Calcula_total
        Iva5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva5.Text = Pusing("#,###,###.##", Iva5.Text)
        Call Calcula_total
        Iva27.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva27_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva27.Text = Pusing("#,###,###.##", Iva27.Text)
        Call Calcula_total
        Iva105.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva105_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva105.Text = Pusing("#,###,###.##", Iva105.Text)
        Call Calcula_total
        Ib.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ib_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ib.Text = Pusing("#,###,###.##", Ib.Text)
        Call Calcula_total
        Exento.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Exento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Exento.Text = Pusing("#,###,###.##", Exento.Text)
        Call Calcula_total
        Despacho.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Despacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entra = "S"
        For iRow = 1 To 50
            ZZCuenta = Trim(WVector1.TextMatrix(iRow, 1))
            ZZDebito = WVector1.TextMatrix(iRow, 3)
            ZZCredito = WVector1.TextMatrix(iRow, 4)
            If Trim(ZZCuenta) <> "" Or Val(ZZDebito) <> 0 Or Val(ZZCredito) <> 0 Then
                Entra = "N"
            End If
        Next iRow
        If Entra = "S" Then
        
            ZZLugar = 0
            
            If Val(Total.Caption) <> 0 Then
                ZZLugar = ZZLugar + 1
                If Letra.Text = "I" Then
                    WVector1.TextMatrix(ZZLugar, 1) = "2010"
                        Else
                    WVector1.TextMatrix(ZZLugar, 1) = "2001"
                End If
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Total.Caption)
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Total.Caption)
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva21.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "151"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva21.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva21.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva27.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "151"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva27.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva27.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva105.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "151"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva105.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva105.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva5.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "152"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva5.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva5.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Ib.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "164"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Ib.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Ib.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
        End If
        
        Call Calcula_Click
        
        WVector1.Col = 1
        WVector1.Row = ZZLugar + 1
        Call StartEdit
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroInterno_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(NroInterno.Text) <> 0 Then
        
            XNroInterno = NroInterno.Text
            XProveedor = Proveedor.Text
            Rem XTipo = Tipo.Text
            XLetra = Letra.Text
            XPunto = Punto.Text
            XNumero = Numero.Text
            WNumero = Numero.Text
            Call Ceros(WNumero, 8)
            Numero.Text = WNumero
            
            spIvaComp = "ConsultaIvacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                    NroInterno.Text = XNroInterno
                    Proveedor.Text = XProveedor
                    Rem Tipo.Text = XTipo
                    Letra.Text = XLetra
                    Punto.Text = XPunto
                    Numero.Text = XNumero
                    '
                    rstIvaComp.Close
                    Call Imprime_Datos
                    '
                    Existe = "S"
                        Else
                    CmdLimpiar_Click
                    NroInterno.Text = XNroInterno
                    Proveedor.Text = XProveedor
                    Rem Tipo.Text = XTipo
                    Letra.Text = XLetra
                    Punto.Text = XPunto
                    Numero.Text = XNumero
                    Existe = "N"
                    Call Imprime_Descripcion
            End If
            
        End If
        
        Proveedor.SetFocus
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Cuentas Contables"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spProveedor = "ListaProveedoresOrdConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
        
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
Rem by nan
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "      " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
            
            With rstCuenta
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCuenta!Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCuenta.Close
            
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                    Proveedor.Text = RstProveedor!Proveedor
                    Cai.Text = IIf(IsNull(RstProveedor!Cai), "", RstProveedor!Cai)
                    Cai.Text = Trim(Cai.Text)
                    VtoCai.Text = IIf(IsNull(RstProveedor!VtoCai), "  /  /    ", RstProveedor!VtoCai)
                    RstProveedor.Close
                    Call Imprime_Descripcion
                        Else
                    CmdLimpiar_Click
                    Proveedor.Text = Claveven$
            End If
            Proveedor.SetFocus
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            WCuenta = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstCuenta!Cuenta
                WVector1.Col = 2
                WVector1.Text = rstCuenta!Descripcion
                WVector1.Col = 3
                Call StartEdit
                rstCuenta.Close
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_VectorII

    NroInterno.Text = ""
    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Vencimiento1.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Iva105.Text = ""
    Ib.Text = ""
    Exento.Text = ""
    Total.Caption = ""
    Contado1.Value = False
    Contado2.Value = True
    Contado3.Value = False
    Desproveedor.Caption = ""
    Paridad.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    Despacho.Text = ""
    Remito.Text = ""
    Cuotas.Text = ""
    MesCuota.Text = ""
    AnoCuota.Text = ""
    
    NroInterno.Text = ""
    Rem spIvacomp = "ListaIvacompNumero"
    Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstIvacomp.RecordCount > 0 Then
    Rem     With rstIvacomp
    Rem         .MoveLast
    Rem         NroInterno.Text = rstIvacomp!NroInterno + 1
    Rem     End With
    Rem     rstIvacomp.Close
    Rem End If
    
    TipoComp.Clear
    
    TipoComp.AddItem "Factura"
    TipoComp.AddItem "N.Debito"
    TipoComp.AddItem "N.Credito"
    
    TipoComp.ListIndex = 0
    
    TipoCompII.Clear
    
    TipoCompII.AddItem "Factura"
    TipoCompII.AddItem "N.Debito"
    TipoCompII.AddItem "N.Credito"
    
    TipoCompII.ListIndex = 0
    
    Pago.Clear
    
    Pago.AddItem ""
    Pago.AddItem "Pesos"
    Pago.AddItem "Clausula Dolar"
    
    Pago.ListIndex = 0
    

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
        Call Calcula_Click
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
        Case 4
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstCuenta!Descripcion
                WVector1.Col = 2
                rstCuenta.Close
                    Else
                WControl = "N"
            End If
            
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
    For iRow = 50 To 1 Step -1
        
        ZLegajo = WVector1.TextMatrix(iRow, 1)
            
        If ZLegajo <> "" Then
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
        For da = 0 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
    Call Calcula_Click
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem ""
    Opcion.AddItem ""

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
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
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 51
    
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Cuenta"
                WVector1.ColWidth(Ciclo) = 1700
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
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Debito"
                WVector1.ColWidth(Ciclo) = 1700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 4
                WVector1.Text = "Credito"
                WVector1.ColWidth(Ciclo) = 1700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
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

Private Sub Calcula_Click()

    SumaDebito = 0
    SumaCredito = 0
        
    For iRow = 1 To 50
        Debito = WVector1.TextMatrix(iRow, 3)
        SumaDebito = SumaDebito + Val(Debito)
        Credito = WVector1.TextMatrix(iRow, 4)
        SumaCredito = SumaCredito + Val(Credito)
    Next iRow
                    
    Call Redondeo(SumaDebito)
    Call Redondeo(SumaCredito)
    
    TotalDebito.Caption = Str$(SumaDebito)
    TotalCredito.Caption = Str$(SumaCredito)
    
    TotalDebito.Caption = Pusing("#,###,###.##", TotalDebito.Caption)
    TotalCredito.Caption = Pusing("#,###,###.##", TotalCredito.Caption)

End Sub

Private Sub CerrarBusquedaNro_Click()
    BusquedaNro.Visible = False
End Sub

Private Sub ConsultaII_Click()

    ProveedorII.Text = ""
    DesProveedorII.Caption = ""
    LetraII.Text = ""
    NumeroII.Text = ""
    PuntoII.Text = ""
    TipoII.Text = ""
    TipoCompII.ListIndex = 0
    
    BusquedaNro.Visible = True
    
    ProveedorII.SetFocus

End Sub


Private Sub ProveedorII_KeyPress(KeyAscii As Integer)

    WProveedor = ProveedorII.Text
    ProveedorII.Text = WProveedor

    spProveedor = "ConsultaProveedores " + "'" + ProveedorII.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        DesProveedorII.Caption = RstProveedor!Nombre
        LetraII.SetFocus
        RstProveedor.Close
            Else
        ProveedorII.SetFocus
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub LetraII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Left$(LetraII.Text, 1) = "A" Or Left$(LetraII.Text, 1) = "B" Or Left$(LetraII.Text, 1) = "C" Or Left$(LetraII.Text, 1) = "X" Or Left$(LetraII.Text, 1) = "M" Or Left$(LetraII.Text, 1) = "I" Then
            PuntoII.SetFocus
                Else
            LetraII.SetFocus
        End If
    End If
End Sub

Private Sub PuntoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = PuntoII.Text
        Call Ceros(WPunto, 4)
        PuntoII.Text = WPunto
        NumeroII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NumeroII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        TipoII.Text = TipoCompII.ListIndex + 1
        WTipo = TipoII.Text
        Call Ceros(WTipo, 2)
        
        WPunto = PuntoII.Text
        Call Ceros(WPunto, 4)
        
        WNumero = NumeroII.Text
        Call Ceros(WNumero, 8)
        
        ZSql = "Select *"
        ZSql = ZSql + " FROM Ivacomp"
        ZSql = ZSql + " Where Ivacomp.Proveedor = " + "'" + ProveedorII.Text + "'"
        ZSql = ZSql + " and Ivacomp.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Ivacomp.Letra = " + "'" + LetraII.Text + "'"
        ZSql = ZSql + " and Ivacomp.Punto = " + "'" + WPunto + "'"
        ZSql = ZSql + " and Ivacomp.Numero = " + "'" + WNumero + "'"
        spIvaComp = ZSql
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            NroInterno.Text = rstIvaComp!NroInterno
            rstIvaComp.Close
            BusquedaNro.Visible = False
            NroInterno_Keypress (13)
                Else
            m$ = "Factura no ingresada"
            A% = MsgBox(m$, 64, "Ingreso de Comprobantes")
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Verifica_Pyme()
    
    ZZCuotas = 0
    ZZMesCuota = 0
    ZZAnoCuota = 0
    
    Erase ZNroRemito
    Erase ZNroOrden
    ZLugarII = 0
    ZLugarIII = 0
    
    ZPyme = "N"
    
    ZZCargaRemito = Trim(ZZPasaRemito)
    
    If Val(ZZCargaRemito) = 0 Then
        Exit Sub
    End If
    
    Do
        MyPos = InStr(ZZCargaRemito, ",")
        If MyPos = 0 Then
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = ZZCargaRemito
            Exit Do
                Else
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = Mid$(ZZCargaRemito, 1, MyPos - 1)
            ZZCargaRemito = Mid$(ZZCargaRemito, MyPos + 1, 100)
        End If
    Loop
    
    Call Busca_Empresa
    XEmpresa = Wempresa
    
    Select Case Val(EmpresaTrabajo)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    For CicloII = 1 To ZLugarII
    
        ZZRemito = ZNroRemito(CicloII)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Informe.Remito = " + "'" + ZZRemito + "'"
        ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            With rstInforme
                .MoveFirst
                Do
                    If .EOF = False Then
                        
                        If rstInforme!Cantidad <> 0 Then
                            WLugarIII = WLugarIII + 1
                            ZNroOrden(WLugarIII, 1) = Str$(rstInforme!Orden)
                            ZNroOrden(WLugarIII, 2) = rstInforme!Articulo
                        End If
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstInforme.Close
        End If
        
    Next CicloII
    
    For Ciclo = 1 To WLugarIII
    
        ZZOrden = ZNroOrden(Ciclo, 1)
        ZZArticulo = ZNroOrden(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZOrden + "'"
        ZSql = ZSql + " and Orden.Articulo = " + "'" + ZZArticulo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            ZTarjeta = IIf(IsNull(rstOrden!Tarjeta), "0", rstOrden!Tarjeta)
            If ZTarjeta = 1 Then
                ZPyme = "S"
                
                ZZCuotas = IIf(IsNull(rstOrden!Cuotas), "", rstOrden!Cuotas)
                ZZMesCuota = IIf(IsNull(rstOrden!MesCuota), "", rstOrden!MesCuota)
                ZZAnoCuota = IIf(IsNull(rstOrden!AnoCuota), "", rstOrden!AnoCuota)
                
            End If
            rstOrden.Close
        End If
    
    Next Ciclo
        
    Call Conecta_Empresa
    
End Sub



Private Sub Busca_Empresa()

    EmpresaTrabajo = 0
    EmpresaAnterior = Wempresa
    XEmpresa = Wempresa
    
    If EmpresaAnterior = 1 Then

        For Va = 1 To 7
    
            Select Case Va
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
                Case 7
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = Wempresa
                rstInforme.Close
                Exit For
            End If
        
        Next Va
        
            Else
        
        For Va = 1 To 4
    
            Select Case Va
                Case 1
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = Wempresa
                rstInforme.Close
            End If
        
        Next Va
        
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub

Private Sub Contado3_Click()
    PantaPyme.Visible = True
    Cuotas.SetFocus
End Sub

Private Sub CierraPyme_Click()
    PantaPyme.Visible = False
    Pago.SetFocus
End Sub










Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 14
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1, 2
            
        Case 3
            WVector2.Text = UCase(WVector2.Text)
            If WVector2.Text = "FC" Or WVector2.Text = "ND" Or WVector2.Text = "NC" Then
                Rem ok
                    Else
                WControlII = "N"
            End If
            
        Case 4
            WVector2.Text = UCase(WVector2.Text)
            If WVector2.Text = "A" Or WVector2.Text = "B" Or WVector2.Text = "C" Or WVector2.Text = "A" Then
                Rem ok
                    Else
                WControlII = "N"
            End If
            
        Case 5, 6
            If Val(WVector2.Text) = 0 Then
                WControlII = "N"
            End If
                
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Ensayo = WVector2.TextMatrix(iRow, 1)
            
        If Ensayo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For da = 0 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorraII(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 15
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Cuit"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Razon Social"
                WVector2.ColWidth(Ciclo) = 2300
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 600
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 2
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Letra"
                WVector2.ColWidth(Ciclo) = 600
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 1
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 5
                WVector2.Text = "Punto"
                WVector2.ColWidth(Ciclo) = 600
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 6
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 850
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 8
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 7
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 10
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 2
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 8
                WVector2.Text = "Neto"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 9
                WVector2.Text = "Iva 21"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 10
                WVector2.Text = "Iva 27"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 11
                WVector2.Text = "Iva 10.5"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 12
                WVector2.Text = "Perc. Iva"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 13
                WVector2.Text = "Perc. IB"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 14
                WVector2.Text = "Exento"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub


