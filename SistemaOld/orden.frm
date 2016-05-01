VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Ordenes de Compra"
   ClientHeight    =   8325
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4800
      TabIndex        =   135
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Bultos 
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
      Left            =   10920
      MaxLength       =   6
      TabIndex        =   132
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame CargaAsunto 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ingrese el Asunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   129
      Top             =   2400
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox Asunto 
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
         Left            =   240
         TabIndex        =   130
         Top             =   360
         Width           =   10695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11520
      TabIndex        =   128
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Dada 
      Height          =   285
      Left            =   10080
      TabIndex        =   127
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   10800
      TabIndex        =   126
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame PantaPyme 
      Height          =   2295
      Left            =   2880
      TabIndex        =   119
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
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
         TabIndex        =   125
         Text            =   " "
         Top             =   960
         Width           =   735
      End
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
         TabIndex        =   124
         Text            =   " "
         Top             =   960
         Width           =   495
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
         TabIndex        =   121
         Text            =   " "
         Top             =   480
         Width           =   1095
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
         TabIndex        =   120
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label35 
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
         TabIndex        =   123
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Mes / Año"
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
         TabIndex        =   122
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.ComboBox Tarjeta 
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
      Left            =   8280
      TabIndex        =   117
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame CargaSolicitud 
      Height          =   2535
      Left            =   840
      TabIndex        =   109
      Top             =   2760
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Solicitud3 
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   113
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Solicitud2 
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   112
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Solicitud1 
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   110
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "Nro Solicitud"
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
         Left            =   960
         TabIndex        =   111
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame IngreDerechos 
      Height          =   1095
      Left            =   3360
      TabIndex        =   106
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WPorceDerechos 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   107
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Derechos"
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
         TabIndex        =   108
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame DatosImpo 
      Height          =   4935
      Left            =   7440
      TabIndex        =   89
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox DJai 
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   137
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox Flete 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   104
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ComboBox TipoPago 
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
         TabIndex        =   102
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox PedidoImpo 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   93
         Top             =   1320
         Width           =   2175
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   92
         Top             =   360
         Width           =   2175
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
         Left            =   1440
         TabIndex        =   91
         Top             =   840
         Width           =   1335
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
         Left            =   1440
         TabIndex        =   90
         Top             =   2280
         Width           =   2175
      End
      Begin MSMask.MaskEdBox FechaImpo 
         Height          =   285
         Left            =   1440
         TabIndex        =   94
         Top             =   1800
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
      Begin MSMask.MaskEdBox FechaDJai 
         Height          =   285
         Left            =   1440
         TabIndex        =   139
         Top             =   4320
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
      Begin VB.Label Label39 
         Caption         =   "Fecha DJai"
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
         TabIndex        =   140
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "DJai"
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
         TabIndex        =   138
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Flete U$S"
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
         TabIndex        =   105
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Tipo Pago"
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
         TabIndex        =   103
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label29 
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
         Left            =   240
         TabIndex        =   99
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label28 
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
         Left            =   240
         TabIndex        =   98
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label27 
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
         Left            =   240
         TabIndex        =   97
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Pedido"
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
         TabIndex        =   96
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label25 
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
         Left            =   240
         TabIndex        =   95
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton ImpreRed 
      Caption         =   "Impre Red"
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
      Left            =   4200
      TabIndex        =   88
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Primera 
      Caption         =   "Impre 1 Cop."
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
      Left            =   3360
      TabIndex        =   87
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox Carpeta 
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
      MaxLength       =   6
      TabIndex        =   86
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox TipoOrden 
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
      Left            =   9840
      TabIndex        =   84
      Top             =   120
      Width           =   1815
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
      Left            =   6720
      TabIndex        =   82
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame XCotart 
      Caption         =   "Cotizaciones por Articulo"
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
      TabIndex        =   56
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ComboBox Moneda2 
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
         Left            =   960
         TabIndex        =   78
         Top             =   840
         Width           =   1575
      End
      Begin MSMask.MaskEdBox XArt2 
         Height          =   255
         Left            =   960
         TabIndex        =   62
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin VB.CommandButton XAcepta2 
         Caption         =   "Acepta"
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
         Left            =   600
         TabIndex        =   59
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton XCancela2 
         Caption         =   "Cancela"
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
         Left            =   2880
         TabIndex        =   58
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton XConsulta2 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   57
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label20 
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
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label XDesArt2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   60
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame XCotPrv 
      Caption         =   "Cotizaciones por Proveedor"
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
      TabIndex        =   48
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ComboBox Moneda1 
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
         Left            =   1200
         TabIndex        =   76
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton XConsulta1 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   54
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton XCancela1 
         Caption         =   "Cancela"
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
         Left            =   3000
         TabIndex        =   53
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton XAcepta1 
         Caption         =   "Acepta"
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
         Left            =   720
         TabIndex        =   52
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox XProv1 
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
         TabIndex        =   50
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label17 
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
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   975
      End
      Begin VB.Label XDesProv1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   51
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame XCoti 
      Caption         =   "Ingreso de Cotizaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ComboBox Moneda3 
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
         Left            =   1920
         TabIndex        =   80
         Top             =   2880
         Width           =   2055
      End
      Begin MSMask.MaskEdBox XArti 
         Height          =   255
         Left            =   1920
         TabIndex        =   66
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
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
      Begin VB.CommandButton XCancela 
         Caption         =   "Cancela"
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
         Left            =   3960
         TabIndex        =   46
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton XAcepta 
         Caption         =   "Acepta"
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
         Left            =   1800
         TabIndex        =   45
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox XObservaciones 
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
         MaxLength       =   40
         TabIndex        =   44
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox XCondicion 
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
         MaxLength       =   40
         TabIndex        =   43
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox XPrec 
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
         MaxLength       =   10
         TabIndex        =   42
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox XProve 
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
         MaxLength       =   11
         TabIndex        =   41
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label21 
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
         Left            =   240
         TabIndex        =   79
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label XDesArti 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3960
         TabIndex        =   65
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label XDesProve 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3960
         TabIndex        =   47
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Condicion de Pago"
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
         TabIndex        =   39
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Precio"
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
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame XCc 
      Caption         =   "Cuenta Corriente de Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   68
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox XProv3 
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
         TabIndex        =   72
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton XAcepta3 
         Caption         =   "Acepta"
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
         Left            =   600
         TabIndex        =   71
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton XCancela3 
         Caption         =   "Cancela"
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
         Left            =   3000
         TabIndex        =   70
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton XConsulta3 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   69
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label XDesProv3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   73
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.CommandButton Ctacte 
      Caption         =   "Cta.Cte de Prov."
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
      Left            =   2280
      TabIndex        =   67
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Ingrecot 
      Caption         =   "Ingreso de Cotizaciones"
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
      Left            =   3360
      TabIndex        =   64
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Cotart 
      Caption         =   "Cot. x Art."
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
      Left            =   4200
      TabIndex        =   63
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Cotprv 
      Caption         =   "Cot. x Prv."
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
      Left            =   3360
      TabIndex        =   55
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Email2 
      Caption         =   "EMail Pantalla"
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
      Left            =   120
      TabIndex        =   34
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Tercera 
      Caption         =   "Impresion 3er Copia"
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
      Left            =   1200
      TabIndex        =   33
      Top             =   7080
      Width           =   975
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
      Left            =   5160
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   6255
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
      Left            =   1080
      MaxLength       =   11
      TabIndex        =   20
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "orden.rpt"
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   17
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
      TabIndex        =   18
      Top             =   6480
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
      Left            =   5520
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   14
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
   Begin VB.TextBox Orden 
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
      MaxLength       =   6
      TabIndex        =   0
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
      TabIndex        =   11
      Top             =   5880
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
      TabIndex        =   10
      Top             =   6480
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
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   11655
      Begin VB.TextBox WBultos 
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
         Left            =   10800
         MaxLength       =   10
         TabIndex        =   134
         Text            =   " "
         Top             =   600
         Width           =   735
      End
      Begin MSMask.MaskEdBox WFecha2 
         Height          =   300
         Left            =   8160
         TabIndex        =   22
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox WFecha1 
         Height          =   300
         Left            =   6960
         TabIndex        =   21
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.TextBox WCantidad 
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
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bultos"
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
         Left            =   10800
         TabIndex        =   133
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   9360
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult. Fecha"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1ra Fecha"
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
         Left            =   6960
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
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
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
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
         Left            =   4920
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
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   3495
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
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WPrecio 
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
         Height          =   300
         Left            =   6000
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label WCondicion 
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
         Left            =   9360
         TabIndex        =   23
         Top             =   600
         Width           =   1455
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
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   3495
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
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9480
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
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
      ItemData        =   "orden.frx":0000
      Left            =   5160
      List            =   "orden.frx":0007
      TabIndex        =   2
      Top             =   6240
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
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton LeePedido 
      Caption         =   "Lee     Pedidos"
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
      Height          =   495
      Left            =   10320
      TabIndex        =   100
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Complemento 
      Caption         =   "Datos Compleme."
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
      Left            =   120
      TabIndex        =   101
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton ComplementoOtro 
      Caption         =   "Datos Impresion"
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
      Left            =   2640
      TabIndex        =   114
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton OrdenImportacion 
      Caption         =   "Orden Importacion"
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
      Left            =   3840
      TabIndex        =   115
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton MiraSolicitud 
      Caption         =   "Solicitudes de Compra"
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
      Left            =   1320
      TabIndex        =   116
      Top             =   7680
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid WVector 
      Height          =   3735
      Left            =   0
      TabIndex        =   136
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      _Version        =   327680
      Rows            =   4000
      Cols            =   6
   End
   Begin VB.Label Label36 
      Caption         =   "Bultos"
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
      Left            =   10200
      TabIndex        =   131
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label34 
      Caption         =   "Forma Pago"
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
      Left            =   7080
      TabIndex        =   118
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label24 
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
      Left            =   5160
      TabIndex        =   85
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Tipo Orden"
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
      TabIndex        =   83
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label22 
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
      Left            =   5880
      TabIndex        =   81
      Top             =   120
      Width           =   1815
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
      Left            =   2640
      TabIndex        =   16
      Top             =   480
      Width           =   2415
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
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1575
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
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private XMoneda As Integer
Private WMonedaOrden As String
Private WTipoOrden As String
Private WTipoPago As String
Private Cantidad As Single
Private WOrdenprecio As String
Private XPrecio As String
Private XCantidad As String

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim rstNroCarpeta As Recordset
Dim spNroCarpeta As String

Dim XParam As String
Dim ZEmail As String
Dim Vector(100, 10) As String
Dim ZVector(100, 10) As String
Dim XPorceDerechos(100) As String
Dim XSolicitud(100, 3) As String
Dim ZZSolicitud(5000, 7) As String
Dim CargaEmpresa(12, 2) As String
Dim Empe(12, 10) As String

Dim ZZArtiCerti(100) As String
Dim ZEnsayo(100) As String
Dim ZDescriII(100) As String

Private TipoConsulta As String
Private XVector(10, 5) As String
Private Auxi As String
Private WAuxi As String
Private WSaldo As Double
Private Desdelugar As Integer
Dim Tabla(10000) As String
Private WEntre As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim Paridad As Double
Dim ParidadII As Double
Dim WDerechos As String
Dim XDerechos As Double
Dim ZDespacho As Double
Dim ZEntra As String
Dim ZSolicitud As String
Dim ZZArticulo As String

Dim ZZFechaLlegada As String
Dim ZZPagoDespacho As String
Dim ZZImpoDespacho As String
Dim ZZVtoDespacho As String
Dim ZZPagoLetra As String
Dim ZZImpoLetra As String
Dim ZZVtoLetra As String

Dim ZZCosto As Double
Dim ZParidad As Double
Dim ZParidadII As Double
Dim p As Object
Dim ZCoeParidad As Double


Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim MSubject As String
Dim MBody As String
Dim MAttach As String
Dim MAttachI As String
Dim MAttachII As String
Dim MAttachIII As String
Dim MAttachIV As String
Dim MAttachV As String
Dim AllPath As String

Dim ZZPasaTo As String
Dim ZZPasaCC As String
Dim ZZPasaBody As String
Dim ZZPasaFile As String


Private Sub Borra_Click()

    WVector.Col = 1
    WVector.Text = ""
    
    WVector.Col = 2
    WVector.Text = ""

    WVector.Col = 3
    WVector.Text = ""
    
    WVector.Col = 4
    WVector.Text = ""
    
    WVector.Col = 5
    WVector.Text = ""
    
    WVector.Col = 6
    WVector.Text = ""
    
    WVector.Col = 7
    WVector.Text = ""
    
    WVector.Col = 8
    WVector.Text = ""
    
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    WBultos.Text = ""
    WLinea.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgOrden.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    Pantalla.Clear
    WIndice.Clear
        
    Dim CargaEmpresa(100, 10) As String
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0003"
    CargaEmpresa(2, 2) = "Empresa03"
    CargaEmpresa(3, 1) = "0005"
    CargaEmpresa(3, 2) = "Empresa05"
    CargaEmpresa(4, 1) = "0006"
    CargaEmpresa(4, 2) = "Empresa06"
    CargaEmpresa(5, 1) = "0007"
    CargaEmpresa(5, 2) = "Empresa07"
    CargaEmpresa(6, 1) = "0010"
    CargaEmpresa(6, 2) = "Empresa10"
    CargaEmpresa(7, 1) = "0011"
    CargaEmpresa(7, 2) = "Empresa11"
    
    XEmpresa = Wempresa
    ZHasta = 7
    
    For ZCiclo = 1 To ZHasta
    
        Wempresa = CargaEmpresa(ZCiclo, 1)
        txtOdbc = CargaEmpresa(ZCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.NroDespacho LIKE " + "'" + "%" + Dada.Text + "%" + "'"
        ZSql = ZSql + " Order by Articulo, Orden"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                        WLiberada = rstLaudo!Liberada
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        If WLiberadaAnt > 0 Then
                            ZLiberada = WLiberadaAnt
                                Else
                            ZLiberada = WLiberada
                        End If
                        IngresaItem = Wempresa + " " + Str$(rstLaudo!Orden) + "   " + rstLaudo!Articulo + "  " + Str$(ZLiberada) + "  " + rstLaudo!NroDespacho
                        Pantalla.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLaudo.Close
        End If
    
    Next ZCiclo
    
    Call Conecta_Empresa
    
    Pantalla.Visible = True

End Sub

Private Sub Command2_Click()
    Call EMail_Click
End Sub

Private Sub Command3_Click()

    Dim ZZZVector(10000, 2) As String
    
    WHasta = "20141231"
    WDesde = "20140101"

    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " Lista = " + " '" + "" + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    ZZZLugar = 0
    Erase ZZZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Fechaord >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Fechaord <= " + "'" + WHasta + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZZLugar = ZZZLugar + 1
                    ZZZVector(ZZZLugar, 1) = rstInforme!Articulo
                    ZZZVector(ZZZLugar, 2) = Str$(rstInforme!Orden)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If


    For Ciclo = 1 To ZZZLugar
    
        ZArticulo = ZZZVector(Ciclo, 1)
        ZOrden = ZZZVector(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Orden SET "
        ZSql = ZSql + " Lista = " + " '" + "S" + "'"
        ZSql = ZSql + " Where Orden = " + "'" + ZOrden + "'"
        ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo



Stop



End Sub

Private Sub Complemento_Click()

    If TipoOrden.ListIndex <> 1 Then
        m$ = "La Orden no es de importacion"
        a% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
        Exit Sub
    End If

    WPasaOrden = Orden.Text
    WPasaCarpeta = Carpeta.Text
    WPasaOrigen = 1
    PrgOrdenComplemento.Show
    Rem Call PrgOrdenComplemento.Orden_Keypress(13)
    
End Sub

Private Sub ComplementoOtro_Click()

    If TipoOrden.ListIndex <> 1 Then
        m$ = "La Orden no es de importacion"
        a% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
        Exit Sub
    End If

    WPasaOrden = Orden.Text
    WPasaCarpeta = Carpeta.Text
    WPasaProveedor = Proveedor.Text
    WPasaOrigen = 1
    PrgOrdenComplementoImpo.Show

End Sub

Private Sub Consulta_Click()

    TipoConsulta = "0"

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"

     Opcion.Visible = True
     
 End Sub

Private Sub Cotart_Click()

    Moneda2.ListIndex = 0

    XCotart.Height = 1935
    XCotart.Left = 2280
    XCotart.Top = 1800
    XCotart.Width = 8655

    XCotart.Visible = True
    XArt2.Text = "  -   -   "
    XDesArt2.Caption = ""
    XArt2.SetFocus

End Sub

Private Sub Cotprv_Click()

    Moneda1.ListIndex = 0

    XCotPrv.Height = 2300
    XCotPrv.Left = 2280
    XCotPrv.Top = 1800
    XCotPrv.Width = 8655

    XCotPrv.Visible = True
    XProv1.Text = ""
    XDesProv1.Caption = ""
    XProv1.SetFocus

End Sub

Private Sub EMail_Click()
    
    Renglon = 0
        
    For a = 1 To 99
        
        Articulo = UCase(WVector.TextMatrix(a, 1))
                
        If Articulo <> "" Then
                    
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
                
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    Rem by nan
                    Rem WDescripcion = ""
                End If
                    
                XParam = "'" + Articulo + "','" _
                        + WDescripcion + "'"
                    
                spArticulo = "ModificaArticuloDescriComercial " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                  
        End If
            
    Next a
    
    MiRuta = CurDir + "\"
    MiRutaII = Left$(CurDir, 1)
    
    
    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Auxi = Orden.Text
    Call Ceros(Auxi, 6)
    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    
    Auxi1 = "c:\orden\" + Auxi + ".rtf"
    
    ZZEstado = Dir(Auxi1)
    If ZZEstado <> "" Then
        Kill Auxi1
    End If
    
    Rem Listado.Destination = crptMapi
    Rem Listado.PrintFileType = crptExcel50
    Listado.Destination = crptToFile
    Rem by nan *********88888
    Rem Listado.PrintFileType = crptWinWord
    Rem Listado.PrintFileType=
    Listado.PrintFileType = 15
    Listado.PrintFileName = Auxi1
    
    Rem Listado.EMailToList = ZEmail
    Rem Listado.EMailSubject = "ORDEN DE COMPRA NUMERO " + Orden.Text
    Rem Listado.EMailMessage = "Se remite por la presente la orden de compra " + Orden.Text
    
    Select Case Val(Wempresa)
        Case 1
            Listado.ReportFileName = "Orden1.rpt"
        Case 2
            Listado.ReportFileName = "Orden11.rpt"
        Case 3
            Listado.ReportFileName = "Orden2.rpt"
        Case 4
            Listado.ReportFileName = "Orden22.rpt"
        Case 5
            Listado.ReportFileName = "Orden3.rpt"
        Case 6
            Listado.ReportFileName = "Orden4.rpt"
        Case 7
            Listado.ReportFileName = "Orden7.rpt"
        Case 8
            Listado.ReportFileName = "Orden8.rpt"
        Case 9
            Listado.ReportFileName = "Orden9.rpt"
        Case 10
            Listado.ReportFileName = "Orden110.rpt"
        Case 11
            Listado.ReportFileName = "Orden111.rpt"
        Case Else
            Listado.ReportFileName = "Orden.rpt"
    End Select

    Orden.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, " + _
                            "Articulo.Descripcion, Proveedor.Nombre, Proveedor.CategoriaI " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + Orden.Text + " AND " + _
                            "Orden.Orden <= " + Orden.Text + " "
                            
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 0
    AAAAA = 1
    
    ChDrive MiRutaII
    ChDir MiRuta

    Auxi1 = "c:\Orden\" + Auxi
    sTo = LCase(ZEmail)
    sCC = ""
    sBCC = ""
    sSubject = "ORDEN DE COMPRA NUMERO " + Orden.Text
    sBody = "Se remite por la presente la orden de compra " + Orden.Text
            
  Rem  Shell "c:\orden\ConvertDoc /S " + Auxi1 + ".DOC /T " + Auxi1 + ".PDF /F9 /C12 /M2"
            
  Rem  sfile = Auxi1 + ".pdf"
   
   SFile = Auxi1 + ".rtf"
   
    

    
    Rem **************************
    
    Rem EmailAddress = sTo
    Rem CopiaAddress = sCC
    Rem MSubject = sSubject
    Rem MBody = sBody
    Rem MAttach = SFile
    Rem MAttachI = SFile
    Rem MAttachII = ""
    Rem MAttachIII = ""
    Rem MAttachIV = ""
    Rem MAttachV = ""

    Rem SendEmail
    ZZPasaTo = sTo
    ZZPasaCC = sCC
    ZZPasaBody = sSubject
    ZZPasaFile = SFile
    
    
    
    
    Asunto.Text = sSubject
    CargaAsunto.Visible = True
    Asunto.SetFocus

    Rem DADA
    Rem DADA
    Rem DADA
    
            
End Sub

Private Sub ImpreRed_Click()

    Renglon = 0
        
    For a = 1 To 99
        
        Articulo = UCase(WVector.TextMatrix(a, 1))
                
        If Articulo <> "" Then
                    
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
                
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    Rem WDescripcion = ""
                End If
                    
                XParam = "'" + Articulo + "','" _
                        + WDescripcion + "'"
                    
                spArticulo = "ModificaArticuloDescriComercial " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                  
        End If
            
    Next a

    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Select Case Val(Wempresa)
        Case 1
            Listado.ReportFileName = "OrdenImpre1.rpt"
        Case 2
            Listado.ReportFileName = "OrdenImpre11.rpt"
        Case 3
            Listado.ReportFileName = "OrdenImpre2.rpt"
        Case 4
            Listado.ReportFileName = "OrdenImpre22.rpt"
        Case 5
            Listado.ReportFileName = "OrdenImpre3.rpt"
        Case 6
            Listado.ReportFileName = "OrdenImpre4.rpt"
        Case 7
            Listado.ReportFileName = "OrdenImpre7.rpt"
        Case 8
            Listado.ReportFileName = "OrdenImpre8.rpt"
        Case 9
            Listado.ReportFileName = "OrdenImpre9.rpt"
        Case 10
            Listado.ReportFileName = "OrdenImpre110.rpt"
        Case 11
            Listado.ReportFileName = "OrdenImpre111.rpt"
        Case Else
            Listado.ReportFileName = "OrdenImpre.rpt"
    End Select

    Orden.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, " + _
                            "Articulo.Descripcion, Proveedor.Nombre, Proveedor.CategoriaI " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + Orden.Text + " AND " + _
                            "Orden.Orden <= " + Orden.Text + " "
                            
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 1
    AAAAA = 1

End Sub


Private Sub Impresion_Crystal()

    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = Wempresa
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Impretit = !Nombre
                Else
            Impretit = ""
        End If
    End With

    ZZImpreEmpresa = Impretit
    ZZImpreLeyenda = Leyenda.Text
    ZZImpreTipoOrden = ""
    If TipoOrden.ListIndex = 1 Then
        ZZImpreTipoOrden = "(IMPORTACION)"
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " ImpreEmpresa = " + "'" + ZZImpreEmpresa + "',"
    ZSql = ZSql + " ImpreLeyenda = " + "'" + ZZImpreLeyenda + "',"
    ZSql = ZSql + " ImpreTipoOrden = " + "'" + ZZImpreTipoOrden + "'"
    ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)


    Renglon = 0
        
    For a = 1 To 99
        
        Articulo = UCase(WVector.TextMatrix(a, 1))
                
        If Articulo <> "" Then
                        
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
                
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    Rem WDescripcion = ""
                End If
                    
                XParam = "'" + Articulo + "','" _
                        + WDescripcion + "'"
                    
                spArticulo = "ModificaArticuloDescriComercial " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                  
        End If
            
    Next a
    
    
    

    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "ImpreOrdenCrystal.rpt"

    Orden.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, Orden.ImpreEmpresa, Orden.ImpreLeyenda, Orden.ImpreTipoOrden, " + _
                            "Articulo.Descripcion, Proveedor.Nombre, Proveedor.CategoriaI " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + Orden.Text + " AND " + _
                            "Orden.Orden <= " + Orden.Text + " "
                            
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 1
    AAAAA = 1

End Sub



Private Sub Impresion_Impo()
    
    Lee = "S"
    
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrdenImpo"
    Sql3 = " Where ObservaOrdenImpo.Orden = " + "'" + Orden.Text + "'"
    spObservaOrdenImpo = Sql1 + Sql2 + Sql3
    Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrdenImpo.RecordCount > 0 Then
        Lee = "N"
        rstObservaOrdenImpo.Close
    End If
    
    If Lee = "S" Then
        
        XEmpresa = Wempresa
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Sql1 = "Select *"
        Sql2 = " FROM ProveedorAdicional"
        Sql3 = " Where ProveedorAdicional.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedorAdicional = Sql1 + Sql2 + Sql3 + Sql4
        Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedorAdicional.RecordCount > 0 Then
            
            ZZDescri11 = rstProveedorAdicional!Descri11
            ZZDescri12 = rstProveedorAdicional!Descri12
            ZZDescri13 = rstProveedorAdicional!Descri13
            ZZDescri14 = rstProveedorAdicional!Descri14
            
            ZZDescri21 = rstProveedorAdicional!Descri21
            ZZDescri22 = rstProveedorAdicional!Descri22
            ZZDescri23 = rstProveedorAdicional!Descri23
            
            ZZDescri31 = rstProveedorAdicional!Descri31
            ZZDescri32 = rstProveedorAdicional!Descri32
            ZZDescri33 = rstProveedorAdicional!Descri33
            
            ZZDescri40 = rstProveedorAdicional!Descri40
            ZZDescri41 = rstProveedorAdicional!Descri41
            ZZDescri42 = rstProveedorAdicional!Descri42
            ZZDescri43 = rstProveedorAdicional!Descri43
            ZZDescri44 = rstProveedorAdicional!Descri44
            ZZDescri45 = rstProveedorAdicional!Descri45
            ZZDescri46 = rstProveedorAdicional!Descri46
            ZZDescri47 = rstProveedorAdicional!Descri47
            ZZDescri48 = rstProveedorAdicional!Descri48
            ZZDescri49 = rstProveedorAdicional!Descri49
            
            ZZDescri51 = rstProveedorAdicional!Descri51
            ZZDescri52 = rstProveedorAdicional!Descri52
            ZZDescri53 = rstProveedorAdicional!Descri53
            ZZDescri54 = rstProveedorAdicional!Descri54
            ZZDescri55 = rstProveedorAdicional!Descri55
            ZZDescri56 = rstProveedorAdicional!Descri56
            ZZDescri57 = rstProveedorAdicional!Descri57
        
            rstProveedorAdicional.Close
        End If
    
        Call Conecta_Empresa
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ObservaOrdenImpo ("
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Descri11 ,"
        ZSql = ZSql + "Descri12 ,"
        ZSql = ZSql + "Descri13 ,"
        ZSql = ZSql + "Descri14 ,"
        ZSql = ZSql + "Descri21 ,"
        ZSql = ZSql + "Descri22 ,"
        ZSql = ZSql + "Descri23 ,"
        ZSql = ZSql + "Descri31 ,"
        ZSql = ZSql + "Descri32 ,"
        ZSql = ZSql + "Descri33 ,"
        ZSql = ZSql + "Descri40 ,"
        ZSql = ZSql + "Descri41 ,"
        ZSql = ZSql + "Descri42 ,"
        ZSql = ZSql + "Descri43 ,"
        ZSql = ZSql + "Descri44 ,"
        ZSql = ZSql + "Descri45 ,"
        ZSql = ZSql + "Descri46 ,"
        ZSql = ZSql + "Descri47 ,"
        ZSql = ZSql + "Descri48 ,"
        ZSql = ZSql + "Descri49 ,"
        ZSql = ZSql + "Descri51 ,"
        ZSql = ZSql + "Descri52 ,"
        ZSql = ZSql + "Descri53 ,"
        ZSql = ZSql + "Descri54 ,"
        ZSql = ZSql + "Descri55 ,"
        ZSql = ZSql + "Descri56 ,"
        ZSql = ZSql + "Descri57 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Orden.Text + "',"
        ZSql = ZSql + "'" + ZZDescri11 + "',"
        ZSql = ZSql + "'" + ZZDescri12 + "',"
        ZSql = ZSql + "'" + ZZDescri13 + "',"
        ZSql = ZSql + "'" + ZZDescri14 + "',"
        ZSql = ZSql + "'" + ZZDescri21 + "',"
        ZSql = ZSql + "'" + ZZDescri22 + "',"
        ZSql = ZSql + "'" + ZZDescri23 + "',"
        ZSql = ZSql + "'" + ZZDescri31 + "',"
        ZSql = ZSql + "'" + ZZDescri32 + "',"
        ZSql = ZSql + "'" + ZZDescri33 + "',"
        ZSql = ZSql + "'" + ZZDescri40 + "',"
        ZSql = ZSql + "'" + ZZDescri41 + "',"
        ZSql = ZSql + "'" + ZZDescri42 + "',"
        ZSql = ZSql + "'" + ZZDescri43 + "',"
        ZSql = ZSql + "'" + ZZDescri44 + "',"
        ZSql = ZSql + "'" + ZZDescri45 + "',"
        ZSql = ZSql + "'" + ZZDescri46 + "',"
        ZSql = ZSql + "'" + ZZDescri47 + "',"
        ZSql = ZSql + "'" + ZZDescri48 + "',"
        ZSql = ZSql + "'" + ZZDescri49 + "',"
        ZSql = ZSql + "'" + ZZDescri51 + "',"
        ZSql = ZSql + "'" + ZZDescri52 + "',"
        ZSql = ZSql + "'" + ZZDescri53 + "',"
        ZSql = ZSql + "'" + ZZDescri54 + "',"
        ZSql = ZSql + "'" + ZZDescri55 + "',"
        ZSql = ZSql + "'" + ZZDescri56 + "',"
        ZSql = ZSql + "'" + ZZDescri57 + "')"
        spObservaOrdenImpo = ZSql
        Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
        
    End If

    Listado.WindowTitle = "Impresion de Orden de Compra de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 1
    Listado.Destination = 0
    
    Select Case Val(Wempresa)
        Case 2, 4, 8, 9
           Listado.ReportFileName = "ImpreOrdenImportacionPelli.rpt"
        Rem    Listado.ReportFileName = "nan22.rpt"
     
        Case Else
        Rem by nan
            Listado.ReportFileName = "ImpreOrdenImportacion.rpt"
    End Select
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Orden.Orden, Orden.Renglon, Orden.Fecha, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Moneda, Orden.Carpeta, Orden.ImpreMarca, Orden.Planta, " _
            + "Proveedor.Nombre, Proveedor.Direccion, Proveedor.Localidad, " _
            + "ObservaOrdenImpo.Descri11, ObservaOrdenImpo.Descri12, ObservaOrdenImpo.Descri13, ObservaOrdenImpo.Descri14, ObservaOrdenImpo.Descri21, ObservaOrdenImpo.Descri22, ObservaOrdenImpo.Descri23, ObservaOrdenImpo.Descri31, ObservaOrdenImpo.Descri32, ObservaOrdenImpo.Descri33, ObservaOrdenImpo.Descri40, ObservaOrdenImpo.Descri41, ObservaOrdenImpo.Descri42, ObservaOrdenImpo.Descri43, ObservaOrdenImpo.Descri44, ObservaOrdenImpo.Descri45, ObservaOrdenImpo.Descri46, ObservaOrdenImpo.Descri47, ObservaOrdenImpo.Descri48, ObservaOrdenImpo.Descri49, ObservaOrdenImpo.Descri51, ObservaOrdenImpo.Descri52, ObservaOrdenImpo.Descri53, ObservaOrdenImpo.Descri54, ObservaOrdenImpo.Descri55, ObservaOrdenImpo.Descri56, ObservaOrdenImpo.Descri57 " _
            + "From " _
            + DSQ + ".dbo.Orden Orden, " _
            + DSQ + ".dbo.Proveedor Proveedor, " _
            + DSQ + ".dbo.ObservaOrdenImpo ObservaOrdenImpo " _
            + "Where " _
            + "Orden.Proveedor = Proveedor.Proveedor And " _
            + "Orden.Orden = ObservaOrdenImpo.Orden AND " _
            + "Orden.Orden >= " + Orden.Text + " AND " _
            + "Orden.Orden <= " + Orden.Text

    Listado.Connect = Connect()
    Listado.Action = 1

End Sub

Private Sub CtaCte_Click()

    XCc.Height = 1935
    XCc.Left = 2280
    XCc.Top = 1800
    XCc.Width = 8655

    XCc.Visible = True
    XProv3.Text = ""
    XDesProv3.Caption = ""
    XProv3.SetFocus

End Sub

Private Sub Email2_Click()

    Renglon = 0
    
    For a = 1 To 99
        
        Articulo = UCase(WVector.TextMatrix(a, 1))
                
        If Articulo <> "" Then
        
                    
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
                
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    Rem WDescripcion = ""
                End If
                    
                XParam = "'" + Articulo + "','" _
                        + WDescripcion + "'"
                    
                spArticulo = "ModificaArticuloDescriComercial " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                  
        End If
            
    Next a

    MiRuta = CurDir + "\"
    MiRutaII = Left$(CurDir, 1)
    
    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 0
    
    Listado.EMailToList = ZEmail
    Listado.EMailSubject = "ORDEN DE COMPRA NUMERO" + Orden.Text
    Listado.EMailMessage = "Se remite por la presente la orden de compra " + Orden.Text
    
    Select Case Val(Wempresa)
        Case 1
            Listado.ReportFileName = "Orden1.rpt"
        Case 2
            Listado.ReportFileName = "Orden11.rpt"
        Case 3
            Listado.ReportFileName = "Orden2.rpt"
        Case 4
            Listado.ReportFileName = "Orden22.rpt"
        Case 5
            Listado.ReportFileName = "Orden3.rpt"
        Case 6
            Listado.ReportFileName = "Orden4.rpt"
        Case 7
            Listado.ReportFileName = "Orden7.rpt"
        Case 8
            Listado.ReportFileName = "Orden8.rpt"
        Case 9
            Listado.ReportFileName = "Orden9.rpt"
        Case 10
            Listado.ReportFileName = "Orden110.rpt"
        Case 11
            Listado.ReportFileName = "Orden111.rpt"
        Case Else
            Listado.ReportFileName = "Orden.rpt"
    End Select
    
    Orden.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, " + _
                            "Articulo.Descripcion, Proveedor.Nombre, Proveedor.CategoriaI " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + Orden.Text + " AND " + _
                            "Orden.Orden <= " + Orden.Text + " "
                            
    
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 1
    
    ChDrive MiRutaII
    ChDir MiRuta
    
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgOrden.Caption = "Ingreso de Orden de Compras :  " + !Nombre
        End If
    End With


    OPEN_FILE_Empresa
    OPEN_FILE_Liscot
    OPEN_FILE_ImpCtaCtePrv
    
    Select Case WProcesoOrden
        Case 1
        
            Call Limpia_Vector
    
            Renglon = 0
            Erase Vector
            
            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            For Ciclo = 1 To 100
                If WVectorOrden(Ciclo, 1) <> "" Then
                
                    ZArticulo = WVectorOrden(Ciclo, 1)
                    WPrecio = 0
                    WCondicion = ""
            
                    spCotiza = "ListaCotizaProveedor " + "'" + WProveedorOrden + "'"
                    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCotiza.RecordCount > 0 Then
                        With rstCotiza
                            .MoveFirst
                            Do
                                If .EOF = False Then
                
                                    If ZArticulo = rstCotiza!Articulo Then
                                        If rstCotiza!FechaOrd > WFecha Then
                                            WPrecio = rstCotiza!Precio
                                            WCondicion = rstCotiza!Condicion
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
                    
                    WVectorOrden(Ciclo, 4) = Str$(WPrecio)
                    WVectorOrden(Ciclo, 5) = WCondicion
                    
                End If
            Next Ciclo
            
            Call Conecta_Empresa
            
            For Ciclo = 1 To 99
                If WVectorOrden(Ciclo, 1) <> "" Then
            
                    Renglon = Renglon + 1
                    
                    WVector.Row = Renglon
                
                    WVector.Col = 1
                    WVector.Text = WVectorOrden(Ciclo, 1)
                    Auxi1 = WVectorOrden(Ciclo, 1)
                    
                    spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WVector.Col = 2
                        WVector.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                
                    WVector.Col = 3
                    WVector.Text = Pusing("###,###.##", WVectorOrden(Ciclo, 2))
                
                    WVector.Col = 4
                    WVector.Text = Pusing("###,###.###", WVectorOrden(Ciclo, 4))
                
                    WVector.Col = 5
                    WVector.Text = Fecha.Text
                
                    WVector.Col = 6
                    WVector.Text = WVectorOrden(Ciclo, 3)
                
                    WVector.Col = 7
                    WVector.Text = WVectorOrden(Ciclo, 5)
                
                    WVector.Col = 8
                    WVector.Text = ""
                
                    Vector(Renglon, 1) = Auxi1
                    
                End If
            
            Next Ciclo
    
        Case Else
    End Select
    WProcesoOrden = 0
    WVector.Col = 1
    WVector.Row = 1
    
End Sub


Private Sub Ingrecot_Click()

    Moneda3.ListIndex = 0
    Desdelugar = 0

    XCoti.Height = 4200
    XCoti.Left = 2360
    XCoti.Top = 1320
    XCoti.Width = 7455
    
    XCoti.Visible = True
    
    XProve.Text = ""
    XArti.Text = "  -   -   "
    XPrec.Text = ""
    XCondicion.Text = ""
    XObservaciones.Text = ""
    
    XProve.SetFocus

End Sub

Private Sub LeePedido_Click()
    WProveedorOrden = Proveedor.Text
    WDesProveedorOrden = DesProveedor.Caption
    PrgOrdenII.Show
End Sub

Private Sub MiraSolicitud_Click()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Solic"
    Opcion.AddItem "Representantes"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 2
    
    Call Opcion_Click

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
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spProveedor = "ListaProveedoresOrdConsulta"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = rstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + rstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstProveedor.Close
            
            Call Conecta_Empresa
            Pantalla.Visible = True
            Ayuda.SetFocus
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            spArticulo = "ListaArticuloConsulta"
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
            Ayuda.SetFocus
                      
        Case 2
            Erase ZZSolicitud
            XEmpresa = Wempresa
            
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    CargaEmpresa(6, 1) = "0010"
                    CargaEmpresa(6, 2) = "Empresa10"
                    CargaEmpresa(7, 1) = "0011"
                    CargaEmpresa(7, 2) = "Empresa11"
                    ZHasta = 7
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
            End Select
            
            Erase ZZSolicitud
            WLugar = 0
            For ZCiclo = 1 To ZHasta
            
                Wempresa = CargaEmpresa(ZCiclo, 1)
                txtOdbc = CargaEmpresa(ZCiclo, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                XParam = "'" + "X" + " '"
                spSolic = "ListaSolicitudPendiente " + XParam
                Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                If rstSolic.RecordCount > 0 Then
                    With rstSolic
            
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                            
                                WLugar = WLugar + 1
                                ZZSolicitud(WLugar, 1) = rstSolic!Articulo
                                ZZSolicitud(WLugar, 2) = Wempresa
                                ZZSolicitud(WLugar, 3) = rstSolic!Clave
                                ZZSolicitud(WLugar, 4) = rstSolic!Observaciones
                                ZZSolicitud(WLugar, 5) = rstSolic!Obser
                                ZZSolicitud(WLugar, 6) = Str$(rstSolic!Cantidad)
                                ZZSolicitud(WLugar, 7) = rstSolic!Solicitante
                                .MoveNext
                        
                                If .EOF = True Then
                                    Exit Do
                                End If
                        
                            Loop
                        End If
                
                    End With
                    rstSolic.Close
                End If
                
            Next ZCiclo
            
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            For Ciclo = 1 To WLugar
                
                ZZArticulo = ZZSolicitud(Ciclo, 1)
                ZZEntra = "S"
                
                If Proveedor.Text = "10071011210" Then
                    If Left$(ZZArticulo, 1) = "T" Or Left$(ZZArticulo, 1) = "Z" Then
                        ZZEntra = "N"
                    End If
                End If
                
                If ZZEntra = "S" Then
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cotiza"
                    ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and Articulo = " + "'" + ZZArticulo + "'"
                    spCotiza = ZSql
                    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCotiza.RecordCount > 0 Then
                        rstCotiza.Close
                            Else
                        ZZSolicitud(Ciclo, 1) = ""
                        ZZSolicitud(Ciclo, 2) = ""
                        ZZSolicitud(Ciclo, 3) = ""
                        ZZSolicitud(Ciclo, 4) = ""
                        ZZSolicitud(Ciclo, 5) = ""
                        ZZSolicitud(Ciclo, 6) = ""
                        ZZSolicitud(Ciclo, 7) = ""
                    End If
                    
                End If
            Next Ciclo
            
            Call Conecta_Empresa
            
            For Ciclo = 1 To WLugar
            
                ZZArticulo = ZZSolicitud(Ciclo, 1)
                If ZZArticulo <> "" Then
            
                    ZZEmpresa = ZZSolicitud(Ciclo, 2)
                    ZZClave = ZZSolicitud(Ciclo, 3)
                    ZZObservaciones = ZZSolicitud(Ciclo, 4)
                    ZZObser = ZZSolicitud(Ciclo, 5)
                    ZZCantidad = ZZSolicitud(Ciclo, 6)
                    ZZSolicitante = ZZSolicitud(Ciclo, 7)
                    ZZDescripcion = ""
        
                    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    
                    Auxi$ = ZZCantidad
                    Call Mascara("###,###.##", Auxi$)
                    ZZCantidad = Auxi$
        
                    IngresaItem = ZZArticulo + "  " + ZZCantidad + "  " + Trim(Left$(ZZDescripcion, 15)) + "  " + Trim(Left$(ZZSolicitante, 15)) + "  " + Trim(ZZObservaciones) + "  " + Trim(ZZObser)
                    Pantalla.AddItem IngresaItem
                    IngresaItem = ZZClave + ZZEmpresa
                    WIndice.AddItem IngresaItem
                    
                End If
                
            Next Ciclo
            Pantalla.Visible = True
        
        Case Else
    End Select

End Sub

Private Sub WVector_Click()

    WVector.Col = 1
    If Len(WVector.Text) = 10 Then
        WLinea.Text = WVector.Row
        WArticulo.Text = WVector.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    WVector.Col = 2
    WDescripcion.Caption = WVector.Text

    WVector.Col = 3
    If Val(WVector.Text) <> 0 Then
        WCantidad.Text = WVector.Text
            Else
        WCantidad.Text = ""
    End If
    
    WVector.Col = 4
    WPrecio.Caption = WVector.Text
    
    WVector.Col = 5
    If WVector.Text <> "" Then
        WFecha1.Text = WVector.Text
    End If
    
    WVector.Col = 6
    If WVector.Text <> "" Then
        WFecha2.Text = WVector.Text
    End If
    
    WVector.Col = 7
    WCondicion.Caption = WVector.Text
    
    WVector.Col = 8
    WBultos.Text = WVector.Text
    
    WLugar = WVector.Row
    
    WPorceDerechos.Text = XPorceDerechos(WLugar)
    Solicitud1.Text = XSolicitud(WLugar, 1)
    Solicitud2.Text = XSolicitud(WLugar, 2)
    Solicitud3.Text = XSolicitud(WLugar, 3)
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()
    
    Pantalla.Visible = False
    
    On Error GoTo WError
    
    If Val(Wempresa) = 5 And TipoOrden.ListIndex = 3 Then
        For ZZCicla = 1 To 99
            Articulo = UCase(WVector.TextMatrix(ZZCicla, 1))
            If Articulo <> "" Then
                Select Case UCase(WVector.TextMatrix(ZZCicla, 1))
                    Case "ZE-010-100", "ZE-010-101", "ZE-042-100", "ZE-043-100", "ZE-045-100", "ZE-071-100", "ZE-077-100", "ZE-080-100", "ZE-081-100"
                        m$ = "Los envases ZE-010-100, ZE-010-101, ZE-042-100, ZE-043-100, ZE-045-100, ZE-071-100" + _
                             Chr$(13) + _
                             "ZE-077-100, ZE-080-100, ZE-081-100 en planta III se deben realizar como ordenes de compra normal y no de envases"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        Exit Sub
                    Case Else
                End Select
            End If
        Next ZZCicla
    End If
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha de la orden de compra es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
    If TipoOrden.ListIndex = 1 Then
        If TipoImpo.ListIndex = 0 Then
            m$ = "Se debe informar la via de transporte"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
    End If
    
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 0 Then
            m$ = "Se debe informar la condicion de la importacion"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
    End If
            
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 1 Then
            If Val(Flete.Text) = 0 Then
                m$ = "Se debe informar el monto del flete"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                Exit Sub
            End If
        End If
    End If
    
    If Moneda.ListIndex = 2 Then
        If TipoOrden.ListIndex <> 1 Then
            m$ = "Se puede cargar en Euros solo las O/C de importaciones"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
    End If
    
    If Tarjeta.ListIndex = 1 Then
        If Val(Cuotas.Text) = 0 Then
            m$ = "Cantidad de cuotas incorrecto"
            AAa% = MsgBox(m$, 0, "Datos pagos Pyme Nacion")
            Exit Sub
        End If
        If Val(MesCuota.Text) < 1 Or Val(MesCuota.Text) > 12 Then
            m$ = "Mes Incorrecto"
            AAa% = MsgBox(m$, 0, "Datos pagos Pyme Nacion")
            Exit Sub
        End If
        If Val(AnoCuota.Text) < 2011 Or Val(AnoCuota.Text) > 2050 Then
            m$ = "Ano Incorrecto"
            AAa% = MsgBox(m$, 0, "Datos pagos Pyme Nacion")
            Exit Sub
        End If
    End If
    
    ZParidad = 0
    ZParidadII = 0
    ZCoeParidad = 1
    If TipoOrden.ListIndex = 1 Then
    
        XEmpresa = Wempresa
        
        Select Case Val(XEmpresa)
            Case 2, 4, 8, 9
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        XXFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        spCambios = "ConsultaCambio  " + "'" + XXFecha + "'"
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
            Call Conecta_Empresa
                    Else
            m$ = "Se debe informar la paridad"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Call Conecta_Empresa
            Exit Sub
        End If
        
    End If
    
    
    
    If TipoOrden.ListIndex = 1 Then
        
    
        Dim ZZBusca(10000) As String
        Dim ZZLugarBusca As Integer
        
    
        ' Muestra los nombres en C:\ que representan directorios.
        ZZCodigoExe = "AcroRd32.exe"
        ZZPasaExe = ""
        
        Erase ZZBusca
        ZZLugarBusca = 1
        ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Adobe\"
        CicloBusca = 1
        ZZSalida = "N"
        
        Do
        
            MiRuta = ZZBusca(CicloBusca)
            MiNombre = Dir(MiRuta, vbDirectory) ' Recupera la primera entrada.
            Do While MiNombre <> "" ' Inicia el bucle.
                    
                If MiNombre <> "." And MiNombre <> ".." Then
            
                    If (GetAttr(MiRuta & MiNombre) And vbDirectory) = vbDirectory Then
                        
                        ZZLugarBusca = ZZLugarBusca + 1
                        ZZBusca(ZZLugarBusca) = MiRuta & MiNombre + "\"
                        
                            Else
                            
                        WEspacios = Len(ZZCodigoExe)
                        Da = Len(MiNombre) - WEspacios
                        If UCase(Trim(ZZCodigoExe)) = UCase(Trim(MiNombre)) Then
                            ZZPasaExe = MiRuta & MiNombre
                            ZZSalida = "S"
                            Exit Do
                        End If
                        
                    End If
                
                End If
                MiNombre = Trim(UCase(Dir))  ' Obtiene siguiente entrada.
                
            Loop
    
            If CicloBusca = ZZLugarBusca Or ZZSalida = "S" Then
                Exit Do
                    Else
                CicloBusca = CicloBusca + 1
            End If
    
        Loop
    
    
        
        
        For a = 1 To 99
            
            ZZCodigo = UCase(WVector.TextMatrix(a, 1))
            
            If Trim(ZZCodigo) <> "" Then
            
                If Left$(UCase(ZZCodigo), 2) = "DY" Then
                        
                    MiRuta = "W:\msdssis\msds"  ' Establece la ruta.
                    ZZLee = MiRuta + ZZCodigo + ".pdf"
                    ZZEstado = Dir(ZZLee)
                    If ZZEstado <> "" Then
                        Rem RetVal = Shell(ZZPasaExe + " " + ZZLee + " ", 3)
                            Else
                        m$ = "La materia prima " + ZZCodigo + " no posee hoja de seguridad"
                        AAa% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                        Exit Sub
                    End If
                        
                        Else
            
                    ' Muestra los nombres en C:\ que representan directorios.
                    ZZCodigo = Left$(ZZCodigo, 2) + Mid$(ZZCodigo, 4, 3) + Right$(ZZCodigo, 3)
                    ZZArchiMsDs = ""
                    MiRuta = "W:\msdssis\MSDS mp\"  ' Establece la ruta.
                    MiNombre = Dir(MiRuta, vbDirectory) ' Recupera la primera entrada.
                    Do While MiNombre <> "" ' Inicia el bucle.
                        Dada = MiNombre
                        MiNombre = Trim(UCase(Dir))  ' Obtiene siguiente entrada.
                        
                        WEspacios = Len(ZZCodigo)
                        Da = Len(MiNombre) - WEspacios
                        ZZSalida = "N"
                        
                        For AAa = 1 To Da + 1
                            If Left$(ZZCodigo, WEspacios) = Mid$(MiNombre, AAa, WEspacios) Then
                                ZZArchiMsDs = MiNombre
                                ZZSalida = "S"
                                Exit For
                            End If
                        Next AAa
                        
                        If ZZSalida = "S" Then
                            Exit Do
                        End If
                        
                    Loop
                
                
                    If ZZArchiMsDs <> "" Then
                        ZZLee = MiRuta + ZZArchiMsDs
                        ZZEstado = Dir(ZZLee)
                        If ZZEstado <> "" Then
                            Rem RetVal = Shell(ZZPasaExe + "  " + ZZLee + " ", 3)
                                Else
                            m$ = "La materia prima " + ZZCodigo + " no posee hoja de seguridad"
                            AAa% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                            Exit Sub
                        End If
                            Else
                        m$ = "La materia prima " + ZZCodigo + " no posee hoja de seguridad"
                        AAa% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                        Exit Sub
                    End If
        
                End If
        
            End If
            
        Next a
            
    End If
    
    If Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
    
        For a = 1 To 99
            
            ZZCodigo = UCase(WVector.TextMatrix(a, 1))
            ZZCodSedronar = ""
            
            If Trim(ZZCodigo) <> "" Then
            
                spArticulo = "ConsultaArticulo " + "'" + ZZCodigo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZCodSedronar = rstArticulo!CodSedronar
                    rstArticulo.Close
                End If
                
                If Trim(ZZCodSedronar) <> "" Then
                    m$ = "No se puede efectuar una compra de la materia prima " + ZZCodigo + " en esta planta por tener que informarse luego al sedronar"
                    AAa% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                    Exit Sub
                End If
        
            End If
            
        Next a
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    XEmpresa = Wempresa
    
    ZZFechaLlegada = "  /  /    "
    ZZPagoDespacho = "0"
    ZZImpoDespacho = "0"
    ZZVtoDespacho = "  /  /    "
    ZZPagoLetra = "0"
    ZZImpoLetra = "0"
    ZZVtoLetra = "  /  /    "
    ZZAuxiFecha = "  /  /    "
        
    Rem Borra la Ordenes anteriores
    
    Renglon = 0
    Erase Vector
    
    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                
                    Vector(Renglon, 1) = rstOrden!Articulo
                    Vector(Renglon, 2) = Str$(rstOrden!Cantidad)
                    XDerechos = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                    Vector(Renglon, 3) = Str$(XDerechos)
                    
                    ZZFechaLlegada = IIf(IsNull(rstOrden!FechaLlegada), "  /  /    ", rstOrden!FechaLlegada)
                    ZZPagoDespacho = IIf(IsNull(rstOrden!PagoDespacho), "0", rstOrden!PagoDespacho)
                    ZZImpoDespacho = IIf(IsNull(rstOrden!ImpoDespacho), "0", rstOrden!ImpoDespacho)
                    ZZVtoDespacho = IIf(IsNull(rstOrden!VtoDespacho), "  /  /    ", rstOrden!VtoDespacho)
                    ZZPagoLetra = IIf(IsNull(rstOrden!PagoLetra), "0", rstOrden!PagoLetra)
                    ZZImpoLetra = IIf(IsNull(rstOrden!ImpoLetra), "0", rstOrden!ImpoLetra)
                    ZZVtoLetra = IIf(IsNull(rstOrden!VtoLetra), "  /  /    ", rstOrden!VtoLetra)
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Da = 1 To Renglon
    
        Articulo = Vector(Da, 1)
        Cantidad = Val(Vector(Da, 2))
    
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
                
            WCodigo = Articulo
            WCosto1 = Str$(rstArticulo!Costo1)
            WFecha = ""
            WFecha = rstArticulo!Fecha
            WOrden = Str$(rstArticulo!Orden)
            WPedido = Str$(rstArticulo!Pedido - Cantidad)
            WProveedor = ""
            WProveedor = rstArticulo!Proveedor
            WDate = Date$
            
            rstArticulo.Close
                        
            XParam = "'" + WCodigo + "','" _
                    + WCosto1 + "','" _
                    + WPedido + "','" _
                    + WFecha + "','" _
                    + WOrden + "','" _
                    + WProveedor + "','" _
                    + WDate + "'"
                        
            spArticulo = "ModificaArticuloOrden " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem WCantot = Val(Cantidad)
        Rem WMarca = ""
        Rem WLugar = 0
        Rem Erase Tabla
        Rem
        Rem XParam = "'" + Articulo + "','" _
        rem                 + WMarca + "'"
        Rem spSolic = "ListaSolicitudBajaArticulo " + XParam
        Rem Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstSolic.RecordCount > 0 Then
        Rem
        Rem     With rstSolic
        Rem
        Rem         .MoveFirst
        Rem         If .NoMatch = False Then
        Rem             Do
        Rem
        Rem                 WLugar = WLugar + 1
        Rem                 Tabla(WLugar) = rstSolic!Clave
        Rem
        Rem                 .MoveNext
        Rem
        Rem                 If .EOF = True Then
        Rem                     Exit Do
        Rem                 End If
        Rem
        Rem             Loop
        Rem         End If
        Rem
        Rem     End With
        Rem     rstSolic.Close
        Rem
        Rem End If
        Rem
        Rem For Cicla = WLugar To 1 Step -1
        Rem
        Rem     WClave = Tabla(Cicla)
        Rem
        Rem     spSolic = "ConsultaSolicitud " + "'" + WClave + "'"
        Rem     Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstSolic.RecordCount > 0 Then
        Rem
        Rem         WEntregado = rstSolic!Entregado
        Rem         rstSolic.Close
        Rem
        Rem         If WEntregado <> 0 Then
        Rem
        Rem             If WEntregado >= WCantot Then
        Rem                 WEntregado = WEntregado - WCantot
        Rem                 WMarca = ""
        Rem                 Salida = "S"
        Rem                     Else
        Rem                 WCantot = WCantot - WEntregado
        Rem                 WEntregado = 0
        Rem                 WMarca = ""
        Rem                 Salida = "N"
        Rem             End If
        Rem
        Rem             WEntre = WEntregado
        Rem
        Rem             XParam = "'" + WClave + "','" _
        rem                     + WEntre + "','" _
        rem                     + WMarca + "'"
        Rem
        Rem             spSolic = "ModificaSolicitudEntregado " + XParam
        Rem             Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        Rem
        Rem         End If
        Rem
        Rem         If Salida = "S" Then
        Rem             Exit For
        Rem         End If
        Rem
        Rem     End If
        Rem Next Cicla
        
    Next Da
        
    spOrden = "BorrarOrdenTotal " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenDynaset, dbSQLPassThrough)
    
    If TipoOrden.ListIndex = 1 Then
    
        XEmpresa = Wempresa
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        If Val(Carpeta.Text) = 0 Then
        
            Sql1 = "Select Max(Carpeta) as [CarpetaMayor]"
            Sql2 = " FROM NroCarpeta"
            spNroCarpeta = Sql1 + Sql2
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
            If rstNroCarpeta.RecordCount > 0 Then
                rstNroCarpeta.MoveLast
                ZCarpeta = IIf(IsNull(rstNroCarpeta!CarpetaMayor), "0", rstNroCarpeta!CarpetaMayor)
                Carpeta.Text = ZCarpeta + 1
                rstNroCarpeta.Close
            End If
            
            m$ = "La carpera asignada es la " + Carpeta.Text
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM NroCarpeta"
        ZSql = ZSql + " Where Planta = " + "'" + XEmpresa + "'"
        ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
        spNroCarpeta = ZSql
        Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
        If rstNroCarpeta.RecordCount > 0 Then
            rstNroCarpeta.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE NroCarpeta SET "
            ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "'"
            ZSql = ZSql + " Where Planta = " + "'" + XEmpresa + "'"
            ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
            spNroCarpeta = ZSql
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ZSql + "INSERT INTO NroCarpeta ("
            ZSql = ZSql + "Carpeta ,"
            ZSql = ZSql + "Planta ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Carpeta.Text + "',"
            ZSql = ZSql + "'" + XEmpresa + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "')"
            spNroCarpeta = ZSql
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        Call Conecta_Empresa
        
    End If
    
    
        
    Renglon = 0
    ZSuma = 0
    ZBase = 0
    Erase ZVector
        
    For a = 1 To 99
        
        WVector.Row = a
                
        WVector.Col = 1
        Articulo = UCase(WVector.Text)
        
        WVector.Col = 3
        Cantidad = Val(WVector.Text)
        XCantidad = WVector.Text
                
        WVector.Col = 4
        Precio = Val(WVector.Text)
        XPrecio = WVector.Text
                
        WVector.Col = 5
        Fecha1 = WVector.Text
                
        WVector.Col = 6
        fecha2 = WVector.Text
        If fecha2 <> "" Then
            ZZAuxiFecha = fecha2
        End If
                
        WVector.Col = 7
        Condicion = WVector.Text
                
        WVector.Col = 8
        ZZBultos = WVector.Text
        
        WLugar = WVector.Row
                    
        WWPorceDerechos = XPorceDerechos(WLugar)
        
        WWSolicitud1 = XSolicitud(WLugar, 1)
        WWSolicitud2 = XSolicitud(WLugar, 2)
        WWSolicitud3 = XSolicitud(WLugar, 3)
                
        If Articulo <> "" Then
        
            Renglon = Renglon + 1
            
            If Moneda.ListIndex = 2 Then
                ZSuma = ZSuma + (Val(XCantidad) * (Val(XPrecio) * ZCoeParidad))
                    Else
                ZSuma = ZSuma + (Val(XCantidad) * Val(XPrecio))
            End If
            
            ZVector(Renglon, 1) = Articulo
            ZVector(Renglon, 2) = XCantidad
            ZVector(Renglon, 3) = XPrecio
            ZVector(Renglon, 4) = WWPorceDerechos
        
            WOrden = Orden.Text
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WProveedor = Proveedor.Text
            WArticulo = Articulo
            WCantidad = XCantidad
            WPrecio = XPrecio
            WFecha1 = Fecha1
            WFecha2 = fecha2
            WCondicion = Condicion
            WRecibida = "0"
            XSaldo = "0"
            WLiberada = "0"
            WDevuelta = "0"
            WFechaEntrega = "  /  /    "
            Auxi1 = WOrden
            Auxi = WRenglon
            Call Ceros(Auxi1, 6)
            Call Ceros(Auxi, 2)
            WClave = Auxi1 + Auxi
            WDate = Date$
            WMonedaOrden = Str$(Moneda.ListIndex)
            WTarjeta = Str$(Tarjeta.ListIndex)
            WTipoOrden = Str$(TipoOrden.ListIndex)
            WTipoPago = Str$(TipoPago.ListIndex)
            WCarpeta = Carpeta.Text
            WDerechos = "0"
            WOrigen = Origen.Text
            WCuotas = Cuotas.Text
            WMesCuota = MesCuota.Text
            WAnoCuota = AnoCuota.Text
            WDjai = DJai.Text
            WFechaDjai = FechaDJai.Text
            
            If TipoOrden.ListIndex = 2 Then
                WPrecio = "0"
                WCondicion = ""
            End If
            
            For Cicla = 1 To 100
                If WArticulo = Vector(Cicla, 1) Then
                    WDerechos = Vector(Cicla, 3)
                    Exit For
                End If
            Next Cicla
                     
            XParam = "'" + WClave + "','" _
                     + WOrden + "','" _
                     + WRenglon + "','" _
                     + WFecha + "','" _
                     + WProveedor + "','" _
                     + WArticulo + "','" _
                     + WCantidad + "','" _
                     + WPrecio + "','" _
                     + WFecha1 + "','" _
                     + WFecha2 + "','" _
                     + WCondicion + "','" _
                     + WRecibida + "','" _
                     + XSaldo + "','" _
                     + WFechaord + "','" _
                     + WLiberada + "','" _
                     + WDevuelta + "','" _
                     + WFechaEntrega + "','" _
                     + WDate + "','" _
                     + WMonedaOrden + "','" _
                     + WTipoOrden + "','" _
                     + WCarpeta + "','" _
                     + WDerechos + "','" _
                     + WOrigen + "'"
                     
            spOrden = "AltaOrdenIII " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Orden SET "
            ZSql = ZSql + " Bultos = " + "'" + ZZBultos + "',"
            ZSql = ZSql + " Tarjeta = " + "'" + WTarjeta + "',"
            ZSql = ZSql + " Cuotas = " + "'" + WCuotas + "',"
            ZSql = ZSql + " MesCuota = " + "'" + WMesCuota + "',"
            ZSql = ZSql + " AnoCuota = " + "'" + WAnoCuota + "',"
            ZSql = ZSql + " Derechos = " + "'" + WWPorceDerechos + "',"
            ZSql = ZSql + " Solicitud1 = " + "'" + WWSolicitud1 + "',"
            ZSql = ZSql + " Solicitud2 = " + "'" + WWSolicitud2 + "',"
            ZSql = ZSql + " Solicitud3 = " + "'" + WWSolicitud3 + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            
            WLeyenda = Str$(Leyenda.ListIndex)
            XParam = "'" + WOrden + "','" _
                         + WLeyenda + "'"
            spOrden = "ModificaOrdenLeyenda " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WCodigo = WArticulo
                Rem WCosto1 = Str$(Precio)
                WCosto1 = Str$(rstArticulo!Costo1)
                WFecha = Fecha.Text
                WOrden = Orden.Text
                WPedido = Str$(rstArticulo!Pedido + Cantidad)
                WProveedor = ""
                WProveedor = rstArticulo!Proveedor
                WDate = Date$
                rstArticulo.Close
                
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = WArticulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    WDescripcion = ""
                End If
                
                Select Case Val(Wempresa)
                    Case 1
                        WPlanta = "SI"
                    Case 2
                        WPlanta = "PI"
                    Case 3
                        WPlanta = "SII"
                    Case 4
                        WPlanta = "PII"
                    Case 5
                        WPlanta = "SIII"
                    Case 6
                        WPlanta = "SIV"
                    Case 7
                        WPlanta = "SV"
                    Case 8
                        WPlanta = "PV"
                    Case 9
                        WPlanta = "PIV"
                    Case 10
                        WPlanta = "SVI"
                    Case 11
                        WPlanta = "SVII"
                    Case Else
                        WPlanta = ""
                End Select
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Orden SET "
                ZSql = ZSql + " Impremarca = " + "'" + WDescripcion + "',"
                ZSql = ZSql + " Planta = " + "'" + WPlanta + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
                Rem dada
                Rem dada
                Rem dada
                If TipoOrden.ListIndex = 2 Then
                
                    WPedido = Str$(Cantidad)
                    XParam = "'" + WCodigo + "','" _
                        + WPedido + "'"
                        
                    spArticulo = "ModificaArticuloPedido " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
                
                    XParam = "'" + WCodigo + "','" _
                        + WCosto1 + "','" _
                        + WPedido + "','" _
                        + WFecha + "','" _
                        + WOrden + "','" _
                        + WProveedor + "','" _
                        + WDate + "'"
                    
                    spArticulo = "ModificaArticuloOrden " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    XParam = "'" + WCodigo + "','" _
                            + WDescripcion + "'"
                    
                    spArticulo = "ModificaArticuloDescriComercial " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            End If
        
            Rem Actualiza la solicitud de orden de compra

            Rem WCantot = Val(XCantidad)
            Rem WMarca = "X"
            Rem WLugar = 0
            Rem Erase Tabla
            
            Rem XParam = "'" + WArticulo + "','" _
            rem              + WMarca + "'"
            Rem spSolic = "ListaSolicitudPendienteArticulo " + XParam
            Rem Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstSolic.RecordCount > 0 Then
            Rem
            Rem     With rstSolic
            Rem
            Rem         .MoveFirst
            Rem         If .NoMatch = False Then
            Rem             Do
            Rem
            Rem                 WLugar = WLugar + 1
            Rem                 Tabla(WLugar) = rstSolic!Clave
            Rem
            Rem                 .MoveNext
            Rem
            Rem                 If .EOF = True Then
            Rem                     Exit Do
            Rem                 End If
            Rem
            Rem             Loop
            Rem         End If
            Rem
            Rem     End With
            Rem     rstSolic.Close
            Rem
            Rem End If
            Rem
            Rem For Cicla = 1 To WLugar
            Rem
            Rem     WClave = Tabla(Cicla)
            Rem
            Rem     spSolic = "ConsultaSolicitud " + "'" + WClave + "'"
            Rem     Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstSolic.RecordCount > 0 Then
            Rem
            Rem         WCanti = rstSolic!Cantidad - rstSolic!Entregado
            Rem         WEntregado = rstSolic!Entregado
            Rem         rstSolic.Close
            Rem
            Rem
            Rem         If WCanti > WCantot Then
            Rem             WEntregado = WEntregado + WCantot
            Rem             WMarca = ""
            Rem             Salida = "S"
            Rem                 Else
            Rem             WEntregado = WEntregado + WCanti
            Rem             WCantot = WCantot - WCanti
            Rem             WMarca = "X"
            Rem             Salida = "N"
            Rem         End If
            Rem
            Rem         WEntre = WEntregado
            Rem
            Rem         XParam = "'" + WClave + "','" _
            rem                 + WEntre + "','" _
            rem                 + WMarca + "'"
            Rem
            Rem         spSolic = "ModificaSolicitudEntregado " + XParam
            Rem         Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            Rem
            Rem         If Salida = "S" Then
            Rem             Exit For
            Rem         End If
            Rem     End If
            Rem Next Cicla
            
            If Val(WWSolicitud1) <> 0 Then
                ZSolicitud = WWSolicitud1
                ZZArticulo = WArticulo
                Call Baja_Solicitud
            End If
            
            If Val(WWSolicitud2) <> 0 Then
                ZSolicitud = WWSolicitud2
                ZZArticulo = WArticulo
                Call Baja_Solicitud
            End If
            
            If Val(WWSolicitud3) <> 0 Then
                ZSolicitud = WWSolicitud3
                ZZArticulo = WArticulo
                Call Baja_Solicitud
            End If
                  
        End If
            
    Next a
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " TotalFob = " + "'" + Str$(ZSuma) + "'"
    ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    If TipoOrden.ListIndex = 1 Then
    
        If Leyenda.ListIndex = 1 Or Leyenda.ListIndex = 2 Or Leyenda.ListIndex = 5 Then
            If ZSuma > 75000 Then
                m$ = "Se debe contratar un seguro adicional ya que " + Chr$(13) + _
                     "el monto de la orden de compra supera los U$S 75000.-"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            End If
        End If
        
        ZZFull = "N"
        T$ = "Orden de Compra"
        m$ = "La importacion es Full Container"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            ZZFull = "S"
        End If
        
        
        
        
        
        
    
        ZDespacho = 0
        ZBase = 0
        ZBaseII = 0
        ZBaseIII = 0
        ZBaseIV = 0
        
        ZSuma1 = 0
        ZSuma2 = 0
        ZSuma3 = 0
        ZSuma4 = 0
        ZSuma5 = 0
        ZSuma6 = 0
        ZSuma7 = 0
        
        ZSumaCantidad = 0
        
        If ZSuma <> 0 Then
    
            ZRegion = 0
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            With rstProveedor
                If rstProveedor.RecordCount > 0 Then
                    ZRegion = 0
                    ZRegion = IIf(IsNull(!Region), "0", !Region)
                End If
                rstProveedor.Close
            End With
    
            For Ciclo = 1 To Renglon
    
                XXArticulo = ZVector(Ciclo, 1)
                XXCantidad = ZVector(Ciclo, 2)
                XXPrecio = ZVector(Ciclo, 3)
                ZPorceDerechos = Val(ZVector(Ciclo, 4))
                If ZRegion = 1 Then
                    ZPorceDerechos = 0
                End If
            
                If ZRegion = 0 And ZPorceDerechos = 0 Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Articulo = " + "'" + XXArticulo + "'"
                    ZSql = ZSql + " and Derechos <> 0"
                    ZSql = ZSql + " Order by Orden.FechaOrd"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        With rstOrden
                            .MoveLast
                            ZPorceDerechos = rstOrden!Derechos
                        End With
                        rstOrden.Close
                    End If
                End If
            
                If Moneda.ListIndex = 2 Then
                    ZImpo = Val(XXCantidad) * (Val(XXPrecio) * ZCoeParidad)
                        Else
                    ZImpo = Val(XXCantidad) * Val(XXPrecio)
                End If
                ZBaseII = ZBaseII + ZImpo
            
                ZSeguro = 0
                If Val(Flete.Text) <> 0 Then
                    ZPorce = ZImpo / ZSuma
                    ZFlete = Val(Flete.Text) * ZPorce
                    ZImpo = ZImpo + ZFlete
                End If
                If Leyenda.ListIndex <> 2 Then
                    ZSeguro = ZImpo * 0.01
                End If
                ZBaseIV = ZBaseIV + ZSeguro
        
                ZImpo = ZImpo + ZSeguro
                ZBaseIII = ZBaseIII + ZImpo
            
                ZDerechos = ZImpo * (ZPorceDerechos / 100)
                ZSuma1 = ZSuma1 + ZDerechos
                
                If ZRegion = 1 Then
                    ZEstadistica = 0
                        Else
                    ZEstadistica = ZImpo * 0.005
                End If
                ZSuma2 = ZSuma2 + ZEstadistica
        
                ZImpo = ZImpo + ZDerechos + ZEstadistica
        
                ZIva = ZImpo * 0.21
                ZIvaComp = ZImpo * 0.1
                ZGanancia = ZImpo * 0.03
                ZIBruto = ZImpo * 0.015
                
                ZSuma3 = ZSuma3 + ZIva
                ZSuma4 = ZSuma4 + ZIvaComp
                ZSuma5 = ZSuma5 + ZGanancia
                ZSuma6 = ZSuma6 + ZIBruto
        
                ZImpo = ZImpo + ZIva + ZIvaComp + ZGanancia + ZIBruto
                ZBase = ZBase + ZImpo
                
                ZSumaCantidad = ZSumaCantidad + Val(XXCantidad)
        
            Next Ciclo
            
            ZCargo = 10
            ZBase = ZBase + ZCargo - ZBaseIII
    
            ZImpoII = ZBase * ZParidad
            ZImpoIV = ZBaseIII * ZParidad
    
            If ZZFull = "S" Then
    
                ZGastos = 0
                ZHonorarios = 100 + (ZImpoIV * 0.006)
                Rem ZIvaGastos = (ZGastos + ZHonorarios) * 0.21
                ZIvaGastos = ZHonorarios * 0.21
        
                Select Case TipoImpo.ListIndex
                    Case 1
                        ZViaI = 1200 + 2000
                        ZViaII = 0
                    Case 2
                        ZViaI = 250 + 1500
                        ZViaII = 0
                    Case 3
                        ZViaI = (200 * ZParidad) + 500
                        ZViaII = 145 * ZParidad
                    Case Else
                End Select
    
                ZDespachoI = ZGastos + ZHonorarios + ZIvaGastos + ZViaI + ZViaII + (Val(Flete.Text) * ZParidad)
                Rem  (ZBaseIV * ZParidad)
                ZDespacho = ZImpoII + ZDespachoI
                ZDespacho = ZDespacho / ZParidad
                
                    Else
    
                Rem REPRESENTA FISCAL
                Rem POR NO SER FULL CONTAINER
                Rem 450 + IVA Y ALGINOS GASTOS MAS
                ZGastos = 1000
                ZHonorarios = 100 + (ZImpoIV * 0.006)
                ZIvaGastos = ZHonorarios * 0.21
        
                ZViaI = 0
                ZViaII = 0
    
                ZDespachoI = ZGastos + ZHonorarios + ZIvaGastos + ZViaI + ZViaII + (Val(Flete.Text) * ZParidad)
                Rem  (ZBaseIV * ZParidad)
                ZDespacho = ZImpoII + ZDespachoI
                ZDespacho = ZDespacho / ZParidad
                ZZCargo = 75 * ((ZSumaCantidad / 1000) * 2)
                ZDespacho = ZDespacho + ZZCargo
            
            End If
            
            ZDespachoOrden = ZDespacho * ZParidad
            If Leyenda.ListIndex > 0 Then
                ZZLeyenda = Str$(Leyenda.ListIndex - 1)
                    Else
                ZZLeyenda = "0"
            End If
            ZZMoneda = Moneda.Text
            
            For Ciclo = 1 To Renglon
    
                XXArticulo = ZVector(Ciclo, 1)
                XXCantidad = ZVector(Ciclo, 2)
                XXPrecio = ZVector(Ciclo, 3)
                If Moneda.ListIndex = 2 Then
                    XXImpo = Val(XXCantidad) * (Val(XXPrecio) * ZCoeParidad)
                        Else
                    XXImpo = Val(XXCantidad) * Val(XXPrecio)
                End If
            
                XEmpresa = Wempresa
                
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Leyenda = " + "'" + ZZLeyenda + "',"
                ZSql = ZSql + " Moneda = " + "'" + ZZMoneda + "',"
                ZSql = ZSql + " Flete = " + "'" + XXPrecio + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                Call Conecta_Empresa
                
            Next Ciclo
            
        End If
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
    
        ZGraba = "N"
        T$ = "Orden de Compra"
        m$ = "Desea actualizar el costo en transito"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            ZGraba = "S"
        End If
        
        If ZGraba = "S" Then
    
            ZDespacho = 0
            ZBase = 0
            ZBaseII = 0
            ZBaseIII = 0
            ZBaseIV = 0
        
            ZSuma1 = 0
            ZSuma2 = 0
            ZSuma3 = 0
            ZSuma4 = 0
            ZSuma5 = 0
            ZSuma6 = 0
            ZSuma7 = 0
        
            ZSumaCantidad = 0
        
            If ZSuma <> 0 Then
    
                ZRegion = 0
                spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                With rstProveedor
                    If rstProveedor.RecordCount > 0 Then
                        ZRegion = 0
                        ZRegion = IIf(IsNull(!Region), "0", !Region)
                    End If
                    rstProveedor.Close
                End With
    
                For Ciclo = 1 To Renglon
    
                    XXArticulo = ZVector(Ciclo, 1)
                    XXCantidad = ZVector(Ciclo, 2)
                    XXPrecio = ZVector(Ciclo, 3)
                    ZPorceDerechos = Val(ZVector(Ciclo, 4))
                    If ZRegion = 1 Then
                        ZPorceDerechos = 0
                    End If
            
                    If ZRegion = 0 And ZPorceDerechos = 0 Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Articulo = " + "'" + XXArticulo + "'"
                        ZSql = ZSql + " and Derechos <> 0"
                        ZSql = ZSql + " Order by Orden.FechaOrd"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            With rstOrden
                                .MoveLast
                                ZPorceDerechos = rstOrden!Derechos
                            End With
                            rstOrden.Close
                        End If
                    End If
            
                    If Moneda.ListIndex = 2 Then
                        ZImpo = Val(XXCantidad) * (Val(XXPrecio) * ZCoeParidad)
                            Else
                        ZImpo = Val(XXCantidad) * Val(XXPrecio)
                    End If
                    
                    ZBaseII = ZBaseII + ZImpo
            
                    ZSeguro = 0
                    If Val(Flete.Text) <> 0 Then
                        ZPorce = ZImpo / ZSuma
                        ZFlete = Val(Flete.Text) * ZPorce
                        ZImpo = ZImpo + ZFlete
                    End If
                    If Leyenda.ListIndex <> 2 Then
                        ZSeguro = ZImpo * 0.01
                    End If
                    ZBaseIV = ZBaseIV + ZSeguro
        
                    ZImpo = ZImpo + ZSeguro
                    ZBaseIII = ZBaseIII + ZImpo
            
                    ZDerechos = ZImpo * (ZPorceDerechos / 100)
                    ZSuma1 = ZSuma1 + ZDerechos
                
                    If ZRegion = 1 Then
                        ZEstadistica = 0
                            Else
                        ZEstadistica = ZImpo * 0.005
                    End If
                    ZSuma2 = ZSuma2 + ZEstadistica
        
                    ZImpo = ZImpo + ZDerechos + ZEstadistica
        
                    Rem ZIva = ZImpo * 0.21
                    Rem ZIvaComp = ZImpo * 0.1
                    Rem ZGanancia = ZImpo * 0.03
                    Rem ZIBruto = ZImpo * 0.015
                
                    ZIva = 0
                    ZIvaComp = 0
                    ZGanancia = 0
                    ZIBruto = 0
                
                    ZSuma3 = ZSuma3 + ZIva
                    ZSuma4 = ZSuma4 + ZIvaComp
                    ZSuma5 = ZSuma5 + ZGanancia
                    ZSuma6 = ZSuma6 + ZIBruto
        
                    ZImpo = ZImpo + ZIva + ZIvaComp + ZGanancia + ZIBruto
                    ZBase = ZBase + ZImpo
                
                    ZSumaCantidad = ZSumaCantidad + Val(XXCantidad)
        
                Next Ciclo
            
                ZCargo = 10
                ZBase = ZBase + ZCargo - ZBaseIII
    
                ZImpoII = ZBase * ZParidad
                ZImpoIV = ZBaseIII * ZParidad
    
                If ZZFull = "S" Then
    
                    ZGastos = 0
                    ZHonorarios = 100 + (ZImpoIV * 0.006)
                    Rem ZIvaGastos = (ZGastos + ZHonorarios) * 0.21
                    ZIvaGastos = ZHonorarios * 0.21
        
                    Select Case TipoImpo.ListIndex
                        Case 1
                            ZViaI = 1200 + 2000
                            ZViaII = 0
                        Case 2
                            ZViaI = 250 + 1500
                            ZViaII = 0
                        Case 3
                            ZViaI = (200 * ZParidad) + 500
                            ZViaII = 145 * ZParidad
                        Case Else
                    End Select
    
                    ZDespachoI = ZGastos + ZHonorarios + ZIvaGastos + ZViaI + ZViaII + (Val(Flete.Text) * ZParidad)
                    Rem  (ZBaseIV * ZParidad)
                    ZDespacho = ZImpoII + ZDespachoI
                    ZDespacho = ZDespacho / ZParidad
                
                        Else
    
                    Rem REPRESENTA FISCAL
                    Rem POR NO SER FULL CONTAINER
                    Rem 450 + IVA Y ALGINOS GASTOS MAS
                    ZGastos = 1000
                    ZHonorarios = 100 + (ZImpoIV * 0.006)
                    ZIvaGastos = ZHonorarios * 0.21
        
                    ZViaI = 0
                    ZViaII = 0
    
                    ZDespachoI = ZGastos + ZHonorarios + ZIvaGastos + ZViaI + ZViaII + (Val(Flete.Text) * ZParidad)
                    Rem  (ZBaseIV * ZParidad)
                    ZDespacho = ZImpoII + ZDespachoI
                    ZDespacho = ZDespacho / ZParidad
                    ZZCargo = 75 * ((ZSumaCantidad / 1000) * 2)
                    ZDespacho = ZDespacho + ZZCargo
            
                End If
            
                For Ciclo = 1 To Renglon
    
                    XXArticulo = ZVector(Ciclo, 1)
                    XXCantidad = ZVector(Ciclo, 2)
                    XXPrecio = ZVector(Ciclo, 3)
                    If Moneda.ListIndex = 2 Then
                        XXImpo = Val(XXCantidad) * (Val(XXPrecio) * ZCoeParidad)
                            Else
                        XXImpo = Val(XXCantidad) * Val(XXPrecio)
                    End If
                    
                    
                
                    Rem ZPorce = Val(XXCantidad) / ZSumaCantidad
                    ZPorce = XXImpo / ZSuma
                    ZZCosto = ZDespacho * ZPorce
                    ZZCosto = ((XXImpo + ZZCosto) / Val(XXCantidad)) * 1.03
                    Call Redondeo(ZZCosto)
                
                    If Leyenda.ListIndex > 0 Then
                        XLeyenda = Str$(Leyenda.ListIndex - 1)
                            Else
                        XLeyenda = "0"
                    End If
                    ZZMoneda = Moneda.Text
                
                    XEmpresa = Wempresa
                
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Costo6 = " + "'" + Str$(ZZCosto) + "',"
                    ZSql = ZSql + " Flete = " + "'" + XXPrecio + "',"
                    ZSql = ZSql + " Leyenda = " + "'" + XLeyenda + "',"
                    ZSql = ZSql + " TipoCosto = " + "'" + "1" + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                    Call Conecta_Empresa
                    
                Next Ciclo
    
            End If
        
        End If
        
            Else
            
        ZDespacho = 0
        ZDespachoOrden = 0
        
    End If
    
                    
    
    ZDespacho = Int(ZDespacho)
    ZDespachoOrden = Int(ZDespachoOrden)

    ZOrigen = Origen.Text
    ZDJai = DJai.Text
    ZFechaDJai = FechaDJai.Text
    ZLeyenda = Str$(Leyenda.ListIndex)
    ZPedidoImpo = PedidoImpo.Text
    ZFlete = Flete.Text
    ZFechaImpo = FechaImpo.Text
    ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
    ZTipoImpo = Str$(TipoImpo.ListIndex)
    ZTipoPago = Str$(TipoPago.ListIndex)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " ImpoDespacho = " + "'" + Str$(ZDespachoOrden) + "',"
    ZSql = ZSql + " DJai = " + "'" + ZDJai + "',"
    ZSql = ZSql + " FechaDjai = " + "'" + ZFechaDJai + "',"
    ZSql = ZSql + " Flete = " + "'" + ZFlete + "',"
    ZSql = ZSql + " Origen = " + "'" + ZOrigen + "',"
    ZSql = ZSql + " Leyenda = " + "'" + ZLeyenda + "',"
    ZSql = ZSql + " PedidoImpo = " + "'" + ZPedidoImpo + "',"
    ZSql = ZSql + " FechaImpo = " + "'" + ZFechaImpo + "',"
    ZSql = ZSql + " OrdFechaImpo = " + "'" + ZOrdFechaImpo + "',"
    ZSql = ZSql + " TipoImpo = " + "'" + ZTipoImpo + "',"
    ZSql = ZSql + " TipoPago = " + "'" + ZTipoPago + "'"
    
    ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem ZZFechaLlegada = "  /  /    "
    Rem ZZPagoDespacho = "0"
    Rem ZZImpoDespacho = "0"
    Rem ZZVtoDespacho = "  /  /    "
    Rem ZZImpoLetra = "0"
    Rem ZZPagoLetra = "0"
    Rem ZZVtoLetra = "  /  /    "
    
    If ZZFechaLlegada = "  /  /    " Or Trim(ZZFechaLlegada) = "" Then
        ZZFechaLlegada = ZZAuxiFecha
    End If
    
    If TipoPago.ListIndex = 1 Then
        ZZVtoLetra = Fecha.Text
    End If
    If TipoPago.ListIndex = 2 Then
        ZZVtoLetra = ZZFechaLlegada
    End If
    ZZImpoLetra = Str$(ZSuma)
    ZZVtoDespacho = ZZFechaLlegada

    ZZOrdVtoDespacho = Right$(ZZVtoDespacho, 4) + Mid$(ZZVtoDespacho, 4, 2) + Left$(ZZVtoDespacho, 2)
    ZZOrdVtoLetra = Right$(ZZVtoLetra, 4) + Mid$(ZZVtoLetra, 4, 2) + Left$(ZZVtoLetra, 2)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " FechaLlegada = " + "'" + ZZFechaLlegada + "',"
    ZSql = ZSql + " PagoDespacho = " + "'" + ZZPagoDespacho + "',"
    ZSql = ZSql + " VtoDespacho = " + "'" + ZZVtoDespacho + "',"
    ZSql = ZSql + " OrdVtoDespacho = " + "'" + ZZOrdVtoDespacho + "',"
    ZSql = ZSql + " PagoLetra = " + "'" + ZZPagoLetra + "',"
    ZSql = ZSql + " ImpoLetra = " + "'" + ZZImpoLetra + "',"
    ZSql = ZSql + " VtoLetra = " + "'" + ZZVtoLetra + "',"
    ZSql = ZSql + " OrdVtoLetra = " + "'" + ZZOrdVtoLetra + "'"
    ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    WOrden = Orden.Text
    WImpresion = "N"
    XParam = "'" + WOrden + "','" _
                 + WImpresion + "'"
    spOrden = "ModificaOrdenImpresion " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                
    T$ = "Orden de Compra"
    m$ = "Desea imprimir la Orden de Compra"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion_Crystal
    End If
                
    If TipoOrden.ListIndex = 1 Then
        T$ = "Orden de Compra"
        m$ = "Desea imprimir la Orden de Compra de Importacion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion_Impo
        End If
    End If
        
    T$ = "Orden de Compra"
    m$ = "Desea enviar la O/C via email al proveedor"
    Respuesta% = MsgBox(m$, 256 + 4, T$)
    If Respuesta% = 6 Then
        Call EMail_Click
        Rem ChDir "\\PRUEBA\E\VB"
            Else
        Call Limpia_Click
        Orden.SetFocus
    
        Call Conecta_Empresa
  
    End If
        
        
   Exit Sub

WError:
    Resume Next
    
End Sub


Private Sub Baja_Solicitud()

    XEmpresa = Wempresa
    
    For Cicla = 1 To 11
    
        Select Case Cicla
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
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Solic"
        ZSql = ZSql + " Where Solic.Solicitud = " + "'" + ZSolicitud + "'"
        ZSql = ZSql + " and Solic.Articulo = " + "'" + ZZArticulo + "'"
        ZSql = ZSql + " and Solic.Marca <> " + "'" + "X" + "'"
        ZSql = ZSql + " Order by Solic.Clave"
        spSolic = ZSql
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            ZZClaveSolic = rstSolic!Clave
            rstSolic.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Solic SET "
            ZSql = ZSql + " Entregado = Cantidad" + ","
            ZSql = ZSql + " Marca = " + "'" + "X" + "'"
            ZSql = ZSql + " Where Solic.Clave = " + "'" + ZZClaveSolic + "'"
            spSolic = ZSql
            Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Cicla
    
    Call Conecta_Empresa
    
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    WBultos.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    DatosImpo.Visible = False

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    WBultos.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""
    Cuotas.Text = ""
    MesCuota.Text = ""
    AnoCuota.Text = ""

    Orden.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    ZEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    Flete.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""
    Cuotas.Text = ""
    MesCuota.Text = ""
    AnoCuota.Text = ""
    DJai.Text = ""
    FechaDJai.Text = "  /  /    "
    
    Moneda.ListIndex = 0
    TipoOrden.ListIndex = 0
    TipoPago.ListIndex = 0
    Leyenda.ListIndex = 0
    TipoImpo.ListIndex = 0
    Tarjeta.ListIndex = 0
    
    Call Limpia_Vector
    
    Rem With rstOrden
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Orden.Text = !Orden + 1
    Rem             Else
    Rem         Orden.Text = ""
    Rem     End If
    Rem End With
    
    Orden.Text = "1"
    
    Rem spOrden = "ListaOrdenNUmero"
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstOrden.RecordCount > 0 Then
    Rem     With rstOrden
    Rem         .MoveLast
    Rem         Orden.Text = rstOrden!Orden + 1
    Rem     End With
    Rem     rstOrden.Close
    Rem         Else
    Rem     Orden.Text = "1"
    Rem End If
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden < 800000"
    ZSql = ZSql + " Order by Orden.Clave"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveLast
            Orden.Text = Str$(rstOrden!Orden + 1)
        End With
        rstOrden.Close
    End If
    
    Renglon = 0
    Graba.Enabled = True

    Orden.SetFocus

End Sub

Private Sub OrdenImportacion_Click()
Rem by nan
       TipoOrden.ListIndex = 1
  Rem  If TipoOrden.ListIndex = 1 Then
  Rem      T$ = "Orden de Compra"
  Rem      m$ = "Desea imprimir la Orden de Compra de Importacion"
  Rem      Respuesta% = MsgBox(m$, 32 + 4, T$)
  Rem      If Respuesta% = 6 Then
           Call Impresion_Impo
           Call Impresion_Especificaciones
 Rem       End If
 Rem   End If
End Sub

Private Sub Tarjeta_click()
    If Tarjeta.ListIndex = 1 Then
        PantaPyme.Visible = True
        Cuotas.SetFocus
    End If
End Sub

Private Sub Tarjeta_change()
    If Tarjeta.ListIndex = 1 Then
        PantaPyme.Visible = True
        Cuotas.SetFocus
    End If
End Sub

Private Sub TipoOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If (TipoOrden.ListIndex = 3 Or TipoOrden.ListIndex = 4) And Val(Orden.Text) < 800000 Then
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden < 900000"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveLast
                If rstOrden!Orden >= 800000 Then
                    Orden.Text = Mid$(Str$(rstOrden!Orden + 1), 2, 6)
                        Else
                    Orden.Text = "800000"
                End If
            End With
            rstOrden.Close
        End If
            Else
        If TipoOrden.ListIndex <> 3 And TipoOrden.ListIndex <> 4 And Val(Orden.Text) >= 800000 Then
            ZSql = ""
            ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Orden < 800000"
            ZSql = ZSql + " Order by Orden.Clave"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                    .MoveLast
                    Orden.Text = Str$(rstOrden!Orden + 1)
                End With
                rstOrden.Close
            End If
        End If
    End If
    End If
End Sub

Private Sub TipoOrden_Click()
    If (TipoOrden.ListIndex = 3 Or TipoOrden.ListIndex = 4) And Val(Orden.Text) < 800000 Then
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden < 900000"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveLast
                If rstOrden!Orden >= 800000 Then
                    Orden.Text = Mid$(Str$(rstOrden!Orden + 1), 2, 6)
                        Else
                    Orden.Text = "800000"
                End If
            End With
            rstOrden.Close
        End If
            Else
        If TipoOrden.ListIndex <> 3 And TipoOrden.ListIndex <> 4 And Val(Orden.Text) >= 800000 Then
            ZSql = ""
            ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Orden < 800000"
            ZSql = ZSql + " Order by Orden.Clave"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                    .MoveLast
                    Orden.Text = Str$(rstOrden!Orden + 1)
                End With
                rstOrden.Close
            End If
        End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ingre = "N"
        WArticulo.Text = UCase(WArticulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            Ingre = "S"
            rstArticulo.Close
                Else
            WArticulo.SetFocus
        End If
        If Ingre = "S" Then
            If TipoOrden.ListIndex <> 2 Then
                Call Calcula_Precio(Proveedor.Text, WArticulo.Text, Precio, Condicion, XMoneda)
                If Precio = 0 Then
                    Desdelugar = 1
                    XCoti.Height = 4200
                    XCoti.Left = 2360
                    XCoti.Top = 1320
                    XCoti.Width = 7455
                    XCoti.Visible = True
                    XProve.Text = Proveedor.Text
                    XDesProve.Caption = DesProveedor.Caption
                    XArti = WArticulo.Text
                    XPrec.Text = ""
                    XCondicion.Text = ""
                    XObservaciones.Text = ""
                    XPrec.SetFocus
                        Else
                    If Moneda.ListIndex = 3 Then
                        Moneda.ListIndex = XMoneda
                    End If
                    If Moneda.ListIndex = XMoneda Then
                        WPrecio.Caption = Pusing("###,###.###", Str$(Precio))
                        WCondicion.Caption = Condicion
                        WCantidad.SetFocus
                            Else
                        m$ = "La moneda de la cotizacion no se corresponde a la moneda de los otros productos cargados en esta orden de compra"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        WArticulo.SetFocus
                    End If
                End If
                    Else
                WPrecio.Caption = ""
                WPrecio.Caption = Pusing("###,###.###", WPrecio.Caption)
                WCondicion.Caption = ""
                WCantidad.SetFocus
            End If
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Rem WFecha1.SetFocus
        CargaSolicitud.Height = 2535
        CargaSolicitud.Left = 3480
        CargaSolicitud.Top = 1320
        CargaSolicitud.Width = 3135
        CargaSolicitud.Visible = True
        Solicitud1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WFecha1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(WFecha1.Text, Auxi)
        If Auxi = "S" Then
            WFecha2.SetFocus
                Else
            WFecha1.SetFocus
        End If
    End If
End Sub

Private Sub WFecha2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(WFecha2.Text, Auxi)
        If Auxi = "S" Then
            WBultos.SetFocus
        End If
    End If
End Sub

Private Sub WBultos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = "S"
        If Auxi = "S" Then
            
            If TipoOrden.ListIndex = 1 Then
            
                ZRegion = 0
                ZZPorceDerechos = Val(WPorceDerechos.Text)
                spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                With rstProveedor
                    If rstProveedor.RecordCount > 0 Then
                        ZRegion = 0
                        ZRegion = IIf(IsNull(!Region), "0", !Region)
                    End If
                    rstProveedor.Close
                End With
            
                If ZRegion = 0 And ZZPorceDerechos = 0 Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " and Derechos <> 0"
                    ZSql = ZSql + " Order by Orden.FechaOrd"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        With rstOrden
                            .MoveLast
                            ZZPorceDerechos = rstOrden!Derechos
                        End With
                        rstOrden.Close
                    End If
                End If

                If ZRegion = 0 And ZZPorceDerechos = 0 Then
                    WPorceDerechos.Text = ""
                    IngreDerechos.Visible = True
                    WPorceDerechos.SetFocus
                    Exit Sub
                End If
            
                If ZRegion = 1 Then
                    WPorceDerechos.Text = ""
                End If
            
                WPorceDerechos.Text = Str$(ZZPorceDerechos)
                
            End If
        
            Call Valida_fecha(WFecha1.Text, Auxi)
            If Auxi <> "S" Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha1.SetFocus
                Exit Sub
            End If
        
            Call Valida_fecha(WFecha2.Text, Auxi)
            If Auxi <> "S" Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha2.SetFocus
                Exit Sub
            End If
            
            ZOrdFechaI = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZOrdFechaII = Right$(WFecha1.Text, 4) + Mid$(WFecha1.Text, 4, 2) + Left$(WFecha1.Text, 2)
            If ZOrdFechaII < ZOrdFechaI Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha1.SetFocus
                Exit Sub
            End If
            
            ZOrdFechaI = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZOrdFechaII = Right$(WFecha2.Text, 4) + Mid$(WFecha2.Text, 4, 2) + Left$(WFecha2.Text, 2)
            If ZOrdFechaII < ZOrdFechaI Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha2.SetFocus
                Exit Sub
            End If
            
            ZZIngre = "N"
            
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                rstArticulo.Close
                
                If ZZTipoMp = 1 Then
                
                    XEmpresa = Wempresa
                    Select Case Val(Wempresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            Wempresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            Wempresa = "0004"
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
            
                    
                    ZZRecha = ""
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Homologa"
                    ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and CodigoMp = " + "'" + WArticulo.Text + "'"
                    spHomologa = ZSql
                    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHomologa.RecordCount > 0 Then
                        If rstHomologa!Estado = "1" Then
                            ZZIngre = "S"
                                Else
                            ZZIngre = "N"
                            ZZRecha = "S"
                        End If
                        rstHomologa.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    If ZZIngre = "N" Then
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " and Proveedor = " + "'" + Proveedor.Text + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZIngre = "S"
                            rstOrden.Close
                        End If
                        
                        If ZZIngre = "N" Then
                            
                            XEmpresa = Wempresa
                            For Cicla = 1 To 11
                            
                                Select Case Cicla
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
                                    Case 11
                                        Wempresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                End Select
                                    
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + Proveedor.Text + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "S"
                                    rstOrden.Close
                                End If
                                
                            Next Cicla
                            
                            Call Conecta_Empresa
                            
                        End If
                        
                        
                        If Proveedor.Text = "10071408390" And ZZIngre = "N" Then
                        
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Orden"
                            ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                            ZSql = ZSql + " and Proveedor = " + "'" + "10059462341" + "'"
                            spOrden = ZSql
                            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            If rstOrden.RecordCount > 0 Then
                                ZZIngre = "S"
                                rstOrden.Close
                            End If
                            
                            If ZZIngre = "N" Then
                                
                                XEmpresa = Wempresa
                                For Cicla = 1 To 11
                                
                                    Select Case Cicla
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
                                        Case 11
                                            Wempresa = "0011"
                                            txtOdbc = "Empresa11"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                    End Select
                                        
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Orden"
                                    ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                                    ZSql = ZSql + " and Proveedor = " + "'" + "10059462341" + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstOrden.RecordCount > 0 Then
                                        ZZIngre = "S"
                                        rstOrden.Close
                                    End If
                                    
                                Next Cicla
                                
                                Call Conecta_Empresa
                                
                            End If
                            
                        End If
                            
                    End If
                    
                        Else
                        
                    ZZIngre = "S"
                    
                End If
                    
                If ZZIngre = "N" Or ZZRecha = "S" Then
                
                    If TipoOrden.ListIndex = 1 Then
                    
                        T$ = "Ingreso de Orden de Compra"
                        m$ = "Materia Prima homologable y no existe muestra aceptada. Desea continuar"
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
                            m$ = "Coloque en homologacion los codigos de Materia Prima a Homologar"
                            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                            ZZIngre = "S"
                                Else
                            Exit Sub
                        End If
                        
                            Else
                            
                        m$ = "Materia Prima homologable y no existe muestra aceptada"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        Exit Sub
                        
                    End If
                    
                End If
                
                    Else
                    
                m$ = "Materia Prima Inexistentre"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                Exit Sub
                
            End If
        
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
            
        End If
    End If
End Sub

Private Sub WPorceDerechos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WPorceDerechos.Text) <> 0 Then
            IngreDerechos.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
                Else
            WPorceDerechos.SetFocus
        End If
    End If
End Sub

Private Sub pantalla_Click()
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Select Case Val(TipoConsulta)
                    Case 2
                        XProv1.Text = WProveedor
                        XDesProv1.Caption = rstProveedor!Nombre
                        XProv1.SetFocus
                    Case 3
                        XProve.Text = WProveedor
                        XDesProve.Caption = rstProveedor!Nombre
                        XProve.SetFocus
                    Case 4
                        XProv3.Text = WProveedor
                        XDesProv3.Caption = rstProveedor!Nombre
                        XProv3.SetFocus
                    Case Else
                        Proveedor.Text = WProveedor
                        DesProveedor.Caption = rstProveedor!Nombre
                        ZEmail = rstProveedor!EMail
                        Proveedor.SetFocus
                End Select
                rstProveedor.Close
            End If
            
            Call Verifica_Atrasos
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            Opcion.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Select Case Val(TipoConsulta)
                    Case 2
                        XArt2.Text = rstArticulo!Codigo
                        XDesArt2.Caption = rstArticulo!Descripcion
                        XArt2.SetFocus
                        rstArticulo.Close
                        
                    Case 3
                        XArti.Text = rstArticulo!Codigo
                        XDesArti.Caption = rstArticulo!Descripcion
                        XArti.SetFocus
                        rstArticulo.Close
                    
                    Case Else
                        WArticulo.Text = rstArticulo!Codigo
                        WDescripcion.Caption = rstArticulo!Descripcion
                    
                        WVector.Col = 1
                        WVector.Text = rstArticulo!Codigo
                        WVector.Col = 2
                        WVector.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    
                        Call Alta_Vector
                        Rem ????? anterior
                        Rem ????? anterior
                        Rem ????? anterior
                        Rem ????? anterior
                        WLinea.Text = WAnterior
                        Rem If Val(WLinea.Text) > 0 Then
                        Rem     WVector.Row = Val(WLinea.Text)
                        Rem End If
                    
                        WCantidad.SetFocus
                        
                End Select
                    
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            Opcion.Visible = False
            
        Case 2
            Indice = Pantalla.ListIndex
            ZClaveSolic = Mid$(WIndice.List(Indice), 1, 8)
            ZZEmpresa = Mid$(WIndice.List(Indice), 9, 6)
            
            XEmpresa = Wempresa
            Select Case Val(ZZEmpresa)
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
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Solic"
            ZSql = ZSql + " Where Solic.Clave = " + "'" + ZClaveSolic + "'"
            spSolic = ZSql
            Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            If rstSolic.RecordCount > 0 Then
                WArticulo.Text = rstSolic!Articulo
                WCantidad.Text = rstSolic!Cantidad
                Solicitud1.Text = Str$(Val(Left$(ZClaveSolic, 6)))
                WFecha1.Text = rstSolic!Entrega
                WFecha2.Text = rstSolic!Entrega
                rstSolic.Close
            End If
                            
            Call Conecta_Empresa
                
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
            If TipoOrden.ListIndex <> 2 Then
                Call Calcula_Precio(Proveedor.Text, WArticulo.Text, Precio, Condicion, XMoneda)
                If Precio = 0 Then
                    Desdelugar = 1
                    XCoti.Height = 4200
                    XCoti.Left = 2360
                    XCoti.Top = 1320
                    XCoti.Width = 7455
                    XCoti.Visible = True
                    XProve.Text = Proveedor.Text
                    XDesProve.Caption = DesProveedor.Caption
                    XArti = WArticulo.Text
                    XPrec.Text = ""
                    XCondicion.Text = ""
                    XObservaciones.Text = ""
                    XPrec.SetFocus
                        Else
                    If Moneda.ListIndex = 3 Then
                        Moneda.ListIndex = XMoneda
                    End If
                    If Moneda.ListIndex = XMoneda Then
                        WPrecio.Caption = Pusing("###,###.###", Str$(Precio))
                        WCondicion.Caption = Condicion
                        WCantidad.SetFocus
                            Else
                        m$ = "La moneda de la cotizacion no se corresponde a la moneda de los otros productos cargados en esta orden de compra"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        WArticulo.SetFocus
                    End If
                End If
                    Else
                WPrecio.Caption = ""
                WPrecio.Caption = Pusing("###,###.###", Str$(Precio))
                WCondicion.Caption = ""
                WCantidad.SetFocus
            End If

        Case Else
    End Select
    
    Rem Ayuda.Visible = False
    Rem Pantalla.Visible = False
    
End Sub


Private Sub Form_Load()

    Call Limpia_Vector

    TipoImpo.Clear
    
    TipoImpo.AddItem ""
    TipoImpo.AddItem "Maritimo"
    TipoImpo.AddItem "Terrestre"
    TipoImpo.AddItem "Areo"
    
    TipoImpo.ListIndex = 0
 
    Tarjeta.Clear
    
    Tarjeta.AddItem ""
    Tarjeta.AddItem "PymeNacion"
    
    Tarjeta.ListIndex = 0
 
    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    WBultos.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""

    Orden.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    ZEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    Flete.Text = ""
    Solicitud1.Text = ""
    Solicitud2.Text = ""
    Solicitud3.Text = ""
    Cuotas.Text = ""
    MesCuota.Text = ""
    AnoCuota.Text = ""
    DJai.Text = ""
    FechaDJai.Text = "  /  /    "
    
    TipoImpo.ListIndex = 0
    
    Moneda1.Clear
    
    Moneda1.AddItem "Dolares"
    Moneda1.AddItem "Pesos"
    Moneda1.AddItem "Euros"

    Moneda1.ListIndex = 0
    
    Leyenda.Clear
    
    Leyenda.AddItem ""
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CIF"
    Leyenda.AddItem "CFR"
    Leyenda.AddItem "CPT"
    Leyenda.AddItem "EXW"
    Leyenda.AddItem "FCA"
    
    Leyenda.ListIndex = 0
 
    Moneda2.Clear
    
    Moneda2.AddItem "Dolares"
    Moneda2.AddItem "Pesos"
    Moneda2.AddItem "Euros"

    Moneda2.ListIndex = 0
    
    Moneda3.Clear
    
    Moneda3.AddItem "Dolares"
    Moneda3.AddItem "Pesos"
    Moneda3.AddItem "Euros"

    Moneda3.ListIndex = 0
    
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Euros"
    Moneda.AddItem ""

    Moneda.ListIndex = 0
    
    TipoOrden.Clear
    
    TipoOrden.AddItem "Local"
    TipoOrden.AddItem "Importacion"
    TipoOrden.AddItem "Prestamo"
    TipoOrden.AddItem "Envases y Filtros"
    TipoOrden.AddItem "Drogas Lab."

    TipoOrden.ListIndex = 0
    
    TipoPago.Clear
    
    TipoPago.AddItem ""
    TipoPago.AddItem "Pago Anticipado"
    TipoPago.AddItem "A la vista"
    TipoPago.AddItem "Cuenta Corriente"

    TipoPago.ListIndex = 0
 
    Orden.Text = "1"
    
    Rem spOrden = "ListaOrdenNUmero"
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstOrden.RecordCount > 0 Then
    Rem     With rstOrden
    Rem         .MoveLast
    Rem         Orden.Text = rstOrden!Orden + 1
    Rem     End With
    Rem     rstOrden.Close
    Rem         Else
    Rem     Orden.Text = "1"
    Rem End If
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden < 800000"
    ZSql = ZSql + " Order by Orden.Clave"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveLast
            Orden.Text = Str$(rstOrden!Orden + 1)
        End With
        rstOrden.Close
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgOrden.Caption = "Ingreso de Orden de Compras :  " + !Nombre
        End If
    End With
 
    WVector.Col = 1
    WVector.Row = 1
    
    Graba.Enabled = True
    
    Rem Orden.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Graba.Enabled = True
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Vector
    Erase ZZArtiCerti
    
    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
            
                WVector.Row = Renglon
                
                WVector.Col = 1
                WVector.Text = rstOrden!Articulo
                Auxi1 = rstOrden!Articulo
                
                WVector.Col = 3
                WVector.Text = Pusing("###,###.##", rstOrden!Cantidad)
                
                WVector.Col = 4
                WVector.Text = Pusing("###,###.###", rstOrden!Precio)
                
                WVector.Col = 5
                WVector.Text = rstOrden!Fecha1
                
                WVector.Col = 6
                WVector.Text = rstOrden!fecha2
                
                WVector.Col = 7
                WVector.Text = rstOrden!Condicion
                
                WVector.Col = 8
                WVector.Text = IIf(IsNull(rstOrden!Bultos), "0", rstOrden!Bultos)
                
                WLugar = WVector.Row
                        
                XPorceDerechos(WLugar) = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                
                XSolicitud(WLugar, 1) = IIf(IsNull(rstOrden!Solicitud1), "0", rstOrden!Solicitud1)
                XSolicitud(WLugar, 2) = IIf(IsNull(rstOrden!Solicitud2), "0", rstOrden!Solicitud2)
                XSolicitud(WLugar, 3) = IIf(IsNull(rstOrden!Solicitud3), "0", rstOrden!Solicitud3)
                
                ZZArtiCerti(WLugar) = rstOrden!Articulo
                
                If rstOrden!Recibida > 0 Then
                    Graba.Enabled = False
                End If
                
                Rem ZMarca = IIf(IsNull(rstOrden!Marca), "", rstOrden!Marca)
                Rem If ZMarca = "X" Then
                Rem     Graba.Enabled = False
                Rem End If
                
                Vector(Renglon, 1) = Auxi1
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    
    If Graba.Enabled = False Then
        m$ = "La orden de compra no podra ser actualizada ya que posee productos que fueron cumplidos en forma total o parcial, o cargados datos adicionales referentes a la importacion"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        Auxi1 = Vector(Renglon, 1)
    
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector.TextMatrix(Renglon, 2) = rstArticulo!Descripcion
            WArticulo.SetFocus
            rstArticulo.Close
        End If
    
    Next Da
    
    WArticulo.SetFocus

    WVector.Col = 1
    WVector.Row = 1
    
    Call WVector_Click

End Sub

Private Sub Alta_Vector()


    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            WVector.Row = Renglon
            WAnterior = WVector.Row
            
            WVector.Col = 1
            WVector.Text = WArticulo.Text
            
            WVector.Col = 2
            WVector.Text = WDescripcion.Caption
                
            WVector.Col = 3
            WVector.Text = Pusing("###,###.##", WCantidad.Text)
                
            WVector.Col = 4
            WVector.Text = Pusing("###,###.###", WPrecio.Caption)
            
            WVector.Col = 5
            WVector.Text = WFecha1.Text
            
            WVector.Col = 6
            WVector.Text = WFecha2.Text
            
            WVector.Col = 7
            WVector.Text = WCondicion.Caption
            
            WVector.Col = 8
            WVector.Text = WBultos.Text
            
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                rstArticulo.Close
            End If
    
            Select Case ZZTipoMp
                Case 1
                    Anterior = "N"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " Order by Clave"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If rstOrden!Orden <> Val(Orden.Text) Then
                            Anterior = "S"
                        End If
                        rstOrden.Close
                    End If
                    
                    If Proveedor.Text = "10071408390" And Anterior = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Proveedor = " + "'" + "10059462341" + "'"
                        ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " Order by Clave"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            If rstOrden!Orden <> Val(Orden.Text) Then
                                Anterior = "S"
                            End If
                            rstOrden.Close
                        End If
                    End If
                    
                    If Anterior = "N" Then
                        m$ = "ATENCION !!! " + Chr$(13) + "SE DEBE SOLICITAR LAS ESPECIFICACIONES PARA QUE LABORATORIO VERIFIQUE LOS VALORES"
                        a% = MsgBox(m$, 48, "ORDENES DE COMPRA")
                    End If
        
                Case 2
                    Anterior = "N"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " Order by Clave"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If rstOrden!Orden <> Val(Orden.Text) Then
                            Anterior = "S"
                        End If
                        rstOrden.Close
                    End If
                    If Proveedor.Text = "10071408390" And Anterior = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Proveedor = " + "'" + "10059462341" + "'"
                        ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " Order by Clave"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            If rstOrden!Orden <> Val(Orden.Text) Then
                                Anterior = "S"
                            End If
                            rstOrden.Close
                        End If
                    End If
                    If Anterior = "N" Then
                        m$ = "ATENCION !!! " + Chr$(13) + "SE DEBE SOLICITAR UNA MUESTRA PREVIA PARA QUE LABORATORIO VERIFIQUE LOS VALORES DE LAS ESPECIFICACIONES DEL PRODUCTO"
                        a% = MsgBox(m$, 48, "ORDENES DE COMPRA")
                    End If

                Case Else
            End Select
            
            
            WLugar = WVector.Row
            
            XPorceDerechos(WLugar) = WPorceDerechos.Text
            
            XSolicitud(WLugar, 1) = Solicitud1.Text
            XSolicitud(WLugar, 2) = Solicitud2.Text
            XSolicitud(WLugar, 3) = Solicitud3.Text
            
            WVector.Col = 1
            
                Else
                
            WVector.Row = Val(WLinea.Text)
            WAnterior = WVector.Row
            
            WVector.Col = 1
            WVector.Text = WArticulo.Text
            
            WVector.Col = 2
            WVector.Text = WDescripcion.Caption
                
            WVector.Col = 3
            WVector.Text = Pusing("###,###.##", WCantidad.Text)
            
            WVector.Col = 4
            WVector.Text = Pusing("###,###.###", WPrecio.Caption)
            
            WVector.Col = 5
            WVector.Text = WFecha1.Text
            
            WVector.Col = 6
            WVector.Text = WFecha2.Text
                        
            WVector.Col = 7
            WVector.Text = WCondicion.Caption
                        
            WVector.Col = 8
            WVector.Text = WBultos.Text
            
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                rstArticulo.Close
            End If
    
            Select Case ZZTipoMp
                Case 1
                    Anterior = "N"
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " Order by Clave"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If rstOrden!Orden <> Val(Orden.Text) Then
                            Anterior = "S"
                        End If
                        rstOrden.Close
                    End If
                    
                    If Proveedor.Text = "10071408390" And Anterior = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Proveedor = " + "'" + "10059462341" + "'"
                        ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " Order by Clave"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            If rstOrden!Orden <> Val(Orden.Text) Then
                                Anterior = "S"
                            End If
                            rstOrden.Close
                        End If
                    End If
                    
                    If Anterior = "N" Then
                        m$ = "ATENCION !!! " + Chr$(13) + "SE DEBE SOLICITAR LAS ESPECIFICACIONES PARA QUE LABORATORIO VERIFIQUE LOS VALORES"
                        a% = MsgBox(m$, 48, "ORDENES DE COMPRA")
                    End If
        
                Case 2
                    Anterior = "N"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " Order by Clave"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If rstOrden!Orden <> Val(Orden.Text) Then
                            Anterior = "S"
                        End If
                        rstOrden.Close
                    End If
                    
                    If Proveedor.Text = "10071408390" And Anterior = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden.Proveedor = " + "'" + "10059462341" + "'"
                        ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " Order by Clave"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            If rstOrden!Orden <> Val(Orden.Text) Then
                                Anterior = "S"
                            End If
                            rstOrden.Close
                        End If
                    End If
                    
                    If Anterior = "N" Then
                        m$ = "ATENCION !!! " + Chr$(13) + "SE DEBE SOLICITAR UNA MUESTRA PREVIA PARA QUE LABORATORIO VERIFIQUE LOS VALORES DE LAS ESPECIFICACIONES DEL PRODUCTO"
                        a% = MsgBox(m$, 48, "ORDENES DE COMPRA")
                    End If

                Case Else
            End Select
            
            WLugar = WVector.Row
            
            XPorceDerechos(WLugar) = WPorceDerechos.Text
            
            XSolicitud(WLugar, 1) = Solicitud1.Text
            XSolicitud(WLugar, 2) = Solicitud2.Text
            XSolicitud(WLugar, 3) = Solicitud3.Text
            
            WVector.Col = 1
            
    End If

End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Pantalla.Visible = False
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        Entra = "N"
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            Origen.Text = IIf(IsNull(rstOrden!Origen), "", rstOrden!Origen)
            Carpeta.Text = IIf(IsNull(rstOrden!Carpeta), "", rstOrden!Carpeta)
            Moneda.ListIndex = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            Tarjeta.ListIndex = IIf(IsNull(rstOrden!Tarjeta), "0", rstOrden!Tarjeta)
            Rem by nan quito rem a tipoOrden 5-3-2014
            TipoOrden.ListIndex = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            TipoPago.ListIndex = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
            Leyenda.ListIndex = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
            PedidoImpo.Text = IIf(IsNull(rstOrden!PedidoImpo), "", rstOrden!PedidoImpo)
            FechaImpo.Text = IIf(IsNull(rstOrden!FechaImpo), "  /  /    ", rstOrden!FechaImpo)
            TipoImpo.ListIndex = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
            Flete.Text = IIf(IsNull(rstOrden!Flete), "", rstOrden!Flete)
            Fecha.Text = rstOrden!Fecha
            Proveedor.Text = rstOrden!Proveedor
            Cuotas.Text = IIf(IsNull(rstOrden!Cuotas), "", rstOrden!Cuotas)
            MesCuota.Text = IIf(IsNull(rstOrden!MesCuota), "", rstOrden!MesCuota)
            AnoCuota.Text = IIf(IsNull(rstOrden!AnoCuota), "", rstOrden!AnoCuota)
            DJai.Text = IIf(IsNull(rstOrden!DJai), "", rstOrden!DJai)
            DJai.Text = Trim(DJai.Text)
            FechaDJai.Text = IIf(IsNull(rstOrden!FechaDJai), "  /  /    ", rstOrden!FechaDJai)
            rstOrden.Close
            Entra = "S"
                Else
            WOrden = Orden.Text
            WTipoOrden = TipoOrden.ListIndex
            Call Limpia_Click
            Orden.Text = WOrden
            TipoOrden.ListIndex = WTipoOrden
            Fecha.SetFocus
        End If
        
        If Entra = "S" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Proveedor.Text = rstProveedor!Proveedor
                DesProveedor.Caption = rstProveedor!Nombre
                ZEmail = rstProveedor!EMail
                rstProveedor.Close
            End If
             Call Proceso_Click
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Proveedor.Text = rstProveedor!Proveedor
                DesProveedor.Caption = rstProveedor!Nombre
                ZEmail = rstProveedor!EMail
                rstProveedor.Close
                If TipoOrden.ListIndex = 1 Then
                    Carpeta.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
                Call Verifica_Atrasos
                    Else
                Proveedor.SetFocus
            End If
                Else
            TipoConsulta = "1"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
End Sub

Private Sub Carpeta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DatosImpo.Visible = True
        Origen.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Leyenda.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Leyenda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PedidoImpo.SetFocus
    End If
End Sub

Private Sub PedidoImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaImpo.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaImpo.Text, Auxi)
        If Auxi = "S" Then
            TipoImpo.SetFocus
                Else
            FechaImpo.SetFocus
        End If
    End If
End Sub

Private Sub TipoImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoPago.SetFocus
    End If
End Sub

Private Sub TipoPago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Flete.SetFocus
    End If
End Sub

Private Sub Flete_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DJai.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DJai_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaDJai.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaDjai_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        DatosImpo.Visible = False
        
        ZOrigen = Origen.Text
        ZLeyenda = Str$(Leyenda.ListIndex)
        ZPedidoImpo = PedidoImpo.Text
        ZFechaImpo = FechaImpo.Text
        ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
        ZTipoImpo = Str$(TipoImpo.ListIndex)
        ZTipoPago = Str$(TipoPago.ListIndex)
        ZFlete = Flete.Text
        ZDJai = DJai.Text
        ZFechaDJai = FechaDJai.Text
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Orden SET "
        ZSql = ZSql + " DJai = " + "'" + ZDJai + "',"
        ZSql = ZSql + " FechaDJai = " + "'" + ZFechaDJai + "',"
        ZSql = ZSql + " Flete = " + "'" + ZFlete + "',"
        ZSql = ZSql + " Origen = " + "'" + ZOrigen + "',"
        ZSql = ZSql + " Leyenda = " + "'" + ZLeyenda + "',"
        ZSql = ZSql + " PedidoImpo = " + "'" + ZPedidoImpo + "',"
        ZSql = ZSql + " FechaImpo = " + "'" + ZFechaImpo + "',"
        ZSql = ZSql + " OrdFechaImpo = " + "'" + ZOrdFechaImpo + "',"
        ZSql = ZSql + " TipoImpo = " + "'" + ZTipoImpo + "',"
        ZSql = ZSql + " TipoPago = " + "'" + ZTipoPago + "'"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        WArticulo.SetFocus
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Solicitud1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Solicitud1.Text) <> 0 Then
            ZSolicitud = Solicitud1.Text
            Call Valida_Solicitud
            If ZEntra = "S" Then
                Solicitud2.SetFocus
            End If
                Else
            CargaSolicitud.Visible = False
            WFecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Solicitud1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Solicitud2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Solicitud2.Text) <> 0 Then
            ZSolicitud = Solicitud2.Text
            Call Valida_Solicitud
            If ZEntra = "S" Then
                Solicitud3.SetFocus
            End If
                Else
            CargaSolicitud.Visible = False
            WFecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Solicitud2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Solicitud3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Solicitud3.Text) <> 0 Then
            ZSolicitud = Solicitud3.Text
            Call Valida_Solicitud
            If ZEntra = "S" Then
                CargaSolicitud.Visible = False
                WFecha1.SetFocus
            End If
                Else
            CargaSolicitud.Visible = False
            WFecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Solicitud3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valida_Solicitud()

    XEmpresa = Wempresa
    ZEntra = "N"
    
    For Cicla = 1 To 11
    
        Select Case Cicla
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
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Solic"
        ZSql = ZSql + " Where Solic.Solicitud = " + "'" + ZSolicitud + "'"
        ZSql = ZSql + " and Solic.Articulo = " + "'" + WArticulo.Text + "'"
        spSolic = ZSql
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            rstSolic.Close
            ZEntra = "S"
            Exit For
        End If
        
    Next Cicla
    
    Call Conecta_Empresa
    
End Sub

Private Sub Carpeta_DblClick()
    DatosImpo.Visible = True
    Origen.SetFocus
End Sub

Sub Calcula_Precio(WProveedor, WArticulo As String, WPrecio As Double, WCondicion As String, WMoneda As Integer)

    WPrecio = 0
    WCondicion = ""
    WFecha = ""
    
    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spCotiza = "ListaCotizaProveedor " + "'" + WProveedor + "'"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstCotiza.RecordCount > 0 Then
    With rstCotiza
        .MoveFirst
        Do
            If .EOF = False Then
            
                If WArticulo = rstCotiza!Articulo Then
                    
                    If rstCotiza!FechaOrd > WFecha Then
                        WPrecio = rstCotiza!Precio
                        WCondicion = rstCotiza!Condicion
                        WCotiza = rstCotiza!Cotiza
                        WFecha = rstCotiza!FechaOrd
                        WMoneda = rstCotiza!Moneda
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
    
    a = 1

End Sub


Sub Impresion()

        Rem Open "lpt1" For Output As #1
        Open "dada.txt" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = Wempresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        For Ci = 1 To 1

        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        If TipoOrden.ListIndex = 1 Then
            Print #1, Tab(30); "(IMPORTACION)";
        End If
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        WCategoriaI = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
            WCategoriaI = ""
            If ZCategoriaI = 1 Then
                WCategoriaI = "A"
                    Else
                If ZCategoriaI = 2 Then
                    WCategoriaI = "B"
                        Else
                    If ZCategoriaI = 3 Then
                        WCategoriaI = "C"
                            Else
                        If ZCategoriaI = 4 Then
                            WCategoriaI = "E"
                        End If
                    End If
                End If
            End If
            rstProveedor.Close
        End If
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 33); " Categ.:"; Trim(WCategoriaI);
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 And Val(Carpeta.Text) <> 999999 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        
        If Leyenda.ListIndex <> 0 Then
            Print #1, Tab(50); "("; Leyenda.Text; ")":
        End If
        
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion  |Canti.|   Precio  |1ra Fec.  |Ul.Fecha  |Cond. Pago|"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        
        For a = 1 To 99
        
            WVector.Row = a
                
            WVector.Col = 1
            Articulo = UCase(WVector.Text)
                
            If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
            
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    WVector.Col = 2
                    WDescripcion = WVector.Text
                End If
                    
                WVector.Col = 3
                Cantidad = Val(WVector.Text)
                
                WVector.Col = 4
                Precio = Val(WVector.Text)
                
                WVector.Col = 5
                Fecha1 = WVector.Text
                
                WVector.Col = 6
                fecha2 = WVector.Text
                
                WVector.Col = 7
                Condicion = WVector.Text
                
                WVector.Col = 8
                Bultos = WVector.Text

                WCantidad = WCantidad + 1

                Print #1, Tab(1); "|"; Articulo;
                Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                Print #1, Tab(28); "|"; Alinea("##,###", Str$(Cantidad));
                Select Case Moneda.ListIndex
                    Case 0
                        Print #1, Tab(35); "|U$S"; Alinea("####.###", Str$(Precio));
                    Case 2
                        Print #1, Tab(35); "| € "; Alinea("####.###", Str$(Precio));
                    Case Else
                        Print #1, Tab(35); "| $ "; Alinea("####.###", Str$(Precio));
                End Select
                Print #1, Tab(47); "|"; Fecha1;
                Print #1, Tab(58); "|"; fecha2;
                Print #1, Tab(69); "|"; Left$(Condicion, 10);
                Print #1, Tab(80); "|"

                Valor = Valor + (Cantidad * Precio)

            End If
        
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |               |      |           |          |          |          |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Select Case Moneda.ListIndex
            Case 0
                Print #1, "|          Valor total de la orden : U$S "; Alinea("#####.##", Str$(Valor));
            Case 2
                Print #1, "|          Valor total de la orden :   € "; Alinea("#####.##", Str$(Valor));
            Case Else
                Print #1, "|          Valor total de la orden :   $ "; Alinea("#####.##", Str$(Valor));
        End Select
        Print #1, Tab(80); "|"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Next Ci

ImpreCopia:

        WCantidad = 0
        
        ' Copia 2

        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, "--------------------------------------------------------------------------------"
        
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Impretit;
        Rem Print #1, Tab(80); "|"
        
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Tab(60); "Remito :..........";
        Rem Print #1, Tab(80); "|"
        
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Tab(5); "Orden.....: ";
        Rem Print #1, Tab(20); Alinea("######", Orden.Text);
        Rem If TipoOrden.ListIndex = 1 Then
        Rem     Print #1, Tab(30); "(IMPORTACION)";
        Rem End If
        Rem Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Rem Print #1, Tab(80); "|"
        
        Rem WCategoriaI = ""
        Rem ZSql = ""
        Rem ZSql = ZSql + "Select *"
        Rem ZSql = ZSql + " FROM Proveedor"
        Rem ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        Rem spProveedor = ZSql
        Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        Rem If RstProveedor.RecordCount > 0 Then
        Rem     ZCategoriaI = IIf(IsNull(RstProveedor!CategoriaI), "0", RstProveedor!CategoriaI)
        Rem     WCategoriaI = ""
        Rem     If ZCategoriaI = 1 Then
        Rem         WCategoriaI = "A"
        Rem             Else
        Rem         If ZCategoriaI = 2 Then
        Rem             WCategoriaI = "B"
        Rem                 Else
        Rem             If ZCategoriaI = 3 Then
        Rem                 WCategoriaI = "C"
        Rem                     Else
        Rem                 If ZCategoriaI = 4 Then
        Rem                     WCategoriaI = "E"
        Rem                 End If
        Rem             End If
        Rem         End If
        Rem     End If
        Rem     RstProveedor.Close
        Rem End If
        
        
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Rem Print #1, Tab(35); Left$(DesProveedor.Caption, 20); " ("; WCategoriaI; ")";
        Rem Print #1, Tab(62); "Informe :.....";
        Rem Print #1, Tab(80); "|"
        
        Rem If Val(Carpeta.Text) <> 0 And Val(Carpeta.Text) <> 999999 Then
        Rem     Print #1, Tab(1); "|";
        Rem     Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
        Rem         Else
        Rem     Print #1, Tab(1); "|";
        Rem End If
        Rem Print #1, Tab(80); "|"
        
        Rem Print #1, "--------------------------------------------------------------------------------"
        Rem Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
        Rem Print #1, "--------------------------------------------------------------------------------"

        Rem Cantidad = 0
        Rem Valor = 0
        
        Rem For A = 0 To 9
        Rem
        Rem     Suma = A * 10
        Rem     DBGrid1.FirstRow = Suma
        Rem
        Rem     For iRow = 0 To 9
        Rem
        Rem         WRow = iRow
        Rem         DBGrid1.Row = WRow
        Rem
        Rem         DBGrid1.Col = 0
        Rem         Articulo = UCase(DBGrid1.Text)
        Rem
        Rem         If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
        Rem
        Rem             XProveedor = Proveedor.Text
        Rem             Call Ceros(XProveedor, 11)
        Rem             ClaveMarcas = Articulo + XProveedor
        Rem             spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
        Rem             Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstMarcas.RecordCount > 0 Then
        Rem                 WDescripcion = rstMarcas!Descripcion
        Rem                 rstMarcas.Close
        Rem                     Else
        Rem                 DBGrid1.Col = 1
        Rem                 WDescripcion = DBGrid1.Text
        Rem             End If
        Rem
        Rem             DBGrid1.Col = 2
        Rem             Cantidad = Val(DBGrid1.Text)
        Rem
        Rem             DBGrid1.Col = 3
        Rem             Precio = Val(DBGrid1.Text)
        Rem
        Rem             DBGrid1.Col = 4
        Rem             Fecha1 = DBGrid1.Text
        Rem
        Rem             DBGrid1.Col = 5
        Rem             fecha2 = DBGrid1.Text
        Rem
        Rem             DBGrid1.Col = 6
        Rem             Condicion = DBGrid1.Text
        Rem
        Rem             WUbicacion = ""
        Rem
        Rem             spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Rem             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstArticulo.RecordCount > 0 Then
        Rem                 WUbicacion = rstArticulo!Deposito
        Rem                 rstArticulo.Close
        Rem             End If
        Rem
        Rem             WCantidad = WCantidad + 2
        Rem
        Rem             Print #1, Tab(1); "|"; Articulo;
        Rem             Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
        Rem             Print #1, Tab(41); "|"; Alinea("###,###", Str$(Cantidad));
        Rem             Print #1, Tab(50); "|"; Fecha1;
        Rem             Print #1, Tab(61); "|"; fecha2;
        Rem             Print #1, Tab(72); "|";
        Rem             Print #1, Tab(80); "|"
        Rem
        Rem             Print #1, Tab(1); "|";
        Rem             Print #1, Tab(12); "|"; WUbicacion;
        Rem             Print #1, Tab(50); "|";
        Rem             Print #1, Tab(61); "|";
        Rem             Print #1, Tab(72); "|";
        Rem             Print #1, Tab(80); "|"
        Rem
        Rem         End If
        Rem
        Rem     Next iRow
        Rem Next A

        Rem For Ciclo = WCantidad To 15
        Rem     Print #1, "|          |                            |        |          |          |       |"
        Rem Next Ciclo

        Rem Print #1, "--------------------------------------------------------------------------------"
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""
        Rem Print #1, ""

    Close #1

End Sub

Private Sub Primera_Click()
        
        Rem by nan
        Call Impresion_Crystal
        Exit Sub
        
        Rem Open "lpt1" For Output As #1
        Open "dada.txt" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = Wempresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        If TipoOrden.ListIndex = 1 Then
            Print #1, Tab(30); "(IMPORTACION)";
        End If
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        WCategoriaI = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
            WCategoriaI = ""
            If ZCategoriaI = 1 Then
                WCategoriaI = "A"
                    Else
                If ZCategoriaI = 2 Then
                    WCategoriaI = "B"
                        Else
                    If ZCategoriaI = 3 Then
                        WCategoriaI = "C"
                            Else
                        If ZCategoriaI = 4 Then
                            WCategoriaI = "E"
                        End If
                    End If
                End If
            End If
            rstProveedor.Close
        End If
        
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 33); " Categ.:"; Trim(WCategoriaI);
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 And Val(Carpeta.Text) <> 999999 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        
        If Leyenda.ListIndex <> 0 Then
            Print #1, Tab(50); "("; Leyenda.Text; ")";
        End If
        
        Print #1, Tab(80); "|"
        
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion  |Canti.|   Precio  |1ra Fec.  |Ul.Fecha  |Cond. Pago|"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        
        For a = 1 To 99
        
            WVector.Row = a
                
            WVector.Col = 1
            Articulo = UCase(WVector.Text)
                
            If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
            
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    WVector.Col = 2
                    WDescripcion = WVector.Text
                End If
                
                WVector.Col = 3
                Cantidad = Val(WVector.Text)
                
                WVector.Col = 4
                Precio = Val(WVector.Text)
                
                WVector.Col = 5
                Fecha1 = WVector.Text
                
                WVector.Col = 6
                fecha2 = WVector.Text
                
                WVector.Col = 7
                Condicion = WVector.Text
                
                WVector.Col = 8
                Bultos = WVector.Text

                WCantidad = WCantidad + 1

                Print #1, Tab(1); "|"; Articulo;
                Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                Print #1, Tab(28); "|"; Alinea("##,###", Str$(Cantidad));
                Select Case Moneda.ListIndex
                    Case 0
                        Print #1, Tab(35); "|U$S"; Alinea("#,###.##", Str$(Precio));
                    Case 2
                        Print #1, Tab(35); "| € "; Alinea("#,###.##", Str$(Precio));
                    Case Else
                        Print #1, Tab(35); "| $ "; Alinea("#,###.##", Str$(Precio));
                End Select
                Print #1, Tab(47); "|"; Fecha1;
                Print #1, Tab(58); "|"; fecha2;
                Print #1, Tab(69); "|"; Left$(Condicion, 10);
                Print #1, Tab(80); "|"

                Valor = Valor + (Cantidad * Precio)

            End If
        
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |               |      |           |          |          |          |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Select Case Moneda.ListIndex
            Case 0
                Print #1, "|          Valor total de la orden : U$S "; Alinea("#####.##", Str$(Valor));
            Case 2
                Print #1, "|          Valor total de la orden :   € "; Alinea("#####.##", Str$(Valor));
            Case Else
                Print #1, "|          Valor total de la orden :   $ "; Alinea("#####.##", Str$(Valor));
        End Select
        Print #1, Tab(80); "|"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Close #1


 End Sub

Private Sub Tercera_Click()

        Rem by nan
        Open "lpt1" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = Wempresa
           .Seek "=", Claveven$
           If .NoMatch = False Then
               Impretit = !Nombre
                   Else
               Impretit = ""
           End If
        End With
    
        WCantidad = 0
        
        ' Copia 2

        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, "Empresa : "; Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(60); "Remito :..........";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        WCategoriaI = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
            WCategoriaI = ""
            If ZCategoriaI = 1 Then
                WCategoriaI = "A"
                    Else
                If ZCategoriaI = 2 Then
                    WCategoriaI = "B"
                        Else
                    If ZCategoriaI = 3 Then
                        WCategoriaI = "C"
                            Else
                        If ZCategoriaI = 4 Then
                            WCategoriaI = "E"
                        End If
                    End If
                End If
            End If
            rstProveedor.Close
        End If
        
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 20); " ("; WCategoriaI; ")";
        Print #1, Tab(62); "Informe :.....";
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 And Val(Carpeta.Text) <> 999999 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
        Print #1, "--------------------------------------------------------------------------------"

        Cantidad = 0
        Valor = 0
        
        For a = 1 To 99
        
            WVector.Row = a
                
            WVector.Col = 1
            Articulo = UCase(WVector.Text)
            
            If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
            
                XProveedor = Proveedor.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Articulo + XProveedor
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    WDescripcion = rstMarcas!Descripcion
                    rstMarcas.Close
                        Else
                    WVector.Col = 2
                    WDescripcion = WVector.Text
                End If
                
                WVector.Col = 3
                Cantidad = Val(WVector.Text)
                
                WVector.Col = 4
                Precio = Val(WVector.Text)
                
                WVector.Col = 5
                Fecha1 = WVector.Text
                
                WVector.Col = 6
                fecha2 = WVector.Text
                
                WVector.Col = 7
                Condicion = WVector.Text
                
                WVector.Col = 8
                Bultos = WVector.Text
            
                WUbicacion = ""
            
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WUbicacion = rstArticulo!Deposito
                    rstArticulo.Close
                End If

                WCantidad = WCantidad + 2

                Print #1, Tab(1); "|"; Articulo;
                Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
                Print #1, Tab(41); "|"; Alinea("###,###", Str$(Cantidad));
                Print #1, Tab(50); "|"; Fecha1;
                Print #1, Tab(61); "|"; fecha2;
                Print #1, Tab(72); "|";
                Print #1, Tab(80); "|"
                    
                Print #1, Tab(1); "|";
                Print #1, Tab(12); "|"; WUbicacion;
                Print #1, Tab(50); "|";
                Print #1, Tab(61); "|";
                Print #1, Tab(72); "|";
                Print #1, Tab(80); "|"

            End If
                                        
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |                            |        |          |          |       |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

    Close #1
    
    WVector.Col = 1
    WVector.Row = 1
    Call WVector_Click
    
    Orden.SetFocus

 End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
        
            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstProveedor.RecordCount > 0 Then
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        Da = Len(rstProveedor!Nombre) - WEspacios
                
                        For aa = 1 To Da
                            If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                                Auxi = Str$(rstProveedor!Proveedor)
                                Call Ceros(Auxi, 11)
                                IngresaItem = Auxi + "    " + rstProveedor!Nombre
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Proveedor
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
            
            Call Conecta_Empresa
            
            
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
    
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            Da = Len(rstArticulo!Descripcion) - WEspacios
                
                            For AAa = 1 To Da
                                If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstArticulo!Descripcion), AAa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next AAa
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

Private Sub XAcepta_Click()

    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WCotiza = "1"
    
    spCotiza = "ListaCotizaNumero"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
        With rstCotiza
            .MoveLast
            WCotiza = rstCotiza!Cotiza + 1
        End With
        rstCotiza.Close
    End If

    Articulo = XArti.Text
    Precio = XPrec.Text
    Condicion = XCondicion.Text
    Observaciones = XObservaciones.Text
    XRenglon = 1
    
    Auxi = Str$(XRenglon)
    Call Ceros(Auxi, 2)
    Auxi1 = Str$(WCotiza)
    Call Ceros(Auxi1, 6)
                        
    WCot = Str$(WCotiza)
    WRenglon = Str$(XRenglon)
    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WProveedor = XProve.Text
    WArticulo = XArti.Text
    WPrecio = XPrec.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WCondicion = XCondicion.Text
    WObservaciones = XObservaciones.Text
    WClave = Auxi1 + Auxi
    WDate = Date$
    WMoneda = Str$(Moneda3.ListIndex)
        
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
    
    XCoti.Visible = False
    If Desdelugar = 0 Then
        Orden.SetFocus
            Else
        Rem WCantidad.Text = ""
        WCantidad.SetFocus
    End If
    
End Sub

Private Sub XAcepta1_Click()

    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Da = 0
    With rstLiscot
        .Index = "Clave"
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
    
    WAno = Right$(Date$, 4)
    WDia = Mid$(Date$, 4, 2)
    WMes = Left$(Date$, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            AA1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = rstCambios!Cambio
            ParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            rstCambios.Close
        End With
            Else
        Paridad = 1
        ParidadII = 1
    End If
    
    XParam = "'" + XProv1.Text + "','" _
            + XProv1.Text + "'"
    
    spCotiza = "ListaCotizaProveedorDesdeHasta" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    
    Pasa = 0
    Canti = 0
    
    If rstCotiza.RecordCount > 0 Then
            
        With rstCotiza
            .MoveFirst
    
            Do
            
                If .EOF = True Then
                    Exit Do
                End If

                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                Select Case Moneda1.ListIndex
                    Case 0
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio
                            Case 1
                                WPrecio = !Precio / Paridad
                            Case Else
                                WCoeParidad = ParidadII / Paridad
                                WPrecio = !Precio * WCoeParidad
                        End Select
                    Case 1
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio * Paridad
                            Case 1
                                WPrecio = !Precio
                            Case Else
                                WPrecio = !Precio * ParidadII
                        End Select
                    Case Else
                        Select Case !Moneda
                            Case 0
                                WCoeParidad = Paridad / ParidadII
                                WPrecio = !Precio * WCoeParidad
                            Case 1
                                WPrecio = !Precio / ParidadII
                            Case Else
                                WPrecio = !Precio
                        End Select
                End Select
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                
                        For Da = 1 To 9
                        
                            If XVector(Da, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = XVector(Da, 1)
                                !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                !Precio = Val(XVector(Da, 2))
                                !Condicion = XVector(Da, 3)
                                !Observaciones = XVector(Da, 4)
                                !Clave = !Proveedor + !Articulo
                                !Orden = 0
                                .Update
                            End If
                            
                        Next Da
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For Da = 1 To 2
                        XVector(Da, 1) = XVector(Da + 1, 1)
                        XVector(Da, 2) = XVector(Da + 1, 2)
                        XVector(Da, 3) = XVector(Da + 1, 3)
                        XVector(Da, 4) = XVector(Da + 1, 4)
                    Next Da
                    Canti = 3
                End If
                
                XVector(Canti, 1) = !Fecha
                XVector(Canti, 2) = Str$(WPrecio)
                XVector(Canti, 3) = !Condicion
                XVector(Canti, 4) = !Observaciones
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
        End With
    End If
    
    If Pasa <> 0 Then
        With rstLiscot
                
            For Da = 1 To 3
                    
                If XVector(Da, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = XVector(Da, 1)
                    !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    !Precio = Val(XVector(Da, 2))
                    !Condicion = XVector(Da, 3)
                    !Observaciones = XVector(Da, 4)
                    !Clave = !Proveedor + !Articulo
                    .Update
                End If
                
            Next Da
                        
        End With
    End If
    
    Da = 0
    With rstLiscot
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WDescriProveedor = ""
                WArticulo = !Articulo
                WDescriArticulo = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                
                    WDescriProveedor = rstProveedor!Nombre
                    
                    ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
                    ZCategoriaII = IIf(IsNull(rstProveedor!CategoriaII), "0", rstProveedor!CategoriaII)
                    
                    WCategoriaI = ""
                    WCategoriaII = ""
        
                    If ZCategoriaI = 1 Then
                        WCategoriaI = "A"
                            Else
                        If ZCategoriaI = 2 Then
                            WCategoriaI = "B"
                                Else
                            If ZCategoriaI = 3 Then
                                WCategoriaI = "C"
                                    Else
                                If ZCategoriaI = 4 Then
                                    WCategoriaI = "E"
                                End If
                            End If
                        End If
                    End If
                    
                    WCategoriaII = "S/C"
                    If ZCategoriaII = 1 Then
                        WCategoriaII = "Muy Bueno"
                            Else
                        If ZCategoriaII = 2 Then
                            WCategoriaII = "Bueno"
                                Else
                            If ZCategoriaII = 3 Then
                                WCategoriaII = "Regular"
                                    Else
                                If ZCategoriaII = 4 Then
                                    WCategoriaII = "Malo"
                                End If
                            End If
                        End If
                    End If
                    
                    If WCategoriaI <> "" And WCategoriaII <> "" Then
                        WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                    End If
                    
                    rstProveedor.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DescriProveedor = WDescriProveedor
                !DescriArticulo = WDescriArticulo
                
                Select Case Moneda1.ListIndex
                    Case 0
                        !Titulo = "(En Dolares)"
                    Case 2
                        !Titulo = "(En Euros)"
                    Case Else
                        !Titulo = "(En Pesos)"
                End Select
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cotizaciones por Proveedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Listcot.proveedor} in " + Chr$(34) + XProv1.Text + Chr$(34) + " to " + Chr$(34) + XProv1.Text + Chr$(34)
    
    Listado.ReportFileName = "WCotprv.rpt"
    Rem Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Destination = 0
    Listado.Action = 1
    
    Call Conecta_Empresa
    
End Sub

Private Sub XAcepta2_Click()

    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    XArt2.Text = UCase(XArt2.Text)

    Da = 0
    With rstLiscot
        .Index = "Clave"
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
    
    WAno = Right$(Date$, 4)
    WDia = Mid$(Date$, 4, 2)
    WMes = Left$(Date$, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            AA1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = rstCambios!Cambio
            ParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            rstCambios.Close
        End With
            Else
        Paridad = 1
        ParidadII = 1
    End If
    
    
    Pasa = 0
    Canti = 0
    
    XParam = "'" + XArt2.Text + "','" _
            + XArt2.Text + "'"
    
    spCotiza = "ListaCotizaArticuloDesdeHasta" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
            
    With rstCotiza
    
            .MoveFirst
            
            Do
            
                WCotiza = !Cotiza
                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                Select Case Moneda2.ListIndex
                    Case 0
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio
                            Case 1
                                WPrecio = !Precio / Paridad
                            Case Else
                                WCoeParidad = ParidadII / Paridad
                                WPrecio = !Precio * WCoeParidad
                        End Select
                    Case 1
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio * Paridad
                            Case 1
                                WPrecio = !Precio
                            Case Else
                                WPrecio = !Precio * ParidadII
                        End Select
                    Case Else
                        Select Case !Moneda
                            Case 0
                                WCoeParidad = Paridad / ParidadII
                                WPrecio = !Precio * WCoeParidad
                            Case 1
                                WPrecio = !Precio / ParidadII
                            Case Else
                                WPrecio = !Precio
                        End Select
                End Select
                
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                    
                        Rem If Val(XVector(3, 2)) <> 0 Then
                        Rem     WAuxi = Int(Val(XVector(3, 2)) * 100)
                        Rem             Else
                        Rem     If Val(XVector(2, 2)) <> 0 Then
                        Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
                        Rem             Else
                        Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
                        Rem     End If
                        Rem End If
                        Rem
                        Rem Call Ceros(WAuxi, 9)
                        
                        If XVector(3, 5) <> "" Then
                            WAuxi = XVector(3, 5)
                                    Else
                            If XVector(2, 5) <> "" Then
                                WAuxi = XVector(2, 5)
                                    Else
                                WAuxi = XVector(1, 5)
                            End If
                        End If
                        WAuxi = Str$(Val(WAuxi) - 90000000)
                        Call Ceros(WAuxi, 9)
                    
                        For Da = 1 To 9
                        
                            If XVector(Da, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = XVector(Da, 1)
                                !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                !Precio = Val(XVector(Da, 2))
                                !Condicion = XVector(Da, 3)
                                !Observaciones = XVector(Da, 4)
                                !Clave = !Proveedor + !Articulo
                                !Orden = WAuxi + !Proveedor
                                .Update
                            End If
                            
                        Next Da
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For Da = 1 To 2
                        XVector(Da, 1) = XVector(Da + 1, 1)
                        XVector(Da, 2) = XVector(Da + 1, 2)
                        XVector(Da, 3) = XVector(Da + 1, 3)
                        XVector(Da, 4) = XVector(Da + 1, 4)
                        XVector(Da, 5) = XVector(Da + 1, 5)
                    Next Da
                    Canti = 3
                End If
                
                XVector(Canti, 1) = !Fecha
                XVector(Canti, 2) = Str$(WPrecio)
                XVector(Canti, 3) = !Condicion
                XVector(Canti, 4) = !Observaciones
                XVector(Canti, 5) = !FechaOrd
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    rstCotiza.Close
    
    End If
    
    If Pasa <> 0 Then
        With rstLiscot
        
            Rem If Val(XVector(3, 2)) <> 0 Then
            Rem     WAuxi = Int(Val(XVector(3, 2)) * 100)
            Rem             Else
            Rem     If Val(XVector(2, 2)) <> 0 Then
            Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
            Rem             Else
            Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
            Rem     End If
            Rem End If
            Rem
            Rem Call Ceros(WAuxi, 9)
            
            If XVector(3, 5) <> "" Then
                WAuxi = XVector(3, 5)
                        Else
                If XVector(2, 5) <> "" Then
                    WAuxi = XVector(2, 5)
                        Else
                    WAuxi = XVector(1, 5)
                End If
            End If
            WAuxi = Str$(Val(WAuxi) - 90000000)
            Call Ceros(WAuxi, 9)
                
            For Da = 1 To 9
                    
                If XVector(Da, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = XVector(Da, 1)
                    !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    !Precio = Val(XVector(Da, 2))
                    !Condicion = XVector(Da, 3)
                    !Observaciones = XVector(Da, 4)
                    !Clave = !Proveedor + !Articulo
                    !Orden = WAuxi + !Proveedor
                    .Update
                End If
                
            Next Da
                        
        End With
    End If
    
    Da = 0
    With rstLiscot
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WDescriProveedor = ""
                WArticulo = !Articulo
                WDescriArticulo = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                
                    WDescriProveedor = rstProveedor!Nombre
                    
                    ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
                    ZCategoriaII = IIf(IsNull(rstProveedor!CategoriaII), "0", rstProveedor!CategoriaII)
                    
                    WCategoriaI = ""
                    WCategoriaII = ""
        
                    If ZCategoriaI = 1 Then
                        WCategoriaI = "A"
                            Else
                        If ZCategoriaI = 2 Then
                            WCategoriaI = "B"
                                Else
                            If ZCategoriaI = 3 Then
                                WCategoriaI = "C"
                                    Else
                                If ZCategoriaI = 4 Then
                                    WCategoriaI = "E"
                                End If
                            End If
                        End If
                    End If
                    
                    WCategoriaII = "S/C"
                    If ZCategoriaII = 1 Then
                        WCategoriaII = "Muy Bueno"
                            Else
                        If ZCategoriaII = 2 Then
                            WCategoriaII = "Bueno"
                                Else
                            If ZCategoriaII = 3 Then
                                WCategoriaII = "Regular"
                                    Else
                                If ZCategoriaII = 4 Then
                                    WCategoriaII = "Malo"
                                End If
                            End If
                        End If
                    End If
                    
                    If WCategoriaI <> "" And WCategoriaII <> "" Then
                        WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                    End If
                    
                    rstProveedor.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DescriProveedor = WDescriProveedor
                !DescriArticulo = WDescriArticulo
                
                Select Case Moneda2.ListIndex
                    Case 0
                        !Titulo = "(En Dolares)"
                    Case 2
                        !Titulo = "(En Euros)"
                    Case Else
                        !Titulo = "(En Pesos)"
                End Select
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cotizaciones por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Listcot.Articulo} in " + Chr$(34) + XArt2.Text + Chr$(34) + " to " + Chr$(34) + XArt2.Text + Chr$(34)
   
    Listado.Destination = 0
    Listado.ReportFileName = "WCotart.rpt"
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
    Call Conecta_Empresa

End Sub

Private Sub XAcepta3_Click()

    On Error GoTo WError
    
    XEmpresa = Wempresa
        
    Select Case Val(XEmpresa)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    Da = ""
    With rstImpCtaCtePrv
        .Index = "ClaveImpre"
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
    
    XParam = "'" + XProv3.Text + "','" _
                 + XProv3.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                XProveedor = !Proveedor
                XLetra = !Letra
                XTipo = !Tipo
                XPunto = !Punto
                XNumero = !Numero
                XFecha = !Fecha
                XEstado = !Estado
                Xvencimiento = !Vencimiento
                XVencimiento1 = !Vencimiento1
                XNroInterno = !nrointerno
                XTotal = !Total
                XSaldo = !Saldo
                XClave = !Clave
                XOrdFecha = !OrdFecha
                XOrdVencimiento = !OrdVencimiento
                XImpre = !Impre
                
                With rstImpCtaCtePrv
                
                    .Index = "CtaCte"
                    .Seek "=", XClave
                    If .NoMatch Then
                        .AddNew
                        !Proveedor = XProveedor
                        !Letra = XLetra
                        !Tipo = XTipo
                        !Punto = XPunto
                        !Numero = XNumero
                        !Fecha = XFecha
                        !Estado = XEstado
                        !Vencimiento = Xvencimiento
                        !Vencimiento1 = XVencimiento1
                        !nrointerno = XNroInterno
                        !Total = XTotal
                        !Saldo = XSaldo
                        !Clave = XClave
                        !OrdFecha = XOrdFecha
                        !OrdVencimiento = XOrdVencimiento
                        !Impre = XImpre
                        !Titulo = WTitulo
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstCtaPrv.Close
    
    End If
    
    
    Pasa = 0
    Acumula = 0

    With rstImpCtaCtePrv
            .Index = "ClaveImpre"
            .MoveFirst
            Do
                Rem If !Proveedor > Hasta.Text Then
                Rem    Exit Do
                Rem End If
                If Pasa = 0 Then
                    Pasa = 1
                    Acumula = 0
                    Corte = !Proveedor
                End If
                If Corte <> !Proveedor Then
                    Acumula = 0
                    Corte = !Proveedor
                End If
                .Edit
                !SaldoList = 0
                If !Proveedor >= XProv3.Text And !Proveedor <= XProv3.Text Then
                    WSaldo = !Saldo
                    Call Redondeo(WSaldo)
                    !SaldoList = WSaldo
                    Acumula = Acumula + WSaldo
                    !Acumulado = Acumula
                End If
                
                WProveedor = !Proveedor
                WNombre = ""
                WCheque = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    WNombre = rstProveedor!Nombre
                    WCheque = rstProveedor!NombreCheque
                    rstProveedor.Close
                End If
                
                !Nombre = WNombre
                !Cheque = WCheque
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Listado.GroupSelectionFormula = "{CtaCtePrv.Proveedor} in " + Chr$(34) + XProv3.Text + Chr$(34) + " to " + Chr$(34) + XProv3.Text + Chr$(34) + " and {CtaCtePrv.Saldolist} <> 0.0"
    Listado.Destination = 0
    Listado.DataFiles(0) = XEmpresa + "Auxi.mdb"
    Listado.ReportFileName = "wccprv.rpt"
    
    Listado.Action = 1
     
    Call Conecta_Empresa
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub XArt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If XArt2.Text <> "  -   -   " Then
            XArt2.Text = UCase(XArt2.Text)
            spArticulo = "ConsultaArticulo " + "'" + XArt2.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                XDesArt2.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                Call XAcepta2_Click
            End If
                Else
            TipoConsulta = "2"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 1
            Call Opcion_Click
        End If
    End If
End Sub

Private Sub XCancela_Click()
    XCoti.Visible = False
End Sub

Private Sub XCancela1_Click()
    XCotPrv.Visible = False
End Sub

Private Sub XCancela2_Click()
    XCotart.Visible = False
End Sub

Private Sub XCancela3_Click()
    XCc.Visible = False
End Sub

Private Sub XConsulta1_Click()
    TipoConsulta = "2"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Call Opcion_Click
    Ayuda.SetFocus
End Sub

Private Sub XConsulta2_Click()
    TipoConsulta = "2"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    Call Opcion_Click
End Sub

Private Sub XConsulta3_Click()
    TipoConsulta = "4"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Call Opcion_Click
    Ayuda.SetFocus
End Sub


Private Sub XProv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProv1.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProv1.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProv1.Text = rstProveedor!Proveedor
                    XDesProv1.Caption = rstProveedor!Nombre
                    Call XAcepta1_Click
            End If
                Else
            TipoConsulta = "2"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)

End Sub

Private Sub XProv3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProv3.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProv3.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProv3.Text = rstProveedor!Proveedor
                    XDesProv3.Caption = rstProveedor!Nombre
                    Call XAcepta3_Click
            End If
                Else
            TipoConsulta = "4"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)

End Sub


Private Sub XProve_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProve.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProve.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProve.Text = rstProveedor!Proveedor
                    XDesProve.Caption = rstProveedor!Nombre
                    XArti.SetFocus
                        Else
                    XProve.SetFocus
            End If
                Else
            TipoConsulta = "3"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub XArti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If XArti.Text <> "  -   -   " Then
            XArti.Text = UCase(XArti.Text)
            spArticulo = "ConsultaArticulo " + "'" + XArti.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                XDesArti.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                XPrec.SetFocus
                    Else
                XArti.SetFocus
            End If
                Else
            TipoConsulta = "3"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 1
            Call Opcion_Click
        End If
    End If

End Sub

Private Sub XPrec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XCondicion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub XCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XObservaciones.SetFocus
    End If

End Sub

Private Sub XObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XProve.SetFocus
    End If

End Sub

Private Sub Verifica_Atrasos()

    XEmpresa = Wempresa
    ZEntra = "N"

    ZZFechaII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ZZOrdFechaII = Right$(ZZFechaII, 4) + Mid$(ZZFechaII, 4, 2) + Left$(ZZFechaII, 2)
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
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
    
    For a = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        Wempresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Proveedor = " + "'" + Proveedor.Text + "'"
        ZSql = ZSql + " and Orden.Saldo > 0"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
            
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                        ZZFecha = rstOrden!fecha2
                        ZZOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                        
                        If ZZOrdFecha < ZZOrdFechaII Then
                            ZEntra = "S"
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
        
    Next a
    
    Call Conecta_Empresa
    
    If ZEntra = "S" Then
        m$ = "El proveedor posee ordenes de compra con atraso"
        ddfd = MsgBox(m$, 0, "Ordenes de compra")
    End If
    
End Sub





Private Sub Impresion_Especificaciones()
                            
    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
                
    ZSql = "DELETE CertificadoMp"
    spCertificadoMp = ZSql
    Set rstCertificadoMp = db.OpenRecordset(spCertificadoMp, dbOpenSnapshot, dbSQLPassThrough)
    
    
    For ZZCiclo = 1 To 100
    
        If ZZArtiCerti(ZZCiclo) <> "" Then
        
            ZZZArticulo = ZZArtiCerti(ZZCiclo)
            ZZVersion = 0
            
            Sql1 = "Select EspecificacionesUnificaII.Producto"
            Sql2 = " FROM EspecificacionesUnificaII"
            Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + ZZZArticulo + "'"
            spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnificaII.RecordCount > 0 Then
            
                rstEspecificacionesUnificaII.Close
                                
                Erase ZEnsayo
                    
                Sql1 = "Select *"
                Sql2 = " FROM EspecificacionesUnifica"
                Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + ZZZArticulo + "'"
                spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnifica.RecordCount > 0 Then
                
                    ZZVersion = rstEspecificacionesUnifica!Version
                
                    ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
                    ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
                    ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
                    ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
                    ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
                    ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
                    ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
                    ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
                    ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
                    ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
                    ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
                    ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
                    ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
                    ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
                    ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
                    ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
                    ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
                    ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
                    ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
                    ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
                    
                    rstEspecificacionesUnifica.Close
                                    
                End If
                
                ZZIValor1 = ""
                ZZIValor2 = ""
                ZZIValor3 = ""
                ZZIValor4 = ""
                ZZIValor5 = ""
                ZZIValor6 = ""
                ZZIValor7 = ""
                ZZIValor8 = ""
                ZZIValor9 = ""
                ZZIValor10 = ""
                ZZIValor11 = ""
                ZZIValor12 = ""
                ZZIValor13 = ""
                ZZIValor14 = ""
                ZZIValor15 = ""
                ZZIValor16 = ""
                ZZIValor17 = ""
                ZZIValor18 = ""
                ZZIValor19 = ""
                ZZIValor20 = ""
                
                Rem dada
                Rem dada
                Rem dada
                Rem dada
                Rem dada
                Rem dada
            Rem    ZZZArticulo = "PB-108-100"
                Rem BY NAN
                
                Sql1 = "Select DescripcionIngles,cas,IValor1,Ivalor2,Ivalor3,Ivalor4,Ivalor5,Ivalor6,Ivalor7,Ivalor8,Ivalor9,Ivalor10"
                Sql2 = " FROM EspecificacionesUnificaII"
                Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + ZZZArticulo + "'"
                spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnificaII.RecordCount > 0 Then
                    
                    ZZDescripcionIngles = IIf(IsNull(rstEspecificacionesUnificaII!DescripcionIngles), "", rstEspecificacionesUnificaII!DescripcionIngles)
                    ZZCas = IIf(IsNull(rstEspecificacionesUnificaII!Cas), "", rstEspecificacionesUnificaII!Cas)
                   
                    ZZDescripcionIngles = Trim(ZZDescripcionIngles)
                    ZZCas = Trim(ZZCas)
                    
                    ZDescriII(1) = IIf(IsNull(rstEspecificacionesUnificaII!IValor1), "", rstEspecificacionesUnificaII!IValor1)
                    ZDescriII(2) = IIf(IsNull(rstEspecificacionesUnificaII!IValor2), "", rstEspecificacionesUnificaII!IValor2)
                    ZDescriII(3) = IIf(IsNull(rstEspecificacionesUnificaII!IValor3), "", rstEspecificacionesUnificaII!IValor3)
                    ZDescriII(4) = IIf(IsNull(rstEspecificacionesUnificaII!IValor4), "", rstEspecificacionesUnificaII!IValor4)
                    ZDescriII(5) = IIf(IsNull(rstEspecificacionesUnificaII!IValor5), "", rstEspecificacionesUnificaII!IValor5)
                    ZDescriII(6) = IIf(IsNull(rstEspecificacionesUnificaII!IValor6), "", rstEspecificacionesUnificaII!IValor6)
                    ZDescriII(7) = IIf(IsNull(rstEspecificacionesUnificaII!IValor7), "", rstEspecificacionesUnificaII!IValor7)
                    ZDescriII(8) = IIf(IsNull(rstEspecificacionesUnificaII!IValor8), "", rstEspecificacionesUnificaII!IValor8)
                    ZDescriII(9) = IIf(IsNull(rstEspecificacionesUnificaII!IValor9), "", rstEspecificacionesUnificaII!IValor9)
                    ZDescriII(10) = IIf(IsNull(rstEspecificacionesUnificaII!IValor10), "", rstEspecificacionesUnificaII!IValor10)
                   rstEspecificacionesUnificaII.Close
                  End If
                    
                Sql1 = "Select IValor11,Ivalor12,Ivalor13,Ivalor14,Ivalor15,Ivalor16,Ivalor17,Ivalor18,Ivalor19,Ivalor20"
                Sql2 = " FROM EspecificacionesUnificaII"
                Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + ZZZArticulo + "'"
                spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnificaII.RecordCount > 0 Then
                    
                    
                    ZDescriII(11) = IIf(IsNull(rstEspecificacionesUnificaII!IValor11), "", rstEspecificacionesUnificaII!IValor11)
                    ZDescriII(12) = IIf(IsNull(rstEspecificacionesUnificaII!IValor12), "", rstEspecificacionesUnificaII!IValor12)
                    ZDescriII(13) = IIf(IsNull(rstEspecificacionesUnificaII!IValor13), "", rstEspecificacionesUnificaII!IValor13)
                    ZDescriII(14) = IIf(IsNull(rstEspecificacionesUnificaII!IValor14), "", rstEspecificacionesUnificaII!IValor14)
                    ZDescriII(15) = IIf(IsNull(rstEspecificacionesUnificaII!IValor15), "", rstEspecificacionesUnificaII!IValor15)
                    ZDescriII(16) = IIf(IsNull(rstEspecificacionesUnificaII!IValor16), "", rstEspecificacionesUnificaII!IValor16)
                    ZDescriII(17) = IIf(IsNull(rstEspecificacionesUnificaII!IValor17), "", rstEspecificacionesUnificaII!IValor17)
                    ZDescriII(18) = IIf(IsNull(rstEspecificacionesUnificaII!IValor18), "", rstEspecificacionesUnificaII!IValor18)
                    ZDescriII(19) = IIf(IsNull(rstEspecificacionesUnificaII!IValor19), "", rstEspecificacionesUnificaII!IValor19)
                    ZDescriII(20) = IIf(IsNull(rstEspecificacionesUnificaII!IValor20), "", rstEspecificacionesUnificaII!IValor20)
                    
                    rstEspecificacionesUnificaII.Close
                                    
                End If
                
                
                
             Rem   Sql1 = "Select *"
              Rem  Sql2 = " FROM EspecificacionesUnificaII"
              Rem  Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + ZZZArticulo + "'"
              Rem  spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
              Rem  Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
              Rem  If rstEspecificacionesUnificaII.RecordCount > 0 Then
                    
                 Rem   ZZDescripcionIngles = IIf(IsNull(rstEspecificacionesUnificaII!DescripcionIngles), "", rstEspecificacionesUnificaII!DescripcionIngles)
                 Rem   ZZCas = IIf(IsNull(rstEspecificacionesUnificaII!Cas), "", rstEspecificacionesUnificaII!Cas)
                   
                 Rem   ZZDescripcionIngles = Trim(ZZDescripcionIngles)
                 Rem   ZZCas = Trim(ZZCas)
                    
                 Rem   ZDescriII(1) = IIf(IsNull(rstEspecificacionesUnificaII!IValor1), "", rstEspecificacionesUnificaII!IValor1)
                 Rem   ZDescriII(2) = IIf(IsNull(rstEspecificacionesUnificaII!IValor2), "", rstEspecificacionesUnificaII!IValor2)
                 Rem   ZDescriII(3) = IIf(IsNull(rstEspecificacionesUnificaII!IValor3), "", rstEspecificacionesUnificaII!IValor3)
                 Rem   ZDescriII(4) = IIf(IsNull(rstEspecificacionesUnificaII!IValor4), "", rstEspecificacionesUnificaII!IValor4)
                 Rem   ZDescriII(5) = IIf(IsNull(rstEspecificacionesUnificaII!IValor5), "", rstEspecificacionesUnificaII!IValor5)
                 Rem   ZDescriII(6) = IIf(IsNull(rstEspecificacionesUnificaII!IValor6), "", rstEspecificacionesUnificaII!IValor6)
                 Rem   ZDescriII(7) = IIf(IsNull(rstEspecificacionesUnificaII!IValor7), "", rstEspecificacionesUnificaII!IValor7)
                 Rem   ZDescriII(8) = IIf(IsNull(rstEspecificacionesUnificaII!IValor8), "", rstEspecificacionesUnificaII!IValor8)
                 Rem   ZDescriII(9) = IIf(IsNull(rstEspecificacionesUnificaII!IValor9), "", rstEspecificacionesUnificaII!IValor9)
                 Rem   ZDescriII(10) = IIf(IsNull(rstEspecificacionesUnificaII!IValor10), "", rstEspecificacionesUnificaII!IValor10)
                 Rem   ZDescriII(11) = IIf(IsNull(rstEspecificacionesUnificaII!IValor11), "", rstEspecificacionesUnificaII!IValor11)
                 Rem   ZDescriII(12) = IIf(IsNull(rstEspecificacionesUnificaII!IValor12), "", rstEspecificacionesUnificaII!IValor12)
                 Rem   ZDescriII(13) = IIf(IsNull(rstEspecificacionesUnificaII!IValor13), "", rstEspecificacionesUnificaII!IValor13)
                 Rem   ZDescriII(14) = IIf(IsNull(rstEspecificacionesUnificaII!IValor14), "", rstEspecificacionesUnificaII!IValor14)
                 Rem   ZDescriII(15) = IIf(IsNull(rstEspecificacionesUnificaII!IValor15), "", rstEspecificacionesUnificaII!IValor15)
                 Rem   ZDescriII(16) = IIf(IsNull(rstEspecificacionesUnificaII!IValor16), "", rstEspecificacionesUnificaII!IValor16)
                 Rem   ZDescriII(17) = IIf(IsNull(rstEspecificacionesUnificaII!IValor17), "", rstEspecificacionesUnificaII!IValor17)
                 Rem   ZDescriII(18) = IIf(IsNull(rstEspecificacionesUnificaII!IValor18), "", rstEspecificacionesUnificaII!IValor18)
                 Rem   ZDescriII(19) = IIf(IsNull(rstEspecificacionesUnificaII!IValor19), "", rstEspecificacionesUnificaII!IValor19)
                  Rem  ZDescriII(20) = IIf(IsNull(rstEspecificacionesUnificaII!IValor20), "", rstEspecificacionesUnificaII!IValor20)
                    
                 Rem   rstEspecificacionesUnificaII.Close
                                    
                Rem End If
                Rem Sql1 = "Select *"
                 Rem Sql2 = " FROM EspecificacionesUnificaII"
                Rem Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + ZZZArticulo + "'"
                Rem spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
                Rem Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
                Rem If rstEspecificacionesUnificaII.RecordCount > 0 Then
                    
                Rem    ZZDescripcionIngles = IIf(IsNull(rstEspecificacionesUnificaII!DescripcionIngles), "", rstEspecificacionesUnificaII!DescripcionIngles)
                Rem    ZZCas = IIf(IsNull(rstEspecificacionesUnificaII!Cas), "", rstEspecificacionesUnificaII!Cas)
                   
                Rem    ZZDescripcionIngles = Trim(ZZDescripcionIngles)
                Rem    ZZCas = Trim(ZZCas)
                    
                Rem    ZDescriII(1) = IIf(IsNull(rstEspecificacionesUnificaII!IValor1), "", rstEspecificacionesUnificaII!IValor1)
                Rem    ZDescriII(2) = IIf(IsNull(rstEspecificacionesUnificaII!IValor2), "", rstEspecificacionesUnificaII!IValor2)
                Rem    ZDescriII(3) = IIf(IsNull(rstEspecificacionesUnificaII!IValor3), "", rstEspecificacionesUnificaII!IValor3)
                Rem    ZDescriII(4) = IIf(IsNull(rstEspecificacionesUnificaII!IValor4), "", rstEspecificacionesUnificaII!IValor4)
                Rem    ZDescriII(5) = IIf(IsNull(rstEspecificacionesUnificaII!IValor5), "", rstEspecificacionesUnificaII!IValor5)
                 Rem   ZDescriII(6) = IIf(IsNull(rstEspecificacionesUnificaII!IValor6), "", rstEspecificacionesUnificaII!IValor6)
                 Rem   ZDescriII(7) = IIf(IsNull(rstEspecificacionesUnificaII!IValor7), "", rstEspecificacionesUnificaII!IValor7)
                 Rem   ZDescriII(8) = IIf(IsNull(rstEspecificacionesUnificaII!IValor8), "", rstEspecificacionesUnificaII!IValor8)
                 Rem   ZDescriII(9) = IIf(IsNull(rstEspecificacionesUnificaII!IValor9), "", rstEspecificacionesUnificaII!IValor9)
                 Rem   ZDescriII(10) = IIf(IsNull(rstEspecificacionesUnificaII!IValor10), "", rstEspecificacionesUnificaII!IValor10)
                 Rem   ZDescriII(11) = IIf(IsNull(rstEspecificacionesUnificaII!IValor11), "", rstEspecificacionesUnificaII!IValor11)
                 Rem   ZDescriII(12) = IIf(IsNull(rstEspecificacionesUnificaII!IValor12), "", rstEspecificacionesUnificaII!IValor12)
                 Rem   ZDescriII(13) = IIf(IsNull(rstEspecificacionesUnificaII!IValor13), "", rstEspecificacionesUnificaII!IValor13)
                 Rem   ZDescriII(14) = IIf(IsNull(rstEspecificacionesUnificaII!IValor14), "", rstEspecificacionesUnificaII!IValor14)
                 Rem   ZDescriII(15) = IIf(IsNull(rstEspecificacionesUnificaII!IValor15), "", rstEspecificacionesUnificaII!IValor15)
                 Rem   ZDescriII(16) = IIf(IsNull(rstEspecificacionesUnificaII!IValor16), "", rstEspecificacionesUnificaII!IValor16)
                 Rem   ZDescriII(17) = IIf(IsNull(rstEspecificacionesUnificaII!IValor17), "", rstEspecificacionesUnificaII!IValor17)
                 Rem   ZDescriII(18) = IIf(IsNull(rstEspecificacionesUnificaII!IValor18), "", rstEspecificacionesUnificaII!IValor18)
                 Rem   ZDescriII(19) = IIf(IsNull(rstEspecificacionesUnificaII!IValor19), "", rstEspecificacionesUnificaII!IValor19)
                 Rem   ZDescriII(20) = IIf(IsNull(rstEspecificacionesUnificaII!IValor20), "", rstEspecificacionesUnificaII!IValor20)
                    
                  Rem  rstEspecificacionesUnificaII.Close
                                    
              Rem  End If
                
                
                For Cicla = 1 To 20
                    ZZDescri = ""
                    If Val(ZEnsayo(Cicla)) <> 0 Then
                    
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZZDescri = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            rstEnsayo.Close
                        End If
                        
                        Auxi1 = Str$(Cicla)
                        Call Ceros(Auxi1, 2)
                        
                        ZClave = ZZZArticulo + Auxi1
                            
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO CertificadoMp ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Terminado ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Descripcion ,"
                        ZSql = ZSql + "Examen ,"
                        ZSql = ZSql + "Valor ,"
                        ZSql = ZSql + "Version ,"
                        ZSql = ZSql + "Cas )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZClave + "',"
                        ZSql = ZSql + "'" + ZZZArticulo + "',"
                        ZSql = ZSql + "'" + Str$(Cicla) + "',"
                        ZSql = ZSql + "'" + Trim(ZZDescripcionIngles) + "',"
                        ZSql = ZSql + "'" + Trim(ZZDescri) + "',"
                        ZSql = ZSql + "'" + Trim(ZDescriII(Cicla)) + "',"
                        ZSql = ZSql + "'" + Str$(ZZVersion) + "',"
                        ZSql = ZSql + "'" + Trim(ZZCas) + "')"
                
                        spCertificadoMp = ZSql
                        Set rstCertificadoMp = db.OpenRecordset(spCertificadoMp, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                            
                Next Cicla
                
            End If
        
        End If
        
    Next ZZCiclo
            
    Listado.WindowTitle = "Certificado de Analisis"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.Destination = 1
    Listado.Destination = 0
            
    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
        Listado.ReportFileName = "recuired specs po nro.rpt"
            Else
        Listado.ReportFileName = "CertificadoMpPelli.rpt"
    End If
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT CertificadoMp.Terminado, CertificadoMp.Renglon, CertificadoMp.Descripcion, CertificadoMp.Examen, CertificadoMp.Valor, CertificadoMp.Cas " _
            + "From " _
            + DSQ + ".dbo.CertificadoMp CertificadoMp " _
            + "Where " _
            + "CertificadoMp.Renglon >= 0 AND " _
            + "CertificadoMp.Renglon <= 999999"

    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = ""
    Listado.SelectionFormula = ""
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.Action = 1
                
    Call Conecta_Empresa
        
End Sub

Private Sub Cuotas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Cuotas.Text) <> 0 Then
            MesCuota.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub MesCuota_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(MesCuota.Text) <> 0 Then
            AnoCuota.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub AnoCuota_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Cuotas.Text) <> 0 And Val(MesCuota.Text) <> 0 And Val(AnoCuota.Text) <> 0 Then
            If Val(MesCuota.Text) > 12 Then
                m$ = "Mes Incorrecto"
                a% = MsgBox(m$, 0, "Datos pagos Pyme Nacion")
                Exit Sub
            End If
            If Val(AnoCuota.Text) < 2011 Or Val(AnoCuota.Text) > 2050 Then
                m$ = "Ano Incorrecto"
                a% = MsgBox(m$, 0, "Datos pagos Pyme Nacion")
                Exit Sub
            End If
            PantaPyme.Visible = False
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CierraPyme_Click()
    PantaPyme.Visible = False
End Sub





Public Sub SendEmail()

    Dim objOutlook As Object
    Dim objMailItem

    Dim NumOfPath As Integer, i As Integer
    Dim AtachPath As String

    On Error GoTo 10

    NumOfPath = 0
    AllPath = ""
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
        .To = EmailAddress
        .cc = CopiaAddress
        .Subject = MSubject
        .Body = MBody
        Rem .Attachments.Add MAttach
        If MAttachI <> "" Then
            .Attachments.Add MAttachI
        End If
        If MAttachII <> "" Then
            .Attachments.Add MAttachII
        End If
        If MAttachIII > "" Then
            .Attachments.Add MAttachIII
        End If
        If MAttachIV <> "" Then
            .Attachments.Add MAttachIV
        End If
        If MAttachV <> "" Then
            .Attachments.Add MAttachV
        End If
        .Send
    End With

    Set objMailItem = Nothing
    Set objOutlook = Nothing
            
    Exit Sub

exit10:
    Exit Sub

10:
    If Err.Number = 429 Then
        MsgBox "Error on connecting with Outlook"
            Else
        MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    End If
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    AllPath = ""

    Resume exit10

End Sub
    
    
    
    
    






Private Sub Asunto_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        CargaAsunto.Visible = False
    
        EmailAddress = ZZPasaTo
        CopiaAddress = ZZPasaCC
        MSubject = Asunto.Text
        MBody = ZZPasaBody
        MAttach = ZZPasaFile
        MAttachI = ZZPasaFile
        MAttachII = ""
        MAttachIII = ""
        MAttachIV = ""
        MAttachV = ""
    
        SendEmail
        
        Call Limpia_Click
    
        WVector.Col = 1
        WVector.Row = 1
        Call WVector_Click
        
        Orden.SetFocus
    
        Call Conecta_Empresa
        
    End If
    
End Sub


Private Sub Limpia_Vector()

    Rem wvector.Height = 4095
    Rem wvector.Left = 120
    Rem wvector.Top = 1200
    Rem wvector.Width = 10000

    WVector.Clear
    WVector.Font.Bold = True
    
    WVector.FixedCols = 1
    WVector.Cols = 9
    WVector.FixedRows = 1
    WVector.Rows = 100
    
    WVector.ColWidth(0) = 200
    WVector.Row = 0
    For Ciclo = 1 To WVector.Cols - 1
        WVector.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector.Text = "Producto"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector.Text = "Descripcion"
                WVector.ColWidth(Ciclo) = 3500
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector.Text = "Cantidad"
                WVector.ColWidth(Ciclo) = 1000
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector.Text = "Precio"
                WVector.ColWidth(Ciclo) = 1000
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector.Text = "1ra Fecha"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                WVector.Text = "Ult Fecha"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignlefttCenter
            Case 7
                WVector.Text = "Condicion"
                WVector.ColWidth(Ciclo) = 1200
                WVector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector.Text = "Bultos"
                WVector.ColWidth(Ciclo) = 700
                WVector.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector.Cols - 1
        WAncho = WAncho + WVector.ColWidth(Ciclo)
    Next Ciclo
    WVector.Width = WAncho

    ' Size the columns.
    Font.Name = WVector.Font.Name
    Font.Size = WVector.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector.AllowUserResizing = flexResizeBoth
    
    WVector.Visible = True
    
    WVector.Col = 1
    WVector.Row = 1
    Rem Call WVector_Click
    
End Sub
