VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos en $"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.Frame CargaLote2 
      Caption         =   "Ingreso de Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   960
      TabIndex        =   73
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox ZZEnvase5 
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
         TabIndex        =   94
         Text            =   " "
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox ZZCantiEnv5 
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
         Left            =   5280
         TabIndex        =   93
         Text            =   " "
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox ZZEnvase4 
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
         TabIndex        =   92
         Text            =   " "
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox ZZCantiEnv4 
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
         Left            =   5280
         TabIndex        =   91
         Text            =   " "
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox ZZEnvase3 
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
         TabIndex        =   90
         Text            =   " "
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox ZZCantiEnv3 
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
         Left            =   5280
         TabIndex        =   89
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox ZZEnvase2 
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
         TabIndex        =   88
         Text            =   " "
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox ZZCantiEnv2 
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
         Left            =   5280
         TabIndex        =   87
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox ZZEnvase1 
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
         TabIndex        =   86
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox ZZCantiEnv1 
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
         Left            =   5280
         TabIndex        =   85
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox ZZCanti5 
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
         TabIndex        =   84
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox ZZCanti4 
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
         TabIndex        =   83
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox ZZPartida5 
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
         MaxLength       =   10
         TabIndex        =   82
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox ZZPartida4 
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
         MaxLength       =   10
         TabIndex        =   81
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox ZZCanti3 
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
         TabIndex        =   80
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox ZZCanti2 
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
         TabIndex        =   79
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox ZZCanti1 
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
         TabIndex        =   78
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox ZZPartida3 
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
         MaxLength       =   10
         TabIndex        =   77
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox ZZPartida2 
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
         MaxLength       =   10
         TabIndex        =   76
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox ZZPartida1 
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
         MaxLength       =   10
         TabIndex        =   75
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton FinCarga 
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
         Height          =   495
         Left            =   2880
         TabIndex        =   74
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label ZZDescri5 
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
         Left            =   4200
         TabIndex        =   104
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label ZZDescri4 
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
         Left            =   4200
         TabIndex        =   103
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ZZDescri3 
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
         Left            =   4200
         TabIndex        =   102
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label ZZDescri2 
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
         Left            =   4200
         TabIndex        =   101
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canti."
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
         Left            =   5280
         TabIndex        =   100
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   4200
         TabIndex        =   99
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envase"
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
         TabIndex        =   98
         Top             =   240
         Width           =   975
      End
      Begin VB.Label ZZDescri1 
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
         Left            =   4200
         TabIndex        =   97
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
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
         Left            =   1800
         TabIndex        =   96
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partida"
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
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame PantaMotivo 
      Height          =   1815
      Left            =   600
      TabIndex        =   108
      Top             =   2400
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox DescriMotivo 
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
         MaxLength       =   50
         TabIndex        =   110
         Top             =   720
         Width           =   9855
      End
      Begin VB.ComboBox ConceptoAtraso 
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
         Left            =   240
         TabIndex        =   109
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "MOTIVO DE RETRASO DE CUMPLIMIENTO DEL PEDIDO"
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
         Left            =   240
         TabIndex        =   111
         Top             =   360
         Width           =   9735
      End
   End
   Begin VB.CommandButton ConsultaPedido 
      Caption         =   "Consulta Pedidos"
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
      Left            =   8040
      TabIndex        =   107
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      Height          =   2655
      Left            =   5400
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WCanti5 
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
         Left            =   1320
         TabIndex        =   68
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox WCanti4 
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
         Left            =   1320
         TabIndex        =   67
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WLote5 
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
         Left            =   240
         TabIndex        =   66
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox WLote4 
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
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox WCanti3 
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
         Left            =   1320
         TabIndex        =   64
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WCanti2 
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
         Left            =   1320
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCanti1 
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
         Left            =   1320
         TabIndex        =   62
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Wlote3 
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
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WLote2 
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
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WLote1 
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
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
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
         Left            =   1440
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partida"
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
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Tipoventa 
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
      Left            =   3240
      TabIndex        =   55
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton ReImpre 
      Caption         =   "ReImpresion"
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
      Left            =   10200
      TabIndex        =   54
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Canti5 
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
      TabIndex        =   48
      Text            =   " "
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Canti4 
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
      TabIndex        =   47
      Text            =   " "
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Canti3 
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
      TabIndex        =   46
      Text            =   " "
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Canti2 
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
      TabIndex        =   45
      Text            =   " "
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Canti1 
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
      TabIndex        =   44
      Text            =   " "
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Envase5 
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
      Left            =   120
      TabIndex        =   43
      Text            =   " "
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Envase4 
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
      Left            =   120
      TabIndex        =   42
      Text            =   " "
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Envase3 
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
      Left            =   120
      TabIndex        =   41
      Text            =   " "
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Envase2 
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
      Left            =   120
      TabIndex        =   40
      Text            =   " "
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Envase1 
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
      Left            =   120
      TabIndex        =   39
      Text            =   " "
      Top             =   6000
      Width           =   975
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   34
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
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
      Left            =   9120
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   8760
      TabIndex        =   23
      Top             =   5640
      Width           =   2535
      Begin VB.Label Label26 
         Caption         =   "IB Ciudad"
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
         Left            =   120
         TabIndex        =   106
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label ImpoIbCiudad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   105
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label ImpoIbTucu 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   72
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "IB Tucu."
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
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "IB Bs.As."
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
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Width           =   975
      End
      Begin VB.Label ImpoIb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   69
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Interes"
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
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Dto."
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
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   31
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
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
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Pedido 
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   120
      Width           =   1335
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
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   720
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
      Left            =   6120
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Orden 
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
      MaxLength       =   10
      TabIndex        =   18
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Remito 
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
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.TextBox Cliente 
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
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   7
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
      Height          =   450
      Left            =   8040
      TabIndex        =   5
      Top             =   120
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
      Height          =   450
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   975
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
      Height          =   450
      Left            =   10200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   1320
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
      ItemData        =   "prgfactup.frx":0000
      Left            =   6480
      List            =   "prgfactup.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "prgfactup.frx":0015
      TabIndex        =   2
      Top             =   1800
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Label Descri5 
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
      Left            =   1200
      TabIndex        =   53
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Descri4 
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
      Left            =   1200
      TabIndex        =   52
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Descri3 
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
      Left            =   1200
      TabIndex        =   51
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Descri2 
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
      Left            =   1200
      TabIndex        =   50
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Descri1 
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
      Left            =   1200
      TabIndex        =   49
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label12 
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
      Left            =   5640
      TabIndex        =   33
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Orden de compra"
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
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
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
      Left            =   3240
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesCliente 
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
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      TabIndex        =   10
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
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Factura"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgFactup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private XImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private dada As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private WAceptada As String
Private Stk(19, 4) As String
Private Envase(5, 2) As String
Private parcial As String
Private Auxiliar(100, 15) As String
Private RestaPedido(100, 3) As String
Private ClavePedido(100)
Private BajaLote(5, 2) As String
Private XLote(100, 20) As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim XParam As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WSaldo4 As Double
Dim WSaldo5 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim XSaldo4 As String
Dim XSaldo5 As String
Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim Compara As Double
Private WCodIb As Integer
Private WCodIbTucu As Integer
Private WCodIbCiudad As Integer
Private WImpoIb As Double
Private WImpoIbTucu As Double
Private WImpoIbCiudad As Double
Dim ControlLote(5, 2) As String
Dim WSal As Double
Private WAdicional As Double
Private WTipoPedido As String
Private WPorceIb As Double

Dim ZLote1 As String
Dim ZCantidad1 As String
Dim ZLote2 As String
Dim ZCantidad2 As String
Dim ZLote3 As String
Dim ZCantidad3 As String
Dim ZLote4 As String
Dim ZCantidad4 As String
Dim ZLote5 As String
Dim ZCantidad5 As String
Dim ZEnv1 As String
Dim ZCantiEnv1 As String
Dim ZEnv2 As String
Dim ZCantiEnv2 As String
Dim ZEnv3 As String
Dim ZCantiEnv3 As String
Dim ZEnv4 As String
Dim ZCantiEnv4 As String
Dim ZEnv5 As String
Dim ZCantiEnv5 As String

Dim DiaFeriado(100) As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim VectorCosto(100, 3) As String
Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZZEnvase(10) As String
Dim ZZCanti(10) As String

Dim ZZImpreDespa(100, 5) As String
Dim ZZImpreDespaII(100, 5) As String

Dim ZVersionPedido As Integer
Dim ZVersionAtraso As Integer
Dim ZSedronar As Integer
Dim ZNroSedronar As String

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
        rstPago.Close
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias2 = rstPago!Dias
        WPlazo2 = rstPago!Plazo
        rstPago.Close
   End If
    
    Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)

End Sub

Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 3
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = "S"
    
    XLote(WRow, 1) = ""
    XLote(WRow, 2) = ""
    XLote(WRow, 3) = ""
    XLote(WRow, 4) = ""
    XLote(WRow, 5) = ""
    XLote(WRow, 6) = ""
    XLote(WRow, 7) = ""
    XLote(WRow, 8) = ""
    XLote(WRow, 9) = ""
    XLote(WRow, 11) = ""
    XLote(WRow, 12) = ""
    XLote(WRow, 13) = ""
    XLote(WRow, 14) = ""
    XLote(WRow, 15) = ""
    XLote(WRow, 16) = ""
    XLote(WRow, 17) = ""
    XLote(WRow, 18) = ""
    XLote(WRow, 19) = ""
    XLote(WRow, 20) = ""
    
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            If Val(Cantidad) <> 0 Then
                WNeto = WNeto + (Val(Cantidad) * Val(Precio))
            End If
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0

    Rem If Val(Paridad.Text) <> 0 Then
    Rem    WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    If WTasa <> 0 Then
        WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
        Call Redondeo(WImpoInteres)
        WNeto = WNeto + WImpoInteres
    End If
    
    WIva1 = 0
    WIva2 = 0
    WImpoIb = 0
    WImpoIbTucu = 0
    WImpoIbCiudad = 0
    
    ZFechaCompa = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    If ZFechaCompa >= "20071201" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
            WPorceIb = IIf(IsNull(rstCliente!PorceIb), "0", rstCliente!PorceIb)
            rstCliente.Close
        End If
        
        If ZZIb <> 2 Then
            WImpoIb = WNeto * (WPorceIb / 100)
            Call Redondeo(WImpoIb)
        End If
    
            Else
    
        Select Case WCodIb
            Case 0, 1
                Select Case Val(WCodIva)
                    Case 1
                        WImpoIb = WNeto * 0.025
                    Case 2, 4, 5, 6
                        WImpoIb = WNeto * 0.03
                    Case Else
                        WImpoIb = 0
                End Select
                Call Redondeo(WImpoIb)
            Case Else
                WImpoIb = 0
        End Select
        
    End If
    
    Select Case WCodIbTucu
        Case 1
            WImpoIbTucu = WNeto * 0.0125
            Call Redondeo(WImpoIbTucu)
        Case Else
            WImpoIbTucu = 0
    End Select
    
    Select Case WCodIbCiudad
        Case 1
            WImpoIbCiudad = WNeto * 0.015
            Call Redondeo(WImpoIbCiudad)
        Case 2
            WImpoIbCiudad = WNeto * 0.03
            Call Redondeo(WImpoIbCiudad)
        Case Else
            WImpoIbCiudad = 0
    End Select
    
    Compara = WNeto
    Call Redondeo(Compara)
    If Compara < 100 Then
        WImpoIb = 0
    End If
    If Compara < 500 Then
        WImpoIbCiudad = 0
    End If
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.21
            WIva2 = WNeto * 0.105
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Case 3, 4, 5
            WIva1 = 0
            WIva2 = 0
        Case Else
            WIva1 = WNeto * 0.21
            Call Redondeo(WIva1)
    End Select
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoIb <> 0 Then
        Call Convierte1_datos(Str$(WImpoIb), Auxi)
        ImpoIb.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIb.Caption = "0.00"
    End If
    
    If WImpoIbTucu <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbTucu), Auxi)
        ImpoIbTucu.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbTucu.Caption = "0.00"
    End If
    
    If WImpoIbCiudad <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbCiudad), Auxi)
        ImpoIbCiudad.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbCiudad.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2 + WImpoIb + WImpoIbTucu + WImpoIbCiudad
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    PrgFactup.Hide
    Unload Me
    Menu.Show
    
End Sub


Private Sub ConsultaPedido_Click()
    ZZProcesoFactura = 2
    PrgSeleccionaPedido.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    If Val(WEmpresa) = 1 Then
        OPEN_FILE_Ctacte8
        OPEN_FILE_Numero8
        OPEN_FILE_Esta8
    End If
    
    If ZZProcesoFactura = 99 And Val(Pedido.Text) <> 0 Then
        Call Pedido_KeyPress(13)
        Call Fecha_Keypress(13)
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        Remito.SetFocus
    End If
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    Call Verifica_Lote
    If WEstado = "N" Then
        Call Limpia_Click
        Exit Sub
    End If
    
    If Tipoventa.ListIndex = 1 Then
    
        spConsig = "ListaConsig " + "'" + Remito.Text + "'"
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount = 0 Then
            m$ = "No Existe el Remito de mercaderia en Consignacion Especificado"
            a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
            Exit Sub
                Else
            If Cliente.Text <> rstConsig!Cliente Then
                m$ = "No coincide el cliente informado con el especificado en el remito"
                a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                Exit Sub
            End If
            rstConsig.Close
        End If
        
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WBase = Val(Right$(Articulo, 3))
                WBaseDy = Val(Left$(Articulo, 2))
                Rem If WBase <= 5 And WBaseDy = "PT" Then
                Rem     Articulo = Left$(Articulo, 7) + "100"
                Rem End If
                
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                    XParam = "'" + Remito.Text + "','" _
                            + Articulo + "'"
                    spConsig = "ListaConsigFactura " + XParam
                    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConsig.RecordCount > 0 Then
                        WSaldo = rstConsig!Cantidad - rstConsig!Facturado
                        If Cantidad > WSaldo Then
                            m$ = "Cantidad insuficiente en consignacion Articulo " + Articulo + " Saldo : " + Str$(WSaldo)
                            a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                            Exit Sub
                        End If
                        rstConsig.Close
                            Else
                        m$ = "No existe este producto en consignacion Articulo " + Articulo
                        a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                        Exit Sub
                    End If
                End If
                                        
            Next iRow
        Next a
    End If
    
    If Tipoventa.ListIndex = 0 Then
    
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                WRow = iRow
                DBGrid1.Row = WRow
                
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                    DBGrid1.Col = 6
                    Rem If DBGrid1.Text <> "S" Then
                    Rem     m$ = "No asigno las partidas a todos los productos"
                    Rem     A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    Rem     DBGrid1.Refresh
                    Rem     Exit Sub
                    Rem End If
                End If
                
            Next iRow
        Next a
    End If
    
        Call Calcula_Click

        Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
        Rem     WImporte = WNeto
        Rem     WNeto = WNeto / 1.21
        Rem     Call Redondeo(WNeto)
        Rem     WIva1 = WImporte - WNeto
        Rem     WIva2 = 0
        Rem End If
        
        WTipo = "01"
        WNumero = Numero.Text
        WRenglon = "01"
        WCliente = Cliente.Text
        WFecha = Fecha.Text
        WEstado = "0"
        Rem Wvencimiento = Wvencimiento
        Rem WVencimiento1 = WVencimiento1
        Call Convierte_datos(Str$(Total), Auxi)
        XTotalUs = Str$(WTotal / Val(Paridad.Text))
        XTotal = Str$(WTotal)
        XSaldoUs = Str$(WTotal / Val(Paridad.Text))
        XSaldo = Str$(WTotal)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WImpre = "FC"
        XNet = Str$(WNeto)
        XIva1 = Str$(WIva1)
        XIva2 = Str$(WIva2)
        XImpoIb = Str$(WImpoIb)
        XImpoIbTucu = Str$(WImpoIbTucu)
        XImpoIbCiudad = Str$(WImpoIbCiudad)
        XSeguro = ""
        XFlete = ""
        WPedido = Pedido.Text
        WRemito = Remito.Text
        WOrden = Orden.Text
        WParidad = Paridad.Text
        WProvincia = WProv
        XVendedor = Str$(WVendedor)
        XRubro = Str$(WRubro)
        WComprobante = ""
        WAceptada = Str$(Tipoventa.ListIndex)
        Call Ceros(WAceptada, 1)
        WCosto = ""
        WImporte1 = ""
        WImporte2 = ""
        WImporte3 = ""
        WImporte4 = ""
        WImporte5 = ""
        WImporte6 = ""
        WImporte7 = ""
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        WClave = "01" + Auxi + "01"
        XEmpresa = "1"
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                    + WTipo + "','" + WNumero + "','" _
                    + WRenglon + "','" + WCliente + "','" _
                    + WFecha + "','" + WEstado + "','" _
                    + Wvencimiento + "','" + WVencimiento1 + "','" _
                    + XTotal + "','" + XTotalUs + "','" _
                    + XSaldo + "','" + XSaldoUs + "','" _
                    + WOrdFecha + "','" + WOrdVencimiento + "','" _
                    + WOrdVencimiento1 + "','" + WImpre + "','" _
                    + XEmpresa + "','" _
                    + XNet + "','" + XIva1 + "','" _
                    + XIva2 + "','" + WPedido + "','" _
                    + WRemito + "','" + WOrden + "','" _
                    + WParidad + "','" + WProvincia + "','" _
                    + XVendedor + "','" + XRubro + "','" _
                    + WComprobante + "','" + WAceptada + "','" _
                    + WCosto + "','" _
                    + WImporte1 + "','" + WImporte2 + "','" _
                    + WImporte3 + "','" + WImporte4 + "','" _
                    + WImporte5 + "','" + WImporte6 + "','" _
                    + WImporte7 + "','" + WDate + "','" _
                    + XSeguro + "','" + XFlete + "','" _
                    + XImpoIb + "'"
                        
        spCtacte = "AltaCtacte " + XParam
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + " ImpoIbTucu = " + "'" + XImpoIbTucu + "',"
        ZSql = ZSql + " ImpoIbCiudad = " + "'" + XImpoIbCiudad + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                     
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        
        If WAdicional > 0 Then
            If Val(WEmpresa) = 8 Then
                OPEN_FILE_Ctacte8
                OPEN_FILE_Numero8
                OPEN_FILE_Esta8
            End If
        End If
        
        If WAdicional > 0 Then
        
             With rstNumero8
                .Index = "Codigo"
                Claveven$ = "01"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WNumero8 = Str$(!Numero + 1)
                        Else
                    WNumero8 = "1"
                End If
            End With
            
            With rstNumero8
                .Index = "Codigo"
                .Seek "=", "01"
                If .NoMatch = False Then
                    .Edit
                    !Numero = Val(WNumero8)
                    .Update
                End If
            End With
            
            With rstCtacte8
                .Index = "Clave"
                .AddNew
                !Tipo = "01"
                !Numero = WNumero8
                !Renglon = "00"
                !Cliente = Cliente.Text
                !Fecha = Fecha.Text
                !Estado = "0"
                !Vencimiento = "  /  /    "
                !Vencimiento1 = "  /  /    "
                Call Convierte_datos(Str$(Total), Auxi)
                !Total = (WNeto * WAdicional)
                !Totalus = (WNeto * WAdicional) / Val(Paridad.Text)
                !Saldo = (WNeto * WAdicional)
                !Saldous = (WNeto * WAdicional) / Val(Paridad.Text)
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = "00000000"
                !OrdVencimiento1 = "00000000"
                !Impre = "FC"
                !Neto = (WNeto * WAdicional)
                !Iva1 = 0
                !Iva2 = 0
                !Pedido = 0
                !Remito = 0
                !Orden = ""
                !Paridad = Val(Paridad.Text)
                !Provincia = WProv
                !Vendedor = WVendedor
                !Rubro = WRubro
                !Comprobante = ""
                !Aceptada = ""
                !Costo = 0
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
                !Importe4 = 0
                !Importe5 = 0
                !Importe6 = 0
                !Importe7 = 0
                Auxi = WNumero8
                Call Ceros(Auxi, 8)
                !Clave = "01" + Auxi + "00"
                !WDate = Date$
                !TipoDescarga = 1
                .Update
            End With
            
        End If
        
        Erase Auxiliar
        Erase RestaPedido
        Erase VectorCosto
        Auxi = 0
        
        Suma = 0
        Renglon = 0
        Renglon1 = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WTipoProDy = Left$(Articulo, 2)
                Rem WBase = Val(Right$(Articulo, 3))
                Rem If WBase <= 5 Then
                Rem     Articulo = Left$(Articulo, 7) + "100"
                Rem End If
                    
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
                
                Rem If WDescuento <> 0 Then
                Rem     XImpoDto = Precio * WDescuento / 100
                Rem     Call Redondeo(XImpoDto)
                Rem     Precio = Precio - XImpoDto
                Rem End If
                    
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 5
                RestaCantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    If WTipoProDy = "PT" Then
                    
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                        
                            Else
                            
                        If WTipoProDy = "DY" Then
                            WLinea = 16
                                Else
                            If WTipoProDy = "DS" Then
                                WLinea = 16
                                    Else
                                If WTipoProDy = "DW" Then
                                    WLinea = 17
                                        Else
                                    If WTipoProDy = "DQ" Then
                                        WLinea = 22
                                            Else
                                        WLinea = 5
                                    End If
                                End If
                            End If
                        End If
                        
                        
                    End If
                        
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                            
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                    WTipo = "01"
                    WNumero = Numero.Text
                    XRenglon = Str$(Renglon)
                    WArticulo = Articulo
                    XXCantidad = Str$(Cantidad)
                    XPrecioUs = Str$(Precio / Val(Paridad.Text))
                    XPrecio = Str$(Precio)
                    XImporteUs = Str$((Precio * Cantidad) / Val(Paridad.Text))
                    XImporte = Str$(Precio * Cantidad)
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = ""
                    XCosto1 = ""
                    WCoeficiente = ""
                    WPedido = Pedido.Text
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = Left$(Articulo, 8)
                    If Tipoventa.ListIndex = 1 Then
                        WRemito = "C" + Remito.Text
                            Else
                        WRemito = Remito.Text
                    End If
                    WClave = "01" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    
                    XMarca = ""
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                XMarca = ""
                            Case Else
                                XMarca = "X"
                        End Select
                    End If
                    
                    WLote1 = XLote(Suma, 1)
                    WLote2 = XLote(Suma, 3)
                    Wlote3 = XLote(Suma, 5)
                    WLote4 = XLote(Suma, 7)
                    WLote5 = XLote(Suma, 9)
                    WImpo = XLote(Suma, 2)
                    WCanti1 = Str$(WImpo)
                    WImpo = XLote(Suma, 4)
                    WCanti2 = Str$(WImpo)
                    WImpo = XLote(Suma, 6)
                    WCanti3 = Str$(WImpo)
                    WImpo = XLote(Suma, 8)
                    WCanti4 = Str$(WImpo)
                    WImpo = XLote(Suma, 10)
                    WCanti5 = WImpo
                    
                    XEnv1 = XLote(Suma, 11)
                    XCantiEnv1 = XLote(Suma, 12)
                    XEnv2 = XLote(Suma, 13)
                    XCantiEnv2 = XLote(Suma, 14)
                    XEnv3 = XLote(Suma, 15)
                    XCantiEnv3 = XLote(Suma, 16)
                    XEnv4 = XLote(Suma, 17)
                    XCantiEnv4 = XLote(Suma, 18)
                    XEnv5 = XLote(Suma, 19)
                    XCantiEnv5 = XLote(Suma, 20)
                    
                    If WCliente = "G00007" And WArticulo = "PT-07581-100" Then
                        XLinea = "18"
                    End If
                    If WCliente = "G00065" And WArticulo = "PT-07581-100" Then
                        XLinea = "18"
                    End If
                    If WTipoProDy <> "PT" Then
                        XTipoproDy = "M"
                        XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                            Else
                        XTipoproDy = "T"
                        XArticuloDy = "  -   -   "
                    End If
                    XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + XRenglon + "','" + WArticulo + "','" _
                                + XXCantidad + "','" + XPrecio + "','" _
                                + XPrecioUs + "','" + XImporte + "','" _
                                + XImporteUs + "','" + WCliente + "','" _
                                + WParidad + "','" + XVendedor + "','" _
                                + XRubro + "','" + XLinea + "','" _
                                + XCosto1 + "','" + XCosto2 + "','" _
                                + WCoeficiente + "','" + WPedido + "','" _
                                + WFecha + "','" + WImporte1 + "','" _
                                + WImporte2 + "','" + WImporte3 + "','" _
                                + WImporte4 + "','" + WOrdFecha + "','" _
                                + XArticulo + "','" + WRemito + "','" _
                                + WDate + "','" + XCanti + "','" _
                                + XImpo + "','" _
                                + XImpoUs + "','" _
                                + XMarca + "','" _
                                + WLote1 + "','" + WCanti1 + "','" _
                                + WLote2 + "','" + WCanti2 + "','" _
                                + Wlote3 + "','" + WCanti3 + "','" _
                                + WLote4 + "','" + WCanti4 + "','" _
                                + WLote5 + "','" + WCanti5 + "','" _
                                + XTipoproDy + "','" + XArticuloDy + "'"
                    
                    spEstadistica = "AltaEstadistica " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    VectorCosto(Renglon, 1) = WArticulo
                    VectorCosto(Renglon, 2) = WClave
                    
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                        Select Case WTipoPedido
                            Case "FA", "PT", "BI", "TA"
                            
                                XEmpresa = WEmpresa
                                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                    Select Case WTipoPedido
                                        Case "PG", "CO"
                                            WEmpresa = "0001"
                                            txtOdbc = "Empresa01"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "FA"
                                            WEmpresa = "0005"
                                            txtOdbc = "Empresa05"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "TA"
                                            WEmpresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            WEmpresa = "0007"
                                            txtOdbc = "Empresa07"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                End If
                            
                                XMarca = ""
                                XParam = "'" + WClave + "','" _
                                            + WTipo + "','" + WNumero + "','" _
                                            + XRenglon + "','" + WArticulo + "','" _
                                            + XXCantidad + "','" + XPrecio + "','" _
                                            + XPrecioUs + "','" + XImporte + "','" _
                                            + XImporteUs + "','" + WCliente + "','" _
                                            + WParidad + "','" + XVendedor + "','" _
                                            + XRubro + "','" + XLinea + "','" _
                                            + XCosto1 + "','" + XCosto2 + "','" _
                                            + WCoeficiente + "','" + WPedido + "','" _
                                            + WFecha + "','" + WImporte1 + "','" _
                                            + WImporte2 + "','" + WImporte3 + "','" _
                                            + WImporte4 + "','" + WOrdFecha + "','" _
                                            + XArticulo + "','" + WRemito + "','" _
                                            + WDate + "','" + XCanti + "','" _
                                            + XImpo + "','" _
                                            + XImpoUs + "','" _
                                            + XMarca + "','" _
                                            + WLote1 + "','" + WCanti1 + "','" _
                                            + WLote2 + "','" + WCanti2 + "','" _
                                            + Wlote3 + "','" + WCanti3 + "','" _
                                            + WLote4 + "','" + WCanti4 + "','" _
                                            + WLote5 + "','" + WCanti5 + "','" _
                                            + XTipoproDy + "','" + XArticuloDy + "'"
                        
                                spEstadistica = "AltaEstadistica " + XParam
                                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                                
                                Call Conecta_Empresa
                            
                            Case Else
                        End Select
                    End If
                    
                    If WAdicional > 0 Then
                        Auxi1 = Str$(WNumero8)
                        Call Ceros(Auxi1, 8)
                        With rstEsta8
                            .Index = "Clave"
                            .AddNew
                            !Tipo = "01"
                            !Numero = WNumero8
                            !Renglon = Renglon
                            !Articulo = Articulo
                            !Cantidad = Cantidad
                            !Precio = Precio * WAdicional
                            !PrecioUs = Precio * WAdicional / Val(Paridad.Text)
                            !Importe = Precio * Cantidad * WAdicional
                            !ImporteUs = Precio * Cantidad * WAdicional / Val(Paridad.Text)
                            !Cliente = Cliente.Text
                            !Paridad = Val(Paridad.Text)
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Linea = WLinea
                            !Costo1 = 0
                            !Costo2 = 0
                            !Coeficiente = 0
                            !Pedido = 0
                            !Fecha = Fecha.Text
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            !WArticulo = Left$(Articulo, 8)
                            !Remito = ""
                            !Clave = "01" + Auxi1 + Auxi
                            !WDate = Date$
                            !TipoDescarga = 1
                            !lote1 = 0
                            !lote2 = 0
                            !lote3 = 0
                            !lote4 = 0
                            !lote5 = 0
                            !Canti1 = 0
                            !Canti2 = 0
                            !Canti3 = 0
                            !Canti4 = 0
                            !Canti5 = 0
                            .Update
                        End With
                    End If
                    
                    Auxiliar(Renglon, 1) = Articulo
                    Auxiliar(Renglon, 2) = Cantidad
                    Auxiliar(Renglon, 3) = Precio
                    Auxiliar(Renglon, 4) = WRenglon
                    Auxiliar(Renglon, 5) = WLote1
                    Auxiliar(Renglon, 6) = WCanti1
                    Auxiliar(Renglon, 7) = WLote2
                    Auxiliar(Renglon, 8) = WCanti2
                    Auxiliar(Renglon, 9) = Wlote3
                    Auxiliar(Renglon, 10) = WCanti3
                    Auxiliar(Renglon, 11) = WLote4
                    Auxiliar(Renglon, 12) = WCanti4
                    Auxiliar(Renglon, 13) = WLote5
                    Auxiliar(Renglon, 14) = WCanti5
                    Auxiliar(Renglon, 15) = RestaCantidad
                        
                End If
                
                If RestaCantidad <> 0 Then
                    Renglon1 = Renglon1 + 1
                    RestaPedido(Renglon1, 1) = Articulo
                    RestaPedido(Renglon1, 2) = RestaCantidad
                    RestaPedido(Renglon1, 3) = ClavePedido(WRenglon)
                End If
                                        
            Next iRow
            
        Next a
        
        For DA = 1 To Renglon
        
            Articulo = Auxiliar(DA, 1)
            Cantidad = Auxiliar(DA, 2)
            Precio = Auxiliar(DA, 3)
            WRenglon = Auxiliar(DA, 4)
            WLote1 = Auxiliar(DA, 5)
            WCanti1 = Auxiliar(DA, 6)
            WLote2 = Auxiliar(DA, 7)
            WCanti2 = Auxiliar(DA, 8)
            Wlote3 = Auxiliar(DA, 9)
            WCanti3 = Auxiliar(DA, 10)
            WLote4 = Auxiliar(DA, 11)
            WCanti4 = Auxiliar(DA, 12)
            WLote5 = Auxiliar(DA, 13)
            WCanti5 = Auxiliar(DA, 14)
            RestaCantidad = Auxiliar(DA, 15)
            WTipoProDy = Left$(Articulo, 2)
            If WTipoProDy <> "PT" Then
                XTipoproDy = "M"
                XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                    Else
                XTipoproDy = "T"
                XArticuloDy = "  -   -   "
            End If
            
            Select Case Tipoventa.ListIndex
                Case 1
                    XParam = "'" + Remito.Text + "','" _
                            + Articulo + "'"
                    spConsig = "ListaConsigFactura " + XParam
                    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConsig.RecordCount > 0 Then
                        WClave = rstConsig!Clave
                        WFacturado = Str$(rstConsig!Facturado + Cantidad)
                        rstConsig.Close
                
                        XParam = "'" + WClave + "','" _
                                + WFacturado + "'"
                                           
                        spConsig = "ModificaConsigFacturado " + XParam
                        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    End If
            
                Case Else
                    If XTipoproDy = "M" Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
                        spArticulo = "ConsultaArticulo " + "'" + XArticuloDy + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WCodigo = XArticuloDy
                            WPedido = Str$(rstArticulo!Venta - RestaCantidad)
                            WSalidas = Str$(rstArticulo!Salidas + Cantidad)
                            WDate = Date$
                            rstArticulo.Close
                            XParam = "'" + WCodigo + "','" _
                                    + WPedido + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                            spArticulo = "ModificaArticuloFacturas " + XParam
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        BajaLote(1, 1) = WLote1
                        BajaLote(1, 2) = WCanti1
                        BajaLote(2, 1) = WLote2
                        BajaLote(2, 2) = WCanti2
                        BajaLote(3, 1) = Wlote3
                        BajaLote(3, 2) = WCanti3
                        BajaLote(4, 1) = WLote4
                        BajaLote(4, 2) = WCanti4
                        BajaLote(5, 1) = WLote5
                        BajaLote(5, 2) = WCanti5
                        
                        For XDa = 1 To 5
                        
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                            
                            If Val(lote1) <> 0 Then
                        
                                XParam = "'" + lote1 + "','" _
                                        + XArticuloDy + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad1))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + WClave + "','" _
                                                 + WDate + "','" _
                                                 + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                            Else
                                
                                    XParam = "'" + XArticuloDy + "','" _
                                                 + lote1 + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad1))
                                        WDate = Date$
                                        rstMovguia.Close
                                
                                        XParam = "'" + WClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                            
                                End If
                            End If
                        Next XDa
                    
                        Call Conecta_Empresa
                        
                            Else
                            
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                            
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WCodigo = Articulo
                            WPedido = Str$(rstTerminado!Pedido - RestaCantidad)
                            WSalidas = Str$(rstTerminado!Salidas + Cantidad)
                            WDate = Date$
                        
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                    
                            XParam = "'" + WCodigo + "','" _
                                         + WPedido + "','" _
                                         + WSalidas + "','" _
                                         + WDate + "'"
                                            
                            spTerminado = "ModificaTerminadoFacturas " + XParam
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        BajaLote(1, 1) = WLote1
                        BajaLote(1, 2) = WCanti1
                        BajaLote(2, 1) = WLote2
                        BajaLote(2, 2) = WCanti2
                        BajaLote(3, 1) = Wlote3
                        BajaLote(3, 2) = WCanti3
                        BajaLote(4, 1) = WLote4
                        BajaLote(4, 2) = WCanti4
                        BajaLote(5, 1) = WLote5
                        BajaLote(5, 2) = WCanti5
                        
                        For XDa = 1 To 5
                    
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                
                            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                rstTerminado.Close
                            End If
                            
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                            
                            If WControla = 0 And Val(lote1) <> 0 Then
                                XParam = "'" + lote1 + "','" _
                                             + Articulo + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                
                                    WClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo - Val(Cantidad1))
                                    WDate = Date$
                                    rstHoja.Close
                                    
                                    XParam = "'" + WClave + "','" _
                                                 + WDate + "','" _
                                                 + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                
                                        Else
                                    
                                    XParam = "'" + Articulo + "','" _
                                                 + lote1 + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad1))
                                        WDate = Date$
                                        rstMovguia.Close
                                    
                                        XParam = "'" + WClave + "','" _
                                                     + WDate + "','" _
                                                     + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                
                                End If
                            End If
                    
                        Next XDa
                        
                        Call Conecta_Empresa

                    End If
                    
            End Select
            
            If XTipoproDy = "M" Then
            
                ClavePrecioMp = Cliente.Text + XArticuloDy
            
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePrecioMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
            
                    WFecha1 = ""
                    WFactura1 = ""
                    WPrecio1 = ""
                    WCantidad1 = ""
                
                    WFecha2 = ""
                    WFactura2 = ""
                    WPrecio2 = ""
                    WCantidad2 = ""
                
                    WFecha3 = ""
                    WFactura3 = ""
                    WPrecio3 = ""
                    WCantidad3 = ""
                
                    WFecha4 = ""
                    WFactura4 = ""
                    WPrecio4 = ""
                    WCantidad4 = ""
                
                    WFecha5 = ""
                    WFactura5 = ""
                    WPrecio5 = ""
                    WCantidad5 = ""
                
                    If rstPreciosMp!Cantidad2 <> O Then
                        WFecha1 = rstPreciosMp!fecha2
                        WFactura1 = rstPreciosMp!Factura2
                        WPrecio1 = Str$(rstPreciosMp!Precio2)
                        WCantidad1 = Str$(rstPreciosMp!Cantidad2)
                    End If
                                
                    If rstPreciosMp!Cantidad3 <> O Then
                        WFecha2 = rstPreciosMp!Fecha3
                        WFactura2 = rstPreciosMp!Factura3
                        WPrecio2 = Str$(rstPreciosMp!Precio3)
                        WCantidad2 = Str$(rstPreciosMp!Cantidad3)
                    End If
                                
                    If rstPreciosMp!Cantidad4 <> O Then
                        WFecha3 = rstPreciosMp!Fecha4
                        WFactura3 = rstPreciosMp!Factura4
                        WPrecio3 = Str$(rstPreciosMp!Precio4)
                        WCantidad3 = Str$(rstPreciosMp!Cantidad4)
                    End If
                                
                    If rstPreciosMp!Cantidad5 <> O Then
                        WFecha4 = rstPreciosMp!Fecha5
                        WFactura4 = rstPreciosMp!Factura5
                        WPrecio4 = Str$(rstPreciosMp!Precio5)
                        WCantidad4 = Str$(rstPreciosMp!Cantidad5)
                    End If
                                
                    WFecha5 = Fecha.Text
                    WFactura5 = Numero.Text
                    WPrecio5 = Str$(Precio / Val(Paridad.Text))
                    WCantidad5 = Str$(Cantidad)
                                
                    WDate = Date$
                
                    rstPreciosMp.Close
                
                    XParam = "'" + ClavePrecioMp + "','" _
                            + WFecha1 + "','" _
                            + WFactura1 + "','" _
                            + WPrecio1 + "','" _
                            + WCantidad1 + "','" _
                            + WFecha2 + "','" _
                            + WFactura2 + "','" _
                            + WPrecio2 + "','" _
                            + WCantidad2 + "','" _
                            + WFecha3 + "','" _
                            + WFactura3 + "','" _
                            + WPrecio3 + "','" _
                            + WCantidad3 + "','" _
                            + WFecha4 + "','" _
                            + WFactura4 + "','" _
                            + WPrecio4 + "','" _
                            + WCantidad4 + "','" _
                            + WFecha5 + "','" _
                            + WFactura5 + "','" _
                            + WPrecio5 + "','" _
                            + WCantidad5 + "','" _
                            + WDate + "'"
                                           
                    spPreciosMp = "ModificaPreciosFacturaMp " + XParam
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                    Else
                
                ClavePrecio = Cliente.Text + Articulo
            
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecio + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
            
                    WFecha1 = ""
                    WFactura1 = ""
                    WPrecio1 = ""
                    WCantidad1 = ""
                
                    WFecha2 = ""
                    WFactura2 = ""
                    WPrecio2 = ""
                    WCantidad2 = ""
                
                    WFecha3 = ""
                    WFactura3 = ""
                    WPrecio3 = ""
                    WCantidad3 = ""
                
                    WFecha4 = ""
                    WFactura4 = ""
                    WPrecio4 = ""
                    WCantidad4 = ""
                
                    WFecha5 = ""
                    WFactura5 = ""
                    WPrecio5 = ""
                    WCantidad5 = ""
                
                    If rstPrecios!Cantidad2 <> O Then
                        WFecha1 = rstPrecios!fecha2
                        WFactura1 = rstPrecios!Factura2
                        WPrecio1 = Str$(rstPrecios!Precio2)
                        WCantidad1 = Str$(rstPrecios!Cantidad2)
                    End If
                                
                    If rstPrecios!Cantidad3 <> O Then
                        WFecha2 = rstPrecios!Fecha3
                        WFactura2 = rstPrecios!Factura3
                        WPrecio2 = Str$(rstPrecios!Precio3)
                        WCantidad2 = Str$(rstPrecios!Cantidad3)
                    End If
                                
                    If rstPrecios!Cantidad4 <> O Then
                        WFecha3 = rstPrecios!Fecha4
                        WFactura3 = rstPrecios!Factura4
                        WPrecio3 = Str$(rstPrecios!Precio4)
                        WCantidad3 = Str$(rstPrecios!Cantidad4)
                    End If
                                
                    If rstPrecios!Cantidad5 <> O Then
                        WFecha4 = rstPrecios!Fecha5
                        WFactura4 = rstPrecios!Factura5
                        WPrecio4 = Str$(rstPrecios!Precio5)
                        WCantidad4 = Str$(rstPrecios!Cantidad5)
                    End If
                                
                    WFecha5 = Fecha.Text
                    WFactura5 = Numero.Text
                    WPrecio5 = Str$(Precio)
                    WCantidad5 = Str$(Cantidad)
                                
                    WDate = Date$
                
                    rstPrecios.Close
                
                    XParam = "'" + ClavePrecio + "','" _
                            + WFecha1 + "','" _
                            + WFactura1 + "','" _
                            + WPrecio1 + "','" _
                            + WCantidad1 + "','" _
                            + WFecha2 + "','" _
                            + WFactura2 + "','" _
                            + WPrecio2 + "','" _
                            + WCantidad2 + "','" _
                            + WFecha3 + "','" _
                            + WFactura3 + "','" _
                            + WPrecio3 + "','" _
                            + WCantidad3 + "','" _
                            + WFecha4 + "','" _
                            + WFactura4 + "','" _
                            + WPrecio4 + "','" _
                            + WCantidad4 + "','" _
                            + WFecha5 + "','" _
                            + WFactura5 + "','" _
                            + WPrecio5 + "','" _
                            + WCantidad5 + "','" _
                            + WDate + "'"
                                           
                    spPrecios = "ModificaPreciosFactura " + XParam
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            End If
        Next DA
        
        For DA = 1 To Renglon1
        
            Articulo = RestaPedido(DA, 1)
            Cantidad = RestaPedido(DA, 2)
            WClavePedido = RestaPedido(DA, 3)
            
            XParam = "'" + Left$(WClavePedido, 6) + "','" _
                        + Right$(WClavePedido, 2) + "'"
            spPedido = "ConsultaPedido2 " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                WFacturado = Str$(rstPedido!Facturado + Cantidad)
                If Val(WFacturado) > rstPedido!Cantidad Then
                    WFacturado = Str$(rstPedido!Cantidad)
                End If
                WClavePedido = rstPedido!Clave
                rstPedido.Close
                XParam = "'" + WClavePedido + "','" _
                            + WFacturado + "'"
                                           
                spPedido = "ModificaPedidoFacturas " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " Remito = " + "'" + Remito.Text + "',"
            ZSql = ZSql + " CantidadFac = " + "'" + Cantidad + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClavePedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        Next DA
        
        BajaImpre = "N"
        
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Canti = !Cantidad - !Facturado
                        
                        If Canti > 0 Then
                            BajaImpre = "S"
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
        
        If BajaImpre = "S" Then
            spPedido = "ModificaPedidoVersion " + "'" + Pedido.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        ZSql = ""
        ZSql = ZSql & "UPDATE Pedido SET "
        ZSql = ZSql & "MarcaFactura = " + "'" + "0" + "'"
        ZSql = ZSql & " Where Pedido = " + "'" + Pedido.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        spNumero = "ConsultaNumero " + "'" + "01" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            WCodigo = "01"
            WNumero = Numero.Text
            rstNumero.Close
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
        Rem Listado.Destination = 1
        Rem Listado.Action = 1
        
        Call Calcula_Saldo
        
        Erase Envase
        Envase(1, 1) = Envase1.Text
        Envase(2, 1) = Envase2.Text
        Envase(3, 1) = Envase3.Text
        Envase(4, 1) = Envase4.Text
        Envase(5, 1) = Envase5.Text
        
        Envase(1, 2) = Canti1.Text
        Envase(2, 2) = Canti2.Text
        Envase(3, 2) = Canti3.Text
        Envase(4, 2) = Canti4.Text
        Envase(5, 2) = Canti5.Text
        
        For XDa = 1 To 5
            For DA = 1 To 9
                If Val(Envase(XDa, 1)) = Val(Stk(DA, 1)) Then
                    Stk(DA, 3) = Val(Envase(XDa, 2))
                End If
            Next DA
        Next XDa
        
        For DA = 1 To 9
            Stk(DA, 4) = Str$(Val(Stk(DA, 2)) + Val(Stk(DA, 3)))
        Next DA
        
        Renglon = 0
        
        For DA = 1 To 5
        
            If Val(Envase(DA, 2)) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Remito.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "1"
                WCodigo = Str$(Val(Remito.Text) + 200000)
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnvase = Envase(DA, 1)
                WCantidad = Envase(DA, 2)
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnvase + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next DA
        
        For Ciclo = 1 To 100
        
            If VectorCosto(Ciclo, 1) <> "" Then
            
                ZZZProducto = VectorCosto(Ciclo, 1)
                ZZClave = VectorCosto(Ciclo, 2)
                
                ZZZCosto = 0
                Call Calcula_CostoFactura(ZZZProducto, ZZZCosto)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Estadistica SET "
                ZSql = ZSql + " Costo1 = " + "'" + Str$(ZZZCosto) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next Ciclo
        
        Call Impresion
        If Tipoventa.ListIndex <> 1 Then
            Call Impresion_Remito
        End If
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
    Exit Sub

WError:
     Resume Next
        
End Sub

Private Sub Verifica_Fecha_Entrega()

    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZTipoPedido = rstPedido!Tipoped
        ZFecha = rstPedido!Fecha
        ZFechaEntrega = rstPedido!FecEntrega
        ZOrdFechaEntrega = rstPedido!OrdFecEntrega
        ZFechaActualizacion = IIf(IsNull(rstPedido!FechaActualizacion), "", rstPedido!FechaActualizacion)
        ZOrdFechaActualizacion = IIf(IsNull(rstPedido!OrdFechaActualizacion), "", rstPedido!OrdFechaActualizacion)
        rstPedido.Close
    End If
    
    If ZTipoPedido = 4 Then
        If ZFechaActualizacion <> "" Then
            ZFechaFactu = ZFechaActualizacion
            ZFechaFactuOrd = ZOrdFechaActualizacion
                Else
            ZFechaFactu = Fecha.Text
            ZFechaFactuOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        End If
                Else
        ZFechaFactu = Fecha.Text
        ZFechaFactuOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    End If
        
    WDias = 0
    WSuma2 = "0"
    WFechaHastaOrd = ZFechaFactuOrd
    WFechaDesdeOrd = ZOrdFechaEntrega
    WFechaHasta = ZFechaFactu
    WFechaDesde = ZFechaEntrega
            
    If WFechaHastaOrd > WFechaDesdeOrd Then
            
        WSuma2 = "1"
            
        Do
        
            Feriado = "N"
            For Cicla = 1 To TotalFeriado
                If DiaFeriado(Cicla) = WFechaDesde Then
                    Feriado = "S"
                    Exit For
                End If
            Next Cicla
                    
            Rem 1 - DOMINGO
            Rem 2 - LUNES
            Rem 3 - MARTES
            Rem 4 - MIERCOLES
            Rem 5 - JUEVES
            Rem 6 - VIERNES
            Rem 7 - SABADO
            XFec1 = WFechaDesde
            strDia = Format$(XFec1, "dddd")
            BDia = Format(XFec1, "w")
            If BDia = 1 Or BDia = 7 Then
                Feriado = "S"
            End If
            
            If Feriado = "N" Then
                WDias = WDias + 1
            End If
            SumaDia = 2
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            WFechaDesde = XFec2
                        
            If WFechaDesde = WFechaHasta Then
                Exit Do
            End If
        
        Loop
        
    End If
    
    Fecha.SetFocus
    
    If WDias > 0 Then
    
        ZVersionAtraso = 0
        ZVersionPedido = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Atraso"
        ZSql = ZSql + " Where Atraso.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " Order by Atraso.Numero"
        spAtraso = ZSql
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstAtraso.RecordCount > 0 Then
            With rstAtraso
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        ZOrigen = IIf(IsNull(!Origen), "0", !Origen)
                        If ZOrigen = 0 Then
                            ZVersionAtraso = IIf(IsNull(!VersionPedido), "0", !VersionPedido)
                        End If
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstAtraso.Close
        End If
        
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZVersionPedido = rstPedido!Version
            rstPedido.Close
        End If
            
        If ZVersionPedido <> ZVersionAtraso Then
            ConceptoAtraso.ListIndex = 0
            DescriMotivo.Text = ""
            PantaMotivo.Visible = True
            DescriMotivo.SetFocus
        End If
        
    End If

End Sub

Private Sub DescriMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(DescriMotivo.Text)) >= 10 And ConceptoAtraso.ListIndex > 0 Then
        
            ZZAtraso = "1"
        
            Sql1 = "Select Max(Numero) as [NumeroMayor]"
            Sql2 = " FROM Atraso"
            spAtraso = Sql1 + Sql2
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                ZZAtraso = Str$(rstAtraso!Numeromayor + 1)
                rstAtraso.Close
            End If
    
            ZFecha = Fecha.Text
            ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZFechaEntrega = Fecha.Text
            ZFechaEntregaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZTerminado = "  -     -   "
            ZArticulo = "  -   -   "
            ZDesTerminado = ""
            ZDesArticulo = ""
            ZConcepto = Str$(ConceptoAtraso.ListIndex + 4)
            ZSolicitud = ""
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Atraso ("
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Problema ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "FechaEntrega ,"
            ZSql = ZSql + "OrdFechaEntrega ,"
            ZSql = ZSql + "DesCliente ,"
            ZSql = ZSql + "DesTerminado ,"
            ZSql = ZSql + "DesArticulo ,"
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Solicitud ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "VersionPedido)"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZAtraso + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZFechaOrd + "',"
            ZSql = ZSql + "'" + Pedido.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + DescriMotivo.Text + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZFechaEntrega + "',"
            ZSql = ZSql + "'" + ZFechaEntregaOrd + "',"
            ZSql = ZSql + "'" + Left$(DesCliente.Caption, 50) + "',"
            ZSql = ZSql + "'" + ZDesTerminado + "',"
            ZSql = ZSql + "'" + ZDesArticulo + "',"
            ZSql = ZSql + "'" + ZConcepto + "',"
            ZSql = ZSql + "'" + ZSolicitud + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + "" + "')"
           
            spAtraso = ZSql
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        
            PantaMotivo.Visible = False
            Fecha.SetFocus
        End If
    End If
End Sub



Private Sub Limpia_Click()

    CargaLote.Visible = False
    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    WAdicional = 0
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    Canti4.Text = ""
    Canti5.Text = ""
    
    Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 7
                Rem Select Case KeyCode
                Rem     Case 13
                Rem         DBGrid1.Col = 4
                Rem         DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                Rem         DBGrid1.Col = 5
                Rem         DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                Rem         DBGrid1.Col = 0
                Rem         XTerminado = DBGrid1.Text
                Rem         DBGrid1.Col = 4
                Rem         XCantidad = Val(DBGrid1.Text)
                Rem         WRow = DBGrid1.Row
                Rem
                Rem         Rem If DBGrid1.Row < 40 Then
                Rem         Rem    DBGrid1.Row = DBGrid1.Row + 1
                Rem         Rem    WRow = DBGrid1.Row
                Rem         Rem    DBGrid1.Col = 4
                Rem         Rem    KeyCode = 0
                Rem         Rem End If
                Rem         Rem Call Calcula_Click
                Rem         Rem DBGrid1.Row = WRow
                Rem
                Rem         If Tipoventa.ListIndex = 0 Then
                Rem             CargaLote.Visible = True
                Rem             WLote1.Text = ""
                Rem             WCanti1.Text = ""
                Rem             WLote2.Text = ""
                Rem             WCanti2.Text = ""
                Rem             Wlote3.Text = ""
                Rem             WCanti3.Text = ""
                Rem             WLote4.Text = ""
                Rem             WCanti4.Text = ""
                Rem             WLote5.Text = ""
                Rem             WCanti5.Text = ""
                Rem
                Rem             If Val(xLote(WRow, 1)) <> 0 Then
                Rem                 WLote1.Text = xLote(WRow, 1)
                Rem                 WCanti1.Text = xLote(WRow, 2)
                Rem             End If
                Rem             If Val(xLote(WRow, 3)) <> 0 Then
                Rem                 WLote2.Text = xLote(WRow, 3)
                Rem                 WCanti2.Text = xLote(WRow, 4)
                Rem             End If
                Rem             If Val(xLote(WRow, 5)) <> 0 Then
                Rem                 Wlote3.Text = xLote(WRow, 5)
                Rem                 WCanti3.Text = xLote(WRow, 6)
                Rem             End If
                Rem             If Val(xLote(WRow, 7)) <> 0 Then
                Rem                 WLote4.Text = xLote(WRow, 7)
                Rem                 WCanti4.Text = xLote(WRow, 6)
                Rem             End If
                Rem             If Val(xLote(WRow, 9)) <> 0 Then
                Rem                 WLote5.Text = xLote(WRow, 9)
                Rem                 WCanti5.Text = xLote(WRow, 10)
                Rem             End If
                Rem
                Rem             WLote1.SetFocus
                Rem         End If
                Rem
                Rem     Case Else
                Rem         Rem If KeyCode <> 0 Then Stop
                Rem
                Rem End Select
            
    End Select

    
End Sub

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub

Private Sub Form_Load()

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    Tipoventa.Clear
    
    Tipoventa.AddItem "Venta Normal"
    Tipoventa.AddItem "Mercaderia en Consignacion"
    
    Tipoventa.ListIndex = 0
    
    
    ConceptoAtraso.Clear
    
    ConceptoAtraso.AddItem ""
    ConceptoAtraso.AddItem "Error del Sistema"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Problemas Vehiculos"
    ConceptoAtraso.AddItem "Problemas Logistica"
    ConceptoAtraso.AddItem "Problemas Recepcion Cliente"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Corte de Luz"
    ConceptoAtraso.AddItem "Pedido por el Cliente"
    
    ConceptoAtraso.ListIndex = 0

    

    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 6, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cant. Entregar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Cant. Descontar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    DBGrid1.Font.Bold = True

    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    WAdicional = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                Else
        Paridad.Text = ""
    End If
    
    Numero.SetFocus
     
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    XEntra = "S"
    
    Erase Auxiliar
    Erase ClavePedido
    Erase XLote
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad - !Facturado
                    
                    If Canti <> 0 Then
                    
                        XCanti1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                        XCanti2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                        
                        If XCanti1 <> 0 Or XCanti2 <> 0 Then
                
                            Renglon = Renglon + 1
                
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = !Terminado
                            Auxi1 = !Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Precio * Val(Paridad.Text)))
                
                            XCantidad1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###.##", Str$(XCantidad1))
                    
                            xCantidad2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                            DBGrid1.Col = 5
                            DBGrid1.Text = Pusing("###,###.##", Str$(xCantidad2))
                    
                            DBGrid1.Col = 6
                            DBGrid1.Text = "S"
                    
                            
                            Rem If XEntra = "S" Then
                            Rem     Envase1.Text = IIf(IsNull(!Env1), "", !Env1)
                            Rem     Envase2.Text = IIf(IsNull(!Env2), "", !Env2)
                            Rem     Envase3.Text = IIf(IsNull(!Env3), "", !Env3)
                            Rem     Envase4.Text = IIf(IsNull(!Env4), "", !Env4)
                            Rem     Envase5.Text = IIf(IsNull(!Env5), "", !Env5)
                            Rem     Canti1.Text = IIf(IsNull(!CantiEnv1), "", !CantiEnv1)
                            Rem     Canti2.Text = IIf(IsNull(!CantiEnv2), "", !CantiEnv2)
                            Rem     Canti3.Text = IIf(IsNull(!CantiEnv3), "", !CantiEnv3)
                            Rem     Canti4.Text = IIf(IsNull(!CantiEnv4), "", !CantiEnv4)
                            Rem     Canti5.Text = IIf(IsNull(!CantiEnv5), "", !CantiEnv5)
                            Rem     XEntra = ""
                            Rem End If
                            
                            XLote(Renglon, 1) = IIf(IsNull(!lote1), "", !lote1)
                            XLote(Renglon, 2) = IIf(IsNull(!CantiLote1), "", !CantiLote1)
                            XLote(Renglon, 3) = IIf(IsNull(!lote2), "", !lote2)
                            XLote(Renglon, 4) = IIf(IsNull(!CantiLote2), "", !CantiLote2)
                            XLote(Renglon, 5) = IIf(IsNull(!lote3), "", !lote3)
                            XLote(Renglon, 6) = IIf(IsNull(!CantiLote3), "", !CantiLote3)
                            XLote(Renglon, 7) = IIf(IsNull(!lote4), "", !lote4)
                            XLote(Renglon, 8) = IIf(IsNull(!CantiLote4), "", !CantiLote4)
                            XLote(Renglon, 9) = IIf(IsNull(!lote5), "", !lote5)
                            XLote(Renglon, 10) = IIf(IsNull(!CantiLote4), "", !CantiLote5)
                            
                            XLote(Renglon, 11) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                            XLote(Renglon, 12) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                            XLote(Renglon, 13) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env2)
                            XLote(Renglon, 14) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv2)
                            XLote(Renglon, 15) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env3)
                            XLote(Renglon, 16) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv3)
                            XLote(Renglon, 17) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env4)
                            XLote(Renglon, 18) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv4)
                            XLote(Renglon, 19) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env5)
                            XLote(Renglon, 20) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv5)
                    
                            Auxiliar(Renglon, 1) = Auxi1
                            Auxiliar(Renglon, 2) = Canti
                            
                            ClavePedido(Renglon) = !Clave
                            
                        End If
                        
                    End If
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    
    
    
    
    
    Erase ZZEnvase
    Erase ZZCanti
    
    Canti1.Text = ""
    Envase1.Text = ""
    Canti2.Text = ""
    Envase2.Text = ""
    Canti3.Text = ""
    Envase3.Text = ""
    Canti4.Text = ""
    Envase4.Text = ""
    Canti5.Text = ""
    Envase5.Text = ""
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    ZZLugar = 0
    
    For CicloEnvase = 1 To 100
    
        Rem envase posicion 1
        If Val(XLote(CicloEnvase, 12)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 11) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 12)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 12)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 11)
        End If
        End If
        
        Rem envase posicion 2
        If Val(XLote(CicloEnvase, 14)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 13) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 14)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 14)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 13)
        End If
        End If
        
        Rem envase posicion 3
        If Val(XLote(CicloEnvase, 16)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 15) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 16)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 16)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 15)
        End If
        End If
        
        Rem envase posicion 4
        If Val(XLote(CicloEnvase, 18)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 17) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 18)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 18)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 17)
        End If
        End If
        
        Rem envase posicion 5
        If Val(XLote(CicloEnvase, 20)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 19) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 20)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 20)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 19)
        End If
        End If
        
    Next CicloEnvase
    
    Envase1.Text = ZZEnvase(1)
    Envase2.Text = ZZEnvase(2)
    Envase3.Text = ZZEnvase(3)
    Envase4.Text = ZZEnvase(4)
    Envase5.Text = ZZEnvase(5)
    
    Canti1.Text = ZZCanti(1)
    Canti2.Text = ZZCanti(2)
    Canti3.Text = ZZCanti(3)
    Canti4.Text = ZZCanti(4)
    Canti5.Text = ZZCanti(5)
    
    spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri1.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri1.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri2.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri3.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri4.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri5.Caption = ""
    End If
    
    
    
    
    
    WConpago = 0
    
    WRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                ClavePreciosMp = Cliente.Text + WArti
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                    Precio = rstPreciosMp!Precio * Val(Paridad.Text)
                
                    WConpago = IIf(IsNull(rstPreciosMp!Pago), 0, rstPreciosMp!Pago)
            
                    rstPreciosMp.Close
                End If

                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    
                    rstArticulo.Close
                End If

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
            
            Case "T"
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
            
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                    Precio = rstPrecios!Precio * Val(Paridad.Text)
                
                    WConpago = IIf(IsNull(rstPrecios!Pago), 0, rstPrecios!Pago)
            
                    rstPrecios.Close
                End If

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
        End Select
        
    Next DA
    
    If WConpago <> 0 Then
        WPago1 = WConpago
        WPago2 = WConpago
        
        spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDias1 = rstPago!Dias
            WPlazo1 = rstPago!Plazo
            WTasa = rstPago!Tasa
            WDescuento = rstPago!Descuento
            WPago = rstPago!Nombre
            rstPago.Close
        End If
        
        WFecha = Fecha.Text
        Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
        spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDias2 = rstPago!Dias
            WPlazo2 = rstPago!Plazo
            rstPago.Close
        End If
        
        Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)
        
    End If
    
    Call Calcula_Click

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub Proceso1_Click()

    WNeto = 0

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    Erase XLote
    
    
    XParam = "'" + "01" + "','" _
                + Numero.Text + "'"
    
    spEstadistica = "ConsultaEstadistica1 " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Precio)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("###,###.##", dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!Precio)
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    
                    XLote(Renglon, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    XLote(Renglon, 3) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    XLote(Renglon, 5) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    XLote(Renglon, 7) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    XLote(Renglon, 9) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    
                    XLote(Renglon, 2) = IIf(IsNull(rstEstadistica!Canti1), "", rstEstadistica!Canti1)
                    XLote(Renglon, 4) = IIf(IsNull(rstEstadistica!Canti2), "", rstEstadistica!Canti2)
                    XLote(Renglon, 6) = IIf(IsNull(rstEstadistica!Canti3), "", rstEstadistica!Canti3)
                    XLote(Renglon, 8) = IIf(IsNull(rstEstadistica!Canti4), "", rstEstadistica!Canti4)
                    XLote(Renglon, 10) = IIf(IsNull(rstEstadistica!Canti5), "", rstEstadistica!Canti5)
                    
                    XLote(Renglon, 11) = IIf(IsNull(rstEstadistica!Env1), "", rstEstadistica!Env1)
                    XLote(Renglon, 13) = IIf(IsNull(rstEstadistica!Env2), "", rstEstadistica!Env2)
                    XLote(Renglon, 15) = IIf(IsNull(rstEstadistica!Env3), "", rstEstadistica!Env3)
                    XLote(Renglon, 17) = IIf(IsNull(rstEstadistica!Env4), "", rstEstadistica!Env4)
                    XLote(Renglon, 19) = IIf(IsNull(rstEstadistica!Env5), "", rstEstadistica!Env5)
                    
                    XLote(Renglon, 12) = IIf(IsNull(rstEstadistica!CantiEnv1), "", rstEstadistica!CantiEnv1)
                    XLote(Renglon, 14) = IIf(IsNull(rstEstadistica!CantiEnv2), "", rstEstadistica!CantiEnv2)
                    XLote(Renglon, 16) = IIf(IsNull(rstEstadistica!CantiEnv3), "", rstEstadistica!CantiEnv3)
                    XLote(Renglon, 18) = IIf(IsNull(rstEstadistica!CantiEnv4), "", rstEstadistica!CantiEnv4)
                    XLote(Renglon, 20) = IIf(IsNull(rstEstadistica!CantiEnv5), "", rstEstadistica!CantiEnv5)
                    
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    
    
    
    
    
    
    
    
    Erase ZZEnvase
    Erase ZZCanti
    
    Canti1.Text = ""
    Envase1.Text = ""
    Canti2.Text = ""
    Envase2.Text = ""
    Canti3.Text = ""
    Envase3.Text = ""
    Canti4.Text = ""
    Envase4.Text = ""
    Canti5.Text = ""
    Envase5.Text = ""
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    ZZLugar = 0
    
    For CicloEnvase = 1 To 100
    
        Rem envase posicion 1
        If Val(XLote(CicloEnvase, 12)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 11) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 12)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 12)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 11)
        End If
        End If
        
        Rem envase posicion 2
        If Val(XLote(CicloEnvase, 14)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 13) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 14)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 14)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 13)
        End If
        End If
        
        Rem envase posicion 3
        If Val(XLote(CicloEnvase, 16)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 15) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 16)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 16)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 15)
        End If
        End If
        
        Rem envase posicion 4
        If Val(XLote(CicloEnvase, 18)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 17) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 18)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 18)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 17)
        End If
        End If
        
        Rem envase posicion 5
        If Val(XLote(CicloEnvase, 20)) <> 0 Then
        Entra = "S"
        For CicloEnvaseII = 1 To 5
            If ZZEnvase(CicloEnvaseII) = XLote(CicloEnvase, 19) Then
                ZZCanti(CicloEnvaseII) = Str$(Val(ZZCanti(CicloEnvaseII)) + Val(XLote(CicloEnvase, 20)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            ZZLugar = ZZLugar + 1
            ZZCanti(ZZLugar) = XLote(CicloEnvase, 20)
            ZZEnvase(ZZLugar) = XLote(CicloEnvase, 19)
        End If
        End If
        
    Next CicloEnvase
    
    Envase1.Text = ZZEnvase(1)
    Envase2.Text = ZZEnvase(2)
    Envase3.Text = ZZEnvase(3)
    Envase4.Text = ZZEnvase(4)
    Envase5.Text = ZZEnvase(5)
    
    Canti1.Text = ZZCanti(1)
    Canti2.Text = ZZCanti(2)
    Canti3.Text = ZZCanti(3)
    Canti4.Text = ZZCanti(4)
    Canti5.Text = ZZCanti(5)
    
    spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri1.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri1.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri2.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri3.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri4.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri5.Caption = ""
    End If
    
    
    
    
    
    
    
    
    
    
    
    XRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                    
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                    
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
        End Select
    Next DA
    
    Renglon = Renglon + 1
            
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
                
    DBGrid1.Col = 0
    DBGrid1.Text = ""
                            
    DBGrid1.Col = 1
    DBGrid1.Text = ""
                
    DBGrid1.Col = 2
    DBGrid1.Text = ""
                
    DBGrid1.Col = 3
    DBGrid1.Text = ""
                
    DBGrid1.Col = 4
    DBGrid1.Text = ""
                            
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    Call Calcula_FechaVto
    Call Calcula_Click

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0
    
    DBGrid1.SetFocus
    
    Graba.Enabled = False
    Borra.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Pedido.Text = rstCtacte!Pedido
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Remito.Text = rstCtacte!Remito
            Orden.Text = rstCtacte!Orden
            Paridad.Text = rstCtacte!Paridad
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!Vendedor
                WProv = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WCodIb = rstCliente!Ib
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            Call Proceso1_Click
                    Else
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Facturas")
            Rem     Numero.SetFocus
            Rem        Else
            Rem     WNumero = Numero.Text
            Rem    Rem Call Limpia_Click
            Rem    Numero.Text = WNumero
            Rem    Pedido.SetFocus
            Rem End If
            WNumero = Numero.Text
            Rem Call Limpia_Click
            Numero.Text = WNumero
            Fecha.SetFocus
                
        End If
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                Pedido.SetFocus
                
                    Else
                    
                If rstPedido!Tipoped = 5 Then
                    m$ = "EL PEDIDO ES DE MUESTRAS"
                    a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                    Pedido.SetFocus
                    Exit Sub
                End If
                    
                Cliente.Text = rstPedido!Cliente
                Orden.Text = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                
                Select Case rstPedido!TipoPedido
                    Case 1
                        WTipoPedido = "CO"
                    Case 3
                        WTipoPedido = "BI"
                    Case 4
                        WTipoPedido = "FA"
                    Case 5
                        WTipoPedido = "PG"
                    Case Else
                        WTipoPedido = "PT"
                End Select
                
                If Val(WEmpresa) = 1 And Cliente.Text = "P00005" Then
                    If Left$(rstPedido!Terminado, 4) = "PT-5" Or rstPedido!Terminado = "PT-03000-001" Then
                        WTipoPedido = "PG"
                    End If
                End If
                
                If Left$(rstPedido!Terminado, 4) = "PT-4" Then
                    WTipoPedido = "TA"
                End If
                
                rstPedido.Close
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                    WPago1 = rstCliente!Pago1
                    WPago2 = rstCliente!Pago2
                    WVendedor = rstCliente!Vendedor
                    WRubro = rstCliente!Rubro
                    WCodIva = rstCliente!Iva
                    WCodIb = rstCliente!Ib
                    WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                    WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                    WRazon = rstCliente!Razon
                    WDireccion = rstCliente!Direccion
                    WLocalidad = rstCliente!Localidad
                    WProv = rstCliente!Provincia
                    WPostal = rstCliente!Postal
                    WCuit = rstCliente!Cuit
                    WDirentrega = rstCliente!DirEntrega
                    rstCliente.Close
                End If
                
                Call Proceso_Click
                Call Calcula_FechaVto
                Call Verifica_Fecha_Entrega
                
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case WTipoPedido
                        Case "PG", "CO"
                            m$ = "Coloque Remito de Pta I"
                            a% = MsgBox(m$, 0, "Emision de facturas")
                        Case "FA"
                            m$ = "Coloque Remito de Pta III"
                            a% = MsgBox(m$, 0, "Emision de facturas")
                        Case "TA"
                            m$ = "Coloque Remito de Pta II"
                            a% = MsgBox(m$, 0, "Emision de facturas")
                        Case Else
                            m$ = "Coloque Remito de Pta V"
                            a% = MsgBox(m$, 0, "Emision de facturas")
                    End Select
                End If
                
                Remito.SetFocus
                
            End If
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
                Rem m$ = "No exsite paridad cargada para esta fecha"
                Rem a% = MsgBox(m$, 0, "Emision de facturas")
                Rem Fecha.SetFocus
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Pedido.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub reImpre_Click()
    Call Impresion
    Call Impresion_Remito
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Remito.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Orden.SetFocus
    End If
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion()

    If Val(WEmpresa) = 1 Then
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
            Else
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    End If
    
    Rem Width #1, 255

    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    Paridad = Val(Paridad.Text)
    Impotot = Val(Total.Caption) / Paridad

    For XX% = 1 To 2
    
        If XX% = 1 Then
            Print #1, ""
                Else
            Print #1, ""
        End If

        If Val(WEmpresa) = 1 Then
            Rem Print #1, ""
            Print #1, ""
        End If
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(WEmpresa) = 1 Then
            Print #1, Tab(59); Fecha.Text
                Else
            Print #1, Tab(57); Fecha.Text
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(8); WRazon
        Print #1, Tab(8); WDireccion
        Print #1, Tab(8); Left$(WLocalidad, 33);
        Print #1, Tab(55); Cliente.Text;
        Print #1, Tab(69); Orden.Text
        Print #1, Tab(8); Provincia(Val(WProv)); " ("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(8); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, ""
        Print #1, Tab(5); Left$(WPago, 40); " ";
        Print #1, Vencimiento.Text;
        Print #1, Tab(60); Remito.Text
        Print #1, ""
        Print #1, ""
        Print #1, Tab(76); "$"

        Impre = 0
        ImpreDespachoI = ""
        ImpreDespachoII = ""
        Erase ZZImpreDespaII
        ZZLugarDespaII = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(Alinea("##,###.##", DBGrid1.Text))
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    If UCase(Left$(Producto, 2)) = "DY" Then
                    
                        ZProductoDy = Left$(Producto, 3) + Right$(Producto, 7)
                    
                        For CicloLote = 1 To 5
                        
                            Select Case CicloLote
                                Case 1
                                    ZZLote = XLote(Suma, 1)
                                Case 2
                                    ZZLote = XLote(Suma, 3)
                                Case 3
                                    ZZLote = XLote(Suma, 5)
                                Case 4
                                    ZZLote = XLote(Suma, 7)
                                Case 5
                                    ZZLote = XLote(Suma, 9)
                                Case Else
                            End Select
                            
                            If Val(ZZLote) <> 0 Then
                            
                                Erase ZZImpreDespa
                                ZLugarDespa = 0
                                ZZPartiOri = ""
                                ZZCantidad = 0
                                ZZSaldo = 0
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Laudo"
                                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZProductoDy + "'"
                                ZSql = ZSql + " and Laudo.Lote = " + "'" + ZZLote + "'"
                                spLaudo = ZSql
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    ZZPartiOri = rstLaudo!PartiOri
                                    rstLaudo.Close
                                End If
                    
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Laudo"
                                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZProductoDy + "'"
                                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + ZZPartiOri + "'"
                                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                                spLaudo = ZSql
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveFirst
                                        Do
                                            If .EOF = False Then
                                            
                                                ZZNroDespacho = IIf(IsNull(rstLaudo!NroDespacho), "", rstLaudo!NroDespacho)
                                                ZZProcedencia = IIf(IsNull(rstLaudo!Procedencia), "", rstLaudo!Procedencia)
                                                
                                                If Trim(ZZNroDespacho) <> "" Then
                                            
                                                    ZLugarDespa = ZLugarDespa + 1
                                                
                                                    ZZImpreDespa(ZLugarDespa, 1) = rstLaudo!Lote
                                                    ZZImpreDespa(ZLugarDespa, 2) = rstLaudo!Liberada
                                                    ZZImpreDespa(ZLugarDespa, 3) = rstLaudo!Saldo
                                                    ZZImpreDespa(ZLugarDespa, 4) = ZZNroDespacho
                                                    ZZImpreDespa(ZLugarDespa, 5) = ZZProcedencia
                                                    
                                                    ZZCantidad = ZZCantidad + rstLaudo!Liberada
                                                    ZZSaldo = ZZSaldo + rstLaudo!Saldo
                                                    
                                                End If
                                            
                                                .MoveNext
                                                    Else
                                                Exit Do
                                            End If
                                        Loop
                                    End With
                                    rstLaudo.Close
                                End If
                            
                                If ZZCantidad <> 0 Then
                                    ZZSaldo = ZZSaldo + Cantidad
                                    ZZConsumo = ZZCantidad - ZZSaldo
                                    For CicloCanti = 1 To ZLugarDespa
                                        If ZZConsumo > Val(ZZImpreDespa(CicloCanti, 2)) Then
                                            ZZConsumo = ZZConsumo - Val(ZZImpreDespa(CicloCanti, 2))
                                            ZZImpreDespa(CicloCanti, 2) = "0"
                                                Else
                                            ZZImpreDespa(CicloCanti, 2) = Str$(Val(ZZImpreDespa(CicloCanti, 2)) - ZZConsumo)
                                            Exit For
                                        End If
                                    Next CicloCanti
                                End If
                                
                                ZZTrabajo = Cantidad
                                
                                For CicloCanti = 1 To ZLugarDespa
                                
                                    If Val(ZZImpreDespa(CicloCanti, 2)) <> 0 Then
                                    
                                        If ZZTrabajo > ZZImpreDespa(CicloCanti, 2) Then
                                        
                                            ZZNroDespacho = Trim(UCase(ZZImpreDespa(ZLugarDespa, 4)))
                                            ZZProcedencia = Trim(UCase(ZZImpreDespa(ZLugarDespa, 5)))
                                            Entra = "S"
                                            For AltaLote = 1 To 100
                                                CA = ZZImpreDespaII(AltaLote, 1)
                                                If Trim(UCase(ZZNroDespacho)) = Trim(UCase(ZZImpreDespaII(AltaLote, 1))) Then
                                                    Entra = "N"
                                                    Exit For
                                                End If
                                            Next AltaLote
                                            If Entra = "S" Then
                                                ZZLugarDespaII = ZZLugarDespaII + 1
                                                ZZImpreDespaII(ZZLugarDespaII, 1) = ZZNroDespacho
                                                ZZImpreDespaII(ZZLugarDespaII, 2) = ZZProcedencia
                                            End If
                                                
                                            ZZTrabajo = ZZTrabajo - ZZImpreDespa(CicloCanti, 2)
                                            
                                                Else
                                            
                                            ZZNroDespacho = Trim(UCase(ZZImpreDespa(ZLugarDespa, 4)))
                                            ZZProcedencia = Trim(UCase(ZZImpreDespa(ZLugarDespa, 5)))
                                            Entra = "S"
                                            For AltaLote = 1 To 100
                                                CA = ZZImpreDespaII(AltaLote, 1)
                                                If Trim(UCase(ZZNroDespacho)) = Trim(UCase(ZZImpreDespaII(AltaLote, 1))) Then
                                                    Entra = "N"
                                                    Exit For
                                                End If
                                            Next AltaLote
                                            If Entra = "S" Then
                                                ZZLugarDespaII = ZZLugarDespaII + 1
                                                ZZImpreDespaII(ZZLugarDespaII, 1) = ZZNroDespacho
                                                ZZImpreDespaII(ZZLugarDespaII, 2) = ZZProcedencia
                                            End If
                                            
                                            Exit For
                                            
                                        End If
                                        
                                    End If
                                
                                Next CicloCanti
                                
                            End If
                        
                        Next CicloLote
                        
                    End If
                
                
                    Print #1, Tab(1); Alinea("#####.##", Str$(Cantidad));
                    Print #1, " Kg";
                    Print #1, Tab(15); Left$(Descri, 37);
                    parcial = Str$(Precio * Cantidad)
                    
                    Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem             Else
                    Rem     Precio = Str$(Val(Precio) * 1.21)
                    Rem     Parcial = Str$(Val(Parcial) * 1.21)
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem End If
                    
                    Print #1, Tab(56); " $ "; Alinea("####.##", Str$(Precio));
                    Print #1, Tab(68); Alinea("###,###.##", parcial)
                    
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a

        For aa = Impre To 19
                Print #1, ""
        Next aa
        
        
        For aa = 1 To ZZLugarDespaII
            If ZZImpreDespaII(aa, 1) <> "" Then
                Select Case aa
                    Case 1
                        ImpreDespachoI = "Despacho : " + ZZImpreDespaII(aa, 1) + "  " + ZZImpreDespaII(aa, 2)
                    Case 2, 3, 4
                        ImpreDespachoI = ImpreDespachoI + " / " + ZZImpreDespaII(aa, 1) + "  " + ZZImpreDespaII(aa, 2)
                    Case 5
                        ImpreDespachoII = "Despacho : " + ZZImpreDespaII(aa, 1) + "  " + ZZImpreDespaII(aa, 2)
                    Case 6, 7, 8
                        ImpreDespachoII = ImpreDespachoII + " / " + ZZImpreDespaII(aa, 1) + "  " + ZZImpreDespaII(aa, 2)
                    Case Else
                End Select
            End If
        Next aa
        
        

        Rem M# = Total# / 100
        Rem GoSub 4630
        

        Print #1, ImpreDespachoI
        Print #1, ImpreDespachoII
        Print #1, Tab(1); "EL IMPORTE DE ESTA FACTURA REPRESENTA U$S ";
        Print #1, Alinea("###,###.##", Str$(Impotot))
        Print #1, Tab(1); "CALCULADOS A UNA PARIDAD DE $ ";
        Print #1, Alinea("##.##", Str$(Paridad))
        Print #1, Tab(1); "Y DEBERA SER CANCELADO A SU VENCIMIENTO EN DOLARES"
        Print #1, Tab(1); "BILLETE  ESTADOUNIDENSES  O  EN  PESOS  AL  CAMBIO"
        Print #1, Tab(1); "OFICIAL  DEL DIA  DE ACREDITACION DE  LOS  VALORES"
        Print #1, Tab(1); "RECIBIDOS."
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        
        Print #1, Tab(65); " $ "; Alinea("###,###.##", Str$(XNeto))

        If Val(Dto.Caption) <> 0 Then
                Print #1, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Dto.Caption)
                        Else
                Print #1, ""
        End If

        If Val(Interes.Caption) <> 0 Then
                Print #1, Tab(56); "Interes";
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Interes.Caption)
                                                  Else
                Print #1, ""
        End If

        Print #1, Tab(3); M1;
        Print #1, Tab(65); " $ "; Alinea("###,###.##", Neto.Caption)
        Print #1, Tab(3); M2;
        If Val(Iva1.Caption) <> 0 Then
                Print #1, Tab(61); "21";
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Iva1.Caption)
                        Else
                Print #1, ""
        End If

        Select Case XX
                Case 1
                        Print #1, Tab(3); "ORIGINAL";
                Case 2
                        Print #1, Tab(3); "DUPLICADO";
                Case 3
                        Print #1, Tab(3); "TRIPLICADO";
                Case Else
        End Select

        If Val(ImpoIbCiudad.Caption) <> 0 Then
                Print #1, Tab(14); "P.Ciudad";
                Print #1, Tab(23); "U$S"; Alinea("##,###.##", ImpoIbCiudad.Caption);
        End If
        If Val(ImpoIbTucu.Caption) <> 0 Then
                Print #1, Tab(36); "P.Tucuman";
                Print #1, Tab(46); "U$S"; Alinea("##,###.##", ImpoIbTucu.Caption);
        End If
        If Val(ImpoIb.Caption) <> 0 Then
                Print #1, Tab(60); "I.B.";
                Print #1, Tab(65); " $ "; Alinea("##,###.##", ImpoIb.Caption)
                        Else
                If Val(Iva2.Caption) <> 0 Then
                    Print #1, Tab(60); "10.5";
                    Print #1, Tab(65); " $ "; Alinea("##,###.##", Iva2.Caption)
                        Else
                    Print #1, ""
                End If
        End If

        Print #1, Tab(65); " $ "; Alinea("###,###.##", Total.Caption); Chr$(12)

        Next XX%

        Close #1
        
End Sub

Sub Impresion_Remito()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(40) + "19U"
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(7); WRazon
        Print #1, Tab(7); Left$(WDireccion, 33)
        Print #1, Tab(7); Left$(WLocalidad, 33);
        Print #1, Tab(44); Pedido.Text;
        Print #1, Tab(57); Cliente.Text;
        Print #1, Tab(68); Orden.Text
        Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(7); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, Tab(30); WDirentrega;
        Print #1, ""
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                        
                        Print #1, Tab(14); Left$(Descri, 40);
                        Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg";
                        Print #1, Tab(71); "Netos"
                        Impre = Impre + 1
                End If
                
                If FF = 1 Then
                
                    ZLote1 = XLote(Suma, 1)
                    ZCantidad1 = XLote(Suma, 2)
                    ZLote2 = XLote(Suma, 3)
                    ZCantidad2 = XLote(Suma, 4)
                    ZLote3 = XLote(Suma, 5)
                    ZCantidad3 = XLote(Suma, 6)
                    ZLote4 = XLote(Suma, 7)
                    ZCantidad4 = XLote(Suma, 8)
                    ZLote5 = XLote(Suma, 9)
                    ZCantidad5 = XLote(Suma, 10)
                    ZLote6 = ""
                    ZCantidad6 = ""
                    ZLote7 = ""
                    ZCantidad7 = ""
                    ZLote8 = ""
                    ZCantidad8 = ""
                    ZLote9 = ""
                    ZCantidad9 = ""
                    ZLote10 = ""
                    ZCantidad10 = ""
                    ZLote11 = ""
                    ZCantidad11 = ""
                    ZLote12 = ""
                    ZCantidad12 = ""
                    
                    If Trim(Producto) <> "" Then
                    
                    If Left$(Producto, 2) = "DY" Then
                    
                        For ZZZCiclo = 1 To 12
                        
                            Select Case ZZZCiclo
                                Case 1
                                    ZZZLote = ZLote1
                                Case 2
                                    ZZZLote = ZLote2
                                Case 3
                                    ZZZLote = ZLote3
                                Case 4
                                    ZZZLote = ZLote4
                                Case 5
                                    ZZZLote = ZLote5
                                Case 6
                                    ZZZLote = ZLote6
                                Case 7
                                    ZZZLote = ZLote7
                                Case 8
                                    ZZZLote = ZLote8
                                Case 9
                                    ZZZLote = ZLote9
                                Case 10
                                    ZZZLote = ZLote10
                                Case 11
                                    ZZZLote = ZLote11
                                Case Else
                                    ZZZLote = ZLote12
                            End Select
                    
                            ZZZArti = Left$(Producto, 3) + Right$(Producto, 7)
                            XParam = "'" + ZZZLote + "','" _
                                         + ZZZArti + "'"
                            spLaudo = "ListaLaudoArticulo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                ZZZPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                                rstLaudo.Close
                                
                                If Trim(ZZZPartiOri) <> "" Then
                                
                                    Select Case ZZZCiclo
                                        Case 1
                                            ZLote1 = ZZZPartiOri
                                        Case 2
                                            ZLote2 = ZZZPartiOri
                                        Case 3
                                            ZLote3 = ZZZPartiOri
                                        Case 4
                                            ZLote4 = ZZZPartiOri
                                        Case 5
                                            ZLote5 = ZZZPartiOri
                                        Case 6
                                            ZLote6 = ZZZPartiOri
                                        Case 7
                                            ZLote7 = ZZZPartiOri
                                        Case 8
                                            ZLote8 = ZZZPartiOri
                                        Case 9
                                            ZLote9 = ZZZPartiOri
                                        Case 10
                                            ZLote10 = ZZZPartiOri
                                        Case 11
                                            ZLote11 = ZZZPartiOri
                                        Case Else
                                            ZLote12 = ZZZPartiOri
                                    End Select
                                    
                                End If
                            End If
                            
                        Next ZZZCiclo
                        
                            Else
                            
                        For ZZZCiclo = 1 To 12
                        
                            Select Case ZZZCiclo
                                Case 1
                                    ZZZLote = ZLote1
                                Case 2
                                    ZZZLote = ZLote2
                                Case 3
                                    ZZZLote = ZLote3
                                Case 4
                                    ZZZLote = ZLote4
                                Case 5
                                    ZZZLote = ZLote5
                                Case 6
                                    ZZZLote = ZLote6
                                Case 7
                                    ZZZLote = ZLote7
                                Case 8
                                    ZZZLote = ZLote8
                                Case 9
                                    ZZZLote = ZLote9
                                Case 10
                                    ZZZLote = ZLote10
                                Case 11
                                    ZZZLote = ZLote11
                                Case Else
                                    ZZZLote = ZLote12
                            End Select
                            
                            XParam = "'" + ZZZLote + "','" _
                                         + Producto + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                            
                                ZZZPartiOri = IIf(IsNull(rstHoja!LoteColorante), "", rstHoja!LoteColorante)
                                
                                rstHoja.Close
                                
                                If Trim(ZZZPartiOri) <> "" Then
                                
                                    Select Case ZZZCiclo
                                        Case 1
                                            ZLote1 = ZZZPartiOri
                                        Case 2
                                            ZLote2 = ZZZPartiOri
                                        Case 3
                                            ZLote3 = ZZZPartiOri
                                        Case 4
                                            ZLote4 = ZZZPartiOri
                                        Case 5
                                            ZLote5 = ZZZPartiOri
                                        Case 6
                                            ZLote6 = ZZZPartiOri
                                        Case 7
                                            ZLote7 = ZZZPartiOri
                                        Case 8
                                            ZLote8 = ZZZPartiOri
                                        Case 9
                                            ZLote9 = ZZZPartiOri
                                        Case 10
                                            ZLote10 = ZZZPartiOri
                                        Case 11
                                            ZLote11 = ZZZPartiOri
                                        Case Else
                                            ZLote12 = ZZZPartiOri
                                    End Select
                                    
                                End If
                            End If
                            
                        Next ZZZCiclo
                        
                    End If
                    
                    End If
                    
                    
                    
                    ZEnv1 = XLote(Suma, 11)
                    ZCantiEnv1 = XLote(Suma, 12)
                    ZEnv2 = XLote(Suma, 13)
                    ZCantiEnv2 = XLote(Suma, 14)
                    ZEnv3 = XLote(Suma, 15)
                    ZCantiEnv3 = XLote(Suma, 16)
                    ZEnv4 = XLote(Suma, 17)
                    ZCantiEnv4 = XLote(Suma, 18)
                    ZEnv5 = XLote(Suma, 19)
                    ZCantiEnv5 = XLote(Suma, 20)
                    
                    ZDescri1 = ""
                    ZDescri2 = ""
                    ZDescri3 = ""
                    ZDescri4 = ""
                    ZDescri5 = ""
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv1 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri1 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv2 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri2 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv3 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri3 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv4 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri4 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    spEnvases = "ConsultaEnvases " + "'" + ZEnv5 + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        ZDescri5 = Left$(rstEnvases!Abreviatura, 8)
                        rstEnvases.Close
                    End If
                    
                    If Val(ZCantidad1) <> 0 Then
                    
                        Print #1, Tab(1); Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72);
                        
                        Print #1, Tab(10); Alinea("#####", ZCantidad1); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote1, 8);
                        If Val(ZCantiEnv1) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv1);
                            Print #1, " x "; ZDescri1;
                        End If
                    
                        If Val(ZCantidad2) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad2); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote2, 8);
                            If Val(ZCantiEnv2) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv2);
                                Print #1, " x "; ZDescri2;
                            End If
                        End If
                    
                        If Val(ZCantidad3) <> 0 Then
                            Print #1, Tab(86); "|";
                            Print #1, Tab(88); Alinea("#####", ZCantidad3); " Kg";
                            Print #1, Tab(97); "Lote:"; Left$(ZLote3, 8);
                            If Val(ZCantiEnv3) <> 0 Then
                                Print #1, Tab(109); Alinea("##", ZCantiEnv3);
                                Print #1, " x "; ZDescri3;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                    If Val(ZCantidad4) <> 0 Then
                    
                        Print #1, Tab(10); Alinea("#####", ZCantidad4); " Kg";
                        Print #1, Tab(19); "Lote:"; Left$(ZLote4, 8);
                        If Val(ZCantiEnv4) <> 0 Then
                            Print #1, Tab(33); Alinea("##", ZCantiEnv4);
                            Print #1, " x "; ZDescri4;
                        End If
                    
                        If Val(ZCantidad5) <> 0 Then
                            Print #1, Tab(48); "|";
                            Print #1, Tab(50); Alinea("#####", ZCantidad5); " Kg";
                            Print #1, Tab(59); "Lote:"; Left$(ZLote5, 8);
                            If Val(ZCantiEnv5) <> 0 Then
                                Print #1, Tab(71); Alinea("##", ZCantiEnv5);
                                Print #1, " x "; ZDescri5;
                            End If
                        End If
                            
                        Print #1, ""
                        Impre = Impre + 1
                        
                    End If
                    
                End If
                
                    
            Next iRow
            
        Next a
        
        If FF = 2 Then
        
            If Val(WEmpresa) = 4 Or Val(WEmpresa) = 8 Then
                For aa = Impre To 10
                    Impre = Impre + 1
                    Print #1, ""
                Next aa
                    Else
                For aa = Impre To 12
                    Impre = Impre + 1
                    Print #1, ""
                Next aa
            End If
            
            If Val(WEmpresa) = 1 Then
            
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "15" + Chr$(72);
                Print #1, "  -----------------------------------------------------------------------------------------------------------------"
                Print #1, "  |   ETIQUETADO          | Si/No | TRANSPORTE                    | Si/No | SI HAY SUSTANCIAS PELIGROSAS  | Si/No |"
                Print #1, "  -----------------------------------------------------------------------------------------------------------------"
                Print #1, "  | Cliente               |       | Conductor                     |       | Ficha de Intervencion         |       |"
                Print #1, "  | Nombre                |       | H.de Ruta/Guia Traslasdo      |       | Rotulos Externos              |       |"
                Print #1, "  | Codigo                |       | Remitos                       |       |----------------------------------------"
                Print #1, "  | Partida               |       | Facturas                      |       | OBSERVACIONES :                       |"
                Print #1, "  | Neto                  |       | Certificado de Analisis       |       |                                       |"
                Print #1, "  | Vencimiento           |       | Hoja de Seguridad             |       |                                       |"
                Print #1, "  | Etiq./Irradiacion     |       | Certificado de Irradicacion   |       |---------------------------------------|"
                Print #1, "  |                       |       | Van Muestras                  |       | VERIFICO :                            |"
                Print #1, "  -----------------------------------------------------------------------------------------------------------------"
                Impre = Impre + 13
                
                    Else
                    
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "15" + Chr$(72);
                Print #1, "  -----------------------------------------------------------------------------------------------------------------"
                Print #1, "  |   ETIQUETADO          | Si/No | TRANSPORTE                    | Si/No | SI HAY SUSTANCIAS PELIGROSAS  | Si/No |"
                Print #1, "  -----------------------------------------------------------------------------------------------------------------"
                Print #1, "  | Cliente               |       | Conductor                     |       | Ficha de Intervencion         |       |"
                Print #1, "  | Nombre                |       | H.de Ruta/Guia Traslasdo      |       | Rotulos Externos              |       |"
                Print #1, "  | Codigo                |       | Remitos                       |       |----------------------------------------"
                Print #1, "  | Partida               |       | Facturas                      |       | VERIFICO :                            |"
                Print #1, "  | Neto                  |       | Certificado de Analisis       |       |---------------------------------------|"
                Print #1, "  |                       |       | Hoja de Seguridad             |       | ENTREGA ENVASES               | Si/No |"
                Print #1, "  |                       |       | Van Muestras                  |       | MOTIVO :                              |"
                Print #1, "  |                       |       |                               |       |                                       |"
                Print #1, "  ------------------------------------------------------------------------|                                       |"
                Print #1, "  |                       |       |                               |       | Firma                                 |"
                Impre = Impre + 13
                    
            End If
        
        End If
        
        
        Select Case Val(WEmpresa)
            Case 4, 8
                If FF = 1 Then
                    For aa = Impre To 17
                        Print #1, ""
                    Next aa

                    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                    Print #1, Tab(4); "Pellital S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                    Print #1, Tab(4); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                    Print #1, Tab(4); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                    Print #1, Tab(4); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                    Print #1, Tab(4); "vigentes del ámbito municipal, provincial y nacional"
                    Print #1, ""
                End If

                For XDa = 1 To 1
                        For DA = 1 To 9
                                If Val(Stk(DA, 4)) <> 0 Then
                                        
                                        Select Case DA
                                                Case 1
                                                        Lugar = 25
                                                Case 2
                                                        Lugar = 36
                                                Case 3
                                                        Lugar = 47
                                                Case 4
                                                        Lugar = 58
                                                Case 5
                                                        Lugar = 69
                                                Case 6
                                                        Lugar = 80
                                                Case 7
                                                        Lugar = 92
                                                Case 8
                                                        Lugar = 104
                                                Case 9
                                                        Lugar = 116
                                                Case Else
                                        End Select
                                                            
                                        If DA = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(DA, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(DA, XDa);
                                        End If
                                    End If
        
                        Next DA
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For DA = 1 To 9
            
                                If Val(Stk(DA, 4)) <> 0 Then
        
                                        Select Case DA
                                            Case 1
                                                Lugar = 16
                                            Case 2
                                                Lugar = 23
                                            Case 3
                                                Lugar = 31
                                            Case 4
                                                Lugar = 38
                                            Case 5
                                                Lugar = 45
                                            Case 6
                                                Lugar = 52
                                            Case 7
                                                Lugar = 59
                                            Case 8
                                                Lugar = 66
                                            Case 9
                                                Lugar = 73
                                            Case Else
                                    End Select
        
                                    If Val(Stk(DA, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(DA, XDa))));
                                    End If
        
                            End If
                    Next DA
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, ""
                Print #1, ""
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
            
            Case Else
                If FF = 1 Then
                    For aa = Impre To 19
                        Print #1, ""
                    Next aa

                    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                    Print #1, Tab(3); "Surfactan S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                    Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                    Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                    Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                    Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional"
                    Print #1, ""
                End If

                For XDa = 1 To 1
                        For DA = 1 To 9
                                If Val(Stk(DA, 4)) <> 0 Then
                                        
                                        Select Case DA
                                                Case 1
                                                        Lugar = 22
                                                Case 2
                                                        Lugar = 33
                                                Case 3
                                                        Lugar = 44
                                                Case 4
                                                        Lugar = 55
                                                Case 5
                                                        Lugar = 66
                                                Case 6
                                                        Lugar = 77
                                                Case 7
                                                        Lugar = 89
                                                Case 8
                                                        Lugar = 101
                                                Case 9
                                                        Lugar = 113
                                                Case Else
                                        End Select
                                                            
                                        If DA = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(DA, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(DA, XDa);
                                        End If
                                    End If
        
                        Next DA
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For DA = 1 To 9
            
                                If Val(Stk(DA, 4)) <> 0 Then
        
                                        Select Case DA
                                            Case 1
                                                Lugar = 14
                                            Case 2
                                                Lugar = 21
                                            Case 3
                                                Lugar = 29
                                            Case 4
                                                Lugar = 36
                                            Case 5
                                                Lugar = 43
                                            Case 6
                                                Lugar = 50
                                            Case 7
                                                Lugar = 57
                                            Case 8
                                                Lugar = 64
                                            Case 9
                                                Lugar = 71
                                            Case Else
                                    End Select
        
                                    If Val(Stk(DA, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(DA, XDa))));
                                    End If
        
                            End If
                    Next DA
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
                
        End Select

        Next FF

        Close #1

End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo

    Erase Stk

    If Val(WEmpresa) = 8 Then
        Stk(1, 1) = "005"
        Stk(2, 1) = "011"
        Stk(3, 1) = "021"
        Stk(4, 1) = "027"
        Stk(5, 1) = "004"
        Stk(6, 1) = "012"
        Stk(7, 1) = "000"
        Stk(8, 1) = "000"
        Stk(9, 1) = "000"
            Else
        Stk(1, 1) = "020"
        Stk(2, 1) = "021"
        Stk(3, 1) = "022"
        Stk(4, 1) = "023"
        Stk(5, 1) = "024"
        Stk(6, 1) = "025"
        Stk(7, 1) = "026"
        Stk(8, 1) = "030"
        Stk(9, 1) = "028"
    End If

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For DA = 1 To 9
                        If Val(Stk(DA, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) + !Cantidad)
                                    Else
                                Stk(DA, 2) = Str$(Val(Stk(DA, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next DA
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If

End Sub

Private Sub Verifica_Lote()

    Renglon = 0
    Renglon1 = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    ZVeriSedronar = "N"
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
            
            Suma = Suma + 1
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            WTipoProDy = Left$(Articulo, 2)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
                    
            If Cantidad <> 0 Then
            
                ZSedronar = 0
                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZSedronar = IIf(IsNull(rstTerminado!Sedronar), "0", rstTerminado!Sedronar)
                    rstTerminado.Close
                End If
                
                If ZSedronar = 1 Then
                    ZVeriSedronar = "S"
                End If
            
                WEstado = "N"
                SumaCant = 0
    
                WLote1 = XLote(Suma, 1)
                WLote2 = XLote(Suma, 3)
                Wlote3 = XLote(Suma, 5)
                WLote4 = XLote(Suma, 7)
                WLote5 = XLote(Suma, 9)
                WImpo = XLote(Suma, 2)
                WCanti1 = Str$(WImpo)
                WImpo = XLote(Suma, 4)
                WCanti2 = Str$(WImpo)
                WImpo = XLote(Suma, 6)
                WCanti3 = Str$(WImpo)
                WImpo = XLote(Suma, 8)
                WCanti4 = Str$(WImpo)
                WImpo = XLote(Suma, 10)
                WCanti5 = Str$(WImpo)
    
                If Val(WLote1) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti1)
                End If
                If Val(WLote2) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti2)
                End If
                If Val(Wlote3) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti3)
                End If
                If Val(WLote4) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti4)
                End If
                If Val(WLote5) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti5)
                End If
    
                If SumaCant = Cantidad Then
                    WEstado = "S"
                        Else
                    WEstado = "N"
                    m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                    Exit Sub
                End If
    
                If WEstado = "S" Then
    
                    Erase ControlLote
                    ControlLote(1, 1) = WLote1
                    ControlLote(1, 2) = WCanti1
                    ControlLote(2, 1) = WLote2
                    ControlLote(2, 2) = WCanti2
                    ControlLote(3, 1) = Wlote3
                    ControlLote(3, 2) = WCanti3
                    ControlLote(4, 1) = WLote4
                    ControlLote(4, 2) = WCanti4
                    ControlLote(5, 1) = WLote5
                    ControlLote(5, 2) = WCanti5
    
                    For Ciclo1 = 1 To 5
                        If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                            For Ciclo2 = 1 To 5
                                If Ciclo1 <> Ciclo2 Then
                                    If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                                        m$ = "A asignado una misma partida 2 veces"
                                        a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        WEstado = "N"
                                        Exit Sub
                                    End If
                                End If
                            Next Ciclo2
                        End If
                        If WEstado = "N" Then
                            Exit For
                        End If
                    Next Ciclo1
        
                End If

                If WEstado = "S" Then
    
                    Erase ControlLote
                    ControlLote(1, 1) = WLote1
                    ControlLote(1, 2) = WCanti1
                    ControlLote(2, 1) = WLote2
                    ControlLote(2, 2) = WCanti2
                    ControlLote(3, 1) = Wlote3
                    ControlLote(3, 2) = WCanti3
                    ControlLote(4, 1) = WLote4
                    ControlLote(4, 2) = WCanti4
                    ControlLote(5, 1) = WLote5
                    ControlLote(5, 2) = WCanti5
    
                    For Ciclo1 = 1 To 5
    
                        WLote = ControlLote(Ciclo1, 1)
                        WCanti = Val(ControlLote(Ciclo1, 2))
            
                        If Val(WLote) <> 0 Or Val(WCanti) <> 0 Then
            
                        If Left$(Articulo, 2) <> "PT" Then
                            WTipopro = "M"
                                Else
                            WTipopro = "T"
                        End If
            
                        Select Case WTipopro
                            Case "M"
                            
                                XEmpresa = WEmpresa
                                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                    Select Case WTipoPedido
                                        Case "PG", "CO"
                                            WEmpresa = "0001"
                                            txtOdbc = "Empresa01"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "FA"
                                            WEmpresa = "0005"
                                            txtOdbc = "Empresa05"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "TA"
                                            WEmpresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            WEmpresa = "0007"
                                            txtOdbc = "Empresa07"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                End If
                            
                                WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                                WEntra = "N"
                                XParam = "'" + WLote + "','" _
                                             + WArti + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WSal = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Call Redondeo(WSal)
                                    WEntra = "S"
                                    rstLaudo.Close
                                    If WSal < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        WEstado = "N"
                                        Call Conecta_Empresa
                                        Exit Sub
                                    End If
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArti + "','" _
                                                 + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        rstMovguia.Close
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Call Conecta_Empresa
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                
                                Call Conecta_Empresa
                                
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                            Case Else
                                WEntra = "N"
                                WControla = 0
                                
                                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                    rstTerminado.Close
                                End If
                
                                If WControla = 0 Then
                                
                                    XEmpresa = WEmpresa
                                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                        Select Case WTipoPedido
                                            Case "PG", "CO"
                                                WEmpresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case "FA"
                                                WEmpresa = "0005"
                                                txtOdbc = "Empresa05"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case "TA"
                                                WEmpresa = "0003"
                                                txtOdbc = "Empresa03"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                                WEmpresa = "0007"
                                                txtOdbc = "Empresa07"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        End Select
                                    End If
                                
                                    XParam = "'" + WLote + "','" _
                                            + Articulo + "'"
                                    spHoja = "ListaHojaProducto " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        rstHoja.Close
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Call Conecta_Empresa
                                            Exit Sub
                                        End If
                                    End If
                
                                    If WEntra = "N" Then
                                        XParam = "'" + Articulo + "','" _
                                                    + WLote + "'"
                                        spMovguia = "ListaMovguiaLote1 " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                            WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                            Call Redondeo(WSal)
                                            WEntra = "S"
                                            rstMovguia.Close
                                            If WSal < WCanti Then
                                                m$ = "La cantidad informada supera al saldo disponible"
                                                a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                                WEstado = "N"
                                                Call Conecta_Empresa
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                    
                                    Call Conecta_Empresa
                
                                        Else
                                        
                                    WEntra = "S"
                                    
                                End If
                                
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    a = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                        End Select
            
                        End If
            
                    Next Ciclo1

                End If
                
            End If
                                        
        Next iRow
            
    Next a
    
    If ZVeriSedronar = "S" Then
        ZNroSedronar = ""
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZNroSedronar = Trim(IIf(IsNull(rstCliente!NroSedronar), "", rstCliente!NroSedronar))
            rstCliente.Close
        End If
        If Trim(ZNroSedronar) = "" Then
            m$ = "Atencion: El cliente debe estar inscripto en el Sedronar para adquirir estos productos"
            Aaa% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
            WEstado = "N"
            Exit Sub
        End If
    End If
    
    
End Sub



Private Sub DBGrid1_DblClick()
    
    WLugar = DBGrid1.Row + 1
    
    CargaLote2.Visible = True
                        
    ZZPartida1.Text = ""
    ZZCanti1.Text = ""
    ZZPartida2.Text = ""
    ZZCanti2.Text = ""
    ZZPartida3.Text = ""
    ZZCanti3.Text = ""
    ZZPartida4.Text = ""
    ZZCanti4.Text = ""
    ZZPartida5.Text = ""
    ZZCanti5.Text = ""
                        
    ZZEnvase1.Text = ""
    ZZCantiEnv1.Text = ""
    ZZDescri1.Caption = ""
    ZZEnvase2.Text = ""
    ZZCantiEnv2.Text = ""
    ZZDescri2.Caption = ""
    ZZEnvase3.Text = ""
    ZZCantiEnv3.Text = ""
    ZZDescri3.Caption = ""
    ZZEnvase4.Text = ""
    ZZCantiEnv4.Text = ""
    ZZDescri4.Caption = ""
    ZZEnvase5.Text = ""
    ZZCantiEnv5.Text = ""
    ZZDescri5.Caption = ""
                       
    If XLote(WLugar, 1) <> "" Then
        ZZPartida1.Text = XLote(WLugar, 1)
        ZZCanti1.Text = XLote(WLugar, 2)
    End If
    If XLote(WLugar, 3) <> "" Then
        ZZPartida2.Text = XLote(WLugar, 3)
        ZZCanti2.Text = XLote(WLugar, 4)
    End If
    If XLote(WLugar, 5) <> "" Then
        ZZPartida3.Text = XLote(WLugar, 5)
        ZZCanti3.Text = XLote(WLugar, 6)
    End If
    If XLote(WLugar, 7) <> "" Then
        ZZPartida4.Text = XLote(WLugar, 7)
        ZZCanti4.Text = XLote(WLugar, 8)
    End If
    If XLote(WLugar, 9) <> "" Then
        ZZPartida5.Text = XLote(WLugar, 9)
        ZZCanti5.Text = XLote(WLugar, 10)
    End If
                        
    If Val(XLote(WLugar, 11)) <> 0 Then
        ZZEnvase1.Text = XLote(WLugar, 11)
        ZZCantiEnv1.Text = XLote(WLugar, 12)
        spEnvases = "ConsultaEnvases " + "'" + ZZEnvase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            ZZDescri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
        End If
    End If
    
    If Val(XLote(WLugar, 13)) <> 0 Then
        ZZEnvase2.Text = XLote(WLugar, 13)
        ZZCantiEnv2.Text = XLote(WLugar, 14)
        spEnvases = "ConsultaEnvases " + "'" + ZZEnvase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            ZZDescri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
        End If
    End If
    
    If Val(XLote(WLugar, 15)) <> 0 Then
        ZZEnvase3.Text = XLote(WLugar, 15)
        ZZCantiEnv3.Text = XLote(WLugar, 16)
        spEnvases = "ConsultaEnvases " + "'" + ZZEnvase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            ZZDescri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
        End If
    End If
    
    If Val(XLote(WLugar, 17)) <> 0 Then
        ZZEnvase4.Text = XLote(WLugar, 17)
        ZZCantiEnv4.Text = XLote(WLugar, 18)
        spEnvases = "ConsultaEnvases " + "'" + ZZEnvase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            ZZDescri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
        End If
    End If
    
    If Val(XLote(WLugar, 19)) <> 0 Then
        ZZEnvase5.Text = XLote(WLugar, 19)
        ZZCantiEnv5.Text = XLote(WLugar, 20)
        spEnvases = "ConsultaEnvases " + "'" + ZZEnvase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            ZZDescri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
        End If
    End If
    
End Sub

Private Sub FinCarga_Click()
    CargaLote2.Visible = False
End Sub






