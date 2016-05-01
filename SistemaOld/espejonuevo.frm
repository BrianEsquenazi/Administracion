VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form EspejoNuevo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos en U$S"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
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
      TabIndex        =   110
      Top             =   1200
      Width           =   2055
   End
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
      Left            =   840
      TabIndex        =   76
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
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
         TabIndex        =   107
         Top             =   2520
         Width           =   975
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
         TabIndex        =   96
         Top             =   600
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
         TabIndex        =   95
         Top             =   960
         Width           =   1335
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
         TabIndex        =   94
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   93
         Top             =   600
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
         TabIndex        =   92
         Top             =   960
         Width           =   975
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
         TabIndex        =   91
         Top             =   1320
         Width           =   975
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
         TabIndex        =   90
         Top             =   1680
         Width           =   1335
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
         TabIndex        =   89
         Top             =   2040
         Width           =   1335
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
         TabIndex        =   88
         Top             =   1680
         Width           =   975
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
         TabIndex        =   87
         Top             =   2040
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
         TabIndex        =   86
         Text            =   " "
         Top             =   600
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
         TabIndex        =   85
         Text            =   " "
         Top             =   600
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
         TabIndex        =   84
         Text            =   " "
         Top             =   960
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
         TabIndex        =   83
         Text            =   " "
         Top             =   960
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
         TabIndex        =   82
         Text            =   " "
         Top             =   1320
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
         TabIndex        =   81
         Text            =   " "
         Top             =   1320
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
         TabIndex        =   80
         Text            =   " "
         Top             =   1680
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
         TabIndex        =   79
         Text            =   " "
         Top             =   1680
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
         TabIndex        =   78
         Text            =   " "
         Top             =   2040
         Width           =   855
      End
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
         TabIndex        =   77
         Text            =   " "
         Top             =   2040
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
         TabIndex        =   106
         Top             =   240
         Width           =   1335
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
         TabIndex        =   105
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
         TabIndex        =   104
         Top             =   600
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
         TabIndex        =   103
         Top             =   240
         Width           =   975
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
         TabIndex        =   102
         Top             =   240
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
         TabIndex        =   101
         Top             =   240
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
         TabIndex        =   100
         Top             =   960
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
         TabIndex        =   99
         Top             =   1320
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
         TabIndex        =   98
         Top             =   1680
         Width           =   855
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
         TabIndex        =   97
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.Frame PantaMotivo 
      Height          =   1095
      Left            =   480
      TabIndex        =   71
      Top             =   2400
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox DescriMotivo 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   72
         Top             =   720
         Width           =   9855
      End
      Begin VB.Label Label17 
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
         TabIndex        =   73
         Top             =   360
         Width           =   9735
      End
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
      Top             =   600
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
      Top             =   600
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   75
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         TabIndex        =   74
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "IB Bs Aa"
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
      Left            =   4080
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
      Top             =   600
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
      Left            =   4200
      TabIndex        =   19
      Top             =   5760
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
      Left            =   6360
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
      Top             =   0
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
      Top             =   0
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
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
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
      ItemData        =   "espejonuevo.frx":0000
      Left            =   5280
      List            =   "espejonuevo.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "espejonuevo.frx":0015
      TabIndex        =   2
      Top             =   1680
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
      Left            =   3240
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
      Left            =   5640
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
Attribute VB_Name = "EspejoNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Impresion_Varios()

    ZZVersion = 0
  Rem BY NAN
  Rem  ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
  Rem  ZZEstado = Dir(ZZRuta)
 Rem   ZZEstado = Trim(ZZEstado)
    If ZZEstado <> "" Then
        ZZVersion = 1
            Else
        ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe"
        ZZEstado = Dir(ZZRuta)
        ZZEstado = Trim(ZZEstado)
        If ZZEstado <> "" Then
            ZZVersion = 2
                Else
            ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe"
            ZZEstado = Dir(ZZRuta)
            ZZEstado = Trim(ZZEstado)
            If ZZEstado <> "" Then
                ZZVersion = 3
                    Else
                ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe"
                ZZEstado = Dir(ZZRuta)
                ZZEstado = Trim(ZZEstado)
                If ZZEstado <> "" Then
                    ZZVersion = 4
                        Else
                    ZZRuta = "C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe"
                    ZZEstado = Dir(ZZRuta)
                    ZZEstado = Trim(ZZEstado)
                    If ZZEstado <> "" Then
                        ZZVersion = 5
                          Rem by nan
                          Else
                     ZZRuta = "C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe"
                     ZZEstado = Dir(ZZRuta)
                     ZZEstado = Trim(ZZEstado)
                         If ZZEstado <> "" Then
                           ZZVersion = 6
                        End If
                        Rem fin by nan
                    
                    End If
                End If
            End If
        End If
    End If
    
    Erase ZImpreFicha
    ZLugarFicha = 0
    
    For DA = 1 To 99
    
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
        WLote6 = Auxiliar(DA, 15)
        WCanti6 = Auxiliar(DA, 16)
        WLote7 = Auxiliar(DA, 17)
        WCanti7 = Auxiliar(DA, 18)
        WLote8 = Auxiliar(DA, 19)
        WCanti8 = Auxiliar(DA, 20)
        WLote9 = Auxiliar(DA, 21)
        WCanti9 = Auxiliar(DA, 22)
        WLote10 = Auxiliar(DA, 23)
        WCanti10 = Auxiliar(DA, 24)
        WLote11 = Auxiliar(DA, 25)
        WCanti11 = Auxiliar(DA, 26)
        WLote12 = Auxiliar(DA, 27)
        WCanti12 = Auxiliar(DA, 28)
        ZZDescriArticulo = Trim(Auxiliar(DA, 30))
        
        If Articulo = "" Then Exit For
    
        Rem
        Rem Hoja de Seguridad
        Rem
        
        ZZRequiereCertificado = 0
        ZZRequiereMsds = 0
        ZZRequiereMsdsCada = 0
        ZZRequiereHoja = 0
        ZZBusqueda = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteEspecif"
        ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + Cliente.Text + "'"
        spClienteEspecif = ZSql
        Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteEspecif.RecordCount > 0 Then
            ZZRequiereCertificado = IIf(IsNull(rstClienteEspecif!RequiereCertificado), "0", rstClienteEspecif!RequiereCertificado)
            ZZRequiereMsds = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
            ZZRequiereMsdsCada = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
            ZZRequiereHoja = IIf(IsNull(rstClienteEspecif!RequiereHoja), "0", rstClienteEspecif!RequiereHoja)
            rstClienteEspecif.Close
        End If
        
        If ZZRequiereMsdsCada = 1 Then
            ZZBusqueda = "S"
                Else
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Cliente = " + "'" + Cliente.Text + "'"
            ZSql = ZSql + " and Estadistica.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " and Estadistica.Numero <> " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                rstEstadistica.Close
                    Else
                ZZBusqueda = "S"
            End If
        End If
        
        If ZZBusqueda = "S" Then
            
            If Left$(UCase(Articulo), 2) = "PT" Then

                Es = ZZDescriArticulo
                x = ""
                For XX = 1 To Len(Es)
                    Y = Mid$(Es, XX, 1)
                    If Y <> " " And Y <> "/" Then
                        x = x + Y
                    End If
                Next
                ZZCodArt = x + Mid$(Articulo, 4, 5) + Right$(Articulo, 3)
                
                    Else
                    
                ZZCodArt = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                
            End If
            
            ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArt + ".PDF"
            ZZEstado = Dir(ZZRuta)
            ZZEstado = Trim(ZZEstado)
            If ZZEstado <> "" Then
                Select Case ZZVersion
                    Case 1
                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Case 2
                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Case 3
                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Case 4
                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Case 5
                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                    Case Else
                        RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                
                End Select
                    Else
                m$ = "El MSDS  (" + ZZCodArt + ")  no se ha encontrado"
                a% = MsgBox(m$, 0, "Impresion de comprobantes varios")
            End If
            
        End If
    
                    
        Rem
        Rem ficha de intevencion
        Rem
        If Left$(UCase(Articulo), 2) = "PT" Then
                
            ZZIntervencion = ""
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                rstTerminado.Close
            End If
            
            If Val(ZZIntervencion) <> 0 Then
            
                ZEntraFicha = "S"
                For ZZCicloFicha = 1 To ZLugarFicha
                    If ZImpreFicha(ZZCicloFicha) = Val(ZZIntervencion) Then
                        ZEntraFicha = "N"
                        Exit For
                    End If
                Next ZZCicloFicha
                If ZEntraFicha = "S" Then
                    ZLugarFicha = ZLugarFicha + 1
                    ZImpreFicha(ZLugarFicha) = ZZIntervencion
                End If
                
            End If
            
                Else
                
            Rem toto
                
            ZZCodArt = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
            ZZIntervencion = ""
            spArticulo = "ConsultaArticulo " + "'" + ZZCodArt + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                rstArticulo.Close
            End If
            
            If Val(ZZIntervencion) <> 0 Then
            
                ZEntraFicha = "S"
                For ZZCicloFicha = 1 To ZLugarFicha
                    If ZImpreFicha(ZZCicloFicha) = Val(ZZIntervencion) Then
                        ZEntraFicha = "N"
                        Exit For
                    End If
                Next ZZCicloFicha
                If ZEntraFicha = "S" Then
                    ZLugarFicha = ZLugarFicha + 1
                    ZImpreFicha(ZLugarFicha) = ZZIntervencion
                End If
                
            End If
                
        End If
                    
            
    
        Rem
        Rem certificado de analisis
        Rem
    
        For ZZCiclo = 1 To 12
            
            Select Case ZZCiclo
                Case 1
                    ZZLugar = 5
                Case 2
                    ZZLugar = 7
                Case 3
                    ZZLugar = 9
                Case 4
                    ZZLugar = 11
                Case 5
                    ZZLugar = 13
                Case 6
                    ZZLugar = 15
                Case 7
                    ZZLugar = 17
                Case 8
                    ZZLugar = 19
                Case 9
                    ZZLugar = 21
                Case 10
                    ZZLugar = 23
                Case 11
                    ZZLugar = 25
                Case Else
                    ZZLugar = 27
            End Select
            
            If Val(Auxiliar(DA, ZZLugar)) <> 0 Then
        
                ZZEntra = "N"
        
                If Left$(UCase(Articulo), 2) = "PT" Then
                
                    XCodigo = Val(Mid$(Articulo, 4, 5))
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 11999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    If XCodigo >= 40000 And XCodigo <= 41000 Then
                                        XTipoPro = "TA"
                                            Else
                                        XTipoPro = "PT"
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    If Left$(Articulo, 2) = "YQ" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YH" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YP" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YF" Then
                        XTipoPro = "FA"
                    End If
            
                    ZLinea = 0
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
            
                    Select Case ZLinea
                        Case 8
                            XTipoPro = "PG"
                        Case 10, 20
                            XTipoPro = "FA"
                        Case Else
                    End Select
            
                    If XTipoPro <> "FA" And XTipoPro <> "CO" And XTipoPro <> "TA" Then
                    
                        XEmpresa = WEmpresa
                        
                        Select Case Val(WEmpresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                            
                        Erase ZOpcion
                        Erase ZValor
                        Erase ZEnsayo
                        Erase ZStd
                        Erase ZDescri
                        Erase ZDescriII
                            
                        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        ZVersion = 0
                        
                        ZZEntra = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM AltaCertificado"
                        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                        ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                        spAltaCertificado = ZSql
                        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstAltaCertificado.RecordCount > 0 Then
                            ZOpcion(1) = rstAltaCertificado!Opcion1
                            ZOpcion(2) = rstAltaCertificado!Opcion2
                            ZOpcion(3) = rstAltaCertificado!Opcion3
                            ZOpcion(4) = rstAltaCertificado!Opcion4
                            ZOpcion(5) = rstAltaCertificado!Opcion5
                            ZOpcion(6) = rstAltaCertificado!Opcion6
                            ZOpcion(7) = rstAltaCertificado!Opcion7
                            ZOpcion(8) = rstAltaCertificado!Opcion8
                            ZOpcion(9) = rstAltaCertificado!Opcion9
                            ZOpcion(10) = rstAltaCertificado!Opcion10
                            rstAltaCertificado.Close
                            ZZEntra = "S"
                        End If
                        
                        If ZZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                        End If
                        
                        If ZZEntra = "S" Then
                            If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                ZZEntra = "N"
                            End If
                        End If
                                
                        If ZZEntra = "N" Then
                            m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                        End If
                                        
                        If ZZEntra = "S" Then
                            Select Case Val(XEmpresa)
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
                                    ZHasta1 = 7
                                Case Else
                                    CargaEmpresa(1, 1) = "0002"
                                    CargaEmpresa(1, 2) = "Empresa02"
                                    CargaEmpresa(2, 1) = "0004"
                                    CargaEmpresa(2, 2) = "Empresa04"
                                    CargaEmpresa(3, 1) = "0008"
                                    CargaEmpresa(3, 2) = "Empresa08"
                                    CargaEmpresa(4, 1) = "0009"
                                    CargaEmpresa(4, 2) = "Empresa09"
                                    ZHasta1 = 4
                            End Select
                                
                            For ZCiclo = 1 To ZHasta1
                                
                                WEmpresa = CargaEmpresa(ZCiclo, 1)
                                txtOdbc = CargaEmpresa(ZCiclo, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        
                                If Val(ZLote) > 99999 Then
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 6)
                                        Else
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 5)
                                End If
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Prueter"
                                ZSql = ZSql + " Where Prueter.Lote = " + "'" + ZLote + "'"
                                spPrueter = ZSql
                                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPrueter.RecordCount > 0 Then
                                
                                    If Left$(rstPrueter!prueba, 1) = "2" Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                    If rstPrueter!Producto <> ZProducto Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                        
                                    WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
                                            
                                    ZValor(1) = rstPrueter!Valor1
                                    ZValor(2) = rstPrueter!valor2
                                    ZValor(3) = rstPrueter!Valor3
                                    ZValor(4) = rstPrueter!valor4
                                    ZValor(5) = rstPrueter!valor5
                                    ZValor(6) = rstPrueter!valor6
                                    ZValor(7) = rstPrueter!valor7
                                    ZValor(8) = rstPrueter!valor8
                                    ZValor(9) = rstPrueter!valor9
                                    ZValor(10) = rstPrueter!valor10
                                        
                                    rstPrueter.Close
                                    
                                    WFechaElaboracion = ""
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Hoja"
                                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZLote + "'"
                                    spHoja = ZSql
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                                        ZZHoja = rstHoja!Hoja
                                        ZZProducto = rstHoja!Producto
                                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                        ZZFecha = rstHoja!Fecha
                                        ZZMeses = ""
                                        rstHoja.Close
                                        
                                        spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                                            rstTerminado.Close
                                        End If
                                        
                                        If Val(ZZMeses) <> 0 Then
                                        
                                            If Val(ZZRevalida) <> 0 Then
                                                WVida = Val(ZZMesesRevalida)
                                                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                                WAno = Val(Right$(ZZFechaRevalida, 4))
                                                    Else
                                                WVida = Val(ZZMeses)
                                                WMes = Val(Mid$(ZZFecha, 4, 2))
                                                WAno = Val(Right$(ZZFecha, 4))
                                            End If
                                            
                                            For Ciclo = 1 To WVida
                                                WMes = WMes + 1
                                                If WMes > 12 Then
                                                    WAno = WAno + 1
                                                    WMes = 1
                                                End If
                                            Next Ciclo
                                            ZMes = Str$(WMes)
                                            ZAno = Str$(WAno)
                                            Call Ceros(ZMes, 2)
                                            Call Ceros(ZAno, 4)
                                            WFechaElaboracion = ZMes + "/" + ZAno
                                            
                                        End If
                                        
                                    End If
                                        
                                    If Left$(ZArticulo, 2) = "DW" Then
                                        WProducto = "DW" + Mid$(ZArticulo, 3, 10)
                                            Else
                                        If Left$(ZArticulo, 2) = "SE" Then
                                            WProducto = "SE" + Mid$(ZArticulo, 3, 10)
                                                Else
                                            WProducto = "PT" + Mid$(ZArticulo, 3, 10)
                                        End If
                                    End If
                                        
                                    Select Case Val(WEmpresa)
                                        Case 1, 3, 5, 6, 7, 10, 11
                                            WEmpresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            WEmpresa = "0004"
                                            txtOdbc = "Empresa04"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                        
                                    LlamaImprime = "N"
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM EspecifUnificaVersion"
                                    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
                                    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                                    spEspecifUnificaVersion = ZSql
                                    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEspecifUnificaVersion.RecordCount > 0 Then
                                        With rstEspecifUnificaVersion
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    
                                                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                                                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                                                            
                                                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                                
                                                    End If
                                                        
                                                    If WDesde > WFechaord And LlamaImprime = "N" Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                    End If
                                                    
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstEspecifUnificaVersion.Close
                                    End If
                                    
                                    If LlamaImprime = "N" Then
                                    
                                        Sql1 = "Select *"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                
                                            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
                                            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
                                            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
                                            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
                                            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
                                            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
                                            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
                                            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
                                            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
                                            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
                                                                
                                            ZStd(1, 1) = rstEspecifUnifica!Valor1
                                            ZStd(2, 1) = rstEspecifUnifica!valor2
                                            ZStd(3, 1) = rstEspecifUnifica!Valor3
                                            ZStd(4, 1) = rstEspecifUnifica!valor4
                                            ZStd(5, 1) = rstEspecifUnifica!valor5
                                            ZStd(6, 1) = rstEspecifUnifica!valor6
                                            ZStd(7, 1) = rstEspecifUnifica!valor7
                                            ZStd(8, 1) = rstEspecifUnifica!valor8
                                            ZStd(9, 1) = rstEspecifUnifica!valor9
                                            ZStd(10, 1) = rstEspecifUnifica!valor10
                                                                
                                            ZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                                            ZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                                            ZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                                            ZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                                            ZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                                            ZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                                            ZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                                            ZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                                            ZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                                            ZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                                                                
                                            ZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                                            ZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                                            ZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                                            ZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                                            ZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                                            ZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                                            ZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                                            ZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                                            ZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                                            ZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                                    
                                            ZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                                            ZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!HAsta2), "", rstEspecifUnifica!HAsta2)
                                            ZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                                            ZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                                            ZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                                            ZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                                            ZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                                            ZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                                            ZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                                            ZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                                                
                                            ZVersion = rstEspecifUnifica!Version
                                            rstEspecifUnifica.Close
                                            LlamaImprime = "S"
                                        End If
                                    
                                    End If
                                    
                                    If LlamaImprime = "S" Then
                                        
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(1) = rstEnsayo!Descripcion
                                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(2) = rstEnsayo!Descripcion
                                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(3) = rstEnsayo!Descripcion
                                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(4) = rstEnsayo!Descripcion
                                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(5) = rstEnsayo!Descripcion
                                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(6) = rstEnsayo!Descripcion
                                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(7) = rstEnsayo!Descripcion
                                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(8) = rstEnsayo!Descripcion
                                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(9) = rstEnsayo!Descripcion
                                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(10) = rstEnsayo!Descripcion
                                            ZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                                                
                                        Call Conecta_Empresa
                                        
                                        XEmpresa = WEmpresa
                                        Select Case Val(XEmpresa)
                                            Case 1, 3, 5, 6, 7, 10, 11
                                                WEmpresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case 2, 4, 8, 9
                                                WEmpresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                        End Select
                                        
                                        ZImpreVto = 0
                                        ZRazon = ""
                                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstCliente.RecordCount > 0 Then
                                            ZRazon = Left$(rstCliente!Razon, 50)
                                            ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                                            rstCliente.Close
                                        End If
                                        
                                        ZZImpreVtoTermi = 0
                                        spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
                                            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
                                            rstTerminado.Close
                                        End If
                                        
                                        If ZZImpreVtoTermi = 0 Then
                                            If ZImpreVto <> 1 Then
                                                Rem WFechaElaboracion = ""
                                            End If
                                        End If
                                            
                                        ZCliente = UCase(ZCliente)
                                        ZArticulo = UCase(ZArticulo)
                                        ZClave = ZCliente + ZArticulo
                        
                                        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
                                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstPrecios.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
                                            rstPrecios.Close
                                        End If
                                        
                                        Call Conecta_Empresa
                                                
                                        ZSql = "DELETE Certificado"
                                        spCertificado = ZSql
                                        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                            
                                        LugarMetodo = 0
                                                
                                        For CiclaMetodo = 1 To 10
                                                
                                            If ZOpcion(CiclaMetodo) = 1 Then
                                                
                                                LugarMetodo = LugarMetodo + 1
                                                    
                                                ZOrden = ""
                                                ZClave1 = ZLote
                                                Call Ceros(ZClave1, 6)
                                                ZClave2 = Str$(LugarMetodo)
                                                Call Ceros(ZClave2, 2)
                                                ZClave = ZClave1 + ZClave2
                                                ZMetodo = ZEnsayo(CiclaMetodo)
                                                
                                                If Val(ZStd(CiclaMetodo, 3)) <> 0 And Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                                    ZValorNormalI = " " + Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo))
                                                    ZValorNormalII = ""
                                                        Else
                                                    ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                End If
                                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 50)
                                                
                                                ZValorNormalI = Trim(ZValorNormalI)
                                                ZCanti = 50 - Len(ZValorNormalI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalI = Space$(ZCanti) + ZValorNormalI
                                                
                                                ZValorNormalII = Trim(ZValorNormalII)
                                                ZCanti = 50 - Len(ZValorNormalII)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalII = Space$(ZCanti) + ZValorNormalII
                                                
                                                ZValorPartidaI = Trim(ZValorPartidaI)
                                                ZCanti = 50 - Len(ZValorPartidaI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorPartidaI = Space$(ZCanti) + ZValorPartidaI
                                                
                                                ZValorPartidaII = ""
                                                ZObservacionesI = ""
                                                ZObservacionesII = ""
                                                ZObservacionesIII = "Version " + ZVersion
                                                ZObservacionesIV = ""
                                                ZObservacionesV = ""
                                                ZObservacionesVI = ""
                                                If Val(WEmpresa) = 1 Then
                                                    ZEmpresa = "Surfactan S.A."
                                                        Else
                                                    ZEmpresa = "Pellital S.A."
                                                End If
                                                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                                ZFechaII = WFechaElaboracion
                                                
                                                ZExamen = Trim(ZDescri(CiclaMetodo))
                                                ZExamenII = ""
                                                ZHasta = Len(Trim(ZExamen))
                                                If ZHasta > 25 Then
                                                    For Cicla = ZHasta To 1 Step -1
                                                        If Mid(ZExamen, Cicla, 1) = Space(1) Then
                                                            ZExamenII = Mid(ZExamen, Cicla - Desde + 1, 25)
                                                            ZExamen = Mid(ZExamen, 1, Cicla - Desde)
                                                            Exit For
                                                        End If
                                                    Next Cicla
                                                End If
                                                        
                                                        
                                                        
                                                aa = Auxiliar(1, 1)
                                                aa = Auxiliar(1, 2)
                                                aa = Auxiliar(1, 3)
                                                aa = Auxiliar(1, 4)
                                                aa = Auxiliar(1, 5)
                                                aa = Auxiliar(1, 6)
                                                aa = Auxiliar(1, 7)
                                                aa = Auxiliar(1, 8)
                                                aa = Auxiliar(1, 9)
                                                aa = Auxiliar(1, 10)
                                                aa = Auxiliar(1, 11)
                                                aa = Auxiliar(1, 12)
                                                aa = Auxiliar(1, 13)
                                                aa = Auxiliar(1, 14)
                                                aa = Auxiliar(1, 15)
                                                aa = Auxiliar(1, 16)
                                                aa = Auxiliar(1, 17)
                                                aa = Auxiliar(1, 18)
                                                aa = Auxiliar(1, 19)
                                                aa = Auxiliar(1, 20)
                                                        
                                                        
                                                ZSql = ""
                                                ZSql = ZSql + "INSERT INTO Certificado ("
                                                ZSql = ZSql + "Clave ,"
                                                ZSql = ZSql + "Partida ,"
                                                ZSql = ZSql + "Renglon ,"
                                                ZSql = ZSql + "Razon ,"
                                                ZSql = ZSql + "Orden ,"
                                                ZSql = ZSql + "Terminado ,"
                                                ZSql = ZSql + "Descripcion ,"
                                                ZSql = ZSql + "Fecha ,"
                                                ZSql = ZSql + "FechaII ,"
                                                ZSql = ZSql + "Cantidad ,"
                                                ZSql = ZSql + "Examen ,"
                                                ZSql = ZSql + "ExamenII ,"
                                                ZSql = ZSql + "ValorPartidaI ,"
                                                ZSql = ZSql + "ValorPartidaII ,"
                                                ZSql = ZSql + "ValorNormalI ,"
                                                ZSql = ZSql + "ValorNormalII ,"
                                                ZSql = ZSql + "Observaciones1 ,"
                                                ZSql = ZSql + "Observaciones2 ,"
                                                ZSql = ZSql + "Observaciones3 ,"
                                                ZSql = ZSql + "Observaciones4 ,"
                                                ZSql = ZSql + "Observaciones5 ,"
                                                ZSql = ZSql + "Observaciones6 ,"
                                                ZSql = ZSql + "Metodo ,"
                                                ZSql = ZSql + "Empresa )"
                                                ZSql = ZSql + "Values ("
                                                ZSql = ZSql + "'" + ZClave + "',"
                                                ZSql = ZSql + "'" + ZLote + "',"
                                                ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                                ZSql = ZSql + "'" + ZRazon + "',"
                                                ZSql = ZSql + "'" + ZOrden + "',"
                                                ZSql = ZSql + "'" + ZArticulo + "',"
                                                ZSql = ZSql + "'" + ZDesArticulo + "',"
                                                ZSql = ZSql + "'" + ZFecha + "',"
                                                ZSql = ZSql + "'" + ZFechaII + "',"
                                                ZSql = ZSql + "'" + ZCantidad + "',"
                                                ZSql = ZSql + "'" + ZExamen + "',"
                                                ZSql = ZSql + "'" + ZExamenII + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                                ZSql = ZSql + "'" + ZValorNormalI + "',"
                                                ZSql = ZSql + "'" + ZValorNormalII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesI + "',"
                                                ZSql = ZSql + "'" + ZObservacionesII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                                ZSql = ZSql + "'" + ZMetodo + "',"
                                                ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                                spCertificado = ZSql
                                                Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                        
                                            End If
                                                            
                                        Next CiclaMetodo
                                            
                                        Do
                                            
                                            If LugarMetodo = 10 Then
                                                Exit Do
                                            End If
                                                
                                            LugarMetodo = LugarMetodo + 1
                                                    
                                            ZOrden = ""
                                            ZClave1 = ZLote
                                            Call Ceros(ZClave1, 6)
                                            ZClave2 = Str$(LugarMetodo)
                                            Call Ceros(ZClave2, 2)
                                            ZClave = ZClave1 + ZClave2
                                            ZMetodo = ""
                                            ZExamen = ""
                                            ZValorNormalI = ""
                                            ZValorNormalII = ""
                                            ZValorPartidaI = ""
                                            ZValorPartidaII = ""
                                            ZObservacionesI = ""
                                            ZObservacionesII = ""
                                            ZObservacionesIII = "Version " + ZVersion
                                            ZObservacionesIV = ""
                                            ZObservacionesV = ""
                                            ZObservacionesVI = ""
                                            If Val(WEmpresa) = 1 Then
                                                ZEmpresa = "Surfactan S.A."
                                                    Else
                                                ZEmpresa = "Pellital S.A."
                                            End If
                                            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                            ZFechaII = WFechaElaboracion
                                            ZExamenII = ""
                                                        
                                            ZSql = ""
                                            ZSql = ZSql + "INSERT INTO Certificado ("
                                            ZSql = ZSql + "Clave ,"
                                            ZSql = ZSql + "Partida ,"
                                            ZSql = ZSql + "Renglon ,"
                                            ZSql = ZSql + "Razon ,"
                                            ZSql = ZSql + "Orden ,"
                                            ZSql = ZSql + "Terminado ,"
                                            ZSql = ZSql + "Descripcion ,"
                                            ZSql = ZSql + "Fecha ,"
                                            ZSql = ZSql + "Cantidad ,"
                                            ZSql = ZSql + "Examen ,"
                                            ZSql = ZSql + "ValorPartidaI ,"
                                            ZSql = ZSql + "ValorPartidaII ,"
                                            ZSql = ZSql + "ValorNormalI ,"
                                            ZSql = ZSql + "ValorNormalII ,"
                                            ZSql = ZSql + "Observaciones1 ,"
                                            ZSql = ZSql + "Observaciones2 ,"
                                            ZSql = ZSql + "Observaciones3 ,"
                                            ZSql = ZSql + "Observaciones4 ,"
                                            ZSql = ZSql + "Observaciones5 ,"
                                            ZSql = ZSql + "Observaciones6 ,"
                                            ZSql = ZSql + "Metodo ,"
                                            ZSql = ZSql + "Empresa )"
                                            ZSql = ZSql + "Values ("
                                            ZSql = ZSql + "'" + ZClave + "',"
                                            ZSql = ZSql + "'" + ZLote + "',"
                                            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                            ZSql = ZSql + "'" + ZRazon + "',"
                                            ZSql = ZSql + "'" + ZOrden + "',"
                                            ZSql = ZSql + "'" + ZArticulo + "',"
                                            ZSql = ZSql + "'" + ZDesArticulo + "',"
                                            ZSql = ZSql + "'" + ZFecha + "',"
                                            ZSql = ZSql + "'" + ZCantidad + "',"
                                            ZSql = ZSql + "'" + ZExamen + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                            ZSql = ZSql + "'" + ZValorNormalI + "',"
                                            ZSql = ZSql + "'" + ZValorNormalII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesI + "',"
                                            ZSql = ZSql + "'" + ZObservacionesII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                            ZSql = ZSql + "'" + ZMetodo + "',"
                                            ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                            spCertificado = ZSql
                                            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                
                                        Loop
                                                
                                        Listado.WindowTitle = "Certificado de Analisis"
                                        Listado.WindowTop = 0
                                        Listado.WindowLeft = 0
                                        Listado.WindowWidth = Screen.Width
                                        Listado.WindowHeight = Screen.Height
                        
                                        Listado.Destination = 1
                                        Rem Listado.Destination = 0
                                                
                                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                            Listado.ReportFileName = "Certificado.rpt"
                                                Else
                                            Listado.ReportFileName = "CertificadoPelli.rpt"
                                        End If
                                                    
                                        DbConnect = db.Connect
                                        DSQ = getDatabase(DbConnect)
                        
                                        Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                                        + "From " _
                                                        + DSQ + ".dbo.Certificado Certificado " _
                                                        + "Where " _
                                                        + "Certificado.Partida >= 0 AND " _
                                                        + "Certificado.Partida <= 999999"
                                                        
                                        Listado.Destination = 1
                                        Rem Listado.Destination = 0
                                        Listado.CopiesToPrinter = 1
                        
                                        Listado.Connect = Connect()
                                        Listado.Action = 1
                                                
                                    End If
                                          
                                End If
                                    
                            Next ZCiclo
                            
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                            
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                
                        Else
                        
                    If Left$(Articulo, 2) = "DY" Then
                        
                        Auxi = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                        ZZCambia = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM Laudo"
                        ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
                        ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZPartiOri = Trim(rstLaudo!PartiOri)
                            rstLaudo.Close
                            ZZCambia = "S"
                            ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                            ZZEstado = Dir(ZZRuta)
                            ZZEstado = Trim(ZZEstado)
                            If ZZEstado <> "" Then
                                Select Case ZZVersion
                                    Case 1
                                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                    Case 2
                                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Case 3
                                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Case 4
                                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Case 5
                                        RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Case Else
                                        RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                End Select
                                    Else
                                m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                            End If
                        End If
                        
                        If ZZCambia = "N" Then
                                        
                            XEmpresa = WEmpresa
                                    
                            WEmpresa = "0006"
                            txtOdbc = "Empresa06"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM Laudo"
                            ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
                            ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                ZZPartiOri = Trim(rstLaudo!PartiOri)
                                rstLaudo.Close
                                ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                                ZZEstado = Dir(ZZRuta)
                                ZZEstado = Trim(ZZEstado)
                                If ZZEstado <> "" Then
                                    Select Case ZZVersion
                                        Case 1
                                            RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                        Case 2
                                            RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                        Case 3
                                            RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                        Case 4
                                            RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                        Case 5
                                            RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                        Case Else
                                            RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    End Select
                                        Else
                                    m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                    a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                End If
                            End If
                            
                            Call Conecta_Empresa
                        
                        End If
                        
                    End If
                    
                End If
            End If
                
        Next ZZCiclo
        
    Next DA
    
    For ZZCicloFicha = 1 To ZLugarFicha
    
        ZZIntervencion = ZImpreFicha(ZZCicloFicha)
        
        Call Ceros(ZZIntervencion, 3)
        
        ZZRuta = "w:\FICHASIS\GUIADEEMERGENCIA" + ZZIntervencion + ".PDF"
        ZZEstado = Dir(ZZRuta)
        ZZEstado = Trim(ZZEstado)
        If ZZEstado <> "" Then
            Select Case ZZVersion
                Case 1
                    RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Case 2
                    RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Case 3
                    RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Case 4
                    RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Case 5
                    RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                Case Else
                    RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
            End Select
            Rem RetVal = Shell("cmd.exe /c Taskkill /f /IM AcroRd32.exe", 1)
                Else
            m$ = "El articulo " + Articulo + " posee la Ficha de Emergencia Nro " + ZZIntervencion + " y no se ha encontrado"
            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
        End If
    
    Next ZZCicloFicha
    
    PrgFactu.Show
    Numero.SetFocus
    


