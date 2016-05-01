VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedidoSol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Solicitud de Pedidos de Ventas"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   300
   ClientWidth     =   11775
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11775
   Visible         =   0   'False
   Begin VB.Frame EntraNombreMl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1560
      TabIndex        =   109
      Top             =   3840
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox Referencia 
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
         MaxLength       =   50
         TabIndex        =   114
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox OrdenTrabajo 
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
         MaxLength       =   50
         TabIndex        =   112
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox NombreComercialMl 
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
         MaxLength       =   50
         TabIndex        =   110
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label32 
         Caption         =   "Referencia"
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
         Left            =   240
         TabIndex        =   115
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Orden Trabajo"
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
         Left            =   240
         TabIndex        =   113
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "Nombre Comercial"
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
         Left            =   240
         TabIndex        =   111
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid PantallaPro 
      Height          =   855
      Left            =   6120
      TabIndex        =   98
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin VB.Frame EntraNombre 
      Caption         =   " Nombre Comecial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   103
      Top             =   2520
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox NombreComercial 
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
         TabIndex        =   104
         Top             =   360
         Width           =   7695
      End
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
      Left            =   7320
      TabIndex        =   102
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame PantaDirEntrega 
      Caption         =   "Seleccion de Lugar de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   96
      Top             =   3600
      Visible         =   0   'False
      Width           =   8895
      Begin VB.ListBox ListaDirEntrega 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   97
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame IngreEspe 
      Caption         =   "Ingreso de Especificaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   94
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox Especificaciones 
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
         MaxLength       =   30
         TabIndex        =   95
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame XClave 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3120
      TabIndex        =   81
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
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
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   83
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
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
         Left            =   840
         TabIndex        =   82
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Este   pedido   al     ser modificado  debera  ser nuevamente autorizado para la facturacion  del mismo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   85
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Ingrese su Password"
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
         Left            =   840
         TabIndex        =   84
         Top             =   360
         Width           =   1935
      End
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
      Height          =   1260
      Left            =   5040
      TabIndex        =   80
      Top             =   6840
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   5520
      TabIndex        =   67
      Text            =   " "
      Top             =   6480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton BorraConsulta 
      Caption         =   "Borra Consulta"
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
      Left            =   10800
      TabIndex        =   66
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton ConsultaPro 
      Caption         =   "Consulta Producto"
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
      Left            =   9720
      TabIndex        =   58
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton ConsultaCli 
      Caption         =   "Consulta Cliente"
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
      Left            =   10800
      TabIndex        =   65
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Baja 
      Caption         =   "  Baja  Pedido"
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
      Left            =   9720
      TabIndex        =   64
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame IngreEnvases 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   9120
      TabIndex        =   36
      Top             =   3600
      Width           =   2535
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   44
         Text            =   " "
         Top             =   2760
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   43
         Text            =   " "
         Top             =   2520
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   42
         Text            =   " "
         Top             =   2280
         Width           =   855
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
         Left            =   360
         MaxLength       =   4
         TabIndex        =   41
         Text            =   " "
         Top             =   2760
         Width           =   855
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
         Left            =   360
         MaxLength       =   4
         TabIndex        =   40
         Text            =   " "
         Top             =   2520
         Width           =   855
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
         Left            =   360
         MaxLength       =   4
         TabIndex        =   39
         Text            =   " "
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99 - Env. Indistinto"
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
         Left            =   360
         TabIndex        =   108
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label WAbre6 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   63
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label WAbre5 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   62
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label WAbre4 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   61
         Top             =   840
         Width           =   855
      End
      Begin VB.Label WAbre3 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   60
         Top             =   600
         Width           =   855
      End
      Begin VB.Label WAbre2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.Label WAbre1 
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   57
         Top             =   120
         Width           =   855
      End
      Begin VB.Label WCapa6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label WCapa5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   55
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label WCapa4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   54
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Wcapa3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   53
         Top             =   600
         Width           =   855
      End
      Begin VB.Label WCapa2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
      Begin VB.Label WCapa1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   51
         Top             =   120
         Width           =   855
      End
      Begin VB.Label WEnvase6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label WEnvase5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label WEnvase4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   615
      End
      Begin VB.Label WEnvase3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   615
      End
      Begin VB.Label WEnvase2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.Label WEnvase1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label11 
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
         Left            =   1320
         TabIndex        =   38
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Envase 
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
         Left            =   360
         TabIndex        =   37
         Top             =   2040
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8280
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Frame Datos 
      BackColor       =   &H8000000A&
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      TabIndex        =   29
      Top             =   6360
      Width           =   4215
      Begin VB.CommandButton AvisoError 
         Caption         =   "Sistema sin Conexion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         Picture         =   "prgpedidoSol.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Wstock7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   119
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Stock7 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   118
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Wstock6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   117
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Stock6 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   116
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Produccion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1320
         TabIndex        =   101
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Produccion"
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
         TabIndex        =   100
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label WStock4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   93
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Stock4 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   92
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Stock5 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   91
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label WStock5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   90
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label WStock3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   89
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Stock3 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   88
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Stock2 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   71
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Stock1 
         Caption         =   "Stock"
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
         Left            =   2400
         TabIndex        =   70
         Top             =   120
         Width           =   615
      End
      Begin VB.Label WStock2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   69
         Top             =   360
         Width           =   975
      End
      Begin VB.Label WStock1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   68
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Disponible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1320
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.Label StkPedido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1320
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Stock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1320
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible"
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Stock"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
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
      Left            =   10800
      TabIndex        =   28
      Top             =   960
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
      Height          =   1260
      Left            =   4800
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   25
      Text            =   " "
      Top             =   1200
      Width           =   7695
   End
   Begin VB.TextBox Hora 
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
      MaxLength       =   5
      TabIndex        =   23
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox FecEntrega 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
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
   Begin VB.TextBox Cliente 
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
      TabIndex        =   15
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   9720
      TabIndex        =   10
      Top             =   0
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
      Left            =   10800
      TabIndex        =   9
      Top             =   480
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
      Left            =   9720
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
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
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   8895
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
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
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   4
         Text            =   " "
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
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   240
         Width           =   1215
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
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   4095
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
      Left            =   10800
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8640
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox OrdenCpa 
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
      TabIndex        =   77
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Precio 
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox Tipoped 
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
      Left            =   7440
      TabIndex        =   73
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Version 
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Inserta 
      Caption         =   "Inserta Renglon"
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
      Left            =   9720
      TabIndex        =   87
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton WImpres 
      Caption         =   "Impresion de Pedidos"
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
      Left            =   9720
      TabIndex        =   86
      Top             =   2760
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3735
      Left            =   120
      TabIndex        =   105
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
      _Version        =   327680
   End
   Begin VB.ComboBox Via 
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
      Left            =   7440
      TabIndex        =   106
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "Mod. Precio"
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
      Left            =   3720
      TabIndex        =   79
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label16 
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
      TabIndex        =   76
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Version"
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
      Left            =   6360
      TabIndex        =   74
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Tipo Pedido"
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
      Left            =   6360
      TabIndex        =   72
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Hora"
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
      TabIndex        =   22
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha Entrega"
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
      TabIndex        =   20
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesPago 
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
      Left            =   6840
      TabIndex        =   19
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Pago 
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
      Left            =   6000
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "C.Pago"
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
      TabIndex        =   17
      Top             =   840
      Width           =   975
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
      Left            =   3120
      TabIndex        =   16
      Top             =   480
      Width           =   3135
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
      TabIndex        =   14
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
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de pedido"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1815
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
      Left            =   6360
      TabIndex        =   107
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "PrgPedidoSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Auxi As String
Private WImpre(10) As String
Private WEnvase(10) As String
Private WVector(6, 4) As String
Private XEnvase(100, 6) As String
Private XEspecificaciones(100) As String
Private XDatosMuestra(100, 3) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private WTermi As String
Private WStkPedido As Double
Private WProduccion As Double
Dim ZZPasaMsds As String

Dim rstOt As Recordset
Dim spOt As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedidoSol As Recordset
Dim spPedidoSol As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstEspeCli As Recordset
Dim spEspeCli As String
Dim rstFeriado As Recordset
Dim spFeriado As String
Dim rstImprePed As Recordset
Dim spImprePed As String
Dim WGraba As String
Dim XParam As String
Dim WVersion As String
Dim WPasa(100, 6) As String
Dim IngreVector(20000, 5) As String
Dim EntraVector As Integer
Dim Partida(100, 3) As String
Dim LugarPartida As String
Dim WSaldo As Double
Dim ImpreEnvase(10) As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Dim WCalcula As String
Dim ZPrecio As String
Dim ZProceso As Integer
Dim ZEstado As String

Dim ZClave As String
Dim ZTipo As String
Dim ZPedido As String
Dim ZRenglon As String
Dim ZEmpresa As String
Dim ZVersion As String
Dim ZCliente As String
Dim ZNombre As String
Dim ZFecha As String
Dim ZFechaent As String
Dim ZTipoPedido As String
Dim ZCondicion As String
Dim ZEntrega As String
Dim ZObservaciones1 As String
Dim ZObservaciones2 As String
Dim ZOrden As String
Dim ZArticulo As String
Dim ZDescripcion As String
Dim ZCantidad As String
Dim ZEnvase As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim WEspecif(100) As String
Dim ZZRenglon As Integer

Dim Producto As String
Dim Costo As Double
Dim ZTipoCosto As Integer
Dim ZSedronar As Integer
Dim ZNroSedronar As String

Dim ZZRequiereCertificado As String
Dim ZZRequiereMsds As String
Dim ZZRequiereMsdsCada As String
Dim ZZRequiereHoja As String
Dim ZZPermiteParcial As String
Dim ZZPartidasVarias As String

Dim ZZEmailCertificado As String
Dim ZZEmailMsds As String
Dim ZZEmailHoja As String
Dim ZZDiasI As String
Dim ZZDiasII As String
Dim ZZDiasIII As String
Dim ZZEnvasesI As String
Dim ZZEnvasesII As String
Dim ZZEnvasesIII As String
Dim ZZEtiquetaI As String
Dim ZZEtiquetaII As String
Dim ZZEspecif1 As String
Dim ZZEspecif2 As String
Dim ZZEspecif3 As String
Dim ZZEspecif4 As String
Dim ZZEspecif5 As String
Dim ZZCantidadPartidas As String

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim WNombreEmail As String
Dim MAttach As String

Private Sub Baja_Click()

    ZSql = ""
    ZSql = ZSql + "DELETE PedidoSol"
    ZSql = ZSql + " Where PedidoSol.Pedido = " + "'" + Pedido.Text + "'"
    spPedidoSol = ZSql
    Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Limpia_Click
    Pedido.SetFocus

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Borra_Click()

    WVector1.Col = 1
    WVector1.Text = ""
    
    WVector1.Col = 2
    WVector1.Text = ""

    WVector1.Col = 3
    WVector1.Text = ""
    
    WVector1.Col = 4
    WVector1.Text = ""
    
    WLugar = WVector1.Row
    
    XEnvase(WLugar, 1) = ""
    XEnvase(WLugar, 2) = ""
    XEnvase(WLugar, 3) = ""
    XEnvase(WLugar, 4) = ""
    XEnvase(WLugar, 5) = ""
    XEnvase(WLugar, 6) = ""
    
    XEspecificaciones(WLugar) = ""
    
    XDatosMuestra(WLugar, 1) = ""
    XDatosMuestra(WLugar, 2) = ""
    XDatosMuestra(WLugar, 3) = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLinea.Text = ""
    
    NombreComercialMl.Text = ""
    OrdenTrabajo.Text = ""
    Referencia.Text = ""
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    
    Especificaciones.Text = ""

    WStock1.Caption = ""
    WStock2.Caption = ""
    WStock3.Caption = ""
    WStock4.Caption = ""
    WStock5.Caption = ""
    Wstock6.Caption = ""
    Wstock7.Caption = ""
    Stock.Caption = ""
    StkPedido.Caption = ""
    Produccion.Caption = ""
    Disponible.Caption = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub BorraConsulta_Click()
    Pantalla.Visible = False
End Sub

Private Sub cmdClose_Click()
    PrgPedidoSol.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub ListaDirEntrega_Click()
    ZLugarDirEntrega = ListaDirEntrega.ListIndex + 1
    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
    PantaDirEntrega.Visible = False
    Tipoped.SetFocus
End Sub

Private Sub ConsultaCli_Click()

    XIndice = 0

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    
    Ayuda.Height = 285
    Ayuda.Left = 2040
    Ayuda.Top = 0
    Ayuda.Width = 8055
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    spClientes = "ListaClienteConsulta"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then

    With rstClientes
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                Pantalla.AddItem IngresaItem
                IngresaItem = rstClientes!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstClientes.Close

    End If
    
    Pantalla.Visible = True
    
    Pantalla.Height = 7740
    Pantalla.Left = 2040
    Pantalla.Top = 360
    Pantalla.Width = 8175
    
    Ayuda.SetFocus
    
End Sub

Private Sub ConsultaProOLD_Click()

    XIndice = 1
    
    Dim IngresaItem As String
    WIndice.Clear
    
    Call Limpia_PantallaPro
    LugarPantalla = 0
    
    Ayuda.Height = 200
    Ayuda.Left = 4100
    Ayuda.Top = 6350
    Ayuda.Width = 3200
    
    Estado.Clear
    
    Estado.AddItem "Activo"
    Estado.AddItem "Historico"
    Estado.AddItem "Cotizacion"
    
    Estado.ListIndex = 0
    
    Ayuda.Text = ""
    
    Sql1 = "Select Cliente, Terminado, Descripcion, Precio, Fecha, Estado"
    Sql2 = " FROM Precios"
    Sql3 = " Where Precios.Cliente = " + "'" + Cliente.Text + "'"
    Sql4 = " Order by Terminado"
    spPrecios = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPrecios!Cliente Then
                        ZEstado = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
                        If rstPrecios!Precio <> Null Then
                            ZPrecio = Str$(rstPrecios!Precio)
                                Else
                            ZPrecio = IIf(IsNull(rstPrecios!Precio), "0", Str$(rstPrecios!Precio))
                        End If
                        If ZEstado = 0 And Val(ZPrecio) <> 0 Then
                            ZTerminado = rstPrecios!Terminado
                            ZDescripcion = rstPrecios!Descripcion
                            ZFecha = IIf(IsNull(rstPrecios!Fecha), "  /  /    ", rstPrecios!Fecha)
                            ZPrecio = Pusing("###,###.##", ZPrecio)
                            LugarPantalla = LugarPantalla + 1
                            PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
                            PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
                            PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
                            PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
                            IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                            WIndice.AddItem IngresaItem
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    
    
    Erase IngreVector
    EntraVector = 0
    
    
    Sql1 = "Select Cliente, Articulo, Precio, Fecha, Estado"
    Sql2 = " FROM PreciosMp"
    Sql3 = " Where PreciosMp.Cliente = " + "'" + Cliente.Text + "'"
    Sql4 = " Order by Articulo"
    spPreciosMp = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
    
        With rstPreciosMp
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPreciosMp!Cliente Then
                        ZEstado = IIf(IsNull(rstPreciosMp!Estado), "0", rstPreciosMp!Estado)
                        If ZEstado = 0 Then
                            ZArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                            EntraVector = EntraVector + 1
                            IngreVector(EntraVector, 1) = ZArticulo
                            IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                            IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                            IngreVector(EntraVector, 4) = IIf(IsNull(rstPreciosMp!Precio), "0", Str$(rstPreciosMp!Precio))
                            IngreVector(EntraVector, 5) = IIf(IsNull(rstPreciosMp!Fecha), "", rstPreciosMp!Fecha)
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosMp.Close
    End If
    
    For CicloVector = 1 To EntraVector
        
        ZTerminado = IngreVector(CicloVector, 1)
        WCliente = IngreVector(CicloVector, 2)
        WArti = IngreVector(CicloVector, 3)
        ZPrecio = IngreVector(CicloVector, 4)
        ZPrecio = Pusing("###,###.##", ZPrecio)
        ZFecha = IngreVector(CicloVector, 5)
        ZDescripcion = ""
        
        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        LugarPantalla = LugarPantalla + 1
        PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
        PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
        PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
        PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
        
        IngresaItem = WCliente + WArti
        WIndice.AddItem IngresaItem
        
    Next CicloVector
    
    Rem PantallaPro.Height = 1740
    
    PantallaPro.Height = 1650
    PantallaPro.Left = 4100
    PantallaPro.Top = 6720
    PantallaPro.Width = 7575
    
    PantallaPro.Col = 1
    PantallaPro.Row = 1
    PantallaPro.TopRow = 1
    
    PantallaPro.Visible = True
    Estado.Visible = True
    Ayuda.Visible = True
    
    Ayuda.SetFocus

End Sub

Private Sub ConsultaPro_Click()

    Estado.Clear
    
    Estado.AddItem "Activo"
    Estado.AddItem "Historico"
    Estado.AddItem "Cotizacion"
    
    Estado.ListIndex = 0

End Sub

Private Sub Estado_click()

    XIndice = 1
    
    Dim IngresaItem As String
    WIndice.Clear
    
    Call Limpia_PantallaPro
    LugarPantalla = 0
    
    Ayuda.Height = 200
    Ayuda.Left = 4100
    Ayuda.Top = 6350
    Ayuda.Width = 3200
    
    Ayuda.Text = ""
    
    Sql1 = "Select Cliente, Terminado, Descripcion, Precio, Fecha, Estado"
    Sql2 = " FROM Precios"
    Sql3 = " Where Precios.Cliente = " + "'" + Cliente.Text + "'"
    Sql4 = " Order by Terminado"
    spPrecios = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPrecios!Cliente Then
                        ZEstado = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
                        If rstPrecios!Precio <> Null Then
                            ZPrecio = Str$(rstPrecios!Precio)
                                Else
                            ZPrecio = IIf(IsNull(rstPrecios!Precio), "0", Str$(rstPrecios!Precio))
                        End If
                        If Val(ZEstado) = Estado.ListIndex And Val(ZPrecio) <> 0 Then
                            ZTerminado = rstPrecios!Terminado
                            ZDescripcion = rstPrecios!Descripcion
                            ZFecha = IIf(IsNull(rstPrecios!Fecha), "  /  /    ", rstPrecios!Fecha)
                            ZPrecio = Pusing("###,###.##", ZPrecio)
                            LugarPantalla = LugarPantalla + 1
                            PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
                            PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
                            PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
                            PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
                            IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                            WIndice.AddItem IngresaItem
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    
    
    Erase IngreVector
    EntraVector = 0
    
    
    Sql1 = "Select Cliente, Articulo, Precio, Fecha, Estado"
    Sql2 = " FROM PreciosMp"
    Sql3 = " Where PreciosMp.Cliente = " + "'" + Cliente.Text + "'"
    Sql4 = " Order by Articulo"
    spPreciosMp = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
    
        With rstPreciosMp
            .MoveFirst
            Do
                If .EOF = False Then
                    If Cliente.Text = rstPreciosMp!Cliente Then
                        ZEstado = IIf(IsNull(rstPreciosMp!Estado), "0", rstPreciosMp!Estado)
                        If Val(ZEstado) = Estado.ListIndex Then
                            ZArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                            EntraVector = EntraVector + 1
                            IngreVector(EntraVector, 1) = ZArticulo
                            IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                            IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                            IngreVector(EntraVector, 4) = IIf(IsNull(rstPreciosMp!Precio), "0", Str$(rstPreciosMp!Precio))
                            IngreVector(EntraVector, 5) = IIf(IsNull(rstPreciosMp!Fecha), "", rstPreciosMp!Fecha)
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosMp.Close
    End If
    
    For CicloVector = 1 To EntraVector
        
        ZTerminado = IngreVector(CicloVector, 1)
        WCliente = IngreVector(CicloVector, 2)
        WArti = IngreVector(CicloVector, 3)
        ZPrecio = IngreVector(CicloVector, 4)
        ZPrecio = Pusing("###,###.##", ZPrecio)
        ZFecha = IngreVector(CicloVector, 5)
        ZDescripcion = ""
        
        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        LugarPantalla = LugarPantalla + 1
        PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
        PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
        PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
        PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
        
        IngresaItem = WCliente + WArti
        WIndice.AddItem IngresaItem
        
    Next CicloVector
    
    Rem PantallaPro.Height = 1740
    
    PantallaPro.Height = 1650
    PantallaPro.Left = 4100
    PantallaPro.Top = 6720
    PantallaPro.Width = 7575
    
    PantallaPro.Col = 1
    PantallaPro.Row = 1
    PantallaPro.TopRow = 1
    
    PantallaPro.Visible = True
    Estado.Visible = True
    Ayuda.Visible = True

    Ayuda.SetFocus
    
End Sub

Private Sub Inserta_Click()

    WLugar = WVector1.Row
    
    For Cicla = 79 To WLugar + 1 Step -1
    
        XEnvase(Cicla, 1) = XEnvase(Cicla - 1, 1)
        XEnvase(Cicla, 2) = XEnvase(Cicla - 1, 2)
        XEnvase(Cicla, 3) = XEnvase(Cicla - 1, 3)
        XEnvase(Cicla, 4) = XEnvase(Cicla - 1, 4)
        XEnvase(Cicla, 5) = XEnvase(Cicla - 1, 5)
        XEnvase(Cicla, 6) = XEnvase(Cicla - 1, 6)
        
        XEspecificaciones(Cicla) = XEspecificaciones(Cicla - 1)
        
        XDatosMuestra(Cicla, 1) = XDatosMuestra(Cicla - 1, 1)
        XDatosMuestra(Cicla, 2) = XDatosMuestra(Cicla - 1, 2)
        XDatosMuestra(Cicla, 3) = XDatosMuestra(Cicla - 1, 3)
        
        WVector1.TextMatrix(Cicla, 1) = WVector1.TextMatrix(Cicla - 1, 1)
        WVector1.TextMatrix(Cicla, 2) = WVector1.TextMatrix(Cicla - 1, 2)
        WVector1.TextMatrix(Cicla, 3) = WVector1.TextMatrix(Cicla - 1, 3)
        WVector1.TextMatrix(Cicla, 4) = WVector1.TextMatrix(Cicla - 1, 4)
        
    Next Cicla
    
    XEnvase(WLugar, 1) = ""
    XEnvase(WLugar, 2) = ""
    XEnvase(WLugar, 3) = ""
    XEnvase(WLugar, 4) = ""
    XEnvase(WLugar, 5) = ""
    XEnvase(WLugar, 6) = ""
    
    XEspecificaciones(WLugar) = ""
    
    XDatosMuestra(WLugar, 1) = ""
    XDatosMuestra(WLugar, 2) = ""
    XDatosMuestra(WLugar, 3) = ""
    
    WVector1.TextMatrix(WLugar, 1) = ""
    WVector1.TextMatrix(WLugar, 2) = ""
    WVector1.TextMatrix(WLugar, 3) = ""
    WVector1.TextMatrix(WLugar, 4) = ""
        
    Call Ingresa_Click
    
End Sub

Private Sub WVector1_Click()

    WVector1.Col = 1
    If Len(WVector1.Text) = 12 Then
        WLinea.Text = WVector1.Row
        WArticulo.Text = WVector1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    WVector1.Col = 2
    WDescripcion.Caption = WVector1.Text

    WVector1.Col = 3
    If Val(WVector1.Text) <> 0 Then
        WCantidad.Text = Pusing("###,###.##", WVector1.Text)
            Else
        WCantidad.Text = ""
    End If
    
    WVector1.Col = 4
    WPrecio.Caption = Pusing("###,###.##", WVector1.Text)
    
    WLugar = WVector1.Row
    
    Envase1.Text = XEnvase(WLugar, 1)
    Canti1.Text = XEnvase(WLugar, 2)
    Envase2.Text = XEnvase(WLugar, 3)
    Canti2.Text = XEnvase(WLugar, 4)
    Envase3.Text = XEnvase(WLugar, 5)
    Canti3.Text = XEnvase(WLugar, 6)
    
    Especificaciones.Text = XEspecificaciones(WLugar)
    
    NombreComercialMl.Text = XDatosMuestra(WLugar, 1)
    OrdenTrabajo.Text = XDatosMuestra(WLugar, 2)
    Referencia.Text = XDatosMuestra(WLugar, 3)
    
    WTermi = WArticulo.Text
    Call StkPed
    StkPedido.Caption = WStkPedido
    
    If Left$(WTermi, 2) <> "ML" Then
        Call Busca_Stock
    End If
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    If FecEntrega.Text = "00/00/0000" Or FecEntrega.Text = "  /  /    " Then
        m$ = "No esta informada la fecha de entrega"
        a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
        Exit Sub
    End If
    
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdFecEntrega = Right$(FecEntrega.Text, 4) + Mid$(FecEntrega.Text, 4, 2) + Left$(FecEntrega.Text, 2)
    If WFechaord > WOrdFecEntrega Then
        m$ = "La fecha de entrega no puede ser menor a la fecha del pedido"
        a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
        Exit Sub
    End If
    
    If Val(Wempresa) = 1 Then
        Call Verifica_Msds
        If ZZPasaMsds = "N" Then
            Exit Sub
        End If
    End If
    
    WFechaInicial = FecEntrega.Text
    WOrdFechaInicial = WOrdFecEntrega
    
    XPasa = "S"
    WLLave = 0
    ZNroSedronar = ""
    
    ZImpreVto = 0
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZNroSedronar = Trim(IIf(IsNull(rstCliente!NroSedronar), "", rstCliente!NroSedronar))
        ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
        rstCliente.Close
    End If
    
    ZVeriSedronar = "N"
    For a = 1 To 99
        
        WRow = a
        WVector1.Row = WRow
                
        WVector1.Col = 1
        Articulo = UCase(WVector1.Text)
                
        WVector1.Col = 3
        Cantidad = WVector1.Text
        
        If Val(Cantidad) <> 0 Then
        
            WCliente = UCase(Cliente.Text)
            WTerminado = UCase(Articulo)
            WClave = WCliente + WTerminado
            Xpago = 0

            WEnvase1 = XEnvase(a, 1)
            If Val(WEnvase1) = 0 Then
                m$ = "Se debe informar envases"
                CA% = MsgBox(m$, 0, "Emision de Facturas")
                Exit Sub
            End If
            
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Xpago = IIf(IsNull(rstPrecios!Pago), 0, rstPrecios!Pago)
                If Xpago = Val(Pago.Caption) Then
                    Xpago = 0
                End If
                rstPrecios.Close
            End If
            
            XCodigo = Val(Mid$(WTerminado, 4, 5))
            If Left$(WTerminado, 2) <> "PT" Then
                Select Case Left$(WTerminado, 2)
                    Case "DY", "DS"
                        XTipoPro = "CO"
                    Case "QC"
                        XTipoPro = "FA"
                    Case Else
                        XTipoPro = "PT"
                End Select
                    Else
                If XCodigo >= 0 And XCodigo <= 999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 11000 And XCodigo <= 12999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 25000 And XCodigo <= 25999 Then
                            XTipoPro = "FA"
                                Else
                            If XCodigo >= 2300 And XCodigo <= 2399 Then
                                XTipoPro = "BI"
                                    Else
                                XTipoPro = "PT"
                            End If
                        End If
                    End If
                End If
            End If
            
            If Left$(WTerminado, 2) = "YQ" Then
                XTipoPro = "PT"
            End If
            If Left$(WTerminado, 2) = "YH" Then
                XTipoPro = "PT"
            End If
            If Left$(WTerminado, 2) = "YP" Then
                XTipoPro = "PT"
            End If
            If Left$(WTerminado, 2) = "YF" Then
                XTipoPro = "FA"
            End If
            
            ZSedronar = 0
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZLinea = rstTerminado!Linea
                ZSedronar = IIf(IsNull(rstTerminado!Sedronar), "0", rstTerminado!Sedronar)
                rstTerminado.Close
            End If
            
            If ZSedronar = 1 Then
                ZVeriSedronar = "S"
            End If
            
            Select Case ZLinea
                Case 8
                    XTipoPro = "PG"
                Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                    XTipoPro = "FA"
                Case Else
            End Select
            
            Rem If UCase(Cliente.Text) = "S00130" Then
            Rem     XTipoPro = "CO"
            Rem End If
            
            If WLLave = 0 Then
                WLLave = 1
                WConpago = Xpago
                WTipopro = XTipoPro
            End If
            
            If WConpago <> Xpago Then
                XPasa = "1"
            End If
            
            If WTipopro <> XTipoPro Then
                XPasa = "2"
            End If
            
            If Left$(WTerminado, 4) = "PT-5" Then
            
                If Val(Wempresa) = 1 And Cliente.Text = "P00005" Then
                
                    ZPasa = "S"
                
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZZTerminado = "PT-0" + Mid$(WTerminado, 5, 8)
            
                    spTerminado = "ConsultaTerminado " + "'" + ZZTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                        WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                        
                        rstTerminado.Close
                        
                        If WEstadoI <> "S" Or WEstadoIII <> "S" Then
                            m$ = "El Producto Terminado no se encuentra autorizado para la Faturacion"
                            If WEstadoI <> "S" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                            End If
                            If WEstadoIII <> "S" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                            End If
                            CA% = MsgBox(m$, 0, "Emision de Facturas")
                            ZPasa = "N"
                        End If
                        
                            Else
                            
                        m$ = "Producto Terminado Inexistente"
                        CA% = MsgBox(m$, 0, "Emision de Facturas")
                        ZPasa = "N"
                        
                    End If
                    
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    If ZPasa = "N" Then
                        Exit Sub
                    End If
                    
                        Else
                        
                    If Left$(UCase(WTerminado), 2) = "PT" Then
                        
                        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                            WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                        
                            rstTerminado.Close
                        
                            If WEstadoI <> "S" Or WEstadoIII <> "S" Then
                                m$ = "El Producto Terminado no se encuentra autorizado para la Faturacion"
                                If WEstadoI <> "S" Then
                                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                                End If
                                If WEstadoIII <> "S" Then
                                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                                End If
                                CA% = MsgBox(m$, 0, "Emision de Facturas")
                                Exit Sub
                            End If
                        End If
                    
                    End If
                        
                End If
                
                    Else
                
                If Left$(UCase(WTerminado), 2) = "PT" Then
                
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                
                        WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                        WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                        
                        rstTerminado.Close
                        
                        If WEstadoI <> "S" Or WEstadoIII <> "S" Then
                            m$ = "El Producto Terminado no se encuentra autorizado para la Faturacion"
                            If WEstadoI <> "S" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                            End If
                            If WEstadoIII <> "S" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                            End If
                            CA% = MsgBox(m$, 0, "Emision de Facturas")
                            Exit Sub
                        End If
                        
                    End If
                        
                End If
                        
            End If
            
            If ZImpreVto = 1 Then
        
                ZVida = 0
                
                If Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Or Left$(WTerminado, 2) = "YP" Or Left$(WTerminado, 2) = "YH" Then
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                        rstTerminado.Close
                    End If
                        Else
                    WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZVida = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
                        rstArticulo.Close
                    End If
                End If
                
                If ZVida = 0 Then
                    m$ = "Atencion: El producto terminado " + WTerminado + " no posee vida util y el cliente lo exige"
                    a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
                    Exit Sub
                End If
                
            End If
        End If
            
    Next a
    
    If Val(Wempresa) = 1 Then
    
        If XPasa = "1" Then
            m$ = "Se cargaron articulos con distinta condicion de pago"
            a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
            Exit Sub
        End If

        If XPasa = "2" Then
            m$ = "Se cargaron articulos PT, Biosidas, Farma, Pigmentos o Colorantes en forma conjunta un mismo Pedido"
            a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
            Exit Sub
        End If
    
    End If
    
    If ZVeriSedronar = "S" Then
        If Trim(ZNroSedronar) = "" Then
            m$ = "Atencion: El cliente debe estar inscripto en el Sedronar para adquirir estos productos"
            a% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
        End If
    End If

    Xversion = 0
    
    If Val(Pedido.Text) = 0 Then
        ZSql = ""
        ZSql = ZSql + "Select Max(Pedido) as [PedidoMayor]"
        ZSql = ZSql + " FROM PedidoSol"
        spPedidoSol = ZSql
        Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoSol.RecordCount > 0 Then
            rstPedidoSol.MoveLast
            WPedidoMayor = IIf(IsNull(rstPedidoSol!PedidoMayor), "0", rstPedidoSol!PedidoMayor)
            Pedido.Text = Trim(Str$(WPedidoMayor + 1))
            rstPedidoSol.Close
        End If
            Else
        Pedido.Text = "1"
    End If
    
    spPedidoSol = "BorrarPedido " + "'" + Pedido.Text + "'"
    Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenDynaset, dbSQLPassThrough)
    
    Erase Auxiliar
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 99
        
        WRow = a
        WVector1.Row = WRow
        WLugar = a
                
        WVector1.Col = 1
        Articulo = UCase(WVector1.Text)
        
        WVector1.Col = 2
        NombreComercial = WVector1.Text
                
        WVector1.Col = 3
        Cantidad = WVector1.Text
                
        WVector1.Col = 4
        Precio = WVector1.Text
        
        If Val(Cantidad) <> 0 Then
        
            Renglon = Renglon + 1
            WRenglon = WRenglon + 1
                
            Auxiliar(WRenglon, 1) = Articulo
            Auxiliar(WRenglon, 2) = Cantidad
                
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                    
            Auxi1 = Str$(Pedido)
            Call Ceros(Auxi1, 6)
                
            WPedido = Pedido.Text
            WRenglon = Str$(Renglon)
            WCliente = Cliente.Text
            WFecha = Fecha.Text
            WFecEntrega = FecEntrega.Text
            WHora = Hora.Text
            WObservaciones = Observaciones.Text
            WOrdenCpa = OrdenCpa.Text
            WTerminado = Articulo
            WCantidad = Cantidad
            Rem aca se reemplaza la rutina de cambio envase
            WEnvase1 = XEnvase(WLugar, 1)
            WCanti1 = XEnvase(WLugar, 2)
            WEnvase2 = XEnvase(WLugar, 3)
            WCanti2 = XEnvase(WLugar, 4)
            WEnvase3 = XEnvase(WLugar, 5)
            WCanti3 = XEnvase(WLugar, 6)
            WEspecificaciones = XEspecificaciones(WLugar)
            WOrdenTrabajo = XDatosMuestra(WLugar, 2)
            WReferencia = XDatosMuestra(WLugar, 3)
            WEnvase4 = 0
            WCanti4 = ""
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WOrdFecEntrega = Right$(FecEntrega.Text, 4) + Mid$(FecEntrega.Text, 4, 2) + Left$(FecEntrega.Text, 2)
            WPrecio = Precio
            WWLinea = Linea
            WFacturado = ""
            WImporte = ""
            WClave = Auxi1 + Auxi
            WAutorizo = "N"
            WImpresion = "N"
            WTipoPed = Str$(Tipoped.ListIndex)
            WCantidad1 = ""
            WCantidad2 = ""
            WLote1 = "0"
            WLote2 = "0"
            Wlote3 = "0"
            WLote4 = "0"
            WLote5 = "0"
            WCantiLote1 = "0"
            WCantiLote2 = "0"
            WCantiLote3 = "0"
            WCantiLote4 = "0"
            WCantiLote5 = "0"
            WEnv1 = "0"
            WEnv2 = "0"
            WEnv3 = "0"
            WEnv4 = "0"
            WEnv5 = "0"
            WCantiEnv1 = "0"
            WCantiEnv2 = "0"
            WCantiEnv3 = "0"
            WCantiEnv4 = "0"
            WCantiEnv5 = "0"
            WVersion = Str$(Xversion + 1)
            If Left$(Articulo, 2) <> "PT" And Left$(Articulo, 2) <> "YQ" And Left$(Articulo, 2) <> "YF" And Left$(Articulo, 2) <> "YP" And Left$(Articulo, 2) <> "YH" Then
                WTipopro = "M"
                WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                    Else
                WTipopro = "T"
                WArti = "  -   -   "
            End If
            WVia = Str$(Via.ListIndex)
            
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO PedidoSol ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FecEntrega ,"
            ZSql = ZSql + "Hora ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Envase1 ,"
            ZSql = ZSql + "Canti1 ,"
            ZSql = ZSql + "Envase2 ,"
            ZSql = ZSql + "Canti2 ,"
            ZSql = ZSql + "Envase3 ,"
            ZSql = ZSql + "Canti3 ,"
            ZSql = ZSql + "Envase4 ,"
            ZSql = ZSql + "Canti4 ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Facturado ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Autorizo ,"
            ZSql = ZSql + "Impresion ,"
            ZSql = ZSql + "TipoPed ,"
            ZSql = ZSql + "Cantidad1 ,"
            ZSql = ZSql + "Cantidad2 ,"
            ZSql = ZSql + "Lote1 ,"
            ZSql = ZSql + "CantiLote1 ,"
            ZSql = ZSql + "Lote2 ,"
            ZSql = ZSql + "CantiLote2 ,"
            ZSql = ZSql + "Lote3 ,"
            ZSql = ZSql + "CantiLote3 ,"
            ZSql = ZSql + "Lote4 ,"
            ZSql = ZSql + "CantiLote4 ,"
            ZSql = ZSql + "Lote5 ,"
            ZSql = ZSql + "CantiLote5 ,"
            ZSql = ZSql + "Env1 ,"
            ZSql = ZSql + "CantiEnv1 ,"
            ZSql = ZSql + "Env2 ,"
            ZSql = ZSql + "CantiEnv2 ,"
            ZSql = ZSql + "Env3 ,"
            ZSql = ZSql + "CantiEnv3 ,"
            ZSql = ZSql + "Env4 ,"
            ZSql = ZSql + "CantiEnv4 ,"
            ZSql = ZSql + "Env5 ,"
            ZSql = ZSql + "CantiEnv5 ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "OrdFecEntrega ,"
            ZSql = ZSql + "OrdenCpa ,"
            ZSql = ZSql + "TipoPro ,"
            ZSql = ZSql + "Articulo )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WPedido + "',"
            ZSql = ZSql + "'" + WRenglon + "',"
            ZSql = ZSql + "'" + WCliente + "',"
            ZSql = ZSql + "'" + WFecha + "',"
            ZSql = ZSql + "'" + WFecEntrega + "',"
            ZSql = ZSql + "'" + WHora + "',"
            ZSql = ZSql + "'" + WObservaciones + "',"
            ZSql = ZSql + "'" + WTerminado + "',"
            ZSql = ZSql + "'" + WCantidad + "',"
            ZSql = ZSql + "'" + WEnvase1 + "',"
            ZSql = ZSql + "'" + WCanti1 + "',"
            ZSql = ZSql + "'" + WEnvase2 + "',"
            ZSql = ZSql + "'" + WCanti2 + "',"
            ZSql = ZSql + "'" + WEnvase3 + "',"
            ZSql = ZSql + "'" + WCanti3 + "',"
            ZSql = ZSql + "'" + WEnvase4 + "',"
            ZSql = ZSql + "'" + WCanti4 + "',"
            ZSql = ZSql + "'" + WFechaord + "',"
            ZSql = ZSql + "'" + WPrecio + "',"
            ZSql = ZSql + "'" + WLinea + "',"
            ZSql = ZSql + "'" + WFacturado + "',"
            ZSql = ZSql + "'" + WImporte + "',"
            ZSql = ZSql + "'" + WAutorizo + "',"
            ZSql = ZSql + "'" + WImpresion + "',"
            ZSql = ZSql + "'" + WTipoPed + "',"
            ZSql = ZSql + "'" + WCantidad1 + "',"
            ZSql = ZSql + "'" + WCantidad2 + "',"
            ZSql = ZSql + "'" + WLote1 + "',"
            ZSql = ZSql + "'" + WCantiLote1 + "',"
            ZSql = ZSql + "'" + WLote2 + "',"
            ZSql = ZSql + "'" + WCantiLote2 + "',"
            ZSql = ZSql + "'" + Wlote3 + "',"
            ZSql = ZSql + "'" + WCantiLote3 + "',"
            ZSql = ZSql + "'" + WLote4 + "',"
            ZSql = ZSql + "'" + WCantiLote4 + "',"
            ZSql = ZSql + "'" + WLote5 + "',"
            ZSql = ZSql + "'" + WCantiLote5 + "',"
            ZSql = ZSql + "'" + WEnv1 + "',"
            ZSql = ZSql + "'" + WCantiEnv1 + "',"
            ZSql = ZSql + "'" + WEnv2 + "',"
            ZSql = ZSql + "'" + WCantiEnv2 + "',"
            ZSql = ZSql + "'" + WEnv3 + "',"
            ZSql = ZSql + "'" + WCantiEnv3 + "',"
            ZSql = ZSql + "'" + WEnv4 + "',"
            ZSql = ZSql + "'" + WCantiEnv4 + "',"
            ZSql = ZSql + "'" + WEnv5 + "',"
            ZSql = ZSql + "'" + WCantiEnv5 + "',"
            ZSql = ZSql + "'" + WVersion + "',"
            ZSql = ZSql + "'" + WOrdFecEntrega + "',"
            ZSql = ZSql + "'" + WOrdenCpa + "',"
            ZSql = ZSql + "'" + WTipopro + "',"
            ZSql = ZSql + "'" + WArti + "')"
       
            spPedidoSol = ZSql
            Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
            
            ZZVendedor = ""
            ZZDesVendedor = ""
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZZDesCliente = rstCliente!Razon
                ZZVendedor = Str$(rstCliente!vendedor)
                rstCliente.Close
            End If
            
            spVendedor = "ConsultaVendedor " + "'" + ZZVendedor + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                ZZDesVendedor = rstVendedor!Nombre
                rstVendedor.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE PedidoSol SET "
            ZSql = ZSql + "DesCliente = " + "'" + ZZDesCliente + "',"
            ZSql = ZSql + "DesVendedor = " + "'" + ZZDesVendedor + "',"
            ZSql = ZSql + "Vendedor = " + "'" + ZZVendedor + "',"
            ZSql = ZSql + "NombreComercial = " + "'" + NombreComercial + "',"
            ZSql = ZSql + "OrdenTrabajo = " + "'" + WOrdenTrabajo + "',"
            ZSql = ZSql + "Referencia = " + "'" + WReferencia + "',"
            ZSql = ZSql + "Via = " + "'" + WVia + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                 
            spPedidoSol = ZSql
            Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
            
            
        End If
            
    Next a
    
    T$ = "Solicitud de Pedidos de Clientes"
    m$ = "Desea Imprimir la solicitud del pedido"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
        
    Call Limpia_Click
    Pedido.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    Stock.Caption = "0"
    StkPedido.Caption = "0"
    Produccion.Caption = "0"
    Disponible.Caption = "0"
    WStock1.Caption = "0"
    WStock2.Caption = "0"
    WStock3.Caption = "0"
    WStock4.Caption = "0"
    WStock5.Caption = "0"
    Wstock6.Caption = "0"
    Wstock7.Caption = "0"
    
    NombreComercialMl.Text = ""
    OrdenTrabajo.Text = ""
    Referencia.Text = ""
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    
    Especificaciones.Text = ""

    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    Version.Text = ""
    Stock.Caption = "0"
    StkPedido.Caption = "0"
    Produccion.Caption = "0"
    Disponible.Caption = "0"
    WStock1.Caption = "0"
    WStock2.Caption = "0"
    WStock3.Caption = "0"
    WStock4.Caption = "0"
    WStock5.Caption = "0"
    Wstock6.Caption = "0"
    Wstock7.Caption = "0"
    
    NombreComercialMl.Text = ""
    OrdenTrabajo.Text = ""
    Referencia.Text = ""
    
    
    WEnvase1.Caption = ""
    WEnvase2.Caption = ""
    WEnvase3.Caption = ""
    WEnvase4.Caption = ""
    WEnvase5.Caption = ""
    WEnvase6.Caption = ""
    
    WAbre1.Caption = ""
    WAbre2.Caption = ""
    WAbre3.Caption = ""
    WAbre4.Caption = ""
    WAbre5.Caption = ""
    WAbre6.Caption = ""
    
    WCapa1.Caption = ""
    WCapa2.Caption = ""
    Wcapa3.Caption = ""
    WCapa4.Caption = ""
    WCapa5.Caption = ""
    WCapa6.Caption = ""
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    
    Especificaciones.Text = ""
    
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Precio.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FecEntrega.Text = "  /  /    "
    Hora.Text = ""
    Pago.Caption = ""
    DesPago.Caption = ""
    Observaciones.Text = ""
    OrdenCpa.Text = ""
    WCalcula = "N"
    FecEntrega.Enabled = False
    WCalcula = "S"
    
    Tipoped.ListIndex = 0
    Via.ListIndex = 0
    
    Pantalla.Visible = False
    PantallaPro.Visible = False
    Ayuda.Visible = False
    
    Erase XEnvase
    Erase XEspecificaciones
    Erase XDatosMuestra
    
    Call Limpia_Vector
    
    Pedido.Text = ""
    
    Renglon = 0
    Pedido.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        WCliente = Cliente.Text
        WTerminado = WArticulo.Text
        WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
        WClave = Cliente.Text + WArticulo.Text
        WClaveMp = Cliente.Text + WArti
        
        If Left$(WArticulo.Text, 2) <> "PT" And Left$(WArticulo.Text, 2) <> "YQ" And Left$(WArticulo.Text, 2) <> "YF" And Left$(WArticulo.Text, 2) <> "YP" And Left$(WArticulo.Text, 2) <> "YH" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    WEntra = "S"
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        WPrecio.Caption = Pusing("###,###.##", "0")
                            Else
                        WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    End If
                    rstPreciosMp.Close
                    If Left$(WArti, 2) <> "ML" Then
                        Call Busca_Stock
                    End If
                    WCantidad.SetFocus
                        Else
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WDescripcion.Caption = rstArticulo!Descripcion
                            If Left$(UCase(WArti), 2) = "ML" Then
                                rstArticulo.Close
                                WEntra = "S"
                                WPrecio.Caption = Pusing("###,###.##", "0")
                                EntraNombreMl.Visible = True
                                Rem NombreComercialMl.Text = ""
                                Rem OrdenTrabajo.Text = ""
                                Rem Referencia.Text = ""
                                NombreComercialMl.SetFocus
                                    Else
                                rstArticulo.Close
                                WEntra = "S"
                                WPrecio.Caption = Pusing("###,###.##", "0")
                                Call Busca_Stock
                                WCantidad.SetFocus
                            End If
                                Else
                            WArticulo.SetFocus
                        End If
                            Else
                        WArticulo.SetFocus
                    End If
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WEntra = "S"
                    WDescripcion.Caption = rstPrecios!Descripcion
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        WPrecio.Caption = Pusing("###,###.##", "0")
                            Else
                        WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    End If
                    rstPrecios.Close
                    Call Busca_Stock
                    WCantidad.SetFocus
                        Else
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        T$ = "Pedidos de Clientes"
                        m$ = "Articulo sin nombre comercial. Desea ingresarlo :"
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
                            EntraNombre.Visible = True
                            NombreComercial.Text = ""
                            NombreComercial.SetFocus
                                Else
                            WArticulo.SetFocus
                        End If
                            Else
                        WArticulo.SetFocus
                    End If
                End If
            
        End Select
        
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        
        ZLinea = 0
        spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZLinea = rstTerminado!Linea
            rstTerminado.Close
        End If
        
                
        XCodigo = Val(Mid$(WArticulo.Text, 4, 5))
        Rem If (XCodigo >= 25000 And XCodigo <= 25999) Or ZLinea = 10 Or ZLinea = 20 Then
            If Val(Wempresa) = 1 Then
                IngreEspe.Visible = True
                If Especificaciones.Text = "" Then
                    XParam = "'" + Cliente.Text + "','" _
                                 + WArticulo.Text + "'"
                    spEspeCli = "ConsultaEspeCli " + XParam
                    Set rstEspeCli = db.OpenRecordset(spEspeCli, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspeCli.RecordCount > 0 Then
                        Especificaciones.Text = rstEspeCli!Especificaciones
                        rstEspeCli.Clone
                    End If
                End If
                Especificaciones.SetFocus
                    Else
                Envase1.SetFocus
            End If
        Rem         Else
        Rem     Envase1.SetFocus
        Rem End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NombreComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Cliente.Text = UCase(Cliente.Text)
        WArticulo.Text = UCase(WArticulo.Text)
    
        ZCliente = Cliente.Text
        ZArticulo = WArticulo.Text
        ZClave = Cliente.Text + WArticulo.Text
        
        ZFecha1 = ""
        ZFactura1 = ""
        ZPrecio1 = ""
        ZCantidad1 = ""
    
        ZFecha2 = ""
        ZFactura2 = ""
        ZPrecio2 = ""
        ZCantidad2 = ""
    
        ZFecha3 = ""
        ZFactura3 = ""
        ZPrecio3 = ""
        ZCantidad3 = ""
    
        ZFecha4 = ""
        ZFactura4 = ""
        ZPrecio4 = ""
        ZCantidad4 = ""
    
        ZFecha5 = ""
        ZFactura5 = ""
        ZPrecio5 = ""
        ZCantidad5 = ""
    
        ZFecha = Date$
    
        XParam = "'" + ZClave + "','" + Cliente.Text + "','" + WArticulo.Text + "','" + "0" + "','" _
                 + NombreComercial.Text + "','" _
                 + ZFecha1 + "','" + ZFactura1 + "','" + ZPrecio1 + "','" + ZCantidad1 + "','" _
                 + ZFecha2 + "','" + ZFactura2 + "','" + ZPrecio2 + "','" + ZCantidad2 + "','" _
                 + ZFecha3 + "','" + ZFactura3 + "','" + ZPrecio3 + "','" + ZCantidad3 + "','" _
                 + ZFecha4 + "','" + ZFactura4 + "','" + ZPrecio4 + "','" + ZCantidad4 + "','" _
                 + ZFecha5 + "','" + ZFactura5 + "','" + ZPrecio5 + "','" + ZCantidad5 + "','" _
                 + Date$ + "','" + ZFecha + "','" + "0" + "'"
        Set rstPrecios = db.OpenRecordset("AltaPrecios1 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    
        EntraNombre.Visible = False
        Call WArticulo_KeyPress(13)
        
    End If
End Sub

Private Sub NombreComercialMl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OrdenTrabajo.SetFocus
    End If
    If KeyAscii = 27 Then
        NombreComercialMl.Text = ""
    End If
End Sub

Private Sub OrdenTrabajo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(OrdenTrabajo.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ot"
            ZSql = ZSql + " Where Ot.Codigo = " + "'" + OrdenTrabajo.Text + "'"
            spOt = ZSql
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            If rstOt.RecordCount > 0 Then
                rstOt.Close
                Referencia.SetFocus
                    Else
                m$ = "Orden de trabajo invalida"
                a% = MsgBox(m$, 0, "Pedido de Muestras")
            End If
                Else
            Referencia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        OrdenTrabajo.Text = ""
    End If
End Sub

Private Sub Referencia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(OrdenTrabajo.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ot"
            ZSql = ZSql + " Where Ot.Codigo = " + "'" + OrdenTrabajo.Text + "'"
            spOt = ZSql
            Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
            If rstOt.RecordCount > 0 Then
                rstOt.Close
                    Else
                m$ = "Orden de trabajo invalida"
                a% = MsgBox(m$, 0, "Pedido de Muestras")
                Exit Sub
            End If
        End If
        WDescripcion.Caption = NombreComercialMl.Text
        WCantidad.SetFocus
        EntraNombreMl.Visible = False
    End If
    If KeyAscii = 27 Then
        Referencia.Text = ""
    End If
End Sub

Private Sub Especificaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
        IngreEspe.Visible = False
    End If
End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Envase1.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase1.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Or Val(Envase1.Text) = 99 Then
                Canti1.SetFocus
                    Else
                Envase1.SetFocus
            End If
                Else
            Rem Call Alta_Vector
            Rem Call Ingresa_Click
            Rem WArticulo.SetFocus
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
        If Val(Envase2.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase2.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Then
                Canti2.SetFocus
                    Else
                Envase2.SetFocus
            End If

                Else
            If Val(Envase1.Text) <> 0 Then
                Call Alta_Vector
                Call Ingresa_Click
                WArticulo.SetFocus
            End If
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
        If Val(Envase3.Text) <> 0 Then
            Ingre = "N"
            For DA = 1 To 6
                If Val(WVector(DA, 1)) = Val(Envase3.Text) Then
                    Ingre = "S"
                    Exit For
                End If
            Next DA
            If Ingre = "S" Then
                Canti3.SetFocus
                    Else
                Envase3.SetFocus
            End If
            
                Else
            If Val(Envase1.Text) <> 0 Then
                Call Alta_Vector
                Call Ingresa_Click
                WArticulo.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub PantallaPro_Click()

    Indice = PantallaPro.Row - 1
    Claveven$ = WIndice.List(Indice)
            
    If Mid$(Claveven$, 7, 2) <> "PT" And Mid$(Claveven$, 7, 2) <> "YQ" And Mid$(Claveven$, 7, 2) <> "YF" And Mid$(Claveven$, 7, 2) <> "YP" And Mid$(Claveven$, 7, 2) <> "YH" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
            
    Select Case WTipopro
        Case "M"
            Claveven$ = Left$(Claveven$, 9) + Right$(Claveven$, 7)
            spPreciosMp = "ConsultaPreciosMp " + "'" + Claveven$ + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                WArti = rstPreciosMp!Articulo
                WArticulo.Text = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                    WPrecio.Caption = Pusing("###,###.##", "0")
                        Else
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                End If
                
                rstPreciosMp.Close
                        
                If Left$(WArticulo.Text, 2) <> "ML" Then
                    Call Busca_Stock
                End If
                
                WCantidad.SetFocus
                    
            End If
            
        Case "T"
            spPrecios = "ConsultaPrecios " + "'" + Claveven$ + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                WArticulo.Text = rstPrecios!Terminado
                WDescripcion.Caption = rstPrecios!Descripcion
                If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                    WPrecio.Caption = "0"
                        Else
                    WPrecio.Caption = Str$(rstPrecios!Precio)
                End If
                    
                rstPrecios.Close
                        
                Call Busca_Stock
                    
                WCantidad.SetFocus
                    
            End If
            
        Case Else
    End Select
    
    If Val(Wempresa) <> 1 Then
        PantallaPro.Visible = False
    End If
    
End Sub

Private Sub pantalla_Click()
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
            
                Cliente.Text = Claveven$
                DesCliente.Caption = rstCliente!Razon
                Pago.Caption = rstCliente!Pago1
                Observaciones.Text = RTrim(rstCliente!Observaciones)
                Precio.Text = IIf(IsNull(rstCliente!Precio), "", rstCliente!Precio)
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                End If
                
                Erase WEspecif
                
                WEspecif(1) = ""
                WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
                WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
                WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
                WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
                WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
                WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
                WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
                WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
                WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
                For CicloEspecif = 1 To 10
                    WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
                Next CicloEspecif
                
                rstCliente.Close
                
                spPago = "ConsultaPago " + "'" + Pago.Caption + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstPago.RecordCount > 0 Then
                    DesPago.Caption = rstPago!Nombre
                    rstPago.Close
                End If
                Tipoped.SetFocus
            End If
            Pantalla.Visible = False
            Ayuda.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            If Mid$(Claveven$, 7, 2) <> "PT" And Mid$(Claveven$, 7, 2) <> "YQ" And Mid$(Claveven$, 7, 2) <> "YF" And Mid$(Claveven$, 7, 2) <> "YP" And Mid$(Claveven$, 7, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    Claveven$ = Left$(Claveven$, 9) + Right$(Claveven$, 7)
                    spPreciosMp = "ConsultaPreciosMp " + "'" + Claveven$ + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPreciosMp.RecordCount > 0 Then
                        WArti = rstPreciosMp!Articulo
                        WArticulo.Text = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                        If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                            WPrecio.Caption = "0"
                                Else
                            WPrecio.Caption = Str$(rstPreciosMp!Precio)
                        End If
                    
                        rstPreciosMp.Close
                        
                        Call Busca_Stock
                        
                        WCantidad.SetFocus
                    
                    End If
            
                Case "T"
                    spPrecios = "ConsultaPrecios " + "'" + Claveven$ + "'"
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPrecios.RecordCount > 0 Then
                        WArticulo.Text = rstPrecios!Terminado
                        WDescripcion.Caption = rstPrecios!Descripcion
                        If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                            WPrecio.Caption = "0"
                                Else
                            WPrecio.Caption = Str$(rstPrecios!Precio)
                        End If
                    
                        rstPrecios.Close
                        
                        Call Busca_Stock
                        
                        WCantidad.SetFocus
                    
                    End If
            
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub


Private Sub Form_Load()

    Call Limpia_Vector

     WEnvase(1) = 20
     WEnvase(2) = 21
     WEnvase(3) = 22
     WEnvase(4) = 23
     WEnvase(5) = 24
     WEnvase(6) = 25
     WEnvase(7) = 26
     WEnvase(8) = 30
     WEnvase(9) = 28
    
     For Cicla = 1 To 9
        spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
            rstEnvase.Close
                Else
            WImpre(Cicla) = ""
        End If
    Next Cicla

    Erase XEnvase
    Erase XEspecificaciones
    Erase XDatosMuestra

    Pantalla.Visible = False
    Pedido.Text = ""
    
    Tipoped.Clear
    
    Tipoped.AddItem "Normal"
    Tipoped.AddItem "a Fecha"
    Tipoped.AddItem "Fecha Limite"
    Tipoped.AddItem "Urgente"
    Tipoped.AddItem "Retira Cliente"
    Tipoped.AddItem "Muestra"
    Tipoped.AddItem "Muestra Retira"
    
    Tipoped.ListIndex = 0
    
    Via.Clear
    
    Via.AddItem ""
    Via.AddItem "Terrestre"
    Via.AddItem "Maritimo"
    Via.AddItem "Aereo"
    
    Via.ListIndex = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FecEntrega.Text = "  /  /    "
    Version.Text = ""
    
    WCalcula = "N"
    FecEntrega.Enabled = False
    WCalcula = "S"
    
End Sub

Private Sub Proceso_Click()

    Erase XEnvase
    Erase XEspecificaciones
    Erase XDatosMuestra
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Auxiliar
    WRenglon = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PedidoSol"
    ZSql = ZSql + " Where PedidoSol.Pedido = " + "'" + Pedido.Text + "'"
    ZSql = ZSql + " Order by PedidoSol.Clave"
    spPedidoSol = ZSql
    Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoSol.RecordCount > 0 Then
        WGraba = "N"
        With rstPedidoSol
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
            
                    WVector1.Row = Renglon
            
                    WVector1.Col = 1
                    WVector1.Text = rstPedidoSol!Terminado
                    Auxi1 = rstPedidoSol!Terminado
                    
                    WVector1.Col = 2
                    WVector1.Text = ""
            
                    WVector1.Col = 3
                    WVector1.Text = Pusing("###,###.##", rstPedidoSol!Cantidad - rstPedidoSol!Facturado)
                    
                    Rem aca se reemplaza la rutina de cambio envase

                    XEnvase(Renglon, 1) = rstPedidoSol!Envase1
                    XEnvase(Renglon, 2) = rstPedidoSol!Canti1
                    XEnvase(Renglon, 3) = rstPedidoSol!Envase2
                    XEnvase(Renglon, 4) = rstPedidoSol!Canti2
                    XEnvase(Renglon, 5) = rstPedidoSol!Envase3
                    XEnvase(Renglon, 6) = rstPedidoSol!Canti3
                    
                    XEspecificaciones(Renglon) = IIf(IsNull(rstPedidoSol!Especificaciones), "0", rstPedidoSol!Especificaciones)
                    
                    XDatosMuestra(Renglon, 1) = IIf(IsNull(rstPedidoSol!NombreComercial), "", rstPedidoSol!NombreComercial)
                    XDatosMuestra(Renglon, 2) = IIf(IsNull(rstPedidoSol!OrdenTrabajo), "", rstPedidoSol!OrdenTrabajo)
                    XDatosMuestra(Renglon, 3) = IIf(IsNull(rstPedidoSol!Referencia), "", rstPedidoSol!Referencia)
                    
                    WRenglon = WRenglon + 1
                
                    Auxiliar(WRenglon, 1) = rstPedidoSol!Cliente
                    Auxiliar(WRenglon, 2) = rstPedidoSol!Terminado
                    If Left$(rstPedidoSol!Terminado, 2) = "ML" Then
                        Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedidoSol!NombreComercial), "", rstPedidoSol!NombreComercial)
                    End If
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedidoSol.Close
    End If
    
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        WLugar = DA
        Cliente = Auxiliar(WLugar, 1)
        Terminado = Auxiliar(WLugar, 2)
        ZZNombreComercial = Auxiliar(WLugar, 3)
        
        If Left$(Terminado, 2) <> "PT" And Left$(Terminado, 2) <> "YQ" And Left$(Terminado, 2) <> "YF" And Left$(Terminado, 2) <> "YP" And Left$(Terminado, 2) <> "YH" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + WArti + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                
                    WVector1.Row = WLugar
                
                    WVector1.Col = 4
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        WVector1.Text = Pusing("###,###.##", "0")
                            Else
                        WVector1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    End If
                    
                    rstPreciosMp.Close
                    
                    WArticulo.SetFocus
                    
                        Else
        
                    WVector1.Row = WLugar
                        
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", "0")
                    
                    WArticulo.SetFocus
                    
                End If
                    
                If Trim(ZZNombreComercial) = "" Then
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                        Else
                    WVector1.Col = 2
                    WVector1.Text = ZZNombreComercial
                End If
            
            Case Else
                ZZDescripcion = ""
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
        
                    WVector1.Row = WLugar
                
                    WVector1.Col = 2
                    WVector1.Text = rstPrecios!Descripcion
                    ZZDescripcion = rstPrecios!Descripcion
            
                    WVector1.Col = 4
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        WVector1.Text = Pusing("###,###.##", "0")
                            Else
                        WVector1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    End If
                    
                    rstPrecios.Close
                    
                    If Trim(ZZDescripcion) = "" Then
                        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WVector1.Col = 2
                            WVector1.Text = rstTerminado!Descripcion
                            rstTerminado.Close
                        End If
                    End If
                    
                    WArticulo.SetFocus
                    
                End If
                
        End Select
        
    Next DA

    WArticulo.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then
    
        XCodigo = Val(Mid$(WArticulo.Text, 4, 5))
        If XCodigo >= 11000 And XCodigo <= 11999 Then
            If Tipoped.ListIndex = 0 Then
                XFec1 = FecEntrega.Text
                SumaDia = 2
                Do
                    SumaDia = SumaDia + 1
                    Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                    FecEntrega.Text = XFec2
                    XFec1 = XFec2
                    strDia = Format$(XFec1, "dddd")
                    BDia = Format(XFec1, "w")
                    SumaDia = 1
                    If Val(BDia) <> 7 And Val(BDia) <> 1 Then
                        FecEntrega.Text = XFec1
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        Renglon = 1
        For ZZCiclo = 1 To 80
            If Trim(WVector1.TextMatrix(ZZCiclo, 1)) = "" Then
                Renglon = ZZCiclo
                Exit For
            End If
        Next ZZCiclo
                
        ZZRenglon = Renglon
        WLugar = ZZRenglon
        
        WVector1.Row = ZZRenglon
        
        WVector1.Col = 1
        WVector1.Text = WArticulo.Text
        
        WVector1.Col = 2
        WVector1.Text = WDescripcion.Caption
            
        WVector1.Col = 3
        WVector1.Text = Pusing("###,###.##", WCantidad.Text)
        
        WVector1.Col = 4
        WVector1.Text = Pusing("###,###.##", WPrecio.Caption)
        
        XEnvase(WLugar, 1) = Envase1.Text
        XEnvase(WLugar, 2) = Canti1.Text
        XEnvase(WLugar, 3) = Envase2.Text
        XEnvase(WLugar, 4) = Canti2.Text
        XEnvase(WLugar, 5) = Envase3.Text
        XEnvase(WLugar, 6) = Canti3.Text
        
        XEspecificaciones(WLugar) = Especificaciones.Text
        
        XDatosMuestra(WLugar, 1) = NombreComercialMl.Text
        XDatosMuestra(WLugar, 2) = OrdenTrabajo.Text
        XDatosMuestra(WLugar, 3) = Referencia.Text
        
            Else
            
        WVector1.Row = Val(WLinea.Text)
        WLugar = Val(WLinea.Text)
        
        WVector1.Col = 1
        WVector1.Text = WArticulo.Text
        
        WVector1.Col = 2
        WVector1.Text = WDescripcion.Caption
        
        WVector1.Col = 3
        WVector1.Text = Pusing("###,###.##", WCantidad.Text)
        
        WVector1.Col = 4
        WVector1.Text = Pusing("###,###.##", WPrecio.Caption)
        
        XEnvase(WLugar, 1) = Envase1.Text
        XEnvase(WLugar, 2) = Canti1.Text
        XEnvase(WLugar, 3) = Envase2.Text
        XEnvase(WLugar, 4) = Canti2.Text
        XEnvase(WLugar, 5) = Envase3.Text
        XEnvase(WLugar, 6) = Canti3.Text
        
        XEspecificaciones(WLugar) = Especificaciones.Text
        
        XDatosMuestra(WLugar, 1) = NombreComercialMl.Text
        XDatosMuestra(WLugar, 2) = OrdenTrabajo.Text
        XDatosMuestra(WLugar, 3) = Referencia.Text
            
    End If

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoSol"
        ZSql = ZSql + " Where PedidoSol.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " Order by PedidoSol.Clave"
        spPedidoSol = ZSql
        Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoSol.RecordCount > 0 Then
            Fecha.Text = rstPedidoSol!Fecha
            Cliente.Text = rstPedidoSol!Cliente
            FecEntrega.Text = rstPedidoSol!FecEntrega
            Hora.Text = rstPedidoSol!Hora
            Observaciones.Text = rstPedidoSol!Observaciones
            OrdenCpa.Text = IIf(IsNull(rstPedidoSol!OrdenCpa), "", rstPedidoSol!OrdenCpa)
            Tipoped.ListIndex = IIf(IsNull(rstPedidoSol!Tipoped), "0", rstPedidoSol!Tipoped)
            Version.Text = IIf(IsNull(rstPedidoSol!Version), "0", rstPedidoSol!Version)
            ZLugarDirEntrega = IIf(IsNull(rstPedidoSol!DirEntrega), "1", rstPedidoSol!DirEntrega)
            Via.ListIndex = IIf(IsNull(rstPedidoSol!Via), "0", rstPedidoSol!Via)
            rstPedidoSol.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                Pago.Caption = rstCliente!Pago1
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                Rem Observaciones.Text = rstCliente!Observaciones
                Precio.Text = IIf(IsNull(rstCliente!Precio), "", rstCliente!Precio)
                
                Erase WEspecif
                WEspecif(1) = ""
                WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
                WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
                WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
                WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
                WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
                WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
                WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
                WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
                WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
                For CicloEspecif = 1 To 10
                    WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
                Next CicloEspecif
                
                rstCliente.Close
                
                spPago = "ConsultaPago " + "'" + Pago.Caption + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstPago.RecordCount > 0 Then
                    DesPago.Caption = rstPago!Nombre
                    rstPago.Close
                End If
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Call Limpia_Click
            Pedido.Text = WPedido
            Cliente.SetFocus
        End If
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                Pago.Caption = rstCliente!Pago1
                Observaciones.Text = RTrim(rstCliente!Observaciones)
                Precio.Text = IIf(IsNull(rstCliente!Precio), "", rstCliente!Precio)
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                End If
                
                
                Erase WEspecif
                WEspecif(1) = ""
                WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
                WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
                WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
                WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
                WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
                WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
                WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
                WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
                WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
                For CicloEspecif = 1 To 10
                    WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
                Next CicloEspecif
                
                
                
                
                
                rstCliente.Close
                
                spPago = "ConsultaPago " + "'" + Pago.Caption + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstPago.RecordCount > 0 Then
                    DesPago.Caption = rstPago!Nombre
                    rstPago.Close
                End If
                
                Tipoped.SetFocus
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
        End If
    End If
End Sub

Private Sub TipoPed_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Tipoped.ListIndex = 0 Then
            If FecEntrega.Text = "  /  /    " And Tipoped.ListIndex = 0 And Fecha.Text <> "  /  /    " Then
                Call Calcula_FecEntrega
                Call Calcula_Feriado
            End If
            FecEntrega.Enabled = False
            Hora.SetFocus
                Else
            FecEntrega.Enabled = True
            FecEntrega.SetFocus
        End If
    End If
End Sub

Private Sub Tipoped_Click()
    If WCalcula = "S" Then
        If Tipoped.ListIndex = 0 Then
            If FecEntrega.Text = "  /  /    " And Tipoped.ListIndex = 0 And Fecha.Text <> "  /  /    " Then
                Call Calcula_FecEntrega
                Call Calcula_Feriado
            End If
            FecEntrega.Enabled = False
            Rem Hora.SetFocus
                Else
            If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                If FecEntrega.Text = "  /  /    " And Fecha.Text <> "  /  /    " Then
                    Call Calcula_FecEntrega_Muestra
                    Call Calcula_Feriado
                End If
                FecEntrega.Enabled = True
                FecEntrega.SetFocus
                    Else
                FecEntrega.Enabled = True
                FecEntrega.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Calcula_FecEntrega()

    Rem 1 - DOMINGO
    Rem 2 - LUNES
    Rem 3 - MARTES
    Rem 4 - MIERCOLES
    Rem 5 - JUEVES
    Rem 6 - VIERNES
    Rem 7 - SABADO
    
    ZProvincia = 0
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZProvincia = rstCliente!Provincia
        rstCliente.Close
    End If
    
    If ZProvincia <> 24 Then
    
        XFec1 = Fecha.Text
        strDia = Format$(XFec1, "dddd")
        BDia = Format(XFec1, "w")
        Select Case BDia
            Case 2, 3, 4
                SumaDia = 2
            Case 5, 6
                SumaDia = 4
            Case 7
                SumaDia = 3
            Case 1
                SumaDia = 2
            Case Else
        End Select
        SumaDia = SumaDia + 1
        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
        FecEntrega.Text = XFec2
        
            Else
            
        XFec1 = Fecha.Text
        SumaDia = 15
        Do
            SumaDia = SumaDia + 1
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            FecEntrega.Text = XFec2
            XFec1 = XFec2
            strDia = Format$(XFec1, "dddd")
            BDia = Format(XFec1, "w")
            SumaDia = 1
            If Val(BDia) <> 7 And Val(BDia) <> 1 Then
                Exit Do
            End If
        Loop
    End If

End Sub

Private Sub Calcula_FecEntregaColo()

    Rem 1 - DOMINGO
    Rem 2 - LUNES
    Rem 3 - MARTES
    Rem 4 - MIERCOLES
    Rem 5 - JUEVES
    Rem 6 - VIERNES
    Rem 7 - SABADO
    
    ZProvincia = 0
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZProvincia = rstCliente!Provincia
        rstCliente.Close
    End If
    
    If ZProvincia <> 24 Then
    
        XFec1 = Fecha.Text
        strDia = Format$(XFec1, "dddd")
        BDia = Format(XFec1, "w")
        Select Case BDia
            Case 1, 2, 3
                SumaDia = 3
            Case 4, 5, 6
                SumaDia = 5
            Case 7
                SumaDia = 4
            Case Else
        End Select
        SumaDia = SumaDia + 1
        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
        FecEntrega.Text = XFec2
        
            Else
            
        XFec1 = Fecha.Text
        SumaDia = 15
        Do
            SumaDia = SumaDia + 1
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            FecEntrega.Text = XFec2
            XFec1 = XFec2
            strDia = Format$(XFec1, "dddd")
            BDia = Format(XFec1, "w")
            SumaDia = 1
            If Val(BDia) <> 7 And Val(BDia) <> 1 Then
                Exit Do
            End If
        Loop
    End If

End Sub

Private Sub Calcula_FecEntrega_Muestra()

    Rem 1 - DOMINGO
    Rem 2 - LUNES
    Rem 3 - MARTES
    Rem 4 - MIERCOLES
    Rem 5 - JUEVES
    Rem 6 - VIERNES
    Rem 7 - SABADO
    
    ZProvincia = 0
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZProvincia = rstCliente!Provincia
        rstCliente.Close
    End If
    
    If ZProvincia <> 24 Then
    
        XFec1 = Fecha.Text
        strDia = Format$(XFec1, "dddd")
        BDia = Format(XFec1, "w")
        Rem Select Case BDia
        Rem     Case 2, 3, 4
        Rem         SumaDia = 2
        Rem     Case 5, 6
        Rem         SumaDia = 4
        Rem     Case 7
        Rem         SumaDia = 3
        Rem     Case 1
        Rem         SumaDia = 2
        Rem     Case Else
        Rem End Select
        Rem SumaDia = SumaDia + 1
        SumaDia = 8
        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
        FecEntrega.Text = XFec2
        
            Else
            
        XFec1 = Fecha.Text
        SumaDia = 15
        Do
            SumaDia = SumaDia + 1
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            FecEntrega.Text = XFec2
            XFec1 = XFec2
            strDia = Format$(XFec1, "dddd")
            BDia = Format(XFec1, "w")
            SumaDia = 1
            If Val(BDia) <> 7 And Val(BDia) <> 1 Then
                Exit Do
            End If
        Loop
    End If

End Sub


Private Sub Calcula_Feriado()

    Erase DiaFeriado
    TotalFeriado = 0
    
    spFeriado = "ListaFeriado"
    Set rstFeriado = db.OpenRecordset(spFeriado, dbOpenSnapshot, dbSQLPassThrough)
    If rstFeriado.RecordCount > 0 Then
        With rstFeriado
            .MoveFirst
            Do
                If .EOF = False Then
                    TotalFeriado = TotalFeriado + 1
                    DiaFeriado(TotalFeriado) = rstFeriado!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFeriado.Close
    End If
    
    Do
    
        Feriado = "N"
        For Ciclo = 1 To TotalFeriado
            If DiaFeriado(Ciclo) = FecEntrega.Text Then
                Feriado = "S"
                Exit For
            End If
        Next Ciclo
                
        Rem 1 - DOMINGO
        Rem 2 - LUNES
        Rem 3 - MARTES
        Rem 4 - MIERCOLES
        Rem 5 - JUEVES
        Rem 6 - VIERNES
        Rem 7 - SABADO
        XFec1 = FecEntrega.Text
        strDia = Format$(XFec1, "dddd")
        BDia = Format(XFec1, "w")
        If BDia = 1 Or BDia = 7 Then
            Feriado = "S"
        End If
        
        If Feriado = "S" Then
            SumaDia = 2
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            FecEntrega.Text = XFec2
                Else
            Exit Do
        End If
        
    Loop

End Sub


Private Sub FecEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If FecEntrega.Text = "  /  /    " Then
            FecEntrega.Text = "00/00/0000"
        End If
        Call Valida_fecha(FecEntrega.Text, Auxi)
        If Auxi = "S" Or FecEntrega.Text = "00/00/0000" Then
            Hora.SetFocus
                Else
            FecEntrega.SetFocus
        End If
    End If
End Sub

Private Sub Hora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OrdenCpa.SetFocus
    End If
End Sub

Private Sub OrdenCpa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
End Sub

Sub Carga_Envases()

 ZZDa = 0

 For Cicla = 1 To 6
    spEnvase = "ConsultaEnvases " + "'" + WVector(Cicla, 1) + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        WVector(Cicla, 2) = rstEnvase!Kilos
        WVector(Cicla, 3) = rstEnvase!Abreviatura
        rstEnvase.Close
                Else
        If ZZDa = 0 Then
            WVector(Cicla, 1) = "99"
            WVector(Cicla, 2) = "0"
            WVector(Cicla, 3) = "Indis."
            ZZDa = 1
                Else
            WVector(Cicla, 2) = ""
            WVector(Cicla, 3) = ""
        End If
    End If
Next Cicla

WEnvase1.Caption = WVector(1, 1)
WEnvase2.Caption = WVector(2, 1)
WEnvase3.Caption = WVector(3, 1)
WEnvase4.Caption = WVector(4, 1)
WEnvase5.Caption = WVector(5, 1)
WEnvase6.Caption = WVector(6, 1)

Rem WCapa1.Caption = WVector(1, 2)
Rem WCapa2.Caption = WVector(2, 2)
Rem Wcapa3.Caption = WVector(3, 2)
Rem WCapa4.Caption = WVector(4, 2)
Rem WCapa5.Caption = WVector(5, 2)
Rem WCapa6.Caption = WVector(6, 2)

WCapa1.Caption = WVector(1, 2)
WCapa2.Caption = WVector(2, 2)
Wcapa3.Caption = WVector(3, 2)
WCapa4.Caption = WVector(4, 2)
WCapa5.Caption = WVector(5, 2)
WCapa6.Caption = WVector(6, 2)

WAbre1.Caption = WVector(1, 3)
WAbre2.Caption = WVector(2, 3)
WAbre3.Caption = WVector(3, 3)
WAbre4.Caption = WVector(4, 3)
WAbre5.Caption = WVector(5, 3)
WAbre6.Caption = WVector(6, 3)


End Sub

Private Sub Impresion()

    On Error GoTo WError
    
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    WObservaciones = Left$(Observaciones.Text + Space$(100), 100)
    Select Case Tipoped.ListIndex
        Case 0
            WTipoPedido = " (Normal)"
        Case 1
            WTipoPedido = " (A fecha)"
        Case 2
            WTipoPedido = " (Fecha Limite)"
        Case 3
            WTipoPedido = " (Urgente)"
        Case 4
            WTipoPedido = " (Retira Cliente)"
        Case 5
            WTipoPedido = " (Muestra)"
        Case Else
    End Select
    
    WVia = ""
    Select Case Via.ListIndex
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For a = 1 To 99
        
        WCounter = WCounter + 1
        WVector1.Row = a
                
        WVector1.Col = 1
                
        If WVector1.Text <> "" Then
                
            WArticulo = WVector1.Text
                
            WVector1.Col = 2
            WDescripcion = WVector1.Text
                
            WVector1.Col = 3
            WCantidad = Val(WVector1.Text)
                
            WVector1.Col = 4
            WPrecio = Val(WVector1.Text)
            
            Rem aca se reemplaza la rutina de cambio envase

            WLugar = WVector1.Row
            
            WEspecificaciones = XEspecificaciones(WLugar)
                
            If WCantidad <> 0 Then
            
                Erase ImpreEnvase
                LugarEnvase = 0
            
                For Cicla = 1 To 6 Step 2
                    If Val(XEnvase(WLugar, Cicla)) <> 0 Then
                    Rem If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                        LugarEnvase = LugarEnvase + 1
                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WLugar, Cicla) + "'"
                        Rem spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                        
                        If rstEnvase.RecordCount > 0 Then
                            WAbre = rstEnvase!Abreviatura
                            rstEnvase.Close
                                Else
                            WAbre = ""
                        End If
                        ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WLugar, Cicla + 1))) + " " + Left$(WAbre, 8)
                        Rem ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8)
                    End If
                Next Cicla
                
                WRenglon = WRenglon + 1
                
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = Pedido.Text
                ZRenglon = Str$(WRenglon)
                ZEmpresa = WNombreEmpresa
                ZVersion = WVersion
                ZCliente = Cliente.Text
                ZNombre = DesCliente.Caption
                ZFecha = Fecha.Text
                ZFechaent = FecEntrega.Text
                ZTipoPedido = WTipoPedido
                ZCondicion = DesPago.Caption
                ZEntrega = WDirentrega
                ZObservaciones1 = Left$(WObservaciones, 50)
                ZObservaciones2 = Right$(WObservaciones, 50)
                ZOrden = OrdenCpa.Text
                ZArticulo = WArticulo
                ZDescripcion = WDescripcion
                ZPrecio = Str$(WPrecio)
                ZCantidad = Str$(WCantidad)
                ZEnvase = ImpreEnvase(1)
                
                spImprePed = "INSERT INTO ImprePed (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                            
                Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                
                If WEspecificaciones <> "" And WEspecificaciones <> "0" Then
                
                    WRenglon = WRenglon + 1
                
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = WVersion
                    ZCliente = Cliente.Text
                    ZNombre = DesCliente.Caption
                    ZFecha = Fecha.Text
                    ZFechaent = FecEntrega.Text
                    ZTipoPedido = WTipoPedido
                    ZCondicion = DesPago.Caption
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = OrdenCpa.Text
                    ZArticulo = "Especif.:"
                    ZDescripcion = WEspecificaciones
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                    
                    spImprePed = "INSERT INTO ImprePed (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                            
                    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                For Ciclo = 2 To LugarEnvase
                
                    WRenglon = WRenglon + 1
                
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = WVersion
                    ZCliente = Cliente.Text
                    ZNombre = DesCliente.Caption
                    ZFecha = Fecha.Text
                    ZFechaent = FecEntrega.Text
                    ZTipoPedido = WTipoPedido
                    ZCondicion = DesPago.Caption
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = OrdenCpa.Text
                    ZArticulo = ""
                    ZDescripcion = ""
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ImpreEnvase(Ciclo)
                    
                    spImprePed = "INSERT INTO ImprePed (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                            
                    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                    
                Next Ciclo
                    
            End If
                
        End If
            
    Next a
    
    SumaEspe = 0
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
        SumaEspe = SumaEspe + 1
    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = Pedido.Text
        ZRenglon = Str$(WRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = WVersion
        ZCliente = Cliente.Text
        ZNombre = DesCliente.Caption
        ZFecha = Fecha.Text
        ZFechaent = FecEntrega.Text
        ZTipoPedido = WTipoPedido
        ZCondicion = DesPago.Caption
        ZEntrega = WDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = OrdenCpa.Text
        ZArticulo = ""
        Rem ZDescripcion = WEspecif(SumaEspe)
        ZDescripcion = ""
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
                        
        spImprePed = "INSERT INTO ImprePed (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePed SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "'"
    spImprePed = ZSql
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
        Listado.ReportFileName = "ImprepedsqlMuestra.rpt"
            Else
        Listado.ReportFileName = "Imprepedsql.rpt"
    End If
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
    
    
    ZZRequiereCertificado = ""
    ZZRequiereMsds = ""
    ZZRequiereMsdsCada = ""
    ZZRequiereHoja = ""
    ZZPermiteParcial = ""
    ZZPartidasVarias = ""

    ZZEmailCertificado = ""
    ZZEmailMsds = ""
    ZZEmailHoja = ""
    ZZDiasI = ""
    ZZDiasII = ""
    ZZDiasIII = ""
    ZZEnvasesI = ""
    ZZEnvasesII = ""
    ZZEnvasesIII = ""
    ZZEtiquetaI = ""
    ZZEtiquetaII = ""
    ZZEspecif1 = ""
    ZZEspecif2 = ""
    ZZEspecif3 = ""
    ZZEspecif4 = ""
    ZZEspecif5 = ""
    ZZCantidadPartidas = ""
    
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
        ZZPermiteParcial = IIf(IsNull(rstClienteEspecif!PermiteParcial), "0", rstClienteEspecif!PermiteParcial)
        ZZPartidasVarias = IIf(IsNull(rstClienteEspecif!PartidaVarias), "0", rstClienteEspecif!PartidaVarias)

        ZZEmailCertificado = IIf(IsNull(rstClienteEspecif!EmailCertificado), "", rstClienteEspecif!EmailCertificado)
        ZZEmailMsds = IIf(IsNull(rstClienteEspecif!EmailMsds), "", rstClienteEspecif!EmailMsds)
        ZZEmailHoja = IIf(IsNull(rstClienteEspecif!EmailHoja), "", rstClienteEspecif!EmailHoja)
        ZZDiasI = IIf(IsNull(rstClienteEspecif!DiasI), "", rstClienteEspecif!DiasI)
        ZZDiasII = IIf(IsNull(rstClienteEspecif!DiasII), "", rstClienteEspecif!DiasII)
        ZZDiasIII = IIf(IsNull(rstClienteEspecif!DiasIII), "", rstClienteEspecif!DiasIII)
        ZZEnvasesI = IIf(IsNull(rstClienteEspecif!EnvasesI), "", rstClienteEspecif!EnvasesI)
        ZZEnvasesII = IIf(IsNull(rstClienteEspecif!EnvasesII), "", rstClienteEspecif!EnvasesII)
        ZZEnvasesIII = IIf(IsNull(rstClienteEspecif!EnvasesIII), "", rstClienteEspecif!EnvasesIII)
        ZZEtiquetaI = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
        ZZEtiquetaII = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
        ZZEspecif1 = IIf(IsNull(rstClienteEspecif!Especif1), "", rstClienteEspecif!Especif1)
        ZZEspecif2 = IIf(IsNull(rstClienteEspecif!Especif2), "", rstClienteEspecif!Especif2)
        ZZEspecif3 = IIf(IsNull(rstClienteEspecif!Especif3), "", rstClienteEspecif!Especif3)
        ZZEspecif4 = IIf(IsNull(rstClienteEspecif!Especif4), "", rstClienteEspecif!Especif4)
        ZZEspecif5 = IIf(IsNull(rstClienteEspecif!Especif5), "", rstClienteEspecif!Especif5)
        ZZCantidadPartidas = IIf(IsNull(rstClienteEspecif!CantidadPartidas), "", rstClienteEspecif!CantidadPartidas)
        
        rstClienteEspecif.Close
        
    End If
    
    ZZImprime = "N"
    
    If Val(ZZRequiereCertificado) <> 0 Or Val(ZZRequiereMsds) <> 0 Or Val(ZZRequiereMsdsCada) <> 0 Or Val(ZZRequiereHoja) <> 0 Or Val(ZZPermiteParcial) <> 0 Or Val(ZZPartidasVarias) <> 0 Then
        ZZImprime = "S"
    End If
    If Trim(ZZDiasI) <> "" Or Trim(ZZDiasII) <> "" Or Trim(ZZDiasIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEnvasesI) <> "" Or Trim(ZZEnvasesII) <> "" Or Trim(ZZEnvasesIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEtiquetaI) <> "" Or Trim(ZZEtiquetaII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEspecif1) <> "" Or Trim(ZZEspecif2) <> "" Or Trim(ZZEspecif3) <> "" Or Trim(ZZEspecif4) <> "" Or Trim(ZZEspecif5) <> "" Then
        ZZImprime = "S"
    End If
    If Val(ZZCantidadPartidas) <> 0 Then
        ZZImprime = "S"
    End If
    
    If ZZImprime = "S" Then
     Rem by nan no se necesita para ventas
    
    Rem    DbConnect = db.Connect
    Rem    DSQ = getDatabase(DbConnect)
    Rem    Listado.SQLQuery = "SELECT ImprePed.Clave, ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.TipoPedido, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via, " _
                + "ClienteEspecif.RequiereCertificado, ClienteEspecif.RequiereMsds, ClienteEspecif.RequiereMsdsCada, ClienteEspecif.RequiereHoja, ClienteEspecif.PermiteParcial, ClienteEspecif.DiasI, ClienteEspecif.DiasII, ClienteEspecif.DiasIII, ClienteEspecif.Especif1, ClienteEspecif.Especif2, ClienteEspecif.Especif3, ClienteEspecif.Especif4, ClienteEspecif.Especif5, ClienteEspecif.PartidaVarias, ClienteEspecif.CantidadPartidas, ClienteEspecif.EnvasesI, ClienteEspecif.EnvasesII, ClienteEspecif.EnvasesIII, ClienteEspecif.EtiquetaI, ClienteEspecif.EtiquetaII " _
                + "From " _
                + DSQ + ".dbo.ImprePed ImprePed, " _
                + DSQ + ".dbo.ClienteEspecif ClienteEspecif " _
                + "Where " _
                + "ImprePed.Cliente = ClienteEspecif.Cliente AND " _
                + "ImprePed.Pedido >= 0 AND " _
    rem            + "ImprePed.Pedido <= 999999"
                            
    Rem    Listado.Connect = Connect()
   
  
   Rem     Listado.ReportFileName = "ImprepedsqlEspecif.rpt"
   Rem     Listado.Destination = 1
        Rem Listado.Destination = 0
   Rem     Listado.CopiesToPrinter = 1
  Rem      Listado.Action = 1
        
    End If
    
        
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        WEspacios = Len(Ayuda.Text)
        WIndice.Clear
        
        Select Case XIndice
            Case 0
                Pantalla.Clear
                
                spCliente = "ListaClienteConsulta"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    With rstCliente
                        .MoveFirst
                        Do
                            If .EOF = False Then
                
                                DA = Len(rstCliente!Razon) - WEspacios
                    
                                For aa = 1 To DA
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                                        Auxi = rstCliente!Cliente
                                        IngresaItem = Auxi + "    " + rstCliente!Razon
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstCliente!Cliente
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
                    rstCliente.Close
                End If
                
            Case 1
                Call Limpia_PantallaPro
                LugarPantalla = 0
                
                Sql1 = "Select Cliente, Terminado, Descripcion, Precio, Fecha, Estado"
                Sql2 = " FROM Precios"
                Sql3 = " Where Precios.Cliente = " + "'" + Cliente.Text + "'"
                Sql4 = " Order by Terminado"
                spPrecios = Sql1 + Sql2 + Sql3 + Sql4
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
        
                    With rstPrecios
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If Cliente.Text = rstPrecios!Cliente Then
                                
                                ZEstado = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
                                If Val(ZEstado) = Estado.ListIndex Then
                                
                                
                                    DA = Len(rstPrecios!Descripcion) - WEspacios
                                    WIngresa = "N"
                                    For Aaa = 1 To DA
                                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstPrecios!Descripcion, Aaa, WEspacios) Then
                                            WIngresa = "S"
                                            Exit For
                                        End If
                                    Next Aaa
                                    
                                    If WIngresa = "S" Then
                                        ZTerminado = rstPrecios!Terminado
                                        ZDescripcion = rstPrecios!Descripcion
                                        If rstPrecios!Precio <> Null Then
                                            ZPrecio = Str$(rstPrecios!Precio)
                                                Else
                                            ZPrecio = IIf(IsNull(rstPrecios!Precio), "0", Str$(rstPrecios!Precio))
                                        End If
                                        ZFecha = IIf(IsNull(rstPrecios!Fecha), "  /  /    ", rstPrecios!Fecha)
                                        ZPrecio = Pusing("###,###.##", ZPrecio)
                                        LugarPantalla = LugarPantalla + 1
                                        PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
                                        PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
                                        PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
                                        PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
                                        IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                                        WIndice.AddItem IngresaItem
                                    End If
                                    
                                End If
                                
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstPrecios.Close
                End If
        
                Erase IngreVector
                EntraVector = 0
        
                Sql1 = "Select Cliente, Articulo, Precio, Fecha, Estado"
                Sql2 = " FROM PreciosMp"
                Sql3 = " Where PreciosMp.Cliente = " + "'" + Cliente.Text + "'"
                Sql4 = " Order by Articulo"
                spPreciosMp = Sql1 + Sql2 + Sql3 + Sql4
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                
                    With rstPreciosMp
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                If Cliente.Text = rstPreciosMp!Cliente Then
                                
                                ZEstado = IIf(IsNull(rstPreciosMp!Estado), "0", rstPreciosMp!Estado)
                                If Val(ZEstado) = Estado.ListIndex Then
                                
                                    ZArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                                    EntraVector = EntraVector + 1
                                    IngreVector(EntraVector, 1) = ZArticulo
                                    IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                                    IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                                    IngreVector(EntraVector, 4) = IIf(IsNull(rstPreciosMp!Precio), "0", Str$(rstPreciosMp!Precio))
                                    IngreVector(EntraVector, 5) = IIf(IsNull(rstPreciosMp!Fecha), "  /  /    ", rstPreciosMp!Fecha)
                                    
                                End If
                                
                                End If
                                
                                    .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstPreciosMp.Close
                End If
        
                For CicloVector = 1 To EntraVector
            
                    ZTerminado = IngreVector(CicloVector, 1)
                    WCliente = IngreVector(CicloVector, 2)
                    WArti = IngreVector(CicloVector, 3)
                    ZPrecio = IngreVector(CicloVector, 4)
                    ZFecha = IngreVector(CicloVector, 5)
                    ZDescripcion = ""
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    
                    DA = Len(ZDescripcion) - WEspacios
                    WIngresa = "N"
                    For Aaa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(ZDescripcion, Aaa, WEspacios) Then
                            WIngresa = "S"
                            Exit For
                        End If
                    Next Aaa
                    
                    If WIngresa = "S" Then
                        LugarPantalla = LugarPantalla + 1
                        PantallaPro.TextMatrix(LugarPantalla, 1) = ZTerminado
                        PantallaPro.TextMatrix(LugarPantalla, 2) = ZDescripcion
                        PantallaPro.TextMatrix(LugarPantalla, 3) = ZPrecio
                        PantallaPro.TextMatrix(LugarPantalla, 4) = Mid$(ZFecha, 4, 2) + "/" + Left$(ZFecha, 2) + "/" + Right$(ZFecha, 2)
            
                        IngresaItem = WCliente + WArti
                        WIndice.AddItem IngresaItem
                    End If
                
                Next CicloVector
                
                PantallaPro.Col = 1
                PantallaPro.Row = 1
                PantallaPro.TopRow = 1
            
            Case Else
            
        End Select
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
Rem BY NAN

End Sub

Private Sub StkPed()

    WTermi = WArticulo.Text
    WArti = Left$(WArticulo.Text, 3) + Right$(WArticulo.Text, 7)
    WStkPedido = 0

    spPedidoSol = "ListaPedidoTerminado " + "'" + WTermi + "'"
    Set rstPedidoSol = db.OpenRecordset(spPedidoSol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoSol.RecordCount > 0 Then
    
        With rstPedidoSol
        
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XPed = rstPedidoSol!Cantidad - rstPedidoSol!Facturado
                
                If XPed <> 0 Then
                    If Pedido.Text <> rstPedidoSol!Pedido Then
                        WStkPedido = WStkPedido + XPed
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        rstPedidoSol.Close
    End If

End Sub

Private Sub Calcula_Produccion()

    WTermi = WArticulo.Text
    WProduccion = 0

    Sql1 = "Select *"
    Sql2 = " FROM CargaSolicitud"
    Sql3 = " Where CargaSolicitud.Articulo = " + "'" + WArticulo.Text + "'"
    spCargaSolicitud = Sql1 + Sql2 + Sql3
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSolicitud.RecordCount > 0 Then
        With rstCargaSolicitud
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZSaldo = IIf(IsNull(rstCargaSolicitud!Saldo), "0", rstCargaSolicitud!Saldo)
                    WProduccion = WProduccion + ZSaldo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaSolicitud.Close
    End If

End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
    Pedido.SetFocus
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave.Text = "BMW" Then
            WGraba = "S"
            XClave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
            WClave.SetFocus
        End If
    End If
End Sub

Private Sub WImpres_Click()
    T$ = "Pedidos de Clientes"
    m$ = "Desea Imprimir el pedido"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
End Sub

Private Sub Busca_Stock()


    WTermi = WArticulo
    Call StkPed
    StkPedido.Caption = WStkPedido
    
    If Val(Wempresa) = 8 Then
        Call Calcula_Produccion
        Produccion.Caption = Str$(WProduccion)
            Else
        Produccion.Caption = ""
    End If

    Erase WVector
    If Val(Wempresa) = 1 Then
        Stock1.Caption = "SI"
        Stock2.Caption = "SII"
        Stock3.Caption = "SIII"
        Stock4.Caption = "SIV"
        Stock5.Caption = "SV"
        Stock6.Caption = "SVI"
        Stock7.Caption = "SVII"
            Else
        Stock1.Caption = "PI"
        Stock2.Caption = "PII"
        Stock3.Caption = "PV"
        Stock4.Caption = "PVI"
        Stock5.Caption = ""
        Stock6.Caption = ""
        Stock7.Caption = ""
    End If
    
    WStock1.Caption = ""
    WStock2.Caption = ""
    WStock3.Caption = ""
    WStock4.Caption = ""
    WStock5.Caption = ""
    Wstock6.Caption = ""
    Wstock7.Caption = ""
    
    If WArticulo = "  -     -   " Then Exit Sub

    If Left$(WArticulo, 2) <> "PT" And Left$(WArticulo, 2) <> "YQ" And Left$(WArticulo, 2) <> "YF" And Left$(WArticulo, 2) <> "YP" And Left$(WArticulo, 2) <> "YH" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
    
    Select Case WTipopro
        Case "M"
            WArti = Left$(WArticulo, 3) + Right$(WArticulo, 7)
            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
    
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        rstArticulo.Close
                            Else
                        WStock1.Caption = "0"
                    End If
            
                Case 8
                    Rem WEmpresa = "0002"
                    Rem txtOdbc = "Empresa02"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    Rem spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstArticulo.RecordCount > 0 Then
                    Rem     WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    Rem     rstArticulo.Close
                    Rem         Else
                    Rem     WStock1.Caption = "0"
                    Rem End If
    
                    Rem WEmpresa = "0004"
                    Rem txtOdbc = "Empresa04"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    Rem spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstArticulo.RecordCount > 0 Then
                    Rem     WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    Rem     rstArticulo.Close
                    Rem         Else
                    Rem     WStock2.Caption = "0"
                    Rem End If
                    
                    Rem WEmpresa = "0008"
                    Rem txtOdbc = "Empresa08"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        rstArticulo.Close
                            Else
                        WStock3.Caption = "0"
                    End If
            
                    Rem WEmpresa = "0008"
                    Rem txtOdbc = "Empresa08"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                Case Else
            End Select
            
            Stock.Caption = Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(Wstock6.Caption) + Val(Wstock7.Caption))
            Disponible.Caption = Str$(Val(Stock.Caption) - Val(StkPedido.Caption) + Val(Produccion.Caption))
            
            Rem Busca que envases hay
            
            WVector(1, 1) = ""
            WVector(2, 1) = ""
            WVector(3, 1) = ""
            WVector(4, 1) = ""
            WVector(5, 1) = ""
            WVector(6, 1) = ""
    
            XParam = "'" + WArti + "','" _
                 + WArti + "'"
            spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
    
                With rstLaudo
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                                Else
                    
                            If rstLaudo!Articulo = WArti Then
                
                                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                Call Redondeo(WSaldo)
                                
                                If WSaldo <> 0 Then
                                
                                    XEnv = IIf(IsNull(rstLaudo!Envase), "0", rstLaudo!Envase)
                                    WEnv = Str$(XEnv)
                                    For CicloEnvase = 1 To 6
                                        If Val(WEnv) = Val(WVector(CicloEnvase, 1)) Then
                                            WVector(CicloEnvase, 4) = Str$(Val(WVector(CicloEnvase, 4)) + WSaldo)
                                            Exit For
                                        End If
                                        If Val(WVector(CicloEnvase, 1)) = 0 Then
                                            WVector(CicloEnvase, 1) = WEnv
                                            WVector(CicloEnvase, 4) = Str$(WSaldo)
                                            Exit For
                                        End If
                                    Next CicloEnvase
                                
                                End If
                                
                            End If
                
                        End If
                
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                    End If
                End With
                rstLaudo.Close
            End If
            
            Call Carga_Envases
            
        Case "T"
            spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector(1, 1) = rstTerminado!Envase1
                WVector(2, 1) = rstTerminado!Envase2
                WVector(3, 1) = rstTerminado!Envase3
                WVector(4, 1) = rstTerminado!Envase4
                WVector(5, 1) = rstTerminado!Envase5
                WVector(6, 1) = rstTerminado!Envase6
                rstTerminado.Close
                Call Carga_Envases
            End If
            
            WSalidaError = ""
            On Error GoTo Control_error
    
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock1.Caption = "0"
                    End If
                    
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                         WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock2.Caption = "0"
                    End If
            
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock3.Caption = "0"
                    End If
                    
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock4.Caption = "0"
                    End If
                    
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock5.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock5.Caption = "0"
                    End If
                    
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        Wstock6.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        Wstock6.Caption = "0"
                    End If
                    
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        Wstock7.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        Wstock7.Caption = "0"
                    End If
                    
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                Case 8
                    Rem WEmpresa = "0002"
                    Rem txtOdbc = "Empresa02"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    Rem spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstTerminado.RecordCount > 0 Then
                    Rem     WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    Rem     rstTerminado.Close
                    Rem         Else
                    Rem     WStock1.Caption = "0"
                    Rem End If
    
                    Rem WEmpresa = "0004"
                    Rem txtOdbc = "Empresa04"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    Rem spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstTerminado.RecordCount > 0 Then
                    Rem     WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    Rem     rstTerminado.Close
                    Rem         Else
                    Rem     WStock2.Caption = "0"
                    Rem End If
                    
                    Rem WEmpresa = "0008"
                    Rem txtOdbc = "Empresa08"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spTerminado = "ConsultaTerminado " + "'" + WArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                            Else
                        WStock3.Caption = "0"
                    End If
            
                    Rem WEmpresa = "0008"
                    Rem txtOdbc = "Empresa08"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                Case Else
            End Select
            
            On Error GoTo 0
            
            ZZClave = Cliente.Text + WArticulo
            
            Stock.Caption = Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(Wstock6.Caption) + Val(Wstock7.Caption))
            Disponible.Caption = Str$(Val(Stock.Caption) - Val(StkPedido.Caption) + Val(Produccion.Caption))
            
        Case Else
    End Select

    For WDa = 1 To 6
        If Val(WVector(WDa, 1)) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + WVector(WDa, 1) + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                WVector(WDa, 3) = rstEnvase!Abreviatura
                rstEnvase.Close
            End If
        End If
    Next WDa
    
    Stock.Caption = Pusing("###,###.##", Stock.Caption)
    WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
    WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
    Wstock6.Caption = Pusing("###,###.##", Wstock6.Caption)
    Wstock7.Caption = Pusing("###,###.##", Wstock7.Caption)
    StkPedido.Caption = Pusing("###,###.##", StkPedido.Caption)
    Produccion.Caption = Pusing("###,###.##", Produccion.Caption)
    Disponible.Caption = Pusing("###,###.##", Disponible.Caption)
    
    If Val(Wempresa) <> 1 Then
        PantallaPro.Visible = False
    End If

    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Stock1.Visible = False
    Stock2.Visible = False
    Stock3.Visible = False
    Stock4.Visible = False
    Stock5.Visible = False
    Stock6.Visible = False
    Stock7.Visible = False
    WStock1.Visible = False
    WStock2.Visible = False
    WStock3.Visible = False
    WStock4.Visible = False
    WStock5.Visible = False
    Wstock6.Visible = False
    Wstock7.Visible = False
    Label7.Visible = False
    Disponible.Visible = False
    Resume Next

End Sub

Private Sub Busca_Partida()

    Erase Partida
    LugarPartida = 0
    
    XArti = Left$(WArticulo.Text, 3) + Right$(WArticulo.Text, 7)
    
    XParam = "'" + XArti + "','" _
                 + XArti + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                        Else
                    If rstLaudo!Articulo = XArti Then
                        WLote = rstLaudo!Laudo
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        Call Redondeo(WSaldo)
                        WAno = Right$(!Fecha, 4)
                        WMes = Mid$(!Fecha, 4, 2)
                        WDia = Left$(!Fecha, 2)
                        WFecha = WAno + WMes + WDia
                        If WSaldo <> 0 Then
                            LugarPartida = LugarPartida + 1
                            Partida(LugarPartida, 1) = Str$(WLote)
                            Partida(LugarPartida, 2) = Str$(WSaldo)
                            Partida(LugarPartida, 3) = WFecha
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    
    XParam = "'" + XArti + "','" _
                + XArti + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                        Else
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = XArti Then
                        WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                        WAno = Right$(!Fecha, 4)
                        WMes = Mid$(!Fecha, 4, 2)
                        WDia = Left$(!Fecha, 2)
                        WFecha = WAno + WMes + WDia
                        If WSaldo <> 0 Then
                            LugarPartida = LugarPartida + 1
                            Partida(LugarPartida, 1) = Str$(WLaudo)
                            Partida(LugarPartida, 2) = Str$(WSaldo)
                            Partida(LugarPartida, 3) = WFecha
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    For CicloPartida = 1 To LugarPartida
        For Dada = CicloPartida + 1 To LugarPartida
            If Partida(CicloPartida, 3) > Partida(Dada, 3) Then
                Auxi1 = Partida(CicloPartida, 1)
                Auxi2 = Partida(CicloPartida, 2)
                Auxi3 = Partida(CicloPartida, 3)
                
                Partida(CicloPartida, 1) = Partida(Dada, 1)
                Partida(CicloPartida, 2) = Partida(Dada, 2)
                Partida(CicloPartida, 3) = Partida(Dada, 3)
                
                Partida(Dada, 1) = Auxi1
                Partida(Dada, 2) = Auxi2
                Partida(Dada, 3) = Auxi3
            End If
        Next Dada
    Next CicloPartida
    
End Sub

Private Sub Limpia_PantallaPro()

    PantallaPro.Clear

    PantallaPro.FixedCols = 1
    PantallaPro.Cols = 5
    PantallaPro.FixedRows = 1
    PantallaPro.Rows = 1001
    
    PantallaPro.ColWidth(0) = 50
    PantallaPro.Row = 0
    For Ciclo = 1 To PantallaPro.Cols - 1
        PantallaPro.Col = Ciclo
        Select Case Ciclo
            Case 1
                PantallaPro.Text = "P.Terminado"
                PantallaPro.ColWidth(Ciclo) = 1400
                PantallaPro.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                PantallaPro.Text = "Descripcion"
                PantallaPro.ColWidth(Ciclo) = 3380
                PantallaPro.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                PantallaPro.Text = "Precio"
                PantallaPro.ColWidth(Ciclo) = 1200
                PantallaPro.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                PantallaPro.Text = "Fecha"
                PantallaPro.ColWidth(Ciclo) = 1200
                PantallaPro.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem WAncho = 400
    Rem For Ciclo = 0 To PantallaPro.Cols - 1
    Rem     WAncho = WAncho + PantallaPro.ColWidth(Ciclo)
    Rem Next Ciclo
    Rem PantallaPro.Width = WAncho

    PantallaPro.Col = 1
    PantallaPro.Row = 1
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
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

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim ZZVector(100, 2) As String
    Dim ZZAuxiliar(100, 3) As String
    
    Erase ZZAuxiliar
    ZZRenglon = 0
    
    ZZVector(1, 1) = Producto
    ZZVector(1, 2) = "1"
    ZZLugar = 1
    ZZCicla = 0
    
    Costo = 0
    
    Do
        ZZCicla = ZZCicla + 1
        If ZZVector(ZZCicla, 1) <> "" Then
    
            ZZEntra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + ZZVector(ZZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            If rstComposicion.RecordCount > 0 Then
                With rstComposicion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            ZZEntra = "N"
                        
                            ZZTipo = rstComposicion!Tipo
                            ZZArticulo1 = rstComposicion!Articulo1
                            ZZArticulo2 = rstComposicion!Articulo2
                            ZZCantidad = rstComposicion!Cantidad
                            
                            Select Case ZZTipo
                                Case "T"
                                    If Producto <> ZZArticulo2 Then
                                        ZZLugar = ZZLugar + 1
                                        ZZVector(ZZLugar, 1) = ZZArticulo2
                                        ZZVector(ZZLugar, 2) = Str$(ZZCantidad * Val(ZZVector(ZZCicla, 2)))
                                    End If
                                Case "M"
                                    ZZRenglon = ZZRenglon + 1
                                    ZZAuxiliar(ZZRenglon, 1) = ZZArticulo1
                                    ZZAuxiliar(ZZRenglon, 2) = ZZCantidad
                                    ZZAuxiliar(ZZRenglon, 3) = ZZVector(ZZCicla, 2)
                                Case Else
                            End Select
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To ZZRenglon
        ZZArticulo = ZZAuxiliar(DA, 1)
        ZZCantidad = ZZAuxiliar(DA, 2)
        ZZCantidadII = ZZAuxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Select Case ZTipoCosto
                Case 1
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
                Case 2
                    WCosto = (ZZCantidad * rstArticulo!Costo1 * Val(ZZCantidadII))
                Case 3
                    Costo4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
                    If Costo4 = 0 Then
                        Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                    End If
                    WCosto = (ZZCantidad * Costo4 * Val(ZZCantidadII))
                Case Else
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
            End Select
            Costo = Costo + WCosto
            rstArticulo.Close
        End If
    Next DA
    
    
End Sub

Private Sub CerrarPanta_Click()
    MuestraCosto.Visible = False
End Sub


Private Sub Verifica_Msds()
    
    ZZRequiereMsds = 0
    ZZRequiereMsdsCada = 0
    ZZBusqueda = ""
    ZZPasaMsds = "S"
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteEspecif"
    ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + Cliente.Text + "'"
    spClienteEspecif = ZSql
    Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstClienteEspecif.RecordCount > 0 Then
        ZZRequiereMsds = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
        ZZRequiereMsdsCada = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
        rstClienteEspecif.Close
    End If
    
    For a = 1 To 99
        
        Articulo = UCase(WVector1.TextMatrix(a, 1))
        ZZDescriArticulo = UCase(WVector1.TextMatrix(a, 2))
        Cantidad = WVector1.TextMatrix(a, 3)
        
        ZZValida = "S"
        WTerminado = Articulo
        XCodigo = Val(Mid$(WTerminado, 4, 5))
        If Left$(WTerminado, 2) = "PT" Then
        
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                        XTipoPro = "FA"
                            Else
                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                            XTipoPro = "BI"
                                Else
                            If XCodigo >= 40000 And XCodigo <= 49999 Then
                                XTipoPro = "TA"
                                    Else
                                XTipoPro = "PT"
                            End If
                        End If
                    End If
                End If
            End If
        
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZLinea = rstTerminado!Linea
                rstTerminado.Close
            End If
        
            Select Case ZLinea
                Case 8
                    XTipoPro = "PG"
                Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                    XTipoPro = "FA"
                Case Else
            End Select
    
            If XTipoPro = "FA" Or XTipoPro = "TA" Then
                ZZValida = "N"
            End If
            
        End If
        
        If Left$(WTerminado, 2) = "YQ" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "YH" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "YP" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "YF" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "ML" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "QC" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "ZE" Then
            ZZValida = "N"
        End If
        If Left$(WTerminado, 2) = "ZT" Then
            ZZValida = "N"
        End If
        
        If Val(Cantidad) <> 0 And ZZValida = "S" Then
            
            If Val(ZZRequiereMsdsCada) = 1 Then
            
                ZZBusqueda = "S"
                
                    Else
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Cliente = " + "'" + Cliente.Text + "'"
                ZSql = ZSql + " and Estadistica.Articulo = " + "'" + Articulo + "'"
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
                If ZZEstado = "" Then
                
                    ZZPasaMsds = "N"
                    
                    m$ = "El MSDS  (" + ZZCodArt + ")  no se ha encontrado"
                    AAAAA% = MsgBox(m$, 0, "Impresion de comprobantes varios")
                        
                    sTo = "dsuarez@surfactan.com.ar; pcorna@surfactan.com.ar; ebiglieri@surfactan.com.ar; isocalidad@surfacatn.com.ar; lalmiron@surfactan.com.ar; textil@surfactan.com.ar"
                    Rem sTo = "d_esquenazi@yahoo.com"
                    sCC = ""
                    sBCC = ""
                    sSubject = "Falta de MSDS"
                    sBody = "Falta de MSDS del " + ZZCodArt + "para el cliente " + DesCliente.Caption + "  Pedido Nro.:" + Pedido.Text + " fecha de entrega:" + FecEntrega.Text + "  (" + Tipoped.Text + ")"
                    SFile = ""
            
                    EmailAddress = sTo
                    CopiaAddress = sCC
                    MSubject = sSubject
                    MBody = sBody
                    MAttach = ""
                    MAttachI = ""
                    MAttachII = ""
                    MAttachIII = ""
                    MAttachIV = ""
                    MAttachVI = ""
                    MAttachVII = ""
                    MAttachVIII = ""
                    
                    SendEmail
                    
                End If
                
            End If
        End If
        
    Next a
    

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
        Rem If MAttachI <> "" Then
        Rem     .Attachments.Add MAttachI
        Rem End If
        Rem If MAttachII <> "" Then
        Rem     .Attachments.Add MAttachII
        Rem End If
        Rem If MAttachIII > "" Then
        Rem     .Attachments.Add MAttachIII
        Rem End If
        Rem If MAttachIV <> "" Then
        Rem     .Attachments.Add MAttachIV
        Rem End If
        Rem If MAttachV <> "" Then
        Rem     .Attachments.Add MAttachV
        Rem End If
        Rem If MAttachVI <> "" Then
        Rem     .Attachments.Add MAttachVI
        Rem End If
        Rem If MAttachVII <> "" Then
        Rem     .Attachments.Add MAttachVII
        Rem End If
        Rem If MAttachVIII <> "" Then
        Rem     .Attachments.Add MAttachVIII
        Rem End If
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




