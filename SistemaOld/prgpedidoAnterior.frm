VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedido 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Pedidos"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   300
   ClientWidth     =   11775
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11775
   Visible         =   0   'False
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
      Left            =   1200
      TabIndex        =   112
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
         TabIndex        =   113
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Adicionales 
      Caption         =   "Datos Adicionales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1560
      TabIndex        =   69
      Top             =   2280
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox Marca1 
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
         MaxLength       =   20
         TabIndex        =   75
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox Destino 
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
         MaxLength       =   30
         TabIndex        =   74
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox Marca3 
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
         MaxLength       =   20
         TabIndex        =   72
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Marca2 
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
         MaxLength       =   20
         TabIndex        =   71
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label13 
         Caption         =   "Destino"
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
         TabIndex        =   73
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Marca"
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
         TabIndex        =   70
         Top             =   480
         Width           =   975
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
      TabIndex        =   111
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid PantallaPro 
      Height          =   855
      Left            =   4200
      TabIndex        =   107
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      _Version        =   327680
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
      TabIndex        =   105
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
         TabIndex        =   106
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
      TabIndex        =   103
      Top             =   1920
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
         TabIndex        =   104
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
      TabIndex        =   90
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   94
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
         TabIndex        =   93
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
      Left            =   5160
      TabIndex        =   89
      Top             =   7080
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   59
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
      TabIndex        =   66
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
      TabIndex        =   65
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CtaCte 
      Caption         =   "Cuenta Corriente"
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
      TabIndex        =   45
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame IngreEnvases 
      Caption         =   "Ingreso de Envases"
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
      Left            =   9120
      TabIndex        =   36
      Top             =   3960
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
         TabIndex        =   64
         Top             =   1560
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
         TabIndex        =   63
         Top             =   1320
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
         TabIndex        =   62
         Top             =   1080
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
         TabIndex        =   61
         Top             =   840
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
         TabIndex        =   60
         Top             =   600
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
         TabIndex        =   58
         Top             =   360
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
         TabIndex        =   57
         Top             =   1560
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
         TabIndex        =   56
         Top             =   1320
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
         TabIndex        =   55
         Top             =   1080
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
         TabIndex        =   54
         Top             =   840
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
         TabIndex        =   53
         Top             =   600
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
         TabIndex        =   52
         Top             =   360
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
         TabIndex        =   51
         Top             =   1560
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
         TabIndex        =   50
         Top             =   1320
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
         TabIndex        =   49
         Top             =   1080
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
         TabIndex        =   48
         Top             =   840
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
         TabIndex        =   47
         Top             =   600
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
         TabIndex        =   46
         Top             =   360
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
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   6480
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
         Left            =   1680
         Picture         =   "prgpedidoAnterior.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   102
         Top             =   960
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
         TabIndex        =   101
         Top             =   960
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
         TabIndex        =   100
         Top             =   1200
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
         TabIndex        =   99
         Top             =   1200
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
         TabIndex        =   98
         Top             =   720
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
         TabIndex        =   97
         Top             =   720
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
         TabIndex        =   80
         Top             =   480
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
         TabIndex        =   79
         Top             =   240
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
         TabIndex        =   78
         Top             =   480
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
         TabIndex        =   77
         Top             =   240
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
      Left            =   240
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
      TabIndex        =   86
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
      TabIndex        =   87
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
      TabIndex        =   82
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
      TabIndex        =   84
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
      TabIndex        =   96
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
      TabIndex        =   95
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Entra 
      Caption         =   "Datos Adicionales"
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
      TabIndex        =   76
      Top             =   3120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   114
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
      _Version        =   327680
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
      TabIndex        =   88
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
      TabIndex        =   85
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
      TabIndex        =   83
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
      TabIndex        =   81
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
End
Attribute VB_Name = "PrgPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 4 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WImpre(10) As String
Private WEnvase(10) As String
Private WVector(6, 4) As String
Private XEnvase(100, 6) As String
Private XEspecificaciones(100) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private WTermi As String
Private WStkPedido As Double
Private WProduccion As Double

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
Dim rstPago As Recordset
Dim spPago As String
Dim rstCtacte As Recordset
Dim spCtacte As String
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

Private Sub Baja_Click()

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Erase Auxiliar
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
    
            With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstPedido!Terminado
                    Auxiliar(WRenglon, 2) = rstPedido!Cantidad
                    Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For DA = 1 To WRenglon
        Select Case Auxiliar(DA, 3)
            Case "M"
                Articulo = Left$(Auxiliar(DA, 1), 3) + Right$(Auxiliar(DA, 1), 7)
                Cantidad = Auxiliar(DA, 2)
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WVenta = Str$(rstArticulo!Venta - Cantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                            + WVenta + "','" _
                            + WDate + "'"
                    rstArticulo.Close
                    spArticulo = "ModificaArticuloVenta " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
            Case Else
                Terminado = Auxiliar(DA, 1)
                Cantidad = Auxiliar(DA, 2)
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WPedido = Str$(rstTerminado!Pedido - Cantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WDate + "'"
                    rstTerminado.Close
                    spTerminado = "ModificaTerminadoPedido " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
        End Select
    Next DA
    
    spPedido = "BorrarPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenDynaset, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "DELETE Muestra"
    ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
    spMuestra = ZSql
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Pedido.SetFocus

End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = "  -     -   "
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLinea.Text = ""
    
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
    Stock.Caption = ""
    StkPedido.Caption = ""
    Produccion.Caption = ""
    Disponible.Caption = ""
    
    WArticulo.SetFocus
    Call DBGrid1_GotFocus
    
End Sub

Private Sub BorraConsulta_Click()
    Pantalla.Visible = False
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click
    PrgPedido.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Pedido SET "
    ZSql = ZSql + "HojaRuta = 1"
                     
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Rem WDesde = "00000000"
    Rem WHasta = "00000000"

    Rem XParam = "'" + WDesde + "','" _
    REM             + WHasta + "'"
    Rem spPedido = "ModificaPedidonousar " + XParam
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem XParam = "'" + Pedido.Text + "','" _
    rem              + "1" + " '"
    Rem spPedido = "ModificaPedidoProceso1 " + XParam
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem Call Limpia_Click

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

Private Sub CtaCte_Click()
    PCliente = Cliente.Text
    PTipo = 1
    Rem PrgConsCcte.Show
    PrgCC1.Show
    PTipo = 0
End Sub

Private Sub Entra_Click()
    Adicionales.Visible = True
    Marca1.SetFocus
End Sub

Private Sub Inserta_Click()

    WPrimer = DBGrid1.FirstRow
    WFila = DBGrid1.Row
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    Erase WPasa
    
    For Cicla = 99 To WLugar + 1 Step -1
            XEnvase(Cicla, 1) = XEnvase(Cicla - 1, 1)
            XEnvase(Cicla, 2) = XEnvase(Cicla - 1, 2)
            XEnvase(Cicla, 3) = XEnvase(Cicla - 1, 3)
            XEnvase(Cicla, 4) = XEnvase(Cicla - 1, 4)
            XEnvase(Cicla, 5) = XEnvase(Cicla - 1, 5)
            XEnvase(Cicla, 6) = XEnvase(Cicla - 1, 6)
            XEspecificaciones(Cicla) = XEspecificaciones(Cicla - 1)
    Next Cicla
    
    XEnvase(WLugar, 1) = ""
    XEnvase(WLugar, 2) = ""
    XEnvase(WLugar, 3) = ""
    XEnvase(WLugar, 4) = ""
    XEnvase(WLugar, 5) = ""
    XEnvase(WLugar, 6) = ""
    XEspecificaciones(WLugar) = ""
        
    DBGrid1.Refresh

    Erase WPasa
    Salida = "N"
    
    XCounter = 0
    XLugar = 0

    For A = 0 To 5
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
    
            WRow = iRow
            DBGrid1.Col = 0
            DBGrid1.Row = iRow
    
            If DBGrid1.Text = "" Then
                Salida = "S"
                Exit For
            End If
    
            XLugar = XLugar + 1
            XCounter = XCounter + 1
                
            WRow = iRow
            DBGrid1.Row = iRow
            
            DBGrid1.Col = 0
            WPasa(XCounter, 1) = DBGrid1.Text
            DBGrid1.Text = ""
            
            DBGrid1.Col = 1
            WPasa(XCounter, 2) = DBGrid1.Text
            DBGrid1.Text = ""
            
            DBGrid1.Col = 2
            WPasa(XCounter, 3) = DBGrid1.Text
            DBGrid1.Text = ""
            
            DBGrid1.Col = 3
            WPasa(XCounter, 4) = DBGrid1.Text
            DBGrid1.Text = ""
                
            If XLugar = WLugar - 1 Then
                XCounter = XCounter + 1
                WPasa(XCounter, 1) = "  -     -   "
                WPasa(XCounter, 2) = ""
                WPasa(XCounter, 3) = ""
                WPasa(XCounter, 4) = ""
            End If
    
        Next iRow
        If Salida = "S" Then
            Exit For
        End If
    Next A
     
    WLugar = 0
    Salida = "N"
            
    For A = 0 To 5
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            WLugar = WLugar + 1
            DBGrid1.Row = iRow
            DBGrid1.Col = 0
            DBGrid1.Text = WPasa(WLugar, 1)
            DBGrid1.Col = 1
            DBGrid1.Text = WPasa(WLugar, 2)
            DBGrid1.Col = 2
            DBGrid1.Text = WPasa(WLugar, 3)
            DBGrid1.Col = 3
            DBGrid1.Text = WPasa(WLugar, 4)
            If WLugar = XCounter Then
                Salida = "S"
                Exit For
            End If
        Next iRow
            
        If Salida = "S" Then
            Exit For
        End If
            
    Next A
    
    Renglon = Renglon + 1
    
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    DBGrid1.Col = 0
    
    DBGrid1.FirstRow = WPrimer
    DBGrid1.Row = WFila
    DBGrid1.Col = 0
    
    Call DBGrid1_GotFocus
    
End Sub

Private Sub marca1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Marca2.SetFocus
    End If
End Sub

Private Sub marca2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Marca3.SetFocus
    End If
End Sub

Private Sub marca3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Destino.SetFocus
    End If
End Sub

Private Sub Destino_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pedido.SetFocus
        Adicionales.Visible = False
    End If
End Sub

Private Sub DBGrid1_GotFocus()

    aa = Len(DBGrid1.Text)

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = Pusing("###,###.##", DBGrid1.Text)
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WPrecio.Caption = Pusing("###,###.##", DBGrid1.Text)
    
    Rem aca se reemplaza la rutina de cambio envase
    
    WPrimer = DBGrid1.FirstRow
    WFila = DBGrid1.Row
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    Envase1.Text = XEnvase(WLugar, 1)
    Canti1.Text = XEnvase(WLugar, 2)
    Envase2.Text = XEnvase(WLugar, 3)
    Canti2.Text = XEnvase(WLugar, 4)
    Envase3.Text = XEnvase(WLugar, 5)
    Canti3.Text = XEnvase(WLugar, 6)
    
    Especificaciones.Text = XEspecificaciones(WLugar)
    
    Rem If Val(WLinea.Text) <> 0 Then
    Rem     Envase1.Text = XEnvase(Val(WLinea.Text) + WInicio, 1)
    Rem     Canti1.Text = XEnvase(Val(WLinea.Text) + WInicio, 2)
    Rem     Envase2.Text = XEnvase(Val(WLinea.Text) + WInicio, 3)
    Rem     Canti2.Text = XEnvase(Val(WLinea.Text) + WInicio, 4)
    Rem     Envase3.Text = XEnvase(Val(WLinea.Text) + WInicio, 5)
    Rem     Canti3.Text = XEnvase(Val(WLinea.Text) + WInicio, 6)
    Rem End If
    
    WTermi = WArticulo.Text
    Call StkPed
    StkPedido.Caption = WStkPedido
    
    If Left$(WTermi, 2) <> "ML" Then
        Call Busca_Stock
    End If
    
    WInicio = DBGrid1.FirstRow
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    If FecEntrega.Text = "00/00/0000" Or FecEntrega.Text = "  /  /    " Then
        m$ = "No esta informada la fecha de entrega"
        A% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
        Exit Sub
    End If
    
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdFecEntrega = Right$(FecEntrega.Text, 4) + Mid$(FecEntrega.Text, 4, 2) + Left$(FecEntrega.Text, 2)
    If WFechaord > WOrdFecEntrega Then
        m$ = "La fecha de entrega no puede ser menor a la fecha del pedido"
        A% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
        Exit Sub
    End If
    
    WFechaInicial = FecEntrega.Text
    WOrdFechaInicial = WOrdFecEntrega
    
    If Val(WEmpresa) = 8 Then
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            rstPedido.Close
            If WGraba <> "S" Then
                Call Ingresa_clave
                Exit Sub
            End If
        End If
    End If

    Renglon = ZZRenglon
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    
    Rem verifica que no halla dos articulos con
    Rem distinto tipo de condicion de pago
    
    XPasa = "S"
    WLLave = 0
    
    ZImpreVto = 0
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
        rstCliente.Close
    End If
    
    For A = 0 To 5
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
            
            If Val(Cantidad) <> 0 Then
            
                WCliente = UCase(Cliente.Text)
                WTerminado = UCase(Articulo)
                WClave = WCliente + WTerminado
                Xpago = 0
    
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
                        Case "DY", "DW", "DS"
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
                        If XCodigo >= 11000 And XCodigo <= 11999 Then
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
                    Case 10, 20
                        XTipoPro = "FA"
                    Case Else
                End Select
                
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
                
                    If Val(WEmpresa) = 1 And Cliente.Text = "P00005" Then
                    
                        ZPasa = "S"
                    
                        WEmpresa = "0008"
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
                                ca% = MsgBox(m$, 0, "Emision de Facturas")
                                ZPasa = "N"
                            End If
                            
                                Else
                                
                            m$ = "Producto Terminado Inexistente"
                            ca% = MsgBox(m$, 0, "Emision de Facturas")
                            ZPasa = "N"
                            
                        End If
                        
                        WEmpresa = "0001"
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
                                    ca% = MsgBox(m$, 0, "Emision de Facturas")
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
                                ca% = MsgBox(m$, 0, "Emision de Facturas")
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
                        A% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
                        Exit Sub
                    End If
                    
                End If
                        
            End If
                                        
        Next iRow
            
    Next A
    
    If Val(WEmpresa) = 1 Then
    
        If XPasa = "1" Then
    
            m$ = "Se cargaron articulos condistinta condicion de pago"
            A% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
    
            Renglon = Renglon + 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = ""
            Renglon = Renglon - 1

            Exit Sub
        
        End If

        If XPasa = "2" Then
    
            m$ = "Se cargaron articulos PT, Biosidas, Farma, Pigmentos o Colorantes en forma conjunta un mismo Pedido"
            A% = MsgBox(m$, 0, "INGRESO DE PEDIDOS")
    
            Renglon = Renglon + 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                    
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = ""
            Renglon = Renglon - 1

            Exit Sub
        
        End If
    
    End If

    Xversion = 0
    Erase Auxiliar
    WRenglon = 0
    
    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
    
            Xversion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
            
            XFechaInicial = IIf(IsNull(rstPedido!FechaInicial), "", rstPedido!FechaInicial)
            XOrdFechaInicial = IIf(IsNull(rstPedido!OrdFechaInicial), "", rstPedido!OrdFechaInicial)
            
            If XFechaInicial <> "" Then
                WFechaInicial = XFechaInicial
            End If
            If XOrdFechaInicial <> "" Then
                WOrdFechaInicial = XOrdFechaInicial
            End If
    
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Terminado
                        Auxiliar(WRenglon, 2) = rstPedido!Cantidad
                        Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
    End If
    
    For DA = 1 To WRenglon
        Select Case Auxiliar(DA, 3)
            Case "M"
                Articulo = Left$(Auxiliar(DA, 1), 3) + Right$(Auxiliar(DA, 1), 7)
                Cantidad = Auxiliar(DA, 2)
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WVenta = Str$(rstArticulo!Venta - Cantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                            + WVenta + "','" _
                            + WDate + "'"
                    rstArticulo.Close
                    spArticulo = "ModificaArticuloVenta " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            Case Else
                Terminado = Auxiliar(DA, 1)
                Cantidad = Auxiliar(DA, 2)
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WPedido = Str$(rstTerminado!Pedido - Cantidad)
                    WDate = Date$
                    rstTerminado.Close
                    XParam = "'" + WCodigo + "','" _
                                + WPedido + "','" _
                                + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoPedido " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
        End Select
    Next DA
    
    spPedido = "BorrarPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenDynaset, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "DELETE Muestra"
    ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
    spMuestra = ZSql
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)

    Erase Auxiliar
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    DoEvents
        
    For A = 0 To 5
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
            
            DBGrid1.Col = 1
            NombreComercial = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                                            
            WInicio = DBGrid1.FirstRow
                        
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
                Rem WEnvase1 = XEnvase(Renglon, 1)
                Rem WCanti1 = XEnvase(Renglon, 2)
                Rem WEnvase2 = XEnvase(Renglon, 3)
                Rem WCanti2 = XEnvase(Renglon, 4)
                Rem WEnvase3 = XEnvase(Renglon, 5)
                Rem WCanti3 = XEnvase(Renglon, 6)
                WEnvase4 = 0
                WCanti4 = ""
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WOrdFecEntrega = Right$(FecEntrega.Text, 4) + Mid$(FecEntrega.Text, 4, 2) + Left$(FecEntrega.Text, 2)
                WPrecio = Precio
                WWLinea = Linea
                WFacturado = ""
                WImporte = ""
                WClave = Auxi1 + Auxi
                WMarca1 = Marca1.Text
                WMarca2 = Marca2.Text
                WMarca3 = Marca3.Text
                WDestino = Destino.Text
                WAutorizo = "N"
                WImpresion = "N"
                WTipoped = Str$(Tipoped.ListIndex)
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
                
                XParam = "'" + WClave + "','" _
                             + WPedido + "','" + WRenglon + "','" + WCliente + "','" + WFecha + "','" _
                             + WFecEntrega + "','" + WHora + "','" + WObservaciones + "','" + WTerminado + "','" _
                             + WCantidad + "','" + WEnvase1 + "','" + WCanti1 + "','" _
                             + WEnvase2 + "','" + WCanti2 + "','" _
                             + WEnvase3 + "','" + WCanti3 + "','" _
                             + WEnvase4 + "','" + WCanti4 + "','" _
                             + WFechaord + "','" + WPrecio + "','" + WWLinea + "','" + WFacturado + "','" _
                             + WImporte + "','" + WMarca1 + "','" _
                             + WMarca2 + "','" + WMarca3 + "','" + WDestino + "','" _
                             + WAutorizo + "','" + WImpresion + "','" + WTipoped + "','" _
                             + WCantidad1 + "','" + WCantidad2 + "','" _
                             + WLote1 + "','" + WCantiLote1 + "','" + WLote2 + "','" + WCantiLote2 + "','" _
                             + Wlote3 + "','" + WCantiLote3 + "','" + WLote1 + "','" + WCantiLote4 + "','" _
                             + WLote5 + "','" + WCantiLote5 + "','" _
                             + WEnv1 + "','" + WCantiEnv1 + "','" + WEnv2 + "','" + WCantiEnv2 + "','" _
                             + WEnv3 + "','" + WCantiEnv3 + "','" + WEnv4 + "','" + WCantiEnv4 + "','" _
                             + WEnv5 + "','" + WCantiEnv5 + "','" _
                             + WVersion + "','" _
                             + WOrdFecEntrega + "','" _
                             + WOrdenCpa + "','" _
                             + WTipopro + "','" _
                             + WArti + "'"

                spPedido = "AltaPedido " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + "NombreComercial = " + "'" + NombreComercial + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                     
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                 
                ZLinea = 0
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZLinea = rstTerminado!Linea
                    rstTerminado.Close
                End If
                
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                If (XCodigo >= 25000 And XCodigo <= 25999) Or ZLinea = 10 Or ZLinea = 20 Then
                    If Val(WEmpresa) = 1 Then
                
                        XParam = "'" + WClave + "','" _
                                     + WEspecificaciones + "'"
                        spPedido = "ModificaPedidoEspecificaciones " + XParam
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    
                        XParam = "'" + WCliente + "','" _
                                     + WTerminado + "'"
                        spEspeCli = "ConsultaEspeCli " + XParam
                        Set rstEspeCli = db.OpenRecordset(spEspeCli, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEspeCli.RecordCount > 0 Then
                            rstEspeCli.Close
                            XParam = "'" + WCliente + "','" _
                                         + WTerminado + "','" _
                                         + WEspecificaciones + "'"
                            spEspeCli = "ModificaEspeCli " + XParam
                            Set rstEspeCli = db.OpenRecordset(spEspeCli, dbOpenSnapshot, dbSQLPassThrough)
                                Else
                            XParam = "'" + WCliente + "','" _
                                         + WTerminado + "','" _
                                         + WEspecificaciones + "'"
                            spEspeCli = "AltaEspeCli " + XParam
                            Set rstEspeCli = db.OpenRecordset(spEspeCli, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                End If
                
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                If Left$(WTerminado, 2) <> "PT" Then
                    Select Case Left$(WTerminado, 2)
                        Case "DY", "DW", "DS"
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
                        If XCodigo >= 11000 And XCodigo <= 11999 Then
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
                    Case 10, 20
                        XTipoPro = "FA"
                    Case Else
                End Select
                
                If Val(WEmpresa) = 8 Then
                    If ZLinea = 16 Then
                        XTipoPro = "CO"
                            Else
                        XTipoPro = "PT"
                    End If
                End If
                
                If Left$(WTerminado, 2) = "ML" Then
                    Select Case XCodigo
                        Case 6
                            XTipoPro = "CO"
                        Case 8
                            XTipoPro = "PG"
                        Case 10
                            XTipoPro = "FA"
                        Case 14
                            XTipoPro = "BI"
                        Case Else
                            XTipoPro = "PT"
                    End Select
                End If
                
            End If
                                        
        Next iRow
            
    Next A
    
    Select Case XTipoPro
        Case "CO"
            XParam = "'" + WPedido + "','" _
                         + "1" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        Case "FA"
            XParam = "'" + WPedido + "','" _
                         + "4" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        Case "BI"
            XParam = "'" + WPedido + "','" _
                         + "3" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        Case "PT"
            XParam = "'" + WPedido + "','" _
                         + "2" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        Case "PG"
            XParam = "'" + WPedido + "','" _
                         + "5" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            WMarca = "X"
            XParam = "'" + WPedido + "','" _
                         + WMarca + "'"
            spPedido = "ModificaPedidoPigmentos " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        Case Else
            XParam = "'" + WPedido + "','" _
                         + "0" + "'"
            spPedido = "ModificaPedidoTipoPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    For DA = 1 To WRenglon
    
        Terminado = Auxiliar(DA, 1)
        Cantidad = Val(Auxiliar(DA, 2))
        If Left$(Terminado, 2) <> "PT" And Left$(Terminado, 2) <> "YQ" And Left$(Terminado, 2) <> "YF" And Left$(Terminado, 2) <> "YP" And Left$(Terminado, 2) <> "YH" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
    
        Select Case WTipopro
            Case "M"
                Articulo = Left$(Terminado, 3) + Right$(Terminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    XVenta = IIf(IsNull(rstArticulo!Venta), 0, rstArticulo!Venta)
                    WVenta = Str$(XVenta + Cantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                            + WVenta + "','" _
                            + WDate + "'"
                    rstArticulo.Close
                    spArticulo = "ModificaArticuloVenta " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            Case Else
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WPedido = Str$(rstTerminado!Pedido + Cantidad)
                    WDate = Date$
                    rstTerminado.Close
                    XParam = "'" + WCodigo + "','" _
                                    + WPedido + "','" _
                                    + WDate + "'"
                                            
                    spTerminado = "ModificaTerminadoPedido " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End Select
    Next DA
    
    If Tipoped.ListIndex = 6 Then
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + "HojaRuta = 999999" + ","
        ZSql = ZSql + "TipoPed = 5" + ","
        ZSql = ZSql + "FechaInicial = " + "'" + WFechaInicial + "',"
        ZSql = ZSql + "OrdFechaInicial = " + "'" + WOrdFechaInicial + "',"
        ZSql = ZSql + "DirEntrega = " + "'" + Str$(ZLugarDirEntrega) + "'"
        ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
                     
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

            Else
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + "HojaRuta = 0" + ","
        ZSql = ZSql + "FechaInicial = " + "'" + WFechaInicial + "',"
        ZSql = ZSql + "OrdFechaInicial = " + "'" + WOrdFechaInicial + "',"
        ZSql = ZSql + "DirEntrega = " + "'" + Str$(ZLugarDirEntrega) + "'"
        ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
                        
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    T$ = "Pedidos de Clientes"
    m$ = "Desea Imprimir el pedido"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
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
    
    Marca1.Text = ""
    Marca2.Text = ""
    Marca3.Text = ""
    Destino.Text = ""
    
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
    
    Pantalla.Visible = False
    PantallaPro.Visible = False
    Ayuda.Visible = False
    
    Erase XEnvase
    Erase XEspecificaciones
    
    For A = 0 To 5
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 3
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Pedido.Text = ""
    spPedido = "ListaPedidoNumero"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveLast
            Pedido.Text = rstPedido!Pedido + 1
        End With
        rstPedido.Close
    End If
    
    DBGrid1.FirstRow = 0
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
                                EntraNombre.Visible = True
                                NombreComercial.Text = ""
                                NombreComercial.SetFocus
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
        If (XCodigo >= 25000 And XCodigo <= 25999) Or ZLinea = 10 Or ZLinea = 20 Then
            If Val(WEmpresa) = 1 Then
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
                Else
            Envase1.SetFocus
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NombreComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Left$(WArticulo.Text, 2) = "ML" Then
        
            WDescripcion.Caption = NombreComercial.Text
            WCantidad.SetFocus
            EntraNombre.Visible = False
        
                Else
    
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
            If Ingre = "S" Then
                Canti1.SetFocus
                    Else
                Envase1.SetFocus
            End If
                Else
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
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
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
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
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
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
                
                DBGrid1.Col = 0
                DBGrid1.Text = WArticulo.Text
                DBGrid1.Col = 3
                If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                    DBGrid1.Text = Pusing("###,###.##", "0")
                        Else
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                End If
                
                rstPreciosMp.Close
                        
                If Left$(WArticulo.Text, 2) <> "ML" Then
                    Call Busca_Stock
                End If
                        
                DBGrid1.Col = 1
                DBGrid1.Text = WDescripcion.Caption
                
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
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
                    
                DBGrid1.Col = 0
                DBGrid1.Text = rstPrecios!Terminado
                DBGrid1.Col = 1
                DBGrid1.Text = rstPrecios!Descripcion
                DBGrid1.Col = 3
                If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                    DBGrid1.Text = Pusing("###,###.##", "0")
                        Else
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                End If
                
                rstPrecios.Close
                        
                Call Busca_Stock
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case Else
    End Select

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
                    
                        DBGrid1.Col = 0
                        DBGrid1.Text = WArticulo.Text
                        DBGrid1.Col = 3
                        If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                            DBGrid1.Text = Pusing("###,###.##", "0")
                                Else
                            DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                        End If
                
                        rstPreciosMp.Close
                        
                        Call Busca_Stock
                        
                        DBGrid1.Col = 1
                        DBGrid1.Text = WDescripcion.Caption
                
                        Call Alta_Vector
                        WLinea.Text = WAnterior + 1
                        If Val(WLinea.Text) > 0 Then
                            DBGrid1.Row = Val(WLinea.Text) - 1
                        End If
                    
                        Call DBGrid1.SetFocus
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
                    
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstPrecios!Terminado
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstPrecios!Descripcion
                        DBGrid1.Col = 3
                        If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                            DBGrid1.Text = Pusing("###,###.##", "0")
                                Else
                            DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                        End If
                
                        rstPrecios.Close
                        
                        Call Busca_Stock
                    
                        Call Alta_Vector
                        WLinea.Text = WAnterior + 1
                        If Val(WLinea.Text) > 0 Then
                            DBGrid1.Row = Val(WLinea.Text) - 1
                        End If
                    
                        Call DBGrid1.SetFocus
                        WCantidad.SetFocus
                    
                    End If
            
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 3, 0 To 80)

mTotalRows& = 80

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
For i = 0 To 3
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 4000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
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

    Pantalla.Visible = False
    Pedido.Text = ""
    spPedido = "ListaPedidoNumero"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveLast
            Pedido.Text = rstPedido!Pedido + 1
        End With
        rstPedido.Close
    End If
    
    Tipoped.Clear
    
    Tipoped.AddItem "Normal"
    Tipoped.AddItem "a Fecha"
    Tipoped.AddItem "Fecha Limite"
    Tipoped.AddItem "Urgente"
    Tipoped.AddItem "Retira Cliente"
    Tipoped.AddItem "Muestra"
    Tipoped.AddItem "Muestra Retira"
    
    Tipoped.ListIndex = 0

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FecEntrega.Text = "  /  /    "
    Version.Text = ""
    
    WCalcula = "N"
    FecEntrega.Enabled = False
    WCalcula = "S"
    
    Pedido.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Erase XEnvase
    Erase XEspecificaciones

    For A = 0 To 5
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 3
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Auxiliar
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
            WGraba = "N"
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstPedido!Terminado
                        Auxi1 = rstPedido!Terminado
                        
                        DBGrid1.Col = 1
                        DBGrid1.Text = ""
                
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        
                        Rem aca se reemplaza la rutina de cambio envase
    
                        WPrimer = DBGrid1.FirstRow
                        WFila = DBGrid1.Row
                        WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
                        XEnvase(WLugar, 1) = rstPedido!Envase1
                        XEnvase(WLugar, 2) = rstPedido!Canti1
                        XEnvase(WLugar, 3) = rstPedido!Envase2
                        XEnvase(WLugar, 4) = rstPedido!Canti2
                        XEnvase(WLugar, 5) = rstPedido!Envase3
                        XEnvase(WLugar, 6) = rstPedido!Canti3
                        
                        XEspecificaciones(WLugar) = IIf(IsNull(rstPedido!Especificaciones), "0", rstPedido!Especificaciones)
                        
                        Rem XEnvase(Renglon, 1) = rstPedido!Envase1
                        Rem XEnvase(Renglon, 2) = rstPedido!Canti1
                        Rem XEnvase(Renglon, 3) = rstPedido!Envase2
                        Rem XEnvase(Renglon, 4) = rstPedido!Canti2
                        Rem XEnvase(Renglon, 5) = rstPedido!Envase3
                        Rem XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!Terminado
                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                            Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
                        End If
                        
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
    End If
    
    Renglon = 0
    
    For DA = 1 To WRenglon
    
        Cliente = Auxiliar(DA, 1)
        Terminado = Auxiliar(DA, 2)
        ZZNombreComercial = Auxiliar(DA, 3)
        
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
        
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 3
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        DBGrid1.Text = Pusing("###,###.##", "0")
                            Else
                        DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                    End If
                    
                    rstPreciosMp.Close
                    
                    WArticulo.SetFocus
                    
                        Else
        
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                        
                    DBGrid1.Col = 0
                    DBGrid1.Text = Terminado
                        
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", "0")
                    
                    WArticulo.SetFocus
                    
                End If
                    
                If Trim(ZZNombreComercial) = "" Then
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                        Else
                    DBGrid1.Col = 1
                    DBGrid1.Text = ZZNombreComercial
                End If
            
            Case Else
                ZZDescripcion = ""
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
        
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
                    ZZDescripcion = rstPrecios!Descripcion
            
                    DBGrid1.Col = 3
                    If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
                        DBGrid1.Text = Pusing("###,###.##", "0")
                            Else
                        DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                    End If
                    
                    rstPrecios.Close
                    
                    If Trim(ZZDescripcion) = "" Then
                        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            DBGrid1.Col = 1
                            DBGrid1.Text = rstTerminado!Descripcion
                            rstTerminado.Close
                        End If
                    End If
                    
                    WArticulo.SetFocus
                    
                End If
                
        End Select
        
    Next DA

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    Call DBGrid1_GotFocus
   
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WArticulo.SetFocus
    Call DBGrid1_GotFocus

End Sub

Private Sub Alta_Vector()

    GoTo DA

    If Left$(WArticulo.Text, 2) = "PT" And Mid$(WArticulo.Text, 4, 2) <> "25" Then
        spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            
            WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
            WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
            WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
            rstTerminado.Close
                
            If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                m$ = "El Producto Terminado no se encuentra autorizado para la Venta"
                If WEstadoI = "N" Then
                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                End If
                If WEstadoII = "N" Then
                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                End If
                If WEstadoIII = "N" Then
                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                End If
                ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                
                Exit Sub
            End If
                
        End If
    End If
    
DA:

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

            Renglon = Renglon + 1
            ZZRenglon = Renglon
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            Rem aca se reemplaza la rutina de cambio envase
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
            
            XEnvase(WLugar, 1) = Envase1.Text
            XEnvase(WLugar, 2) = Canti1.Text
            XEnvase(WLugar, 3) = Envase2.Text
            XEnvase(WLugar, 4) = Canti2.Text
            XEnvase(WLugar, 5) = Envase3.Text
            XEnvase(WLugar, 6) = Canti3.Text
            
            XEspecificaciones(WLugar) = Especificaciones.Text
            
            Rem XEnvase(Renglon, 1) = Envase1.Text
            Rem XEnvase(Renglon, 2) = Canti1.Text
            Rem XEnvase(Renglon, 3) = Envase2.Text
            Rem XEnvase(Renglon, 4) = Canti2.Text
            Rem XEnvase(Renglon, 5) = Envase3.Text
            Rem XEnvase(Renglon, 6) = Canti3.Text
            
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
    
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            If Renglon = 1 Then
                 DBGrid1.Row = DBGrid1.Row + 1
                 DBGrid1.Col = 0
                 DBGrid1.Text = ""
            End If
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
            
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            Rem aca se reemplaza la rutina de cambio envase
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
            
            XEnvase(WLugar, 1) = Envase1.Text
            XEnvase(WLugar, 2) = Canti1.Text
            XEnvase(WLugar, 3) = Envase2.Text
            XEnvase(WLugar, 4) = Canti2.Text
            XEnvase(WLugar, 5) = Envase3.Text
            XEnvase(WLugar, 6) = Canti3.Text
            
            XEspecificaciones(WLugar) = Especificaciones.Text
            
            Rem XEnvase(Val(WLinea.Text) + WInicio, 1) = Envase1.Text
            Rem XEnvase(Val(WLinea.Text) + WInicio, 2) = Canti1.Text
            Rem XEnvase(Val(WLinea.Text) + WInicio, 3) = Envase2.Text
            Rem XEnvase(Val(WLinea.Text) + WInicio, 4) = Canti2.Text
            Rem XEnvase(Val(WLinea.Text) + WInicio, 5) = Envase3.Text
            Rem XEnvase(Val(WLinea.Text) + WInicio, 6) = Canti3.Text
                
            Rem DbGrid1.Row = Renglon
            DBGrid1.Row = Lugar2 - 1
            DBGrid1.Col = 0
            WInicio = DBGrid1.FirstRow
            
            DBGrid1.Row = DBGrid1.Row + 1
            
    End If

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Fecha.Text = rstPedido!Fecha
            Cliente.Text = rstPedido!Cliente
            FecEntrega.Text = rstPedido!FecEntrega
            Hora.Text = rstPedido!Hora
            Observaciones.Text = rstPedido!Observaciones
            OrdenCpa.Text = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
            Marca1.Text = rstPedido!Marca1
            Marca2.Text = rstPedido!Marca2
            Marca3.Text = rstPedido!Marca3
            Destino.Text = rstPedido!Destino
            Tipoped.ListIndex = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
            Version.Text = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
            ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            rstPedido.Close
            
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
        Hora.SetFocus
            Else
        FecEntrega.Enabled = True
        FecEntrega.SetFocus
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

 For Cicla = 1 To 6
    spEnvase = "ConsultaEnvases " + "'" + WVector(Cicla, 1) + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        WVector(Cicla, 2) = rstEnvase!Kilos
        WVector(Cicla, 3) = rstEnvase!Abreviatura
        rstEnvase.Close
                Else
        WVector(Cicla, 2) = ""
        WVector(Cicla, 3) = ""
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

WCapa1.Caption = WVector(1, 4)
WCapa2.Caption = WVector(2, 4)
Wcapa3.Caption = WVector(3, 4)
WCapa4.Caption = WVector(4, 4)
WCapa5.Caption = WVector(5, 4)
WCapa6.Caption = WVector(6, 4)

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
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For A = 0 To 5
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WCounter = WCounter + 1
        
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
                    
            If DBGrid1.Text <> "" Then
                    
                WArticulo = DBGrid1.Text
                    
                DBGrid1.Col = 1
                WDescripcion = DBGrid1.Text
                    
                DBGrid1.Col = 2
                WCantidad = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 3
                WPrecio = Val(DBGrid1.Text)
                
                Rem aca se reemplaza la rutina de cambio envase
    
                WPrimer = DBGrid1.FirstRow
                WFila = DBGrid1.Row
                WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                
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
                        ZArticulo = ""
                        ZDescripcion = "Especificaciones: " + WEspecificaciones
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
                                        
        Next iRow
            
    Next A
    
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
        ZDescripcion = WEspecif(SumaEspe)
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
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase " _
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
        
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
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

End Sub

Private Sub StkPed()

    WTermi = WArticulo.Text
    WArti = Left$(WArticulo.Text, 3) + Right$(WArticulo.Text, 7)
    WStkPedido = 0

    spPedido = "ListaPedidoTerminado " + "'" + WTermi + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
        
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XPed = rstPedido!Cantidad - rstPedido!Facturado
                
                If XPed <> 0 Then
                    If Pedido.Text <> rstPedido!Pedido Then
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
        rstPedido.Close
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
            A% = MsgBox(m$, 0, "Actualizacion de Pedidos")
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
    
    If Val(WEmpresa) = 8 Then
        Call Calcula_Produccion
        Produccion.Caption = Str$(WProduccion)
            Else
        Produccion.Caption = ""
    End If

    Erase WVector
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 10 Then
        Stock1.Caption = "SI"
        Stock2.Caption = "SII"
        Stock3.Caption = "SIII"
        Stock4.Caption = "SIV"
        Stock5.Caption = "SV"
            Else
        Stock1.Caption = "PI"
        Stock2.Caption = "PII"
        Stock3.Caption = "PV"
        Stock4.Caption = "PVI"
        Stock5.Caption = ""
    End If
    
    WStock1.Caption = ""
    WStock2.Caption = ""
    WStock3.Caption = ""
    WStock4.Caption = ""
    WStock5.Caption = ""
    
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
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
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
            
            Stock.Caption = Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption))
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
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1
                    WEmpresa = "0001"
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
                    
                    WEmpresa = "0003"
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
            
                    WEmpresa = "0005"
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
                    
                    WEmpresa = "0006"
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
                    
                    WEmpresa = "0007"
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
                    
                    WEmpresa = "0001"
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
            
            Stock.Caption = Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption))
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
    StkPedido.Caption = Pusing("###,###.##", StkPedido.Caption)
    Produccion.Caption = Pusing("###,###.##", Produccion.Caption)
    Disponible.Caption = Pusing("###,###.##", Disponible.Caption)

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
    WStock1.Visible = False
    WStock2.Visible = False
    WStock3.Visible = False
    WStock4.Visible = False
    WStock5.Visible = False
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
        For dada = CicloPartida + 1 To LugarPartida
            If Partida(CicloPartida, 3) > Partida(dada, 3) Then
                Auxi1 = Partida(CicloPartida, 1)
                Auxi2 = Partida(CicloPartida, 2)
                Auxi3 = Partida(CicloPartida, 3)
                
                Partida(CicloPartida, 1) = Partida(dada, 1)
                Partida(CicloPartida, 2) = Partida(dada, 2)
                Partida(CicloPartida, 3) = Partida(dada, 3)
                
                Partida(dada, 1) = Auxi1
                Partida(dada, 2) = Auxi2
                Partida(dada, 3) = Auxi3
            End If
        Next dada
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

