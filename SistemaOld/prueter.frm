VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPruter 
   Caption         =   "Ingreso de Ensayos de Productos Terminados"
   ClientHeight    =   8370
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   11880
   Visible         =   0   'False
   Begin VB.TextBox VersionLaboII 
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
      Left            =   7680
      TabIndex        =   123
      Text            =   " "
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
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
      Left            =   9360
      TabIndex        =   118
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox ValorNumero1 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   117
      Top             =   360
      Width           =   800
   End
   Begin VB.TextBox ValorNumero2 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   116
      Top             =   840
      Width           =   800
   End
   Begin VB.TextBox ValorNumero3 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   115
      Top             =   1320
      Width           =   800
   End
   Begin VB.TextBox ValorNumero4 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   114
      Top             =   1800
      Width           =   800
   End
   Begin VB.TextBox ValorNumero5 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   113
      Top             =   2280
      Width           =   800
   End
   Begin VB.TextBox ValorNumero6 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   112
      Top             =   2760
      Width           =   800
   End
   Begin VB.TextBox ValorNumero7 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   111
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox ValorNumero8 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   110
      Top             =   3720
      Width           =   800
   End
   Begin VB.TextBox ValorNumero9 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   109
      Top             =   4200
      Width           =   800
   End
   Begin VB.TextBox ValorNumero10 
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
      Left            =   11000
      MaxLength       =   8
      TabIndex        =   108
      Top             =   4680
      Width           =   800
   End
   Begin VB.TextBox VersionLabo 
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
      Left            =   7080
      TabIndex        =   106
      Text            =   " "
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   2160
      TabIndex        =   31
      Top             =   6360
      Visible         =   0   'False
      Width           =   6375
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   300
         Left            =   4800
         TabIndex        =   78
         Top             =   600
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   300
         Left            =   4800
         TabIndex        =   77
         Top             =   240
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1560
         TabIndex        =   74
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Frame Frame3 
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
         Height          =   975
         Left            =   240
         TabIndex        =   70
         Top             =   1080
         Width           =   1695
         Begin VB.OptionButton ImprePantalla 
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
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton ImpreListado 
            Caption         =   "Lisatado"
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
            TabIndex        =   71
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   975
         Left            =   2160
         TabIndex        =   67
         Top             =   1080
         Width           =   1695
         Begin VB.OptionButton Rechazo 
            Caption         =   "Rechazados"
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
            TabIndex        =   69
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Aprobado 
            Caption         =   "Aprobados"
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
            TabIndex        =   68
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   5160
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
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
         TabIndex        =   76
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
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
         TabIndex        =   75
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
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
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   4320
      TabIndex        =   102
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   104
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   840
         TabIndex        =   103
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Ingrese su Password"
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
         Left            =   360
         TabIndex        =   105
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame PantaNumeroPrueba 
      Height          =   855
      Left            =   3240
      TabIndex        =   99
      Top             =   6960
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox NumeroPrueba 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   100
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label72 
         Caption         =   "Numero de Prueba"
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
         TabIndex        =   101
         Top             =   360
         Width           =   1815
      End
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   6840
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   6840
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
      Index           =   3
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   6840
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
      Index           =   4
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox WPantalla 
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
      Height          =   1620
      ItemData        =   "prueter.frx":0000
      Left            =   3720
      List            =   "prueter.frx":0007
      TabIndex        =   84
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   1815
      Left            =   1200
      TabIndex        =   83
      Top             =   6480
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3201
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "prueter.frx":0015
      Left            =   480
      List            =   "prueter.frx":001C
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "Modifica Prueba"
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
      Left            =   10440
      TabIndex        =   82
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Actualiza 
      Caption         =   "Actualiza Datos"
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
      Left            =   8280
      TabIndex        =   81
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Partida 
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
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Impensayo 
      Caption         =   "Impresion Prueba"
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
      Left            =   10440
      TabIndex        =   73
      Top             =   5880
      Width           =   1215
   End
   Begin MSMask.MaskEdBox fecha 
      Height          =   285
      Left            =   3120
      TabIndex        =   46
      Top             =   0
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
   Begin VB.CommandButton CmdAddRechazo 
      Caption         =   "Graba Rechazo"
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
      Left            =   9360
      TabIndex        =   44
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Confecciono 
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
      Top             =   6120
      Width           =   3975
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
      MaxLength       =   100
      TabIndex        =   42
      Text            =   " "
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox Aspecto 
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
      TabIndex        =   41
      Text            =   " "
      Top             =   5640
      Width           =   3975
   End
   Begin VB.TextBox Ensayo 
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
      TabIndex        =   40
      Text            =   " "
      Top             =   5400
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
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
   Begin Crystal.CrystalReport lista 
      Left            =   9960
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPruter.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
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
      Height          =   1500
      Left            =   480
      TabIndex        =   30
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   10320
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox valor9 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox valor8 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox valor7 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox valor6 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox valor5 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox valor4 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Valor3 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox valor2 
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
      Left            =   8780
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      Text            =   " "
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Valor1 
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
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
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
      Left            =   7200
      TabIndex        =   5
      Top             =   5280
      Width           =   975
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
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
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
      Left            =   6120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
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
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdAddlote 
      Caption         =   "Graba   Prueba"
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
      Left            =   9360
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin MSMask.MaskEdBox Vto 
      Height          =   285
      Left            =   10320
      TabIndex        =   120
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox NroRevalida 
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
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   119
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label77 
      Caption         =   "Rev."
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
      TabIndex        =   122
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
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
      TabIndex        =   107
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Std1010 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   98
      Top             =   4920
      Width           =   5350
   End
   Begin VB.Label Std99 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   97
      Top             =   4440
      Width           =   5350
   End
   Begin VB.Label Std88 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   96
      Top             =   3960
      Width           =   5350
   End
   Begin VB.Label Std77 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   95
      Top             =   3480
      Width           =   5350
   End
   Begin VB.Label Std66 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   94
      Top             =   3000
      Width           =   5350
   End
   Begin VB.Label Std55 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   93
      Top             =   2520
      Width           =   5350
   End
   Begin VB.Label Std44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   92
      Top             =   2040
      Width           =   5350
   End
   Begin VB.Label Std33 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   91
      Top             =   1560
      Width           =   5350
   End
   Begin VB.Label Std22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   90
      Top             =   1080
      Width           =   5350
   End
   Begin VB.Label Std11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   89
      Top             =   600
      Width           =   5350
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   79
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Ensayo10 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   66
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Ensayo9 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   65
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Ensayo8 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   64
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Ensayo7 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   63
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Ensayo6 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   62
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Ensayo5 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   61
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Ensayo4 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   60
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Ensayo3 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   59
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Ensayo2 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   58
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Ensayo1 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   57
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   56
      Top             =   4680
      Width           =   5350
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   55
      Top             =   4200
      Width           =   5350
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   54
      Top             =   3720
      Width           =   5350
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   53
      Top             =   3240
      Width           =   5350
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   52
      Top             =   2760
      Width           =   5350
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   51
      Top             =   2280
      Width           =   5350
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   50
      Top             =   1800
      Width           =   5350
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   49
      Top             =   1320
      Width           =   5350
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   840
      Width           =   5350
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   360
      Width           =   5350
   End
   Begin VB.Label Label15 
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
      Left            =   2520
      TabIndex        =   45
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Confecciono"
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
      TabIndex        =   39
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
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
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descri10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   4680
      Width           =   2340
   End
   Begin VB.Label Descri9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   4200
      Width           =   2340
   End
   Begin VB.Label Descri8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   3720
      Width           =   2340
   End
   Begin VB.Label Descri7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   3240
      Width           =   2340
   End
   Begin VB.Label Descri6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2760
      Width           =   2340
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   2340
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1800
      Width           =   2340
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Width           =   2340
   End
   Begin VB.Label descri2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   2340
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   360
      Width           =   2340
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label78 
      Caption         =   "Vto."
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
      TabIndex        =   121
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "PrgPruter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim WProceso As String
Dim ZEnsayo1 As String
Dim ZEnsayo2 As String
Dim ZEnsayo3 As String
Dim ZEnsayo4 As String
Dim ZEnsayo5 As String
Dim ZEnsayo6 As String
Dim ZEnsayo7 As String
Dim ZEnsayo8 As String
Dim ZEnsayo9 As String
Dim ZEnsayo10 As String

Dim ZVersionI As String
Dim ZVersionII As String

Dim ZMesesTerminado As Integer
Dim ZMesesRevalida As Integer
Dim ZMes As String
Dim ZAno As String

Dim EmpresaActual As String

Dim ZCarga(20000, 3) As String
Dim CargaEmpresa(12, 2) As String

Dim ZEnsayo(10) As String
Dim ZDesde(10) As String
Dim ZHasta(10) As String
Dim ZUnidad(10) As String
Dim ZValorNumero(10) As String

Dim ZZSalvaOri(100000) As String

Dim ZZZZVencimiento As String
Dim ZZZZFechaVto As String
Dim XMes As String
Dim XAno As String

Dim WBorra(1000, 10) As String



Private Sub Acepta_Click()

    If Aprobado.Value = True Then
        Desdepru = "1000000"
        HastaPru = "1999999"
            Else
        Desdepru = "2000000"
        HastaPru = "2999999"
    End If
    
    WAno = Right$(Desdefec.Text, 4)
    WMes = Mid$(Desdefec.Text, 4, 2)
    WDia = Left$(Desdefec.Text, 2)
    FDesde = WAno + WMes + WDia
    WAno = Right$(Hastafec.Text, 4)
    WMes = Mid$(Hastafec.Text, 4, 2)
    WDia = Left$(Hastafec.Text, 2)
    FHasta = WAno + WMes + WDia

    Lista.WindowTitle = "Listado de Controles de Materias Primas"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height
    
    Lista.ReportFileName = "WPruter.rpt"
    
    Desde.Text = UCase(Desde.Text)
    
    With rstPrueba
        .Index = "Clave"
        .Seek ">=", "0"
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
    
    If Left$(Desde.Text, 2) = "DW" Then
        WProducto = "DW" + Mid$(Desde.Text, 3, 10)
            Else
        If Left$(Desde.Text, 2) = "SE" Then
            WProducto = "SE" + Mid$(Desde.Text, 3, 10)
                Else
            WProducto = "PT" + Mid$(Desde.Text, 3, 10)
        End If
    End If

    Sql1 = "Select *"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        ZEnsayo1 = rstEspecifUnifica!Ensayo1
        ZEnsayo2 = rstEspecifUnifica!Ensayo2
        ZEnsayo3 = rstEspecifUnifica!Ensayo3
        ZEnsayo4 = rstEspecifUnifica!Ensayo4
        ZEnsayo5 = rstEspecifUnifica!Ensayo5
        ZEnsayo6 = rstEspecifUnifica!Ensayo6
        ZEnsayo7 = rstEspecifUnifica!Ensayo7
        ZEnsayo8 = rstEspecifUnifica!Ensayo8
        ZEnsayo9 = rstEspecifUnifica!Ensayo9
        ZEnsayo10 = rstEspecifUnifica!Ensayo10
        rstEspecifUnifica.Close
    End If
    
    ZDesValor1 = ""
    ZDesValor2 = ""
    ZDesValor3 = ""
    ZDesValor4 = ""
    ZDesValor5 = ""
    ZDesValor6 = ""
    ZDesValor7 = ""
    ZDesValor8 = ""
    ZDesValor9 = ""
    ZDesValor10 = ""
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo1 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor1 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo2 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor2 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo3 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor3 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo4 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor4 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo5 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor5 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo6 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor6 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo7 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor7 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo8 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor8 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo9 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor9 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo10 + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZDesValor10 = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
    Suma = 0
    
    ZSql = ""
    ZSql = "Select PrueTer.Prueba, Prueter.Producto, Prueter.Fecha, Prueter.Valor1,  Prueter.Valor2,  Prueter.Valor3,  Prueter.Valor4,  Prueter.Valor5,  Prueter.Valor6,  Prueter.Valor7,  Prueter.Valor8,  Prueter.Valor9,  Prueter.Valor10, Terminado.Descripcion as [DesProducto]  "
    ZSql = ZSql & " FROM Prueter, Terminado"
    ZSql = ZSql & " Where Prueter.Producto = " + "'" + Desde.Text + "'"
    ZSql = ZSql & " and Prueter.FechaOrd >= " + "'" + FDesde + "'"
    ZSql = ZSql & " and Prueter.FechaOrd <= " + "'" + FHasta + "'"
    ZSql = ZSql & " and Prueter.Producto = Terminado.Codigo"
    
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
    
        With rstPrueter
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Suma = Suma + 1
                    
                    ZPrueba = rstPrueter!Prueba
                    ZProducto = rstPrueter!Producto
                    ZFecha = rstPrueter!Fecha
                    ZValor1 = rstPrueter!Valor1
                    ZValor2 = rstPrueter!valor2
                    ZValor3 = rstPrueter!Valor3
                    ZValor4 = rstPrueter!valor4
                    ZValor5 = rstPrueter!valor5
                    ZValor6 = rstPrueter!valor6
                    ZValor7 = rstPrueter!valor7
                    ZValor8 = rstPrueter!valor8
                    ZValor9 = rstPrueter!valor9
                    ZValor10 = rstPrueter!valor10
                    ZDesProducto = rstPrueter!DesProducto
                    
                    With rstPrueba
                        .AddNew
                        !Clave = Suma
                        !Prueba = ZPrueba
                        !Producto = ZProducto
                        !Fecha = ZFecha
                        !Valor1 = ZValor1
                        !valor2 = ZValor2
                        !Valor3 = ZValor3
                        !valor4 = ZValor4
                        !valor5 = ZValor5
                        !valor6 = ZValor6
                        !valor7 = ZValor7
                        !valor8 = ZValor8
                        !valor9 = ZValor9
                        !valor10 = ZValor10
                        !DesValor1 = ZDesValor1
                        !Desvalor2 = ZDesValor2
                        !DesValor3 = ZDesValor3
                        !Desvalor4 = ZDesValor4
                        !Desvalor5 = ZDesValor5
                        !Desvalor6 = ZDesValor6
                        !Desvalor7 = ZDesValor7
                        !Desvalor8 = ZDesValor8
                        !Desvalor9 = ZDesValor9
                        !Desvalor10 = ZDesValor10
                        !DesProducto = ZDesProducto
                        .Update
                    End With
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstPrueter.Close
        
    End If
    
    Rem Uno = "{Prueter.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Dos = " and {Prueter.Fechaord} in " + Chr$(34) + FDesde + Chr$(34) + " to " + Chr$(34) + FHasta + Chr$(34)
    Rem Tres = " and {Prueter.Prueba} in " + Chr$(34) + Desdepru + Chr$(34) + " to " + Chr$(34) + HastaPru + Chr$(34)
    Rem lista.GroupSelectionFormula = Uno + Dos + Tres

    If ImpreListado.Value = True Then
        Lista.Destination = 1
            Else
        Lista.Destination = 0
    End If
    
    a = Lista.ReportFileName
    
    
    Lista.DataFiles(0) = Wempresa + "auxi.mdb"
    Lista.Connect = Connect()
    
    Lista.Action = 1
    Frame2.Visible = False
End Sub


Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()
Rem by nan 26-05-11 verifico datos no vacios


    If Valor1 = "" And ValorNumero1 <> "" Then
      m$ = "Existe un campo vacio "
       a% = MsgBox(m$, 0, "Verificacion de Datos")
       Exit Sub
    End If
   
    If valor2 = "" And ValorNumero2 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If

    If Valor3 = "" And ValorNumero3 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor4 = "" And ValorNumero4 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor5 = "" And ValorNumero5 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor6 = "" And ValorNumero6 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor7 = "" And ValorNumero7 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor8 = "" And ValorNumero8 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor9 = "" And ValorNumero9 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
    If valor10 = "" And ValorNumero10 <> "" Then
        m$ = "Existe un campo vacio "
        a% = MsgBox(m$, 0, "Verificacion de Datos")
        Exit Sub
    End If
    
Rem fin by nan 26-05-11
    WPasa = "S"
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Partida.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    spHoja = "ListaHoja " + "'" + Partida.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        If rstHoja!Producto <> Producto.Text Then
            m$ = "El Codigo de Producto de la partida no coincide con el informado"
            a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
        End If
        rstHoja.Close
                    Else
        m$ = "Codigo de Partida invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If WPasa = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Prueter"
        ZSql = ZSql + " Where Prueter.Lote = " + "'" + Partida.Text + "'"
        rsPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            m$ = "Prueba ya ingresada"
            a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
            rstPrueter.Close
        End If
    End If

     If WPasa = "S" Then
    
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
    
        Sql1 = "Select *"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
        
            ZDesde(1) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZDesde(2) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZDesde(3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZDesde(4) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZDesde(5) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZDesde(6) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZDesde(7) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZDesde(8) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZDesde(9) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZDesde(10) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            
            ZHasta(1) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZHasta(2) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZHasta(3) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZHasta(4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZHasta(5) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZHasta(6) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZHasta(7) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZHasta(8) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZHasta(9) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZHasta(10) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            ZDesde(1) = Trim(ZDesde(1))
            ZDesde(2) = Trim(ZDesde(2))
            ZDesde(3) = Trim(ZDesde(3))
            ZDesde(4) = Trim(ZDesde(4))
            ZDesde(5) = Trim(ZDesde(5))
            ZDesde(6) = Trim(ZDesde(6))
            ZDesde(7) = Trim(ZDesde(7))
            ZDesde(8) = Trim(ZDesde(8))
            ZDesde(9) = Trim(ZDesde(9))
            ZDesde(10) = Trim(ZDesde(10))
            
            ZHasta(1) = Trim(ZHasta(1))
            ZHasta(2) = Trim(ZHasta(2))
            ZHasta(3) = Trim(ZHasta(3))
            ZHasta(4) = Trim(ZHasta(4))
            ZHasta(5) = Trim(ZHasta(5))
            ZHasta(6) = Trim(ZHasta(6))
            ZHasta(7) = Trim(ZHasta(7))
            ZHasta(8) = Trim(ZHasta(8))
            ZHasta(9) = Trim(ZHasta(9))
            ZHasta(10) = Trim(ZHasta(10))
            
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
            
            rstEspecifUnifica.Close
        End If
        
        Call Conecta_Empresa
    
        ZValorNumero(1) = ValorNumero1.Text
        ZValorNumero(2) = ValorNumero2.Text
        ZValorNumero(3) = ValorNumero3.Text
        ZValorNumero(4) = ValorNumero4.Text
        ZValorNumero(5) = ValorNumero5.Text
        ZValorNumero(6) = ValorNumero6.Text
        ZValorNumero(7) = ValorNumero7.Text
        ZValorNumero(8) = ValorNumero8.Text
        ZValorNumero(9) = ValorNumero9.Text
        ZValorNumero(10) = ValorNumero10.Text
        
        Rem For WWCiclo = 1 To 10
        Rem     If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
        Rem         If Val(ZValorNumero(WWCiclo)) = 0 Then
        Rem             m$ = "No se informado valor de control en una de las ensayos que requiere validacion"
        Rem             A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Rem             Exit Sub
        Rem         End If
        Rem     End If
        Rem Next WWCiclo
        
        
        
        
        
        
        
        
        For WWCiclo = 1 To 10
        
            If ZEnsayo(WWCiclo) <> 0 Then
            
                If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
                
                    ZNumeI = InStr(Trim(ZDesde(WWCiclo)), ".")
                    ZNumeII = Len(Trim(ZDesde(WWCiclo)))
                    If ZNumeI <> 0 Then
                        ZDife = ZNumeII - ZNumeI
                            Else
                        ZDife = 0
                    End If
                    
                    Select Case ZDife
                        Case 1
                            ZValorNumero(WWCiclo) = Pusing("###,###.#", ZValorNumero(WWCiclo))
                        Case 2
                            ZValorNumero(WWCiclo) = Pusing("###,###.##", ZValorNumero(WWCiclo))
                        Case 3
                            ZValorNumero(WWCiclo) = Pusing("###,###.###", ZValorNumero(WWCiclo))
                        Case 4
                            ZValorNumero(WWCiclo) = Pusing("###,###.####", ZValorNumero(WWCiclo))
                        Case 5
                            ZValorNumero(WWCiclo) = Pusing("###,###.#####", ZValorNumero(WWCiclo))
                        Case 6
                            ZValorNumero(WWCiclo) = Pusing("###,###.######", ZValorNumero(WWCiclo))
                        Case Else
                            ZValorNumero(WWCiclo) = Pusing("###,###", ZValorNumero(WWCiclo))
                    End Select
                    
                    Select Case WWCiclo
                        Case 1
                            Valor1.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 2
                            valor2.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 3
                            Valor3.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 4
                            valor4.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 5
                            valor5.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 6
                            valor6.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 7
                            valor7.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 8
                            valor8.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 9
                            valor9.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case 10
                            valor10.Text = ZValorNumero(WWCiclo) + " " + ZUnidad(WWCiclo)
                        Case Else
                    End Select
                
                    If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                        aa = Val(ZValorNumero(WWCiclo))
                        If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Or Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                                
                    If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) = 0 Then
                        If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                            
                    If Val(ZDesde(WWCiclo)) = 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                        If Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                    
                        Else
                        
                    If ZValorNumero(WWCiclo) = "S" Or ZValorNumero(WWCiclo) = "N" Then
                    
                        If ZValorNumero(WWCiclo) = "S" Then
                            ZZZZValor = "Cumple"
                                Else
                            ZZZZValor = "No Cumple"
                        End If
                    
                        Select Case WWCiclo
                            Case 1
                                Valor1.Text = ZZZZValor
                            Case 2
                                valor2.Text = ZZZZValor
                            Case 3
                                Valor3.Text = ZZZZValor
                            Case 4
                                valor4.Text = ZZZZValor
                            Case 5
                                valor5.Text = ZZZZValor
                            Case 6
                                valor6.Text = ZZZZValor
                            Case 7
                                valor7.Text = ZZZZValor
                            Case 8
                                valor8.Text = ZZZZValor
                            Case 9
                                valor9.Text = ZZZZValor
                            Case 10
                                valor10.Text = ZZZZValor
                            Case Else
                        End Select
                        
                    End If
                        
                    If Trim(UCase(ZValorNumero(WWCiclo))) <> "S" Then
                        m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Exit Sub
                    End If
                    
                End If
                
            End If
        
        Next WWCiclo


        ZSql = ""
        ZSql = ZSql + "Select PrueTer.Prueba, PrueTer.Lote"
        ZSql = ZSql + " FROM PrueTer"
        ZSql = ZSql + " Where PrueTer.Prueba <= " + "'" + "1199999" + "'"
        ZSql = ZSql + " Order by PrueTer.Prueba"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            With rstPrueter
                .MoveLast
                Lote = Str$(rstPrueter!Lote + 1)
            End With
            rstPrueter.Close
                Else
            Lote = "1"
        End If



        Rem spPrueter = "ConsultaPrueterMenor " + "'" + "1999999" + "'"
        Rem Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstPrueter.RecordCount > 0 Then
        Rem     Lote = Str$(rstPrueter!Lote + 1)
        Rem     rstPrueter.Close
        Rem         Else
        Rem     Lote = "1"
        Rem End If
    
        If Val(Partida.Text) <> 0 Then
            Auxi1 = Partida.Text
            Call Ceros(Auxi1, 6)
            Lote = Auxi1
                Else
            Auxi1 = Lote
            Call Ceros(Auxi1, 6)
            Lote = Auxi1
        End If
    
        Auxi = "1"
        
        WPrueba = Auxi + Lote
        WProducto = Producto.Text
        WFecha = Fecha.Text
        WValor1 = Valor1.Text
        WValor2 = valor2.Text
        WValor3 = Valor3.Text
        WValor4 = valor4.Text
        WValor5 = valor5.Text
        WValor6 = valor6.Text
        WValor7 = valor7.Text
        WValor8 = valor8.Text
        WValor9 = valor9.Text
        WValor10 = valor10.Text
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = ""
        WLote = Lote
        WRechazo = Lote
        WDate = Date$
        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        XParam = "'" + WPrueba + "','" _
                + WProducto + "','" _
                + WFecha + "','" _
                + WValor1 + "','" _
                + WValor2 + "','" _
                + WValor3 + "','" _
                + WValor4 + "','" _
                + WValor5 + "','" _
                + WValor6 + "','" _
                + WValor7 + "','" _
                + WValor8 + "','" _
                + WValor9 + "','" _
                + WValor10 + "','" _
                + WEnsayo + "','" _
                + WAspecto + "','" _
                + WObservaciones + "','" _
                + WConfecciono + "','" _
                + WLiberada + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WFechaord + "','" _
                + WDate + "'"
        Set rstPrueter = db.OpenRecordset("AltaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueter SET "
        ZSql = ZSql + " ValorNumero1 = " + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + " ValorNumero2 = " + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + " ValorNumero3 = " + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + " ValorNumero4 = " + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + " ValorNumero5 = " + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + " ValorNumero6 = " + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + " ValorNumero7 = " + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + " ValorNumero8 = " + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + " ValorNumero9 = " + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + " ValorNumero10 = " + "'" + ValorNumero10.Text + "'"
        ZSql = ZSql + " Where Prueba = " + "'" + WPrueba + "'"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueter SET "
        ZSql = ZSql + " ValorOriginal1 = " + "'" + WValor1 + "',"
        ZSql = ZSql + " ValorOriginal2 = " + "'" + WValor2 + "',"
        ZSql = ZSql + " ValorOriginal3 = " + "'" + WValor3 + "',"
        ZSql = ZSql + " ValorOriginal4 = " + "'" + WValor4 + "',"
        ZSql = ZSql + " ValorOriginal5 = " + "'" + WValor5 + "',"
        ZSql = ZSql + " ValorOriginal6 = " + "'" + WValor6 + "',"
        ZSql = ZSql + " ValorOriginal7 = " + "'" + WValor7 + "',"
        ZSql = ZSql + " ValorOriginal8 = " + "'" + WValor8 + "',"
        ZSql = ZSql + " ValorOriginal9 = " + "'" + WValor9 + "',"
        ZSql = ZSql + " ValorOriginal10 = " + "'" + WValor10 + "',"
        ZSql = ZSql + " ValorNumeroOriginal1 = " + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal2 = " + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal3 = " + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal4 = " + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal5 = " + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal6 = " + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal7 = " + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal8 = " + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal9 = " + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal10 = " + "'" + ValorNumero10.Text + "'"
        ZSql = ZSql + " Where Prueba = " + "'" + WPrueba + "'"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
        Call CmdLimpiar_Click
        Producto.SetFocus
    
    End If
        
End Sub

Private Sub cmdAddRechazo_Click()
    
    WPasa = "S"
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Partida.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select PrueTer.Prueba, PrueTer.Lote"
    ZSql = ZSql + " FROM PrueTer"
    ZSql = ZSql + " Where PrueTer.Prueba <= " + "'" + "2999999" + "'"
    ZSql = ZSql + " Order by PrueTer.Prueba"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
        With rstPrueter
            .MoveLast
            Lote = Str$(rstPrueter!Lote + 1)
        End With
        rstPrueter.Close
            Else
        Lote = "1"
    End If
    
    Rem spPrueter = "ConsultaPrueterMenor " + "'" + "2999999" + "'"
    Rem Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPrueter.RecordCount > 0 Then
    Rem     Lote = Str$(rstPrueter!Lote + 1)
    Rem     rstPrueter.Close
    Rem         Else
    Rem     Lote = "1"
    Rem End If
    
    Auxi1 = Lote
    Call Ceros(Auxi1, 6)
    Lote = Auxi1
        
    Auxi = "2"
        
    WPrueba = Auxi + Lote
    WProducto = Producto.Text
    WFecha = Fecha.Text
    WValor1 = Valor1.Text
    WValor2 = valor2.Text
    WValor3 = Valor3.Text
    WValor4 = valor4.Text
    WValor5 = valor5.Text
    WValor6 = valor6.Text
    WValor7 = valor7.Text
    WValor8 = valor8.Text
    WValor9 = valor9.Text
    WValor10 = valor10.Text
    WEnsayo = Ensayo.Text
    WAspecto = Aspecto.Text
    WObservaciones = Observaciones.Text
    WConfecciono = Confecciono.Text
    WLote = Lote
    WLiberada = ""
    WRechazo = Lote
    WDate = Date$
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    
    
    XParam = "'" + WPrueba + "','" _
                + WProdcuto + "','" _
                + WFecha + "','" _
                + WValor1 + "','" _
                + WValor2 + "','" _
                + WValor3 + "','" _
                + WValor4 + "','" _
                + WValor5 + "','" _
                + WValor6 + "','" _
                + WValor7 + "','" _
                + WValor8 + "','" _
                + WValor9 + "','" _
                + WValor10 + "','" _
                + WEnsayo + "','" _
                + WAspecto + "','" _
                + WObservaciones + "','" _
                + WConfecciono + "','" _
                + WLiberada + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WFechaord + "','" _
                + WDate + "'"
    Set rstPrueter = db.OpenRecordset("AltaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Prueter SET "
    ZSql = ZSql + " ValorNumero1 = " + "'" + ValorNumero1.Text + "',"
    ZSql = ZSql + " ValorNumero2 = " + "'" + ValorNumero2.Text + "',"
    ZSql = ZSql + " ValorNumero3 = " + "'" + ValorNumero3.Text + "',"
    ZSql = ZSql + " ValorNumero4 = " + "'" + ValorNumero4.Text + "',"
    ZSql = ZSql + " ValorNumero5 = " + "'" + ValorNumero5.Text + "',"
    ZSql = ZSql + " ValorNumero6 = " + "'" + ValorNumero6.Text + "',"
    ZSql = ZSql + " ValorNumero7 = " + "'" + ValorNumero7.Text + "',"
    ZSql = ZSql + " ValorNumero8 = " + "'" + ValorNumero8.Text + "',"
    ZSql = ZSql + " ValorNumero9 = " + "'" + ValorNumero9.Text + "',"
    ZSql = ZSql + " ValorNumero10 = " + "'" + ValorNumero10.Text + "'"
    ZSql = ZSql + " Where Prueba = " + "'" + WPrueba + "'"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
    Call CmdLimpiar_Click
    Producto.SetFocus
        
End Sub

Private Sub CmdLimpiar_Click()

    Producto.Text = "  -     -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Ensayo1.Caption = ""
    Valor1.Text = ""
    Ensayo2.Caption = ""
    valor2.Text = ""
    Ensayo3.Caption = ""
    Valor3.Text = ""
    Ensayo4.Caption = ""
    valor4.Text = ""
    Ensayo5.Caption = ""
    valor5.Text = ""
    Ensayo6.Caption = ""
    valor6.Text = ""
    Ensayo7.Caption = ""
    valor7.Text = ""
    Ensayo8.Caption = ""
    valor8.Text = ""
    Ensayo9.Caption = ""
    valor9.Text = ""
    Ensayo10.Caption = ""
    valor10.Text = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Std1.Caption = ""
    Std2.Caption = ""
    Std3.Caption = ""
    Std4.Caption = ""
    Std5.Caption = ""
    Std6.Caption = ""
    Std7.Caption = ""
    Std8.Caption = ""
    Std9.Caption = ""
    Std10.Caption = ""
    Std11.Caption = ""
    Std22.Caption = ""
    Std33.Caption = ""
    Std44.Caption = ""
    Std55.Caption = ""
    Std66.Caption = ""
    Std77.Caption = ""
    Std88.Caption = ""
    Std99.Caption = ""
    Std1010.Caption = ""
    Partida.Text = ""
    VersionLabo.Text = ""
    VersionLaboII.Text = ""
    ZVersionI = ""
    ZVersionII = ""
    
    ValorNumero1.Text = ""
    ValorNumero2.Text = ""
    ValorNumero3.Text = ""
    ValorNumero4.Text = ""
    ValorNumero5.Text = ""
    ValorNumero6.Text = ""
    ValorNumero7.Text = ""
    ValorNumero8.Text = ""
    ValorNumero9.Text = ""
    ValorNumero10.Text = ""
    
    NroRevalida.Text = ""
    Vto.Text = "  /  /    "
    
    cmdAddlote.Enabled = True
    CmdAddRechazo.Enabled = True
    Actualiza.Enabled = False
    
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgPruter.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command2_Click()



































End Sub

Private Sub Ensayo2_Click()
posicion = 2
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Partida.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub NroRevalida_DblClick()
    If Val(NroRevalida.Text) <> 0 Then
        ZLoteRevalida = Partida.Text
        ZArticuloRevalida = Producto.Text
        ZNroRevalida = NroRevalida.Text
        PrgRevalidaPtConsulta.Show
    End If
End Sub

Private Sub NroRevalida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZLoteRevalida = Partida.Text
        ZArticuloRevalida = Producto.Text
        ZNroRevalida = NroRevalida.Text
        PrgRevalidaPtConsulta.Show
    End If
End Sub

Private Sub Partida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPasa = "S"
        spHoja = "ListaHoja " + "'" + Partida.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            If UCase(rstHoja!Producto) <> UCase(Producto.Text) Then
                m$ = "El Codigo de Producto de la partida no coincide con el informado"
                a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
                WPasa = "N"
            End If
            rstHoja.Close
                        Else
            m$ = "Codigo de Partida invalido"
            a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Prueter"
            ZSql = ZSql + " Where Prueter.Lote = " + "'" + Partida.Text + "'"
            rsPrueter = ZSql
            Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
                m$ = "Prueba ya ingresada"
                a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
                Partida.Text = ""
                WPasa = "N"
                rstPrueter.Close
            End If
        End If
        
        If WPasa = "S" Then
            ValorNumero1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
    OPEN_FILE_Empresa
    OPEN_FILE_PRUEBA
End Sub

Private Sub Impensayo_Click()

    If Val(Auxi) = 0 Then
        Auxi = "0"
    End If
    
    If Val(Lote) = 0 Then
        Lote = "000000"
    End If

    Rem lista.ReportFileName = "Ensayoter.rpt"
    Rem lista.GroupSelectionFormula = "{Prueter.Prueba} in " + Chr$(34) + Auxi + Lote + Chr$(34) + " to " + Chr$(34) + Auxi + Lote + Chr$(34)
    Rem lista.Destination = 1
    Rem lista.Action = 1
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = "12"
    Printer.Print Tab(1); ""
    Printer.FontSize = "10"
    
    Printer.Print Tab(1); "Empresa : " + WAuxiliar
    Printer.Print Tab(1); ""
    Printer.Print Tab(20); "ENSAYO DE PRODUCTO TERMINADO"
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Prueba"; Tab(15); Lote
    Printer.Print Tab(1); "Producto"; Tab(15); Producto.Text
    Printer.Print Tab(1); "Fecha"; Tab(15); Fecha.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo1.Caption; Tab(25); Descri1.Caption; Tab(80); Std1.Caption; Tab(105); Valor1.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo2.Caption; Tab(25); descri2.Caption; Tab(80); Std2.Caption; Tab(105); valor2.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo3.Caption; Tab(25); Descri3.Caption; Tab(80); Std3.Caption; Tab(105); Valor3.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo4.Caption; Tab(25); Descri4.Caption; Tab(80); Std4.Caption; Tab(105); valor4.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo5.Caption; Tab(25); Descri5.Caption; Tab(80); Std5.Caption; Tab(105); valor5.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo6.Caption; Tab(25); Descri6.Caption; Tab(80); Std6.Caption; Tab(105); valor6.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo7.Caption; Tab(25); Descri7.Caption; Tab(80); Std7.Caption; Tab(105); valor7.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo8.Caption; Tab(25); Descri8.Caption; Tab(80); Std8.Caption; Tab(105); valor8.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo9.Caption; Tab(25); Descri9.Caption; Tab(80); Std9.Caption; Tab(105); valor9.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo10.Caption; Tab(25); Descri10.Caption; Tab(80); Std10.Caption; Tab(105); valor10.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Observaciones"; Tab(20); Ensayo.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Aspecto.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Observaciones.Text
    Printer.Print Tab(1); "Confecciono"; Tab(20); Confecciono.Text
    Printer.Print Tab(1); ""
    
    Printer.EndDoc

End Sub

Private Sub Listado_Click()
    Desde.Text = "  -     -   "
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Aprobado.Value = True
    Rechazo.Value = False
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastafec.SetFocus
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub


Private Sub Revalida_Click()
    If Val(Partida.Text) <> 0 Then
            
        WEmpresaRevalida = ""
        ZProgramaOrigen = 0
        ZLoteRevalida = Partida.Text
        ZFechaRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZArticuloRevalida = Producto.Text
        Rem ZDesArticuloRevalida = Descriprod.Caption
        
        ZZRenglon = 0
        ZZTipo = ""
        ZZTerminado = ""
        ZZArticulo = ""
        ZZCantidad = 0
        ZZCantidadLote = 0
        ZZLote = ""
        spHoja = "ListaHoja " + "'" + Partida.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZZRenglon = ZZRenglon + 1
                        ZZTipo = rstHoja!Tipo
                        ZZTerminado = rstHoja!Terminado
                        ZZArticulo = rstHoja!Articulo
                        ZZCantidad = rstHoja!Cantidad
                        ZZCantidadLote = rstHoja!Canti1
                        ZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
        
        Rem veo si es mono
        Rem If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
        Rem     m$ = "Atencion : Este PT es monoproducto y se debe revalida la Materia Prima correspondiente"
        Rem     A% = MsgBox(m$, 64, "Revalidas de Productos Terminados")
        Rem     Exit Sub
        Rem End If
        
        PrgRevalidaPt.Show
        
    End If
End Sub

Private Sub ValorNumero1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(1)), ".")
            ZNumeII = Len(Trim(ZDesde(1)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero1.Text = Pusing("###,###.#", ValorNumero1.Text)
                Case 2
                    ValorNumero1.Text = Pusing("###,###.##", ValorNumero1.Text)
                Case 3
                    ValorNumero1.Text = Pusing("###,###.###", ValorNumero1.Text)
                Case 4
                    ValorNumero1.Text = Pusing("###,###.####", ValorNumero1.Text)
                Case 5
                    ValorNumero1.Text = Pusing("###,###.#####", ValorNumero1.Text)
                Case 6
                    ValorNumero1.Text = Pusing("###,###.######", ValorNumero1.Text)
                Case Else
                    ValorNumero1.Text = Pusing("###,###", ValorNumero1.Text)
            End Select
            
            Valor1.Text = ValorNumero1.Text + " " + ZUnidad(1)
            
            ValorNumero2.SetFocus
            
                Else
                
            If ValorNumero1.Text = "S" Or ValorNumero1.Text = "N" Then
                If ValorNumero1.Text = "S" Then
                    Valor1.Text = "Cumple"
                        Else
                    Valor1.Text = "No Cumple"
                End If
                ValorNumero2.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero1.Text = ""
    End If
    
    If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(2)), ".")
            ZNumeII = Len(Trim(ZDesde(2)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero2.Text = Pusing("###,###.#", ValorNumero2.Text)
                Case 2
                    ValorNumero2.Text = Pusing("###,###.##", ValorNumero2.Text)
                Case 3
                    ValorNumero2.Text = Pusing("###,###.###", ValorNumero2.Text)
                Case 4
                    ValorNumero2.Text = Pusing("###,###.####", ValorNumero2.Text)
                Case 5
                    ValorNumero2.Text = Pusing("###,###.#####", ValorNumero2.Text)
                Case 6
                    ValorNumero2.Text = Pusing("###,###.######", ValorNumero2.Text)
                Case Else
                    ValorNumero2.Text = Pusing("###,###", ValorNumero2.Text)
            End Select
            
            valor2.Text = ValorNumero2.Text + " " + ZUnidad(2)
            
            ValorNumero3.SetFocus
            
                Else
                
            If ValorNumero2.Text = "S" Or ValorNumero2.Text = "N" Then
                If ValorNumero2.Text = "S" Then
                    valor2.Text = "Cumple"
                        Else
                    valor2.Text = "No Cumple"
                End If
                ValorNumero3.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero2.Text = ""
    End If
    
    If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(3)), ".")
            ZNumeII = Len(Trim(ZDesde(3)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero3.Text = Pusing("###,###.#", ValorNumero3.Text)
                Case 2
                    ValorNumero3.Text = Pusing("###,###.##", ValorNumero3.Text)
                Case 3
                    ValorNumero3.Text = Pusing("###,###.###", ValorNumero3.Text)
                Case 4
                    ValorNumero3.Text = Pusing("###,###.####", ValorNumero3.Text)
                Case 5
                    ValorNumero3.Text = Pusing("###,###.#####", ValorNumero3.Text)
                Case 6
                    ValorNumero3.Text = Pusing("###,###.######", ValorNumero3.Text)
                Case Else
                    ValorNumero3.Text = Pusing("###,###", ValorNumero3.Text)
            End Select
            
            Valor3.Text = ValorNumero3.Text + " " + ZUnidad(3)
            
            ValorNumero4.SetFocus
            
                Else
                
            If ValorNumero3.Text = "S" Or ValorNumero3.Text = "N" Then
                If ValorNumero3.Text = "S" Then
                    Valor3.Text = "Cumple"
                        Else
                    Valor3.Text = "No Cumple"
                End If
                ValorNumero4.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero3.Text = ""
    End If
    
    If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(4)), ".")
            ZNumeII = Len(Trim(ZDesde(4)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero4.Text = Pusing("###,###.#", ValorNumero4.Text)
                Case 2
                    ValorNumero4.Text = Pusing("###,###.##", ValorNumero4.Text)
                Case 3
                    ValorNumero4.Text = Pusing("###,###.###", ValorNumero4.Text)
                Case 4
                    ValorNumero4.Text = Pusing("###,###.####", ValorNumero4.Text)
                Case 5
                    ValorNumero4.Text = Pusing("###,###.#####", ValorNumero4.Text)
                Case 6
                    ValorNumero4.Text = Pusing("###,###.######", ValorNumero4.Text)
                Case Else
                    ValorNumero4.Text = Pusing("###,###", ValorNumero4.Text)
            End Select
            
            valor4.Text = ValorNumero4.Text + " " + ZUnidad(4)
            
            ValorNumero5.SetFocus
            
                Else
                
            If ValorNumero4.Text = "S" Or ValorNumero4.Text = "N" Then
                If ValorNumero4.Text = "S" Then
                    valor4.Text = "Cumple"
                        Else
                    valor4.Text = "No Cumple"
                End If
                ValorNumero5.SetFocus
            End If
            
        End If
        
    End If
    
    If KeyAscii = 27 Then
        ValorNumero4.Text = ""
    End If
    
    If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(5)), ".")
            ZNumeII = Len(Trim(ZDesde(5)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero5.Text = Pusing("###,###.#", ValorNumero5.Text)
                Case 2
                    ValorNumero5.Text = Pusing("###,###.##", ValorNumero5.Text)
                Case 3
                    ValorNumero5.Text = Pusing("###,###.###", ValorNumero5.Text)
                Case 4
                    ValorNumero5.Text = Pusing("###,###.####", ValorNumero5.Text)
                Case 5
                    ValorNumero5.Text = Pusing("###,###.#####", ValorNumero5.Text)
                Case 6
                    ValorNumero5.Text = Pusing("###,###.######", ValorNumero5.Text)
                Case Else
                    ValorNumero5.Text = Pusing("###,###", ValorNumero5.Text)
            End Select
            
            valor5.Text = ValorNumero5.Text + " " + ZUnidad(5)
            
            ValorNumero6.SetFocus
            
                Else
                
            If ValorNumero5.Text = "S" Or ValorNumero5.Text = "N" Then
                If ValorNumero5.Text = "S" Then
                    valor5.Text = "Cumple"
                        Else
                    valor5.Text = "No Cumple"
                End If
                ValorNumero6.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero5.Text = ""
    End If
    
    If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(6)), ".")
            ZNumeII = Len(Trim(ZDesde(6)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero6.Text = Pusing("###,###.#", ValorNumero6.Text)
                Case 2
                    ValorNumero6.Text = Pusing("###,###.##", ValorNumero6.Text)
                Case 3
                    ValorNumero6.Text = Pusing("###,###.###", ValorNumero6.Text)
                Case 4
                    ValorNumero6.Text = Pusing("###,###.####", ValorNumero6.Text)
                Case 5
                    ValorNumero6.Text = Pusing("###,###.#####", ValorNumero6.Text)
                Case 6
                    ValorNumero6.Text = Pusing("###,###.######", ValorNumero6.Text)
                Case Else
                    ValorNumero6.Text = Pusing("###,###", ValorNumero6.Text)
            End Select
            
            valor6.Text = ValorNumero6.Text + " " + ZUnidad(6)
            
            ValorNumero7.SetFocus
            
                Else
                
            If ValorNumero6.Text = "S" Or ValorNumero6.Text = "N" Then
                If ValorNumero6.Text = "S" Then
                    valor6.Text = "Cumple"
                        Else
                    valor6.Text = "No Cumple"
                End If
                ValorNumero7.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero6.Text = ""
    End If
    
    If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(7)), ".")
            ZNumeII = Len(Trim(ZDesde(7)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero7.Text = Pusing("###,###.#", ValorNumero7.Text)
                Case 2
                    ValorNumero7.Text = Pusing("###,###.##", ValorNumero7.Text)
                Case 3
                    ValorNumero7.Text = Pusing("###,###.###", ValorNumero7.Text)
                Case 4
                    ValorNumero7.Text = Pusing("###,###.####", ValorNumero7.Text)
                Case 5
                    ValorNumero7.Text = Pusing("###,###.#####", ValorNumero7.Text)
                Case 6
                    ValorNumero7.Text = Pusing("###,###.######", ValorNumero7.Text)
                Case Else
                    ValorNumero7.Text = Pusing("###,###", ValorNumero7.Text)
            End Select
            
            valor7.Text = ValorNumero7.Text + " " + ZUnidad(7)
            
            ValorNumero8.SetFocus
            
                Else
                
            If ValorNumero7.Text = "S" Or ValorNumero7.Text = "N" Then
                If ValorNumero7.Text = "S" Then
                    valor7.Text = "Cumple"
                        Else
                    valor7.Text = "No Cumple"
                End If
                ValorNumero8.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero7.Text = ""
    End If
    
    If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(8)), ".")
            ZNumeII = Len(Trim(ZDesde(8)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero8.Text = Pusing("###,###.#", ValorNumero8.Text)
                Case 2
                    ValorNumero8.Text = Pusing("###,###.##", ValorNumero8.Text)
                Case 3
                    ValorNumero8.Text = Pusing("###,###.###", ValorNumero8.Text)
                Case 4
                    ValorNumero8.Text = Pusing("###,###.####", ValorNumero8.Text)
                Case 5
                    ValorNumero8.Text = Pusing("###,###.#####", ValorNumero8.Text)
                Case 6
                    ValorNumero8.Text = Pusing("###,###.######", ValorNumero8.Text)
                Case Else
                    ValorNumero8.Text = Pusing("###,###", ValorNumero8.Text)
            End Select
            
            valor8.Text = ValorNumero8.Text + " " + ZUnidad(8)
            
            ValorNumero9.SetFocus
            
                Else
                
            If ValorNumero8.Text = "S" Or ValorNumero8.Text = "N" Then
                If ValorNumero8.Text = "S" Then
                    valor8.Text = "Cumple"
                        Else
                    valor8.Text = "No Cumple"
                End If
                ValorNumero9.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero8.Text = ""
    End If
    
    If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(9)), ".")
            ZNumeII = Len(Trim(ZDesde(9)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero9.Text = Pusing("###,###.#", ValorNumero9.Text)
                Case 2
                    ValorNumero9.Text = Pusing("###,###.##", ValorNumero9.Text)
                Case 3
                    ValorNumero9.Text = Pusing("###,###.###", ValorNumero9.Text)
                Case 4
                    ValorNumero9.Text = Pusing("###,###.####", ValorNumero9.Text)
                Case 5
                    ValorNumero9.Text = Pusing("###,###.#####", ValorNumero9.Text)
                Case 6
                    ValorNumero9.Text = Pusing("###,###.######", ValorNumero9.Text)
                Case Else
                    ValorNumero9.Text = Pusing("###,###", ValorNumero9.Text)
            End Select
            
            valor9.Text = ValorNumero9.Text + " " + ZUnidad(9)
            
            ValorNumero10.SetFocus
            
                Else
                
            If ValorNumero9.Text = "S" Or ValorNumero9.Text = "N" Then
                If ValorNumero9.Text = "S" Then
                    valor9.Text = "Cumple"
                        Else
                    valor9.Text = "No Cumple"
                End If
                ValorNumero10.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero9.Text = ""
    End If
    
    If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(10)), ".")
            ZNumeII = Len(Trim(ZDesde(10)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero10.Text = Pusing("###,###.#", ValorNumero10.Text)
                Case 2
                    ValorNumero10.Text = Pusing("###,###.##", ValorNumero10.Text)
                Case 3
                    ValorNumero10.Text = Pusing("###,###.###", ValorNumero10.Text)
                Case 4
                    ValorNumero10.Text = Pusing("###,###.####", ValorNumero10.Text)
                Case 5
                    ValorNumero10.Text = Pusing("###,###.#####", ValorNumero10.Text)
                Case 6
                    ValorNumero10.Text = Pusing("###,###.######", ValorNumero10.Text)
                Case Else
                    ValorNumero10.Text = Pusing("###,###", ValorNumero10.Text)
            End Select
            
            valor10.Text = ValorNumero10.Text + " " + ZUnidad(10)
            
            ValorNumero1.SetFocus
            
                Else
                
            If ValorNumero10.Text = "S" Or ValorNumero10.Text = "N" Then
                If ValorNumero10.Text = "S" Then
                    valor10.Text = "Cumple"
                        Else
                    valor10.Text = "No Cumple"
                End If
                ValorNumero1.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero10.Text = ""
    End If
    
    If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub

















Private Sub Ensayo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Aspecto.SetFocus
    End If
End Sub

Private Sub Aspecto_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Confecciono.SetFocus
    End If
End Sub

Private Sub Confecciono_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.SetFocus
    End If
End Sub

Private Sub imprime_Click()

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

    If Left$(Producto.Text, 2) = "DW" Then
        WProducto = "DW" + Mid$(Producto.Text, 3, 10)
            Else
        If Left$(Producto.Text, 2) = "SE" Then
            WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                Else
            WProducto = "PT" + Mid$(Producto.Text, 3, 10)
        End If
    End If
    
    ZLee = "N"
        
    If Val(ZVersionI) <> Val(ZVersionII) And Val(ZVersionI) <> 0 Then
                
        Sql1 = "Select *"
        Sql2 = " FROM EspecifUnificaVersion"
        Sql3 = " Where EspecifUnificaVersion.Producto = " + "'" + Producto.Text + "'"
        Sql4 = " and EspecifUnificaVersion.Version = " + "'" + ZVersionI + "'"
        spEspecifUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
        Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnificaVersion.RecordCount > 0 Then
                
            ZLee = "S"
            
            Ensayo1.Caption = rstEspecifUnificaVersion!Ensayo1
            Ensayo2.Caption = rstEspecifUnificaVersion!Ensayo2
            Ensayo3.Caption = rstEspecifUnificaVersion!Ensayo3
            Ensayo4.Caption = rstEspecifUnificaVersion!Ensayo4
            Ensayo5.Caption = rstEspecifUnificaVersion!Ensayo5
            Ensayo6.Caption = rstEspecifUnificaVersion!Ensayo6
            Ensayo7.Caption = rstEspecifUnificaVersion!Ensayo7
            Ensayo8.Caption = rstEspecifUnificaVersion!Ensayo8
            Ensayo9.Caption = rstEspecifUnificaVersion!Ensayo9
            Ensayo10.Caption = rstEspecifUnificaVersion!Ensayo10
            
            ZStd1 = rstEspecifUnificaVersion!Valor1
            ZStd2 = rstEspecifUnificaVersion!valor2
            ZStd3 = rstEspecifUnificaVersion!Valor3
            ZStd4 = rstEspecifUnificaVersion!valor4
            ZStd5 = rstEspecifUnificaVersion!valor5
            ZStd6 = rstEspecifUnificaVersion!valor6
            ZStd7 = rstEspecifUnificaVersion!valor7
            ZStd8 = rstEspecifUnificaVersion!valor8
            ZStd9 = rstEspecifUnificaVersion!valor9
            ZStd10 = rstEspecifUnificaVersion!valor10
            
            ZDesde1 = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
            ZDesde2 = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
            ZDesde3 = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
            ZDesde4 = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
            ZDesde5 = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
            ZDesde6 = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
            ZDesde7 = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
            ZDesde8 = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
            ZDesde9 = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
            ZDesde10 = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
            
            ZHasta1 = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
            ZHasta2 = IIf(IsNull(rstEspecifUnificaVersion!Hasta2), "", rstEspecifUnificaVersion!Hasta2)
            ZHasta3 = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
            ZHasta4 = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
            ZHasta5 = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
            ZHasta6 = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
            ZHasta7 = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
            ZHasta8 = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
            ZHasta9 = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
            ZHasta10 = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
            
            Std11.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
            Std22.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
            Std33.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
            Std44.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
            Std55.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
            Std66.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
            Std77.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
            Std88.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
            Std99.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
            Std1010.Caption = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
            
            rstEspecifUnificaVersion.Close
            
                Else
                
            m$ = "Atencion: Se mostraran las especificaciones actuales del producto"
            a% = MsgBox(m$, 64, "Consulta de Ensayos de Productos Terminados")
            ZLee = "N"
            
        End If
        
    End If
        
    If ZLee = "N" Then
        
        Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,Ensayo4,Ensayo5,Ensayo6,Ensayo7,Ensayo8,Ensayo9,Ensayo10,Valor1,Valor2,Valor3,Valor4,Valor5,Valor6,Valor7,Valor8,Valor9,Valor10,Valor11,Valor22,Valor33,Valor44,Valor55,Valor66,Valor77,Valor88,Valor99,Valor1010"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
        
            Ensayo1.Caption = rstEspecifUnifica!Ensayo1
            Ensayo2.Caption = rstEspecifUnifica!Ensayo2
            Ensayo3.Caption = rstEspecifUnifica!Ensayo3
            Ensayo4.Caption = rstEspecifUnifica!Ensayo4
            Ensayo5.Caption = rstEspecifUnifica!Ensayo5
            Ensayo6.Caption = rstEspecifUnifica!Ensayo6
            Ensayo7.Caption = rstEspecifUnifica!Ensayo7
            Ensayo8.Caption = rstEspecifUnifica!Ensayo8
            Ensayo9.Caption = rstEspecifUnifica!Ensayo9
            Ensayo10.Caption = rstEspecifUnifica!Ensayo10
            
            ZStd1 = rstEspecifUnifica!Valor1
            ZStd2 = rstEspecifUnifica!valor2
            ZStd3 = rstEspecifUnifica!Valor3
            ZStd4 = rstEspecifUnifica!valor4
            ZStd5 = rstEspecifUnifica!valor5
            ZStd6 = rstEspecifUnifica!valor6
            ZStd7 = rstEspecifUnifica!valor7
            ZStd8 = rstEspecifUnifica!valor8
            ZStd9 = rstEspecifUnifica!valor9
            ZStd10 = rstEspecifUnifica!valor10
            
            
            
            Std11.Caption = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
            Std22.Caption = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
            Std33.Caption = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
            Std44.Caption = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
            Std55.Caption = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
            Std66.Caption = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
            Std77.Caption = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
            Std88.Caption = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
            Std99.Caption = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
            Std1010.Caption = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
            
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
            
            rstEspecifUnifica.Close
            
        End If
            
            
            
        Sql1 = "Select Desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,Desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            
            ZDesde(1) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZDesde(2) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZDesde(3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZDesde(4) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZDesde(5) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZDesde(6) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZDesde(7) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZDesde(8) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZDesde(9) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZDesde(10) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            
            ZHasta(1) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZHasta(2) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZHasta(3) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZHasta(4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZHasta(5) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZHasta(6) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZHasta(7) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZHasta(8) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZHasta(9) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZHasta(10) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            ZDesde(1) = Trim(ZDesde(1))
            ZDesde(2) = Trim(ZDesde(2))
            ZDesde(3) = Trim(ZDesde(3))
            ZDesde(4) = Trim(ZDesde(4))
            ZDesde(5) = Trim(ZDesde(5))
            ZDesde(6) = Trim(ZDesde(6))
            ZDesde(7) = Trim(ZDesde(7))
            ZDesde(8) = Trim(ZDesde(8))
            ZDesde(9) = Trim(ZDesde(9))
            
            ZHasta(1) = Trim(ZHasta(1))
            ZHasta(2) = Trim(ZHasta(2))
            ZHasta(3) = Trim(ZHasta(3))
            ZHasta(4) = Trim(ZHasta(4))
            ZHasta(5) = Trim(ZHasta(5))
            ZHasta(6) = Trim(ZHasta(6))
            ZHasta(7) = Trim(ZHasta(7))
            ZHasta(8) = Trim(ZHasta(8))
            ZHasta(9) = Trim(ZHasta(9))
            ZHasta(10) = Trim(ZHasta(10))
            
            ZDesde1 = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZDesde2 = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZDesde3 = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZDesde4 = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZDesde5 = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZDesde6 = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZDesde7 = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZDesde8 = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZDesde9 = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZDesde10 = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            
            ZHasta1 = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZHasta2 = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZHasta3 = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZHasta4 = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZHasta5 = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZHasta6 = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZHasta7 = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZHasta8 = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZHasta9 = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZHasta10 = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            rstEspecifUnifica.Close
        End If
        
    End If
    
    
    If Val(Ensayo1.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri1 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo2.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri2 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo3.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri3 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo4.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri4 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo5.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri5 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo6.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri6 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo7.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri7 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo8.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri8 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo9.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri9 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(Ensayo10.Caption) <> 0 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDescri10 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            rstEnsayo.Close
        End If
    End If
    
    If Val(ZDesde1) <> 0 Or Val(ZHasta1) <> 0 Then
        Std1.Caption = Trim(ZDesde1) + " - " + Trim(ZHasta1) + " " + Trim(ZDescri1) + " " + Left$(ZStd1, 50)
            Else
        Std1.Caption = Left$(ZStd1, 50)
    End If
    
    If Val(ZDesde2) <> 0 Or Val(ZHasta2) <> 0 Then
        Std2.Caption = Trim(ZDesde2) + " - " + Trim(ZHasta2) + " " + Trim(ZDescri2) + " " + Left$(ZStd2, 50)
            Else
        Std2.Caption = Left$(ZStd2, 50)
    End If
    
    If Val(ZDesde3) <> 0 Or Val(ZHasta3) <> 0 Then
        Std3.Caption = Trim(ZDesde3) + " - " + Trim(ZHasta3) + " " + Trim(ZDescri3) + " " + Left$(ZStd3, 50)
            Else
        Std3.Caption = Left$(ZStd3, 50)
    End If
    
    If Val(ZDesde4) <> 0 Or Val(ZHasta4) <> 0 Then
        Std4.Caption = Trim(ZDesde4) + " - " + Trim(ZHasta4) + " " + Trim(ZDescri4) + " " + Left$(ZStd4, 50)
            Else
        Std4.Caption = Left$(ZStd4, 50)
    End If
    
    If Val(ZDesde5) <> 0 Or Val(ZHasta5) <> 0 Then
        Std5.Caption = Trim(ZDesde5) + " - " + Trim(ZHasta5) + " " + Trim(ZDescri5) + " " + Left$(ZStd5, 50)
            Else
        Std5.Caption = Left$(ZStd5, 50)
    End If
    
    If Val(ZDesde6) <> 0 Or Val(ZHasta6) <> 0 Then
        Std6.Caption = Trim(ZDesde6) + " - " + Trim(ZHasta6) + " " + Trim(ZDescri6) + " " + Left$(ZStd6, 50)
            Else
        Std6.Caption = Left$(ZStd6, 50)
    End If
    
    If Val(ZDesde7) <> 0 Or Val(ZHasta7) <> 0 Then
        Std7.Caption = Trim(ZDesde7) + " - " + Trim(ZHasta7) + " " + Trim(ZDescri7) + " " + Left$(ZStd7, 50)
            Else
        Std7.Caption = Left$(ZStd7, 50)
    End If
    
    If Val(ZDesde8) <> 0 Or Val(ZHasta8) <> 0 Then
        Std8.Caption = Trim(ZDesde8) + " - " + Trim(ZHasta8) + " " + Trim(ZDescri8) + " " + Left$(ZStd8, 50)
            Else
        Std8.Caption = Left$(ZStd8, 50)
    End If
    
    If Val(ZDesde9) <> 0 Or Val(ZHasta9) <> 0 Then
        Std9.Caption = Trim(ZDesde9) + " - " + Trim(ZHasta9) + " " + Trim(ZDescri9) + " " + Left$(ZStd9, 50)
            Else
        Std9.Caption = Left$(ZStd9, 50)
    End If
    
    If Val(ZDesde10) <> 0 Or Val(ZHasta10) <> 0 Then
        Std10.Caption = Trim(ZDesde10) + " - " + Trim(ZHasta10) + " " + Trim(ZDescri10) + " " + Left$(ZStd10, 50)
            Else
        Std10.Caption = Left$(ZStd10, 50)
    End If
    
    Call ImprimeII_Click
    
    Call Conecta_Empresa

End Sub

Private Sub ImprimeII_Click()

    Erase ZUnidad
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        ZUnidad(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        descri2.Caption = rstEnsayo!Descripcion
        ZUnidad(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        ZUnidad(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        ZUnidad(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        ZUnidad(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        ZUnidad(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        ZUnidad(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        ZUnidad(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        ZUnidad(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        ZUnidad(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If

End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            Producto.Text = UCase(Producto.Text)
            WProducto = Producto.Text
            
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
            
            If Left$(Producto.Text, 2) = "DW" Then
                WProducto = "DW" + Mid$(Producto.Text, 3, 10)
                    Else
                If Left$(Producto.Text, 2) = "SE" Then
                    
                    WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                    
                        Else
                        
                    Rem BY NAN PARA RE
                    If WProducto = "RE-25012-994" Then
                        WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                            Else
                        WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                    End If
                    
                End If
            End If
            
            Rem by nan 9-9-2013
            Sql1 = "Select producto"
            Sql2 = " FROM EspecifUnifica"
            Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
            spEspecifUnifica = Sql1 + Sql2 + Sql3
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                rstEspecifUnifica.Close
                Call Conecta_Empresa
                Call imprime_Click
                    Else
                Call Conecta_Empresa
                CmdLimpiar_Click
                Producto.Text = WProducto
            End If
            
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                rstTerminado.Close
                    Else
                Producto.SetFocus
                Exit Sub
            End If
            
        End If
        Fecha.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    WPantalla.Visible = False
    Muestra.Visible = False
    Pantalla.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Pruebas"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    WPantalla.Visible = False
    Opcion.Visible = False
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        Rem IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                        IngresaItem = rstTerminado!Codigo
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstTerminado!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            
            End If
        
        Case 1
            Call Limpia_Vector
            LugarVector = 0
            spPrueter = "ListaPrueterConsulta"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
            
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueter!Producto <> "" Then
                            If rstPrueter!Producto <> "  -     -   " Then
                                If rstPrueter!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueter!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueter!Lote)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueter!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueter!Fecha
                                    IngresaItem = rstPrueter!Prueba
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
            rstPrueter.Close
            
            End If
        
        Case Else
    End Select
            
    If XIndice = 0 Then
        Pantalla.Visible = True
            Else
        Muestra.Visible = True
    End If

End Sub


Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.Top = 100
    Muestra.Height = 5000
    
    WPantalla.Top = 500
    WPantalla.Height = 3500
    
    Muestra.FixedCols = 1
    Muestra.Cols = 5
    Muestra.FixedRows = 1
    Muestra.Rows = 50000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Tipo"
                Muestra.ColWidth(Ciclo) = 1300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Nro.Prueba"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Producto"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    
    Muestra.Row = 0
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        WTitulo(Ciclo).Text = Muestra.Text
        WTitulo(Ciclo).Left = Muestra.CellLeft + Muestra.Left
        WTitulo(Ciclo).Top = Muestra.CellTop + Muestra.Top
        WTitulo(Ciclo).Width = Muestra.CellWidth
        WTitulo(Ciclo).Height = Muestra.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA muestra
    
    WAncho = 340
    For Ciclo = 0 To Muestra.Cols - 1
        WAncho = WAncho + Muestra.ColWidth(Ciclo)
    Next Ciclo
    Muestra.Width = WAncho

    ' Size the columns.
    Font.Name = Muestra.Font.Name
    Font.Size = Muestra.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
    
End Sub


Private Sub WPantalla_Click()

    If WPantalla.ListIndex <> 0 Then
        Seleccion = WPantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    WPantalla.Visible = False
    WIndice.Clear
    
    Select Case ColumnaOpcion
        Case 0
            Call Limpia_Vector
            LugarVector = 0
            spPrueter = "ListaPrueterConsulta"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
            
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueter!Producto <> "" Then
                            If rstPrueter!Producto <> "  -     -   " Then
                                If rstPrueter!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueter!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueter!Lote)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueter!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueter!Fecha
                                    IngresaItem = rstPrueter!Prueba
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
            rstPrueter.Close
            
            End If
            
        Case 3
            Call Limpia_Vector
            LugarVector = 0
            
            Sql1 = "Select *"
            Sql2 = " FROM Prueter"
            Sql3 = " Where Producto = " + "'" + Seleccion + "'"
            Sql4 = " Order by Producto, Fechaord"
            spPrueter = Sql1 + Sql2 + Sql3 + Sql4
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueter!Producto <> "" Then
                            If rstPrueter!Producto <> "  -     -   " Then
                                If rstPrueter!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueter!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueter!Lote)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueter!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueter!Fecha
                                    IngresaItem = rstPrueter!Prueba
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
            rstPrueter.Close
            Muestra.TopRow = 1
            Muestra.Row = 1
            Muestra.Col = 1
            
            End If
    
        Case 4
            Call Limpia_Vector
            LugarVector = 0
            
            Sql1 = "Select *"
            Sql2 = " FROM Prueter"
            Sql3 = " Where Fecha = " + "'" + Seleccion + "'"
            Sql4 = " Order by Producto, Fechaord"
            spPrueter = Sql1 + Sql2 + Sql3 + Sql4
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueter!Producto <> "" Then
                            If rstPrueter!Producto <> "  -     -   " Then
                                If rstPrueter!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueter!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueter!Lote)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueter!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueter!Fecha
                                    IngresaItem = rstPrueter!Prueba
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
            rstPrueter.Close
            Muestra.TopRow = 1
            Muestra.Row = 1
            Muestra.Col = 1
            
            End If
            
        Case Else
        
    End Select
    
End Sub

Private Sub NumeroPrueba_Keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    
        WIndice.Clear
        PantaNumeroPrueba.Visible = False
        WPantalla.Visible = False
        Call Limpia_Vector
        LugarVector = 0
    
        Sql1 = "Select *"
        Sql2 = " FROM PrueTer"
        Sql3 = " WHERE Lote = " + "'" + NumeroPrueba.Text + "'"
        spPrueter = Sql1 + Sql2 + Sql3
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            If rstPrueter!Producto <> "" Then
                If rstPrueter!Producto <> "  -     -   " Then
                    If rstPrueter!Producto <> Space$(12) Then
                        LugarVector = LugarVector + 1
                        If Left$(rstPrueter!Prueba, 1) = "1" Then
                            Muestra.TextMatrix(LugarVector, 1) = "OK"
                        End If
                        Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueter!Lote)
                        Muestra.TextMatrix(LugarVector, 3) = rstPrueter!Producto
                        Muestra.TextMatrix(LugarVector, 4) = rstPrueter!Fecha
                        IngresaItem = rstPrueter!Prueba
                        WIndice.AddItem IngresaItem
                    End If
                End If
            End If
            
            rstPrueter.Close
            
        End If
        
        Muestra.TopRow = 1
        Muestra.Row = 1
        Muestra.Col = 1
        
    End If
    
End Sub

Private Sub WTitulo_dblClick(Index As Integer)

    If Index = 2 Then
        PantaNumeroPrueba.Height = 855
        PantaNumeroPrueba.Left = 3200
        PantaNumeroPrueba.Top = 2000
        PantaNumeroPrueba.Width = 4095
        PantaNumeroPrueba.Visible = True
        NumeroPrueba.Text = ""
        NumeroPrueba.SetFocus
    End If
    
    If Index = 3 Then
        ColumnaOpcion = 3
        Call Busqueda
    End If
    
    If Index = 4 Then
        ColumnaOpcion = 4
        Call Busqueda
    End If
    
End Sub

Private Sub Busqueda()

    WPantalla.Clear
    Select Case ColumnaOpcion
        Case 3
            WPantalla.AddItem ""
            Sql1 = "Select DISTINCT Producto"
            Sql2 = " FROM Prueter"
            Sql3 = " Order by Producto"
            spPrueter = Sql1 + Sql2 + Sql3
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueter!Producto <> "" Then
                            If rstPrueter!Producto <> "  -     -   " Then
                                If rstPrueter!Producto <> Space$(12) Then
                                    IngresaItem = rstPrueter!Producto
                                    WPantalla.AddItem IngresaItem
                                End If
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueter.Close
            
        Case 4
            WPantalla.AddItem ""
            Sql1 = "Select DISTINCT FechaOrd"
            Sql2 = " FROM Prueter"
            Sql3 = " Order by FechaOrd"
            spPrueter = Sql1 + Sql2 + Sql3
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueter
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Right$(rstPrueter!FechaOrd, 2) + "/" + Mid$(rstPrueter!FechaOrd, 5, 2) + "/" + Left$(rstPrueter!FechaOrd, 4)
                        WPantalla.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueter.Close
            
            
        Case Else
        
    End Select
            
    WPantalla.Visible = True
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    WPantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Clavepro$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Clavepro$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                rstTerminado.Close
                Call imprime_Click
                    Else
                CmdLimpiar_Click
                Producto.Text = "  -     -   "
            End If
            Producto.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Muestra_Click()

    If Muestra.Row <> 0 Then
    
    Muestra.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    Select Case XIndice
        Case 1
            Indice = Muestra.Row - 1
            ClavePrue$ = WIndice.List(Indice)
            spPrueter = "ConsultaPrueter " + "'" + ClavePrue$ + "'"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
                Partida.Text = rstPrueter!Lote
                Producto.Text = rstPrueter!Producto
                Fecha.Text = rstPrueter!Fecha
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                
                Valor1.Text = rstPrueter!Valor1
                valor2.Text = rstPrueter!valor2
                Valor3.Text = rstPrueter!Valor3
                valor4.Text = rstPrueter!valor4
                valor5.Text = rstPrueter!valor5
                valor6.Text = rstPrueter!valor6
                valor7.Text = rstPrueter!valor7
                valor8.Text = rstPrueter!valor8
                valor9.Text = rstPrueter!valor9
                valor10.Text = rstPrueter!valor10
                
                ZZValor1 = IIf(IsNull(rstPrueter!ValorOriginal1), "", rstPrueter!ValorOriginal1)
                ZZValor2 = IIf(IsNull(rstPrueter!ValorOriginal2), "", rstPrueter!ValorOriginal2)
                ZZValor3 = IIf(IsNull(rstPrueter!ValorOriginal3), "", rstPrueter!ValorOriginal3)
                ZZValor4 = IIf(IsNull(rstPrueter!ValorOriginal4), "", rstPrueter!ValorOriginal4)
                ZZValor5 = IIf(IsNull(rstPrueter!ValorOriginal5), "", rstPrueter!ValorOriginal5)
                ZZValor6 = IIf(IsNull(rstPrueter!ValorOriginal6), "", rstPrueter!ValorOriginal6)
                ZZValor7 = IIf(IsNull(rstPrueter!ValorOriginal7), "", rstPrueter!ValorOriginal7)
                ZZValor8 = IIf(IsNull(rstPrueter!ValorOriginal8), "", rstPrueter!ValorOriginal8)
                ZZValor9 = IIf(IsNull(rstPrueter!ValorOriginal9), "", rstPrueter!ValorOriginal9)
                ZZValor10 = IIf(IsNull(rstPrueter!ValorOriginal10), "", rstPrueter!ValorOriginal10)
                
                
                
                ValorNumero1.Text = IIf(IsNull(rstPrueter!ValorNumero1), "", rstPrueter!ValorNumero1)
                ValorNumero2.Text = IIf(IsNull(rstPrueter!ValorNumero2), "", rstPrueter!ValorNumero2)
                ValorNumero3.Text = IIf(IsNull(rstPrueter!ValorNumero3), "", rstPrueter!ValorNumero3)
                ValorNumero4.Text = IIf(IsNull(rstPrueter!ValorNumero4), "", rstPrueter!ValorNumero4)
                ValorNumero5.Text = IIf(IsNull(rstPrueter!ValorNumero5), "", rstPrueter!ValorNumero5)
                ValorNumero6.Text = IIf(IsNull(rstPrueter!ValorNumero6), "", rstPrueter!ValorNumero6)
                ValorNumero7.Text = IIf(IsNull(rstPrueter!ValorNumero7), "", rstPrueter!ValorNumero7)
                ValorNumero8.Text = IIf(IsNull(rstPrueter!ValorNumero8), "", rstPrueter!ValorNumero8)
                ValorNumero9.Text = IIf(IsNull(rstPrueter!ValorNumero9), "", rstPrueter!ValorNumero9)
                ValorNumero10.Text = IIf(IsNull(rstPrueter!ValorNumero10), "", rstPrueter!ValorNumero10)
                
                
                ZZValorNumero1 = IIf(IsNull(rstPrueter!ValorNumeroOriginal1), "", rstPrueter!ValorNumeroOriginal1)
                ZZValorNumero2 = IIf(IsNull(rstPrueter!ValorNumeroOriginal2), "", rstPrueter!ValorNumeroOriginal2)
                ZZValorNumero3 = IIf(IsNull(rstPrueter!ValorNumeroOriginal3), "", rstPrueter!ValorNumeroOriginal3)
                ZZValorNumero4 = IIf(IsNull(rstPrueter!ValorNumeroOriginal4), "", rstPrueter!ValorNumeroOriginal4)
                ZZValorNumero5 = IIf(IsNull(rstPrueter!ValorNumeroOriginal5), "", rstPrueter!ValorNumeroOriginal5)
                ZZValorNumero6 = IIf(IsNull(rstPrueter!ValorNumeroOriginal6), "", rstPrueter!ValorNumeroOriginal6)
                ZZValorNumero7 = IIf(IsNull(rstPrueter!ValorNumeroOriginal7), "", rstPrueter!ValorNumeroOriginal7)
                ZZValorNumero8 = IIf(IsNull(rstPrueter!ValorNumeroOriginal8), "", rstPrueter!ValorNumeroOriginal8)
                ZZValorNumero9 = IIf(IsNull(rstPrueter!ValorNumeroOriginal9), "", rstPrueter!ValorNumeroOriginal9)
                ZZValorNumero10 = IIf(IsNull(rstPrueter!ValorNumeroOriginal10), "", rstPrueter!ValorNumeroOriginal10)
                
                If Trim(ZZValor1) <> "" Then
                    Valor1.Text = ZZValor1
                End If
                If Trim(ZZValor2) <> "" Then
                    valor2.Text = ZZValor2
                End If
                If Trim(ZZValor3) <> "" Then
                    Valor3.Text = ZZValor3
                End If
                If Trim(ZZValor4) <> "" Then
                    valor4.Text = ZZValor4
                End If
                If Trim(ZZValor5) <> "" Then
                    valor5.Text = ZZValor5
                End If
                If Trim(ZZValor6) <> "" Then
                    valor6.Text = ZZValor6
                End If
                If Trim(ZZValor7) <> "" Then
                    valor7.Text = ZZValor7
                End If
                If Trim(ZZValor8) <> "" Then
                    valor8.Text = ZZValor8
                End If
                If Trim(ZZValor9) <> "" Then
                    valor9.Text = ZZValor9
                End If
                If Trim(ZZValor10) <> "" Then
                    valor10.Text = ZZValor10
                End If
                
                
                If Trim(ZZValorNumero1) <> "" Then
                    ValorNumero1.Text = ZZValorNumero1
                End If
                
                
                
                ValorNumero1.Text = Trim(ValorNumero1.Text)
                ValorNumero2.Text = Trim(ValorNumero2.Text)
                ValorNumero3.Text = Trim(ValorNumero3.Text)
                ValorNumero4.Text = Trim(ValorNumero4.Text)
                ValorNumero5.Text = Trim(ValorNumero5.Text)
                ValorNumero6.Text = Trim(ValorNumero6.Text)
                ValorNumero7.Text = Trim(ValorNumero7.Text)
                ValorNumero8.Text = Trim(ValorNumero8.Text)
                ValorNumero9.Text = Trim(ValorNumero9.Text)
                ValorNumero10.Text = Trim(ValorNumero10.Text)
                
                Ensayo.Text = rstPrueter!Ensayo
                Aspecto.Text = rstPrueter!Aspecto
                Observaciones.Text = rstPrueter!Observaciones
                Confecciono.Text = rstPrueter!Confecciono
                Auxi = Left$(rstPrueter!Prueba, 1)
                Lote = rstPrueter!Lote
                
                rstPrueter.Close
                
                ZVersionI = "0"
                ZVersionII = "0"
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Partida.Text + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    ZVersionI = IIf(IsNull(rstHoja!VersionIII), "0", rstHoja!VersionIII)
                    ZFechaHoja = IIf(IsNull(rstHoja!Fecha), "  /  /    ", rstHoja!Fecha)
                    ZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                    WRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                    NroRevalida.Text = Str$(WRevalida)
                    ZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                    If Val(NroRevalida.Text) <> 0 Then
                        WFechaord = Right$(ZFechaRevalida, 4) + Mid$(ZFechaRevalida, 4, 2) + Left$(ZFechaRevalida, 2)
                    End If
                    rstHoja.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Terminado"
                ZSql = ZSql + " Where Terminado.Codigo = " + "'" + Producto.Text + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZMesesTerminado = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                    ZVersionII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
                    rstTerminado.Close
                End If
                
                If ZMesesRevalida > 0 Then
                    WVida = ZMesesRevalida
                    WMes = Val(Mid$(ZFechaRevalida, 4, 2))
                    WAno = Val(Right$(ZFechaRevalida, 4))
                        Else
                    WVida = ZMesesTerminado
                    WMes = Val(Mid$(ZFechaHoja, 4, 2))
                    WAno = Val(Right$(ZFechaHoja, 4))
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
                ZFechaVencimiento = "01/" + ZMes + "/" + ZAno
                
                Rem Call Calcula_Mono
                Rem
                Rem If ZZZZVencimiento <> "" Then
                Rem     WMes = Val(Mid$(ZZZZVencimiento, 4, 2))
                Rem     WAno = Val(Right$(ZZZZVencimiento, 4))
                Rem     ZMes = Str$(WMes)
                Rem     ZAno = Str$(WAno)
                Rem     Call Ceros(ZMes, 2)
                Rem     Call Ceros(ZAno, 4)
                Rem     ZFechaVencimiento = "01/" + ZMes + "/" + ZAno
                Rem End If
                
                Vto.Text = ZFechaVencimiento
                
                If Left$(Producto.Text, 2) = "DW" Then
                    WProducto = "DW" + Mid$(Producto.Text, 3, 10)
                        Else
                    If Left$(Producto.Text, 2) = "SE" Then
                        WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                            Else
                        WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                    End If
                End If
                
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
                
                LlamaImprime = "N"
                
                If Left$(Producto.Text, 2) = "DW" Then
                    WProducto = "DW" + Mid$(Producto.Text, 3, 10)
                        Else
                    If Left$(Producto.Text, 2) = "SE" Then
                        WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                            Else
                        WProducto = "PT" + Mid$(Producto.Text, 3, 10)
                    End If
                End If
                
                ZEnsayo1 = ""
                ZEnsayo2 = ""
                ZEnsayo3 = ""
                ZEnsayo4 = ""
                ZEnsayo5 = ""
                ZEnsayo6 = ""
                ZEnsayo7 = ""
                ZEnsayo8 = ""
                ZEnsayo9 = ""
                ZEnsayo10 = ""
                ZStd1 = ""
                ZStd2 = ""
                ZStd3 = ""
                ZStd4 = ""
                ZStd5 = ""
                ZStd6 = ""
                ZStd7 = ""
                ZStd8 = ""
                ZStd9 = ""
                ZStd10 = ""
                ZStd11 = ""
                ZStd22 = ""
                ZStd33 = ""
                ZStd44 = ""
                ZStd55 = ""
                ZStd66 = ""
                ZStd77 = ""
                ZStd88 = ""
                ZStd99 = ""
                ZStd1010 = ""
                ZDesde1 = ""
                ZHasta1 = ""
                ZDesde2 = ""
                ZHasta2 = ""
                ZDesde3 = ""
                ZHasta3 = ""
                ZDesde4 = ""
                ZHasta4 = ""
                ZDesde5 = ""
                ZHasta5 = ""
                ZDesde6 = ""
                ZHasta6 = ""
                ZDesde7 = ""
                ZHasta7 = ""
                ZDesde8 = ""
                ZHasta8 = ""
                ZDesde9 = ""
                ZHasta9 = ""
                ZDesde10 = ""
                ZHasta10 = ""
                ZVersion = 0
                
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
                                If WDesde <= WFechaord And WHasta > WFechaord Then
                                
                                    ZEnsayo1 = rstEspecifUnificaVersion!Ensayo1
                                    ZEnsayo2 = rstEspecifUnificaVersion!Ensayo2
                                    ZEnsayo3 = rstEspecifUnificaVersion!Ensayo3
                                    ZEnsayo4 = rstEspecifUnificaVersion!Ensayo4
                                    ZEnsayo5 = rstEspecifUnificaVersion!Ensayo5
                                    ZEnsayo6 = rstEspecifUnificaVersion!Ensayo6
                                    ZEnsayo7 = rstEspecifUnificaVersion!Ensayo7
                                    ZEnsayo8 = rstEspecifUnificaVersion!Ensayo8
                                    ZEnsayo9 = rstEspecifUnificaVersion!Ensayo9
                                    ZEnsayo10 = rstEspecifUnificaVersion!Ensayo10
                                    
                                    ZStd1 = rstEspecifUnificaVersion!Valor1
                                    ZStd2 = rstEspecifUnificaVersion!valor2
                                    ZStd3 = rstEspecifUnificaVersion!Valor3
                                    ZStd4 = rstEspecifUnificaVersion!valor4
                                    ZStd5 = rstEspecifUnificaVersion!valor5
                                    ZStd6 = rstEspecifUnificaVersion!valor6
                                    ZStd7 = rstEspecifUnificaVersion!valor7
                                    ZStd8 = rstEspecifUnificaVersion!valor8
                                    ZStd9 = rstEspecifUnificaVersion!valor9
                                    ZStd10 = rstEspecifUnificaVersion!valor10
                                    ZStd11 = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                    ZStd22 = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                    ZStd33 = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                    ZStd44 = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                    ZStd55 = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                    ZStd66 = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                    ZStd77 = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                    ZStd88 = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                    ZStd99 = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                    ZStd1010 = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                    
                                    ZDesde1 = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                    ZDesde2 = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                    ZDesde3 = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                    ZDesde4 = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                    ZDesde5 = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                    ZDesde6 = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                    ZDesde7 = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                    ZDesde8 = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                    ZDesde9 = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                    ZDesde10 = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                    
                                    ZHasta1 = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                    ZHasta2 = IIf(IsNull(rstEspecifUnificaVersion!Hasta2), "", rstEspecifUnificaVersion!Hasta2)
                                    ZHasta3 = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                    ZHasta4 = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                    ZHasta5 = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                    ZHasta6 = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                    ZHasta7 = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                    ZHasta8 = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                    ZHasta9 = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                    ZHasta10 = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                    
                                    ZVersion = rstEspecifUnificaVersion!Version
                                    LlamaImprime = "S"
                                End If
                                
                                If WDesde > WFechaord And LlamaImprime = "N" Then
                                
                                    ZEnsayo1 = rstEspecifUnificaVersion!Ensayo1
                                    ZEnsayo2 = rstEspecifUnificaVersion!Ensayo2
                                    ZEnsayo3 = rstEspecifUnificaVersion!Ensayo3
                                    ZEnsayo4 = rstEspecifUnificaVersion!Ensayo4
                                    ZEnsayo5 = rstEspecifUnificaVersion!Ensayo5
                                    ZEnsayo6 = rstEspecifUnificaVersion!Ensayo6
                                    ZEnsayo7 = rstEspecifUnificaVersion!Ensayo7
                                    ZEnsayo8 = rstEspecifUnificaVersion!Ensayo8
                                    ZEnsayo9 = rstEspecifUnificaVersion!Ensayo9
                                    ZEnsayo10 = rstEspecifUnificaVersion!Ensayo10
                                    
                                    ZStd1 = rstEspecifUnificaVersion!Valor1
                                    ZStd2 = rstEspecifUnificaVersion!valor2
                                    ZStd3 = rstEspecifUnificaVersion!Valor3
                                    ZStd4 = rstEspecifUnificaVersion!valor4
                                    ZStd5 = rstEspecifUnificaVersion!valor5
                                    ZStd6 = rstEspecifUnificaVersion!valor6
                                    ZStd7 = rstEspecifUnificaVersion!valor7
                                    ZStd8 = rstEspecifUnificaVersion!valor8
                                    ZStd9 = rstEspecifUnificaVersion!valor9
                                    ZStd10 = rstEspecifUnificaVersion!valor10
                                    ZStd11 = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                    ZStd22 = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                    ZStd33 = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                    ZStd44 = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                    ZStd55 = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                    ZStd66 = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                    ZStd77 = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                    ZStd88 = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                    ZStd99 = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                    ZStd1010 = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                    
                                    ZDesde1 = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                    ZDesde2 = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                    ZDesde3 = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                    ZDesde4 = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                    ZDesde5 = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                    ZDesde6 = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                    ZDesde7 = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                    ZDesde8 = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                    ZDesde9 = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                    ZDesde10 = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                    
                                    ZHasta1 = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                    ZHasta2 = IIf(IsNull(rstEspecifUnificaVersion!Hasta2), "", rstEspecifUnificaVersion!Hasta2)
                                    ZHasta3 = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                    ZHasta4 = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                    ZHasta5 = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                    ZHasta6 = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                    ZHasta7 = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                    ZHasta8 = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                    ZHasta9 = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                    ZHasta10 = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                    
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
                
                
                Rem BY NAN 16-8-2013
                
                If LlamaImprime = "N" Then
                
                    Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,Ensayo4,Ensayo5,Ensayo6,ensayo7,Ensayo8,Ensayo9,Ensayo10"
                    Sql2 = " FROM EspecifUnifica"
                    Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
                    spEspecifUnifica = Sql1 + Sql2 + Sql3
                    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspecifUnifica.RecordCount > 0 Then
                    
                        ZEnsayo1 = rstEspecifUnifica!Ensayo1
                        ZEnsayo2 = rstEspecifUnifica!Ensayo2
                        ZEnsayo3 = rstEspecifUnifica!Ensayo3
                        ZEnsayo4 = rstEspecifUnifica!Ensayo4
                        ZEnsayo5 = rstEspecifUnifica!Ensayo5
                        ZEnsayo6 = rstEspecifUnifica!Ensayo6
                        ZEnsayo7 = rstEspecifUnifica!Ensayo7
                        ZEnsayo8 = rstEspecifUnifica!Ensayo8
                        ZEnsayo9 = rstEspecifUnifica!Ensayo9
                        ZEnsayo10 = rstEspecifUnifica!Ensayo10
                        rstEspecifUnifica.Close
                        
                        LlamaImprime = "S"
                 
                        Sql1 = "Select Valor1,Valor2,Valor3,Valor4,Valor5,Valor6,Valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010,version"
                        Sql2 = " FROM EspecifUnifica"
                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEspecifUnifica.RecordCount > 0 Then
                        
                            ZStd1 = rstEspecifUnifica!Valor1
                            ZStd2 = rstEspecifUnifica!valor2
                            ZStd3 = rstEspecifUnifica!Valor3
                            ZStd4 = rstEspecifUnifica!valor4
                            ZStd5 = rstEspecifUnifica!valor5
                            ZStd6 = rstEspecifUnifica!valor6
                            ZStd7 = rstEspecifUnifica!valor7
                            ZStd8 = rstEspecifUnifica!valor8
                            ZStd9 = rstEspecifUnifica!valor9
                            ZStd10 = rstEspecifUnifica!valor10
                            
                            ZStd11 = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                            ZStd22 = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                            ZStd33 = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                            ZStd44 = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                            ZStd55 = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                            ZStd66 = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                            ZStd77 = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                            ZStd88 = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                            ZStd99 = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                            ZStd1010 = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                            
                            ZVersion = rstEspecifUnifica!Version
                            rstEspecifUnifica.Close
                        End If
                        
                        Sql1 = "Select Desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,Desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10"
                        Sql2 = " FROM EspecifUnifica"
                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEspecifUnifica.RecordCount > 0 Then
                        
                            ZDesde1 = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                            ZDesde2 = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                            ZDesde3 = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                            ZDesde4 = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                            ZDesde5 = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                            ZDesde6 = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                            ZDesde7 = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                            ZDesde8 = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                            ZDesde9 = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                            ZDesde10 = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                            
                            ZHasta1 = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                            ZHasta2 = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
                            ZHasta3 = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                            ZHasta4 = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                            ZHasta5 = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                            ZHasta6 = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                            ZHasta7 = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                            ZHasta8 = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                            ZHasta9 = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                            ZHasta10 = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                    
                            rstEspecifUnifica.Close
                        End If
                    
                    End If
                    
                End If
                
                If LlamaImprime = "S" Then
                
                    Ensayo1.Caption = ZEnsayo1
                    Ensayo2.Caption = ZEnsayo2
                    Ensayo3.Caption = ZEnsayo3
                    Ensayo4.Caption = ZEnsayo4
                    Ensayo5.Caption = ZEnsayo5
                    Ensayo6.Caption = ZEnsayo6
                    Ensayo7.Caption = ZEnsayo7
                    Ensayo8.Caption = ZEnsayo8
                    Ensayo9.Caption = ZEnsayo9
                    Ensayo10.Caption = ZEnsayo10
                
                    If Val(Ensayo1.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri1 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo2.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri2 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo3.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri3 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo4.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri4 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo5.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri5 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo6.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri6 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo7.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri7 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo8.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri8 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo9.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri9 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(Ensayo10.Caption) <> 0 Then
                        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            ZDescri10 = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                    End If
                    
                    If Val(ZDesde1) <> 0 Or Val(ZHasta1) <> 0 Then
                        Std1.Caption = Trim(ZDesde1) + " - " + Trim(ZHasta1) + " " + Trim(ZDescri1) + " " + Left$(ZStd1, 50)
                            Else
                        Std1.Caption = Left$(ZStd1, 50)
                    End If
                    
                    If Val(ZDesde2) <> 0 Or Val(ZHasta2) <> 0 Then
                        Std2.Caption = Trim(ZDesde2) + " - " + Trim(ZHasta2) + " " + Trim(ZDescri2) + " " + Left$(ZStd2, 50)
                            Else
                        Std2.Caption = Left$(ZStd2, 50)
                    End If
                    
                    If Val(ZDesde3) <> 0 Or Val(ZHasta3) <> 0 Then
                        Std3.Caption = Trim(ZDesde3) + " - " + Trim(ZHasta3) + " " + Trim(ZDescri3) + " " + Left$(ZStd3, 50)
                            Else
                        Std3.Caption = Left$(ZStd3, 50)
                    End If
                    
                    If Val(ZDesde4) <> 0 Or Val(ZHasta4) <> 0 Then
                        Std4.Caption = Trim(ZDesde4) + " - " + Trim(ZHasta4) + " " + Trim(ZDescri4) + " " + Left$(ZStd4, 50)
                            Else
                        Std4.Caption = Left$(ZStd4, 50)
                    End If
                    
                    If Val(ZDesde5) <> 0 Or Val(ZHasta5) <> 0 Then
                        Std5.Caption = Trim(ZDesde5) + " - " + Trim(ZHasta5) + " " + Trim(ZDescri5) + " " + Left$(ZStd5, 50)
                            Else
                        Std5.Caption = Left$(ZStd5, 50)
                    End If
                    
                    If Val(ZDesde6) <> 0 Or Val(ZHasta6) <> 0 Then
                        Std6.Caption = Trim(ZDesde6) + " - " + Trim(ZHasta6) + " " + Trim(ZDescri6) + " " + Left$(ZStd6, 50)
                            Else
                        Std6.Caption = Left$(ZStd6, 50)
                    End If
                    
                    If Val(ZDesde7) <> 0 Or Val(ZHasta7) <> 0 Then
                        Std7.Caption = Trim(ZDesde7) + " - " + Trim(ZHasta7) + " " + Trim(ZDescri7) + " " + Left$(ZStd7, 50)
                            Else
                        Std7.Caption = Left$(ZStd7, 50)
                    End If
                    
                    If Val(ZDesde8) <> 0 Or Val(ZHasta8) <> 0 Then
                        Std8.Caption = Trim(ZDesde8) + " - " + Trim(ZHasta8) + " " + Trim(ZDescri8) + " " + Left$(ZStd8, 50)
                            Else
                        Std8.Caption = Left$(ZStd8, 50)
                    End If
                    
                    If Val(ZDesde9) <> 0 Or Val(ZHasta9) <> 0 Then
                        Std9.Caption = Trim(ZDesde9) + " - " + Trim(ZHasta9) + " " + Trim(ZDescri9) + " " + Left$(ZStd9, 50)
                            Else
                        Std9.Caption = Left$(ZStd9, 50)
                    End If
                    
                    If Val(ZDesde10) <> 0 Or Val(ZHasta10) <> 0 Then
                        Std10.Caption = Trim(ZDesde10) + " - " + Trim(ZHasta10) + " " + Trim(ZDescri10) + " " + Left$(ZStd10, 50)
                            Else
                        Std10.Caption = Left$(ZStd10, 50)
                    End If
                                
                    Std11.Caption = ZStd11
                    Std22.Caption = ZStd22
                    Std33.Caption = ZStd33
                    Std44.Caption = ZStd44
                    Std55.Caption = ZStd55
                    Std66.Caption = ZStd66
                    Std77.Caption = ZStd77
                    Std88.Caption = ZStd88
                    Std99.Caption = ZStd99
                    Std1010.Caption = ZStd1010
                    
                    Call ImprimeII_Click
                    
                End If
                        
                Call Conecta_Empresa
                
                VersionLabo.Text = ZVersion
                VersionLaboII.Text = ZVersionII
                
                cmdAddlote.Enabled = False
                CmdAddRechazo.Enabled = False
                Actualiza.Enabled = True
                    
                    Else
                    
                Call CmdLimpiar_Click
                
            End If
            Producto.SetFocus
        
        Case Else
    End Select
    
    End If

End Sub

Private Sub Form_Load()

    Producto.Text = "  -     -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Ensayo1.Caption = ""
    Valor1.Text = ""
    Ensayo2.Caption = ""
    valor2.Text = ""
    Ensayo3.Caption = ""
    Valor3.Text = ""
    Ensayo4.Caption = ""
    valor4.Text = ""
    Ensayo5.Caption = ""
    valor5.Text = ""
    Ensayo6.Caption = ""
    valor6.Text = ""
    Ensayo7.Caption = ""
    valor7.Text = ""
    Ensayo8.Caption = ""
    valor8.Text = ""
    Ensayo9.Caption = ""
    valor9.Text = ""
    Ensayo10.Caption = ""
    valor10.Text = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Std1.Caption = ""
    Std2.Caption = ""
    Std3.Caption = ""
    Std4.Caption = ""
    Std5.Caption = ""
    Std6.Caption = ""
    Std7.Caption = ""
    Std8.Caption = ""
    Std9.Caption = ""
    Std10.Caption = ""
    Std11.Caption = ""
    Std22.Caption = ""
    Std33.Caption = ""
    Std44.Caption = ""
    Std55.Caption = ""
    Std66.Caption = ""
    Std77.Caption = ""
    Std88.Caption = ""
    Std99.Caption = ""
    Std1010.Caption = ""
    Partida.Text = ""
    VersionLabo.Text = ""
    VersionLaboII.Text = ""
    ZVersionI = ""
    ZVersionII = ""
    
    ValorNumero1.Text = ""
    ValorNumero2.Text = ""
    ValorNumero3.Text = ""
    ValorNumero4.Text = ""
    ValorNumero5.Text = ""
    ValorNumero6.Text = ""
    ValorNumero7.Text = ""
    ValorNumero8.Text = ""
    ValorNumero9.Text = ""
    ValorNumero10.Text = ""
    
    NroRevalida.Text = ""
    Vto.Text = "  /  /    "

    cmdAddlote.Enabled = True
    CmdAddRechazo.Enabled = True
    Actualiza.Enabled = False

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgPruter.Caption = "Ingreso de Ensayos de Producto Terminado :  " + !Nombre
        End If
    End With
    EmpresaActual = Wempresa
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    VersionLabo.Text = ""
    VersionLaboII.Text = ""
    
    Partida.Text = ""
    Valor1.Text = ""
    
End Sub

Private Sub Modifica_Click()
    WProceso = 0
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub Actualiza_Click()
    WProceso = 1
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case WProceso
            Case 0
                If WClave.Text = "SEGURO" Then
                    Pass.Visible = False
                    Call ModificaPrueba
                End If
            Case Else
                If WClave.Text = "SEGURO" Then
                    Pass.Visible = False
                    Call ActualizaPrueba
                End If
        End Select
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub ModificaPrueba()

    ZSql = ""
    ZSql = ZSql + "UPDATE Prueter SET "
    ZSql = ZSql + " Ensayo = " + "'" + Ensayo.Text + "',"
    ZSql = ZSql + " Aspecto = " + "'" + Aspecto.Text + "',"
    ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
    ZSql = ZSql + " Confecciono = " + "'" + Confecciono.Text + "'"
    ZSql = ZSql + " Where Lote = " + "'" + Partida.Text + "'"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
    Call CmdLimpiar_Click
    Producto.SetFocus
    
End Sub

Private Sub ActualizaPrueba()

    WPasa = "N"
    WTipo = ""
    ZPrueba = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Prueter"
    ZSql = ZSql + " Where Prueter.Lote = " + "'" + Partida.Text + "'"
    rsPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
        WTipo = Left$(rstPrueter!Prueba, 1)
        WPasa = "S"
        ZPrueba = rstPrueter!Prueba
        rstPrueter.Close
    End If

    If WPasa = "N" Then
        m$ = "Prueba no ingresada"
        a% = MsgBox(m$, 0, "Actualizacion de Pruebas de Prodcuto Terminado")
            Else
        WPrueba = ZPrueba
        WValor1 = Valor1.Text
        WValor2 = valor2.Text
        WValor3 = Valor3.Text
        WValor4 = valor4.Text
        WValor5 = valor5.Text
        WValor6 = valor6.Text
        WValor7 = valor7.Text
        WValor8 = valor8.Text
        WValor9 = valor9.Text
        WValor10 = valor10.Text
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WDate = Date$
        
        XParam = "'" + WPrueba + "','" _
                + WValor1 + "','" _
                + WValor2 + "','" _
                + WValor3 + "','" _
                + WValor4 + "','" _
                + WValor5 + "','" _
                + WValor6 + "','" _
                + WValor7 + "','" _
                + WValor8 + "','" _
                + WValor9 + "','" _
                + WValor10 + "','" _
                + WEnsayo + "','" _
                + WAspecto + "','" _
                + WObservaciones + "','" _
                + WConfecciono + "','" _
                + WDate + "'"
        Set rstPrueter = db.OpenRecordset("ModificaPrueterValores " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    
        Call CmdLimpiar_Click
        Producto.SetFocus
    
    End If

End Sub






Private Sub Command1_Click()

    Erase ZZSalvaOri
    ZZLugar = 0

    spPrueter = "ListaPrueter"
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
        With rstPrueter
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Rem ZZPasa = IIf(IsNull(rstPrueter!ValorOriginal1), "S", "N")
                    Rem If ZZPasa = "S" Then
                    
                    Rem ZZValor4 = "34"
                    Rem ZZValor5 = "54"
                    
                    If Trim(rstPrueter!ValorOriginal4) = "34" And Trim(rstPrueter!ValorOriginal5) = "54" Then
                        ZZLugar = ZZLugar + 1
                        ZZSalvaOri(ZZLugar) = rstPrueter!Prueba
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrueter.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        spPrueter = "ConsultaPrueter " + "'" + ZZSalvaOri(Ciclo) + "'"
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            
            ZZValor1 = rstPrueter!Valor1
            ZZValor2 = rstPrueter!valor2
            ZZValor3 = rstPrueter!Valor3
            ZZValor4 = rstPrueter!valor4
            ZZValor5 = rstPrueter!valor5
            ZZValor6 = rstPrueter!valor6
            ZZValor7 = rstPrueter!valor7
            ZZValor8 = rstPrueter!valor8
            ZZValor9 = rstPrueter!valor9
            ZZValor10 = rstPrueter!valor10
            
            ZZValorNumero1 = IIf(IsNull(rstPrueter!ValorNumero1), "", rstPrueter!ValorNumero1)
            ZZValorNumero2 = IIf(IsNull(rstPrueter!ValorNumero2), "", rstPrueter!ValorNumero2)
            ZZValorNumero3 = IIf(IsNull(rstPrueter!ValorNumero3), "", rstPrueter!ValorNumero3)
            ZZValorNumero4 = IIf(IsNull(rstPrueter!ValorNumero4), "", rstPrueter!ValorNumero4)
            ZZValorNumero5 = IIf(IsNull(rstPrueter!ValorNumero5), "", rstPrueter!ValorNumero5)
            ZZValorNumero6 = IIf(IsNull(rstPrueter!ValorNumero6), "", rstPrueter!ValorNumero6)
            ZZValorNumero7 = IIf(IsNull(rstPrueter!ValorNumero7), "", rstPrueter!ValorNumero7)
            ZZValorNumero8 = IIf(IsNull(rstPrueter!ValorNumero8), "", rstPrueter!ValorNumero8)
            ZZValorNumero9 = IIf(IsNull(rstPrueter!ValorNumero9), "", rstPrueter!ValorNumero9)
            ZZValorNumero10 = IIf(IsNull(rstPrueter!ValorNumero10), "", rstPrueter!ValorNumero10)
                
            rstPrueter.Close
            
            Rem ZZValor4 = "80C 35 32C 38"
            
            Rem ZZValor2 = "26"
            Rem ZZValor3 = "24"
            Rem ZZValor4 = "34"
            Rem ZZValor5 = "54"
                
            ZSql = ""
            ZSql = ZSql + "UPDATE Prueter SET "
            ZSql = ZSql + " ValorOriginal1 = " + "'" + ZZValor1 + "',"
            ZSql = ZSql + " ValorOriginal2 = " + "'" + ZZValor2 + "',"
            ZSql = ZSql + " ValorOriginal3 = " + "'" + ZZValor3 + "',"
            ZSql = ZSql + " ValorOriginal4 = " + "'" + ZZValor4 + "',"
            ZSql = ZSql + " ValorOriginal5 = " + "'" + ZZValor5 + "',"
            ZSql = ZSql + " ValorOriginal6 = " + "'" + ZZValor6 + "',"
            ZSql = ZSql + " ValorOriginal7 = " + "'" + ZZValor7 + "',"
            ZSql = ZSql + " ValorOriginal8 = " + "'" + ZZValor8 + "',"
            ZSql = ZSql + " ValorOriginal9 = " + "'" + ZZValor9 + "',"
            ZSql = ZSql + " ValorOriginal10 = " + "'" + ZZValor10 + "',"
            ZSql = ZSql + " ValorNumeroOriginal1 = " + "'" + ZZValorNumero1 + "',"
            ZSql = ZSql + " ValorNumeroOriginal2 = " + "'" + ZZValorNumero2 + "',"
            ZSql = ZSql + " ValorNumeroOriginal3 = " + "'" + ZZValorNumero3 + "',"
            ZSql = ZSql + " ValorNumeroOriginal4 = " + "'" + ZZValorNumero4 + "',"
            ZSql = ZSql + " ValorNumeroOriginal5 = " + "'" + ZZValorNumero5 + "',"
            ZSql = ZSql + " ValorNumeroOriginal6 = " + "'" + ZZValorNumero6 + "',"
            ZSql = ZSql + " ValorNumeroOriginal7 = " + "'" + ZZValorNumero7 + "',"
            ZSql = ZSql + " ValorNumeroOriginal8 = " + "'" + ZZValorNumero8 + "',"
            ZSql = ZSql + " ValorNumeroOriginal9 = " + "'" + ZZValorNumero9 + "',"
            ZSql = ZSql + " ValorNumeroOriginal10 = " + "'" + ZZValorNumero10 + "'"
            ZSql = ZSql + " Where Prueba = " + "'" + ZZSalvaOri(Ciclo) + "'"
            spPrueter = ZSql
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next Ciclo
    

End Sub


Private Sub Calcula_Mono()

    
    ZZZZVencimiento = ""
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
    
        Wempresa = CargaEmpresa(ZCiclo, 1)
        txtOdbc = CargaEmpresa(ZCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        spHoja = "ListaHoja " + "'" + Partida.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            rstHoja.Close
            
            ZZZZRenglon = 0
            ZZZZCantidad = 0
            ZZZZCantidadLote = 0
            ZZZZTipo = ""
            ZZZZLote = ""
            ZZZZTerminado = ""
            ZZZZArticulo = ""
            
            spHoja = "ListaHoja " + "'" + Partida.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                With rstHoja
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZZZRenglon = ZZZZRenglon + 1
                            ZZZZCantidad = rstHoja!Cantidad
                            ZZZZCantidadLote = rstHoja!Canti1
                            ZZZZLote = IIf(IsNull(rstHoja!lote1), 0, rstHoja!lote1)
                            ZZZZTipo = rstHoja!Tipo
                            ZZZZTerminado = rstHoja!Terminado
                            ZZZZArticulo = rstHoja!Articulo
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstHoja.Close
            End If
            
        End If
        
    Next ZCiclo
    
    Call Conecta_Empresa
    
    If ZZZZRenglon = 1 And ZZZZCantidad = ZZZZCantidadLote And ZZZZTipo = "M" Then
             
         ZZZZVto = ""
         ZZZZLaudo = ZZZZLote
         ZZZZFecha = ""
         ZZZZFechaVto = ""
     
         For ZCiclo = 1 To ZHasta1
         
             Wempresa = CargaEmpresa(ZCiclo, 1)
             txtOdbc = CargaEmpresa(ZCiclo, 2)
             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
             
            ZSql = ""
             ZSql = ZSql + "Select *"
             ZSql = ZSql + " FROM Laudo"
             ZSql = ZSql + " Where Laudo = " + "'" + Str$(ZZZZLaudo) + "'"
             ZSql = ZSql + " and Articulo = " + "'" + ZZZZArticulo + "'"
             spLaudo = ZSql
             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
             If rstLaudo.RecordCount > 0 Then
                 ZZZZFecha = rstLaudo!Fecha
                 ZZZZFechaVto = IIf(IsNull(rstLaudo!fechavencimiento), "", rstLaudo!fechavencimiento)
                 rstLaudo.Close
                 Exit For
             End If
                 
         Next ZCiclo
                 
         Call Conecta_Empresa
         
         ZZZZVto = ""
         ZZZZOrdFecha = Right$(ZZZZFecha, 4) + Mid$(ZZZZFecha, 4, 2) + Left$(ZZZZFecha, 2)
         If ZZZZFechaVto <> "" And ZZZZFechaVto <> "  /  /    " And ZZZZFechaVto <> "00/00/0000" Then
             Call Valida_fecha(ZZZZFechaVto, Auxi)
             If Auxi = "S" Then
                 ZZZZVto = ZZZZFechaVto
             End If
        End If
             
        If ZZZZVto = "" Then
             
             ZZZZMeses = 0
             ZSql = ""
             ZSql = ZSql + "Select *"
             ZSql = ZSql + " FROM Articulo"
             ZSql = ZSql + " Where Codigo = " + "'" + ZZZZArticulo + "'"
             spArticulo = ZSql
             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
             If rstArticulo.RecordCount > 0 Then
                 ZZZZMeses = rstArticulo!Meses
                 rstArticulo.Close
             End If
             
             WMes = Val(Mid$(ZZZZFecha, 4, 2))
             WAno = Val(Right$(ZZZZFecha, 4))
             For ZCiclo = 1 To ZZZZMeses
                 WMes = WMes + 1
                 If WMes > 12 Then
                     WAno = WAno + 1
                     WMes = 1
                 End If
             Next ZCiclo
             
             XMes = Str$(WMes)
             XAno = Str$(WAno)
             Call Ceros(XMes, 2)
             Call Ceros(XAno, 4)
             If Val(Left$(ZZZZFecha, 2)) <= 30 Then
                 If Val(XMes) = 2 And Val(Left$(ZZZZFecha, 2)) > 28 Then
                     ZZZZVto = "28/" + XMes + "/" + XAno
                         Else
                     ZZZZVto = Left$(ZZZZFecha, 3) + XMes + "/" + XAno
                 End If
                     Else
                 If Val(XMes) = 2 Then
                     ZZZZVto = "28/" + XMes + "/" + XAno
                         Else
                     ZZZZVto = "30/" + XMes + "/" + XAno
                 End If
             End If
             
        End If
         
        ZZZZVencimiento = ZZZZVto
        
    End If

End Sub

