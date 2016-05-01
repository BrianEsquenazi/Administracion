VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPruterFarma 
   Caption         =   "Ingreso de Ensayos de Productos Terminados de Farma"
   ClientHeight    =   8340
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   Begin VB.CommandButton Confiirmacion 
      Caption         =   "Confirmacion Direccion Tecnica"
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
      Left            =   9600
      TabIndex        =   71
      Top             =   7680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TipoOri 
      BeginProperty Font 
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
      MaxLength       =   2
      TabIndex        =   70
      Text            =   " "
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox VersionI 
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
      Left            =   8400
      TabIndex        =   67
      Text            =   " "
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox VersionII 
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
      Left            =   11160
      TabIndex        =   66
      Text            =   " "
      Top             =   0
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   9600
      TabIndex        =   60
      Top             =   5880
      Width           =   1935
      Begin VB.CommandButton ConfirmaResultado 
         Caption         =   "Confirma"
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
         Left            =   360
         TabIndex        =   62
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Resultado 
         BeginProperty Font 
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
         MaxLength       =   2
         TabIndex        =   61
         Text            =   " "
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Producto"
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
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
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
      TabIndex        =   58
      Top             =   1260
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   57
      Top             =   1200
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
      Left            =   3000
      TabIndex        =   56
      Top             =   1260
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   6375
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   300
         Left            =   4800
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   33
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
         TabIndex        =   30
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
            TabIndex        =   32
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
            TabIndex        =   31
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
         TabIndex        =   27
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
            TabIndex        =   29
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
            TabIndex        =   28
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   4320
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   53
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   840
         TabIndex        =   52
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
         TabIndex        =   54
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame PantaNumeroPrueba 
      Height          =   855
      Left            =   3240
      TabIndex        =   48
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
         MaxLength       =   6
         TabIndex        =   49
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
         TabIndex        =   50
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      ItemData        =   "prueterfarma.frx":0000
      Left            =   3720
      List            =   "prueterfarma.frx":0007
      TabIndex        =   43
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   1815
      Left            =   1200
      TabIndex        =   42
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
      ItemData        =   "prueterfarma.frx":0015
      Left            =   480
      List            =   "prueterfarma.frx":001C
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
      Left            =   7200
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Actualiza 
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
      Left            =   8280
      TabIndex        =   40
      Top             =   5880
      Width           =   1095
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
      Left            =   5400
      TabIndex        =   39
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   26
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
      Left            =   8400
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      Left            =   10800
      Top             =   7680
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
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   10320
      TabIndex        =   9
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
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
      Left            =   8280
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3720
      TabIndex        =   55
      Top             =   1260
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
      Height          =   4815
      Left            =   240
      TabIndex        =   59
      Top             =   360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.CommandButton Registro 
      Caption         =   "Registro de Produccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   68
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Original"
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
      Left            =   9600
      TabIndex        =   69
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label ImpreVersionII 
      Caption         =   "Version Especif Actual"
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
      Left            =   9120
      TabIndex        =   65
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label ImpreVersionI 
      Caption         =   "Version Especif Orig."
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
      Left            =   6480
      TabIndex        =   64
      Top             =   0
      Width           =   1815
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
      Left            =   4680
      TabIndex        =   38
      Top             =   0
      Width           =   1095
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
      TabIndex        =   25
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgPruterFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrueterFarma As Recordset
Dim spPrueterFarma As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCargaV As Recordset
Dim spCargaV As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim ColumnaOpcion As Integer
Dim Seleccion As String
Dim WProceso As String
Dim WPartida As String

Dim EmpresaActual As String
Dim ZPasaEnsayo(5000, 2) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Acepta_Click()

    Erase ZPasaEnsayo
    ZLugar = 0

    XEmpresa = Wempresa
    Wempresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ensayos"
    spEnsayos = ZSql
    Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayos.RecordCount > 0 Then
        With rstEnsayos
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZLugar = ZLugar + 1
                    ZPasaEnsayo(ZLugar, 1) = rstEnsayos!Codigo
                    ZPasaEnsayo(ZLugar, 2) = rstEnsayos!Descripcion
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstEnsayos.Close
    End If
    
    Call Conecta_Empresa
    
    Rem dada
    
    For Ciclo = 1 To ZLugar
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ensayos"
        ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + ZPasaEnsayo(Ciclo, 1) + "'"
        spEnsayos = ZSql
        Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayos.RecordCount > 0 Then
            rstEnsayos.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Ensayos SET "
            ZSql = ZSql + "Descripcion = " + "'" + ZPasaEnsayo(Ciclo, 2) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + ZPasaEnsayo(Ciclo, 1) + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Ensayos ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZPasaEnsayo(Ciclo, 1) + "',"
            ZSql = ZSql + "'" + ZPasaEnsayo(Ciclo, 2) + "')"
                
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
    Next Ciclo


    If Aprobado.Value = True Then
        Desdepru = "100000"
        HastaPru = "199999"
            Else
        Desdepru = "200000"
        HastaPru = "299999"
    End If
    
    WAno = Right$(Desdefec.Text, 4)
    WMes = Mid$(Desdefec.Text, 4, 2)
    WDia = Left$(Desdefec.Text, 2)
    FDesde = WAno + WMes + WDia
    WAno = Right$(Hastafec.Text, 4)
    WMes = Mid$(Hastafec.Text, 4, 2)
    WDia = Left$(Hastafec.Text, 2)
    FHasta = WAno + WMes + WDia

    Lista.WindowTitle = "Listado de Controles de Producto Terminado"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height
    
    Lista.ReportFileName = "WPrueTerFarma.rpt"
    
    If Aprobado.Value = True Then
        WTipo = "1"
            Else
        WTipo = "2"
    End If
    
    Desde.Text = UCase(Desde.Text)
    
    Uno = "{PrueTerFarma.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Desde.Text + Chr$(34)
    Dos = " and {PrueTerFarma.FechaOrd} in " + Chr$(34) + FDesde + Chr$(34) + " to " + Chr$(34) + FHasta + Chr$(34)
    Tres = " and {PrueTerFarma.Tipo} in " + WTipo + " to " + WTipo
    
    Lista.GroupSelectionFormula = Uno + Dos + Tres
    Lista.SelectionFormula = Uno + Dos + Tres
   
    If ImpreListado.Value = True Then
        Lista.Destination = 1
            Else
        Lista.Destination = 0
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Lista.SQLQuery = "SELECT PrueTerFarma.Tipo, PrueTerFarma.Partida, PrueTerFarma.Producto, PrueTerFarma.Fecha, PrueTerFarma.FechaOrd, PrueTerFarma.Codigo, PrueTerFarma.Valor, PrueTerFarma.Resultado, " _
                + "Ensayos.Descripcion, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.PrueTerFarma PrueTerFarma, " _
                + DSQ + ".dbo.Ensayos Ensayos, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "PrueTerFarma.Codigo = Ensayos.Codigo AND " _
                + "PrueTerFarma.Producto = Terminado.Codigo AND " _
                + "PrueTerFarma.Tipo >= " + WTipo + " AND " _
                + "PrueTerFarma.Tipo <= " + WTipo + " AND " _
                + "PrueTerFarma.Producto >= '" + Desde.Text + "' AND " _
                + "PrueTerFarma.Producto <= '" + Desde.Text + "' AND " _
                + "PrueTerFarma.FechaOrd >= '" + FDesde + "' AND " _
                + "PrueTerFarma.FechaOrd <= '" + FHasta + "'"
    
    Lista.Connect = Connect()
    
    Lista.Action = 1
    
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()

    WPasa = "S"
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Partida.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    spHoja = "ListaHoja " + "'" + Partida.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        If rstHoja!Producto <> Producto.Text Then
            m$ = "El Codigo de Producto de la partida no coincide con el informado"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
        End If
        rstHoja.Close
                    Else
        m$ = "Codigo de Partida invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If WPasa = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PrueTerFarma"
        ZSql = ZSql + " Where PrueTerFarma.Partida = " + "'" + Partida.Text + "'"
        spPrueterFarma = ZSql
        Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueterFarma.RecordCount > 0 Then
            m$ = "Prueba ya ingresada"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
            rstPrueterFarma.Close
        End If
    End If

    If WPasa = "S" Then
    
        WTipo = "1"
        WPartida = Partida.Text
        WProducto = Producto.Text
        WFecha = Fecha.Text
        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = ""
        
        WRenglon = 0
        For iRow = 1 To 100
    
            WVector1.Row = iRow
            
            WVector1.Col = 1
            WCodigo = WVector1.Text
        
            WVector1.Col = 2
            WDesEnsayo = WVector1.Text
        
            WVector1.Col = 3
            WValor = Trim(WVector1.Text)
            
            WVector1.Col = 4
            WResultado = Trim(WVector1.Text)
        
            If Val(WCodigo) <> 0 Or WResultado <> "" Then
        
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WPartida = Partida.Text
                Call Ceros(WPartida, 6)
                        
                WClave = WTipo + Producto.Text + WPartida + Auxi
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO PrueTerFarma ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Partida ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Producto ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Valor ,"
                ZSql = ZSql + "Resultado ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "Aspecto ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Confecciono ,"
                ZSql = ZSql + "Liberada )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + WPartida + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + Producto.Text + "',"
                ZSql = ZSql + "'" + WFecha + "',"
                ZSql = ZSql + "'" + WFechaord + "',"
                ZSql = ZSql + "'" + WCodigo + "',"
                ZSql = ZSql + "'" + Left$(WValor, 50) + "',"
                ZSql = ZSql + "'" + Left$(WResultado, 50) + "',"
                ZSql = ZSql + "'" + WEnsayo + "',"
                ZSql = ZSql + "'" + WAspecto + "',"
                ZSql = ZSql + "'" + WObservaciones + "',"
                ZSql = ZSql + "'" + WConfecciono + "',"
                ZSql = ZSql + "'" + WLiberada + "')"
                
                spPrueterFarma = ZSql
                Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        
            End If
            
        Next iRow
        
        Rem by nan 27-8
    
        Call Registro_Click
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
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Partida.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    spHoja = "ListaHoja " + "'" + Partida.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        If rstHoja!Producto <> Producto.Text Then
            m$ = "El Codigo de Producto de la partida no coincide con el informado"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
        End If
        rstHoja.Close
                    Else
        m$ = "Codigo de Partida invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If WPasa = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PrueTerFarma"
        ZSql = ZSql + " Where PrueTerFarma.Partida = " + "'" + Partida.Text + "'"
        spPrueterFarma = ZSql
        Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueterFarma.RecordCount > 0 Then
            m$ = "Prueba ya ingresada"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
            rstPrueterFarma.Close
        End If
    End If

    If WPasa = "S" Then
    
        WTipo = "1"
        WPartida = Partida.Text
        WProducto = Producto.Text
        WFecha = Fecha.Text
        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = ""
        
        WRenglon = 0
        For iRow = 1 To 100
    
            WVector1.Row = iRow
            
            WVector1.Col = 1
            WCodigo = WVector1.Text
        
            WVector1.Col = 2
            WDesEnsayo = WVector1.Text
        
            WVector1.Col = 3
            WValor = Trim(WVector1.Text)
            
            WVector1.Col = 4
            WResultado = Trim(WVector1.Text)
        
            If Val(WCodigo) <> 0 Or WResultado <> "" Then
        
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WPartida = Partida.Text
                Call Ceros(WPartida, 6)
                        
                WClave = WTipo + Producto.Text + WPartida + Auxi
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO PrueTerFarma ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Partida ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Producto ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Valor ,"
                ZSql = ZSql + "Resultado ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "Aspecto ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Confecciono ,"
                ZSql = ZSql + "Liberada )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + WPartida + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + Producto.Text + "',"
                ZSql = ZSql + "'" + WFecha + "',"
                ZSql = ZSql + "'" + WFechaord + "',"
                ZSql = ZSql + "'" + WCodigo + "',"
                ZSql = ZSql + "'" + WValor + "',"
                ZSql = ZSql + "'" + WResultado + "',"
                ZSql = ZSql + "'" + WEnsayo + "',"
                ZSql = ZSql + "'" + WAspecto + "',"
                ZSql = ZSql + "'" + WObservaciones + "',"
                ZSql = ZSql + "'" + WConfecciono + "',"
                ZSql = ZSql + "'" + WLiberada + "')"
                
                spPrueterFarma = ZSql
                Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        
            End If
            
        Next iRow
        
        Call CmdLimpiar_Click
        Producto.SetFocus
    
    End If
        
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -     -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Partida.Text = ""
    TipoOri.Text = ""
    Resultado.Text = ""
    
    Rem Producto.BackColor = &H80000005
    
    VersionI.Visible = False
    VersionII.Visible = False
    ImpreVersionI.Visible = False
    ImpreVersionII.Visible = False
    
    Call Limpia_Vector
    
    cmdAddlote.Enabled = True
    CmdAddRechazo.Enabled = True
    Actualiza.Enabled = False
    
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgPruterFarma.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub ConfirmaResultado_Click()

    If Val(Partida.Text) <> 0 Then

        Resultado.Text = UCase(Resultado.Text)

        If Resultado.Text <> "PT" Then
            If Resultado.Text <> "NK" Then
                If Resultado.Text <> "RE" Then
                    If Resultado.Text <> "SE" Then
                        ca% = MsgBox("El resultado debe ser PT, SE, NK, RE", 0, "Ingreso de Hoja de Produccion")
                        Exit Sub
                    End If
                End If
            End If
        End If
    
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM Hoja"
        ZSql = ZSql & " Where Hoja.Hoja = " + "'" + Partida.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WMarcaLabora = IIf(IsNull(rstHoja!MarcaLabora), "", rstHoja!MarcaLabora)
            WProducto = rstHoja!Producto
            WTeorico = rstHoja!Teorico
            rstHoja.Close
            Rem If WMarcaLabora = "S" Then
            Rem     ca% = MsgBox("La Hoja de Produccion ya fue actualizada", 0, "Ingreso de Hoja de Produccion")
            Rem    Exit Sub
            Rem End If
                Else
            ca% = MsgBox("Hoja de Produccion Inexistente", 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
        ZZTipoOri = Left$(Producto.Text, 2)
        If ZZTipoOri = "PT" Or ZZTipoOri = "SE" Then
            TipoOri.Text = ZZTipoOri
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + "TipoOri = " + "'" + ZZTipoOri + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + Partida.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        WProductoNuevo = Resultado.Text + Mid$(WProducto, 3, 10)
    
        If WProducto <> WProductoNuevo Then
            
            Producto.Text = WProductoNuevo
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + "Producto = " + "'" + WProductoNuevo + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + Partida.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            spTerminado = "ConsultaTerminado " + "'" + WProductoNuevo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                WProceso = Str$(rstTerminado!Proceso + WTeorico)
                WEntradas = Str$(rstTerminado!Entradas)
            End If
            WDate = Date$
            rstTerminado.Close
                        
            XParam = "'" + WCodigo + "','" _
                     + WEntradas + "','" _
                     + WProceso + "','" _
                     + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoHoja " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        
            spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                WProceso = Str$(rstTerminado!Proceso - WTeorico)
                WEntradas = Str$(rstTerminado!Entradas)
            End If
            WDate = Date$
            rstTerminado.Close
                            
            XParam = "'" + WCodigo + "','" _
                     + WEntradas + "','" _
                     + WProceso + "','" _
                     + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoHoja " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        WMarcaLabora = "S"
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + "MarcaLabora = " + "'" + WMarcaLabora + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Partida.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        Rem by nan 27-8-2012
        Rem    Call Registro_Click
    
    End If
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


Private Sub Partida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPasa = "S"
        spHoja = "ListaHoja " + "'" + Partida.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            If UCase(rstHoja!Producto) <> UCase(Producto.Text) Then
                m$ = "El Codigo de Producto de la partida no coincide con el informado"
                A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
                WPasa = "N"
            End If
            rstHoja.Close
                        Else
            m$ = "Codigo de Partida invalido"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueTerFarma"
            ZSql = ZSql + " Where PrueTerFarma.Partida = " + "'" + Partida.Text + "'"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
                m$ = "Prueba ya ingresada"
                A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
                Partida.Text = ""
                WPasa = "N"
                rstPrueterFarma.Close
            End If
        End If
        
        If WPasa = "S" Then
            WVector1.Col = 4
            WVector1.Row = 1
            Call StartEdit
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
        Ensayo.SetFocus
    End If
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
        
            Producto.Text = UCase(Producto.Text)
            WProducto = Producto.Text
            
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                rstTerminado.Close
                    Else
                Producto.SetFocus
                Exit Sub
            End If
            
            Call Limpia_Vector
    
            WRenglon = 0
    
    
            Rem If Left$(Producto.Text, 2) = "SE" Then
            Rem     WProducto = "SE" + Mid$(Producto.Text, 3, 10)
            Rem         Else
            Rem     WProducto = "PT" + Mid$(Producto.Text, 3, 10)
            Rem End If
            WProducto = Producto.Text
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaV"
            ZSql = ZSql + " Where CargaV.Terminado = " + "'" + WProducto + "'"
            ZSql = ZSql + " and CargaV.Paso = " + "'" + "99" + "'"
            ZSql = ZSql + " Order by CargaV.Clave"
    
            rsCargaV = ZSql
            Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaV.RecordCount > 0 Then
                With rstCargaV
                    .MoveFirst
                    Do
                        If .EOF = False Then
                
                            WRenglon = WRenglon + 1
                    
                            WVector1.Row = WRenglon
                            Renglon = WRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = Trim(Str$(rstCargaV!Ensayo))
                
                            WVector1.Col = 2
                            WVector1.Text = ""
                
                            WVector1.Col = 3
                            WVector1.Text = Trim(rstCargaV!Valor)
                            
                            WVector1.Col = 4
                            WVector1.Text = ""
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCargaV.Close
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
    
            For Ciclo = 1 To WRenglon
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Ensayos"
                ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
                spEnsayos = ZSql
                Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayos.RecordCount > 0 Then
                    WVector1.TextMatrix(Ciclo, 2) = Trim(rstEnsayos!Descripcion)
                    rstEnsayos.Close
                End If
        
            Next Ciclo
    
            Call Conecta_Empresa
            
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
            Call Limpia_VectorII
            LugarVector = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueTerFarma"
            ZSql = ZSql + " Where PrueTerFarma.Renglon = 1"
            ZSql = ZSql + " Order by PrueTerFarma.Partida"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
            
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueterFarma!Producto <> "" Then
                            If rstPrueterFarma!Producto <> "  -     -   " Then
                                If rstPrueterFarma!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Val(rstPrueterFarma!Tipo) = 1 Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueterFarma!Partida)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueterFarma!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueterFarma!Fecha
                                    IngresaItem = rstPrueterFarma!Clave
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
            rstPrueterFarma.Close
            
            End If
        
        Case Else
    End Select
            
    If XIndice = 0 Then
        Pantalla.Visible = True
            Else
        Muestra.Visible = True
    End If

End Sub

Private Sub Limpia_VectorII()

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

Private Sub Registro_Click()

    ZZProducto = Producto.Text
    If Left(ZZProducto, 2) = "RE" Or Left(ZZProducto, 2) = "NK" Then
        TipoOri.Text = UCase(TipoOri.Text)
        If TipoOri.Text = "PT" Or TipoOri.Text = "SE" Then
            ZZProducto = TipoOri.Text + Mid$(Producto.Text, 3, 10)
                Else
            m$ = "Tipo de Producto original no informado"
            A% = MsgBox(m$, 0, "Actualizacion de Pruebas de Prodcuto Terminado")
            Exit Sub
        End If
    End If

    WRenglon = 0
    For iRow = 1 To 100

        WCodigo = WVector1.TextMatrix(iRow, 1)
        WDesEnsayo = WVector1.TextMatrix(iRow, 2)
        WValor = Trim(WVector1.TextMatrix(iRow, 3))
        WResultado = Trim(WVector1.TextMatrix(iRow, 4))
    
        If Val(WCodigo) <> 0 Or WResultado <> "" Then
    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
    
            WClave = ZZProducto + "0099" + Auxi
        
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaV SET "
            ZSql = ZSql + " Resultado = " + "'" + WResultado + "',"
            ZSql = ZSql + " ObservaI = " + "'" + Ensayo.Text + "',"
            ZSql = ZSql + " ObservaII = " + "'" + Aspecto.Text + "',"
            ZSql = ZSql + " ObservaIII = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " ObservaIV = " + "'" + Confecciono.Text + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spCargaV = ZSql
            Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
    Next iRow




    WTeorico = ""
    spHoja = "ListaHoja " + "'" + Partida.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WTeorico = Str$(rstHoja!Teorico)
        rstHoja.Close
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " ImpreTerminado = " + "'" + Producto.Text + "',"
    ZSql = ZSql + " Partida = " + "'" + Partida.Text + "',"
    ZSql = ZSql + " Fechaing = " + "'" + Fecha.Text + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + WTeorico + "',"
    ZSql = ZSql + " ImprePaso = 99"
    ZSql = ZSql + " Where Terminado = " + "'" + ZZProducto + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Lista.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Valor, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, CargaV.Resultado, CargaV.ObservaI, CargaV.ObservaII, CargaV.ObservaIII, CargaV.ObservaIV, CargaV.ImpreTerminado, " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaV CargaV, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaV.Terminado = Terminado.Codigo AND " _
                    + "CargaV.Terminado >= '" + ZZProducto + "' AND " _
                    + "CargaV.Terminado <= '" + ZZProducto + "' AND " _
                    + "CargaV.Paso = 99"

    Lista.ReportFileName = "ImpreCalidadResultado.rpt"
    
    Uno = "{CargaV.Terminado} in " + Chr$(34) + ZZProducto + Chr$(34) + " to " + Chr$(34) + ZZProducto + Chr$(34)
    Dos = " and {CargaV.Paso} = 99"
    
    Lista.GroupSelectionFormula = Uno + Dos
    Lista.SelectionFormula = Uno + Dos
    
    Lista.Connect = Connect()
    Lista.Destination = 1
    Lista.Action = 1


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
            Call Limpia_VectorII
            LugarVector = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueTerFarma"
            ZSql = ZSql + " Where PrueTerFarma.Renglon = 1"
            ZSql = ZSql + " Order by PrueTerFarma.Partida"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
            
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueterFarma!Producto <> "" Then
                            If rstPrueterFarma!Producto <> "  -     -   " Then
                                If rstPrueterFarma!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Val(rstPrueterFarma!Tipo) = 1 Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueterFarma!Partida)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueterFarma!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueterFarma!Fecha
                                    IngresaItem = rstPrueterFarma!Clave
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
            rstPrueterFarma.Close
            
            End If
            
        Case 3
            Call Limpia_VectorII
            LugarVector = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueterFarma"
            ZSql = ZSql + " Where Producto = " + "'" + Seleccion + "'"
            ZSql = ZSql + " and PrueTerFarma.Renglon = 1"
            ZSql = ZSql + " Order by Producto, Fechaord"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueterFarma!Producto <> "" Then
                            If rstPrueterFarma!Producto <> "  -     -   " Then
                                If rstPrueterFarma!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Val(rstPrueterFarma!Tipo) = 1 Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueterFarma!Partida)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueterFarma!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueterFarma!Fecha
                                    IngresaItem = rstPrueterFarma!Clave
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
            rstPrueterFarma.Close
            Muestra.TopRow = 1
            Muestra.Row = 1
            Muestra.Col = 1
            
            End If
    
        Case 4
            Call Limpia_VectorII
            LugarVector = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueterFarma"
            ZSql = ZSql + " Where Fecha = " + "'" + Seleccion + "'"
            ZSql = ZSql + " and PrueTerFarma.Renglon = 1"
            ZSql = ZSql + " Order by Producto, Fechaord"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueterFarma!Producto <> "" Then
                            If rstPrueterFarma!Producto <> "  -     -   " Then
                                If rstPrueterFarma!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Val(rstPrueterFarma!Tipo) = 1 Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueterFarma!Partida)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueterFarma!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueterFarma!Fecha
                                    IngresaItem = rstPrueterFarma!Clave
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
            rstPrueterFarma.Close
            
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
        Call Limpia_VectorII
        LugarVector = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PrueTerFarma"
        ZSql = ZSql + " WHERE Partida = " + "'" + NumeroPrueba.Text + "'"
        spPrueterFarma = ZSql
        Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueterFarma.RecordCount > 0 Then
            If rstPrueterFarma!Producto <> "" Then
                If rstPrueterFarma!Producto <> "  -     -   " Then
                    If rstPrueterFarma!Producto <> Space$(12) Then
                        LugarVector = LugarVector + 1
                        If Val(rstPrueterFarma!Tipo) = 1 Then
                            Muestra.TextMatrix(LugarVector, 1) = "OK"
                        End If
                        Muestra.TextMatrix(LugarVector, 2) = Str$(rstPrueterFarma!Partida)
                        Muestra.TextMatrix(LugarVector, 3) = rstPrueterFarma!Producto
                        Muestra.TextMatrix(LugarVector, 4) = rstPrueterFarma!Fecha
                        IngresaItem = rstPrueterFarma!Clave
                        WIndice.AddItem IngresaItem
                    End If
                End If
            End If
            
            rstPrueterFarma.Close
            
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
            
            ZSql = ""
            ZSql = ZSql + "Select DISTINCT Producto"
            ZSql = ZSql + " FROM PrueterFarma"
            ZSql = ZSql + " Order by Producto"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueterFarma!Producto <> "" Then
                            If rstPrueterFarma!Producto <> "  -     -   " Then
                                If rstPrueterFarma!Producto <> Space$(12) Then
                                    IngresaItem = rstPrueterFarma!Producto
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
            rstPrueterFarma.Close
            
        Case 4
            WPantalla.AddItem ""
            
            ZSql = ""
            ZSql = ZSql + "Select DISTINCT FechaOrd"
            ZSql = ZSql + " FROM PrueterFarma"
            ZSql = ZSql + " Order by FechaOrd"
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueterFarma
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Right$(rstPrueterFarma!FechaOrd, 2) + "/" + Mid$(rstPrueterFarma!FechaOrd, 5, 2) + "/" + Left$(rstPrueterFarma!FechaOrd, 4)
                        WPantalla.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueterFarma.Close
            
            
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
                Rem Call imprime_Click
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
            WPartida = Mid$(ClavePrue$, 14, 6)
            
            Call Limpia_Vector
            WRenglon = 0
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PrueterFarma"
            ZSql = ZSql + " Where Partida = " + "'" + WPartida + "'"
            ZSql = ZSql + " Order by Clave"
    
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueterFarma.RecordCount > 0 Then
                With rstPrueterFarma
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Partida.Text = rstPrueterFarma!Partida
                            Producto.Text = rstPrueterFarma!Producto
                            Fecha.Text = rstPrueterFarma!Fecha
                            Ensayo.Text = Trim(rstPrueterFarma!Ensayo)
                            Aspecto.Text = Trim(rstPrueterFarma!Aspecto)
                            Observaciones.Text = Trim(rstPrueterFarma!Observaciones)
                            Confecciono.Text = Trim(rstPrueterFarma!Confecciono)
                            Rem ZConfirmacion = IIf(IsNull(rstPrueterFarma!Confirmacion), "", rstPrueterFarma!Confirmacion)
                            Rem If ZConfirmacion = "S" Then
                            Rem Producto.BackColor = &HFF00&
                            Rem         Else
                            Rem     Producto.BackColor = &HFF&
                            Rem End If
                
                            WRenglon = WRenglon + 1
                    
                            WVector1.Row = WRenglon
                            Renglon = WRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = Trim(Str$(rstPrueterFarma!Codigo))
                
                            WVector1.Col = 2
                            WVector1.Text = ""
                
                            WVector1.Col = 3
                            WVector1.Text = Trim(rstPrueterFarma!Valor)
                            
                            WVector1.Col = 4
                            WVector1.Text = Trim(rstPrueterFarma!Resultado)
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPrueterFarma.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja = " + "'" + WPartida + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                VersionI.Text = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
                rstHoja.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Terminado"
            ZSql = ZSql + " Where Codigo = " + "'" + Producto.Text + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                VersionII.Text = rstTerminado!VersionII
                rstTerminado.Close
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
    
            For Ciclo = 1 To WRenglon
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Ensayos"
                ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
                spEnsayos = ZSql
                Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayos.RecordCount > 0 Then
                    WVector1.TextMatrix(Ciclo, 2) = Trim(rstEnsayos!Descripcion)
                    rstEnsayos.Close
                End If
        
            Next Ciclo
    
            Call Conecta_Empresa
  
            cmdAddlote.Enabled = False
            CmdAddRechazo.Enabled = False
            Actualiza.Enabled = True
        
            VersionI.Visible = True
            VersionII.Visible = True
            ImpreVersionI.Visible = True
            ImpreVersionII.Visible = True
                    
            Producto.SetFocus
        
        Case Else
    End Select
    
    End If

End Sub

Private Sub Form_Load()

    Producto.Text = "  -     -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Partida.Text = ""
    TipoOri.Text = ""
    Resultado.Text = ""
    
    VersionI.Visible = False
    VersionII.Visible = False
    ImpreVersionI.Visible = False
    ImpreVersionII.Visible = False

    cmdAddlote.Enabled = True
    CmdAddRechazo.Enabled = True
    Actualiza.Enabled = False
    
    Call Limpia_Vector

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgPruterFarma.Caption = "Ingreso de Ensayos de Productos Terminados de Farma :  " + !Nombre
        End If
    End With
    
    EmpresaActual = Wempresa
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
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

Private Sub Confiirmacion_Click()
    WProceso = 2
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub


Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case WProceso
            Case 0
                If UCase(WClave.Text) = "SEGURO" Then
                    Pass.Visible = False
                    Call ModificaPrueba
                End If
            Case 1
                If UCase(WClave.Text) = "SEGURO" Then
                    Pass.Visible = False
                    Call ActualizaPrueba
                End If
            Case Else
                If UCase(WClave.Text) = "DT24JF" Then
                    Pass.Visible = False
                    Call ConfirmaPrueba
                End If
        End Select
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub ModificaPrueba()

    ZSql = ""
    ZSql = ZSql + "UPDATE PrueTerFarma SET "
    ZSql = ZSql + " Ensayo = " + "'" + Ensayo.Text + "',"
    ZSql = ZSql + " Aspecto = " + "'" + Aspecto.Text + "',"
    ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
    ZSql = ZSql + " Confecciono = " + "'" + Confecciono.Text + "'"
    ZSql = ZSql + " Where Tipo = " + "'" + "1" + "'"
    ZSql = ZSql + " and Partida = " + "'" + Partida.Text + "'"
    spPrueterFarma = ZSql
    Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)

    Call CmdLimpiar_Click
    Producto.SetFocus
    
End Sub


Private Sub ConfirmaPrueba()

    ZSql = ""
    ZSql = ZSql + "UPDATE PrueTerFarma SET "
    ZSql = ZSql + " Confirmacion = " + "'" + "S" + "'"
    ZSql = ZSql + " Where Tipo = " + "'" + "1" + "'"
    ZSql = ZSql + " and Partida = " + "'" + Partida.Text + "'"
    spPrueterFarma = ZSql
    Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)

    m$ = "Desbloqueo de la partida para facturacion realizado"
    A% = MsgBox(m$, 0, "Actualizacion de Pruebas de Prodcuto Terminado")

    Call CmdLimpiar_Click
    Producto.SetFocus
    
End Sub



Private Sub ActualizaPrueba()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PrueTerFarma"
    ZSql = ZSql + " Where PrueTerFarma.Partida = " + "'" + Partida.Text + "'"
    spPrueterFarma = ZSql
    Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueterFarma.RecordCount > 0 Then
        Select Case rstPrueterFarma!Tipo
            Case 2
                WTipo = "2"
            Case Else
                WTipo = "1"
        End Select
        WLiberada = Str$(rstPrueterFarma!Liberada)
        rstPrueterFarma.Close
            Else
        m$ = "Prueba no ingresada"
        A% = MsgBox(m$, 0, "Actualizacion de Pruebas de Prodcuto Terminado")
        Exit Sub
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE PrueTerFarma"
    ZSql = ZSql + " Where Partida = " + "'" + Partida.Text + "'"
    spPrueterFarma = ZSql
    Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    WPartida = Partida.Text
    WProducto = Producto.Text
    WFecha = Fecha.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WEnsayo = Ensayo.Text
    WAspecto = Aspecto.Text
    WObservaciones = Observaciones.Text
    WConfecciono = Confecciono.Text
        
    WRenglon = 0
    For iRow = 1 To 100
    
        WVector1.Row = iRow
            
        WVector1.Col = 1
        WCodigo = WVector1.Text
        
        WVector1.Col = 2
        WDesEnsayo = WVector1.Text
        
        WVector1.Col = 3
        WValor = Trim(WVector1.Text)
            
        WVector1.Col = 4
        WResultado = Trim(WVector1.Text)
        
        If Val(WCodigo) <> 0 Or WResultado <> "" Then
        
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WPartida = Partida.Text
            Call Ceros(WPartida, 6)
                        
            WClave = WTipo + Producto.Text + WPartida + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO PrueTerFarma ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Partida ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Producto ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Valor ,"
            ZSql = ZSql + "Resultado ,"
            ZSql = ZSql + "Ensayo ,"
            ZSql = ZSql + "Aspecto ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Confecciono ,"
            ZSql = ZSql + "Liberada )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WPartida + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Producto.Text + "',"
            ZSql = ZSql + "'" + WFecha + "',"
            ZSql = ZSql + "'" + WFechaord + "',"
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WValor + "',"
            ZSql = ZSql + "'" + WResultado + "',"
            ZSql = ZSql + "'" + WEnsayo + "',"
            ZSql = ZSql + "'" + WAspecto + "',"
            ZSql = ZSql + "'" + WObservaciones + "',"
            ZSql = ZSql + "'" + WConfecciono + "',"
            ZSql = ZSql + "'" + WLiberada + "')"
                
            spPrueterFarma = ZSql
            Set rstPrueterFarma = db.OpenRecordset(spPrueterFarma, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
            
    Next iRow
    
    Call CmdLimpiar_Click
    Producto.SetFocus

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
            Rem WVector1.Col = 1
        Case Else
            Rem If WVector1.Col < WVector1.Cols - 1 Then
            Rem     WVector1.Col = WVector1.Col + 1
            Rem End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
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
                WVector1.Text = "Ensayo"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Valor"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Resultado"
                WVector1.ColWidth(Ciclo) = 2600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
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




