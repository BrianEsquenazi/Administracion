VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgLiberaTerminado 
   Caption         =   "Ingreso de Ensayos de Productos Terminados"
   ClientHeight    =   8520
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   11880
   Begin VB.TextBox NroDevol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9000
      TabIndex        =   94
      Text            =   " "
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TipoPro 
      BeginProperty Font 
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
      TabIndex        =   92
      Text            =   " "
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Cantidad 
      BeginProperty Font 
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
      MaxLength       =   10
      TabIndex        =   90
      Text            =   " "
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Observa 
      BeginProperty Font 
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
      TabIndex        =   88
      Text            =   " "
      Top             =   360
      Width           =   3495
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
      Left            =   8880
      MaxLength       =   6
      TabIndex        =   85
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   4320
      TabIndex        =   81
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   83
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   840
         TabIndex        =   82
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
         TabIndex        =   84
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame PantaNumeroPrueba 
      Height          =   855
      Left            =   3240
      TabIndex        =   78
      Top             =   7200
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
         TabIndex        =   79
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
         TabIndex        =   80
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
      TabIndex        =   67
      Top             =   7080
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
      TabIndex        =   66
      Top             =   7080
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
      TabIndex        =   65
      Top             =   7080
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
      TabIndex        =   64
      Top             =   7080
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
      ItemData        =   "LiberaTerminado.frx":0000
      Left            =   3720
      List            =   "LiberaTerminado.frx":0007
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   1815
      Left            =   1200
      TabIndex        =   62
      Top             =   6720
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
      ItemData        =   "LiberaTerminado.frx":0015
      Left            =   480
      List            =   "LiberaTerminado.frx":001C
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Partida 
      BeginProperty Font 
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
      MaxLength       =   10
      TabIndex        =   61
      Text            =   " "
      Top             =   0
      Width           =   1455
   End
   Begin MSMask.MaskEdBox fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   39
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
      TabIndex        =   37
      Text            =   " "
      Top             =   6360
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
      TabIndex        =   36
      Text            =   " "
      Top             =   6120
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
      TabIndex        =   35
      Text            =   " "
      Top             =   5880
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
      TabIndex        =   34
      Text            =   " "
      Top             =   5640
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1080
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
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   10320
      TabIndex        =   27
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   5040
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   4560
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   4080
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   3600
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   3120
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   2640
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   20
      Text            =   " "
      Top             =   2160
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   1680
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   18
      Text            =   " "
      Top             =   1200
      Width           =   3975
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
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   17
      Text            =   " "
      Top             =   720
      Width           =   3975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
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
      Left            =   6960
      TabIndex        =   3
      Top             =   5760
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
      Left            =   8160
      TabIndex        =   2
      Top             =   5760
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
      Left            =   5760
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Nro.Ent.Dev."
      Height          =   255
      Left            =   7920
      TabIndex        =   95
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Left            =   10320
      TabIndex        =   93
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5280
      TabIndex        =   91
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   89
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   7920
      TabIndex        =   87
      Top             =   0
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9960
      TabIndex        =   86
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Std1010 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   77
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Std99 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   76
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label Std88 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   75
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Std77 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   74
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Std66 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   73
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Std55 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   72
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label Std44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   71
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Std33 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   70
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Std22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   69
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Std11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   68
      Top             =   960
      Width           =   4335
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
      Left            =   5280
      TabIndex        =   60
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Ensayo10 
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
      Left            =   120
      TabIndex        =   59
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Ensayo9 
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
      Left            =   120
      TabIndex        =   58
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Ensayo8 
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
      Left            =   120
      TabIndex        =   57
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Ensayo7 
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
      Left            =   120
      TabIndex        =   56
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Ensayo6 
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
      Left            =   120
      TabIndex        =   55
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Ensayo5 
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
      Left            =   120
      TabIndex        =   54
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Ensayo4 
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
      Left            =   120
      TabIndex        =   53
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Ensayo3 
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
      Left            =   120
      TabIndex        =   52
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Ensayo2 
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
      Left            =   120
      TabIndex        =   51
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Ensayo1 
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
      Left            =   120
      TabIndex        =   50
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   49
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   46
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   45
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   44
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   43
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3360
      TabIndex        =   40
      Top             =   720
      Width           =   4335
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
      Left            =   2760
      TabIndex        =   38
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
      TabIndex        =   33
      Top             =   6360
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
      TabIndex        =   32
      Top             =   6120
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
      TabIndex        =   31
      Top             =   5880
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
      TabIndex        =   30
      Top             =   5640
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
      TabIndex        =   29
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descri10 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   5040
      Width           =   2340
   End
   Begin VB.Label Descri9 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   4560
      Width           =   2340
   End
   Begin VB.Label Descri8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   4080
      Width           =   2340
   End
   Begin VB.Label Descri7 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   3600
      Width           =   2340
   End
   Begin VB.Label Descri6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   3120
      Width           =   2340
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2640
      Width           =   2340
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2160
      Width           =   2340
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   2340
   End
   Begin VB.Label descri2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   2340
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   2340
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgLiberaTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstLiberaTerminado As Recordset
Dim spLiberaTerminado As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecificacionesUnifica As Recordset
Dim spEspecificacionesUnifica As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim XParam As String
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
Dim WCodigoMayor As Double

Dim EmpresaActual As String

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()

    If Trim(TipoPro.Text) <> "" Then
        If Trim(TipoPro.Text) <> "DK" And Trim(TipoPro.Text) <> "NS" And Trim(TipoPro.Text) <> "NW" And Trim(TipoPro.Text) <> "NQ" Then
            m$ = "Tipo de Producto incorrecto"
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            Exit Sub
        End If
            Else
        TipoPro.Text = Left$(Producto.Text, 2)
    End If

    Rem If Val(Cantidad.Text) = 0 Then
    Rem     m$ = "La Cantidad a Liberar no puede ser igual a 0"
    Rem     A% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
    Rem     Exit Sub
    Rem End If

    WPasa = "S"
    
    If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
    
    
    
    
    
    
    
    
    
    
        If Left$(Producto.Text, 2) = "DY" Then
            WArti = "DY-" + Right$(Producto.Text, 7)
            WArtiDev = "DK-" + Right$(Producto.Text, 7)
                Else
            If Left$(Producto.Text, 2) = "DS" Then
                WArti = "DS-" + Right$(Producto.Text, 7)
                WArtiDev = "NS-" + Right$(Producto.Text, 7)
                    Else
                WArti = "DW-" + Right$(Producto.Text, 7)
                WArtiDev = "NW-" + Right$(Producto.Text, 7)
            End If
        End If
    
        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            rstArticulo.Close
                Else
            m$ = "Codigo de Producto invalido"
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
        End If
    
        If Partida.Text = "" Then
            m$ = "Codigo de Partida invalido"
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
        End If
    
        WEntra = "N"
        WEstado = ""
        
        Sql1 = "Select *"
        Sql2 = " FROM Laudo"
        Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
        Sql4 = " and Laudo.PartiOri = " + "'" + Partida.Text + "'"
        spLaudo = Sql1 + Sql2 + Sql3 + Sql4
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstLaudo!Estado), "N", rstLaudo!Estado)
            rstLaudo.Close
        End If
                
        If WEntra = "N" Then
            Sql1 = "Select *"
            Sql2 = " FROM Guia"
            Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
            Sql4 = " and Guia.PartiOri = " + "'" + Partida.Text + "'"
            spMovguia = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "N", rstMovguia!Estado)
                rstMovguia.Close
            End If
        End If
        
        Rem If WEntra = "S" And WEstado <> "N" Then
        Rem     Sql1 = "Select *"
        Rem     Sql2 = " FROM EntDev"
        Rem     Sql3 = " Where EntDev.Terminado = " + "'" + "NK" + Mid$(WArti, 3, 10) + "'"
        Rem     Sql4 = " and EntDev.PartiOri = " + "'" + Partida.Text + "'"
        Rem     spEntdev = Sql1 + Sql2 + Sql3 + Sql4
        Rem     Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstEntdev.RecordCount > 0 Then
        Rem         If rstEntdev!Saldo > 0 Then
        Rem             WEstado = "N"
        Rem        End If
        Rem        rstEntdev.Close
        Rem     End If
        Rem End If
        
        If WEntra = "N" Then
            m$ = "Codigo de Partida o Producto Invalido "
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
                Else
            If WEstado <> "N" Then
                m$ = "La partida ya se encuentra liberada"
                a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                WPasa = "N"
            End If
        End If
        
        If WPasa = "S" Then
            If Cliente.Text <> "" Then
            
                Auxi1 = Left$(WArtiDev, 3) + "00" + Right$(WArtiDev, 7)
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM EntDev"
                ZSql = ZSql + " Where EntDev.Terminado = " + "'" + Auxi1 + "'"
                ZSql = ZSql + " and EntDev.PartiOri = " + "'" + Partida.Text + "'"
                ZSql = ZSql + " and EntDev.Cliente = " + "'" + Cliente.Text + "'"
                spEntdev = ZSql
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                     Cantidad.Text = Str$(rstEntdev!Cantidad)
                     NroDevol.Text = rstEntdev!Codigo
                     rstEntdev.Close
                         Else
                     m$ = "No se encontro datos de entrada de devolucion que coincidan con los informados"
                     a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                     WPasa = "N"
                 End If
                 
            End If
        End If
    
        If WPasa = "S" Then
        
            Sql1 = "Select Max(Codigo) as [CodigoMayor]"
            Sql2 = " FROM LiberaTerminado"
            spLiberaTerminado = Sql1 + Sql2
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstLiberaTerminado.RecordCount > 0 Then
                rstLiberaTerminado.MoveLast
                WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
                Lote = Str$(WCodigoMayor)
                rstLiberaTerminado.Close
                    Else
                Lote = "0"
            End If
        
            WCodigo = Str$(Val(Lote) + 1)
            WProducto = Producto.Text
            WFecha = Fecha.Text
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WPartida = ""
            WPartiOri = Partida.Text
            
            WValor1 = Valor1.Text
            WValor2 = Valor2.Text
            WValor3 = Valor3.Text
            WValor4 = Valor4.Text
            WValor5 = Valor5.Text
            WValor6 = Valor6.Text
            WValor7 = Valor7.Text
            WValor8 = Valor8.Text
            WValor9 = Valor9.Text
            WValor10 = Valor10.Text
            WEnsayo = Ensayo.Text
            WAspecto = Aspecto.Text
            WObservaciones = Observaciones.Text
            WConfecciono = Confecciono.Text
            WMarca = "N"
            WCliente = Cliente.Text
            WObserva = Observa.Text
            WCantidad = Cantidad.Text
            WFacturado = "0"
            WOrigen = "L"
            WTipo = TipoPro.Text
            WImpreProdI = "N"
            WImpreProdII = "N"
            WImpreProdIII = "N"
            WImpreVentas = "N"
            WTipoPro = ""
            
            XTipoPro = ""
            XCodigo = Val(Mid$(Producto.Text, 4, 5))
            If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
                XTipoPro = "CO"
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
                
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
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
            
            WTipoPro = XTipoPro
            
            Select Case WTipoPro
                Case "CO", "PG"
                    WImpreProdI = "S"
                Case "BI", "PT"
                    WImpreProdII = "S"
                Case "FA"
                    WImpreProdIII = "S"
                Case Else
            End Select
                    
            ZSql = ""
            ZSql = ZSql & "INSERT INTO LiberaTerminado ("
            ZSql = ZSql & "Codigo, "
            ZSql = ZSql & "Producto, "
            ZSql = ZSql & "Fecha, "
            ZSql = ZSql & "OrdFecha, "
            ZSql = ZSql & "Partida, "
            ZSql = ZSql & "PartiOri, "
            ZSql = ZSql & "PedidoDevol, "
            ZSql = ZSql & "Valor1, "
            ZSql = ZSql & "Valor2, "
            ZSql = ZSql & "Valor3, "
            ZSql = ZSql & "Valor4, "
            ZSql = ZSql & "Valor5, "
            ZSql = ZSql & "Valor6, "
            ZSql = ZSql & "Valor7, "
            ZSql = ZSql & "Valor8, "
            ZSql = ZSql & "Valor9, "
            ZSql = ZSql & "Valor10, "
            ZSql = ZSql & "Ensayo, "
            ZSql = ZSql & "Aspecto, "
            ZSql = ZSql & "Observaciones, "
            ZSql = ZSql & "Confecciono, "
            ZSql = ZSql & "Marca, "
            ZSql = ZSql & "Cliente, "
            ZSql = ZSql & "Cantidad, "
            ZSql = ZSql & "Facturado, "
            ZSql = ZSql & "Observa, "
            ZSql = ZSql & "Origen, "
            ZSql = ZSql & "Tipo, "
            ZSql = ZSql & "ImpreProdI, "
            ZSql = ZSql & "ImpreProdII, "
            ZSql = ZSql & "ImpreProdIII, "
            ZSql = ZSql & "ImpreVentas, "
            ZSql = ZSql & "TipoPro) "
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WCodigo + "',"
            ZSql = ZSql & "'" + WProducto + "',"
            ZSql = ZSql & "'" + WFecha + "',"
            ZSql = ZSql & "'" + WOrdFecha + "',"
            ZSql = ZSql & "'" + WPartida + "',"
            ZSql = ZSql & "'" + WPartiOri + "',"
            ZSql = ZSql & "'" + NroDevol.Text + "',"
            ZSql = ZSql & "'" + WValor1 + "',"
            ZSql = ZSql & "'" + WValor2 + "',"
            ZSql = ZSql & "'" + WValor3 + "',"
            ZSql = ZSql & "'" + WValor4 + "',"
            ZSql = ZSql & "'" + WValor5 + "',"
            ZSql = ZSql & "'" + WValor6 + "',"
            ZSql = ZSql & "'" + WValor7 + "',"
            ZSql = ZSql & "'" + WValor8 + "',"
            ZSql = ZSql & "'" + WValor9 + "',"
            ZSql = ZSql & "'" + WValor10 + "',"
            ZSql = ZSql & "'" + WEnsayo + "',"
            ZSql = ZSql & "'" + WAspecto + "',"
            ZSql = ZSql & "'" + WObservaciones + "',"
            ZSql = ZSql & "'" + WConfecciono + "',"
            ZSql = ZSql & "'" + WMarca + "',"
            ZSql = ZSql & "'" + WCliente + "',"
            ZSql = ZSql & "'" + WCantidad + "',"
            ZSql = ZSql & "'" + WFacturado + "',"
            ZSql = ZSql & "'" + WObserva + "',"
            ZSql = ZSql & "'" + WOrigen + "',"
            ZSql = ZSql & "'" + WTipo + "',"
            ZSql = ZSql & "'" + WImpreProdI + "',"
            ZSql = ZSql & "'" + WImpreProdII + "',"
            ZSql = ZSql & "'" + WImpreProdIII + "',"
            ZSql = ZSql & "'" + WImpreVentas + "',"
            ZSql = ZSql & "'" + WTipoPro + "')"
          
            spLiberaTerminado = ZSql
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            WEntra = "N"
            
            Sql1 = "Select *"
            Sql2 = " FROM Laudo"
            Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
            Sql4 = " and Laudo.PartiOri = " + "'" + Partida.Text + "'"
            spLaudo = Sql1 + Sql2 + Sql3 + Sql4
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.Close
                WEntra = "S"
                WMarcaEstado = ""
                Sql1 = "UPDATE Laudo SET "
                Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                Sql4 = " and Laudo.PartiOri = " + "'" + Partida.Text + "'"
                spLaudo = Sql1 + Sql2 + Sql3 + Sql4
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                
            If WEntra = "N" Then
                Sql1 = "Select *"
                Sql2 = " FROM Guia"
                Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                Sql4 = " and Guia.PartiOri = " + "'" + Partida.Text + "'"
                spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    rstMovguia.Close
                    WMarcaEstado = ""
                    Sql1 = "UPDATE Guia SET "
                    Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                    Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Guia.PartiOri = " + "'" + Partida.Text + "'"
                    spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
                
            Auxi1 = Left$(WArtiDev, 3) + "00" + Right$(WArtiDev, 7)
            ZSql = ""
            ZSql = ZSql + "UPDATE EntDev SET "
            ZSql = ZSql + "Estado  = 'PT'"
            ZSql = ZSql + " Where EntDev.Terminado = " + "'" + Auxi1 + "'"
            ZSql = ZSql + " and EntDev.PartiOri = " + "'" + Partida.Text + "'"
            ZSql = ZSql + " and EntDev.Cliente = " + "'" + Cliente.Text + "'"
            spEntdev = ZSql
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
        
        
        
        
        
        
    
            Else
            
            
            
            
            
            
            
            
            
    
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            rstTerminado.Close
                Else
            m$ = "Codigo de Producto invalido"
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
        End If
    
        If Val(Partida.Text) = 0 Then
            m$ = "Codigo de Partida invalido"
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
        End If
    
        WEntra = "N"
        WEstado = ""
        
        XParam = "'" + Partida.Text + "','" _
                     + Producto.Text + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "N", rstHoja!Estado)
            rstHoja.Close
        End If
                
        If WEntra = "N" Then
            XParam = "'" + Producto.Text + "','" _
                     + Partida.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "N", rstMovguia!Estado)
                rstMovguia.Close
            End If
        End If
        
        If WEntra = "S" And WEstado <> "N" Then
            Sql1 = "Select *"
            Sql2 = " FROM EntDev"
            Sql3 = " Where EntDev.Terminado = " + "'" + "NK" + Mid$(Producto, 3, 10) + "'"
            Sql4 = " and EntDev.Lote = " + "'" + Partida.Text + "'"
            Sql5 = " and EntDev.Cliente = " + "'" + Cliente.Text + "'"
            spEntdev = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            If rstEntdev.RecordCount > 0 Then
                Cantidad.Text = rstEntdev!Saldo
                NroDevol.Text = rstEntdev!Codigo
                If rstEntdev!Saldo > 0 Then
                    WEstado = "N"
                End If
                rstEntdev.Close
            End If
        End If
        
        If WEntra = "N" Then
            m$ = "Codigo de Partida o Producto Invalido "
            a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
            WPasa = "N"
                Else
            If WEstado <> "N" Then
                m$ = "La partida ya se encuentra liberada"
                a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                WPasa = "N"
            End If
        End If
        
        If WPasa = "S" Then
            If Cliente.Text <> "" Then
            
                If Left$(Producto.Text, 2) = "PT" Then
                    WArtiDev = "NK" + Right$(Producto.Text, 10)
                        Else
                    If Left$(Producto.Text, 2) = "DY" Then
                        WArtiDev = "DK" + Right$(Producto.Text, 10)
                            Else
                        If Left$(Producto.Text, 2) = "DS" Then
                            WArtiDev = "NS" + Right$(Producto.Text, 10)
                                Else
                            WArtiDev = "NW" + Right$(Producto.Text, 10)
                        End If
                    End If
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM EntDev"
                ZSql = ZSql + " Where EntDev.Terminado = " + "'" + WArtiDev + "'"
                ZSql = ZSql + " and EntDev.Lote = " + "'" + Partida.Text + "'"
                ZSql = ZSql + " and EntDev.Cliente = " + "'" + Cliente.Text + "'"
                spEntdev = ZSql
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                     rstEntdev.Close
                         Else
                     m$ = "No se encontro datos de entrada de devolucion que coincidan con los informados"
                     a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                     WPasa = "N"
                 End If
                 
            End If
        End If
    
        If WPasa = "S" Then
    
            Sql1 = "Select Max(Codigo) as [CodigoMayor]"
            Sql2 = " FROM LiberaTerminado"
            spLiberaTerminado = Sql1 + Sql2
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstLiberaTerminado.RecordCount > 0 Then
                rstLiberaTerminado.MoveLast
                WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
                Lote = Str$(WCodigoMayor)
                rstLiberaTerminado.Close
                    Else
                Lote = "0"
            End If
        
            WCodigo = Str$(Val(Lote) + 1)
            WProducto = Producto.Text
            WFecha = Fecha.Text
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WPartida = Partida.Text
            WPartiOri = ""
            WValor1 = Valor1.Text
            WValor2 = Valor2.Text
            WValor3 = Valor3.Text
            WValor4 = Valor4.Text
            WValor5 = Valor5.Text
            WValor6 = Valor6.Text
            WValor7 = Valor7.Text
            WValor8 = Valor8.Text
            WValor9 = Valor9.Text
            WValor10 = Valor10.Text
            WEnsayo = Ensayo.Text
            WAspecto = Aspecto.Text
            WObservaciones = Observaciones.Text
            WConfecciono = Confecciono.Text
            WMarca = "N"
            WCliente = Cliente.Text
            WObserva = Observa.Text
            WCantidad = Cantidad.Text
            WFacturado = "0"
            WOrigen = "L"
            WTipo = "PT"
            WImpreProdI = "N"
            WImpreProdII = "N"
            WImpreProdIII = "N"
            WImpreVentas = "N"
            WTipoPro = ""
            
            XTipoPro = ""
            XCodigo = Val(Mid$(Producto.Text, 4, 5))
            If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
                XTipoPro = "CO"
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
                
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
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
                
            WTipoPro = XTipoPro
            
            Select Case WTipoPro
                Case "CO", "PG"
                    WImpreProdI = "S"
                Case "BI", "PT"
                    WImpreProdII = "S"
                Case "FA"
                    WImpreProdIII = "S"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql & "INSERT INTO LiberaTerminado ("
            ZSql = ZSql & "Codigo, "
            ZSql = ZSql & "Producto, "
            ZSql = ZSql & "Fecha, "
            ZSql = ZSql & "OrdFecha, "
            ZSql = ZSql & "Partida, "
            ZSql = ZSql & "PartiOri, "
            ZSql = ZSql & "PedidoDevol, "
            ZSql = ZSql & "Valor1, "
            ZSql = ZSql & "Valor2, "
            ZSql = ZSql & "Valor3, "
            ZSql = ZSql & "Valor4, "
            ZSql = ZSql & "Valor5, "
            ZSql = ZSql & "Valor6, "
            ZSql = ZSql & "Valor7, "
            ZSql = ZSql & "Valor8, "
            ZSql = ZSql & "Valor9, "
            ZSql = ZSql & "Valor10, "
            ZSql = ZSql & "Ensayo, "
            ZSql = ZSql & "Aspecto, "
            ZSql = ZSql & "Observaciones, "
            ZSql = ZSql & "Confecciono, "
            ZSql = ZSql & "Marca, "
            ZSql = ZSql & "Cliente, "
            ZSql = ZSql & "Cantidad, "
            ZSql = ZSql & "Facturado, "
            ZSql = ZSql & "Observa, "
            ZSql = ZSql & "Origen, "
            ZSql = ZSql & "Tipo, "
            ZSql = ZSql & "ImpreProdI, "
            ZSql = ZSql & "ImpreProdII, "
            ZSql = ZSql & "ImpreProdIII, "
            ZSql = ZSql & "ImpreVentas, "
            ZSql = ZSql & "TipoPro) "
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WCodigo + "',"
            ZSql = ZSql & "'" + WProducto + "',"
            ZSql = ZSql & "'" + WFecha + "',"
            ZSql = ZSql & "'" + WOrdFecha + "',"
            ZSql = ZSql & "'" + WPartida + "',"
            ZSql = ZSql & "'" + WPartiOri + "',"
            ZSql = ZSql & "'" + NroDevol.Text + "',"
            ZSql = ZSql & "'" + WValor1 + "',"
            ZSql = ZSql & "'" + WValor2 + "',"
            ZSql = ZSql & "'" + WValor3 + "',"
            ZSql = ZSql & "'" + WValor4 + "',"
            ZSql = ZSql & "'" + WValor5 + "',"
            ZSql = ZSql & "'" + WValor6 + "',"
            ZSql = ZSql & "'" + WValor7 + "',"
            ZSql = ZSql & "'" + WValor8 + "',"
            ZSql = ZSql & "'" + WValor9 + "',"
            ZSql = ZSql & "'" + WValor10 + "',"
            ZSql = ZSql & "'" + WEnsayo + "',"
            ZSql = ZSql & "'" + WAspecto + "',"
            ZSql = ZSql & "'" + WObservaciones + "',"
            ZSql = ZSql & "'" + WConfecciono + "',"
            ZSql = ZSql & "'" + WMarca + "',"
            ZSql = ZSql & "'" + WCliente + "',"
            ZSql = ZSql & "'" + WCantidad + "',"
            ZSql = ZSql & "'" + WFacturado + "',"
            ZSql = ZSql & "'" + WObserva + "',"
            ZSql = ZSql & "'" + WOrigen + "',"
            ZSql = ZSql & "'" + WTipo + "',"
            ZSql = ZSql & "'" + WImpreProdI + "',"
            ZSql = ZSql & "'" + WImpreProdII + "',"
            ZSql = ZSql & "'" + WImpreProdIII + "',"
            ZSql = ZSql & "'" + WImpreVentas + "',"
            ZSql = ZSql & "'" + WTipoPro + "')"
          
            spLiberaTerminado = ZSql
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            WEntra = "N"
            XParam = "'" + Partida.Text + "','" _
                         + Producto.Text + "'"
                         
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                rstHoja.Close
                WEntra = "S"
                WMarcaEstado = ""
                Sql1 = "UPDATE Hoja SET "
                Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                Sql3 = " Where Hoja.Producto = " + "'" + Producto.Text + "'"
                Sql4 = " and Hoja.Hoja = " + "'" + Partida.Text + "'"
                spHoja = Sql1 + Sql2 + Sql3 + Sql4
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            If WEntra = "N" Then
                XParam = "'" + Producto.Text + "','" _
                                 + Partida.Text + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    rstMovguia.Close
                    WMarcaEstado = ""
                    Sql1 = "UPDATE Guia SET "
                    Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
                    Sql3 = " Where Guia.Terminado = " + "'" + Producto.Text + "'"
                    Sql4 = " and Guia.Lote = " + "'" + Partida.Text + "'"
                    spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
            End If
                
            If Left$(Producto.Text, 2) = "PT" Then
                WArtiDev = "NK" + Right$(Producto.Text, 10)
                    Else
                If Left$(Producto.Text, 2) = "DY" Then
                    WArtiDev = "DK" + Right$(Producto.Text, 10)
                        Else
                    If Left$(Producto.Text, 2) = "DS" Then
                        WArtiDev = "NS" + Right$(Producto.Text, 10)
                            Else
                        WArtiDev = "NW" + Right$(Producto.Text, 10)
                    End If
                End If
            End If
            
            aa = Wempresa
            
            ZSql = ""
            ZSql = ZSql + "UPDATE EntDev SET "
            ZSql = ZSql + "Estado  = 'PT'"
            ZSql = ZSql + " Where EntDev.Terminado = " + "'" + WArtiDev + "'"
            ZSql = ZSql + " and EntDev.Lote = " + "'" + Partida.Text + "'"
            ZSql = ZSql + " and EntDev.Cliente = " + "'" + Cliente.Text + "'"
            spEntdev = ZSql
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
        
        
        
        
        
        
    
    End If
    
    
    
    
    
    
    Call CmdLimpiar_Click
    Producto.SetFocus
        
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -     -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Ensayo1.Caption = ""
    Valor1.Text = ""
    Ensayo2.Caption = ""
    Valor2.Text = ""
    Ensayo3.Caption = ""
    Valor3.Text = ""
    Ensayo4.Caption = ""
    Valor4.Text = ""
    Ensayo5.Caption = ""
    Valor5.Text = ""
    Ensayo6.Caption = ""
    Valor6.Text = ""
    Ensayo7.Caption = ""
    Valor7.Text = ""
    Ensayo8.Caption = ""
    Valor8.Text = ""
    Ensayo9.Caption = ""
    Valor9.Text = ""
    Ensayo10.Caption = ""
    Valor10.Text = ""
    Descri1.Caption = ""
    Descri2.Caption = ""
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
    Cantidad.Text = ""
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
    Observa.Text = ""
    Cantidad.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    TipoPro.Text = ""
    
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgLiberaTerminado.Hide
    Unload Me
    Menu.Show
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
    
        If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
    
            If Left$(Producto.Text, 2) = "DY" Then
                WArti = "DY-" + Right$(Producto.Text, 7)
                WArtiDev = "DK-" + Right$(Producto.Text, 7)
                    Else
                If Left$(Producto.Text, 2) = "DS" Then
                    WArti = "DS-" + Right$(Producto.Text, 7)
                    WArtiDev = "NS-" + Right$(Producto.Text, 7)
                        Else
                    WArti = "DW-" + Right$(Producto.Text, 7)
                    WArtiDev = "NW-" + Right$(Producto.Text, 7)
                End If
            End If
    
            WPasa = "S"
            WEntra = "N"
            WEstado = ""
        
            Sql1 = "Select *"
            Sql2 = " FROM Laudo"
            Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
            Sql4 = " and Laudo.PartiOri = " + "'" + Partida.Text + "'"
            spLaudo = Sql1 + Sql2 + Sql3 + Sql4
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstLaudo!Estado), "N", rstLaudo!Estado)
                rstLaudo.Close
            End If
                
            If WEntra = "N" Then
                Sql1 = "Select *"
                Sql2 = " FROM Guia"
                Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                Sql4 = " and Guia.PartiOri = " + "'" + Partida.Text + "'"
                spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WEntra = "S"
                    WEstado = IIf(IsNull(rstMovguia!Estado), "N", rstMovguia!Estado)
                    rstMovguia.Close
                End If
            End If
            
            Rem If WEntra = "S" And WEstado <> "N" Then
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EntDev"
            Rem     Sql3 = " Where EntDev.Terminado = " + "'" + WArtiDev + "'"
            Rem     Sql4 = " and EntDev.PartiOri = " + "'" + Partida.Text + "'"
            Rem     spEntdev = Sql1 + Sql2 + Sql3 + Sql4
            Rem     Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEntdev.RecordCount > 0 Then
            Rem         If rstEntdev!Saldo > 0 Then
            Rem             WEstado = "N"
            Rem         End If
            Rem         rstEntdev.Close
            Rem     End If
            Rem End If
            
            If WEntra = "N" Then
                m$ = "Codigo de Partida o Producto Invalido "
                a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                WPasa = "N"
                    Else
                If WEstado <> "N" Then
                    m$ = "La partida ya se encuentra liberada"
                    a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                    WPasa = "N"
                End If
            End If
            
                Else
                
            WPasa = "S"
            WEntra = "N"
            WEstado = ""
        
            XParam = "'" + Partida.Text + "','" _
                         + Producto.Text + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstHoja!Estado), "N", rstHoja!Estado)
                rstHoja.Close
            End If
                
            If WEntra = "N" Then
                XParam = "'" + Producto.Text + "','" _
                            + Partida.Text + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WEntra = "S"
                    WEstado = IIf(IsNull(rstMovguia!Estado), "N", rstMovguia!Estado)
                    rstMovguia.Close
                End If
            End If
            
            If WEntra = "S" And WEstado <> "N" Then
                Sql1 = "Select *"
                Sql2 = " FROM EntDev"
                Sql3 = " Where EntDev.Terminado = " + "'" + "NK" + Mid$(Producto.Text, 3, 10) + "'"
                Sql4 = " and EntDev.Lote = " + "'" + Partida.Text + "'"
                spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                    If rstEntdev!Saldo > 0 Then
                        WEstado = "N"
                    End If
                    rstEntdev.Close
                End If
            End If
        
            If WEntra = "N" Then
                m$ = "Codigo de Partida o Producto Invalido "
                a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                WPasa = "N"
                    Else
                If WEstado <> "N" Then
                    m$ = "La partida ya se encuentra liberada"
                    a% = MsgBox(m$, 0, "Liberacion de Partidas de Prodcuto Terminado")
                    WPasa = "N"
                End If
            End If
                
        End If
        
        If WPasa = "S" Then
            Cliente.SetFocus
        End If
        
    End If
    
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
            Cliente.Text = UCase(Cliente.Text)
            If Cliente.Text <> "" Then
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!razon
                    rstCliente.Close
                    Observa.SetFocus
                        Else
                    Cliente.Text = Claveven$
                    Cliente.SetFocus
                End If
            End If
                Else
            Observa.SetFocus
        End If
    End If
End Sub

Private Sub Observa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1.SetFocus
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


Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor2.SetFocus
    End If
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor3.SetFocus
    End If
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor4.SetFocus
    End If
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor5.SetFocus
    End If
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor6.SetFocus
    End If
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor7.SetFocus
    End If
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor8.SetFocus
    End If
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor9.SetFocus
    End If
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor10.SetFocus
    End If
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo.SetFocus
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
    
    If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
    
        If Left$(Producto.Text, 2) = "DY" Then
            WArti = "DY-" + Right$(Producto.Text, 7)
                Else
            If Left$(Producto.Text, 2) = "DS" Then
                WArti = "DS-" + Right$(Producto.Text, 7)
                    Else
                WArti = "DW-" + Right$(Producto.Text, 7)
            End If
        End If

        Sql1 = "Select *"
        Sql2 = " FROM EspecificacionesUnifica"
        Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + WArti + "'"
        spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnifica.RecordCount > 0 Then
            Rem Producto.Text = rstEspecifUnifica!Producto
            Ensayo1.Caption = rstEspecificacionesUnifica!Ensayo1
            Ensayo2.Caption = rstEspecificacionesUnifica!Ensayo2
            Ensayo3.Caption = rstEspecificacionesUnifica!Ensayo3
            Ensayo4.Caption = rstEspecificacionesUnifica!Ensayo4
            Ensayo5.Caption = rstEspecificacionesUnifica!Ensayo5
            Ensayo6.Caption = rstEspecificacionesUnifica!Ensayo6
            Ensayo7.Caption = rstEspecificacionesUnifica!Ensayo7
            Ensayo8.Caption = rstEspecificacionesUnifica!Ensayo8
            Ensayo9.Caption = rstEspecificacionesUnifica!Ensayo9
            Ensayo10.Caption = rstEspecificacionesUnifica!Ensayo10
            Std1.Caption = rstEspecificacionesUnifica!Valor1
            Std2.Caption = rstEspecificacionesUnifica!Valor2
            Std3.Caption = rstEspecificacionesUnifica!Valor3
            Std4.Caption = rstEspecificacionesUnifica!Valor4
            Std5.Caption = rstEspecificacionesUnifica!Valor5
            Std6.Caption = rstEspecificacionesUnifica!Valor6
            Std7.Caption = rstEspecificacionesUnifica!Valor7
            Std8.Caption = rstEspecificacionesUnifica!Valor8
            Std9.Caption = rstEspecificacionesUnifica!Valor9
            Std10.Caption = rstEspecificacionesUnifica!Valor10
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
            rstEspecificacionesUnifica.Close
        End If
    
        
            Else

        If Left$(Producto.Text, 2) = "DW" Then
            WProducto = "DW" + Mid$(Producto.Text, 3, 10)
                Else
            If Left$(Producto.Text, 2) = "SE" Then
                WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                    Else
                WProducto = "PT" + Mid$(Producto.Text, 3, 10)
            End If
        End If

        Sql1 = "Select *"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            Rem Producto.Text = rstEspecifUnifica!Producto
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
            Std1.Caption = rstEspecifUnifica!Valor1
            Std2.Caption = rstEspecifUnifica!Valor2
            Std3.Caption = rstEspecifUnifica!Valor3
            Std4.Caption = rstEspecifUnifica!Valor4
            Std5.Caption = rstEspecifUnifica!Valor5
            Std6.Caption = rstEspecifUnifica!Valor6
            Std7.Caption = rstEspecifUnifica!Valor7
            Std8.Caption = rstEspecifUnifica!Valor8
            Std9.Caption = rstEspecifUnifica!Valor9
            Std10.Caption = rstEspecifUnifica!Valor10
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
            rstEspecifUnifica.Close
        End If
        
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
    
    Call Conecta_Empresa

End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Producto.Text = UCase(Producto.Text)
        If Producto.Text <> "" Then
        
            If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "DS" Then
    
                If Left$(Producto.Text, 2) = "DY" Then
                    WProducto = Producto.Text
                    WArti = "DY-" + Right$(Producto.Text, 7)
                        Else
                    If Left$(Producto.Text, 2) = "DS" Then
                        WProducto = Producto.Text
                        WArti = "DS-" + Right$(Producto.Text, 7)
                            Else
                        WProducto = Producto.Text
                        WArti = "DW-" + Right$(Producto.Text, 7)
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
            
                Sql1 = "Select *"
                Sql2 = " FROM EspecificacionesUnifica"
                Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + WArti + "'"
                spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnifica.RecordCount > 0 Then
                    rstEspecificacionesUnifica.Close
                    Call Conecta_Empresa
                    Call imprime_Click
                        Else
                    Call Conecta_Empresa
                    CmdLimpiar_Click
                    Producto.Text = WProducto
                End If
            
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    rstArticulo.Close
                        Else
                    Producto.SetFocus
                    Exit Sub
                End If
        
                    Else
        
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
            
                Sql1 = "Select *"
                Sql2 = " FROM EspecifUnifica"
                Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
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
            
        End If
        Partida.SetFocus
    End If
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgLiberaTerminado.Caption = "Ingreso de Ensayos de Producto Terminado :  " + !Nombre
        End If
    End With
    EmpresaActual = Wempresa
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    TipoPro.Text = ""
End Sub




