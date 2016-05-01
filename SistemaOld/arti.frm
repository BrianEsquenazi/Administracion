VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgArti 
   Caption         =   "Ingreso de Materias Primas"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   Begin VB.CommandButton Command124 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10680
      TabIndex        =   243
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CargaTipo 
      Caption         =   "Etiqueta"
      Height          =   255
      Left            =   8880
      TabIndex        =   242
      Top             =   50
      Width           =   1095
   End
   Begin VB.TextBox TipoEti 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      MaxLength       =   10
      TabIndex        =   241
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Command1"
      Height          =   735
      Left            =   480
      TabIndex        =   238
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton LeeMsds 
      Caption         =   "Msds"
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
      Left            =   2160
      TabIndex        =   237
      Top             =   6600
      Width           =   975
   End
   Begin VB.CheckBox Restriccion 
      Caption         =   "Restriccion"
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
      TabIndex        =   236
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CheckBox DatosPrv 
      Caption         =   "Fec Vto y Elab Prv"
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
      TabIndex        =   235
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Frame PantaHistoria 
      Caption         =   "Historial"
      Height          =   1935
      Left            =   3240
      TabIndex        =   212
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox HistoriaOrden 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   218
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox HistoriaInforme 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   217
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox HistoriaRemito 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   216
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox HistoriaFactura 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   215
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton PantaHistoriaCierra 
         Caption         =   "Cierra"
         Height          =   495
         Left            =   2040
         TabIndex        =   214
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox HistoriaCarpeta 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   213
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin MSMask.MaskEdBox HistoriaFechaOrden 
         Height          =   285
         Left            =   3720
         TabIndex        =   219
         Top             =   360
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
      Begin MSMask.MaskEdBox HistoriaFechaInforme 
         Height          =   285
         Left            =   3720
         TabIndex        =   220
         Top             =   720
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
      Begin MSMask.MaskEdBox HistoriaFechaFactura 
         Height          =   285
         Left            =   3720
         TabIndex        =   221
         Top             =   1440
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
      Begin VB.Label DesproveedorH 
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
         TabIndex        =   228
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label60 
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
         TabIndex        =   227
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label59 
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
         TabIndex        =   226
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "Informe de Recepcion"
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
         TabIndex        =   225
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label57 
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
         Left            =   120
         TabIndex        =   224
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label52 
         Caption         =   "Factura"
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
         TabIndex        =   223
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label51 
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
         Left            =   120
         TabIndex        =   222
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame Clave 
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
      Height          =   1815
      Left            =   2280
      TabIndex        =   58
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
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
         TabIndex        =   61
         Top             =   1200
         Width           =   1935
      End
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
         TabIndex        =   60
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label16 
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
         Left            =   960
         TabIndex        =   59
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2760
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   4695
      Begin MSMask.MaskEdBox Desdecodigo 
         Height          =   285
         Left            =   1800
         TabIndex        =   45
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HastaCodigo 
         Height          =   285
         Left            =   1800
         TabIndex        =   46
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2520
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
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
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   360
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
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Panord 
      Caption         =   "Control Fecha"
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
      Left            =   3480
      TabIndex        =   75
      Top             =   2760
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Acepta1 
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
         Height          =   255
         Left            =   3480
         TabIndex        =   81
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Cancela1 
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
         Height          =   255
         Left            =   3480
         TabIndex        =   80
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
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
         Height          =   375
         Left            =   1080
         TabIndex        =   79
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Impresora"
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
         Left            =   2640
         TabIndex        =   78
         Top             =   1200
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   76
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   77
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label24 
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
         Height          =   375
         Left            =   360
         TabIndex        =   83
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
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
         Left            =   360
         TabIndex        =   82
         Top             =   720
         Width           =   1575
      End
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
      Height          =   2790
      ItemData        =   "arti.frx":0000
      Left            =   3480
      List            =   "arti.frx":0007
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   4695
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
      Left            =   2760
      TabIndex        =   70
      Top             =   3000
      Visible         =   0   'False
      Width           =   4695
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
      ItemData        =   "arti.frx":0015
      Left            =   3480
      List            =   "arti.frx":0017
      TabIndex        =   48
      Top             =   3960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame PantaEtiDy 
      Caption         =   "Datos Etiqueta DY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   142
      Top             =   720
      Visible         =   0   'False
      Width           =   7815
      Begin VB.ComboBox TipoBarra 
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
         TabIndex        =   162
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton EtiCancela 
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
         Height          =   615
         Left            =   5760
         TabIndex        =   159
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton EtiImpre2 
         Caption         =   "Codigo Barra"
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
         TabIndex        =   158
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton EtiImpre1 
         Caption         =   "Etiqueta"
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
         Left            =   840
         TabIndex        =   157
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox EtiCantidadEti 
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
         MaxLength       =   10
         TabIndex        =   155
         Text            =   " "
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox EtiCantidad 
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
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   153
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox EtiArticulo 
         BeginProperty Font 
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
         TabIndex        =   152
         Text            =   " "
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox EtiDescri6 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   150
         Text            =   " "
         Top             =   2760
         Width           =   6495
      End
      Begin VB.TextBox EtiDescri5 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   149
         Text            =   " "
         Top             =   2400
         Width           =   6495
      End
      Begin VB.TextBox EtiDescri4 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   148
         Text            =   " "
         Top             =   2040
         Width           =   6495
      End
      Begin VB.TextBox EtiDescri3 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   147
         Text            =   " "
         Top             =   1680
         Width           =   6495
      End
      Begin VB.TextBox EtiDescri2 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   146
         Text            =   " "
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox EtiDescri1 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   145
         Text            =   " "
         Top             =   960
         Width           =   6495
      End
      Begin VB.TextBox EtiPartida 
         BeginProperty Font 
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
         TabIndex        =   143
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label42 
         Caption         =   "Tipo Barra"
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
         TabIndex        =   161
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label43 
         Caption         =   "Etiquetas"
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
         TabIndex        =   156
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label zscfzc 
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
         Left            =   3360
         TabIndex        =   154
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "Articulo Nro"
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
         TabIndex        =   151
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label40 
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
         Left            =   360
         TabIndex        =   144
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame PantaMensaje 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   3600
      TabIndex        =   229
      Top             =   2880
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox MensajeII 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   231
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox MensajeI 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   230
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame ConsultaMarcas 
      Caption         =   "Consulta de Marcas"
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
      Left            =   11880
      TabIndex        =   90
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton GrabaMarcas 
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
         Height          =   375
         Left            =   4920
         TabIndex        =   109
         Top             =   4200
         Width           =   1695
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   2040
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   2040
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox WTexto2 
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
         Left            =   2160
         TabIndex        =   103
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   102
         Top             =   2040
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
         Left            =   2760
         TabIndex        =   101
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton FinConsulta 
         Caption         =   "Fin Consulta"
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
         Left            =   2520
         TabIndex        =   93
         Top             =   4200
         Width           =   1695
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   1560
         TabIndex        =   107
         Top             =   1440
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
         Height          =   3495
         Left            =   360
         TabIndex        =   108
         Top             =   720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6165
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label WCampo2 
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
         Left            =   2760
         TabIndex        =   94
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label WCampo1 
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
         Left            =   1200
         TabIndex        =   92
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label26 
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
         TabIndex        =   91
         Top             =   360
         Width           =   855
      End
   End
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
      Height          =   1455
      Left            =   3720
      Picture         =   "arti.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   120
      TabIndex        =   203
      Top             =   2520
      Width           =   2055
      Begin VB.TextBox Costo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   205
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label53 
         Caption         =   "Costo Std U$S"
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
         TabIndex        =   204
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox ZCosto2 
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
      Left            =   12120
      MaxLength       =   10
      TabIndex        =   198
      Text            =   " "
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox ZCostoAnterior 
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
      Left            =   12120
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   197
      Text            =   " "
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox WCosto2 
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
      Left            =   12120
      MaxLength       =   10
      TabIndex        =   196
      Text            =   " "
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox CostoAnterior 
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
      Left            =   12120
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   195
      Text            =   " "
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "Compras Locales"
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
      Left            =   2280
      TabIndex        =   180
      Top             =   3480
      Width           =   9375
      Begin VB.TextBox WCosto1Dol 
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
         MaxLength       =   10
         TabIndex        =   207
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox WCosto1 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   186
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox ZCosto1 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   185
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox OrdenII 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   184
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox OrdenIII 
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   183
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox FechaOrdenII 
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   182
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox FechaOrdenIII 
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   181
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label55 
         Caption         =   "F.Laudo"
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
         TabIndex        =   192
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Costo Ult. Compra $"
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
         TabIndex        =   191
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label47 
         Caption         =   "Costo Ult. Compra U$S"
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
         TabIndex        =   190
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label TituloOrdenII 
         Caption         =   "Orden"
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
         Left            =   4920
         TabIndex        =   189
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label TituloOrdenIII 
         Caption         =   "Orden"
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
         Left            =   4920
         TabIndex        =   188
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label54 
         Caption         =   "F.Laudo"
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
         TabIndex        =   187
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame CargaCuadro 
      Height          =   1095
      Left            =   3000
      TabIndex        =   139
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox PartidaCuadro 
         BeginProperty Font 
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
         TabIndex        =   140
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label39 
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
         Left            =   600
         TabIndex        =   141
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame CargaPartida 
      Height          =   1095
      Left            =   4560
      TabIndex        =   135
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   136
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label38 
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
         Left            =   600
         TabIndex        =   137
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton AvisoErrorII 
      Caption         =   "No se puede ejecutar el procedimiento. Sistema sin Conexion con las otras plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3480
      Picture         =   "arti.frx":075B
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importaciones"
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
      Left            =   2280
      TabIndex        =   171
      Top             =   2520
      Width           =   9375
      Begin VB.TextBox Costo6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   210
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Costo4 
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   208
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox FechaOrdenI 
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   193
         Top             =   600
         Width           =   1215
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
         Left            =   1080
         TabIndex        =   177
         Top             =   240
         Width           =   1095
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
         Left            =   2280
         TabIndex        =   176
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Moneda 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         TabIndex        =   175
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox OrdenI 
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
         MaxLength       =   10
         TabIndex        =   174
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Costo1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   172
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label50 
         Caption         =   "Reposicion"
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
         TabIndex        =   209
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label49 
         Caption         =   "Costo En Transito"
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
         Left            =   3600
         TabIndex        =   206
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label56 
         Caption         =   "F. Laudo"
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
         TabIndex        =   194
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label TituloOrdenI 
         Caption         =   "Orden"
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
         Left            =   4560
         TabIndex        =   178
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Costo Ult. Compra "
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
         TabIndex        =   173
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.TextBox parance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   170
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton EtiquetaDy 
      Caption         =   "Etiqueta Dy"
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
      Left            =   1560
      TabIndex        =   160
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Derechos 
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
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   132
      Text            =   " "
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox TipoMp 
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
      Left            =   5160
      TabIndex        =   130
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ComboBox Sedronar 
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
      Left            =   2280
      TabIndex        =   128
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox Reventa 
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
      Left            =   4560
      TabIndex        =   127
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Meses 
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
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   124
      Text            =   " "
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton GrabaMinimo 
      Caption         =   "Graba Minimo Planta"
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
      Left            =   5520
      TabIndex        =   123
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox CodigoDy 
      BeginProperty Font 
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
      TabIndex        =   121
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
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
      Left            =   3240
      TabIndex        =   19
      Top             =   6480
      Width           =   3615
      Begin VB.CommandButton Anterior 
         Caption         =   "Registro Anterior"
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
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Registro Siguiente"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Registro"
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
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Registro"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   7080
      TabIndex        =   112
      Top             =   4440
      Width           =   4575
      Begin VB.TextBox Caracteristicas 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   239
         Text            =   " "
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Embalaje 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   116
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Naciones 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   115
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Intervencion 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   114
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Clase 
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   113
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label62 
         Caption         =   "Caracteristicas"
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
         Height          =   300
         Left            =   240
         TabIndex        =   240
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "Grupo Embalaje"
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
         TabIndex        =   120
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label29 
         Caption         =   "Nro. N.Unidas"
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
         TabIndex        =   119
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "F.Intervencion"
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
         TabIndex        =   118
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Clase"
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
         TabIndex        =   117
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton PedPen 
      Caption         =   " Pedidos Ventas"
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
      Left            =   10320
      TabIndex        =   100
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton LaboPendiente 
      Caption         =   "Pend. Laboratorio"
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
      Left            =   10200
      TabIndex        =   97
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton ocpend 
      Caption         =   "O/C Pendientes"
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
      TabIndex        =   96
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton ComMarcas 
      Caption         =   "   Consulta de        Marcas"
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
      Left            =   8640
      TabIndex        =   95
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton ConCoti1 
      Caption         =   "Consulta de Cotizaciones U$S"
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
      Left            =   7080
      TabIndex        =   89
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox WCosto3 
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
      Left            =   11880
      MaxLength       =   10
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton ConCpa 
      Caption         =   "Consulta de Compras"
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
      Left            =   8640
      TabIndex        =   74
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton ConCoti 
      Caption         =   "Consulta de Cotizaciones $"
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
      Left            =   7080
      TabIndex        =   73
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Costo3 
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
      Left            =   11880
      MaxLength       =   10
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Densidad 
      BeginProperty Font 
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
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox Controla 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame StockCons 
      Caption         =   "Stock Consolidado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   8400
      TabIndex        =   62
      Top             =   720
      Width           =   3255
      Begin VB.Label WStock7 
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
         Left            =   2280
         TabIndex        =   166
         Top             =   720
         Width           =   900
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1680
         TabIndex        =   165
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label WStock6 
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
         Left            =   2280
         TabIndex        =   164
         Top             =   480
         Width           =   900
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1680
         TabIndex        =   163
         Top             =   480
         Width           =   1095
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1680
         TabIndex        =   87
         Top             =   240
         Width           =   495
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
         Left            =   2280
         TabIndex        =   86
         Top             =   240
         Width           =   900
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   85
         Top             =   960
         Width           =   615
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
         Left            =   720
         TabIndex        =   84
         Top             =   960
         Width           =   900
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
         Left            =   720
         TabIndex        =   68
         Top             =   720
         Width           =   900
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
         Left            =   720
         TabIndex        =   67
         Top             =   480
         Width           =   900
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
         Left            =   720
         TabIndex        =   66
         Top             =   240
         Width           =   900
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   1095
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   1095
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Proveedor 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   12
      Text            =   " "
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Rs 
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
      Left            =   480
      MaxLength       =   2
      TabIndex        =   8
      Text            =   " "
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Envase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   11
      Text            =   " "
      Top             =   5040
      Width           =   615
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.TextBox Deposito 
      BeginProperty Font 
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
      MaxLength       =   20
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Unidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Minimo 
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
      MaxLength       =   10
      TabIndex        =   5
      Text            =   " "
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Salidas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Height          =   285
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   44
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Entradas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   43
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Inicial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   42
      Text            =   " "
      Top             =   720
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11520
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WArticulo.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   11280
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton lista 
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
      Height          =   300
      Left            =   8160
      TabIndex        =   25
      Top             =   6480
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
      Height          =   300
      Left            =   7080
      TabIndex        =   24
      Top             =   6480
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
      Height          =   300
      Left            =   9240
      TabIndex        =   13
      Top             =   6480
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
      Height          =   300
      Left            =   8160
      TabIndex        =   18
      Top             =   6120
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
      Height          =   300
      Left            =   9240
      TabIndex        =   17
      Top             =   6120
      Width           =   975
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
      Height          =   300
      Left            =   7080
      TabIndex        =   16
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
      BeginProperty Font 
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
      MaxLength       =   40
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox Minimo1 
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
      Left            =   7200
      MaxLength       =   10
      TabIndex        =   6
      Text            =   " "
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Historial 
      Caption         =   "Historial"
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
      Left            =   4200
      TabIndex        =   134
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Cuadro 
      Caption         =   "Cuadro"
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
      Left            =   2880
      TabIndex        =   138
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox responsa 
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
      Left            =   4800
      TabIndex        =   167
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton DatosComple 
      Caption         =   "Datos Imprtacion"
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
      TabIndex        =   232
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox CodSedronar 
      BeginProperty Font 
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
      TabIndex        =   233
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label61 
      Caption         =   "Cod. Sedronar"
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
      TabIndex        =   234
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Rs"
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
      TabIndex        =   211
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label48 
      Caption         =   "Costo Estandar U$S"
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
      Left            =   12120
      TabIndex        =   202
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label46 
      Caption         =   "Costo Anterior"
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
      Left            =   12120
      TabIndex        =   201
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "Costo Estandar $"
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
      Left            =   12120
      TabIndex        =   200
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label37 
      Caption         =   "Costo Anterior"
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
      Left            =   12120
      TabIndex        =   199
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TituloStd 
      Caption         =   "Costo Estandar U$S"
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
      Left            =   -2280
      TabIndex        =   179
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label45 
      Caption         =   "P.Aranc."
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
      Left            =   7200
      TabIndex        =   169
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label36 
      Caption         =   "% Der."
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
      Left            =   10080
      TabIndex        =   133
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "Tipo M.P."
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
      Left            =   4200
      TabIndex        =   131
      Top             =   5775
      Width           =   975
   End
   Begin VB.Label Label34 
      Caption         =   "Incluye Sedronar"
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
      TabIndex        =   129
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label33 
      Caption         =   "Reventa"
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
      TabIndex        =   126
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "Meses Vencimiento"
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
      TabIndex        =   125
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label31 
      Caption         =   "Codigo Prv."
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
      Left            =   7800
      TabIndex        =   122
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Pedidos Pendientes"
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
      TabIndex        =   99
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Venta 
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
      Left            =   5760
      TabIndex        =   98
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "Costo Promedio $"
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
      Left            =   11880
      TabIndex        =   88
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Costo Promedio U$S"
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
      Left            =   12000
      TabIndex        =   72
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Densidad"
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
      TabIndex        =   71
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Controla Lote"
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
      Left            =   3960
      TabIndex        =   69
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label DesProveedor 
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
      Left            =   4080
      TabIndex        =   57
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor Ult. Cpa."
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
      TabIndex        =   56
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Pedido 
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
      Left            =   2280
      TabIndex        =   55
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Laboratorio 
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
      Left            =   2280
      TabIndex        =   54
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label DesEnvase 
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
      Left            =   3000
      TabIndex        =   53
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label14 
      Caption         =   "Rs"
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
      TabIndex        =   52
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "O/C Pendientes"
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
      TabIndex        =   50
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label aa 
      Caption         =   "Stock Laboratorio"
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
      Top             =   1800
      Width           =   1695
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
      Left            =   2280
      TabIndex        =   47
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Cantidad Final"
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
      TabIndex        =   41
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Deposito"
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
      Left            =   3960
      TabIndex        =   40
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Unidad de Medida"
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
      TabIndex        =   39
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Minimo Consol./Planta"
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
      TabIndex        =   38
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad Salida"
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
      Left            =   3960
      TabIndex        =   37
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad Entrada"
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
      TabIndex        =   36
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad Inicial"
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
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label Label44 
      Caption         =   "Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3480
      TabIndex        =   168
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private WGraba As String
Private WProceso As String
Private Auxi As String
Private WAuxi As String
Private XVector(3, 5) As String
Dim Empe(12, 2) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim Vector(10000, 10) As String
Dim Vector1(10000, 4) As Double
Dim XCosto1 As String
Dim XCosto2 As String
Dim XCosto3 As String
Dim XCosto4 As String


Dim ZZCosto6 As Double
Dim ZZCosto1 As Double
Dim ZZPtaOrdenI As Integer
Dim ZZPtaOrdenII As Integer
Dim ZZPtaOrdenIII As Integer

Dim rstCambios As Recordset
Dim spCambios As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstPeligroso As Recordset
Dim spPeligroso As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String

Dim Paridad As Double
Dim ParidadII As Double
Dim WWVector(10000, 4) As String
Dim WWVectorII(1000, 12) As String
Dim XProveedor As String
Dim CargaEmpresa(12, 2) As String
Dim ZCostoCompara As Double
Dim ZCostoActual As Double

Dim ZCampo1 As String
Dim ZCampo2 As String
Dim ZCampo3 As String
Dim ZCampo4 As String
Dim ZCampo5 As String
Dim ZCampo6 As String
Dim ZCampo7 As String
Dim ZCampo8 As String
Dim ZCampo9 As String
Dim ZCampo10 As String
Dim ZCampo11 As String
Dim ZCampo12 As String
Dim ZCampo13 As String
Dim ZCampo14 As String
Dim ZCampo15 As String
Dim ZCampo16 As String
Dim ZCampo17 As String
Dim ZCampo18 As String
Dim ZCampo19 As String
Dim ZCampo20 As String
Dim ZCampo21 As String
Dim ZCampo22 As String
Dim ZCampo23 As String
Dim ZCampo24 As String
Dim PasaLeyenda As String
Dim ZAyuda As String

Dim ZTipoCosto As Integer
Private XLote(100, 7) As String

Dim ZZLote(100) As String
Dim ZZCanti(100) As String

Dim WVerifica(200) As String
Dim ZZZZImpre(100) As String
Dim ZZZZOrden As String
Dim ZZZZLaudo As String
Dim ZZZZInforme As String
Dim ZZZZEmpresa As String
Dim ZZZZPtaOrden As Integer

Dim ZZRestriccion As Integer
Dim ZZDatosPrv As Integer
Dim WDatosPrv As String
Dim WRestriccion As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Sub Imprime_Datos()

    spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Codigo.Text = rstArticulo!Codigo
        Descripcion.Text = Trim(rstArticulo!Descripcion)
        Unidad.Text = rstArticulo!Unidad
        Deposito.Text = rstArticulo!Deposito
        Inicial.Text = Str$(rstArticulo!Inicial)
        Entradas.Text = Str$(rstArticulo!Entradas)
        Salidas.Text = Str$(rstArticulo!Salidas)
        Minimo.Text = Str$(rstArticulo!Minimo)
        Minimo1.Text = IIf(IsNull(rstArticulo!Minimo1), "0", rstArticulo!Minimo1)
        Rem Laboratorio.Caption = Str$(rstArticulo!Laboratorio)
        Rem Pedido.Caption = Str$(rstArticulo!Pedido)
        Venta.Caption = IIf(IsNull(rstArticulo!Venta), "0", rstArticulo!Venta)
        Envase.Text = rstArticulo!Envase
        
        Costo1.Text = Str$(rstArticulo!Costo1)
        
        ZZCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
        WCosto1.Text = Str$(ZZCosto1)
        
        ZZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
        ZCosto1.Text = Str$(ZZCosto1)
        
        ZZCosto1 = IIf(IsNull(rstArticulo!Costo6), "0", rstArticulo!Costo6)
        Costo6.Text = Str$(ZZCosto1)
        
        Costo2.Text = Str$(rstArticulo!Costo2)
        WCosto2.Text = IIf(IsNull(rstArticulo!WCosto2), "0", rstArticulo!WCosto2)
        
        Costo3.Text = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
        WCosto3.Text = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
        
        Costo4.Text = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
        
        OrdenI.Text = IIf(IsNull(rstArticulo!OrdenI), "", rstArticulo!OrdenI)
        OrdenII.Text = IIf(IsNull(rstArticulo!OrdenII), "", rstArticulo!OrdenII)
        OrdenIII.Text = IIf(IsNull(rstArticulo!OrdenIII), "", rstArticulo!OrdenIII)
        ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
        ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
        ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
        
        Rs.Text = rstArticulo!Rs
        Flete.Text = Str$(rstArticulo!Flete)
        Moneda.Text = rstArticulo!Moneda
        Controla.ListIndex = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
        Reventa.ListIndex = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
        Sedronar.ListIndex = IIf(IsNull(rstArticulo!Sedronar), "0", rstArticulo!Sedronar)
        TipoMp.ListIndex = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
        CodSedronar.Text = IIf(IsNull(rstArticulo!CodSedronar), "", rstArticulo!CodSedronar)
        Densidad.Text = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
        TipoEti.Text = IIf(IsNull(rstArticulo!TipoEti), "", rstArticulo!TipoEti)
        CodigoDy.Text = IIf(IsNull(rstArticulo!CodigoDy), "", rstArticulo!CodigoDy)
        Leyenda.ListIndex = IIf(IsNull(rstArticulo!Leyenda), "0", rstArticulo!Leyenda)
        If Leyenda.ListIndex = 0 And Val(Flete.Text) = 0 Then
            Leyenda.ListIndex = 6
        End If
        Clase.Text = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
        Intervencion.Text = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
        Caracteristicas.Text = IIf(IsNull(rstArticulo!Descrionu), "", rstArticulo!Descrionu)
        Naciones.Text = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
        Embalaje.Text = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
        Meses.Text = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
        WDerechos = IIf(IsNull(rstArticulo!Derechos), "0", rstArticulo!Derechos)
        Derechos.Text = Str$(WDerechos)
        Derechos.Text = Pusing("###,###.##", Derechos.Text)
        CostoAnterior.Text = IIf(IsNull(rstArticulo!Costo2Anterior), "0", rstArticulo!Costo2Anterior)
        CostoAnterior.Text = Pusing("###,###.##", CostoAnterior.Text)
        Rem by nan
        parance.Text = IIf(IsNull(rstArticulo!Posarance), "0", rstArticulo!Posarance)
        Clase.Text = Trim(Clase.Text)
        Intervencion.Text = Trim(Intervencion.Text)
        Naciones.Text = Trim(Naciones.Text)
        Embalaje.Text = Trim(Embalaje.Text)
        ZZRestriccion = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
        ZZDatosPrv = IIf(IsNull(rstArticulo!DatosPrv), "0", rstArticulo!DatosPrv)
        Restriccion.Value = ZZRestriccion
        DatosPrv.Value = ZZDatosPrv
                    
        
        
        ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
        If ZTipoCosto = 1 Then
            TituloStd.Caption = "Std.  Estimado U$S"
                Else
            TituloStd.Caption = "Costo Estandar U$S"
        End If
        
        If rstArticulo!Proveedor <> "" Then
            Proveedor.Text = rstArticulo!Proveedor
                Else
            Proveedor.Text = ""
        End If
        
        rstArticulo.Close
        
        
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        responsa.Text = ""
        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            responsa.Text = IIf(IsNull(rstArticulo!Responsable), "", rstArticulo!Responsable)
            rstArticulo.Close
        End If
        
        Call Conecta_Empresa
        
        FechaOrdenI.Text = ""
        FechaOrdenII.Text = ""
        FechaOrdenIII.Text = ""
        
        TituloOrdenI.Caption = ""
        TituloOrdenII.Caption = ""
        TituloOrdenIII.Caption = ""
        
        ZZProveedor = 0
        ZZEnvase = 0
        ZZMoneda = ""
        ZZLeyenda = ""
        ZZFlete = 0
        
        ZZLaboratorioI = 0
        ZZLaboratorioII = 0
        ZZLaboratorioIII = 0
        ZZLaboratorioIV = 0
        ZZLaboratorioV = 0
        ZZLaboratorioVI = 0
        ZZLaboratorioVII = 0
        
        ZZPedidoI = 0
        ZZPedidoII = 0
        ZZPedidoIII = 0
        ZZPedidoIV = 0
        ZZPedidoV = 0
        ZZPedidoVI = 0
        ZZPedidoVII = 0
        
        If ZZPtaOrdenI <> 0 And Val(OrdenI.Text) <> 0 Then
        
            ZZImpre = ""
            
            Select Case ZZPtaOrdenI
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenI.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                TituloOrdenI.Caption = "Orden " + ZZImpre
                FechaOrdenI.Text = rstOrden!Fecha
                ZZProveedorI = rstOrden!Proveedor
                Flete.Text = Str$(rstOrden!Precio)
                Select Case rstOrden!Moneda
                    Case 0
                        Moneda.Text = "U$S"
                    Case 1
                        Moneda.Text = "$"
                    Case Else
                        Moneda.Text = "Eur"
                End Select
                If rstOrden!Leyenda > 0 Then
                    Leyenda.ListIndex = rstOrden!Leyenda - 1
                End If
                rstOrden.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenI.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                ZZEnvaseI = rstInforme!Envase
                rstInforme.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenI.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                If rstLaudo!Liberada <> 0 Then
                    FechaOrdenI.Text = rstLaudo!Fecha
                    ZZFechaOrdenI = rstLaudo!Fecha
                End If
                rstLaudo.Close
            End If
            
            
            Call Conecta_Empresa
            
        End If
        
        
        If ZZPtaOrdenII <> 0 And Val(OrdenII.Text) <> 0 Then
            
            Select Case ZZPtaOrdenII
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenII.Text + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                TituloOrdenII.Caption = "Orden " + ZZImpre
                FechaOrdenII.Text = rstOrden!Fecha
                ZZProveedorII = rstOrden!Proveedor
                rstOrden.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenII.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                ZZEnvaseII = rstInforme!Envase
                rstInforme.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenII.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                If rstLaudo!Liberada <> 0 Then
                    FechaOrdenII.Text = rstLaudo!Fecha
                    ZZFechaOrdenII = rstLaudo!Fecha
                End If
                rstLaudo.Close
            End If
            
            Call Conecta_Empresa
            
        End If
        
        
        
        If ZZPtaOrdenIII <> 0 And Val(OrdenIII.Text) <> 0 Then
            
            Select Case ZZPtaOrdenIII
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenIII.Text + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                TituloOrdenIII.Caption = "Orden " + ZZImpre
                FechaOrdenIII.Text = rstOrden!Fecha
                ZZProveedorIII = rstOrden!Proveedor
                rstOrden.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenIII.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                ZZEnvaseIII = rstInforme!Envase
                rstInforme.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Orden = " + "'" + OrdenIII.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + Codigo.Text + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                If rstLaudo!Liberada <> 0 Then
                    FechaOrdenIII.Text = rstLaudo!Fecha
                    ZZFechaOrdenIII = rstLaudo!Fecha
                End If
                rstLaudo.Close
            End If
            
            Call Conecta_Empresa
            
        End If
        
        Rem DADA
        Rem spCambio = "ConsultaCambio " + "'" + ZZFecha + "'"
        Rem Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCambio.RecordCount > 0 Then
        Rem     ZZParidad = rstCambio!Cambio
        Rem     rstCambio.Close
        Rem End If
        
        Rem If ZZParidad <> 0 Then
        Rem     ZZCostoPesos = WPrecio * ZZParidad
        Rem End If
                
        Rem If FechaOrdenI.Text <> "" Then
        Rem     WFechaOrdI = Right$(FechaOrdenI.Text, 4) + Mid$(FechaOrdenI.Text, 4, 2) + Left$(FechaOrdenI.Text, 2)
        Rem         Else
        Rem     WFechaOrdI = ""
        Rem End If
        Rem If FechaOrdenII.Text <> "" Then
        Rem     WFechaOrdII = Right$(FechaOrdenII.Text, 4) + Mid$(FechaOrdenII.Text, 4, 2) + Left$(FechaOrdenII.Text, 2)
        Rem         Else
        Rem     WFechaOrdII = ""
        Rem End If
        Rem If FechaOrdenIII.Text <> "" Then
        Rem     WFechaOrdIII = Right$(FechaOrdenIII.Text, 4) + Mid$(FechaOrdenIII.Text, 4, 2) + Left$(FechaOrdenIII.Text, 2)
        Rem         Else
        Rem WFechaOrdIII = ""
        Rem End If
        
        If ZZFechaOrdenI <> "" Then
            WFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                Else
            WFechaOrdI = ""
        End If
        If ZZFechaOrdenII <> "" Then
            WFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                Else
            WFechaOrdII = ""
        End If
        If ZZFechaOrdenIII <> "" Then
            WFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                Else
            WFechaOrdIII = ""
        End If
        
        
        If WFechaOrdI <> "" And WFechaOrdI > WFechaOrdII And WFechaOrdI > WFechaOrdIII Then
            FechaOrdenI.BackColor = &HFFFF&
            OrdenI.BackColor = &HFFFF&
            Costo1.BackColor = &HFFFF&
            ZZProveedor = ZZProveedorI
            ZZEnvase = ZZEnvaseI
                Else
            FechaOrdenI.BackColor = &H80000005
            OrdenI.BackColor = &H80000005
            Costo1.BackColor = &H80000005
        End If
        
        If WFechaOrdII <> "" And WFechaOrdII > WFechaOrdI And WFechaOrdII > WFechaOrdIII Then
            FechaOrdenII.BackColor = &HFFFF&
            OrdenII.BackColor = &HFFFF&
            WCosto1.BackColor = &HFFFF&
            WCosto1Dol.BackColor = &HFFFF&
            ZZProveedor = ZZProveedorII
            ZZEnvase = ZZEnvaseII
                Else
            FechaOrdenII.BackColor = &H80000005
            OrdenII.BackColor = &H80000005
            WCosto1.BackColor = &H80000005
            WCosto1Dol.BackColor = &H80000005
        End If
        
        If WFechaOrdIII <> "" And WFechaOrdIII > WFechaOrdI And WFechaOrdIII > WFechaOrdII Then
            FechaOrdenIII.BackColor = &HFFFF&
            OrdenIII.BackColor = &HFFFF&
            ZCosto1.BackColor = &HFFFF&
            ZZProveedor = ZZProveedorIII
            ZZEnvase = ZZEnvaseIII
                Else
            FechaOrdenIII.BackColor = &H80000005
            OrdenIII.BackColor = &H80000005
            ZCosto1.BackColor = &H80000005
        End If
        
        If ZZProveedor <> 0 Then
            Proveedor.Text = ZZProveedor
            Envase.Text = ZZEnvase
        End If
        
        Call Format_datos
        Call Imprime_Descripcion
    
    End If
    
    WSalidaError = ""
    On Error GoTo Control_Error
    XEmpresa = Wempresa
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioI = rstArticulo!Laboratorio
                ZZPedidoI = rstArticulo!Pedido
                WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock1.Caption = "0"
            End If
        
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioII = rstArticulo!Laboratorio
                ZZPedidoII = rstArticulo!Pedido
                WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioIII = rstArticulo!Laboratorio
                ZZPedidoIII = rstArticulo!Pedido
                WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock3.Caption = "0"
            End If
    
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioIV = rstArticulo!Laboratorio
                ZZPedidoIV = rstArticulo!Pedido
                WStock4.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock4.Caption = "0"
            End If
    
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioV = rstArticulo!Laboratorio
                ZZPedidoV = rstArticulo!Pedido
                WStock5.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock5.Caption = "0"
            End If
    
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioVI = rstArticulo!Laboratorio
                ZZPedidoVI = rstArticulo!Pedido
                WStock6.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock6.Caption = "0"
            End If
    
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioVII = rstArticulo!Laboratorio
                ZZPedidoVII = rstArticulo!Pedido
                WStock7.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock7.Caption = "0"
            End If
    
        Case 2, 4, 8, 9
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioI = rstArticulo!Laboratorio
                ZZPedidoI = rstArticulo!Pedido
                WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                      Else
                WStock1.Caption = "0"
            End If
    
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioII = rstArticulo!Laboratorio
                ZZPedidoII = rstArticulo!Pedido
                WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioIII = rstArticulo!Laboratorio
                ZZPedidoIII = rstArticulo!Pedido
                WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock3.Caption = "0"
            End If
            
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLaboratorioIV = rstArticulo!Laboratorio
                ZZPedidoIV = rstArticulo!Pedido
                WStock4.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock4.Caption = "0"
            End If
    
        Case Else
    End Select
    
    On Error GoTo 0
    Call Conecta_Empresa
    
    WCosto1Dol.Text = ""
    If Val(WCosto1.Text) <> 0 Then
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        spCambios = "ConsultaCambio  " + "'" + FechaOrdenII.Text + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZZParidad = rstCambios!Cambio
            rstCambios.Close
            If ZZParidad <> 0 Then
                ZZCosto1Dol = Val(WCosto1.Text) / ZZParidad
                WCosto1Dol.Text = Str$(ZZCosto1Dol)
                WCosto1Dol.Text = Pusing("###,###.##", WCosto1Dol.Text)
            End If
        End If
        
        Call Conecta_Empresa
        
    End If
    
    
    
    
    
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Select Case Val(Wempresa)
                Case 1
                    Laboratorio.Caption = Str$(ZZLaboratorioI + ZZLaboratorioII + ZZLaboratorioIII + ZZLaboratorioIV + ZZLaboratorioV + ZZLaboratorioVI + ZZLaboratorioVII)
                    Pedido.Caption = Str$(ZZPedidoI + ZZPedidoII + ZZPedidoIII + ZZPedidoIV + ZZPedidoV + ZZPedidoVI + ZZPedidoVII)
                Case 3
                    Laboratorio.Caption = Str$(ZZLaboratorioII)
                    Pedido.Caption = Str$(ZZPedidoII)
                Case 5
                    Laboratorio.Caption = Str$(ZZLaboratorioIII)
                    Pedido.Caption = Str$(ZZPedidoIII)
                Case 6
                    Laboratorio.Caption = Str$(ZZLaboratorioIV)
                    Pedido.Caption = Str$(ZZPedidoIV)
                Case 7
                    Laboratorio.Caption = Str$(ZZLaboratorioV)
                    Pedido.Caption = Str$(ZZPedidoV)
                Case 10
                    Laboratorio.Caption = Str$(ZZLaboratorioVI)
                    Pedido.Caption = Str$(ZZPedidoVI)
                Case 11
                    Laboratorio.Caption = Str$(ZZLaboratorioVII)
                    Pedido.Caption = Str$(ZZPedidoVII)
                Case Else
            End Select
                
    
        Case 2, 4, 8, 9
            Select Case Val(Wempresa)
                Case 2
                    Laboratorio.Caption = Str$(ZZLaboratorioI + ZZLaboratorioII + ZZLaboratorioIII + ZZLaboratorioIV)
                    Pedido.Caption = Str$(ZZPedidoI + ZZPedidoII + ZZPedidoIII + ZZPedidoIV)
                Case 4
                    Laboratorio.Caption = Str$(ZZLaboratorioII)
                    Pedido.Caption = Str$(ZZPedidoII)
                Case 8
                    Laboratorio.Caption = Str$(ZZLaboratorioIII)
                    Pedido.Caption = Str$(ZZPedidoIII)
                Case 9
                    Laboratorio.Caption = Str$(ZZLaboratorioIV)
                    Pedido.Caption = Str$(ZZPedidoIV)
                Case Else
            End Select
    
        Case Else
    End Select
    
    Laboratorio.Caption = Pusing("###,###.##", Laboratorio.Caption)
    Pedido.Caption = Pusing("###,###.##", Pedido.Caption)
    
    WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
    WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
    WStock6.Caption = Pusing("###,###.##", WStock6.Caption)
    WStock7.Caption = Pusing("###,###.##", WStock7.Caption)
    
    PrgArti.WindowState = 0

    Rem by nan
    Wempresa = XEmpresa
    
    Exit Sub
    
Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    StockCons.Visible = False
    Resume Next
    
End Sub

Sub Imprime_Descripcion()

    spEnvase = "ConsultaEnvases " + "'" + Envase.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        DesEnvase.Caption = rstEnvase!Descripcion
        rstEnvase.Close
                Else
        DesEnvase.Caption = ""
    End If
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        rstProveedor.Close
                Else
        DesProveedor.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Inicial.Text) = 0 Then
        Inicial.Text = "0"
    End If
    If Val(Entradas.Text) = 0 Then
        Entradas.Text = "0"
    End If
    If Val(Salidas.Text) = 0 Then
        Salidas.Text = "0"
    End If
    If Val(Minimo.Text) = 0 Then
        Minimo.Text = "0"
    End If
    If Val(Minimo1.Text) = 0 Then
        Minimo1.Text = "0"
    End If
    If Val(Laboratorio.Caption) = 0 Then
        Laboratorio.Caption = "0"
    End If
    If Val(Pedido.Caption) = 0 Then
        Pedido.Caption = "0"
    End If
    If Val(Envase.Text) = 0 Then
        Envase.Text = "0"
    End If
    If Val(Costo1.Text) = 0 Then
        Costo1.Text = "0"
    End If
    If Val(Costo2.Text) = 0 Then
        Costo2.Text = "0"
    End If
    If Val(Costo3.Text) = 0 Then
        Costo3.Text = "0"
    End If
    If Val(Costo4.Text) = 0 Then
        Costo4.Text = "0"
    End If
    If Val(WCosto1.Text) = 0 Then
        WCosto1.Text = "0"
    End If
    If Val(WCosto2.Text) = 0 Then
        WCosto2.Text = "0"
    End If
    If Val(WCosto3.Text) = 0 Then
        WCosto3.Text = "0"
    End If
    If Val(Flete.Text) = 0 Then
        Flete.Text = "0"
    End If
    If Val(Venta.Caption) = 0 Then
        Venta.Caption = "0"
    End If
End Sub

Sub Format_datos()
    Inicial.Text = Pusing("###,###.##", Inicial.Text)
    Entradas.Text = Pusing("###,###.##", Entradas.Text)
    Salidas.Text = Pusing("###,###.##", Salidas.Text)
    Minimo.Text = Pusing("###,###.##", Minimo.Text)
    Minimo1.Text = Pusing("###,###.##", Minimo1.Text)
    Laboratorio.Caption = Pusing("###,###.##", Laboratorio.Caption)
    Pedido.Caption = Pusing("###,###.##", Pedido.Caption)
    Stock.Caption = Pusing("###,###.##", Str$(Val(Inicial.Text) + Val(Entradas.Text) - Val(Salidas.Text)))
    Venta.Caption = Pusing("###,###.##", Venta.Caption)
    Costo1.Text = Pusing("###,###.##", Costo1.Text)
    WCosto1.Text = Pusing("###,###.##", WCosto1.Text)
    ZCosto1.Text = Pusing("###,###.##", ZCosto1.Text)
    Costo4.Text = Pusing("###,###.##", Costo4.Text)
    Costo6.Text = Pusing("###,###.##", Costo6.Text)
    Costo2.Text = Pusing("###,###.##", Costo2.Text)
    
    Costo2.Text = Pusing("###,###.##", Costo2.Text)
    Costo3.Text = Pusing("###,###.##", Costo3.Text)
    WCosto2.Text = Pusing("###,###.##", WCosto2.Text)
    WCosto3.Text = Pusing("###,###.##", WCosto3.Text)
    
    Flete.Text = Pusing("###,###.##", Flete.Text)
End Sub

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desdecodigo.Text + Chr$(34) + " to " + Chr$(34) + HastaCodigo.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Codigo.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Articulo.Codigo , Articulo.Descripcion, Articulo.Inicial, Articulo.Entradas, Articulo.Salidas, Articulo.Minimo, Articulo.Laboratorio, Articulo.Pedido " _
                        + "From " + DSQ + ".dbo.Articulo Articulo " _
                        + "Where Articulo.Codigo >= ' ' AND Articulo.Codigo <= 'ZZ-ZZZ-ZZZ'"
    
    Listado.DataFiles(0) = ""
    Listado.DataFiles(1) = Wempresa + "auxi.mdb"
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = ""
    Listado.Connect = Connect()
    
    Listado.ReportFileName = "WArticulo.rpt"
    Listado.Action = 1
    
    Frame2.Visible = False
    
End Sub

Private Sub Acepta1_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    XEmpresa = Wempresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
                    
    For Cicla = 1 To 11
        If CargaEmpresa(Cicla, 1) <> "" Then
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    If WSalidaError = "N" Then Exit Sub

    Codigo.SetFocus
    Panord.Visible = False

    WDesdeArt = UCase(Codigo.Text)
    WHastaArt = UCase(Codigo.Text)

    Desdefecha1 = Right$(Desdefecha.Text, 4) + Mid$(Desdefecha.Text, 4, 2) + Left$(Desdefecha.Text, 2)
    Hastafecha1 = Right$(Hastafecha.Text, 4) + Mid$(Hastafecha.Text, 4, 2) + Left$(Hastafecha.Text, 2)
    
    With rstWOrden
        .Index = "Orden"
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
    
    XEmpresa = Wempresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
        XHasta = 11
            Else
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
        XHasta = 11
    End If
    
    For A = 1 To XHasta
    
        Wempresa = Empe(A, 1)
        txtOdbc = Empe(A, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        spOrden = "ListaOrdenTotal "
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        If WDesdeArt <= rstOrden!Articulo And WHastaArt >= rstOrden!Articulo Then
                            If Desdefecha1 <= rstOrden!FechaOrd And Hastafecha1 >= rstOrden!FechaOrd Then
                                WOrden = rstOrden!Orden
                                WArticulo = rstOrden!Articulo
                                WProveedor = rstOrden!Proveedor
                                WFecha = rstOrden!Fecha
                                WCantidad = rstOrden!Cantidad
                                WPrecio = rstOrden!Precio
                                WLiberada = rstOrden!Liberada
                                WDevuelta = rstOrden!devuelta
                                WFechaEntrega = rstOrden!FechaEntrega
                                WDesArticulo = ""
                                WDEsProveedor = ""
                                
                                With rstWOrden
                                    .AddNew
                                    !Orden = WOrden
                                    !Articulo = WArticulo
                                    !Proveedor = WProveedor
                                    !Fecha = WFecha
                                    !Cantidad = WCantidad
                                    !Precio = WPrecio
                                    !Liberada = WLiberada
                                    !devuelta = WDevuelta
                                    !FechaEntrega = WFechaEntrega
                                    !DesArticulo = ""
                                    !DesProveedor = ""
                                    .Update
                                End With
                            
                            End If
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstOrden.Close
        
        End If
        
    Next A
    
    Call Conecta_Empresa
    
    With rstWOrden
        .Index = "Orden"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WProveedor = !Proveedor
                WArticulo = !Articulo
                
                WDEsProveedor = ""
                spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    WDEsProveedor = rstProveedor!Nombre
                    rstProveedor.Close
                End If
                
                WDesArticulo = ""
                spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDesArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DesArticulo = WDesArticulo
                !DesProveedor = WDEsProveedor
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Ordenes de Compra por Materia prima consolidado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Option1.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Panord.Visible = False

    Listado.GroupSelectionFormula = "{WOrden.Orden} in 0 to 999999"
    Listado.SQLQuery = ""
    
    Listado.DataFiles(0) = Wempresa + "auxi.mdb"
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = ""
    Listado.Connect = ""
    
    Listado.ReportFileName = "WOrdartcon.rpt"
    Listado.Action = 1

    Exit Sub

Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoErrorII.Visible = True
    Resume Next

End Sub

Private Sub AvisoErrorII_Click()
    AvisoErrorII.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub Cancela1_click()

    Panord.Visible = False
    Codigo.SetFocus

End Sub

Private Sub CargaTipo_Click()
    PrgArticuloEtiqueta.Show
End Sub

Private Sub cmdAdd_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    XEmpresa = Wempresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
                    
    For Cicla = 1 To 11
        If CargaEmpresa(Cicla, 1) <> "" Then
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    If WSalidaError = "N" Then Exit Sub
    
    On Error GoTo WError
    
    If Val(Wempresa) <> 1 Then
        Rem spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstArticulo.RecordCount > 0 Then
        Rem m$ = "Los cambios efectuados solo se realizaran en la empresa en que se encuentra"
        Rem     G% = MsgBox(m$, 0, "Modificacion de Materia Prima")
        Rem     rstArticulo.Close
        Rem End If
        m$ = "Los cambios solo se podran realizar en la empresa Surfactan Planta I"
        G% = MsgBox(m$, 0, "Modificacion de Materia Prima")
        Exit Sub
    End If
    
    If Codigo.Text <> "" Then
    
    WProceso = 0

    If WGraba <> "S" Then
    
        Call Ingresa_clave
        
            Else
            
        WGraba = ""
        XGraba = "N"
    
        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            ZCostoCompara = rstArticulo!Costo2
            Call Redondeo(ZCostoCompara)
            
            ZCostoActual = Val(Costo2.Text)
            Call Redondeo(ZCostoActual)
            
            If ZCostoActual <> ZCostoCompara Then
                XGraba = "S"
            End If
            
            rstArticulo.Close
            
            WPasa = "S"
            
                Else
                
            ZGraba = "N"
            If ZCampo1 = "S" And ZCampo2 = "S" And ZCampo3 = "S" And ZCampo4 = "S" And ZCampo5 = "S" Then
               If ZCampo6 = "S" And ZCampo7 = "S" And ZCampo8 = "S" And ZCampo9 = "S" And ZCampo10 = "S" And ZCampo11 = "S" Then
                    If ZCampo12 = "S" And ZCampo13 = "S" And ZCampo14 = "S" And ZCampo15 = "S" And ZCampo16 = "S" And ZCampo17 = "S" And ZCampo18 = "S" Then
                        If ZCampo19 = "S" And ZCampo20 = "S" And ZCampo21 = "S" And ZCampo22 = "S" And ZCampo23 = "S" And ZCampo24 = "S" Then
                            ZGraba = "S"
                        End If
                    End If
                End If
            End If
            
            If ZGraba = "N" Then
                m$ = "No se puede dar de alta al no haber confirmado la totalidad de los campos"
                G% = MsgBox(m$, 0, "Ingreso de Materia Prima")
                Exit Sub
            End If
            
            WPasa = "N"
            
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZCostoAnterior = rstArticulo!Costo2
            rstArticulo.Close
        End If
        If ZCostoAnterior <> Val(Costo2.Text) Then
            TituloStd.Caption = "Costo Estandar U$S"
        End If
        
    
        Call Verifica_datos
        Rem by nan
        WCodigo = Codigo.Text
        WDescripcion = Descripcion.Text
        XCosto1 = Costo1.Text
        XCosto2 = Costo2.Text
        XCosto3 = Costo3.Text
        XCosto4 = Costo4.Text
        WWCosto1 = WCosto1.Text
        WWCosto2 = WCosto2.Text
        WWCosto3 = WCosto3.Text
        WWInicial = Inicial.Text
        WEntradas = Entradas.Text
        WSalidas = Salidas.Text
        WMinimo = Minimo.Text
        WMinimo1 = Minimo1.Text
        WLaboratorio = Laboratorio.Caption
        WUnidad = Unidad.Text
        WPedido = Pedido.Caption
        WDeposito = Deposito.Text
        WEnvase = Envase.Text
        WRs = Rs.Text
        WProveedor = Proveedor.Text
        WDate = Date$
        WFlete = Flete.Text
        WMoneda = Moneda.Text
        WControla = Str$(Controla.ListIndex)
        WReventa = Str$(Reventa.ListIndex)
        WSedronar = Str$(Sedronar.ListIndex)
        WTipoMp = Str$(TipoMp.ListIndex)
        WCodSedronar = CodSedronar.Text
        WDensidad = Densidad.Text
        WTipoeti = TipoEti.Text
        WCodigoDy = CodigoDy.Text
        WVenta = Venta.Caption
        WClase = Clase.Text
        WIntervencion = Intervencion.Text
        WNaciones = Naciones.Text
        WEmbalaje = Embalaje.Text
        WMeses = Meses.Text
        Wresponsable = Responsable
        WDatosPrv = DatosPrv.Value
        WRestriccion = Restriccion.Value
        
        If TituloStd.Caption = "Std.  Estimado U$S" Then
            ZZTipoCosto = "1"
                Else
            ZZTipoCosto = "0"
        End If
        
        
        If WPasa = "N" Then
            WFecha = ""
            WOrden = ""
            WDife = "0"
            XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + XCosto1 + "','" _
                         + XCosto2 + "','" _
                         + WInicial + "','" _
                         + WEntradas + "','" _
                         + WSalidas + "','" _
                         + WMinimo + "','" _
                         + WLaboratorio + "','" _
                         + WUnidad + "','" _
                         + WPedido + "','" _
                         + WDeposito + "','" _
                         + WEnvase + "','" _
                         + WRs + "','" _
                         + WFecha + "','" _
                         + WOrden + "','" _
                         + WDife + "','" _
                         + WProveedor + "','" _
                         + WDate + "','" + WFlete + "','" _
                         + WMoneda + "','" + WControla + "','" _
                         + WDensidad + "','" + XCosto3 + "','" _
                         + WWCosto1 + "','" + WWCosto2 + "','" _
                         + WWCosto3 + "','" _
                         + WVenta + "'"
                         
            Set rstArticulo = db.OpenRecordset("AltaArticuloII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
                        Else
                        
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Descripcion = " + "'" + WDescripcion + "',"
            ZSql = ZSql & "Costo1 = " + "'" + XCosto1 + "',"
            ZSql = ZSql & "Costo2 = " + "'" + XCosto2 + "',"
            ZSql = ZSql & "Inicial = " + "'" + "0" + "',"
            ZSql = ZSql & "Entradas = " + "'" + WEntradas + "',"
            ZSql = ZSql & "Salidas = " + "'" + WSalidas + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Laboratorio = " + "'" + WLaboratorio + "',"
            ZSql = ZSql & "Unidad = " + "'" + WUnidad + "',"
            ZSql = ZSql & "Pedido = " + "'" + WPedido + "',"
            ZSql = ZSql & "Deposito = " + "'" + WDeposito + "',"
            ZSql = ZSql & "Envase = " + "'" + WEnvase + "',"
            ZSql = ZSql & "Rs = " + "'" + WRs + "',"
            ZSql = ZSql & "Fecha = " + "'" + WFecha + "',"
            ZSql = ZSql & "Orden = " + "'" + WOrden + "',"
            ZSql = ZSql & "Dife = " + "'" + WDife + "',"
            ZSql = ZSql & "Proveedor = " + "'" + WProveedor + "',"
            ZSql = ZSql & "WDate = " + "'" + WDate + "',"
            ZSql = ZSql & "Flete = " + "'" + WFlete + "',"
            ZSql = ZSql & "Moneda = " + "'" + WMoneda + "',"
            ZSql = ZSql & "Controla = " + "'" + WControla + "',"
            ZSql = ZSql & "Densidad = " + "'" + WDensidad + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "Costo3 = " + "'" + XCosto3 + "',"
            ZSql = ZSql & "WCosto1 = " + "'" + WWCosto1 + "',"
            ZSql = ZSql & "WCosto2 = " + "'" + WWCosto2 + "',"
            ZSql = ZSql & "WCosto3 = " + "'" + WWCosto3 + "',"
            ZSql = ZSql & "Venta = " + "'" + WVenta + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                    
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        XParam = "'" + WCodigo + "','" _
                     + WMinimo1 + "'"
                         
        spArticulo = "ModificaArticuloMinimo1 " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        WLeyenda = Str$(Leyenda.ListIndex)
        XParam = "'" + WCodigo + "','" _
                     + WLeyenda + "'"
                         
        spArticulo = "ModificaArticuloLeyenda " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "Responsable = " + "'" + Wresponsable + "',"
        ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
        ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
        ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
        ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
        ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
        ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
        ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
        ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
        ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
        ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
        ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
        ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
        ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
        ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
        ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
        ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
        ZSql = ZSql & "CodigoDy = " + "'" + CodigoDy.Text + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        If Val(Wempresa) = 1 Then
        
            XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + XCosto1 + "','" _
                         + XCosto2 + "','" _
                         + WUnidad + "','" _
                         + WEnvase + "','" _
                         + WRs + "','" _
                         + WProveedor + "','" _
                         + WFlete + "','" _
                         + WMoneda + "','" _
                         + WControla + "','" _
                         + WDensidad + "','" _
                         + XCosto3 + "','" _
                         + WWCosto1 + "','" _
                         + WWCosto2 + "','" _
                         + WWCosto3 + "'"
                         
            
            
            
            WInicial = ""
            WEntradas = ""
            WSalidas = ""
            
            Rem by nan 28-9-2012
            Rem modificacion minimo todas las plantas
      
            
            Rem       WMinimo = ""
            Rem WMinimo1 = ""
            Rem fin by nan
            
            
            WLaboratorio = ""
            WPedido = ""
            WDeposito = ""
            WDate = Date$
            WFecha = "  /  /    "
            WOrden = ""
            WDife = ""
            WVenta = ""
                         
            XParam1 = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + XCosto1 + "','" _
                         + XCosto2 + "','" _
                         + WInicial + "','" _
                         + WEntradas + "','" _
                         + WSalidas + "','" _
                         + WMinimo + "','" _
                         + WLaboratorio + "','" _
                         + WUnidad + "','" _
                         + WPedido + "','" _
                         + WDeposito + "','" _
                         + WEnvase + "','" _
                         + WRs + "','" _
                         + WFecha + "','" _
                         + WOrden + "','" _
                         + WDife + "','" _
                         + WProveedor + "','" _
                         + WDate + "','" _
                         + WFlete + "','" + WMoneda + "','" _
                         + WControla + "','" + WDensidad + "','" _
                         + XCosto3 + "','" + WWCosto1 + "','" _
                         + WWCosto2 + "','" + WWCosto3 + "','" _
                         + WVenta + "'"
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZTipoCosto + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
            ZSql = ZSql & "CodigoDy = " + "'" + CodigoDy.Text + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            Rem agregado de minimo
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            Rem by nan 28-9-2012
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            
            Rem by nan 28-9 agregado min
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                         
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                spArticulo = "ModificaArticuloVariosII " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                spArticulo = "AltaArticuloII " + XParam1
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Reventa = " + "'" + WReventa + "',"
            ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
            ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
            ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
            ZSql = ZSql & "TipoMp = " + "'" + WTipoMp + "',"
            ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
            ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
            ZSql = ZSql & "Leyenda = " + "'" + WLeyenda + "',"
            ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
            ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
            ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
            ZSql = ZSql & "Meses = " + "'" + Meses.Text + "',"
            ZSql = ZSql & "TipoCosto = " + "'" + ZZTipoCosto + "',"
            ZSql = ZSql & "DatosPrv = " + "'" + WDatosPrv + "',"
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                         
            If XGraba = "S" Then
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo4 = " + "'" + Costo4.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        End If
            
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
    
    End If
    
    Exit Sub

WError:
    Resume Next

Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoErrorII.Visible = True
    Resume Next
    
End Sub

Private Sub cmdDelete_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    XEmpresa = Wempresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
                    
    For Cicla = 1 To 11
        If CargaEmpresa(Cicla, 1) <> "" Then
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    If WSalidaError = "N" Then Exit Sub

    If Codigo.Text <> "" Then

    WProceso = 1
    
    If WGraba <> "S" Then
    
        Call Ingresa_clave
        WClave.SetFocus
        
            Else
            
        WGraba = ""

        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            rstArticulo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
                
                If Val(Wempresa) = 1 Then
        
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
                    
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
        
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
        
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
            
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    spArticulo = "BorrarArticulo " + "'" + Codigo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenDynaset, dbSQLPassThrough)
        
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
              
                End If
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    
    End If
    
    Codigo.SetFocus
    
    Exit Sub

WError:
    Resume Next

Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoErrorII.Visible = True
    Resume Next
    
End Sub

Private Sub CmdLimpiar_Click()

    
    Pantalla.Visible = False
    Codigo.Text = "  -   -   "
    Descripcion.Text = ""
    Unidad.Text = ""
    Deposito.Text = ""
    Inicial.Text = ""
    Entradas.Text = ""
    Salidas.Text = ""
    Minimo.Text = ""
    Minimo1.Text = ""
    Laboratorio.Caption = ""
    Pedido.Caption = ""
    Venta.Caption = ""
    Envase.Text = ""
    Rs.Text = ""
    Costo1.Text = ""
    Costo2.Text = ""
    Costo3.Text = ""
    Costo4.Text = ""
    WCosto1.Text = ""
    WCosto2.Text = ""
    WCosto3.Text = ""
    Flete.Text = ""
    Moneda.Text = ""
    Stock.Caption = ""
    DesEnvase.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    WGraba = ""
    WStock1.Caption = "0"
    WStock2.Caption = "0"
    WStock3.Caption = "0"
    WStock4.Caption = "0"
    WStock5.Caption = "0"
    WStock6.Caption = "0"
    WStock7.Caption = "0"
    CodSedronar.Text = ""
    Densidad.Text = ""
    CodigoDy.Text = ""
    Clase.Text = ""
    Intervencion.Text = ""
    Caracteristicas.Text = ""
    Naciones.Text = ""
    Embalaje.Text = ""
    Meses.Text = ""
    Derechos.Text = ""
    TipoEti.Text = ""
    
    Restriccion.Value = 0
    DatosPrv.Value = 0
    
    OrdenI.Text = ""
    OrdenII.Text = ""
    OrdenIII.Text = ""
    FechaOrdenI.Text = ""
    FechaOrdenII.Text = ""
    FechaOrdenIII.Text = ""
    WCosto1Dol.Text = ""
    ZCosto1.Text = ""
    Costo6.Text = ""
    WCosto1Dol.Text = ""
        
    TituloOrdenI.Caption = ""
    TituloOrdenII.Caption = ""
    TituloOrdenIII.Caption = ""
    
    ZCampo1 = "N"
    ZCampo2 = "N"
    ZCampo3 = "N"
    ZCampo4 = "N"
    ZCampo5 = "N"
    ZCampo6 = "N"
    ZCampo7 = "N"
    ZCampo8 = "N"
    ZCampo9 = "N"
    ZCampo10 = "N"
    ZCampo11 = "N"
    ZCampo12 = "N"
    ZCampo13 = "N"
    ZCampo14 = "S"
    ZCampo15 = "N"
    ZCampo16 = "N"
    ZCampo17 = "S"
    ZCampo18 = "S"
    ZCampo19 = "N"
    ZCampo20 = "N"
    ZCampo21 = "N"
    ZCampo22 = "N"
    ZCampo23 = "N"
    ZCampo24 = "N"
    
    TituloStd.Caption = "Costo Estandar U$S"
    
    Controla.ListIndex = 0
    Reventa.ListIndex = 0
    Sedronar.ListIndex = 0
    TipoMp.ListIndex = 1

    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    Codigo.SetFocus
    PrgArti.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spArticulo = "AnteriorArticulo " + "'" + Codigo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstArticulo
        .MoveLast
        Codigo.Text = rstArticulo!Codigo
    End With
    
    rstArticulo.Close
    Call Imprime_Datos
    Rem  Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Articulos", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus

End Sub

Private Sub Command1_Click()

    WMeses = "60"

    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
    ZSql = ZSql & " Where Codigo >= " + "'" + "DY-000-000" + "'"
    ZSql = ZSql & " and Codigo <= " + "'" + "DY-999-999" + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0007"
    txtOdbc = "Empresa07"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0009"
    txtOdbc = "Empresa09"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"

    Wempresa = "0010"
    txtOdbc = "Empresa10"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    Wempresa = "0011"
    txtOdbc = "Empresa11"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Meses = " + "'" + WMeses + "'"


End Sub

Private Sub Command2_Click()

    Da = 0
    With rstFichaMat
        .Index = "Articulo"
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
    
    Erase WVerifica
    WLugar = 0
    
        
    Rem
    Rem proceso los comodatos
    Rem

    Set appExcel = CreateObject("Excel.application")
    
    Rem modificar aca
    Rem Ruta = Nombre del archivo excel
    Rem
    
    ruta = "C:\PartidaChina.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
            Rem modificar aca
            Rem LugarPlanilla separa los acidos
            Rem de los directos
            Rem
            If LugarPlanilla > 32 Then
            
                ZZCantidad = appExcel.cells(LugarPlanilla, 6).Value
                
                If Val(ZZCantidad) <> 0 Then
                
                    ZZPartida = appExcel.cells(LugarPlanilla, 12).Value
                    
                    Entra = "S"
                    
                    For Ciclo = 1 To WLugar
                        If ZZPartida = WVerifica(Ciclo) Then
                            Entra = "N"
                            Exit For
                        End If
                    Next Ciclo
                    
                    If Entra = "S" Then
                        WLugar = WLugar + 1
                        WVerifica(WLugar) = ZZPartida
                    End If
                                
                End If
                
            End If
            
            If LugarPlanilla = 200 Then Exit Do
            
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
    
    
    For Ciclo = 1 To 200
    
        If WVerifica(Ciclo) <> "" Then
    
            WPartiOri = WVerifica(Ciclo)
            nrolote = ""
            Articulo = ""
            WEntra = "N"
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                With rstLaudo
                    .MoveFirst
                    nrolote = IIf(IsNull(rstLaudo!Laudo), "", Str$(rstLaudo!Laudo))
                    WArticulo = rstLaudo!Articulo
                    WEntra = "S"
                    rstLaudo.Close
                End With
            End If
                    
            If WEntra = "N" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.PartiOri = " + "'" + WPartiOri + "'"
                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    With rstMovguia
                        .MoveFirst
                        WEntra = "S"
                        nrolote = IIf(IsNull(rstMovguia!Lote), "", Str$(rstMovguia!Lote))
                        WArticulo = rstMovguia!Articulo
                        rstMovguia.Close
                    End With
                End If
            End If
            
            If WEntra = "N" Then
                m$ = "Partida no encontrada : " + WPartiOri
                A% = MsgBox(m$, 0, "Archivo de Materias Primas")
            End If
            
            If WEntra = "S" Then
            
                XParam = "'" + WArticulo + "','" _
                     + WArticulo + "'"
            
                spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
            
                    With rstEstadistica
            
                        .MoveFirst
                
                        If .NoMatch = False Then
                        Do
                
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                            If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
                    
                                If rstEstadistica!Tipo = 1 Then
                            
                                    ZZLote(1) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                                    ZZLote(2) = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                                    ZZLote(3) = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                                    ZZLote(4) = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                                    ZZLote(5) = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                                    
                                    ZZCanti(1) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                                    ZZCanti(2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                                    ZZCanti(3) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                                    ZZCanti(4) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                                    ZZCanti(5) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                                
                                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                            
                                    If Len(Trim(WLoteAdicional)) = 98 Then
                                        ZZLote(6) = Mid$(WLoteAdicional, 1, 8)
                                        ZZCanti(6) = Mid$(WLoteAdicional, 9, 6)
                                        ZZLote(7) = Mid$(WLoteAdicional, 15, 8)
                                        ZZCanti(7) = Mid$(WLoteAdicional, 23, 6)
                                        ZZLote(8) = Mid$(WLoteAdicional, 29, 8)
                                        ZZCanti(8) = Mid$(WLoteAdicional, 37, 6)
                                        ZZLote(9) = Mid$(WLoteAdicional, 43, 8)
                                        ZZCanti(9) = Mid$(WLoteAdicional, 51, 6)
                                        ZZLote(10) = Mid$(WLoteAdicional, 57, 8)
                                        ZZCanti(10) = Mid$(WLoteAdicional, 65, 6)
                                        ZZLote(11) = Mid$(WLoteAdicional, 71, 8)
                                        ZZCanti(11) = Mid$(WLoteAdicional, 79, 6)
                                        ZZLote(12) = Mid$(WLoteAdicional, 85, 8)
                                        ZZCanti(12) = Mid$(WLoteAdicional, 93, 6)
                                            Else
                                        ZZLote(6) = "0"
                                        ZZCanti(6) = "0"
                                        ZZLote(7) = "0"
                                        ZZCanti(7) = "0"
                                        ZZLote(8) = "0"
                                        ZZCanti(8) = "0"
                                        ZZLote(9) = "0"
                                        ZZCanti(9) = "0"
                                        ZZLote(10) = "0"
                                        ZZCanti(10) = "0"
                                        ZZLote(11) = "0"
                                        ZZCanti(11) = "0"
                                        ZZLote(12) = "0"
                                        ZZCanti(12) = "0"
                                    End If
                                
                                    For ZZCiclo = 1 To 12
                                
                                        If Val(ZZLote(ZZCiclo)) = Val(nrolote) Then
                    
                                            WFecha = rstEstadistica!Fecha
                                            WCodigo = rstEstadistica!Numero
                                            WObservaciones = rstEstadistica!Cliente
                                            WTipo = rstEstadistica!Tipo
                                            WCantidad = Val(ZZCanti(ZZCiclo))
                                            WPrecio = rstEstadistica!Precio
                                            WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2)
                                            If Trim(WFechaord) >= "200501" Then
                                            
                                                With rstFichaMat
                                                    .AddNew
                                                    !Articulo = WArticulo
                                                    !Fecha = WFecha
                                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2)
                                                    !Tipo = 0
                                                    !Numero = WCodigo
                                                    !Inicial = 0
                                                    !Entrada = WCantidad
                                                    !Salida = WPrecio
                                                    !Lista1 = "Fact"
                                                    !Observaciones = WObservaciones
                                                    !Descripcion = ""
                                                    !Lista2 = ""
                                                    !Lote = Val(nrolote)
                                                    !Saldo = 0
                                                    !Empresa = NombreEmpresa
                                                    !PartiOri = WPartiOri
                                                    .Update
                                                End With
                                            
                                            End If
                                        End If
                                    Next ZZCiclo
                        
                                End If
                    
                            End If
                
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                        End If
                
                    End With
                    rstEstadistica.Close
                End If
            End If
        End If
        
    Next Ciclo
        
        
        
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WObservaciones = !Observaciones
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WObservaciones = rstCliente!Razon
                    rstCliente.Close
                End If
                !Descripcion = WDescripcion
                !Observaciones = WObservaciones
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.ReportFileName = "WLotematResumen.rpt"

    Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    Rem modificar aca
    Rem poner 0 para pantalla
    Rem poner 1 para impresora
    Rem
    
    Listado.Destination = 0
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
     
    Listado.Action = 0
    
    Listado.ReportFileName = "WLotematResumenII.rpt"

    Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem modificar aca
    Rem poner 0 para pantalla
    Rem poner 1 para impresora
    Rem
    
    Listado.Destination = 0
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
    Listado.Action = 1
        

End Sub


Private Sub Command3_Click()


    Dim ZTipoMp(5000, 2) As String

        
    Rem
    Rem proceso los comodatos
    Rem
    LugarPlanilla = 6
    ZZLugar = 0

    Set appExcel = CreateObject("Excel.application")
    
    Rem modificar aca
    Rem Ruta = Nombre del archivo excel
    Rem
    
    ruta = "S:\TipoMp.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
            ZZCodigo = appExcel.cells(LugarPlanilla, 1).Value
            ZZComo = appExcel.cells(LugarPlanilla, 6).Value
            ZZHomo = appExcel.cells(LugarPlanilla, 7).Value
            ZZRepre = appExcel.cells(LugarPlanilla, 8).Value
            
            ZLugar = 0
            If Trim(UCase(ZZComo)) = "X" Then
                ZLugar = 1
            End If
            If Trim(UCase(ZZHomo)) = "X" Then
                ZLugar = 2
            End If
            If Trim(UCase(ZZRepre)) = "X" Then
                ZLugar = 3
            End If
            
            If ZLugar > 0 Then
                ZZLugar = ZZLugar + 1
                ZTipoMp(ZZLugar, 1) = ZZCodigo
                ZTipoMp(ZZLugar, 2) = ZLugar - 1
            End If
            
            If LugarPlanilla = 5000 Then Exit Do
                
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
    

    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo









    Wempresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
            








    Wempresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
            







    Wempresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo






    Wempresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo






    Wempresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo








    Wempresa = "0007"
    txtOdbc = "Empresa07"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo







    Wempresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo






    Wempresa = "0009"
    txtOdbc = "Empresa09"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
            


    Wempresa = "0010"
    txtOdbc = "Empresa10"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo




    Wempresa = "0011"
    txtOdbc = "Empresa11"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    For Ciclo = 1 To ZZLugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Articulo SET "
        ZSql = ZSql & "TipoMp = " + "'" + ZTipoMp(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + ZTipoMp(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo


End Sub






Private Sub Command4_Click()

    aa = Wempresa

    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "FechaCierre = " + "'" + "10/12/2008" + "',"
    ZSql = ZSql & "OrdFechaCierre = " + "'" + "20081210" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Command5_Click()

    Rem dada
    Rem dada
    Rem dada
    Dim ZActualiza(10000, 2) As String
    Dim ZLugar As Integer
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZActualiza(ZLugar, 1) = rstArticulo!Codigo
                    ZActualiza(ZLugar, 2) = ""
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    XEmpresa = Wempresa
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZActualiza(Ciclo, 1)
    
        spArticulo = "ConsultaArticulo " + "'" + ZCodigo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZActualiza(Ciclo, 2) = IIf(IsNull(rstArticulo!Meses), "", rstArticulo!Meses)
            rstArticulo.Close
                Else
            ZActualiza(Ciclo, 1) = ""
        End If
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZActualiza(Ciclo, 1)
        ZMeses = ZActualiza(Ciclo, 2)
        
        If ZCodigo <> "" Then
        
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            ZSql = ZSql & "Meses = " + "'" + ZMeses + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + ZCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    m$ = "Proceso Finalizado"
    G% = MsgBox(m$, 0, "Ingreso de Materia Prima")

End Sub

Private Sub Command22_Click()

    WAyuda = "FOOD"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Marcas"
    ZSql = ZSql + " Where Marcas.Descripcion LIKE " + "'" + "%" + WAyuda + "%" + "'"
    spMarcas = ZSql
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMarcas.RecordCount > 0 Then
        With rstMarcas
            .MoveFirst
            Do
                If .EOF = False Then
                    AA1 = rstMarcas!Proveedor
                    aa2 = rstMarcas!Articulo
                    aa3 = rstMarcas!Descripcion
                    Stop
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMarcas.Close
    End If


End Sub

Private Sub ComMarcas_Click()

    parance.Visible = False
    ConsultaMarcas.Height = 4815
    ConsultaMarcas.Left = 480
    ConsultaMarcas.Top = 960
    ConsultaMarcas.Width = 9495

    ConsultaMarcas.Visible = True
    WCampo1.Caption = Codigo.Text
    WCampo2.Caption = Descripcion.Text
    
    Call Limpia_Vector
    
    Lugar = 0
    
    spMarcas = "ListaMarcasArticulo " + "'" + Codigo.Text + "'"
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMarcas.RecordCount > 0 Then
    With rstMarcas
        .MoveFirst
        Do
            If .EOF = False Then
                Lugar = Lugar + 1
                WVector1.Row = Lugar
                WVector1.Col = 1
                WVector1.Text = !Proveedor
                WVector1.Col = 2
                WVector1.Text = ""
                WVector1.Col = 3
                WVector1.Text = !Descripcion
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstMarcas.Close
    End If
    
    For Ciclo = 1 To Lugar
        WVector1.Row = Ciclo
        WVector1.Col = 1
        spProveedor = "ConsultaProveedores " + "'" + WVector1.Text + "'"
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WVector1.Col = 2
            WVector1.Text = rstProveedor!Nombre
            rstProveedor.Close
        End If
    Next Ciclo
    
End Sub

Private Sub ConCoti_Click()

    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
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
        End With
        rstCambios.Close
            Else
        Paridad = 1
        ParidadII = 1
    End If
    
    Codigo.Text = UCase(Codigo.Text)

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
    
    Pasa = 0
    Canti = 0
    
    XParam = "'" + Codigo.Text + "','" _
            + Codigo.Text + "'"
    
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
                    
                    Select Case !Moneda
                        Case 0
                            WPrecio = !Precio * Paridad
                        Case 1
                            WPrecio = !Precio
                        Case Else
                            WPrecio = !Precio * ParidadII
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
                            
                            If XVector(3, 2) <> "" Then
                                WAuxi = XVector(3, 5)
                                        Else
                                If XVector(2, 2) <> "" Then
                                    WAuxi = XVector(2, 5)
                                        Else
                                    WAuxi = XVector(1, 5)
                                End If
                            End If
                            WAuxi = Str$(Val(WAuxi) - 90000000)
                                
                            Call Ceros(WAuxi, 9)
                        
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
            Rem         Else
            Rem     If Val(XVector(2, 2)) <> 0 Then
            Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
            Rem             Else
            Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
            Rem     End If
            Rem End If
            
            If XVector(3, 2) <> "" Then
                WAuxi = XVector(3, 5)
                    Else
                If XVector(2, 2) <> "" Then
                    WAuxi = XVector(2, 5)
                        Else
                    WAuxi = XVector(1, 5)
                End If
            End If
            WAuxi = Str$(Val(WAuxi) - 90000000)
                            
            Call Ceros(WAuxi, 9)
                
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
                    
                WCategoriaI = ""
                WCategoriaII = ""
                
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
                    
                    rstProveedor.Close
                End If
                
                ZZIngre = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                    rstArticulo.Close
                End If
                
                If ZZTipoMp = 1 Then
                
                    For i = 1 To 2
                    
                        Select Case i
                            Case 1
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
  
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Homologa"
                        ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
                        ZSql = ZSql + " and CodigoMp = " + "'" + WArticulo + "'"
                        ZSql = ZSql + " and Estado = " + "'" + "1" + "'"
                        spHomologa = ZSql
                        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHomologa.RecordCount > 0 Then
                            ZZIngre = "  (H)   "
                            rstHomologa.Close
                            Exit For
                        End If
                        
                    Next i
                    
                    If ZZIngre = "" Then
                    
                        For i = 1 To 11
                        
                            Select Case i
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
                            ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                            ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                            spOrden = ZSql
                            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            If rstOrden.RecordCount > 0 Then
                                ZZIngre = "  (H)   "
                                rstOrden.Close
                                Exit For
                            End If
                            
                        Next i
                                
                    End If
                    
                End If
                
                WDescriProveedor = ZZIngre + Trim(WDescriProveedor)
                
                If WCategoriaI <> "" And WCategoriaII <> "" Then
                    WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                End If
                
                !DescriProveedor = Left$(WDescriProveedor, 50)
                !DescriArticulo = WDescriArticulo
                !Titulo = "(En Pesos)"
                    
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
                        
    Listado.WindowTitle = "Listado de Cotizaciones por Articulo en Pesos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 0
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = ""
    
    Listado.SQLQuery = ""
    Listado.Connect = ""
    
    Codigo.SetFocus
    
    Listado.GroupSelectionFormula = "{Listcot.Articulo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
    
    Listado.ReportFileName = "WCotart.rpt"
    Listado.Action = 1
    
    Call Conecta_Empresa

End Sub

Private Sub ConCoti1_Click()

    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
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
        End With
        rstCambios.Close
            Else
        Paridad = 1
        ParidadII = 1
    End If
    
    Codigo.Text = UCase(Codigo.Text)

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
    
    Pasa = 0
    Canti = 0
    
    XParam = "'" + Codigo.Text + "','" _
            + Codigo.Text + "'"
    
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
                    
                    Select Case !Moneda
                        Case 0
                            WPrecio = !Precio
                        Case 1
                            WPrecio = !Precio / Paridad
                        Case Else
                            WCoeParidad = ParidadII / Paridad
                            WPrecio = !Precio * WCoeParidad
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
                            
                            If XVector(3, 2) <> "" Then
                                WAuxi = XVector(3, 5)
                                        Else
                                If XVector(2, 2) <> "" Then
                                    WAuxi = XVector(2, 5)
                                        Else
                                    WAuxi = XVector(1, 5)
                                End If
                            End If
                            WAuxi = Str$(Val(WAuxi) - 90000000)
                                
                            Call Ceros(WAuxi, 9)
                        
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
            Rem         Else
            Rem     If Val(XVector(2, 2)) <> 0 Then
            Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
            Rem             Else
            Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
            Rem     End If
            Rem End If
            
            If XVector(3, 2) <> "" Then
                WAuxi = XVector(3, 5)
                    Else
                If XVector(2, 2) <> "" Then
                    WAuxi = XVector(2, 5)
                        Else
                    WAuxi = XVector(1, 5)
                End If
            End If
            WAuxi = Str$(Val(WAuxi) - 90000000)
                            
            Call Ceros(WAuxi, 9)
                
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
                    
                WCategoriaI = ""
                WCategoriaII = ""
                
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
                    
                    rstProveedor.Close
                End If
                
                ZZIngre = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                    rstArticulo.Close
                End If
                
                If ZZTipoMp = 1 Then
                
                    For i = 1 To 2
                    
                        Select Case i
                            Case 1
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
  
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Homologa"
                        ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
                        ZSql = ZSql + " and CodigoMp = " + "'" + WArticulo + "'"
                        ZSql = ZSql + " and Estado = " + "'" + "1" + "'"
                        spHomologa = ZSql
                        Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHomologa.RecordCount > 0 Then
                            ZZIngre = "  (H)   "
                            rstHomologa.Close
                            Exit For
                        End If
                        
                    Next i
                    
                    If ZZIngre = "" Then
                    
                        For i = 1 To 11
                        
                            Select Case i
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
                            ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                            ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                            spOrden = ZSql
                            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            If rstOrden.RecordCount > 0 Then
                                ZZIngre = "  (H)   "
                                rstOrden.Close
                                Exit For
                            End If
                            
                        Next i
                                
                    End If
                    
                End If
                
                WDescriProveedor = ZZIngre + Trim(WDescriProveedor)
                
                If WCategoriaI <> "" And WCategoriaII <> "" Then
                    WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                End If
                
                !DescriProveedor = Left$(WDescriProveedor, 50)
                !DescriArticulo = WDescriArticulo
                !Titulo = "(En Dolares)"
                    
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
                        
    Listado.WindowTitle = "Listado de Cotizaciones por Articulo en Dolares"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 0
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = ""
    
    Listado.SQLQuery = ""
    Listado.Connect = ""
    
    Codigo.SetFocus
    
    Listado.GroupSelectionFormula = "{Listcot.Articulo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
    
    Listado.ReportFileName = "WCotart.rpt"
    Listado.Action = 1
    
    Call Conecta_Empresa

End Sub

Private Sub ConCpa_Click()

    Desdefecha.Text = "  /  /    "
    Hastafecha.Text = "  /  /    "
    
    Panord.Visible = True
    Option2.Value = True
    Desdefecha.SetFocus

End Sub

Private Sub Cuadro_Click()
    CargaCuadro.Visible = True
    PartidaCuadro.Text = ""
    PartidaCuadro.SetFocus
End Sub

Private Sub DatosComple_Click()

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 6
    


End Sub

Private Sub EtiCancela_Click()
    PantaEtiDy.Visible = False
    Frame6.Visible = True
    Frame3.Visible = True


End Sub

Private Sub EtiquetaDy_Click()
     
    Frame6.Visible = False
    Frame3.Visible = False
     
    PantaEtiDy.Height = 5055
    PantaEtiDy.Left = 1200
    PantaEtiDy.Top = 840
    PantaEtiDy.Width = 7935

    PantaEtiDy.Visible = True
    
    TipoBarra.Clear
    
    TipoBarra.AddItem "Completo"
    TipoBarra.AddItem "Comprimida"
    TipoBarra.AddItem "Dividida"
    
    TipoBarra.ListIndex = 0
    
    EtiPartida.Text = ""
    EtiCantidad.Text = "25"
    EtiCantidadEti.Text = ""
    EtiDescri1.Text = ""
    EtiDescri2.Text = ""
    EtiDescri3.Text = ""
    EtiDescri4.Text = ""
    EtiDescri5.Text = ""
    EtiDescri6.Text = ""
    EtiArticulo.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM DescriDy"
    ZSql = ZSql + " Where DescriDy.Codigo = " + "'" + Codigo.Text + "'"
    spDescriDy = ZSql
    Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
    If rstDescriDy.RecordCount > 0 Then
        EtiDescri1.Text = rstDescriDy!Descri1
        EtiDescri2.Text = rstDescriDy!Descri2
        EtiDescri3.Text = rstDescriDy!Descri3
        EtiDescri4.Text = rstDescriDy!Descri4
        EtiDescri5.Text = rstDescriDy!Descri5
        EtiDescri6.Text = rstDescriDy!Descri6
        EtiArticulo.Text = rstDescriDy!Articulo
        rstDescriDy.Close
    End If
        
    EtiPartida.SetFocus

End Sub

Private Sub EtiPartida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WPartiOri = PartidaCuadro.Text
        WEntra = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + EtiPartida.Text + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + Codigo.Text + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                WEntra = "S"
                rstLaudo.Close
            End With
        End If
                
        If WEntra = "N" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.PartiOri = " + "'" + EtiPartida.Text + "'"
            ZSql = ZSql + " and Guia.Articulo = " + "'" + Codigo.Text + "'"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                With rstMovguia
                    .MoveFirst
                    WEntra = "S"
                    rstMovguia.Close
                End With
            End If
        End If
        
        If WEntra = "S" Then
            EtiCantidadEti.SetFocus
                Else
            m$ = "Partida Inexistnte"
            A% = MsgBox(m$, 0, "Impresion de Pedidos")
            Exit Sub
        End If

    End If
End Sub

Private Sub EtiPartidaEti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EtiPartida.SetFocus
    End If
End Sub


Private Sub EtiImpre1_Click()

    Rem VERIDICA LA PARTIDA

    WEntra = "N"
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + EtiPartida.Text + "'"
    ZSql = ZSql + " and Laudo.Articulo = " + "'" + Codigo.Text + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            WEntra = "S"
            rstLaudo.Close
        End With
    End If
            
    If WEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.PartiOri = " + "'" + EtiPartida.Text + "'"
        ZSql = ZSql + " and Guia.Articulo = " + "'" + Codigo.Text + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            With rstMovguia
                .MoveFirst
                WEntra = "S"
                rstMovguia.Close
            End With
        End If
    End If
    
    If WEntra = "N" Then
        m$ = "Partida Inexistnte"
        A% = MsgBox(m$, 0, "Impresion de Pedidos")
        Exit Sub
    End If





    Auxi = EtiCantidad.Text
    Call Ceros(Auxi, 5)

    ZZBarra = "10" + Trim(EtiPartida.Text) + "(240)" + Trim(EtiArticulo.Text) + "3101" + Auxi
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM DescriDy"
    ZSql = ZSql + " Where DescriDy.Codigo = " + "'" + Codigo.Text + "'"
    spDescriDy = ZSql
    Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
    If rstDescriDy.RecordCount > 0 Then
    
        rstDescriDy.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE DescriDy SET "
        ZSql = ZSql + " Descri1 = " + "'" + EtiDescri1.Text + "',"
        ZSql = ZSql + " Descri2 = " + "'" + EtiDescri2.Text + "',"
        ZSql = ZSql + " Descri3 = " + "'" + EtiDescri3.Text + "',"
        ZSql = ZSql + " Descri4 = " + "'" + EtiDescri4.Text + "',"
        ZSql = ZSql + " Descri5 = " + "'" + EtiDescri5.Text + "',"
        ZSql = ZSql + " Descri6 = " + "'" + EtiDescri6.Text + "',"
        ZSql = ZSql + " Articulo = " + "'" + EtiArticulo.Text + "',"
        ZSql = ZSql + " Partida = " + "'" + EtiPartida.Text + "',"
        ZSql = ZSql + " Barra = " + "'" + ZZBarra + "',"
        ZSql = ZSql + " Cantidad = " + "'" + EtiCantidad.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spDescriDy = ZSql
        Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO DescriDy ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descri1 ,"
        ZSql = ZSql + "Descri2 ,"
        ZSql = ZSql + "Descri3 ,"
        ZSql = ZSql + "Descri4 ,"
        ZSql = ZSql + "Descri5 ,"
        ZSql = ZSql + "Descri6 ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Barra ,"
        ZSql = ZSql + "Cantidad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + EtiDescri1.Text + "',"
        ZSql = ZSql + "'" + EtiDescri2.Text + "',"
        ZSql = ZSql + "'" + EtiDescri3.Text + "',"
        ZSql = ZSql + "'" + EtiDescri4.Text + "',"
        ZSql = ZSql + "'" + EtiDescri5.Text + "',"
        ZSql = ZSql + "'" + EtiDescri6.Text + "',"
        ZSql = ZSql + "'" + EtiArticulo.Text + "',"
        ZSql = ZSql + "'" + EtiPartida.Text + "',"
        ZSql = ZSql + "'" + ZZBarra + "',"
        ZSql = ZSql + "'" + EtiCantidad.Text + "')"
        
        spDescriDy = ZSql
        Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
    
    End If


    Listado.WindowTitle = "Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 1
   Rem  Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT DescriDy.Codigo, DescriDy.Descri1, DescriDy.Descri2, DescriDy.Descri3, DescriDy.Descri4, DescriDy.Descri5, DescriDy.Descri6, DescriDy.Articulo, DescriDy.Partida, DescriDy.Cantidad " _
            + "From " _
            + DSQ + ".dbo.DescriDy DescriDy " _
            + "Where " _
            + "DescriDy.Codigo >= '" + Codigo.Text + "' AND " _
            + "DescriDy.Codigo <= '" + Codigo.Text + "'"
    
    Listado.GroupSelectionFormula = "{DescriDy.Codigo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
    Listado.SelectionFormula = "{DescriDy.Codigo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)

    Listado.Connect = Connect()
    Listado.ReportFileName = "ImpreEtiDy.rpt"
    
    Listado.CopiesToPrinter = Val(EtiCantidadEti.Text)
    
    Listado.Action = 1

End Sub



Private Sub EtiImpre2_Click()


    Rem VERIDICA LA PARTIDA

    WEntra = "N"
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + EtiPartida.Text + "'"
    ZSql = ZSql + " and Laudo.Articulo = " + "'" + Codigo.Text + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            WEntra = "S"
            rstLaudo.Close
        End With
    End If
            
    If WEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.PartiOri = " + "'" + EtiPartida.Text + "'"
        ZSql = ZSql + " and Guia.Articulo = " + "'" + Codigo.Text + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            With rstMovguia
                .MoveFirst
                WEntra = "S"
                rstMovguia.Close
            End With
        End If
    End If
    
    If WEntra = "N" Then
        m$ = "Partida Inexistnte"
        A% = MsgBox(m$, 0, "Impresion de Pedidos")
        Exit Sub
    End If


    Auxi = EtiCantidad.Text
    Call Ceros(Auxi, 5)

    Select Case TipoBarra.ListIndex
        Case 0
            ZZBarra = "10" + Trim(EtiPartida.Text) + "(240)" + Trim(EtiArticulo.Text) + "3101" + Auxi
        Case 1
            ZZBarra = "       " + Trim(EtiPartida.Text) + Trim(EtiArticulo.Text) + Auxi
        Case Else
            ZZBarra = "       " + Trim(EtiPartida.Text) + "     " + Trim(EtiArticulo.Text)
    End Select
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM DescriDy"
    ZSql = ZSql + " Where DescriDy.Codigo = " + "'" + Codigo.Text + "'"
    spDescriDy = ZSql
    Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
    If rstDescriDy.RecordCount > 0 Then
    
        rstDescriDy.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE DescriDy SET "
        ZSql = ZSql + " Descri1 = " + "'" + EtiDescri1.Text + "',"
        ZSql = ZSql + " Descri2 = " + "'" + EtiDescri2.Text + "',"
        ZSql = ZSql + " Descri3 = " + "'" + EtiDescri3.Text + "',"
        ZSql = ZSql + " Descri4 = " + "'" + EtiDescri4.Text + "',"
        ZSql = ZSql + " Descri5 = " + "'" + EtiDescri5.Text + "',"
        ZSql = ZSql + " Descri6 = " + "'" + EtiDescri6.Text + "',"
        ZSql = ZSql + " Articulo = " + "'" + EtiArticulo.Text + "',"
        ZSql = ZSql + " Partida = " + "'" + EtiPartida.Text + "',"
        ZSql = ZSql + " Barra = " + "'" + ZZBarra + "',"
        ZSql = ZSql + " Cantidad = " + "'" + EtiCantidad.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spDescriDy = ZSql
        Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO DescriDy ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descri1 ,"
        ZSql = ZSql + "Descri2 ,"
        ZSql = ZSql + "Descri3 ,"
        ZSql = ZSql + "Descri4 ,"
        ZSql = ZSql + "Descri5 ,"
        ZSql = ZSql + "Descri6 ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Barra ,"
        ZSql = ZSql + "Cantidad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + EtiDescri1.Text + "',"
        ZSql = ZSql + "'" + EtiDescri2.Text + "',"
        ZSql = ZSql + "'" + EtiDescri3.Text + "',"
        ZSql = ZSql + "'" + EtiDescri4.Text + "',"
        ZSql = ZSql + "'" + EtiDescri5.Text + "',"
        ZSql = ZSql + "'" + EtiDescri6.Text + "',"
        ZSql = ZSql + "'" + EtiArticulo.Text + "',"
        ZSql = ZSql + "'" + EtiPartida.Text + "',"
        ZSql = ZSql + "'" + ZZBarra + "',"
        ZSql = ZSql + "'" + EtiCantidad.Text + "')"
        
        spDescriDy = ZSql
        Set rstDescriDy = db.OpenRecordset(spDescriDy, dbOpenSnapshot, dbSQLPassThrough)
    
    End If


    Listado.WindowTitle = "Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT DescriDy.Codigo, DescriDy.Barra " _
            + "From " _
            + DSQ + ".dbo.DescriDy DescriDy " _
            + "Where " _
            + "DescriDy.Codigo >= '" + Codigo.Text + "' AND " _
            + "DescriDy.Codigo <= '" + Codigo.Text + "'"
    
    Listado.GroupSelectionFormula = "{DescriDy.Codigo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
    Listado.SelectionFormula = "{DescriDy.Codigo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)

    Listado.Connect = Connect()
    Listado.ReportFileName = "ImpreBarraDy.rpt"
    
    Listado.CopiesToPrinter = Val(EtiCantidadEti.Text)
    
    Listado.Action = 1

End Sub



Rem dada
Rem dada
Rem dada
Rem dada
Rem dada
Rem dada




Private Sub FinConsulta_Click()
    ConsultaMarcas.Visible = False
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgArti.Caption = "Ingreso de Materias Primas :  " + !Nombre
        End If
    End With
    
    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
        Stock1.Caption = "Pta I"
        Stock2.Caption = "Pta II"
        Stock3.Caption = "Pta III"
        Stock4.Caption = "Pta IV"
        Stock5.Caption = "Pta V"
        Stock6.Caption = "Pta VI"
        Stock7.Caption = "Pta VII"
            Else
        Stock1.Caption = "Pta I"
        Stock2.Caption = "Pta II"
        Stock3.Caption = "Pta V"
        Stock4.Caption = "Pta IV"
        Stock5.Caption = ""
        Stock6.Caption = ""
        Stock7.Caption = ""
    End If
    
    Controla.Clear
    
    Controla.AddItem "Controla Lote"
    Controla.AddItem "No Controla Lote"
    Controla.AddItem "A Granel"
    
    Controla.ListIndex = 0
    
    Reventa.Clear
    
    Reventa.AddItem ""
    Reventa.AddItem "Si"
    
    Reventa.ListIndex = 0
    
    Sedronar.Clear
    
    Sedronar.AddItem ""
    Sedronar.AddItem "Si"
    
    Sedronar.ListIndex = 0
    
    TipoMp.Clear
    
    TipoMp.AddItem "Comodity"
    TipoMp.AddItem "Homologable"
    TipoMp.AddItem "Representada"
    
    TipoMp.ListIndex = 1
    
    PasaLeyenda = "N"
    
    Leyenda.Clear
    
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CIF"
    Leyenda.AddItem "CFR"
    Leyenda.AddItem "CPT"
    Leyenda.AddItem "EXW"
    Leyenda.AddItem "FCA"
    Leyenda.AddItem ""
    
    Leyenda.ListIndex = 0
    
    PasaLeyenda = "S"
    
    ZCampo1 = "N"
    ZCampo2 = "N"
    ZCampo3 = "N"
    ZCampo4 = "N"
    ZCampo5 = "N"
    ZCampo6 = "N"
    ZCampo7 = "N"
    ZCampo8 = "N"
    ZCampo9 = "N"
    ZCampo10 = "N"
    ZCampo11 = "N"
    ZCampo12 = "N"
    ZCampo13 = "N"
    ZCampo14 = "S"
    ZCampo15 = "N"
    ZCampo16 = "N"
    ZCampo17 = "S"
    ZCampo18 = "S"
    ZCampo19 = "N"
    ZCampo20 = "N"
    ZCampo21 = "N"
    ZCampo22 = "N"
    ZCampo23 = "N"
    ZCampo24 = "N"
    
    Restriccion.Value = 0
    DatosPrv.Value = 0
    
    TituloStd.Caption = "Costo Estandar U$S"
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Liscot
    OPEN_FILE_WOrden
    OPEN_FILE_FichaMat
End Sub

Private Sub GrabaMinimo_Click()

    ZSql = ""
    ZSql = ZSql & "UPDATE Articulo SET "
    ZSql = ZSql & "Minimo1 = " + "'" + Minimo1.Text + "'"
    ZSql = ZSql & " Where Codigo = " + "'" + Codigo.Text + "'"
                
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub

Private Sub Historial_Click()
    CargaPartida.Visible = True
    Partida.Text = ""
    Partida.SetFocus
End Sub

Private Sub LeeMsds_Click()

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

    If Left$(UCase(Codigo.Text), 2) = "DY" Then
            
        MiRuta = "W:\msdssis\msds"  ' Establece la ruta.
        ZZLee = MiRuta + Codigo.Text + ".pdf"
        ZZEstado = Dir(ZZLee)
        If ZZEstado <> "" Then
            RetVal = Shell(ZZPasaExe + " " + ZZLee + " ", 3)
        End If
            
            Else

        ' Muestra los nombres en C:\ que representan directorios.
        ZZCodigo = Left$(Codigo.Text, 2) + Mid$(Codigo.Text, 4, 3) + Right$(Codigo.Text, 3)
        ZZArchiMsDs = ""
        MiRuta = "W:\msdssis\MSDS mp\"  ' Establece la ruta.
        MiNombre = Dir(MiRuta, vbDirectory) ' Recupera la primera entrada.
        Do While MiNombre <> "" ' Inicia el bucle.
            dada = MiNombre
            MiNombre = Trim(UCase(Dir))  ' Obtiene siguiente entrada.
            
            WEspacios = Len(ZZCodigo)
            Da = Len(MiNombre) - WEspacios
            ZZSalida = "N"
            
            For Aaa = 1 To Da + 1
                If Left$(ZZCodigo, WEspacios) = Mid$(MiNombre, Aaa, WEspacios) Then
                    ZZArchiMsDs = MiNombre
                    ZZSalida = "S"
                    Exit For
                End If
            Next Aaa
            
            If ZZSalida = "S" Then
                Exit Do
            End If
            
        Loop
    
    
        If ZZArchiMsDs <> "" Then
            ZZLee = MiRuta + ZZArchiMsDs
            ZZEstado = Dir(ZZLee)
            If ZZEstado <> "" Then
                RetVal = Shell(ZZPasaExe + "  " + ZZLee + " ", 3)
            End If
        End If


    End If

End Sub

Private Sub Partida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        CargaCuadro.Visible = False

        Da = 0
        With rstFichaMat
            .Index = "Articulo"
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
        
        WPartiOri = PartidaCuadro.Text
        nrolote = ""
        Articulo = ""
        WEntra = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                nrolote = IIf(IsNull(rstLaudo!Laudo), "", Str$(rstLaudo!Laudo))
                Articulo = IIf(IsNull(rstLaudo!Articulo), "", Str$(rstLaudo!Articulo))
                WEntra = "S"
                rstLaudo.Close
            End With
        End If
                
        If WEntra = "N" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " and Guia.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                With rstMovguia
                    .MoveFirst
                    nrolote = IIf(IsNull(rstMovguia!Lote), "", Str$(rstMovguia!Lote))
                    Articulo = IIf(IsNull(rstMovguia!Articulo), "", Str$(rstMovguia!Articulo))
                    rstMovguia.Close
                End With
            End If
        End If
        
        
        
        XParam = "'" + Articulo + "','" _
                     + Articulo + "'"
        
        spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
        
            With rstMovvar
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If !Tipo = "M" Then
                    
                        WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        
                        If Val(WLote) = Val(nrolote) Then
                    
                            WArticulo = rstMovvar!Articulo
                            WCantidad = rstMovvar!Cantidad
                            WFecha = rstMovvar!Fecha
                            WCodigo = rstMovvar!Codigo
                            WMovi = rstMovvar!Movi
                            WTipomov = Val(rstMovvar!Tipomov)
                            WObservaciones = rstMovvar!Observaciones
                            WSaldo = 0
                        
                            With rstFichaMat
                        
                                .AddNew
                                !Articulo = WArticulo
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WCodigo
                                !Inicial = 0
                                If WMovi = "E" Then
                                    !Entrada = WCantidad
                                    !Salida = 0
                                        Else
                                    !Entrada = 0
                                    !Salida = WCantidad
                                End If
                                !Observaciones = WObservaciones
                                !Descripcion = WDesArticulo
                                If WTipomov = 0 Or WTipomov = 1 Then
                                    !Lista1 = "Mov.Var"
                                        Else
                                    !Lista1 = "Guia In"
                                End If
                                !Lista2 = ""
                                !Lote = WLote
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                            End With
                        End If
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstMovvar.Close
        End If
        
        
        
        
        
        
        Da = 0
        With rstFichaMat
            .Index = "Articulo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    .Edit
                    WArticulo = !Articulo
                    WObservaciones = !Observaciones
                    WDescripcion = ""
                    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    If !Lista1 = "Devol." Or !Lista1 = "Factura" Then
                        spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCliente.RecordCount > 0 Then
                            WObservaciones = rstCliente!Razon
                            rstCliente.Close
                        End If
                    End If
                    !Descripcion = WDescripcion
                    !Observaciones = WObservaciones
                    .Update
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
        Listado.ReportFileName = "WLotemat.rpt"
    
        Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        Listado.Destination = 0
        Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
        
        Listado.Action = 1

    End If
End Sub

Private Sub PartidaCuadro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        CargaCuadro.Visible = False

        Da = 0
        With rstFichaMat
            .Index = "Articulo"
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
        
        WPartiOri = PartidaCuadro.Text
        nrolote = ""
        Articulo = ""
        WEntra = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                nrolote = IIf(IsNull(rstLaudo!Laudo), "", Str$(rstLaudo!Laudo))
                WArticulo = rstLaudo!Articulo
                WEntra = "S"
                rstLaudo.Close
            End With
        End If
                
        If WEntra = "N" Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " and Guia.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                With rstMovguia
                    .MoveFirst
                    nrolote = IIf(IsNull(rstMovguia!Lote), "", Str$(rstMovguia!Lote))
                    WArticulo = rstMovguia!Articulo
                    rstMovguia.Close
                End With
            End If
            
        End If
        
        
        
        XParam = "'" + WArticulo + "','" _
                     + WArticulo + "'"
        
        spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
        
            With rstMovvar
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If !Tipo = "M" Then
                    
                        WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        
                        If Val(WLote) = Val(nrolote) Then
                    
                            WArticulo = rstMovvar!Articulo
                            WCantidad = rstMovvar!Cantidad
                            WFecha = rstMovvar!Fecha
                            WCodigo = rstMovvar!Codigo
                            WMovi = rstMovvar!Movi
                            WTipomov = Val(rstMovvar!Tipomov)
                            WObservaciones = rstMovvar!Observaciones
                            WSaldo = 0
                        
                            If WMovi = "S" Then
                                
                                With rstFichaMat
                            
                                    .AddNew
                                    !Articulo = WArticulo
                                    !Fecha = WFecha
                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Tipo = 0
                                    !Numero = WCodigo
                                    !Inicial = 0
                                    !Entrada = WCantidad
                                    !Salida = 0
                                    !Observaciones = WObservaciones
                                    !Descripcion = WDesArticulo
                                    If WTipomov = 0 Or WTipomov = 1 Then
                                        !Lista1 = "Mov.Var"
                                            Else
                                        !Lista1 = "Guia In"
                                    End If
                                    !Lista2 = ""
                                    !Lote = WLote
                                    !Saldo = WSaldo
                                    !Empresa = NombreEmpresa
                                    !PartiOri = WPartiOri
                                    .Update
                                End With
                                
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
            rstMovvar.Close
        End If
        
        
        
        
        
        
        
        
        
        XParam = "'" + WArticulo + "','" _
             + WArticulo + "'"

        spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then

            With rstEstadistica

                .MoveFirst
        
                If .NoMatch = False Then
                Do
        
                    If .EOF = True Then
                        Exit Do
                    End If
            
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
            
                        If rstEstadistica!Tipo = 1 Then
                    
                            ZZLote(1) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                            ZZLote(2) = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                            ZZLote(3) = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                            ZZLote(4) = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                            ZZLote(5) = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                            
                            ZZCanti(1) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                            ZZCanti(2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                            ZZCanti(3) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                            ZZCanti(4) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                            ZZCanti(5) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                                
                            WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                    
                            If Len(Trim(WLoteAdicional)) = 98 Then
                                ZZLote(6) = Mid$(WLoteAdicional, 1, 8)
                                ZZCanti(6) = Mid$(WLoteAdicional, 9, 6)
                                ZZLote(7) = Mid$(WLoteAdicional, 15, 8)
                                ZZCanti(7) = Mid$(WLoteAdicional, 23, 6)
                                ZZLote(8) = Mid$(WLoteAdicional, 29, 8)
                                ZZCanti(8) = Mid$(WLoteAdicional, 37, 6)
                                ZZLote(9) = Mid$(WLoteAdicional, 43, 8)
                                ZZCanti(9) = Mid$(WLoteAdicional, 51, 6)
                                ZZLote(10) = Mid$(WLoteAdicional, 57, 8)
                                ZZCanti(10) = Mid$(WLoteAdicional, 65, 6)
                                ZZLote(11) = Mid$(WLoteAdicional, 71, 8)
                                ZZCanti(11) = Mid$(WLoteAdicional, 79, 6)
                                ZZLote(12) = Mid$(WLoteAdicional, 85, 8)
                                ZZCanti(12) = Mid$(WLoteAdicional, 93, 6)
                                    Else
                                ZZLote(6) = "0"
                                ZZCanti(6) = "0"
                                ZZLote(7) = "0"
                                ZZCanti(7) = "0"
                                ZZLote(8) = "0"
                                ZZCanti(8) = "0"
                                ZZLote(9) = "0"
                                ZZCanti(9) = "0"
                                ZZLote(10) = "0"
                                ZZCanti(10) = "0"
                                ZZLote(11) = "0"
                                ZZCanti(11) = "0"
                                ZZLote(12) = "0"
                                ZZCanti(12) = "0"
                            End If
                        
                            For ZZCiclo = 1 To 12
                        
                                If Val(ZZLote(ZZCiclo)) = Val(nrolote) Then
            
                                    WFecha = rstEstadistica!Fecha
                                    WCodigo = rstEstadistica!Numero
                                    WObservaciones = rstEstadistica!Cliente
                                    WTipo = rstEstadistica!Tipo
                                    WCantidad = Val(ZZCanti(ZZCiclo))
                                    WPrecio = rstEstadistica!PrecioUs
                                    
                                    With rstFichaMat
                                        .AddNew
                                        !Articulo = WArticulo
                                        !Fecha = WFecha
                                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        !Tipo = 0
                                        !Numero = WCodigo
                                        !Inicial = 0
                                        !Entrada = WCantidad
                                        !Salida = WPrecio
                                        !Lista1 = "Fact"
                                        !Observaciones = WObservaciones
                                        !Descripcion = ""
                                        !Lista2 = ""
                                        !Lote = nrolote
                                        !Saldo = 0
                                        !Empresa = NombreEmpresa
                                        !PartiOri = WPartiOri
                                        .Update
                                    End With
                                End If
                            Next ZZCiclo
                
                        End If
            
                    End If
        
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
            
                Loop
                End If
        
            End With
            rstEstadistica.Close
        End If
    
    
    
    
    
    
    
    
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                        WFecha = rstLaudo!Fecha
                        WCodigo = rstLaudo!Laudo
                        WLiberada = rstLaudo!Liberada
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        If WLiberadaAnt <> 0 Then
                            WCantidad = WLiberadaAnt
                                Else
                            WCantidad = WLiberada
                        End If
                        With rstFichaMat
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = WCantidad
                            !Entrada = 0
                            !Salida = 0
                            !Lista1 = ""
                            !Observaciones = ""
                            !Descripcion = ""
                            !Lista2 = ""
                            !Lote = nrolote
                            !Saldo = 0
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                        End With
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLaudo.Close
        End If
        
        
        
        
        
    
        XParam = "'" + WArticulo + "','" _
                     + WArticulo + "'"
        
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
                    
                    If rstMovguia!Tipo = "M" Then
                
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WCantidadAnt = IIf(IsNull(rstMovguia!Cantidadant), "0", rstMovguia!Cantidadant)
                        If WCantidadAnt <> 0 Then
                            WCantidad = WCantidadAnt
                        End If
                        
                        Rem WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        Rem WObservaciones = rstMovvar!Observaciones
                            
                        If WMovi = "E" Then
                            WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            WSaldo = 0
                                Else
                            WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            WSaldo = 0
                        End If
    
                            
                        If WMovi = "S" Then
                            Select Case WDestino
                                Case 1
                                    WObservaciones = "Envio a Surfactan"
                                Case 2
                                    WObservaciones = "Envio a Pellital"
                                Case 3
                                    WObservaciones = "Envio a Surfactan II"
                                Case 4
                                    WObservaciones = "Envio a Pellital II"
                                Case 5
                                    WObservaciones = "Envio a Surfactan III"
                                Case 6
                                    WObservaciones = "Envio a Surfactan IV"
                                Case 7
                                    WObservaciones = "Envio a Surfactan V"
                                Case 8
                                    WObservaciones = "Envio a Pellital V"
                                Case 9
                                    WObservaciones = "Envio a Pellital IV"
                                Case 10
                                    WObservaciones = "Envio a Surfactan VI"
                                Case 11
                                    WObservaciones = "Envio a Surfactan VII"
                                Case Else
                            End Select
                                
                                    Else
                                    
                            Select Case WTipomov
                                Case 1
                                    WObservaciones = "Recepcion de Surfactan"
                                Case 2
                                    WObservaciones = "Recepcion de Pellital"
                                Case 3
                                    WObservaciones = "Recepcion de Surfactan II"
                                Case 4
                                    WObservaciones = "Recepcion de Pellital II"
                                Case 5
                                    WObservaciones = "Recepcion de Surfactan III"
                                Case 6
                                    WObservaciones = "Recepcion de Surfactan IV"
                                Case 7
                                    WObservaciones = "Recepcion de Surfactan V"
                                Case 8
                                    WObservaciones = "Recepcion de Pellital V"
                                Case 9
                                    WObservaciones = "Recepcion de Pellital IV"
                                Case 10
                                    WObservaciones = "Recepcion de Surfactan VI"
                                Case 11
                                    WObservaciones = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                                
                        End If
                        
                        With rstFichaMat
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = Right$(WCodigo, 6)
                            !Inicial = WCantidad
                            !Entrada = 0
                            !Salida = 0
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            !Lista1 = "Guia In"
                            !Lista2 = ""
                            !Lote = nrolote
                            !Saldo = 0
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
                            .Update
                            
                        End With
                        
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
        
            
            
            
            
            
        
        
        
        
        
        
        
        
        Da = 0
        With rstFichaMat
            .Index = "Articulo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    .Edit
                    WArticulo = !Articulo
                    WObservaciones = !Observaciones
                    WDescripcion = ""
                    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                    !Descripcion = WDescripcion
                    !Observaciones = WObservaciones
                    .Update
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
        Listado.ReportFileName = "WLotematCuadro.rpt"
    
        Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        Listado.Destination = 0
        Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
        
        Listado.Action = 1

    End If
End Sub


Private Sub Lista_Click()
    Desdecodigo.Text = "AA-000-000"
    HastaCodigo.Text = "ZZ-999-999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desdecodigo.SetFocus
End Sub

Private Sub DesdeCodigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCodigo.SetFocus
    End If
End Sub

Private Sub HastaCodigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdecodigo.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastafecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo1 = "S"
        Unidad.SetFocus
    End If
End Sub


Private Sub Unidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo2 = "S"
        Densidad.SetFocus
    End If
End Sub

Private Sub Densidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo3 = "S"
        CodigoDy.SetFocus
    End If
End Sub

Private Sub CodigoDy_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo23 = "S"
        Deposito.SetFocus
    End If
End Sub

Private Sub Deposito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo4 = "S"
        Minimo.SetFocus
    End If
End Sub

Private Sub Minimo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo5 = "S"
        Minimo.Text = Pusing("###,###.##", Minimo.Text)
        Minimo1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo6 = "S"
        Minimo1.Text = Pusing("###,###.##", Minimo1.Text)
        Controla.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Controla_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo7 = "S"
        CodSedronar.SetFocus
    End If
End Sub

Private Sub CodSedronar_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rs.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Rs_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo8 = "S"
        Leyenda.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Leyenda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo9 = "S"
        Flete.SetFocus
    End If
End Sub

Private Sub Leyenda_Click()
    ZCampo9 = "S"
    If PasaLeyenda <> "N" Then
        Flete.SetFocus
    End If
End Sub

Private Sub Flete_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo10 = "S"
        Flete.Text = Pusing("###,###.##", Flete.Text)
        Moneda.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Moneda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo11 = "S"
        Costo1.SetFocus
    End If
End Sub

Private Sub Costo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo12 = "S"
        Costo1.Text = Pusing("###,###.##", Costo1.Text)
        Costo2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Costo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo13 = "S"
        Costo2.Text = Pusing("###,###.##", Costo2.Text)
        
        Rem spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstArticulo.RecordCount > 0 Then
        Rem     ZCostoAnterior = rstArticulo!Costo2
        Rem     rstArticulo.Close
        Rem End If
        
        Rem If ZCostoAnterior <> Val(Costo2.Text) Then
        Rem     TituloStd.Caption = "Costo Estandar U$S"
       Rem  End If
        
        Costo4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Costo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo3.Text = Pusing("###,###.##", Costo3.Text)
        Costo4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Costo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo4.Text = Pusing("###,###.##", Costo4.Text)
        WCosto1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCosto1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo15 = "S"
        ZCampo16 = "S"
        ZCampo17 = "S"
        ZCampo18 = "S"
        WCosto1.Text = Pusing("###,###.##", WCosto1.Text)
        Meses.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCosto2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo16 = "S"
        ZCampo17 = "S"
        ZCampo18 = "S"
        WCosto2.Text = Pusing("###,###.##", WCosto2.Text)
        Rem WCosto3.SetFocus
        Meses.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCosto3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo17 = "S"
        ZCampo18 = "S"
        WCosto3.Text = Pusing("###,###.##", WCosto3.Text)
        Meses.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Meses_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo24 = "S"
        Naciones.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo18 = "S"
        If Val(Envase.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                DesEnvase.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Naciones.SetFocus
            End If
                Else
            Envase.Text = ""
            DesEnvase.Caption = ""
            Clase.SetFocus
        End If
    End If
End Sub

Private Sub Naciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(Naciones.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
                rstPeligroso.Close
                
                ZLugar = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Peligroso"
                ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
                ZSql = ZSql + " Order by Peligroso.Codigo"
                spPeligroso = ZSql
                Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
                If rstPeligroso.RecordCount > 0 Then
        
                    With rstPeligroso
                    
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                                                            
                                ZLugar = ZLugar + 1
                                ZPeligrosoI = rstPeligroso!Ficha
                                ZPeligrosoII = rstPeligroso!Descripcion
                                ZPeligrosoIII = rstPeligroso!Clase
                                ZPeligrosoIV = rstPeligroso!Secundario
                                ZPeligrosoV = rstPeligroso!Riesgo
                                ZPeligrosoVI = rstPeligroso!Embalaje
                                
                                .MoveNext
                                
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                            Loop
                        End If
                        
                    End With
                    rstPeligroso.Close
        
                End If
                
                If ZLugar = 1 Then
                
                    Pantalla.Visible = False
                
                    Clase.Text = Trim(ZPeligrosoIII)
                    Rem Secundario.Text = Trim(ZPeligrosoIV)
                    Caracteristicas.Text = Left$(ZPeligrosoII, 100)
                    Intervencion.Text = Trim(ZPeligrosoI)
                    Rem Riesgo.Text = Trim(ZPeligrosoV)
                    Embalaje.Text = Trim(ZPeligrosoVI)
                    
                        Else
                        
                    Opcion.Clear
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Rem Opcion.Visible = True
                    Opcion.ListIndex = 4
                        
                End If
                
                    Else
                    
                Clase.Text = ""
                Rem Secundario.Text = ""
                Rem Riesgo.Text = ""
                Embalaje.Text = ""
                Caracteristicas.Text = ""
                Intervencion.Text = ""
            
                m$ = "Nro de Naciones Unidas Inexistente"
                A% = MsgBox(m$, 0, "Archivo de Productos Terminados")
                Exit Sub
                
            End If
            
                Else
                
            Clase.Text = ""
            Embalaje.Text = ""
            Intervencion.Text = ""
            
        End If
    
        ZCampo19 = "S"
        ZCampo20 = "S"
        ZCampo21 = "S"
        ZCampo22 = "S"
        Descripcion.SetFocus
        
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
    If KeyAscii <> 13 Then
    
        ZAyuda = ""
        If KeyAscii > 31 Then
            ZAyuda = Naciones.Text + Chr$(KeyAscii)
                Else
            Select Case KeyAscii
                Case 27
                    Naciones.Text = ""
                    ZAyuda = ""
                Case 8
                    WEspacios = Len(Naciones.Text)
                    If WEspacios > 0 Then
                        WEspacios = WEspacios - 1
                        ZAyuda = Left$(Naciones.Text, WEspacios)
                    End If
                Case Else
                    ZAyuda = Naciones.Text
            End Select
        End If
    
        Clase.Text = ""
        Rem Secundario.Text = ""
        Rem Riesgo.Text = ""
        Embalaje.Text = ""
        Caracteristicas.Text = ""
        Intervencion.Text = ""
        
        ZLugar = 0
        If Trim(ZAyuda) <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + ZAyuda + "'"
            ZSql = ZSql + " Order by Peligroso.Codigo"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
    
                With rstPeligroso
                
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                                                        
                            ZLugar = ZLugar + 1
                            ZPeligrosoI = rstPeligroso!Ficha
                            ZPeligrosoII = rstPeligroso!Descripcion
                            ZPeligrosoIII = rstPeligroso!Clase
                            ZPeligrosoIV = rstPeligroso!Secundario
                            ZPeligrosoV = rstPeligroso!Riesgo
                            ZPeligrosoVI = rstPeligroso!Embalaje
                            
                            .MoveNext
                            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                        Loop
                    End If
                    
                End With
                rstPeligroso.Close
    
            End If
            
            If ZLugar = 1 Then
            
                Pantalla.Visible = False
                        
                Clase.Text = Trim(ZPeligrosoIII)
                Rem Secundario.Text = Trim(ZPeligrosoIV)
                Caracteristicas.Text = Left$(ZPeligrosoII, 100)
                Intervencion.Text = Trim(ZPeligrosoI)
                Rem Riesgo.Text = Trim(ZPeligrosoV)
                Embalaje.Text = Trim(ZPeligrosoVI)
                
                    Else
                    
                Opcion.Clear
                Opcion.AddItem ""
                Opcion.AddItem ""
                Opcion.AddItem ""
                Opcion.AddItem ""
                Opcion.AddItem ""
                Opcion.AddItem ""
                Opcion.AddItem ""
                Rem Opcion.Visible = True
                Opcion.ListIndex = 5
                    
            End If
            
        End If
        
    End If
    
End Sub

Private Sub Clase_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo120 = "S"
        Intervencion.SetFocus
    End If
End Sub

Private Sub Intervencion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo21 = "S"
        Embalaje.SetFocus
    End If
End Sub

Private Sub Embalaje_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo22 = "S"
        Descripcion.SetFocus
    End If
End Sub

Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount <= 0 Then
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
                Descripcion.SetFocus
                        Else
                Codigo.Text = rstArticulo!Codigo
                rstArticulo.Close
                Call Imprime_Datos
                Descripcion.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Materias Primas"
     Opcion.AddItem "Envases"
     Opcion.AddItem "Marcas"
     Opcion.AddItem "Naciones Unidas"

     Opcion.Visible = True
     
 End Sub

 Private Sub Opcion_Click()
    
  sal = 1
    Opcion.Visible = False
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        sal = 0
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
        
        Case 1
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
                       sal = 0
                        IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstEnvases!Envases
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnvases.Close
            
        Case 2
            spMarcas = "ListaMarcas"
            Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
            With rstMarcas
                .MoveFirst
                Do
                    If .EOF = False Then
                        sal = 0
                        IngresaItem = rstMarcas!Articulo + " " + rstMarcas!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstMarcas!Articulo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstMarcas.Close
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
            ZSql = ZSql + " Order by Peligroso.Ficha, Peligroso.Descripcion"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
    
                With rstPeligroso
                
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                             sal = 0
                            IngresaItem = Trim(rstPeligroso!Ficha) + " " + rstPeligroso!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstPeligroso!Codigo
                            WIndice.AddItem IngresaItem
                            
                            .MoveNext
                            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                        Loop
                    End If
                    
                End With
                rstPeligroso.Close
    
            End If
            
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + ZAyuda + "'"
            ZSql = ZSql + " Order by Peligroso.Ficha, Peligroso.Descripcion"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
    
                With rstPeligroso
                
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            sal = 0
                            IngresaItem = Trim(rstPeligroso!Ficha) + " " + rstPeligroso!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstPeligroso!Codigo
                            WIndice.AddItem IngresaItem
                            
                            .MoveNext
                            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                        Loop
                    End If
                    
                End With
                rstPeligroso.Close
    
            End If
        
        Case 6
        
            Erase WWVectorII
            WWLugarII = 0
            
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Select Case Val(Wempresa)
                        Case 1
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
                            HastaEmpre = 7
                        Case 3
                            CargaEmpresa(1, 1) = "0003"
                            CargaEmpresa(1, 2) = "Empresa03"
                            HastaEmpre = 1
                        Case 5
                            CargaEmpresa(1, 1) = "0005"
                            CargaEmpresa(1, 2) = "Empresa05"
                            HastaEmpre = 1
                        Case 6
                            CargaEmpresa(1, 1) = "0006"
                            CargaEmpresa(1, 2) = "Empresa06"
                            HastaEmpre = 1
                        Case 7
                            CargaEmpresa(1, 1) = "0007"
                            CargaEmpresa(1, 2) = "Empresa07"
                            HastaEmpre = 1
                        Case 10
                            CargaEmpresa(1, 1) = "0010"
                            CargaEmpresa(1, 2) = "Empresa10"
                            HastaEmpre = 1
                        Case 11
                            CargaEmpresa(1, 1) = "0011"
                            CargaEmpresa(1, 2) = "Empresa11"
                            HastaEmpre = 1
                        Case Else
                    End Select
                        
            
                Case 2, 4, 8, 9
                    Select Case Val(Wempresa)
                        Case 2
                            CargaEmpresa(1, 1) = "0002"
                            CargaEmpresa(1, 2) = "Empresa02"
                            CargaEmpresa(2, 1) = "0004"
                            CargaEmpresa(2, 2) = "Empresa04"
                            CargaEmpresa(3, 1) = "0009"
                            CargaEmpresa(3, 2) = "Empresa09"
                            CargaEmpresa(4, 1) = "0009"
                            CargaEmpresa(4, 2) = "Empresa09"
                            HastaEmpre = 4
                        Case 4
                            CargaEmpresa(1, 1) = "0004"
                            CargaEmpresa(1, 2) = "Empresa04"
                            HastaEmpre = 1
                        Case 8
                            CargaEmpresa(1, 1) = "0008"
                            CargaEmpresa(1, 2) = "Empresa08"
                            HastaEmpre = 1
                        Case 9
                            CargaEmpresa(1, 1) = "0009"
                            CargaEmpresa(1, 2) = "Empresa09"
                            HastaEmpre = 1
                        Case Else
                    End Select
            
                Case Else
            End Select
        
        
                            
            For Cicla = 1 To HastaEmpre
            
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Cantidad > Orden.Recibida"
                    ZSql = ZSql + " and Orden.Tipo = 1"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + Codigo.Text + "'"
                    ZSql = ZSql + " Order by Clave"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                
                        With rstOrden
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                    sal = 0
                                    IngresaItem = Str$(rstOrden!Carpeta) + "  " + Str$(rstOrden!Orden) + "  " + rstOrden!Fecha
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstOrden!Carpeta
                                    WIndice.AddItem IngresaItem
                                    
                                    WWLugarII = WWLugarII + 1
                                    WWVectorII(WWLugarII, 1) = rstOrden!Carpeta
                                    WWVectorII(WWLugarII, 2) = rstOrden!Orden
                                    
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstOrden.Close
                
                    End If
        
                End If
            Next Cicla
            
        Case Else
    End Select
    If sal = 1 Then
    Pantalla.Visible = False
       Else
     Pantalla.Visible = True
       
       
       End If
    
    Opcion.Visible = False

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0, 2
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Codigo.Text = rstArticulo!Codigo
                rstArticulo.Close
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Codigo.Text = WArticulo
            End If
            
            Ayuda.Visible = False
            Rem Codigo.SetFocus
        
        Case 1
            Rem Indice = Pantalla.ListIndex
            Rem WEnvases = WIndice.List(Indice)
            Rem spEnvases = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
            Rem Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstEnvases.RecordCount > 0 Then
            Rem     Envase.Text = rstEnvases!Envases
            Rem     DesEnvase.Caption = rstEnvases!Descripcion
            Rem     rstEnvases.Close
            Rem End If
            Rem Envase.SetFocus
            
        Case 4, 5
            Indice = Pantalla.ListIndex
            ZZCodigoOnu = WIndice.List(Indice)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.Codigo = " + "'" + ZZCodigoOnu + "'"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
                Naciones.Text = Trim(rstPeligroso!nroonu)
                Clase.Text = Trim(rstPeligroso!Clase)
                Rem Secundario.Text = Trim(rstPeligroso!Secundario)
                Rem Riesgo.Text = Trim(rstPeligroso!Riesgo)
                Embalaje.Text = Trim(rstPeligroso!Embalaje)
                Caracteristicas.Text = Left$(rstPeligroso!Descripcion, 100)
                Intervencion.Text = Trim(rstPeligroso!Ficha)
                rstPeligroso.Close
            End If
            
        Case 6
            Indice = Pantalla.ListIndex + 1
            WPasaCarpeta = WWVectorII(Indice, 1)
            WPasaOrden = WWVectorII(Indice, 2)
            PrgOrdenComplementoConsulta.Show
            
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstArticulo
        .MoveFirst
        Codigo.Text = rstArticulo!Codigo
    End With
    
    rstArticulo.Close
    
    Call Imprime_Datos
    Rem Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Articulo", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()

    On Error GoTo Error_ultimo
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstArticulo
        .MoveLast
        Codigo.Text = rstArticulo!Codigo
    End With
    
    rstArticulo.Close
    Call Imprime_Datos
    Rem Codigo.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Articulo", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Articulo.SetFocus
 End Sub

Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spArticulo = "PosteriorArticulo " + "'" + Codigo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstArticulo
        .MoveFirst
        Codigo.Text = rstArticulo!Codigo
    End With
    
    rstArticulo.Close
    Call Imprime_Datos
    Rem Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Articulo", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
     
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    Clave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave.Visible = False
    Codigo.SetFocus

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        WGraba = "N"
        ZGRABAII = ""
        
        XEmpresa = Wempresa
        
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAi = IIf(IsNull(rstOperador!GrabaI), "", rstOperador!GrabaI)
            Responsable = rstOperador!Descripcion
            rstOperador.Close
        End If
        
        Call Conecta_Empresa
        
        If ZGRABAi = "S" Then
            WGraba = "S"
            Clave.Visible = False
            If WProceso = 0 Then
                Call cmdAdd_Click
                    Else
                If WProceso = 2 Then
                    Call GrabaMarcas_Click
                        Else
                    Call cmdDelete_Click
                End If
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Materia Prima")
            WClave.SetFocus
        End If

    End If

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                spArticulo = "ListaArticuloConsulta"
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
                                        Exit For
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
                
            Case 2
                spMarcas = "ListaMarcasConsulta " + "'" + Ayuda.Text + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    With rstMarcas
                        .MoveFirst
                        Do
                            If .EOF = False Then
            
                                Da = Len(rstMarcas!Descripcion) - WEspacios
                
                                For Aaa = 1 To Da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstMarcas!Descripcion, Aaa, WEspacios) Then
                                        IngresaItem = rstMarcas!Articulo + " " + rstMarcas!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstMarcas!Articulo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next Aaa
                                .MoveNext
                    
                                    Else
                        
                                Exit Do
                
                            End If
                        Loop
                    End With
    
                    rstMarcas.Close
    
                End If
            Case Else
        End Select
    
    End If
    If KeyAscii = 27 Then
        Pantalla.Visible = False
        Opcion.Visible = False
        Ayuda.Visible = False
    End If

End Sub

Private Sub ocpend_Click()
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Select Case Val(Wempresa)
                Case 1
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
                    HastaEmpre = 7
                Case 3
                    CargaEmpresa(1, 1) = "0003"
                    CargaEmpresa(1, 2) = "Empresa03"
                    HastaEmpre = 1
                Case 5
                    CargaEmpresa(1, 1) = "0005"
                    CargaEmpresa(1, 2) = "Empresa05"
                    HastaEmpre = 1
                Case 6
                    CargaEmpresa(1, 1) = "0006"
                    CargaEmpresa(1, 2) = "Empresa06"
                    HastaEmpre = 1
                Case 7
                    CargaEmpresa(1, 1) = "0007"
                    CargaEmpresa(1, 2) = "Empresa07"
                    HastaEmpre = 1
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                    HastaEmpre = 1
                Case 11
                    CargaEmpresa(1, 1) = "0011"
                    CargaEmpresa(1, 2) = "Empresa11"
                    HastaEmpre = 1
                Case Else
            End Select
                
    
        Case 2, 4, 8, 9
            Select Case Val(Wempresa)
                Case 2
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0009"
                    CargaEmpresa(3, 2) = "Empresa09"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    HastaEmpre = 4
                Case 4
                    CargaEmpresa(1, 1) = "0004"
                    CargaEmpresa(1, 2) = "Empresa04"
                    HastaEmpre = 1
                Case 8
                    CargaEmpresa(1, 1) = "0008"
                    CargaEmpresa(1, 2) = "Empresa08"
                    HastaEmpre = 1
                Case 9
                    CargaEmpresa(1, 1) = "0009"
                    CargaEmpresa(1, 2) = "Empresa09"
                    HastaEmpre = 1
                Case Else
            End Select
    
        Case Else
    End Select


                    
    For Cicla = 1 To HastaEmpre
    
        If CargaEmpresa(Cicla, 1) <> "" Then
        
        
        
            MensajeI.Text = "ORDEN DE COMPRA PENDIENTE"
            Select Case Val(CargaEmpresa(Cicla, 1))
                Case 1
                    WObservaciones = "Surfactan I"
                Case 2
                    WObservaciones = "Pellital I"
                Case 3
                    WObservaciones = "Surfactan II"
                Case 4
                    WObservaciones = "Pellital II"
                Case 5
                    WObservaciones = "Surfactan III"
                Case 6
                    WObservaciones = "Surfactan IV"
                Case 7
                    WObservaciones = "Surfactan V"
                Case 8
                    WObservaciones = "Pellital V"
                Case 9
                    WObservaciones = "Pellital IV"
                Case 10
                    WObservaciones = "Surfactan VI"
                Case 11
                    WObservaciones = "Surfactan VII"
                Case Else
            End Select
            MensajeII.Text = WObservaciones
        
            PantaMensaje.Visible = True
            DoEvents
        
        
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)



        
            XParam = "'" + "'"
        
            spOrden = "ModificaOrdenSaldo " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            Erase WWVector
            WWLugar = 0
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Articulo = " + "'" + Codigo.Text + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                     .MoveFirst
                     Do
                         If .EOF = False Then
                         
                            If !FechaOrd > "20130101" Then
                            
                                WWClave = rstOrden!Clave
                                WWOrden = rstOrden!Orden
                                WWFecha2 = rstOrden!fecha2
                                WWSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                                If Val(WWSaldo) > 0 Then
                                   Entra = "S"
                                   For XX = 1 To WWLugar
                                       If Val(WWVector(XX, 1)) = WWOrden Then
                                           Entra = "N"
                                           Exit For
                                       End If
                                   Next XX
                                   
                                   If Entra = "S" Then
                                       WWLugar = WWLugar + 1
                                       WWVector(WWLugar, 1) = WWOrden
                                       WWVector(WWLugar, 2) = Right$(WWFecha2, 4) + Mid$(WWFecha2, 4, 2) + Left$(WWFecha2, 2)
                                   End If
                                   
                                End If
                                
                            End If
                                
                             .MoveNext
                                 Else
                             Exit Do
                         End If
                     Loop
                End With
                rstOrden.Close
            End If
            
            
            PantaMensaje.Visible = False
            
            If WWLugar > 0 Then
            
                m$ = "Hay ordenes de compra pendientes en " + WObservaciones
                G% = MsgBox(m$, 0, "ORDEN DE COMPRA PENDIENTE")
                
                 For XX = 1 To WWLugar
                     WWOrden = WWVector(XX, 1)
                     WWFecha2 = WWVector(XX, 2)
                     XParam = "'" + WWOrden + "','" _
                                  + WWFecha2 + "'"
                 
                     spOrden = "ModificaOrdenFecha2 " + XParam
                     Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                 Next XX
                 
                 Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
                 Listado.WindowTop = 0
                 Listado.WindowLeft = 0
                 Listado.WindowWidth = Screen.Width
                 Listado.WindowHeight = Screen.Height
                 
                 Uno = "{Orden.Articulo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
                 Dos = " and {Orden.FechaOrd} >= " + Chr$(34) + "20130101" + Chr$(34)
                
                 Listado.GroupSelectionFormula = Uno + Dos
                 Listado.SelectionFormula = Uno + Dos
                
                 Listado.Destination = 0
                 
                 DbConnect = db.Connect
                 DSQ = getDatabase(DbConnect)
                 
                 Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Fecha2, Orden.Saldo, Orden.OrdFecha2, Proveedor.Nombre, Articulo.Descripcion " _
                                     + "From " + DSQ + ".dbo.Orden Orden, " _
                                     + DSQ + ".dbo.Proveedor Proveedor, " _
                                     + DSQ + ".dbo.Articulo Articulo " _
                                     + "Where Orden.Proveedor = Proveedor.Proveedor AND " _
                                     + "Orden.Articulo = Articulo.Codigo AND " _
                                     + "Orden.Articulo >= '" + Codigo.Text + "' AND Orden.Articulo <= '" + Codigo.Text + "' AND Orden.Saldo > 0. AND Orden.OrdFecha2 >= '00000000' AND Orden.OrdFecha2 <= '9999999' and " _
                                     + "Orden.FechaOrd >= '20130101'"
                 
                 Listado.DataFiles(0) = ""
                 Listado.DataFiles(1) = ""
                 Listado.DataFiles(2) = ""
                 Listado.DataFiles(3) = Wempresa + "auxi.mdb"
                 Listado.Connect = Connect()
                 
                 Listado.ReportFileName = "WOrdPenArtII.rpt"
                 Listado.Action = 1

            End If

        End If
    Next Cicla
    
    
    
    
    

End Sub

Private Sub LaboPendiente_Click()


    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Select Case Val(Wempresa)
                Case 1
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
                    HastaEmpre = 7
                Case 3
                    CargaEmpresa(1, 1) = "0003"
                    CargaEmpresa(1, 2) = "Empresa03"
                    HastaEmpre = 1
                Case 5
                    CargaEmpresa(1, 1) = "0005"
                    CargaEmpresa(1, 2) = "Empresa05"
                    HastaEmpre = 1
                Case 6
                    CargaEmpresa(1, 1) = "0006"
                    CargaEmpresa(1, 2) = "Empresa06"
                    HastaEmpre = 1
                Case 7
                    CargaEmpresa(1, 1) = "0007"
                    CargaEmpresa(1, 2) = "Empresa07"
                    HastaEmpre = 1
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                    HastaEmpre = 1
                Case 11
                    CargaEmpresa(1, 1) = "0011"
                    CargaEmpresa(1, 2) = "Empresa11"
                    HastaEmpre = 1
                Case Else
            End Select
                
    
        Case 2, 4, 8, 9
            Select Case Val(Wempresa)
                Case 2
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0009"
                    CargaEmpresa(3, 2) = "Empresa09"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    HastaEmpre = 4
                Case 4
                    CargaEmpresa(1, 1) = "0004"
                    CargaEmpresa(1, 2) = "Empresa04"
                    HastaEmpre = 1
                Case 8
                    CargaEmpresa(1, 1) = "0008"
                    CargaEmpresa(1, 2) = "Empresa08"
                    HastaEmpre = 1
                Case 9
                    CargaEmpresa(1, 1) = "0009"
                    CargaEmpresa(1, 2) = "Empresa09"
                    HastaEmpre = 1
                Case Else
            End Select
    
        Case Else
    End Select


                    
    For Cicla = 1 To HastaEmpre
    
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            MensajeI.Text = "PENDIENTE DE LABORATORIO"
            Select Case Val(CargaEmpresa(Cicla, 1))
                Case 1
                    WObservaciones = "Surfactan I"
                Case 2
                    WObservaciones = "Pellital I"
                Case 3
                    WObservaciones = "Surfactan II"
                Case 4
                    WObservaciones = "Pellital II"
                Case 5
                    WObservaciones = "Surfactan III"
                Case 6
                    WObservaciones = "Surfactan IV"
                Case 7
                    WObservaciones = "Surfactan V"
                Case 8
                    WObservaciones = "Pellital V"
                Case 9
                    WObservaciones = "Pellital IV"
                Case 10
                    WObservaciones = "Surfactan VI"
                Case 11
                    WObservaciones = "Surfactan VII"
                Case Else
            End Select
            MensajeII.Text = WObservaciones
        
            PantaMensaje.Visible = True
            DoEvents
            
        
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)



        
        
        
        
            Erase WWVector
            WWRenglon = 0
            
            spInforme = "ModificaInformeProcesoSaldo"
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            
            XParam = "'" + "20020101" + "'"
            spInforme = "ModificaInformeProceso0 " + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            
            XParam = "'" + Codigo.Text + "'"
            spInforme = "ListaInformeArticulo " + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                With rstInforme
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If !FechaOrd > "20130101" Then
                                If rstInforme!Articulo = Codigo.Text Then
                                    WWRenglon = WWRenglon + 1
                                    WWVector(WWRenglon, 1) = rstInforme!Clave
                                    WWVector(WWRenglon, 2) = rstInforme!Informe
                                    WWVector(WWRenglon, 3) = rstInforme!Articulo
                                    WWVector(WWRenglon, 4) = rstInforme!Cantidad
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstInforme.Close
            End If
            
            For Ciclo = 1 To WWRenglon
            
                WClave = WWVector(Ciclo, 1)
                WInforme = WWVector(Ciclo, 2)
                WArticulo = WWVector(Ciclo, 3)
                WCantidad = Val(WWVector(Ciclo, 4))
                WResta = 0
            
                XParam = "'" + WInforme + "','" _
                         + WArticulo + "'"
                spLaudo = "ListaLaudoInforme " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
            
                    With rstLaudo
            
                        .MoveFirst
                    
                        If .NoMatch = False Then
                            Do
                    
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                                WLiberada = rstLaudo!Liberada
                                WDevuelta = rstLaudo!devuelta
                                WSuma = WLiberada + WDevuelta
                                
                                WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                                WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                                WSumaAnt = WLiberadaAnt + WDevueltaAnt
                                
                                If WSumaAnt <> 0 Then
                                    WResta = WResta + WSumaAnt
                                        Else
                                    WResta = WResta + WSuma
                                End If
                                
                                Rem WResta = WResta + rstLaudo!Liberada + rstLaudo!Devuelta
                        
                                .MoveNext
                        
                                If .EOF = True Then
                                    Exit Do
                                End If
                        
                            Loop
                        End If
                    End With
                    rstLaudo.Close
                End If
                
                XParam = "'" + WClave + "','" _
                         + Str$(WResta) + "'"
                spInforme = "ModificaInformeProceso " + XParam
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
            
            spInforme = "ModificaInformeProcesoDife"
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            PantaMensaje.Visible = False
            
            WDesde = "20130101"
            WHasta = "99999999"
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            Rem BY NAN
            
            ZSql = ZSql + " Where Informe.Articulo = " + "'" + Codigo.Text + "'"
            ZSql = ZSql + " and Informe.Dife <> 0"
            Rem BY NAN AGREGADO
            ZSql = ZSql + "and Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "' "
            Rem BY NAN FIN AGREGADO
            Rem  ZSql = ZSql + " where Informe.Dife <> 0"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                ZZPasa = "S"
                Rem by nan
                Rem   rstInforme.Close
                    Else
                ZZPasa = "N"
            End If
            
            If ZZPasa = "S" Then
            
                m$ = "Hay informes pendientes en " + WObservaciones
                G% = MsgBox(m$, 0, "PENDIENTE DE LABORATORIO")
                
                WDesde = "20130101"
                WHasta = "99999999"
                
                Listado.WindowTitle = "Listado de Informe de Recepcion Pendientes de Aprobacion"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
                
                Rem by nan
                Rem  Codigo.Text = "AA-000-000"
                Rem  Codigo1 = "ZZ-999-999'"
                Listado.GroupSelectionFormula = "{Informe.Articulo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34) + " and {Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
                Rem   Listado.GroupSelectionFormula = "{Informe.Articulo} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo1 + Chr$(34) + " and {Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
                
                Listado.Destination = 0
                
                DbConnect = db.Connect
                DSQ = getDatabase(DbConnect)
                
                Listado.SQLQuery = "SELECT Informe.Informe, Informe.Fecha, Informe.Remito, Informe.Proveedor, Informe.Orden, Informe.Articulo, Informe.Cantidad, Informe.Fechaord, Informe.CantidadLaudo, Informe.Dife, " _
                                    + "Articulo.Descripcion, " _
                                    + "Proveedor.Nombre " _
                                    + "From " _
                                    + DSQ + ".dbo.Informe Informe, " _
                                    + DSQ + ".dbo.Articulo Articulo, " _
                                    + DSQ + ".dbo.Proveedor Proveedor " _
                                    + "Where " _
                                    + "Informe.Articulo = Articulo.Codigo AND " _
                                    + "Informe.Proveedor = Proveedor.Proveedor AND " _
                                    + "Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "' AND " _
                                    + "Informe.Dife <> 0."
                                    
                Listado.DataFiles(0) = ""
                Listado.DataFiles(1) = ""
                Listado.DataFiles(2) = ""
                Listado.DataFiles(3) = Wempresa + "auxi.mdb"
                Listado.Connect = Connect()
                
                WListado = Listado.ReportFileName
                Listado.ReportFileName = "Wlistinfpend.rpt"
                Listado.Action = 1
                Listado.ReportFileName = WListado
            
            End If
            
            Rem agregao by nan
            rstInforme.Close
            Rem fin
        
        End If
        
    Next Cicla
        

End Sub

Private Sub PedPen_Click()

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + Codigo.Text + "','" _
                 + Codigo.Text + "'"
    spPedido = "ModificaPedpenDy " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = Left$(WAuxiliar, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Pedidos Pendientes de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Tipoped, " _
                    + "Cliente.Razon, " _
                    + "Articulo.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Articulo Articulo " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.Articulo = Articulo.Codigo AND " _
                    + "Pedido.Importe > 0"
    
    Listado.DataFiles(0) = ""
    Listado.DataFiles(1) = ""
    Listado.DataFiles(2) = ""
    Listado.DataFiles(3) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Destination = 0
    WListado = Listado.ReportFileName
    Listado.ReportFileName = "WPedpenDy.rpt"
    Listado.GroupSelectionFormula = ""
    Listado.Action = 1
    Listado.ReportFileName = WListado
    
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
        Case 3
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
            WVector1.Col = XColumna
            
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
    WVector1.Cols = 4
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
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
                WVector1.Text = "Proveedor"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Razon Social"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Nombre Comercial"
                WVector1.ColWidth(Ciclo) = 4100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
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

Private Sub GrabaMarcas_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    XEmpresa = Wempresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
                    
    For Cicla = 1 To 11
        If CargaEmpresa(Cicla, 1) <> "" Then
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    If WSalidaError = "N" Then Exit Sub

    If WGraba <> "S" Then
        WProceso = 2
        Call Ingresa_clave
            Else
        WGraba = ""
        For Ciclo = 1 To 1000
            WVector1.Row = Ciclo
            WVector1.Col = 1
            If WVector1.Text <> "" Then
            
                XEmpresa = Wempresa
            
                XProveedor = WVector1.Text
                Call Ceros(XProveedor, 11)
                ClaveMarcas = Codigo.Text + XProveedor
                WVector1.Col = 3
                WNombreComercial = WVector1.Text
                
                XParam = "'" + ClaveMarcas + "','" _
                            + Codigo.Text + "','" _
                            + XProveedor + "','" _
                            + WNombreComercial + "'"
                                         
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                
        
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                If rstMarcas.RecordCount > 0 Then
                    rstMarcas.Close
                    spMarcas = "ModificaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    spMarcas = "AltaMarcas " + XParam
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                Call Conecta_Empresa
    
            End If
        Next Ciclo
        Call FinConsulta_Click
    End If
    
    Exit Sub

Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoErrorII.Visible = True
    Resume Next
    
 End Sub

Private Sub Costo1_dblclick()

    spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
        WPasaCarpeta = IIf(IsNull(rstArticulo!Carpeta), "0", rstArticulo!Carpeta)
        WPasaUltimoFob = IIf(IsNull(rstArticulo!UltimoFob), "0", rstArticulo!UltimoFob)
        WPasaFactor = IIf(IsNull(rstArticulo!Factor), "0", rstArticulo!Factor)
        WPasaUltimoCosto = IIf(IsNull(rstArticulo!UltimoCosto), "0", rstArticulo!UltimoCosto)
        WPasaUltimoTipo = IIf(IsNull(rstArticulo!UltimoTipo), "0", rstArticulo!UltimoTipo)
        
        rstArticulo.Close
    End If

    PrgArtiComple.Show
    
End Sub

Private Sub Proveedor_DblClick()
    ZZPasaProveedor = Proveedor.Text
    PrgProveConsulta.Show
End Sub

Private Sub Command1111_Click()

    Dim ZZTrabajo(6000, 4) As String
    Dim ZZTrabajoII(100, 2) As String
    
    Erase ZZTrabajo
    ZZLugar = 0
    
    Stop
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Order by Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZLugar = ZZLugar + 1
                    ZZTrabajo(ZZLugar, 1) = rstArticulo!Codigo
                    WCarpeta = IIf(IsNull(rstArticulo!Carpeta), "", rstArticulo!Carpeta)
                    ZZTrabajo(ZZLugar, 2) = WCarpeta
                    ZZTrabajo(ZZLugar, 3) = Str$(rstArticulo!Costo1)
                    ZZUltimoCosto = IIf(IsNull(rstArticulo!UltimoCosto), "0", rstArticulo!UltimoCosto)
                    ZZTrabajo(ZZLugar, 4) = Str$(ZZUltimoCosto)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
        
    End If
    
    Rem ZZLugar = 50
    
    
    XEmpresa = Wempresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
    
    For ZZCiclo = 1 To ZZLugar
    
        ZZArticulo = ZZTrabajo(ZZCiclo, 1)
        ZZCarpeta = ZZTrabajo(ZZCiclo, 2)
        ZZCosto1 = Val(ZZTrabajo(ZZCiclo, 3))
        ZZUltimoCosto = Val(ZZTrabajo(ZZCiclo, 4))
        
        Rem If Trim(ZZArticulo) = "AA-031-100" Then Stop
        
        ZZOrdenPesos = 0
        ZZPtaOrdenPesos = 0
        ZZFechaOrdenPesos = ""
        ZZOrdFechaOrdenPesos = ""
        ZZImpoPesos = 0
                        
        ZZOrdenDol = 0
        ZZPtaOrdenDol = 0
        ZZFechaOrdenDol = ""
        ZZOrdFechaOrdenDol = ""
        ZZImpoDolar = 0
                        
        ZZOrdenImpo = 0
        ZZPtaOrdenImpo = 0
        ZZFechaOrdenImpo = ""
        ZZOrdFechaOrdenImpo = ""
        ZZImpoImpo = 0
                        
                        
        For Cicla = 1 To 11
        
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Articulo = " + "'" + ZZArticulo + "'"
            ZSql = ZSql + " and Moneda = " + "'" + "1" + "'"
            ZSql = ZSql + " and (Tipo = '0' or Tipo = '3' or Tipo = '4')"
            ZSql = ZSql + " and Liberada <> 0"
            ZSql = ZSql + " Order by FechaOrd desc"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                If rstOrden!FechaOrd > ZZOrdFechaOrdenPesos Then
                    ZZOrdenPesos = rstOrden!Orden
                    ZZPtaOrdenPesos = Cicla
                    ZZFechaOrdenPesos = rstOrden!Fecha
                    ZZOrdFechaOrdenPesos = rstOrden!FechaOrd
                    ZZImpoPesos = rstOrden!Precio
                End If
                rstOrden.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Articulo = " + "'" + ZZArticulo + "'"
            ZSql = ZSql + " and Moneda = " + "'" + "0" + "'"
            ZSql = ZSql + " and (Tipo = '0' or Tipo = '3' or Tipo = '4')"
            ZSql = ZSql + " and Liberada <> 0"
            ZSql = ZSql + " Order by FechaOrd desc"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                If rstOrden!FechaOrd > ZZOrdFechaOrdenDol Then
                    ZZOrdenDol = rstOrden!Orden
                    ZZPtaOrdenDol = Cicla
                    ZZFechaOrdenDol = rstOrden!Fecha
                    ZZOrdFechaOrdenDol = rstOrden!FechaOrd
                    ZZImpoDolar = rstOrden!Precio
                End If
                rstOrden.Close
            End If
            
            If Aaa = 999 Then
                
                If Val(ZZCarpeta) <> 0 Then
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Articulo = " + "'" + ZZArticulo + "'"
                    ZSql = ZSql + " and Carpeta = " + "'" + ZZCarpeta + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        ZZOrdenImpo = rstOrden!Orden
                        ZZPtaOrdenImpo = Cicla
                        ZZFechaOrdenImpo = rstOrden!Fecha
                        ZZOrdFechaOrdenImpo = rstOrden!FechaOrd
                        ZZImpoImpo = ZZUltimoCosto
                        rstOrden.Close
                    End If
                    
                End If
            
            End If
            
        Next Cicla
        
        
        
        Rem If Trim(ZZArticulo) = "AA-031-100" Then Stop
        
        For Cicla = 1 To 11
        
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql & "UPDATE Articulo SET "
            Rem ZSql = ZSql & "OrdenI = " + "'" + Str$(ZZOrdenImpo) + "',"
            Rem ZSql = ZSql & "PtaOrdenI = " + "'" + Str$(ZZPtaOrdenImpo) + "',"
            ZSql = ZSql & "WCosto1 = " + "'" + Str$(ZZImpoPesos) + "',"
            ZSql = ZSql & "OrdenII = " + "'" + Str$(ZZOrdenPesos) + "',"
            ZSql = ZSql & "PtaOrdenII = " + "'" + Str$(ZZPtaOrdenPesos) + "',"
            ZSql = ZSql & "ZCosto1 = " + "'" + Str$(ZZImpoDolar) + "',"
            ZSql = ZSql & "OrdenIII = " + "'" + Str$(ZZOrdenDol) + "',"
            ZSql = ZSql & "PtaOrdenIII = " + "'" + Str$(ZZPtaOrdenDol) + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + ZZArticulo + "'"
                    
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            If Aaa = 999 Then
                If ZZCarpeta = "" Then
                    If ZZCosto1 = ZZImpoDolar Then
                        ZSql = ""
                        ZSql = ZSql & "UPDATE Articulo SET "
                        ZSql = ZSql & "Costo1 = " + "'" + "0" + "'"
                        ZSql = ZSql & " Where Codigo = " + "'" + ZZArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                End If
                
                If ZZCarpeta <> "" Then
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Articulo SET "
                    ZSql = ZSql & "Costo1 = " + "'" + Str$(ZZImpoImpo) + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + ZZArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
            
        Next Cicla
        
        Call Conecta_Empresa
    
    Next ZZCiclo

Stop


End Sub




Private Sub OrdenI_dblclick()

    Erase ZZZZImpre
    
    ZZZZImpre(1) = "Orden " + "SI"
    ZZZZImpre(2) = "Orden " + "PI"
    ZZZZImpre(3) = "Orden " + "SII"
    ZZZZImpre(4) = "Orden " + "PII"
    ZZZZImpre(5) = "Orden " + "SIII"
    ZZZZImpre(6) = "Orden " + "SIV"
    ZZZZImpre(7) = "Orden " + "SV"
    ZZZZImpre(8) = "Orden " + "PIII"
    ZZZZImpre(9) = "Orden " + "PV"
    ZZZZImpre(10) = "Orden " + "SVI"
    ZZZZImpre(11) = "Orden " + "SVII"
    
    If Val(OrdenI.Text) <> 0 And Trim(TituloOrdenI.Caption) <> "" Then
        ZZZZOrden = OrdenI.Text
        For ZZZZCiclo = 1 To 11
            If TituloOrdenI.Caption = ZZZZImpre(ZZZZCiclo) Then
                ZZZZPtaOrden = ZZZZCiclo
            End If
        Next ZZZZCiclo
        Call Muestra_Historial
    End If
    
End Sub

Private Sub OrdenII_dblclick()

    Erase ZZZZImpre
    
    ZZZZImpre(1) = "Orden " + "SI"
    ZZZZImpre(2) = "Orden " + "PI"
    ZZZZImpre(3) = "Orden " + "SII"
    ZZZZImpre(4) = "Orden " + "PII"
    ZZZZImpre(5) = "Orden " + "SIII"
    ZZZZImpre(6) = "Orden " + "SIV"
    ZZZZImpre(7) = "Orden " + "SV"
    ZZZZImpre(8) = "Orden " + "PIII"
    ZZZZImpre(9) = "Orden " + "PV"
    ZZZZImpre(10) = "Orden " + "SVI"
    ZZZZImpre(11) = "Orden " + "SVII"
    
    If Val(OrdenII.Text) <> 0 And Trim(TituloOrdenII.Caption) <> "" Then
        ZZZZOrden = OrdenII.Text
        For ZZZZCiclo = 1 To 11
            If TituloOrdenII.Caption = ZZZZImpre(ZZZZCiclo) Then
                ZZZZPtaOrden = ZZZZCiclo
            End If
        Next ZZZZCiclo
        Call Muestra_Historial
    End If
    
End Sub


Private Sub OrdenIII_dblclick()

    Erase ZZZZImpre
    
    ZZZZImpre(1) = "Orden " + "SI"
    ZZZZImpre(2) = "Orden " + "PI"
    ZZZZImpre(3) = "Orden " + "SII"
    ZZZZImpre(4) = "Orden " + "PII"
    ZZZZImpre(5) = "Orden " + "SIII"
    ZZZZImpre(6) = "Orden " + "SIV"
    ZZZZImpre(7) = "Orden " + "SV"
    ZZZZImpre(8) = "Orden " + "PIII"
    ZZZZImpre(9) = "Orden " + "PV"
    ZZZZImpre(10) = "Orden " + "SVI"
    ZZZZImpre(11) = "Orden " + "SVII"
    
    If Val(OrdenIII.Text) <> 0 And Trim(TituloOrdenIII.Caption) <> "" Then
        ZZZZOrden = OrdenIII.Text
        For ZZZZCiclo = 1 To 11
            If TituloOrdenIII.Caption = ZZZZImpre(ZZZZCiclo) Then
                ZZZZPtaOrden = ZZZZCiclo
            End If
        Next ZZZZCiclo
        Call Muestra_Historial
    End If
    
End Sub



Private Sub Muestra_Historial()

    ZZZZEmpresa = Wempresa

    Select Case ZZZZPtaOrden
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

    ZZZZLaudo = ""

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " and Laudo.orden = " + "'" + ZZZZOrden + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        ZZZZLaudo = Str$(rstLaudo!Laudo)
        rstLaudo.Close
    End If
    
    If Val(ZZZZLaudo) <> 0 Then
        
        HistoriaOrden.Text = ""
        HistoriaInforme.Text = ""
        HistoriaRemito.Text = ""
        HistoriaFactura.Text = ""
        HistoriaCarpeta.Text = ""
        
        HistoriaFechaOrden.Text = "  /  /    "
        HistoriaFechaInforme.Text = "  /  /    "
        HistoriaFechaFactura.Text = "  /  /    "
        
        WOrden = ""
        WInforme = ""
        WRemito = ""
        WFactura = ""
        WProveedor = ""
        WCarpeta = ""
        
        WFechaOrden = "  /  /    "
        WFechaInforme = "  /  /    "
        WFechaFactura = "  /  /    "
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + ZZZZLaudo + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + Codigo.Text + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            ZZZZInforme = Str$(rstLaudo!Informe)
            rstLaudo.Close
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Articulo = " + "'" + Codigo.Text + "'"
            ZSql = ZSql + " and Informe.Informe = " + "'" + ZZZZInforme + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                WRemito = Str$(rstInforme!Remito)
                WFechaInforme = rstInforme!Fecha
                rstInforme.Close
            End If
            
            WProveedor = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Articulo = " + "'" + Codigo.Text + "'"
            ZSql = ZSql + " and Orden.Orden = " + "'" + ZZZZOrden + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WFechaOrden = rstOrden!Fecha
                WProveedor = rstOrden!Proveedor
                WCarpeta = rstOrden!Carpeta
                rstOrden.Close
            End If
            
            If WProveedor <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    DesproveedorH.Caption = rstProveedor!Nombre
                    rstProveedor.Close
                End If
            End If
            
            XEmpresa = Wempresa
            
            If ZZZZPtaOrden = 1 Or ZZZZPtaOrden = 3 Or ZZZZPtaOrden = 5 Or ZZZZPtaOrden = 6 Or ZZZZPtaOrden = 7 Or ZZZZPtaOrden = 10 Or ZZZZPtaOrden = 11 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM IvaComp"
            ZSql = ZSql + " Where IvaComp.Remito LIKE " + "'" + "%" + Trim(WRemito) + "%" + "'"
            ZSql = ZSql + " and IvaComp.Proveedor = " + "'" + WProveedor + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                WFactura = Str$(Val(rstIvaComp!Numero))
                WFechaFactura = rstIvaComp!Fecha
                rstIvaComp.Close
            End If
            
            Call Conecta_Empresa
        
        End If
        
        HistoriaOrden.Text = ZZZZOrden
        HistoriaInforme.Text = ZZZZInforme
        HistoriaRemito.Text = WRemito
        HistoriaFactura.Text = WFactura
        HistoriaCarpeta.Text = WCarpeta
        
        HistoriaFechaOrden.Text = WFechaOrden
        HistoriaFechaInforme.Text = WFechaInforme
        HistoriaFechaFactura.Text = WFechaFactura
        
        PantaHistoria.Height = 3300
        PantaHistoria.Left = 3000
        PantaHistoria.Top = 1080
        PantaHistoria.Width = 5295
        
        PantaHistoria.Visible = True
        
    End If
    
    XEmpresa = ZZZZEmpresa
    Call Conecta_Empresa
    
    
End Sub

Private Sub PantaHistoriaCierra_Click()
    PantaHistoria.Visible = False
End Sub






Private Sub Command124_Click()

    
    XEmpresa = Wempresa
    Erase CargaEmpresa

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


    Rem
    Rem proceso los comodatos
    Rem

    Set appExcel = CreateObject("Excel.application")
    
    Rem modificar aca
    Rem Ruta = Nombre del archivo excel
    Rem
    
    LugarPlanilla = 0
    ruta = "C:\david\eliminar.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
            ZZCodigo = appExcel.cells(LugarPlanilla, 1).Value
                    
            If Trim(ZZCodigo) = "" Then Exit Do
            
            ZZCodigo = Mid$(ZZCodigo, 1, 2) + "-" + Mid$(ZZCodigo, 4, 3) + "-" + Mid$(ZZCodigo, 8, 3)
            
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Articulo SET "
                    ZSql = ZSql & "Minimo = " + "'" + "0" + "',"
                    ZSql = ZSql & "Minimo1 = " + "'" + "0" + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + ZZCodigo + "'"
                            
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
                
            Next Cicla
            
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
    Stop

End Sub

