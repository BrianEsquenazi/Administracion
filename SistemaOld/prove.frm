VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProve 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Proveedores"
   ClientHeight    =   8595
   ClientLeft      =   450
   ClientTop       =   390
   ClientWidth     =   10935
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   10935
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   9840
      TabIndex        =   116
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox IbCiudadII 
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
      Left            =   8640
      TabIndex        =   114
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Frame PantaObserva 
      Height          =   4455
      Left            =   3120
      TabIndex        =   90
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton CierraPantaObserva 
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
         Height          =   615
         Left            =   2280
         TabIndex        =   93
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox ObservacionesII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   91
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones del Proveedor"
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
         Index           =   5
         Left            =   240
         TabIndex        =   92
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.CommandButton AltaCufe 
      Caption         =   "CUFE"
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
      Left            =   8400
      TabIndex        =   112
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Frame PantaCufe 
      Height          =   3135
      Left            =   1920
      TabIndex        =   99
      Top             =   840
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox DirCufeIII 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   109
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox CufeIII 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   107
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox DirCufeII 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   106
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox CufeII 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   104
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox DirCufe 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   103
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Cufe 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   101
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CierraPantaCufe 
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
         Height          =   615
         Left            =   2880
         TabIndex        =   100
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label35 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   3480
         TabIndex        =   111
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label34 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   1800
         TabIndex        =   110
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "CUFE III"
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "CUFE II"
         Height          =   255
         Left            =   240
         TabIndex        =   105
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "CUFE"
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox PorceIbCaba 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   97
      Text            =   " "
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Email Embargo"
      Height          =   375
      Left            =   8400
      TabIndex        =   95
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Observa 
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
      Height          =   375
      Left            =   8400
      TabIndex        =   89
      Top             =   6840
      Width           =   2295
   End
   Begin VB.ComboBox Califica 
      Height          =   315
      Left            =   7920
      TabIndex        =   88
      Text            =   " "
      Top             =   6480
      Width           =   1335
   End
   Begin VB.ComboBox Estado 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   7920
      TabIndex        =   84
      Text            =   " "
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox PorceIb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9000
      MaxLength       =   10
      TabIndex        =   83
      Text            =   " "
      Top             =   3960
      Width           =   855
   End
   Begin VB.ComboBox Region 
      Height          =   315
      Left            =   8640
      TabIndex        =   80
      Text            =   " "
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox Iso 
      Height          =   315
      Left            =   7920
      TabIndex        =   79
      Text            =   " "
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox CategoriaII 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   76
      Text            =   " "
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ComboBox CategoriaI 
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   " "
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ComboBox TipoProv 
      Height          =   315
      ItemData        =   "prove.frx":0000
      Left            =   2280
      List            =   "prove.frx":0002
      TabIndex        =   72
      Text            =   " "
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Cai 
      Height          =   285
      Left            =   8640
      MaxLength       =   14
      TabIndex        =   68
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede actualizar los datos de los proveedores. El sistema se encuentra sin conexion con las demas plantas"
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
      Left            =   3360
      Picture         =   "prove.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3720
      TabIndex        =   63
      Top             =   1080
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   65
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   64
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   66
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   600
      TabIndex        =   30
      Top             =   6000
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox TipoListado 
         Height          =   315
         Left            =   1320
         TabIndex        =   62
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   38
         Text            =   " "
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   37
         Text            =   " "
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Tipo Listado"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox NroInsc 
      Height          =   285
      Left            =   7680
      MaxLength       =   15
      TabIndex        =   60
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox NroIb 
      Height          =   285
      Left            =   5880
      MaxLength       =   20
      TabIndex        =   58
      Top             =   3960
      Width           =   1815
   End
   Begin VB.ComboBox CodIb 
      Height          =   315
      ItemData        =   "prove.frx":0746
      Left            =   1200
      List            =   "prove.frx":0748
      TabIndex        =   55
      Text            =   " "
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   54
      Top             =   5520
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.TextBox NombreCheque 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   14
      Text            =   " "
      Top             =   3600
      Width           =   7095
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   1080
      TabIndex        =   52
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   13
      Text            =   " "
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox Iva 
      Height          =   315
      Left            =   6600
      TabIndex        =   12
      Text            =   " "
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Text            =   " "
      Top             =   2880
      Width           =   2415
   End
   Begin VB.ComboBox Provincia 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Text            =   " "
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Dias 
      Height          =   285
      Left            =   6600
      MaxLength       =   20
      TabIndex        =   7
      Text            =   " "
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox EMail 
      Height          =   285
      Left            =   2280
      MaxLength       =   200
      TabIndex        =   8
      Text            =   " "
      Top             =   2160
      Width           =   8535
   End
   Begin VB.TextBox Telefono 
      Height          =   285
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   6
      Text            =   " "
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Cuit 
      Height          =   285
      Left            =   8640
      MaxLength       =   15
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Postal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   5
      Text            =   " "
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Localidad 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   1080
      Width           =   5175
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox Proveedor 
      Height          =   285
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wprove.rpt"
      Destination     =   1
      WindowTitle     =   "Listados de Proveedores"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "prove.frx":074A
      Left            =   120
      List            =   "prove.frx":0751
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   4440
      TabIndex        =   27
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   3360
      TabIndex        =   26
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   6720
      TabIndex        =   21
      Top             =   6840
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   2280
      TabIndex        =   16
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5520
      TabIndex        =   20
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1200
      TabIndex        =   19
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Nombre 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin MSMask.MaskEdBox VtoCai 
      Height          =   285
      Left            =   8640
      TabIndex        =   71
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox VtoIso 
      Height          =   285
      Left            =   9360
      TabIndex        =   77
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FechaCalifica 
      Height          =   285
      Left            =   9360
      TabIndex        =   86
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FechaCategoria 
      Height          =   285
      Left            =   5880
      TabIndex        =   94
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FechaNroInsc 
      Height          =   285
      Left            =   9480
      TabIndex        =   96
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.ComboBox CodIbCaba 
      Height          =   315
      ItemData        =   "prove.frx":075F
      Left            =   3120
      List            =   "prove.frx":0761
      TabIndex        =   113
      Text            =   " "
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label36 
      Caption         =   "Tipo de Inscrip IB"
      Height          =   375
      Left            =   7320
      TabIndex        =   115
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label30 
      Caption         =   "Porcel IB  CABA"
      Height          =   255
      Left            =   9600
      TabIndex        =   98
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Calificacion"
      Height          =   375
      Left            =   6840
      TabIndex        =   87
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "Estado"
      Height          =   375
      Left            =   6840
      TabIndex        =   85
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "Porcel Ib Pcia"
      Height          =   255
      Left            =   7800
      TabIndex        =   82
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label26 
      Caption         =   "Region"
      Height          =   255
      Left            =   7560
      TabIndex        =   81
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Certificados"
      Height          =   375
      Left            =   6840
      TabIndex        =   78
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "Categoria del Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label23 
      Caption         =   "Tipo de Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "Vto. CAI"
      Height          =   375
      Left            =   6840
      TabIndex        =   70
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "CAI"
      Height          =   255
      Left            =   7320
      TabIndex        =   69
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Nro. Insc. SEDRONAR"
      Height          =   255
      Left            =   5880
      TabIndex        =   59
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Nro. I.b."
      Height          =   255
      Left            =   5280
      TabIndex        =   57
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   "Condicion I.b."
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Cheque a nombre de "
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label DesCuenta 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   4920
      TabIndex        =   51
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label14 
      Caption         =   "Cuenta Contable"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Codigo de Iva"
      Height          =   255
      Left            =   5160
      TabIndex        =   49
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Provincia"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "Dias de Plazo"
      Height          =   255
      Left            =   5160
      TabIndex        =   47
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Tipo de Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "E Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Cuit"
      Height          =   255
      Left            =   7680
      TabIndex        =   42
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo Postal"
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Localidad"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Razon Social"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Proveedor"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "PrgProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTipo As String
Private WProvincia As String
Private WIva As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstTipoProv As Recordset
Dim spTipoProv As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim cParam As String
Dim x As Integer
Dim EmpresaReal As String
Dim CargaEmpresa(12, 2) As String
Dim WCuit As String
Dim ZVectorI(100) As Integer
Dim ZVectorII(100) As Integer

Private WGraba As String
Private WProceso As String

Dim ZZPorceIbCaba As Double

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


Sub Verifica_datos()
    If Val(Cuenta.Text) = 0 Then
        Cuenta.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
    With RstProveedor
        If RstProveedor.RecordCount > 0 Then
            Nombre.BackColor = &HFFFFFF
            Proveedor.Text = IIf(IsNull(!Proveedor), "", !Proveedor)
            Nombre.Text = IIf(IsNull(!Nombre), "", !Nombre)
            Direccion.Text = IIf(IsNull(!Direccion), "", !Direccion)
            Localidad.Text = IIf(IsNull(!Localidad), "", !Localidad)
            Postal.Text = IIf(IsNull(!Postal), "", !Postal)
            Cuit.Text = IIf(IsNull(!Cuit), "", !Cuit)
            Telefono.Text = IIf(IsNull(!Telefono), "", !Telefono)
            EMail.Text = IIf(IsNull(!EMail), "", !EMail)
            EMail.Text = LCase(EMail.Text)
            Observaciones.Text = IIf(IsNull(!Observaciones), "", !Observaciones)
            Dias.Text = IIf(IsNull(!Dias), "", !Dias)
            Cuenta.Text = IIf(IsNull(!Cuenta), "", !Cuenta)
            Iva.ListIndex = IIf(IsNull(!Iva), "", !Iva)
            Tipo.ListIndex = IIf(IsNull(!Tipo), "", !Tipo)
            Provincia.ListIndex = 25
            Provincia.ListIndex = IIf(IsNull(!Provincia), "", !Provincia)
            Region.ListIndex = 0
            PorceIb.Text = Str$(!PorceIb)
            ZZPorceIbCaba = IIf(IsNull(!PorceIbCaba), "0", !PorceIbCaba)
            PorceIbCaba.Text = Str$(ZZPorceIbCaba)
            Region.ListIndex = IIf(IsNull(!Region), "", !Region)
            NombreCheque.Text = IIf(IsNull(!NombreCheque), "", !NombreCheque)
            CodIb.ListIndex = IIf(IsNull(!CodIb), "0", !CodIb)
            CodIbCaba.ListIndex = IIf(IsNull(!CodIbCaba), "0", !CodIbCaba)
            TipoProv.ListIndex = IIf(IsNull(!TipoProv), "0", !TipoProv)
            NroIb.Text = IIf(IsNull(!NroIb), "", !NroIb)
            NroInsc.Text = IIf(IsNull(!NroInsc), "", !NroInsc)
            FechaNroInsc.Text = IIf(IsNull(!FechaNroInsc), "  /  /    ", !FechaNroInsc)
            Cufe.Text = IIf(IsNull(!Cufe), "", !Cufe)
            CufeII.Text = IIf(IsNull(!CufeII), "", !CufeII)
            CufeIII.Text = IIf(IsNull(!CufeIII), "", !CufeIII)
            DirCufe.Text = IIf(IsNull(!DirCufe), "", !DirCufe)
            DirCufeII.Text = IIf(IsNull(!DirCufeII), "", !DirCufeII)
            DirCufeIII.Text = IIf(IsNull(!DirCufeIII), "", !DirCufeIII)
            Cai.Text = IIf(IsNull(!Cai), "", !Cai)
            VtoCai.Text = IIf(IsNull(!VtoCai), "  /  /    ", !VtoCai)
            CategoriaI.ListIndex = IIf(IsNull(!CategoriaI), "0", !CategoriaI)
            CategoriaII.ListIndex = IIf(IsNull(!CategoriaII), "0", !CategoriaII)
            Iso.ListIndex = IIf(IsNull(!Iso), "0", !Iso)
            VtoIso.Text = IIf(IsNull(!VtoIso), "  /  /    ", !VtoIso)
            Estado.ListIndex = IIf(IsNull(!Estado), "0", !Estado)
            Califica.ListIndex = IIf(IsNull(!Califica), "0", !Califica)
            IbCiudadII.ListIndex = IIf(IsNull(!IbCiudadII), "0", !IbCiudadII)
            FechaCalifica.Text = IIf(IsNull(!FechaCalifica), "  /  /    ", !FechaCalifica)
            ObservacionesII.Text = IIf(IsNull(RstProveedor!ObservacionesII), "", RstProveedor!ObservacionesII)
            FechaCategoria.Text = IIf(IsNull(!FechaCategoria), "  /  /    ", !FechaCategoria)
            ZEmbargo = IIf(IsNull(!Embargo), "", !Embargo)
            If ZEmbargo = "S" Then
                Nombre.BackColor = &HFF&
            End If
            If Estado.ListIndex = 2 Then
                Estado.BackColor = &HFF&
                    Else
                Estado.BackColor = &H80000005
            End If
            Call Format_datos
        End If
    End With
    RstProveedor.Close
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
       DesCuenta.Caption = rstCuenta!Descripcion
       rstCuenta.Close
       Nombre.SetFocus
            Else
       Cuenta.SetFocus
    End If
    
    Exit Sub

WError:
    Resume Next
    
End Sub

Sub Imprime_Descripcion()

    spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        DesCuenta.Caption = rstCuenta!Descripcion
        rstCuenta.Close
        Nombre.SetFocus
            Else
        Cuenta.SetFocus
    End If
    
End Sub

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Proveedor.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Proveedor.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Proveedor.Proveedor , Proveedor.Nombre, Proveedor.Direccion, Proveedor.Localidad, Proveedor.Postal, Proveedor.Telefono, Proveedor.Observaciones, Proveedor.NombreCheque " _
                        + "From " + DSQ + ".dbo.Proveedor Proveedor " _
                        + "Where Proveedor.Proveedor >= '0' AND Proveedor.Proveedor <= '99999999999'"
    Listado.DataFiles(1) = Wempresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    If TipoListado.ListIndex = 0 Then
        Listado.ReportFileName = "WProve.rpt"
            Else
        Listado.ReportFileName = "WProveres.rpt"
    End If
    
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub AltaCufe_Click()
    PantaCufe.Visible = True
    Cufe.SetFocus
End Sub

Private Sub AvisoError_Click()
    AvisoError.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub CierraPantaCufe_Click()
    PantaCufe.Visible = False
End Sub

Private Sub CierraPantaObserva_Click()
    PantaObserva.Visible = False
End Sub

Private Sub cmdAdd_Click()

    If WGraba <> "S" Then
    
        WProceso = "A"
        Call Ingresa_clave

               Else

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

        On Error GoTo ErrAltaProveedor

        If Proveedor.Text <> "" Then
    
            Call Verifica_datos
            WPasa = "S"
    
            If Val(Cuenta.Text) <> 0 Then
        
                spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount <= 0 Then
                    WPasa = "N"
                    m$ = "Codigo de Cuenta Contable incorrecto"
                    A% = MsgBox(m$, 0, "Archivo de Proveedores")
                        Else
                    rstCuenta.Close
                End If
            
            End If
            
            If TipoProv.ListIndex <= 0 Then
                m$ = "Codigo de Tipo de Proveedor no Informado"
                A% = MsgBox(m$, 0, "Archivo de Proveedores")
                Exit Sub
            End If
            
            If Trim(NroInsc.Text) <> "" Then
                If FechaNroInsc.Text = "  /  /    " Then
                    m$ = "Se debe informar la fecha de vencimiento del sedronar"
                    A% = MsgBox(m$, 0, "Archivo de Proveedores")
                    Exit Sub
                        Else
                    Call Valida_fecha1(FechaNroInsc.Text, Auxi)
                    If Auxi <> "S" Then
                        m$ = "Fecha de vencimiento del sedronar incorrectra"
                        A% = MsgBox(m$, 0, "Archivo de Proveedores")
                        Exit Sub
                    End If
                End If
            End If
            
            
            WCuit = Cuit.Text
            If Trim(WCuit) <> "" Then
            
                Call Eval
                
                If Len(WCuit) = 11 Then
                
                    Erase ZVectorI
                    ZVectorII(1) = 5
                    ZVectorII(2) = 4
                    ZVectorII(3) = 3
                    ZVectorII(4) = 2
                    ZVectorII(5) = 7
                    ZVectorII(6) = 6
                    ZVectorII(7) = 5
                    ZVectorII(8) = 4
                    ZVectorII(9) = 3
                    ZVectorII(10) = 2
                    ZSuma = 0
                    ZDigitoI = Val(Mid$(WCuit, 11, 1))
                    
                    For Ciclo = 1 To 10
                        ZSuma = ZSuma + (Val(ZVectorII(Ciclo)) * Val(Mid$(WCuit, Ciclo, 1)))
                    Next Ciclo
                    MiResultado = ZSuma Mod 11
                    MiResultadoII = 11 - MiResultado
                    If MiResultadoII = 11 Then
                        ZDigitoII = 0
                            Else
                        If MiResultadoII = 10 Then
                            ZDigitoII = 9
                                Else
                            ZDigitoII = MiResultadoII
                        End If
                    End If
                    
                    If ZDigitoI <> ZDigitoII Then
                        m$ = "Error en Cuit " + Nombre.Text + "  " + WCuit + "  Ok= " + Str$(ZDigitoII)
                        A% = MsgBox(m$, 0, "Ingreso de Proveedores")
                    End If
                
                        Else
                        
                    m$ = "Error en Cuit " + Nombre.Text + "  " + WCuit
                    A% = MsgBox(m$, 0, "Ingreso de Proveedores")
                    
                End If
                
                    Else
                    
                m$ = "No se informo numero de Cuit " + Nombre.Text
                A% = MsgBox(m$, 0, "Ingreso de Proveedores")
                
            End If
            
            
            
                
        
            WIva = "7"
            WTipo = "4"
            WProvincia = "25"
            WRegion = "0"
             
            WCodIb = Str$(CodIb.ListIndex)
            WCodIbCaba = Str$(CodIbCaba.ListIndex)
            WTipoProv = Str$(TipoProv.ListIndex)
            WTipo = Str$(Tipo.ListIndex)
            WProvincia = Str$(Provincia.ListIndex)
            WRegion = Str$(Region.ListIndex)
            WIva = Str$(Iva.ListIndex)
            Call Ceros(WTipo, 1)
            Call Ceros(WProvincia, 2)
            Call Ceros(WIva, 1)
            XEmpresa = "1"
            WImporte1 = ""
            WImporte2 = ""
            WImporte3 = ""
            WImporte4 = ""
            WImporte5 = ""
            WImporte6 = ""
            WDate = Date$
            WEstado = Str$(Estado.ListIndex)
            WCalifica = Str$(Califica.ListIndex)
            WIbCiudadII = Str$(IbCiudadII.ListIndex)
            WFechaCalifica = FechaCalifica.Text
            WOrdFechaCalifica = Right$(FechaCalifica.Text, 4) + Mid$(FechaCalifica.Text, 4, 2) + Left$(FechaCalifica.Text, 2)
        
            EmpresaReal = Wempresa
            
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    For CiclaEmpresa = 1 To 7
                    
                        Select Case CiclaEmpresa
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
                        
                        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
                        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                        If RstProveedor.RecordCount > 0 Then
                            RstProveedor.Close
                            Call Verifica_datos
                            cParam = "'" + Proveedor.Text + "','" + Nombre.Text + "','" + Direccion.Text + "','" _
                                        + Localidad.Text + "','" + WProvincia + "','" + Postal.Text + "','" + Cuit.Text + "','" _
                                        + Telefono.Text + "','" + EMail.Text + "','" + Observaciones.Text + "','" _
                                        + WTipo + "','" + WIva + "','" _
                                        + Dias.Text + "','" + XEmpresa + "','" + Cuenta.Text + "','" _
                                        + WImporte1 + "','" + WImporte2 + "','" _
                                        + WImporte2 + "','" + WImporte4 + "','" _
                                        + WImporte3 + "','" + WImporte6 + "','" _
                                        + NombreCheque.Text + "','" _
                                        + WDate + "','" _
                                        + WCodIb + "','" _
                                        + NroIb.Text + "','" _
                                        + NroInsc.Text + "'"
                            Set RstProveedor = db.OpenRecordset("ModificaProveedor1 " + cParam, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Proveedor SET "
                            ZSql = ZSql + " FechaNroInsc = " + "'" + FechaNroInsc.Text + "',"
                            ZSql = ZSql + " OrdFechaNroInsc = " + "'" + WOrdFechaNroInsc + "',"
                            ZSql = ZSql + " Region = " + "'" + WRegion + "',"
                            ZSql = ZSql + " EMail = " + "'" + EMail.Text + "',"
                            ZSql = ZSql + " Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + " CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + " CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + " DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + " DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + " DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                            ZSql = ZSql + " Cai = " + "'" + Cai.Text + "',"
                            ZSql = ZSql + " VtoCai = " + "'" + VtoCai.Text + "',"
                            ZSql = ZSql + " CodIbCaba = " + "'" + Str$(CodIbCaba.ListIndex) + "',"
                            ZSql = ZSql + " TipoProv = " + "'" + Str$(TipoProv.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaI = " + "'" + Str$(CategoriaI.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaII = " + "'" + Str$(CategoriaII.ListIndex) + "',"
                            ZSql = ZSql + " PorceIb = " + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + " PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + " Iso = " + "'" + Str$(Iso.ListIndex) + "',"
                            ZSql = ZSql + " VtoIso = " + "'" + VtoIso.Text + "',"
                            ZSql = ZSql + " Estado = " + "'" + WEstado + "',"
                            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
                            ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
                            ZSql = ZSql + " IbCiudadII = " + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
                            ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
                            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                            spProveedor = ZSql
                            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                    
                            Call Verifica_datos
                            cParam = "'" + Proveedor.Text + "','" + Nombre.Text + "','" + Direccion.Text + "','" _
                                        + Localidad.Text + "','" + WProvincia + "','" + Postal.Text + "','" + Cuit.Text + "','" _
                                        + Telefono.Text + "','" + EMail.Text + "','" + Observaciones.Text + "','" _
                                        + WTipo + "','" + WIva + "','" _
                                        + Dias.Text + "','" + XEmpresa + "','" + Cuenta.Text + "','" _
                                        + WImporte1 + "','" + WImporte2 + "','" _
                                        + WImporte2 + "','" + WImporte4 + "','" _
                                        + WImporte3 + "','" + WImporte6 + "','" _
                                        + NombreCheque.Text + "','" _
                                        + WDate + "','" _
                                        + WCodIb + "','" _
                                        + NroIb.Text + "','" _
                                        + NroInsc.Text + "'"
                            Set RstProveedor = db.OpenRecordset("AltaProveedor1 " + cParam, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Proveedor SET "
                            ZSql = ZSql + " FechaNroInsc = " + "'" + FechaNroInsc.Text + "',"
                            ZSql = ZSql + " OrdFechaNroInsc = " + "'" + WOrdFechaNroInsc + "',"
                            ZSql = ZSql + " Region = " + "'" + WRegion + "',"
                            ZSql = ZSql + " EMail = " + "'" + EMail.Text + "',"
                            ZSql = ZSql + " Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + " CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + " CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + " DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + " DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + " DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                            ZSql = ZSql + " Cai = " + "'" + Cai.Text + "',"
                            ZSql = ZSql + " VtoCai = " + "'" + VtoCai.Text + "',"
                            ZSql = ZSql + " CodIbCaba = " + "'" + Str$(CodIbCaba.ListIndex) + "',"
                            ZSql = ZSql + " TipoProv = " + "'" + Str$(TipoProv.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaI = " + "'" + Str$(CategoriaI.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaII = " + "'" + Str$(CategoriaII.ListIndex) + "',"
                            ZSql = ZSql + " PorceIb = " + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + " PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + " Iso = " + "'" + Str$(Iso.ListIndex) + "',"
                            ZSql = ZSql + " VtoIso = " + "'" + VtoIso.Text + "',"
                            ZSql = ZSql + " Estado = " + "'" + WEstado + "',"
                            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
                            ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
                            ZSql = ZSql + " IbCiudadII = " + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
                            ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
                            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                            spProveedor = ZSql
                            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                    
                    Next CiclaEmpresa
                            
                
                Case Else
                    For CiclaEmpresa = 1 To 4
                    
                        Select Case CiclaEmpresa
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
                        
                        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
                        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                        If RstProveedor.RecordCount > 0 Then
                            RstProveedor.Close
                            Call Verifica_datos
                            cParam = "'" + Proveedor.Text + "','" + Nombre.Text + "','" + Direccion.Text + "','" _
                                        + Localidad.Text + "','" + WProvincia + "','" + Postal.Text + "','" + Cuit.Text + "','" _
                                        + Telefono.Text + "','" + EMail.Text + "','" + Observaciones.Text + "','" _
                                        + WTipo + "','" + WIva + "','" _
                                        + Dias.Text + "','" + XEmpresa + "','" + Cuenta.Text + "','" _
                                        + WImporte1 + "','" + WImporte2 + "','" _
                                        + WImporte2 + "','" + WImporte4 + "','" _
                                        + WImporte3 + "','" + WImporte6 + "','" _
                                        + NombreCheque.Text + "','" _
                                        + WDate + "','" _
                                        + WCodIb + "','" _
                                        + NroIb.Text + "','" _
                                        + NroInsc.Text + "'"
                            Set RstProveedor = db.OpenRecordset("ModificaProveedor1 " + cParam, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Proveedor SET "
                            ZSql = ZSql + " FechaNroInsc = " + "'" + FechaNroInsc.Text + "',"
                            ZSql = ZSql + " OrdFechaNroInsc = " + "'" + WOrdFechaNroInsc + "',"
                            ZSql = ZSql + " Region = " + "'" + WRegion + "',"
                            ZSql = ZSql + " EMail = " + "'" + EMail.Text + "',"
                            ZSql = ZSql + " Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + " CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + " CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + " DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + " DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + " DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                            ZSql = ZSql + " Cai = " + "'" + Cai.Text + "',"
                            ZSql = ZSql + " VtoCai = " + "'" + VtoCai.Text + "',"
                            ZSql = ZSql + " CodIbCaba = " + "'" + Str$(CodIbCaba.ListIndex) + "',"
                            ZSql = ZSql + " TipoProv = " + "'" + Str$(TipoProv.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaI = " + "'" + Str$(CategoriaI.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaII = " + "'" + Str$(CategoriaII.ListIndex) + "',"
                            ZSql = ZSql + " PorceIb = " + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + " PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + " Iso = " + "'" + Str$(Iso.ListIndex) + "',"
                            ZSql = ZSql + " VtoIso = " + "'" + VtoIso.Text + "',"
                            ZSql = ZSql + " Estado = " + "'" + WEstado + "',"
                            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
                            ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
                            ZSql = ZSql + " IbCiudadII = " + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
                            ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
                            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                            spProveedor = ZSql
                            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                    
                            Call Verifica_datos
                            cParam = "'" + Proveedor.Text + "','" + Nombre.Text + "','" + Direccion.Text + "','" _
                                        + Localidad.Text + "','" + WProvincia + "','" + Postal.Text + "','" + Cuit.Text + "','" _
                                        + Telefono.Text + "','" + EMail.Text + "','" + Observaciones.Text + "','" _
                                        + WTipo + "','" + WIva + "','" _
                                        + Dias.Text + "','" + XEmpresa + "','" + Cuenta.Text + "','" _
                                        + WImporte1 + "','" + WImporte2 + "','" _
                                        + WImporte2 + "','" + WImporte4 + "','" _
                                        + WImporte3 + "','" + WImporte6 + "','" _
                                        + NombreCheque.Text + "','" _
                                        + WDate + "','" _
                                        + WCodIb + "','" _
                                        + NroIb.Text + "','" _
                                        + NroInsc.Text + "'"
                            Set RstProveedor = db.OpenRecordset("AltaProveedor1 " + cParam, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Proveedor SET "
                            ZSql = ZSql + " FechaNroInsc = " + "'" + FechaNroInsc.Text + "',"
                            ZSql = ZSql + " OrdFechaNroInsc = " + "'" + WOrdFechaNroInsc + "',"
                            ZSql = ZSql + " Region = " + "'" + WRegion + "',"
                            ZSql = ZSql + " EMail = " + "'" + EMail.Text + "',"
                            ZSql = ZSql + " Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + " CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + " CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + " DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + " DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + " DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                            ZSql = ZSql + " Cai = " + "'" + Cai.Text + "',"
                            ZSql = ZSql + " VtoCai = " + "'" + VtoCai.Text + "',"
                            ZSql = ZSql + " CodIbCaba = " + "'" + Str$(CodIbCaba.ListIndex) + "',"
                            ZSql = ZSql + " TipoProv = " + "'" + Str$(TipoProv.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaI = " + "'" + Str$(CategoriaI.ListIndex) + "',"
                            ZSql = ZSql + " CategoriaII = " + "'" + Str$(CategoriaII.ListIndex) + "',"
                            ZSql = ZSql + " PorceIb = " + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + " PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + " Iso = " + "'" + Str$(Iso.ListIndex) + "',"
                            ZSql = ZSql + " VtoIso = " + "'" + VtoIso.Text + "',"
                            ZSql = ZSql + " Estado = " + "'" + WEstado + "',"
                            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
                            ZSql = ZSql + " Califica = " + "'" + WCalifica + "',"
                            ZSql = ZSql + " IbCiudadII = " + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + " FechaCalifica = " + "'" + WFechaCalifica + "',"
                            ZSql = ZSql + " OrdFechaCalifica = " + "'" + WOrdFechaCalifica + "'"
                            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                            spProveedor = ZSql
                            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                    
                    Next CiclaEmpresa
                    
            End Select
        
        
            Wempresa = EmpresaReal
            txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Call CmdLimpiar_Click
        
        End If
        
        Proveedor.SetFocus
        
    End If
    
    Exit Sub

ErrAltaProveedor:
    MsgBox Err.Description
    Resume Next
    
Control_Error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
    
End Sub

Private Sub cmdDelete_Click()

    If WGraba <> "S" Then
    
        WProceso = "B"
        Call Ingresa_clave
        
               Else
               
        If Proveedor.Text <> "" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                EmpresaReal = Wempresa
                
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                Wempresa = EmpresaReal
                txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
                spProveedor = "BorrarProveedor " + "'" + Proveedor.Text + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
        
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            End If
            
        End If
    End If
End Sub

Private Sub CmdLimpiar_Click()
    Proveedor.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Cuit.Text = ""
    Telefono.Text = ""
    EMail.Text = ""
    Observaciones.Text = ""
    Dias.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    NombreCheque = ""
    ObservacionesII.Text = ""
    Nombre.BackColor = &HFFFFFF
    
    Estado.BackColor = &H80000005
    
    Iva.ListIndex = 7
    Tipo.ListIndex = 8
    Provincia.ListIndex = 25
    Region.ListIndex = 0
    CodIb.ListIndex = 0
    IbCiudadII.ListIndex = 0
    CodIbCaba.ListIndex = 2
    TipoProv.ListIndex = 0
    CategoriaI.ListIndex = 0
    CategoriaII.ListIndex = 0
    Iso.ListIndex = 0
    WGraba = ""
    WProceso = ""
    
    NroIb.Text = ""
    NroInsc.Text = ""
    FechaNroInsc.Text = "  /  /    "
    Cufe.Text = ""
    CufeII.Text = ""
    CufeIII.Text = ""
    DirCufe.Text = ""
    DirCufeII.Text = ""
    DirCufeIII.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    VtoIso.Text = "  /  /    "
    
    Estado.ListIndex = 0
    Califica.ListIndex = 0
    IbCiudadII.ListIndex = 0
    FechaCalifica.Text = "  /  /    "
    FechaCategoria.Text = "  /  /    "
    
    Proveedor.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Proveedor.SetFocus
    PrgProve.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

    ZZSuma = 0
    Suma = 0
    ZZDirecion = ""

    spProveedor = "ListaProveedores"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                ZZRazon = !Nombre
                ZZEmail = IIf(IsNull(!EMail), "", !EMail)
                ZZEmail = Trim(ZZEmail)
                
                If ZZEmail <> "" Then
                
                    Rem ZZEmail = "hfrias@surfactan.com.ar"
                    
                    sTo = ZZEmail
                    sCC = ""
                    sBCC = ""
                    sSubject = "Aviso de Certificado de exclusión de IVA"
                    sBody = "                                                                                                  Victoria, Mayo  2009" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "Sres : " + ZZRazon + Chr$(13) _
                            + "Presente " + Chr$(13) _
                            + "" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "De nuestra consideración: " + Chr$(13) _
                            + "" + Chr$(13) _
                            + "Por medio de la presente les adjuntamos el certificado de exclusión de IVA" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "Sin mas saludamos cordialmente" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "                        Surfactan S.A."
                            
                    SFile = "c:\email\exclusioniva.pdf"
                
                    EmailAddress = sTo
                    CopiaAddress = sCC
                    MSubject = sSubject
                    MBody = sBody
                    MAttach = SFile
                    MAttachI = ""
                    MAttachII = ""
                    MAttachIII = ""
                    MAttachIV = ""
                    MAttachV = ""
                
                    SendEmail
                
                End If
                
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    
    m$ = "Proceso Finalizado"
    A% = MsgBox(m$, 0, "Envio de Avisio de Embargo")

End Sub

Private Sub Command2_Click()

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
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + " CodIbCaba = 2"
            spProveedor = ZSql
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Cicla

Stop
End Sub

Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "99999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
    End If
End Sub

Private Sub Direccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Localidad.SetFocus
    End If
End Sub

Private Sub Localidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Postal.SetFocus
    End If
End Sub



Private Sub Observa_Click()
    PantaObserva.Visible = True
    ObservacionesII.SetFocus
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Telefono.SetFocus
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dias.SetFocus
    End If
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EMail.SetFocus
    End If
End Sub

Private Sub EMail_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.SetFocus
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuenta.SetFocus
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstCuenta
            spCuenta = "ConsultaCuentas" + "'" + Cuenta.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                DesCuenta.Caption = rstCuenta!Descripcion
                NombreCheque.SetFocus
                    Else
                Cuenta.SetFocus
            End If
        End With
    End If
End Sub

Private Sub NombreCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroIb.SetFocus
    End If
End Sub

Private Sub NroIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroInsc.SetFocus
    End If
End Sub

Private Sub NroInsc_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaNroInsc.SetFocus
    End If
End Sub

Private Sub FechaNroInsc_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaNroInsc.Text, Auxi)
        If Auxi = "S" Then
            Cai.SetFocus
        End If
    End If
End Sub

Private Sub Cai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoCai.SetFocus
    End If
End Sub

Private Sub VtoCai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Proveedor.Text <> "" Then
        
            EmpresaReal = Wempresa
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Wempresa = EmpresaReal
            txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        
            Claveven$ = Proveedor.Text
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                    Proveedor.Text = RstProveedor!Proveedor
                    RstProveedor.Close
                    Call Imprime_Datos
                Else
                    CmdLimpiar_Click
                    Proveedor.Text = Claveven$
            End If
            
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        End If
        Nombre.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cufe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirCufe.SetFocus
    End If
End Sub

Private Sub DirCufe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CufeII.SetFocus
    End If
End Sub

Private Sub CufeII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirCufeII.SetFocus
    End If
End Sub

Private Sub DirCufeII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CufeIII.SetFocus
    End If
End Sub

Private Sub CufeIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirCufeIII.SetFocus
    End If
End Sub

Private Sub DirCufeIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cufe.SetFocus
    End If
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
        
            EmpresaReal = Wempresa
            Wempresa = "0001"
            txtOdbc = "Empresa01"
         Rem by nan
         Rem   strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
         Rem   Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Wempresa = EmpresaReal
            txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        
            spProveedor = "ListaProveedoresOrdConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
            
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            
        Case 1
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)

            With rstCuenta
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Cuenta) + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Select Case XIndice
        Case 0
        
            EmpresaReal = Wempresa
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Wempresa = EmpresaReal
            txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                RstProveedor.Close
                Call Imprime_Datos
                       Else
                CmdLimpiar_Click
                Proveedor.Text = Claveven$
            End If
            Ayuda.Visible = False
            Proveedor.SetFocus
            
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case 1
        
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas" + "'" + Claveven$ + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                Cuenta.Text = rstCuenta!Cuenta
                rstCuenta.Close
                Call Imprime_Descripcion
                    Else
                CmdLimpiar_Click
                Cuenta.Text = Claveven$
            End If

            Cuenta.SetFocus
        
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgProve.Caption = "Ingreso de Proveedores :  " + !Nombre
        End If
    End With

    Proveedor.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Cuit.Text = ""
    Telefono.Text = ""
    EMail.Text = ""
    Observaciones.Text = ""
    Dias.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    NombreCheque.Text = ""
    NroIb.Text = ""
    NroInsc.Text = ""
    FechaNroInsc.Text = "  /  /    "
    Cufe.Text = ""
    CufeII.Text = ""
    CufeIII.Text = ""
    DirCufe.Text = ""
    DirCufeII.Text = ""
    DirCufeIII.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    VtoIso.Text = "  /  /    "
    FechaCalifica.Text = "  /  /    "
    ObservacionesII.Text = ""
    FechaCategoria.Text = "  /  /    "
    Nombre.BackColor = &HFFFFFF
    
    
    WGraba = ""
    WProceso = ""
    
    Iva.Clear
    
    Iva.AddItem "No Inscripto"
    Iva.AddItem "Consumidor Final"
    Iva.AddItem "Resp.Inscripto"
    Iva.AddItem "Exento"
    Iva.AddItem "No Responsable"
    Iva.AddItem "Monotributo"
    Iva.AddItem "No Catalogado"
    Iva.AddItem ""
    
    Provincia.Clear
    
    Provincia.AddItem "Capital Federal"
    Provincia.AddItem "Buenos Aires"
    Provincia.AddItem "Catamarca"
    Provincia.AddItem "Cordoba"
    Provincia.AddItem "Corrientes"
    Provincia.AddItem "Chaco"
    Provincia.AddItem "Chubut"
    Provincia.AddItem "Entre Rios"
    Provincia.AddItem "Formosa"
    Provincia.AddItem "Jujuy"
    Provincia.AddItem "La Pampa"
    Provincia.AddItem "La Rioja"
    Provincia.AddItem "Mendoza"
    Provincia.AddItem "Misiones"
    Provincia.AddItem "Neuquen"
    Provincia.AddItem "Rio Negro"
    Provincia.AddItem "Salta"
    Provincia.AddItem "San Juan"
    Provincia.AddItem "San Luis"
    Provincia.AddItem "Santa Cruz"
    Provincia.AddItem "Santa Fe"
    Provincia.AddItem "Santiago del Estero"
    Provincia.AddItem "Tucuman"
    Provincia.AddItem "Tierra del Fuego"
    Provincia.AddItem "Exterior"
    Provincia.AddItem ""
    
    Region.Clear
    
    Region.AddItem "Fuera Mercosur"
    Region.AddItem "Mercosur"
    
    Tipo.Clear
    
    Tipo.AddItem "Bienes"
    Tipo.AddItem "Servicios"
    Tipo.AddItem "Alquileres"
    Tipo.AddItem "Exento"
    Tipo.AddItem "Despachante"
    Tipo.AddItem "Locacion de Obras"
    Tipo.AddItem "Fletes"
    Tipo.AddItem "Facturas (M)"
    Tipo.AddItem ""
    
    CodIb.Clear
    
    CodIb.AddItem "Bienes"
    CodIb.AddItem "Servicio"
    CodIb.AddItem "Exento"
    CodIb.AddItem "Ciudad Normal"
    CodIb.AddItem "Ciudad Riesgo"
    
    CodIbCaba.Clear
    
    CodIbCaba.AddItem ""
    CodIbCaba.AddItem ""
    CodIbCaba.AddItem "Exento"
    CodIbCaba.AddItem "Retiene"
    
    IbCiudadII.Clear
    
    IbCiudadII.AddItem ""
    IbCiudadII.AddItem "Local"
    IbCiudadII.AddItem "Conv. Multilateral"
    IbCiudadII.AddItem "No Inscripto"
    IbCiudadII.AddItem "Reg. Simplificado"
    
    
    IbCiudadII.ListIndex = 0
    
    
    TipoProv.Clear
    
    
    TipoProv.AddItem ""
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoProv"
    ZSql = ZSql + " Order by Codigo"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        With rstTipoProv
            .MoveFirst
            Do
                If .EOF = False Then
                    TipoProv.AddItem rstTipoProv!Descripcion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTipoProv.Close
    End If
    For Ciclo = 1 To 100
        TipoProv.AddItem ""
    Next Ciclo
    
    TipoListado.Clear
    
    TipoListado.AddItem "Completo"
    TipoListado.AddItem "Resumido"
    
    TipoListado.ListIndex = 0
    
    
    CategoriaI.Clear
    
    CategoriaI.AddItem ""
    CategoriaI.AddItem "A"
    CategoriaI.AddItem "B"
    CategoriaI.AddItem "C"
    CategoriaI.AddItem "E"
    
    CategoriaII.Clear
    
    CategoriaII.AddItem "Sin Calificar"
    CategoriaII.AddItem "Muy Bueno"
    CategoriaII.AddItem "Bueno"
    CategoriaII.AddItem "Regular"
    CategoriaII.AddItem "Malo"
    
    Iso.Clear
    
    Iso.AddItem ""
    Iso.AddItem "ISO 9001"
    Iso.AddItem "ISO 9001/14001"
    Iso.AddItem "ISO 17025"
    Iso.AddItem "SENASA"
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "Habilitado"
    Estado.AddItem "Inhabilitado"
    
    Califica.Clear
    
    Califica.AddItem ""
    Califica.AddItem "Apto"
    Califica.AddItem "Condicional"
    Califica.AddItem "No Apto"
    
    Iva.ListIndex = 7
    Tipo.ListIndex = 8
    Provincia.ListIndex = 25
    Region.ListIndex = 0
    CodIb.ListIndex = 0
    CodIbCaba.ListIndex = 2
    TipoProv.ListIndex = 0
    CategoriaI.ListIndex = 0
    CategoriaII.ListIndex = 0
    Iso.ListIndex = 0
    Estado.ListIndex = 0
    Califica.ListIndex = 0
    IbCiudadII.ListIndex = 0
    Estado.BackColor = &H80000005

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
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
    
    RstProveedor.Close
    
    End If
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    End If

End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
    
    spProveedor = "ListaProveedores"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Proveedor.Text = RstProveedor!Proveedor
            RstProveedor.Close
            Call Imprime_Datos
        End With
    End If
    Proveedor.SetFocus
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Proveedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Proveedor.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
    
    spProveedor = "ListaProveedores"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveLast
            Proveedor.Text = RstProveedor!Proveedor
            RstProveedor.Close
            Call Imprime_Datos
        End With
    End If
    Proveedor.SetFocus
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Proveedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Proveedor.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
    
    spProveedor = "AnteriorProveedor " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveLast
            Proveedor.Text = RstProveedor!Proveedor
            RstProveedor.Close
            Call Imprime_Datos
        End With
    End If
    
    Proveedor.SetFocus
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Proveedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Proveedor.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    EmpresaReal = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Wempresa = EmpresaReal
    txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
    
    spProveedor = "PosteriorProveedor " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Proveedor.Text = RstProveedor!Proveedor
            RstProveedor.Close
            Call Imprime_Datos
        End With
    End If
    
    Proveedor.SetFocus
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Proveedor", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Proveedor.SetFocus
    
End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        ZGrabaIV = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGrabaIV = IIf(IsNull(rstOperador!GrabaIV), "", rstOperador!GrabaIV)
            rstOperador.Close
        End If
        
        If ZGrabaIV = "S" Then
            WGraba = "S"
            XClave.Visible = False
            If WProceso = "A" Then
                Call cmdAdd_Click
            End If
            If WProceso = "B" Then
                Call cmdDelete_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClave.SetFocus
        End If
        
    End If
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
    
    
    
    
    









Private Sub Eval()

    Es = Trim(WCuit)

    XXX = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               XXX = XXX + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               XXX = XXX + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               XXX = XXX + Y: MinusOk = 0

        End If

    Next

    WCuit = XXX

End Sub



