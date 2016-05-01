VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgTermi 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Producto Terminado"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   Begin VB.CommandButton Command123 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1920
      TabIndex        =   192
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   191
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame CargaIngles 
      Caption         =   "Descripciones en Ingles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   156
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton CierraIngles 
         Caption         =   "Cierra Pantalla"
         Height          =   615
         Left            =   3120
         TabIndex        =   162
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox ConservacionIngles 
         Height          =   285
         Left            =   2400
         MaxLength       =   80
         TabIndex        =   161
         Text            =   " "
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox ConservacionIIIngles 
         Height          =   285
         Left            =   2400
         MaxLength       =   80
         TabIndex        =   160
         Text            =   " "
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox DescripcionIngles 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   158
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox DescriEtiquetaIngles 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   157
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label LabelDescriEtiquetaIngles 
         Caption         =   "Desc.P/Farma 2 Reng"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label60 
         Caption         =   "Conservacion"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   186
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label59 
         Caption         =   "Obs.Cliente"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   185
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label51 
         Caption         =   "Obs. Interna"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   184
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Descripcion"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   159
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox Opcion 
      Height          =   1620
      ItemData        =   "terminado.frx":0000
      Left            =   4080
      List            =   "terminado.frx":0002
      TabIndex        =   65
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   2760
      TabIndex        =   94
      Top             =   3720
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      Height          =   2400
      ItemData        =   "terminado.frx":0004
      Left            =   3120
      List            =   "terminado.frx":000B
      TabIndex        =   41
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Commanddada 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2040
      TabIndex        =   183
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox FabricaIII 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   182
      Text            =   " "
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox FabricaII 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   181
      Text            =   " "
      Top             =   2880
      Width           =   855
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
      Left            =   5160
      TabIndex        =   179
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   178
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox ImpreVto 
      Height          =   315
      Left            =   6120
      TabIndex        =   169
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox Sedronar 
      Height          =   315
      Left            =   2760
      TabIndex        =   167
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton AltaIngles 
      Caption         =   "Carga Descripcion en Ingles"
      Height          =   615
      Left            =   7320
      TabIndex        =   163
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   435
      Left            =   960
      TabIndex        =   155
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame PantaLiberaHoja 
      Height          =   1695
      Left            =   3720
      TabIndex        =   151
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaLiberaHoja 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   153
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox WClaveLiberaHoja 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   152
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label50 
         Caption         =   "Ingrese se Password"
         Height          =   375
         Left            =   840
         TabIndex        =   154
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox HojaTecnica 
      Height          =   285
      Left            =   5160
      TabIndex        =   149
      Top             =   7440
      Width           =   3735
   End
   Begin VB.ComboBox Marca 
      Height          =   315
      Left            =   6000
      TabIndex        =   148
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox DescriEtiqueta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   146
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      Height          =   3615
      Left            =   7680
      TabIndex        =   132
      Top             =   -120
      Width           =   4215
      Begin VB.TextBox Caracteristicas 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   176
         Text            =   " "
         Top             =   1750
         Width           =   2775
      End
      Begin VB.TextBox Clase 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   172
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Secundario 
         Height          =   285
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   171
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Riesgo 
         Height          =   285
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   170
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox Carga 
         Height          =   315
         Left            =   1320
         TabIndex        =   145
         Top             =   2640
         Width           =   2775
      End
      Begin VB.ComboBox EstadoProducto 
         Height          =   315
         Left            =   1320
         TabIndex        =   144
         Top             =   2920
         Width           =   2775
      End
      Begin VB.ComboBox ListaProducto 
         Height          =   315
         Left            =   1320
         TabIndex        =   142
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox Seguridad 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   141
         Text            =   " "
         Top             =   2350
         Width           =   2775
      End
      Begin VB.TextBox Intervencion 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2050
         Width           =   1815
      End
      Begin VB.TextBox Naciones 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Embalaje 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   26
         Top             =   1155
         Width           =   1815
      End
      Begin VB.TextBox Observaciones 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Tipoeti 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   23
         Top             =   520
         Width           =   1815
      End
      Begin VB.Label Label44 
         Caption         =   "Caracteristicas"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   177
         Top             =   1750
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Clase"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   175
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label56 
         Caption         =   "R.Sec."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   174
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label57 
         Caption         =   "Riesgo"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2760
         TabIndex        =   173
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label47 
         Caption         =   "Lista Producto"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   143
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "Carga"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   140
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Estado"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   139
         Top             =   2950
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "F.Intervencion"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   2050
         Width           =   2175
      End
      Begin VB.Label Label23 
         Caption         =   "Nro. N.Unidas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Grupo Embalaje"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1155
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "Observaciones"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Tipo de Etiqueta"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "Hoja Seguridad"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   133
         Top             =   2350
         Width           =   1095
      End
   End
   Begin VB.Frame PantaAutoriza 
      Height          =   1695
      Left            =   5040
      TabIndex        =   127
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClaveAutoriza 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   129
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CancelaAutoriza 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   128
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label42 
         Caption         =   "Ingrese se Password"
         Height          =   375
         Left            =   840
         TabIndex        =   130
         Top             =   120
         Width           =   2175
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
      Height          =   1695
      Left            =   2040
      Picture         =   "terminado.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Clave 
      Caption         =   "Control de Grabacion"
      Height          =   1695
      Left            =   4200
      TabIndex        =   87
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   90
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   89
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese se Password"
         Height          =   375
         Left            =   720
         TabIndex        =   88
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.TextBox DesEfluentes 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Efluentes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   15
      Text            =   " "
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Metodo 
      Height          =   285
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   14
      Text            =   " "
      Top             =   3240
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Height          =   4335
      Left            =   9000
      TabIndex        =   102
      Top             =   3600
      Width           =   2775
      Begin VB.TextBox ObservaII 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox EstadoII 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox VersionII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox FechaVersionII 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox ObservaI 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox EstadoI 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox VersionI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox FechaVersionI 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Observa 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Estado 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox FechaVersion 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Version 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label39 
         Caption         =   "E S P E C I F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   120
         TabIndex        =   123
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   120
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Fec. Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   119
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "P R O C E S O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   120
         TabIndex        =   116
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label37 
         Caption         =   "F O R M U L A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label35 
         Caption         =   "Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   112
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Fec. Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   111
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label rty 
         Caption         =   "Fec. Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   106
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label ert 
         Caption         =   "Version"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   104
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox ConservacionII 
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   13
      Text            =   " "
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2295
      Left            =   4800
      TabIndex        =   43
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox HastaLinea 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   85
         Text            =   " "
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox DesdeLinea 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   84
         Text            =   " "
         Top             =   1080
         Width           =   735
      End
      Begin MSMask.MaskEdBox HastaCodigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   62
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdecodigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   61
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   49
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Hasta Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Desde Linea"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Conservacion 
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   12
      Text            =   " "
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Vida 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Minimo1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   6
      Text            =   " "
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Escrito 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton HojaPend 
      Caption         =   "    Hojas de Produccion                   Pendientes"
      Height          =   615
      Left            =   5160
      TabIndex        =   95
      Top             =   6720
      Width           =   2055
   End
   Begin VB.ComboBox Controla 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Impreadi 
      Height          =   285
      Left            =   6960
      MaxLength       =   1
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   67
      Top             =   3960
      Width           =   8775
      Begin VB.TextBox Envase6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         TabIndex        =   21
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Envase5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Text            =   " "
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Envase4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         TabIndex        =   19
         Text            =   " "
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Envase3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Envase2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   " "
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Envase1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   " "
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6240
         TabIndex        =   79
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6240
         TabIndex        =   78
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6240
         TabIndex        =   77
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Envase 6"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   76
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Envase 5"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   75
         Top             =   530
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Envase 4"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   74
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1800
         TabIndex        =   73
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1800
         TabIndex        =   72
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1800
         TabIndex        =   71
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Envase 3"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Envase 2"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   530
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Envase 1"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   200
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox Linea 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Deposito 
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   7
      Text            =   " "
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Unidad 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Minimo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   5
      Text            =   " "
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Salidas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   59
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Entradas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   58
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Inicial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   57
      Text            =   " "
      Top             =   720
      Width           =   1335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8280
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wterminado.rpt"
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
      Left            =   2400
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   5160
      TabIndex        =   40
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5160
      TabIndex        =   39
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   7320
      TabIndex        =   34
      Top             =   5280
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   6240
      TabIndex        =   28
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   33
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   6240
      TabIndex        =   32
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   5160
      TabIndex        =   31
      Top             =   5280
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
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox Fabrica 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   11
      Text            =   " "
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Autoriza 
      Caption         =   "Conformidad de Versiones de P.T."
      Height          =   615
      Left            =   3480
      TabIndex        =   126
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
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
      MaxLength       =   50
      TabIndex        =   189
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label LabelDescriEtiqueta 
      Caption         =   "Desc.P/Farma 2 Reng"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   187
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label58 
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
      Left            =   3720
      TabIndex        =   180
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label55 
      Caption         =   "Vto:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   168
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label54 
      Caption         =   "Sedronar"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2040
      TabIndex        =   166
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label53 
      Caption         =   "Obs.Cliente"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   165
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label52 
      Caption         =   "Obs. Interna"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   164
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label49 
      Caption         =   "Hoja Tecnica"
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
      TabIndex        =   150
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label48 
      Caption         =   "Autoriz."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5160
      TabIndex        =   147
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label43 
      Caption         =   "Lote Fab."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   131
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label41 
      Caption         =   "Efluentes    de Lavado"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   5640
      TabIndex        =   125
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label40 
      Caption         =   "Met.Lav."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   124
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label31 
      Caption         =   "Conservacion"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   100
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "Vida Util Meses"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   99
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Pedidos Pend."
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
      Left            =   4440
      TabIndex        =   98
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label pedido 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5760
      TabIndex        =   97
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Procdim.Farma"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1200
      TabIndex        =   96
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Re 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   93
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "Cantidad Re"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   92
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad Inicial"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "Eti.Adicional"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   91
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Proceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  "
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
      Height          =   300
      Left            =   1440
      TabIndex        =   86
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Nk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4200
      TabIndex        =   81
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label aa 
      Caption         =   "Proceso"
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
      TabIndex        =   80
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label asd 
      Caption         =   "Cantidad Nk"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2880
      TabIndex        =   66
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label DescriLinea 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4440
      TabIndex        =   64
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Stock 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   63
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Linea"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   720
      Width           =   615
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
      TabIndex        =   56
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Deposito"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Unidad Medida"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Minimo Consol/Planta"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   53
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad Salida"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3000
      TabIndex        =   52
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad Entrada"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   30
      Top             =   0
      Width           =   1215
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
      TabIndex        =   29
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label Label61 
      Caption         =   "Cod. Sedronar"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   190
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "PrgTermi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstLineas As Recordset
Dim spLineas As String
Dim rstEfluentes As Recordset
Dim spEfluentes As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstPeligroso As Recordset
Dim spPeligroso As String
Dim XParam As String
Private WGraba As String
Private WProceso As String
Private WNk As String
Private WRe As String
Dim CargaEmpresa(12, 2) As String
Dim ZAyuda As String

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
Dim ZCampo25 As String
Dim ZCampo26 As String
Dim ZCampo27 As String
Dim ZCampo28 As String
Dim ZCampo29 As String
Dim ZCampo30 As String
Dim ZCampo31 As String
Dim ZCampo32 As String

Dim PasaControla As String
Dim PasaEscrito As String
Dim PasaCarga As String
Dim PasaEstadoProducto As String
Dim PasaListaProducto As String

Dim ZVector(5000, 10) As String

Dim ZZPorce As Double

Dim ZZRestriccion As Integer
Dim WRestriccion As String





Sub Imprime_Datos()

    Rem On Error GoTo WError

    ClavePro$ = "NK" + Right$(Codigo.Text, 10)
    spTerminado = "ConsultaTerminado " + "'" + ClavePro$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Nk.Caption = Pusing("###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        Nk.Caption = "0.00"
    End If
    
    ClavePro$ = "RE" + Right$(Codigo.Text, 10)
    spTerminado = "ConsultaTerminado " + "'" + ClavePro$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        RE.Caption = Pusing("###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        RE.Caption = "0.00"
    End If
        
    spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Codigo.Text = rstTerminado!Codigo
        Descripcion.Text = rstTerminado!Descripcion
        DescriEtiqueta.Text = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
        Linea.Text = rstTerminado!Linea
        Unidad.Text = rstTerminado!Unidad
        Inicial.Text = Str$(rstTerminado!Inicial)
        Entradas.Text = Str$(rstTerminado!Entradas)
        Salidas.Text = Str$(rstTerminado!Salidas)
        Minimo.Text = Str$(rstTerminado!Minimo)
        Minimo1.Text = IIf(IsNull(rstTerminado!Minimo1), "0", rstTerminado!Minimo1)
        Fabrica.Text = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
        FabricaII.Text = IIf(IsNull(rstTerminado!FabricaII), "0", rstTerminado!FabricaII)
        FabricaIII.Text = IIf(IsNull(rstTerminado!FabricaIII), "0", rstTerminado!FabricaIII)
        Rem LoteAutorizado.Text = IIf(IsNull(rstTerminado!LoteAutorizado), "", rstTerminado!LoteAutorizado)
        Deposito.Text = IIf(IsNull(rstTerminado!Deposito), "", rstTerminado!Deposito)
        Rem Pedido.text = Str$(rstTerminado!Pedido)
        Rem Envase.text = rstTerminado!Envase
        Envase1.Text = rstTerminado!Envase1
        Envase2.Text = rstTerminado!Envase2
        Envase3.Text = rstTerminado!Envase3
        Envase4.Text = rstTerminado!Envase4
        Envase5.Text = rstTerminado!Envase5
        Envase6.Text = rstTerminado!Envase6
        Proceso.Caption = Str$(rstTerminado!Proceso)
        Impreadi.Text = ""
        Clase.Text = ""
        Secundario.Text = ""
        Riesgo.Text = ""
        Intervencion.Text = ""
        Naciones.Text = ""
        Embalaje.Text = ""
        Impreadi.Text = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
        Clase.Text = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
        Secundario.Text = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
        Riesgo.Text = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
        Intervencion.Text = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
        Naciones.Text = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
        Embalaje.Text = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
        Controla.ListIndex = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
        Sedronar.ListIndex = IIf(IsNull(rstTerminado!Sedronar), "0", rstTerminado!Sedronar)
        CodSedronar.Text = IIf(IsNull(rstTerminado!CodSedronar), "", rstTerminado!CodSedronar)
        ImpreVto.ListIndex = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
        Marca.ListIndex = IIf(IsNull(rstTerminado!Marca), "0", rstTerminado!Marca)
        Observaciones.Text = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
        TipoEti.Text = IIf(IsNull(rstTerminado!TipoEti), "", rstTerminado!TipoEti)
        Escrito.ListIndex = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
        Pedido.Caption = Str$(rstTerminado!Pedido)
        Conservacion.Text = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
        Conservacion.Text = RTrim(Conservacion.Text)
        ConservacionII.Text = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
        ConservacionII.Text = RTrim(ConservacionII.Text)
        Vida.Text = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
        Seguridad.Text = IIf(IsNull(rstTerminado!Seguridad), "", rstTerminado!Seguridad)
        
        Version.Text = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
        VersionI.Text = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
        VersionII.Text = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
        
        FechaVersion.Text = IIf(IsNull(rstTerminado!FechaVersion), "  /  /    ", rstTerminado!FechaVersion)
        FechaVersionI.Text = IIf(IsNull(rstTerminado!FechaVersionI), "  /  /    ", rstTerminado!FechaVersionI)
        FechaVersionII.Text = IIf(IsNull(rstTerminado!FechaVersionII), "  /  /    ", rstTerminado!FechaVersionII)
        
        Estado.Text = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
        EstadoI.Text = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
        EstadoII.Text = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
        
        Observa.Text = IIf(IsNull(rstTerminado!Observa), "", rstTerminado!Observa)
        ObservaI.Text = IIf(IsNull(rstTerminado!ObservaI), "", rstTerminado!ObservaI)
        ObservaII.Text = IIf(IsNull(rstTerminado!ObservaII), "", rstTerminado!ObservaII)
        
        Metodo.Text = IIf(IsNull(rstTerminado!Metodo), "", rstTerminado!Metodo)
        Efluentes.Text = IIf(IsNull(rstTerminado!Efluentes), "", rstTerminado!Efluentes)
        
        Caracteristicas.Text = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
        Carga.ListIndex = IIf(IsNull(rstTerminado!Carga), "0", rstTerminado!Carga)
        EstadoProducto.ListIndex = IIf(IsNull(rstTerminado!EstadoProducto), "0", rstTerminado!EstadoProducto)
        ListaProducto.ListIndex = IIf(IsNull(rstTerminado!ListaProducto), "0", rstTerminado!ListaProducto)
        
        DescripcionIngles.Text = IIf(IsNull(rstTerminado!DescripcionIngles), "", rstTerminado!DescripcionIngles)
        DescriEtiquetaIngles.Text = IIf(IsNull(rstTerminado!DescriEtiquetaIngles), "", rstTerminado!DescriEtiquetaIngles)
        ConservacionIngles.Text = IIf(IsNull(rstTerminado!ConservacionIngles), "", rstTerminado!ConservacionIngles)
        ConservacionIIIngles.Text = IIf(IsNull(rstTerminado!ConservacionIIIngles), "", rstTerminado!ConservacionIIIngles)
        
        responsa.Text = IIf(IsNull(rstTerminado!Responsable), "", rstTerminado!Responsable)
        
        Naciones.Text = Trim(Naciones.Text)
        
        Impreadi.Text = Trim(Impreadi.Text)
        TipoEti.Text = Trim(TipoEti.Text)
        
        ZZRestriccion = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
        Restriccion.Value = ZZRestriccion
        
        rstTerminado.Close
    End If
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            MiRuta = "\\193.168.0.2\Impresion pdf\DOC" + Mid$(Codigo.Text, 4, 5) + Right$(Codigo.Text, 3) + "*.PDF"
            MiNombre = Dir(MiRuta)
            HojaTecnica.Text = MiNombre
            
        Case Else
    End Select
    
        
    
    
    
    
    
    
    WTerminado = Codigo.Text
    XCodigo = Val(Mid$(WTerminado, 4, 5))
    XTipoPro = ""
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
    
    Select Case Val(Linea.Text)
        Case 8
            XTipoPro = "PG"
        Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
            XTipoPro = "FA"
        Case Else
    End Select
    
    If XTipoPro = "FA" Then
        LabelDescriEtiqueta.Visible = True
        DescriEtiqueta.Visible = True
        LabelDescriEtiquetaIngles.Visible = True
        DescriEtiquetaIngles.Visible = True
            Else
        LabelDescriEtiqueta.Visible = False
        DescriEtiqueta.Visible = False
        LabelDescriEtiquetaIngles.Visible = False
        DescriEtiquetaIngles.Visible = False
    End If
    
    
    
    
    
    
    
    Call Format_datos
    
    Call Imprime_Descripcion
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Command9_Click()

    Dim ZCarga(10000) As String
    
    ZLugar = 0
    Erase ZCarga

    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                        
                    ZVida = IIf(IsNull(!Vida), "0", !Vida)
                    
                    If ZVida > 12 Then
                    
                        ZZEntra = "S"
                    
                        ZZCodigo = Mid$(!Codigo, 1, 2)
                        ZZCodigoI = Mid$(!Codigo, 4, 5)
                        ZZCodigoII = Mid$(!Codigo, 10, 3)
                        
                        If ZZCodigo = "PT" Then
                    
                            If Val(ZZCodigoI) >= 0 And Val(ZZCodigoI) <= 999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 3000 And Val(ZZCodigoI) <= 3999 Then
                                If Val(ZZCodigoII) <> 100 Then
                                    ZZEntra = "N"
                                End If
                            End If
                            If Val(ZZCodigoI) >= 9700 And Val(ZZCodigoI) <= 9799 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 9800 And Val(ZZCodigoI) <= 9899 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 11000 And Val(ZZCodigoI) <= 12999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 21000 And Val(ZZCodigoI) <= 21999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 25000 And Val(ZZCodigoI) <= 25999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 27000 And Val(ZZCodigoI) <= 27999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 40000 And Val(ZZCodigoI) <= 49999 Then
                                ZZEntra = "N"
                            End If
                            If Val(ZZCodigoI) >= 50000 And Val(ZZCodigoI) <= 59999 Then
                                ZZEntra = "N"
                            End If
                        
                            If ZZEntra = "S" Then
                                ZLugar = ZLugar + 1
                                ZCarga(ZLugar) = !Codigo
                            End If
                        
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
        
        
        
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
    
    WVida = "12"
        
    For Cicla = 1 To 7

      If CargaEmpresa(Cicla, 1) <> "" Then

          Wempresa = CargaEmpresa(Cicla, 1)
          txtOdbc = CargaEmpresa(Cicla, 2)
          strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
          Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            For Ciclo = 1 To ZLugar
            
                ZZProducto = ZCarga(Ciclo)
          
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "Vida = " + "'" + WVida + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + ZZProducto + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
            Next Ciclo
                  
        End If
        
    Next Cicla
          
        
    Select Case Val(XEmpresa)
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
    
End Sub

Private Sub Command11_Click()

    Erase ZVector
    Lugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Peligroso"
    spPeligroso = Sql1 + Sql2
    Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
    If rstPeligroso.RecordCount > 0 Then
        With rstPeligroso
            .MoveFirst
            Do
                If .EOF = False Then
                    Lugar = Lugar + 1
                    ZVector(Lugar, 1) = rstPeligroso!Codigo
                    ZVector(Lugar, 2) = IIf(IsNull(rstPeligroso!Secundario), "", rstPeligroso!Secundario)
                    ZVector(Lugar, 3) = IIf(IsNull(rstPeligroso!Riesgo), "", rstPeligroso!Riesgo)
                    ZVector(Lugar, 4) = IIf(IsNull(rstPeligroso!Embalaje), "", rstPeligroso!Embalaje)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPeligroso.Close
    End If
    
    For XCiclo = 1 To Lugar
    
        XXCodigo = ZVector(XCiclo, 1)
        XXSecundario = ZVector(XCiclo, 2)
        XXRiesgo = ZVector(XCiclo, 3)
        XXEmbalaje = ZVector(XCiclo, 4)
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Peligroso SET "
        ZSql = ZSql & "Secundario = " + "'" + XXSecundario + "',"
        ZSql = ZSql & "Riesgo = " + "'" + XXRiesgo + "',"
        ZSql = ZSql & "Embalaje = " + "'" + XXEmbalaje + "'"
        ZSql = ZSql & " Where Codigo = " + "'" + XXCodigo + "'"
            
        spPeligroso = ZSql
        Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
        
    Next XCiclo
    
End Sub


Private Sub Command111_Click()
    Panta1.Visible = True
End Sub

Private Sub Command3_Click()

    Rem dada
    Rem dada
    Rem dada
    Dim ZActualiza(10000, 2) As String
    Dim ZLugar As Integer
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If Left$(UCase(rstTerminado!Codigo), 2) = "PT" Then
                        ZLugar = ZLugar + 1
                        ZActualiza(ZLugar, 1) = rstTerminado!Codigo
                        ZActualiza(ZLugar, 2) = ""
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
    
    XEmpresa = Wempresa
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZActualiza(Ciclo, 1)
    
        spTerminado = "ConsultaTerminado " + "'" + ZCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZActualiza(Ciclo, 2) = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
            rstTerminado.Close
                Else
            ZActualiza(Ciclo, 1) = ""
        End If
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZActualiza(Ciclo, 1)
        ZVida = ZActualiza(Ciclo, 2)
        
        If ZCodigo <> "" Then
        
            ZSql = ""
            ZSql = ZSql & "UPDATE Terminado SET "
            ZSql = ZSql & "Vida = " + "'" + ZVida + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + ZCodigo + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    m$ = "Proceso Finalizado"
    G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")

End Sub

Private Sub Command333_Click()
    
    dada1.Clear
    dada2.Clear
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosII"
    ZSql = ZSql + " Where PreciosII.Terminado = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Clave"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosII.RecordCount > 0 Then
    
        With rstPreciosII
            .MoveFirst
            Do
                If .EOF = False Then
    
                    dada1.AddItem rstPreciosII!Nombre
                    dada2.AddItem rstPreciosII!Nombre
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosII.Close
    End If
    
    dada1.ListIndex = 0
    dada2.ListIndex = 0
    
    Panta2.Visible = True
End Sub

Private Sub Command222_Click()

    Dim ZZDada(10000, 10) As String
    
    Erase ZZDada
    ZZLugar = 0
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If Left$(rstTerminado!Codigo, 5) = "PT-25" Then
                        ZZLugar = ZZLugar + 1
                        ZZDada(ZZLugar, 1) = rstTerminado!Codigo
                        ZZDada(ZZLugar, 2) = rstTerminado!loteautoriza
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
        
    End If

    Wempresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    For ZZCiclo = 1 To ZZLugar
    
        ZZCodigo = ZZDada(ZZCiclo, 1)
        ZZLote = ZZDada(ZZCiclo, 2)
    
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM Terminado"
        ZSql = ZSql & " Where Terminado.Producto = " + "'" + ZZCodigo + "'"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZZTerminadoII = rstTerminado!lotefabricacion
            rstEspecifUnifica.Close
                Else
            ZZVersionII = 0
        End If
        
        If Val(ZZVersion) <> Val(ZZVersionII) Then
        
            Dim CargaEmpresa(10, 10) As String
            
            Dada = ZZCodigo
            
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
            
            For Cicla = 1 To 7
            
                Wempresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                WCodigo = WProducto
                WVersionII = ZVersion
                WFechaVersionII = ZFecha
                WEstadoII = ZEstado
                WObservaII = ZObservaciones
            
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "VersionII = " + "'" + Str$(ZZVersionII) + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + ZZCodigo + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
            Next Cicla
            
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        End If
        
    Next ZZCiclo

End Sub

Private Sub Command123_Click()
    
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
    
    LugarPlanilla = 1
    ruta = "C:\david\farma.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
    
            
            
            ZZCodigo = appExcel.cells(LugarPlanilla, 1).Value
            ZZLinea10 = appExcel.cells(LugarPlanilla, 5).Value
            ZZLinea24 = appExcel.cells(LugarPlanilla, 6).Value
            ZZLinea22 = appExcel.cells(LugarPlanilla, 7).Value
            ZZLinea25 = appExcel.cells(LugarPlanilla, 8).Value
            ZZLinea4 = appExcel.cells(LugarPlanilla, 9).Value
            ZZLinea26 = appExcel.cells(LugarPlanilla, 10).Value
            ZZLinea27 = appExcel.cells(LugarPlanilla, 11).Value
            ZZLinea44 = appExcel.cells(LugarPlanilla, 12).Value
            ZZLinea20 = appExcel.cells(LugarPlanilla, 13).Value
                    
            If Trim(ZZCodigo) = "" Then Exit Do
            
            ZZCodigo = "PT-" + ZZCodigo
            
            If UCase(Trim(ZZLinea10)) = "X" Then
                WLinea = "10"
            End If
            
            If UCase(Trim(ZZLinea24)) = "X" Then
                WLinea = "24"
            End If
            
            If UCase(Trim(ZZLinea22)) = "X" Then
                WLinea = "22"
            End If
            
            If UCase(Trim(ZZLinea25)) = "X" Then
                WLinea = "25"
            End If
            
            If UCase(Trim(ZZLinea4)) = "X" Then
                WLinea = "29"
            End If
            
            If UCase(Trim(ZZLinea26)) = "X" Then
                WLinea = "26"
            End If
            
            If UCase(Trim(ZZLinea27)) = "X" Then
                WLinea = "27"
            End If
            
            If UCase(Trim(ZZLinea44)) = "X" Then
                WLinea = "30"
            End If
            
            If UCase(Trim(ZZLinea20)) = "X" Then
                WLinea = "20"
                Stop
            End If
                    
            If Val(WLinea) = 29 Or Val(WLinea) = 30 Then
                        
                For Cicla = 1 To 7
                
                    If CargaEmpresa(Cicla, 1) <> "" Then
                
                        Wempresa = CargaEmpresa(Cicla, 1)
                        txtOdbc = CargaEmpresa(Cicla, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                        Rem ZSql = ""
                        Rem ZSql = ZSql & "UPDATE Terminado SET "
                        Rem ZSql = ZSql & "Linea = " + "'" + WLinea + "'"
                        Rem ZSql = ZSql & " Where Codigo = " + "'" + Trim(UCase(ZZCodigo)) + "'"
                        Rem spTerminado = ZSql
                        Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                
                        ZSql = ""
                        ZSql = ZSql & "UPDATE Estadistica SET "
                        ZSql = ZSql & "Linea = " + "'" + WLinea + "'"
                        ZSql = ZSql & " Where Articulo = " + "'" + Trim(UCase(ZZCodigo)) + "'"
                        ZSql = ZSql & " and OrdFecha >= " + "'" + "20160101" + "'"
                        spEstadistica = ZSql
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                    End If
                    
                Next Cicla
            
            End If
            
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
Stop

End Sub

Private Sub Command2_Click()


    Dim ZZVector(1000)

    ZZLugar = 0
    Erase ZZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Seguro"
    ZSql = ZSql + " Order by Seguro.Codigo"
    
    spSeguro = ZSql
    Set rstSeguro = db.OpenRecordset(spSeguro, dbOpenSnapshot, dbSQLPassThrough)
    If rstSeguro.RecordCount > 0 Then
    
        With rstSeguro
            .MoveFirst
            Do
                If .EOF = False Then
                    If Left$(rstSeguro!Codigo, 2) = "PT" Or Left$(rstSeguro!Codigo, 2) = "SE" Or Left$(rstSeguro!Codigo, 2) = "NK" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = rstSeguro!Codigo
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstSeguro.Close
    
    End If
    
    For ZZCiclo = 1 To ZZLugar
    
        ZZCodigo = ZZVector(ZZCiclo)
        


        spTerminado = "ConsultaTerminado " + "'" + ZZCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Codigo.Text = rstTerminado!Codigo
            ZZDescripcion = rstTerminado!Descripcion
            rstTerminado.Close

            Es = ZZDescripcion
            x = ""
            For XX = 1 To Len(Es)
                Y = Mid$(Es, XX, 1)
                If Y <> " " And Y <> "/" Then
                    x = x + Y
                End If
            Next
            ZZCodArt = x + Mid$(Codigo.Text, 4, 5) + Right$(Codigo.Text, 3)
            
            ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArt + ".PDF"
            ZZEstado = Dir(ZZRuta)
            ZZEstado = Trim(ZZEstado)
            If ZZEstado <> "" Then
                ZZRutaII = "c:\pasa\MSDS" + ZZCodArt + ".PDF"
                FileCopy ZZRuta, ZZRutaII
            End If
        
        End If
    
    Next ZZCiclo
End Sub

Private Sub Command4_Click()

    Dim ZCarga(10000, 10) As String
    
    ZLugar = 0
    Erase ZCarga

    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(!Codigo, 2)) = "PT" Then
                        
                        ZZImpreAdi = ""
                        ZZClase = ""
                        ZZSecundario = ""
                        ZZRiesgo = ""
                        ZZIntervencion = ""
                        ZZNaciones = ""
                        ZZEmbalaje = ""
                        
                        ZZImpreAdi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
                        ZZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                        ZZSecundario = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
                        ZZRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
                        ZZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                        ZZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                        ZZEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                        ZZDescriOnu = IIf(IsNull(rstTerminado!Descrionu), "", rstTerminado!Descrionu)
                        
                        ZLugar = ZLugar + 1
                        ZCarga(ZLugar, 1) = rstTerminado!Codigo
                        ZCarga(ZLugar, 2) = ZZImpreAdi
                        ZCarga(ZLugar, 3) = ZZClase
                        ZCarga(ZLugar, 4) = ZZSecundario
                        ZCarga(ZLugar, 5) = ZZRiesgo
                        ZCarga(ZLugar, 6) = ZZIntervencion
                        ZCarga(ZLugar, 7) = ZZNaciones
                        ZCarga(ZLugar, 8) = ZZEmbalaje
                        ZCarga(ZLugar, 9) = ZZDescriOnu
                        
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
        
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
    
    For Cicla = 1 To 7
    
        If CargaEmpresa(Cicla, 1) <> "" Then
    
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
              
            For Ciclo = 1 To ZLugar
    
                ZZCodigo = ZCarga(Ciclo, 1)
                ZZImpreAdi = ZCarga(Ciclo, 2)
                ZZClase = ZCarga(Ciclo, 3)
                ZZSecundario = ZCarga(Ciclo, 4)
                ZZRiesgo = ZCarga(Ciclo, 5)
                ZZIntervencion = ZCarga(Ciclo, 6)
                ZZNaciones = ZCarga(Ciclo, 7)
                ZZEmbalaje = ZCarga(Ciclo, 8)
                ZZDescriOnu = ZCarga(Ciclo, 9)
                
                For ZZCicla = 1 To 3
                
                    Select Case ZZCicla
                        Case 1
                            WCodigo = "NK" + Right$(ZZCodigo, 10)
                        Case 2
                            WCodigo = "RE" + Right$(ZZCodigo, 10)
                        Case Else
                            WCodigo = "SE" + Right$(ZZCodigo, 10)
                    End Select
        
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "ImpreAdi = " + "'" + ZZImpreAdi + "',"
                    ZSql = ZSql & "Clase = " + "'" + ZZClase + "',"
                    ZSql = ZSql & "Secundario = " + "'" + ZZSecundario + "',"
                    ZSql = ZSql & "Riesgo = " + "'" + ZZRiesgo + "',"
                    ZSql = ZSql & "Intervencion = " + "'" + ZZIntervencion + "',"
                    ZSql = ZSql & "Naciones = " + "'" + ZZNaciones + "',"
                    ZSql = ZSql & "Embalaje = " + "'" + ZZEmbalaje + "',"
                    ZSql = ZSql & "DescriOnu = " + "'" + ZZDescriOnu + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                              
                Next ZZCicla
                
            Next Ciclo
                  
        End If
            
    Next Cicla
        
    Select Case Val(XEmpresa)
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

End Sub

Private Sub Command5_Click()
    AdminNombre.Visible = False
End Sub

Private Sub Commanddada_Click()



    Dim ZZVector(10000)


    ZSql = ""
    ZSql = ZSql & "UPDATE Terminado SET "
    ZSql = ZSql & "PorceSedro = " + "'" + "0" + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)


    ZZLugar = 0
    Erase ZZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Order by Terminado.Codigo"
    
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If Left$(rstTerminado!Codigo, 2) = "PT" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = rstTerminado!Codigo
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    
    End If
    
    For ZZCiclo = 1 To ZZLugar
    
        ZZCodigo = ZZVector(ZZCiclo)
        

        Dim Vector(100, 2) As String
        Dim Auxiliar(100, 4) As String
        
        Erase Auxiliar
        Erase Vector
        Renglon = 0
        
        Producto = ZZCodigo
        
        Rem If ZZCodigo = "PT-03000-058" Or ZZCodigo = "PT-05041-100" Then Stop
        
        Vector(1, 1) = Producto
        Vector(1, 2) = "1"
        Costo = 0
        Lugar = 1
        Cicla = 0
        
        Do
            Cicla = Cicla + 1
            If Vector(Cicla, 1) <> "" Then
        
                Entra = "S"
                
                spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                
                If rstComposicion.RecordCount > 0 Then
                With rstComposicion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Entra = "N"
                            
                            Tipo = rstComposicion!Tipo
                            Articulo1 = rstComposicion!Articulo1
                            Articulo2 = rstComposicion!Articulo2
                            Cantidad = rstComposicion!Cantidad
                            
                            Rem If Left$(Articulo1, 2) = "DW" Then
                            Rem     Tipo = "T"
                            Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                            Rem End If
                            
                            Select Case Tipo
                                Case "T"
                                    If Producto <> Articulo2 Then
                                        Lugar = Lugar + 1
                                        If Lugar > 99 Then
                                            Exit Do
                                        End If
                                        Vector(Lugar, 1) = Articulo2
                                        Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                    End If
                                Case "M"
                                    Renglon = Renglon + 1
                                    If Renglon > 99 Then
                                        Exit Do
                                    End If
                                    Auxiliar(Renglon, 1) = Articulo1
                                    Auxiliar(Renglon, 2) = Str$(Cantidad)
                                    Auxiliar(Renglon, 3) = Vector(Cicla, 2)
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
                
                Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
                Rem     Renglon = Renglon + 1
                Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
                Rem     Auxiliar(Renglon, 2) = "1"
                Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                Rem End If
                
                    Else
                    
                Exit Do
                
            End If
            If Renglon > 99 Then
                Exit Do
            End If
            If Lugar > 99 Then
                Exit Do
            End If
            
        Loop
        
        Suma = 0
        SumaII = 0
                        
        For Da = 1 To Renglon
            Articulo = Auxiliar(Da, 1)
            Cantidad = Val(Auxiliar(Da, 2))
            XVector = Auxiliar(Da, 3)
            
            ZZSedronar = 0
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZSedronar = IIf(IsNull(rstArticulo!Sedronar), "0", rstArticulo!Sedronar)
                rstArticulo.Close
            End If
            
            ZZCantidad = (Cantidad * Val(XVector))
            
            Suma = Suma + ZZCantidad
            If ZZSedronar = 1 Then
                SumaII = SumaII + ZZCantidad
            End If
            
        Next Da
        
        If SumaII <> 0 Then
        
            ZZPorce = SumaII / (Suma / 100)
            Call Redondeo(ZZPorce)
            
            If ZZPorce >= 20 Then
            
                ZZMarca = "N"
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + Producto + "'"
                ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + "20120101" + "'"
                ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + "20141231" + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                    With rstEstadistica
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                
                                ZZMarca = "S"
                                
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstEstadistica.Close
                End If
            
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "PorceSedro = " + "'" + Str$(ZZPorce) + "',"
                ZSql = ZSql & "MarcaSedro = " + "'" + ZZMarca + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + Producto + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
    Next ZZCiclo
    
    
Stop
    
    
End Sub









Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Str$(Cantidad)
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
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
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = "1"
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Val(Auxiliar(Da, 2))
        XVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub





Private Sub HojaTecnica_dblclick()
    
    If Trim(HojaTecnica.Text) = "" Then
            
        m$ = "No existe hoja tecnica"
        G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
    
            Else
    
        ZEstadoHoja = ""
        spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZEstadoHoja = IIf(IsNull(rstTerminado!EstadoHoja), "", rstTerminado!EstadoHoja)
            rstTerminado.Close
        End If
        
        If ZEstadoHoja = "N" Then
        
            T$ = "Hojas Tecnicas"
            m$ = "Se ha modificado las especificaciones del productio y no se verificado la hoja tecnica" + Chr$(13) + "Desea ratificarla "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                WClaveLiberaHoja.Text = ""
                PantaLiberaHoja.Visible = True
                WClaveLiberaHoja.SetFocus
            End If
        
                Else
        
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
                         
            ZZRuta = "W:\" + HojaTecnica.Text
            ZZEstado = Dir(ZZRuta)
            If ZZEstado <> "" Then
                RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
            End If
            
        End If
    End If

End Sub

Sub Imprime_Descripcion()

    spEnvase = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri1.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri1.Caption = ""
    End If
    
    spEnvase = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri2.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri2.Caption = ""
    End If
    
    spEnvase = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri3.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri3.Caption = ""
    End If
    
    spEnvase = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri4.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri4.Caption = ""
    End If
    
    spEnvase = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri5.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri5.Caption = ""
    End If
    
    spEnvase = "ConsultaEnvases " + "'" + Envase6.Text + "'"
    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvase.RecordCount > 0 Then
        Descri6.Caption = rstEnvase!Descripcion
        rstEnvase.Close
            Else
        Descri6.Caption = ""
    End If
    
    spLineas = "ConsultaLinea " + "'" + Linea.Text + "'"
    Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineas.RecordCount > 0 Then
        DescriLinea.Caption = rstLineas!Nombre
        rstLineas.Close
    End If
    
    If Val(Efluentes.Text) = 0 Then
        Efluentes.Text = "0"
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Efluentes"
    ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + Efluentes.Text + "'"
    spEfluentes = ZSql
    Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
    If rstEfluentes.RecordCount > 0 Then
        DesEfluentes.Text = rstEfluentes!Descripcion
        rstEfluentes.Close
        Envase1.SetFocus
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
    If Val(Fabrica.Text) = 0 Then
        Fabrica.Text = "0"
    End If
    If Val(FabricaII.Text) = 0 Then
        FabricaII.Text = "0"
    End If
    If Val(FabricaIII.Text) = 0 Then
        FabricaIII.Text = "0"
    End If
    If Val(Proceso.Caption) = 0 Then
        Proceso.Caption = "0"
    End If
    If Val(Linea.Text) = 0 Then
        Linea.Text = "0"
    End If
    If Val(Envase1.Text) = 0 Then
        Envase1.Text = "0"
    End If
    If Val(Envase2.Text) = 0 Then
        Envase2.Text = "0"
    End If
    If Val(Envase3.Text) = 0 Then
        Envase3.Text = "0"
    End If
    If Val(Envase4.Text) = 0 Then
        Envase4.Text = "0"
    End If
    If Val(Envase5.Text) = 0 Then
        Envase5.Text = "0"
    End If
    If Val(Envase6.Text) = 0 Then
        Envase6.Text = "0"
    End If
    If Val(Efluentes.Text) = 0 Then
        Efluentes.Text = "0"
    End If
End Sub

Sub Format_datos()
    Inicial.Text = Pusing("###,###.##", Inicial.Text)
    Entradas.Text = Pusing("###,###.##", Entradas.Text)
    Salidas.Text = Pusing("###,###.##", Salidas.Text)
    Minimo.Text = Pusing("###,###.##", Minimo.Text)
    Minimo1.Text = Pusing("###,###.##", Minimo1.Text)
    Fabrica.Text = Pusing("###,###.##", Fabrica.Text)
    FabricaII.Text = Pusing("###,###.##", FabricaII.Text)
    FabricaIII.Text = Pusing("###,###.##", FabricaIII.Text)
    Proceso.Caption = Pusing("###,###.##", Proceso.Caption)
    Pedido.Caption = Pusing("###,###.##", Pedido.Caption)
    Stock.Caption = Pusing("###,###.##", Val(Inicial.Text) + Val(Entradas.Text) - Val(Salidas.Text))
End Sub

Private Sub Acepta_Click()

    Desdecodigo.Text = UCase(Desdecodigo.Text)
    HastaCodigo.Text = UCase(HastaCodigo.Text)

    Listado.WindowTitle = "Listado de Rubros"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Terminado.Linea} in " + DesdeLinea.Text + " to " + HastaLinea.Text + " AND {Terminado.Codigo} in " + Chr$(34) + Desdecodigo.Text + Chr$(34) + " to " + Chr$(34) + HastaCodigo.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.Linea, Terminado.Inicial, Terminado.Entradas, Terminado.Salidas, Terminado.Costo, Lineas.Nombre " _
                        + "From " + DSQ + ".dbo.Terminado Terminado, " _
                        + DSQ + ".dbo.Lineas Lineas " _
                        + "Where Terminado.Linea = Lineas.Linea AND Terminado.Codigo >= ' ' AND Terminado.Codigo <= 'ZZ-ZZZZZ-ZZZ' AND Terminado.Linea >= 0 AND Terminado.Linea <= 9999"
    
    Listado.DataFiles(2) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub


Private Sub AvisoErrorII_Click()
    AvisoErrorII.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()

    Select Case Val(Wempresa)
        Case 1, 4, 8
            Rem se puede dar de alta
        Case Else
            m$ = "La empresa no esta habilitada para dar de alta o modificar productos terminados"
            G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
            Exit Sub
    End Select

    spLineas = "ConsultaLinea " + "'" + Linea.Text + "'"
    Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineas.RecordCount > 0 Then
        rstLineas.Close
            Else
        m$ = "Codigo de Linea Invalido"
        G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
        Exit Sub
    End If
    
    If Val(Naciones.Text) <> 0 Or Trim(Clase.Text) <> "" Then
        Impreadi.Text = "S"
    End If
    
    Codigo.Text = UCase(Codigo.Text)
    If Codigo.Text >= "PT-25000-000" And Codigo.Text <= "PT-25999-999" Then
        Rem No controlo nada
            Else
        Conservacion.Text = Trim(Conservacion.Text)
        ConservacionII.Text = Trim(ConservacionII.Text)
        If Len(Conservacion.Text) > 43 Or Len(ConservacionII.Text) > 43 Then
            m$ = "El campo conservacion supera el maximo de caracteres permitido"
            G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
            Exit Sub
        End If
    End If
    
    If Val(Naciones.Text) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Peligroso"
        ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
        spPeligroso = ZSql
        Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
        If rstPeligroso.RecordCount > 0 Then
            rstPeligroso.Close
                Else
            m$ = "Codigo de Nacines Unidas Invalido"
            G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
            Exit Sub
        End If
    End If
    
    If UCase(Trim(Impreadi.Text)) = "S" Then
        If Val(Naciones.Text) = 0 Then
            m$ = "Se ha definido el Producto Terminado como peligroso y no se informo numero de naciones unidas"
            G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
            Exit Sub
        End If
    End If
            
    

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
    
    
    
    WProceso = 0

    If WGraba <> "S" Then
    
        Call Ingresa_clave
        
            Else
            
        Codigo.Text = UCase(Codigo.Text)
        If Codigo.Text <> "" Then
        
            If Val(Wempresa) = 1 Then
                Call Verifica_Msds
            End If
    
            spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
                WPasa = "S"
                rstTerminado.Close
                
                    Else
                    
                Select Case Val(Wempresa)
                    Case 1, 4, 8
                        Rem se puede dar de alta
                    Case Else
                        m$ = "La empresa no esta habilitada para dar de lata productos terminados"
                        G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
                        Exit Sub
                End Select
                    
                ZGraba = "N"
                If ZCampo1 = "S" And ZCampo2 = "S" And ZCampo3 = "S" And ZCampo4 = "S" And ZCampo5 = "S" Then
                    If ZCampo6 = "S" And ZCampo7 = "S" And ZCampo8 = "S" And ZCampo9 = "S" And ZCampo10 = "S" And ZCampo11 = "S" Then
                        If ZCampo12 = "S" And ZCampo13 = "S" And ZCampo14 = "S" And ZCampo15 = "S" And ZCampo16 = "S" And ZCampo17 = "S" And ZCampo18 = "S" Then
                            If ZCampo19 = "S" And ZCampo20 = "S" And ZCampo21 = "S" And ZCampo22 = "S" And ZCampo23 = "S" And ZCampo24 = "S" And ZCampo25 = "S" Then
                                If ZCampo26 = "S" And ZCampo27 = "S" And ZCampo28 = "S" And ZCampo29 = "S" And ZCampo30 = "S" And ZCampo31 = "S" And ZCampo32 = "S" Then
                                    ZGraba = "S"
                                End If
                            End If
                        End If
                    End If
                End If
                If ZGraba = "N" Then
                    m$ = "No se puede dar de alta al no haber confirmado la totalidad de los campos"
                    G% = MsgBox(m$, 0, "Ingreso de Producto Terminado")
                    Exit Sub
                End If
                    
                WPasa = "N"
                
            End If
            
            If UCase(Left$(Codigo.Text, 2)) = "RE" Or UCase(Left$(Codigo.Text, 2)) = "NK" Then
            
                ZZPt = "PT" + Right$(Codigo.Text, 10)
                spTerminado = "ConsultaTerminado " + "'" + ZZPt + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Impreadi.Text = ""
                    Clase.Text = ""
                    Secundario.Text = ""
                    Riesgo.Text = ""
                    Intervencion.Text = ""
                    Naciones.Text = ""
                    Embalaje.Text = ""
                    Impreadi.Text = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
                    Clase.Text = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                    Secundario.Text = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
                    Riesgo.Text = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
                    Intervencion.Text = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                    Naciones.Text = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                    Embalaje.Text = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                    rstTerminado.Close
                End If
                
            End If
    
    
            Rem If UCase(Left$(Codigo.Text, 2)) = "NW" Then
            Rem
            Rem     ZZPt = "DW" + Right$(Codigo.Text, 10)
            Rem     spTerminado = "ConsultaTerminado " + "'" + ZZPt + "'"
            Rem     Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstTerminado.RecordCount > 0 Then
            Rem         Impreadi.Text = ""
            Rem         Clase.Text = ""
            Rem         Secundario.Text = ""
            Rem         Riesgo.Text = ""
            Rem         Intervencion.Text = ""
            Rem         Naciones.Text = ""
            Rem         Embalaje.Text = ""
            Rem         Impreadi.Text = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
            Rem         Clase.Text = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
            Rem rem         Secundario.Text = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
            Rem         Riesgo.Text = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
            Rem         Intervencion.Text = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            Rem         Naciones.Text = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            Rem         Embalaje.Text = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            Rem         rstTerminado.Close
            Rem     End If
            Rem
            Rem End If
    
    
            Call Verifica_datos
            WCodigo = Codigo.Text
            WDescripcion = Descripcion.Text
            WDescriEtiqueta = DescriEtiqueta.Text
            WLinea = Linea.Text
            WUnidad = Unidad.Text
            WInicial = Inicial.Text
            WEntradas = Entradas.Text
            WSalidas = Salidas.Text
            WMinimo = Minimo.Text
            WMinimo1 = Minimo1.Text
            WFabrica = Fabrica.Text
            WFabricaII = FabricaII.Text
            WFabricaIII = FabricaIII.Text
            If Val(FabricaII.Text) <> 0 And Val(FabricaIII.Text) <> 0 Then
                WLoteAutorizado = Trim(FabricaII.Text) + " a " + Trim(FabricaIII.Text)
                    Else
                WLoteAutorizado = ""
            End If
            WDeposito = Deposito.Text
            WPedido = ""
            WEnvase1 = Envase1.Text
            WEnvase2 = Envase2.Text
            WEnvase3 = Envase3.Text
            WEnvase4 = Envase4.Text
            WEnvase5 = Envase5.Text
            WEnvase6 = Envase6.Text
            WProceso = Val(Proceso.Caption)
            WCosto = ""
            WFactor = ""
            WDate = Date$
            WImpreadi = Impreadi.Text
            WIntervencion = Intervencion.Text
            WClase = Clase.Text
            WSecundario = Secundario.Text
            WRiesgo = Riesgo.Text
            WNaciones = Naciones.Text
            WEmbalaje = Embalaje.Text
            WControla = Str$(Controla.ListIndex)
            WSedronar = Str$(Sedronar.ListIndex)
            WCodSedronar = CodSedronar.Text
            WImpreVto = Str$(ImpreVto.ListIndex)
            WMarca = Trim(Str$(Marca.ListIndex))
            WEscrito = Str$(Escrito.ListIndex)
            WObservaciones = Observaciones.Text
            WTipoeti = TipoEti.Text
            WConservacion = Conservacion.Text
            WConservacionII = ConservacionII.Text
            WVida = Vida.Text
            WSeguridad = Seguridad.Text
            
            WVersion = Version.Text
            WVersionI = VersionI.Text
            WVersionII = VersionII.Text
            
            WFechaVersion = FechaVersion.Text
            WFechaVersionI = FechaVersionI.Text
            WFechaVersionII = FechaVersionII.Text
            
            WEstado = Estado.Text
            WEstadoI = EstadoI.Text
            WEstadoII = EstadoII.Text
            
            WObserva = Observa.Text
            WObservaI = ObservaI.Text
            WObservaII = ObservaII.Text
            
            WMetodo = Metodo.Text
            WEfluentes = Efluentes.Text
            WRestriccion = Restriccion.Value
            
            If WPasa = "N" Then
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO Terminado ("
                ZSql = ZSql & "Codigo ,"
                ZSql = ZSql & "Descripcion ,"
                ZSql = ZSql & "DescriEtiqueta ,"
                ZSql = ZSql & "Linea ,"
                ZSql = ZSql & "Unidad ,"
                ZSql = ZSql & "Inicial ,"
                ZSql = ZSql & "Entradas ,"
                ZSql = ZSql & "Salidas ,"
                ZSql = ZSql & "Minimo ,"
                ZSql = ZSql & "Minimo1 ,"
                ZSql = ZSql & "Deposito ,"
                ZSql = ZSql & "Pedido ,"
                ZSql = ZSql & "Envase1 ,"
                ZSql = ZSql & "Envase2 ,"
                ZSql = ZSql & "Envase3 ,"
                ZSql = ZSql & "Envase4 ,"
                ZSql = ZSql & "Envase5 ,"
                ZSql = ZSql & "Envase6 ,"
                ZSql = ZSql & "Proceso ,"
                ZSql = ZSql & "Costo ,"
                ZSql = ZSql & "Factor ,"
                ZSql = ZSql & "WDate ,"
                ZSql = ZSql & "ImpreAdi ,"
                ZSql = ZSql & "Clase ,"
                ZSql = ZSql & "Secundario ,"
                ZSql = ZSql & "Riesgo ,"
                ZSql = ZSql & "Intervencion ,"
                ZSql = ZSql & "Naciones ,"
                ZSql = ZSql & "Embalaje ,"
                ZSql = ZSql & "Controla ,"
                ZSql = ZSql & "Sedronar ,"
                ZSql = ZSql & "CodSedronar ,"
                ZSql = ZSql & "ImpreVto ,"
                ZSql = ZSql & "Marca ,"
                ZSql = ZSql & "Observaciones ,"
                ZSql = ZSql & "TipoEti ,"
                ZSql = ZSql & "Escrito ,"
                ZSql = ZSql & "Fabrica ,"
                ZSql = ZSql & "FabricaII ,"
                ZSql = ZSql & "FabricaIII ,"
                ZSql = ZSql & "LoteAutorizado ,"
                ZSql = ZSql & "Conservacion ,"
                ZSql = ZSql & "ConservacionII ,"
                ZSql = ZSql & "Vida ,"
                ZSql = ZSql & "Seguridad ,"
                ZSql = ZSql & "Version ,"
                ZSql = ZSql & "VersionI ,"
                ZSql = ZSql & "VersionII ,"
                ZSql = ZSql & "FechaVersion ,"
                ZSql = ZSql & "FechaVersionI ,"
                ZSql = ZSql & "FechaVersionII ,"
                ZSql = ZSql & "Estado ,"
                ZSql = ZSql & "EstadoI ,"
                ZSql = ZSql & "EstadoII ,"
                ZSql = ZSql & "Observa ,"
                ZSql = ZSql & "ObservaI ,"
                ZSql = ZSql & "ObservaII ,"
                ZSql = ZSql & "DescripcionIngles ,"
                ZSql = ZSql & "DescriEtiquetaIngles ,"
                ZSql = ZSql & "ConservacionIngles ,"
                ZSql = ZSql & "ConservacionIIIngles ,"
                ZSql = ZSql & "Metodo ,"
                ZSql = ZSql & "Efluentes )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WCodigo + "',"
                ZSql = ZSql & "'" + WDescripcion + "',"
                ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                ZSql = ZSql & "'" + WLinea + "',"
                ZSql = ZSql & "'" + WUnidad + "',"
                ZSql = ZSql & "'" + WInicial + "',"
                ZSql = ZSql & "'" + WEntradas + "',"
                ZSql = ZSql & "'" + WSalidas + "',"
                ZSql = ZSql & "'" + WMinimo + "',"
                ZSql = ZSql & "'" + WMinimo1 + "',"
                ZSql = ZSql & "'" + WDeposito + "',"
                ZSql = ZSql & "'" + WPedido + "',"
                ZSql = ZSql & "'" + WEnvase1 + "',"
                ZSql = ZSql & "'" + WEnvase2 + "',"
                ZSql = ZSql & "'" + WEnvase3 + "',"
                ZSql = ZSql & "'" + WEnvase4 + "',"
                ZSql = ZSql & "'" + WEnvase5 + "',"
                ZSql = ZSql & "'" + WEnvase6 + "',"
                ZSql = ZSql & "'" + WProceso + "',"
                ZSql = ZSql & "'" + WCosto + "',"
                ZSql = ZSql & "'" + WFactor + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WImpreadi + "',"
                ZSql = ZSql & "'" + WClase + "',"
                ZSql = ZSql & "'" + WSecundario + "',"
                ZSql = ZSql & "'" + WRiesgo + "',"
                ZSql = ZSql & "'" + WIntervencion + "',"
                ZSql = ZSql & "'" + WNaciones + "',"
                ZSql = ZSql & "'" + WEmbalaje + "',"
                ZSql = ZSql & "'" + WControla + "',"
                ZSql = ZSql & "'" + WSedronar + "',"
                ZSql = ZSql & "'" + WCodSedronar + "',"
                ZSql = ZSql & "'" + WImpreVto + "',"
                ZSql = ZSql & "'" + WMarca + "',"
                ZSql = ZSql & "'" + WObservaciones + "',"
                ZSql = ZSql & "'" + WTipoeti + "',"
                ZSql = ZSql & "'" + WEscrito + "',"
                ZSql = ZSql & "'" + WFabrica + "',"
                ZSql = ZSql & "'" + WFabricaII + "',"
                ZSql = ZSql & "'" + WFabricaIII + "',"
                ZSql = ZSql & "'" + WLoteAutorizado + "',"
                ZSql = ZSql & "'" + WConservacion + "',"
                ZSql = ZSql & "'" + WConservacionII + "',"
                ZSql = ZSql & "'" + WVida + "',"
                ZSql = ZSql & "'" + WSeguridad + "',"
                ZSql = ZSql & "'" + WVersion + "',"
                ZSql = ZSql & "'" + WVersionI + "',"
                ZSql = ZSql & "'" + WVersionII + "',"
                ZSql = ZSql & "'" + WFechaVersion + "',"
                ZSql = ZSql & "'" + WFechaVersionI + "',"
                ZSql = ZSql & "'" + WFechaVersionII + "',"
                ZSql = ZSql & "'" + WEstado + "',"
                ZSql = ZSql & "'" + WEstadoI + "',"
                ZSql = ZSql & "'" + WEstadoII + "',"
                ZSql = ZSql & "'" + WObserva + "',"
                ZSql = ZSql & "'" + WObservaI + "',"
                ZSql = ZSql & "'" + WObservaII + "',"
                ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                ZSql = ZSql & "'" + WMetodo + "',"
                ZSql = ZSql & "'" + WEfluentes + "')"
      
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                        Else
                        
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "Codigo = " + "'" + WCodigo + "',"
                ZSql = ZSql & "Descripcion = " + "'" + WDescripcion + "',"
                ZSql = ZSql & "DescriEtiqueta = " + "'" + WDescriEtiqueta + "',"
                ZSql = ZSql & "Linea = " + "'" + WLinea + "',"
                ZSql = ZSql & "Unidad = " + "'" + WUnidad + "',"
                ZSql = ZSql & "Inicial = " + "'" + WInicial + "',"
                ZSql = ZSql & "Entradas = " + "'" + WEntradas + "',"
                ZSql = ZSql & "Salidas = " + "'" + WSalidas + "',"
                ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
                ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
                ZSql = ZSql & "Deposito = " + "'" + WDeposito + "',"
                ZSql = ZSql & "Pedido = " + "'" + WPedido + "',"
                ZSql = ZSql & "Envase1 = " + "'" + WEnvase1 + "',"
                ZSql = ZSql & "Envase2 = " + "'" + WEnvase2 + "',"
                ZSql = ZSql & "Envase3 = " + "'" + WEnvase3 + "',"
                ZSql = ZSql & "Envase4 = " + "'" + WEnvase4 + "',"
                ZSql = ZSql & "Envase5 = " + "'" + WEnvase5 + "',"
                ZSql = ZSql & "Envase6 = " + "'" + WEnvase6 + "',"
                ZSql = ZSql & "Proceso = " + "'" + WProceso + "',"
                ZSql = ZSql & "Costo = " + "'" + WCosto + "',"
                ZSql = ZSql & "Factor = " + "'" + WFactor + "',"
                ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                ZSql = ZSql & "ImpreAdi = " + "'" + WImpreadi + "',"
                ZSql = ZSql & "Clase = " + "'" + WClase + "',"
                ZSql = ZSql & "Secundario = " + "'" + WSecundario + "',"
                ZSql = ZSql & "Riesgo = " + "'" + WRiesgo + "',"
                ZSql = ZSql & "Intervencion = " + "'" + WIntervencion + "',"
                ZSql = ZSql & "Naciones = " + "'" + WNaciones + "',"
                ZSql = ZSql & "Embalaje = " + "'" + WEmbalaje + "',"
                ZSql = ZSql & "Controla = " + "'" + WControla + "',"
                ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
                ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
                ZSql = ZSql & "ImpreVto = " + "'" + WImpreVto + "',"
                ZSql = ZSql & "Marca = " + "'" + WMarca + "',"
                ZSql = ZSql & "Observaciones = " + "'" + WObservaciones + "',"
                ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
                ZSql = ZSql & "Escrito = " + "'" + WEscrito + "',"
                ZSql = ZSql & "Fabrica = " + "'" + WFabrica + "',"
                ZSql = ZSql & "FabricaII = " + "'" + WFabricaII + "',"
                ZSql = ZSql & "FabricaIII = " + "'" + WFabricaIII + "',"
                ZSql = ZSql & "LoteAutorizado = " + "'" + WLoteAutorizado + "',"
                ZSql = ZSql & "Conservacion = " + "'" + WConservacion + "',"
                ZSql = ZSql & "ConservacionII = " + "'" + WConservacionII + "',"
                ZSql = ZSql & "Vida = " + "'" + WVida + "',"
                ZSql = ZSql & "Seguridad = " + "'" + WSeguridad + "',"
                ZSql = ZSql & "Version = " + "'" + WVersion + "',"
                ZSql = ZSql & "VersionI = " + "'" + WVersionI + "',"
                ZSql = ZSql & "VersionII = " + "'" + WVersionII + "',"
                ZSql = ZSql & "FechaVersion = " + "'" + WFechaVersion + "',"
                ZSql = ZSql & "FechaVersionI = " + "'" + WFechaVersionI + "',"
                ZSql = ZSql & "FechaVersionII = " + "'" + WFechaVersionII + "',"
                ZSql = ZSql & "Estado = " + "'" + WEstado + "',"
                ZSql = ZSql & "EstadoI = " + "'" + WEstadoI + "',"
                ZSql = ZSql & "EstadoII = " + "'" + WEstadoII + "',"
                ZSql = ZSql & "Observa = " + "'" + WObserva + "',"
                ZSql = ZSql & "ObservaI = " + "'" + WObservaI + "',"
                ZSql = ZSql & "ObservaII = " + "'" + WObservaII + "',"
                ZSql = ZSql & "DescripcionIngles = " + "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql & "DescriEtiquetaIngles = " + "'" + DescriEtiquetaIngles.Text + "',"
                ZSql = ZSql & "ConservacionIngles = " + "'" + ConservacionIngles.Text + "',"
                ZSql = ZSql & "ConservacionIIIngles = " + "'" + ConservacionIIIngles.Text + "',"
                ZSql = ZSql & "Metodo = " + "'" + WMetodo + "',"
                ZSql = ZSql & "Efluentes = " + "'" + WEfluentes + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Terminado SET "
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
            ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
            ZSql = ZSql & "Carga = " + "'" + Str$(Carga.ListIndex) + "',"
            ZSql = ZSql & "EstadoProducto = " + "'" + Str$(EstadoProducto.ListIndex) + "',"
            ZSql = ZSql & "ListaProducto = " + "'" + Str$(ListaProducto.ListIndex) + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZZImpreAdi = WImpreadi
            ZZClase = WClase
            ZZSecundario = WSecundario
            ZZRiesgo = Riesgo
            ZZIntervencion = WIntervencion
            ZZNaciones = WNaciones
            ZZEmbalaje = WEmbalaje
            ZZDescriOnu = Caracteristicas.Text
            
            
            
            Rem da de alta el nk
            
            WNk = "NK" + Right$(Codigo.Text, 10)
        
            spTerminado = "ConsultaTerminado " + "'" + WNk + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount = 0 Then
        
                rstTerminado.Close
        
                WCodigo = WNk
                WDescripcion = Descripcion.Text
                WDescriEtiqueta = DescriEtiqueta.Text
                WLinea = Linea.Text
                WUnidad = Unidad.Text
                WInicial = ""
                WEntradas = ""
                WSalidas = ""
                WMinimo = ""
                WMinimo1 = ""
                WDeposito = ""
                WPedido = ""
                WEnvase1 = Envase1.Text
                WEnvase2 = Envase2.Text
                WEnvase3 = Envase3.Text
                WEnvase4 = Envase4.Text
                WEnvase5 = Envase5.Text
                WEnvase6 = Envase6.Text
                WProceso = ""
                WCosto = ""
                WFactor = ""
                WDate = Date$
                WImpreadi = ""
                WIntervencion = ""
                WClase = ""
                WSecundario = ""
                WRiesgo = ""
                WNaciones = ""
                WEmbalaje = ""
                WVersion = ""
                WFechaVersion = "  /  /    "
                WControla = "0"
                WCodSedronar = ""
                WSedronar = "0"
                WImpreVto = "0"
                WMarca = "0"
                WObservaciones = ""
                WEscrito = Str$(Escrito.ListIndex)
                WFabrica = ""
                WFabricaII = ""
                WFabricaIII = ""
                WLoteAutorizado = ""
                WConservacion = ""
                WConservacionII = ""
                WVida = ""
                WSeguridad = ""
                
                WVersion = ""
                WVersionI = ""
                WVersionII = ""
            
                WFechaVersion = "  /  /    "
                WFechaVersionI = "  /  /    "
                WFechaVersionII = "  /  /    "
            
                WEstado = ""
                WEstadoI = ""
                WEstadoII = ""
            
                WObserva = ""
                WObservaI = ""
                WObservaII = ""
            
                WMetodo = ""
                WEfluentes = "0"
                
                ZSql = ""
                ZSql = ZSql & "INSERT INTO Terminado ("
                ZSql = ZSql & "Codigo ,"
                ZSql = ZSql & "Descripcion ,"
                ZSql = ZSql & "DescriEtiqueta ,"
                ZSql = ZSql & "Linea ,"
                ZSql = ZSql & "Unidad ,"
                ZSql = ZSql & "Inicial ,"
                ZSql = ZSql & "Entradas ,"
                ZSql = ZSql & "Salidas ,"
                ZSql = ZSql & "Minimo ,"
                ZSql = ZSql & "Minimo1 ,"
                ZSql = ZSql & "Deposito ,"
                ZSql = ZSql & "Pedido ,"
                ZSql = ZSql & "Envase1 ,"
                ZSql = ZSql & "Envase2 ,"
                ZSql = ZSql & "Envase3 ,"
                ZSql = ZSql & "Envase4 ,"
                ZSql = ZSql & "Envase5 ,"
                ZSql = ZSql & "Envase6 ,"
                ZSql = ZSql & "Proceso ,"
                ZSql = ZSql & "Costo ,"
                ZSql = ZSql & "Factor ,"
                ZSql = ZSql & "WDate ,"
                ZSql = ZSql & "ImpreAdi ,"
                ZSql = ZSql & "Clase ,"
                ZSql = ZSql & "Secundario ,"
                ZSql = ZSql & "Riesgo ,"
                ZSql = ZSql & "Intervencion ,"
                ZSql = ZSql & "Naciones ,"
                ZSql = ZSql & "Embalaje ,"
                ZSql = ZSql & "Controla ,"
                ZSql = ZSql & "Sedronar ,"
                ZSql = ZSql & "CodSedronar ,"
                ZSql = ZSql & "ImpreVto ,"
                ZSql = ZSql & "Marca ,"
                ZSql = ZSql & "Observaciones ,"
                ZSql = ZSql & "TipoEti ,"
                ZSql = ZSql & "Escrito ,"
                ZSql = ZSql & "Fabrica ,"
                ZSql = ZSql & "FabricaII ,"
                ZSql = ZSql & "FabricaIII ,"
                ZSql = ZSql & "LoteAutorizado ,"
                ZSql = ZSql & "Conservacion ,"
                ZSql = ZSql & "ConservacionII ,"
                ZSql = ZSql & "Vida ,"
                ZSql = ZSql & "Seguridad ,"
                ZSql = ZSql & "Version ,"
                ZSql = ZSql & "VersionI ,"
                ZSql = ZSql & "VersionII ,"
                ZSql = ZSql & "FechaVersion ,"
                ZSql = ZSql & "FechaVersionI ,"
                ZSql = ZSql & "FechaVersionII ,"
                ZSql = ZSql & "Estado ,"
                ZSql = ZSql & "EstadoI ,"
                ZSql = ZSql & "EstadoII ,"
                ZSql = ZSql & "Observa ,"
                ZSql = ZSql & "ObservaI ,"
                ZSql = ZSql & "ObservaII ,"
                ZSql = ZSql & "DescripcionIngles ,"
                ZSql = ZSql & "DescriEtiquetaIngles ,"
                ZSql = ZSql & "ConservacionIngles ,"
                ZSql = ZSql & "ConservacionIIIngles ,"
                ZSql = ZSql & "Metodo ,"
                ZSql = ZSql & "Efluentes )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WCodigo + "',"
                ZSql = ZSql & "'" + WDescripcion + "',"
                ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                ZSql = ZSql & "'" + WLinea + "',"
                ZSql = ZSql & "'" + WUnidad + "',"
                ZSql = ZSql & "'" + WInicial + "',"
                ZSql = ZSql & "'" + WEntradas + "',"
                ZSql = ZSql & "'" + WSalidas + "',"
                ZSql = ZSql & "'" + WMinimo + "',"
                ZSql = ZSql & "'" + WMinimo1 + "',"
                ZSql = ZSql & "'" + WDeposito + "',"
                ZSql = ZSql & "'" + WPedido + "',"
                ZSql = ZSql & "'" + WEnvase1 + "',"
                ZSql = ZSql & "'" + WEnvase2 + "',"
                ZSql = ZSql & "'" + WEnvase3 + "',"
                ZSql = ZSql & "'" + WEnvase4 + "',"
                ZSql = ZSql & "'" + WEnvase5 + "',"
                ZSql = ZSql & "'" + WEnvase6 + "',"
                ZSql = ZSql & "'" + WProceso + "',"
                ZSql = ZSql & "'" + WCosto + "',"
                ZSql = ZSql & "'" + WFactor + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WImpreadi + "',"
                ZSql = ZSql & "'" + WClase + "',"
                ZSql = ZSql & "'" + WSecundario + "',"
                ZSql = ZSql & "'" + WRiesgo + "',"
                ZSql = ZSql & "'" + WIntervencion + "',"
                ZSql = ZSql & "'" + WNaciones + "',"
                ZSql = ZSql & "'" + WEmbalaje + "',"
                ZSql = ZSql & "'" + WControla + "',"
                ZSql = ZSql & "'" + WSedronar + "',"
                ZSql = ZSql & "'" + WCodSedronar + "',"
                ZSql = ZSql & "'" + WImpreVto + "',"
                ZSql = ZSql & "'" + WMarca + "',"
                ZSql = ZSql & "'" + WObservaciones + "',"
                ZSql = ZSql & "'" + WTipoeti + "',"
                ZSql = ZSql & "'" + WEscrito + "',"
                ZSql = ZSql & "'" + WFabrica + "',"
                ZSql = ZSql & "'" + WFabricaII + "',"
                ZSql = ZSql & "'" + WFabricaIII + "',"
                ZSql = ZSql & "'" + WLoteAutorizado + "',"
                ZSql = ZSql & "'" + WConservacion + "',"
                ZSql = ZSql & "'" + WConservacionII + "',"
                ZSql = ZSql & "'" + WVida + "',"
                ZSql = ZSql & "'" + WSeguridad + "',"
                ZSql = ZSql & "'" + WVersion + "',"
                ZSql = ZSql & "'" + WVersionI + "',"
                ZSql = ZSql & "'" + WVersionII + "',"
                ZSql = ZSql & "'" + WFechaVersion + "',"
                ZSql = ZSql & "'" + WFechaVersionI + "',"
                ZSql = ZSql & "'" + WFechaVersionII + "',"
                ZSql = ZSql & "'" + WEstado + "',"
                ZSql = ZSql & "'" + WEstadoI + "',"
                ZSql = ZSql & "'" + WEstadoII + "',"
                ZSql = ZSql & "'" + WObserva + "',"
                ZSql = ZSql & "'" + WObservaI + "',"
                ZSql = ZSql & "'" + WObservaII + "',"
                ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                ZSql = ZSql & "'" + WMetodo + "',"
                ZSql = ZSql & "'" + WEfluentes + "')"
      
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        
            ZSql = ""
            ZSql = ZSql & "UPDATE Terminado SET "
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
            ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
            ZSql = ZSql & "ImpreAdi = " + "'" + ZZImpreAdi + "',"
            ZSql = ZSql & "Clase = " + "'" + ZZClase + "',"
            ZSql = ZSql & "Secundario = " + "'" + ZZSecundario + "',"
            ZSql = ZSql & "Riesgo = " + "'" + ZZRiesgo + "',"
            ZSql = ZSql & "Intervencion = " + "'" + ZZIntervencion + "',"
            ZSql = ZSql & "Naciones = " + "'" + ZZNaciones + "',"
            ZSql = ZSql & "Embalaje = " + "'" + ZZEmbalaje + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + ZZDescriOnu + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WNk + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
            Rem da de alta el Re
        
            WRe = "RE" + Right$(Codigo.Text, 10)
            
            spTerminado = "ConsultaTerminado " + "'" + WRe + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount = 0 Then
        
                rstTerminado.Close
        
                WCodigo = WRe
                WDescripcion = Descripcion.Text
                WDescriEtiqueta = DescriEtiqueta.Text
                WLinea = Linea.Text
                WUnidad = Unidad.Text
                WInicial = ""
                WEntradas = ""
                WSalidas = ""
                WMinimo = ""
                WMinimo1 = ""
                WDeposito = ""
                WPedido = ""
                WEnvase1 = Envase1.Text
                WEnvase2 = Envase2.Text
                WEnvase3 = Envase3.Text
                WEnvase4 = Envase4.Text
                WEnvase5 = Envase5.Text
                WEnvase6 = Envase6.Text
                WProceso = ""
                WCosto = ""
                WFactor = ""
                WDate = Date$
                WImpreadi = ""
                WIntervencion = ""
                WClase = ""
                WSecundario = ""
                WRiesgo = ""
                WNaciones = ""
                WEmbalaje = ""
                WVersion = ""
                WFechaVersion = "  /  /    "
                WControla = "0"
                WSedronar = "0"
                WCodSedronar = ""
                WImpreVto = "0"
                WMarca = "0"
                WObservaciones = ""
                WEscrito = Str$(Escrito.ListIndex)
                
                WFabrica = ""
                WFabricaII = ""
                WFabricaIII = ""
                WLoteAutorizado = ""
                WConservacion = ""
                WConservacionII = ""
                WVida = ""
                WSeguridad = ""
                
                WVersion = ""
                WVersionI = ""
                WVersionII = ""
            
                WFechaVersion = "  /  /    "
                WFechaVersionI = "  /  /    "
                WFechaVersionII = "  /  /    "
            
                WEstado = ""
                WEstadoI = ""
                WEstadoII = ""
            
                WObserva = ""
                WObservaI = ""
                WObservaII = ""
            
                WMetodo = ""
                WEfluentes = "0"
                
                ZSql = ""
                ZSql = ZSql & "INSERT INTO Terminado ("
                ZSql = ZSql & "Codigo ,"
                ZSql = ZSql & "Descripcion ,"
                ZSql = ZSql & "DescriEtiqueta ,"
                ZSql = ZSql & "Linea ,"
                ZSql = ZSql & "Unidad ,"
                ZSql = ZSql & "Inicial ,"
                ZSql = ZSql & "Entradas ,"
                ZSql = ZSql & "Salidas ,"
                ZSql = ZSql & "Minimo ,"
                ZSql = ZSql & "Minimo1 ,"
                ZSql = ZSql & "Deposito ,"
                ZSql = ZSql & "Pedido ,"
                ZSql = ZSql & "Envase1 ,"
                ZSql = ZSql & "Envase2 ,"
                ZSql = ZSql & "Envase3 ,"
                ZSql = ZSql & "Envase4 ,"
                ZSql = ZSql & "Envase5 ,"
                ZSql = ZSql & "Envase6 ,"
                ZSql = ZSql & "Proceso ,"
                ZSql = ZSql & "Costo ,"
                ZSql = ZSql & "Factor ,"
                ZSql = ZSql & "WDate ,"
                ZSql = ZSql & "ImpreAdi ,"
                ZSql = ZSql & "Clase ,"
                ZSql = ZSql & "Secundario ,"
                ZSql = ZSql & "Riesgo ,"
                ZSql = ZSql & "Intervencion ,"
                ZSql = ZSql & "Naciones ,"
                ZSql = ZSql & "Embalaje ,"
                ZSql = ZSql & "Controla ,"
                ZSql = ZSql & "Sedronar ,"
                ZSql = ZSql & "CodSedronar ,"
                ZSql = ZSql & "ImpreVto ,"
                ZSql = ZSql & "Marca ,"
                ZSql = ZSql & "Observaciones ,"
                ZSql = ZSql & "TipoEti ,"
                ZSql = ZSql & "Escrito ,"
                ZSql = ZSql & "Fabrica ,"
                ZSql = ZSql & "FabricaII ,"
                ZSql = ZSql & "FabricaIII ,"
                ZSql = ZSql & "LoteAutorizado ,"
                ZSql = ZSql & "Conservacion ,"
                ZSql = ZSql & "ConservacionII ,"
                ZSql = ZSql & "Vida ,"
                ZSql = ZSql & "Seguridad ,"
                ZSql = ZSql & "Version ,"
                ZSql = ZSql & "VersionI ,"
                ZSql = ZSql & "VersionII ,"
                ZSql = ZSql & "FechaVersion ,"
                ZSql = ZSql & "FechaVersionI ,"
                ZSql = ZSql & "FechaVersionII ,"
                ZSql = ZSql & "Estado ,"
                ZSql = ZSql & "EstadoI ,"
                ZSql = ZSql & "EstadoII ,"
                ZSql = ZSql & "Observa ,"
                ZSql = ZSql & "ObservaI ,"
                ZSql = ZSql & "ObservaII ,"
                ZSql = ZSql & "DescripcionIngles ,"
                ZSql = ZSql & "DescriEtiquetaIngles ,"
                ZSql = ZSql & "ConservacionIngles ,"
                ZSql = ZSql & "ConservacionIIIngles ,"
                ZSql = ZSql & "Metodo ,"
                ZSql = ZSql & "Efluentes )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WCodigo + "',"
                ZSql = ZSql & "'" + WDescripcion + "',"
                ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                ZSql = ZSql & "'" + WLinea + "',"
                ZSql = ZSql & "'" + WUnidad + "',"
                ZSql = ZSql & "'" + WInicial + "',"
                ZSql = ZSql & "'" + WEntradas + "',"
                ZSql = ZSql & "'" + WSalidas + "',"
                ZSql = ZSql & "'" + WMinimo + "',"
                ZSql = ZSql & "'" + WMinimo1 + "',"
                ZSql = ZSql & "'" + WDeposito + "',"
                ZSql = ZSql & "'" + WPedido + "',"
                ZSql = ZSql & "'" + WEnvase1 + "',"
                ZSql = ZSql & "'" + WEnvase2 + "',"
                ZSql = ZSql & "'" + WEnvase3 + "',"
                ZSql = ZSql & "'" + WEnvase4 + "',"
                ZSql = ZSql & "'" + WEnvase5 + "',"
                ZSql = ZSql & "'" + WEnvase6 + "',"
                ZSql = ZSql & "'" + WProceso + "',"
                ZSql = ZSql & "'" + WCosto + "',"
                ZSql = ZSql & "'" + WFactor + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WImpreadi + "',"
                ZSql = ZSql & "'" + WClase + "',"
                ZSql = ZSql & "'" + WSecundario + "',"
                ZSql = ZSql & "'" + WRiesgo + "',"
                ZSql = ZSql & "'" + WIntervencion + "',"
                ZSql = ZSql & "'" + WNaciones + "',"
                ZSql = ZSql & "'" + WEmbalaje + "',"
                ZSql = ZSql & "'" + WControla + "',"
                ZSql = ZSql & "'" + WSedronar + "',"
                ZSql = ZSql & "'" + WCodSedronar + "',"
                ZSql = ZSql & "'" + WImpreVto + "',"
                ZSql = ZSql & "'" + WMarca + "',"
                ZSql = ZSql & "'" + WObservaciones + "',"
                ZSql = ZSql & "'" + WTipoeti + "',"
                ZSql = ZSql & "'" + WEscrito + "',"
                ZSql = ZSql & "'" + WFabrica + "',"
                ZSql = ZSql & "'" + WFabricaII + "',"
                ZSql = ZSql & "'" + WFabricaIII + "',"
                ZSql = ZSql & "'" + WLoteAutorizado + "',"
                ZSql = ZSql & "'" + WConservacion + "',"
                ZSql = ZSql & "'" + WConservacionII + "',"
                ZSql = ZSql & "'" + WVida + "',"
                ZSql = ZSql & "'" + WSeguridad + "',"
                ZSql = ZSql & "'" + WVersion + "',"
                ZSql = ZSql & "'" + WVersionI + "',"
                ZSql = ZSql & "'" + WVersionII + "',"
                ZSql = ZSql & "'" + WFechaVersion + "',"
                ZSql = ZSql & "'" + WFechaVersionI + "',"
                ZSql = ZSql & "'" + WFechaVersionII + "',"
                ZSql = ZSql & "'" + WEstado + "',"
                ZSql = ZSql & "'" + WEstadoI + "',"
                ZSql = ZSql & "'" + WEstadoII + "',"
                ZSql = ZSql & "'" + WObserva + "',"
                ZSql = ZSql & "'" + WObservaI + "',"
                ZSql = ZSql & "'" + WObservaII + "',"
                ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                ZSql = ZSql & "'" + WMetodo + "',"
                ZSql = ZSql & "'" + WEfluentes + "')"
      
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        
            ZSql = ""
            ZSql = ZSql & "UPDATE Terminado SET "
            ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
            ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
            ZSql = ZSql & "ImpreAdi = " + "'" + ZZImpreAdi + "',"
            ZSql = ZSql & "Clase = " + "'" + ZZClase + "',"
            ZSql = ZSql & "Secundario = " + "'" + ZZSecundario + "',"
            ZSql = ZSql & "Riesgo = " + "'" + ZZRiesgo + "',"
            ZSql = ZSql & "Intervencion = " + "'" + ZZIntervencion + "',"
            ZSql = ZSql & "Naciones = " + "'" + ZZNaciones + "',"
            ZSql = ZSql & "Embalaje = " + "'" + ZZEmbalaje + "',"
            ZSql = ZSql & "DescriOnu = " + "'" + ZZDescriOnu + "'"
            ZSql = ZSql & " Where Codigo = " + "'" + WRe + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            
            Rem verifica la alta en todas las empresas
            
            XEmpresa = Wempresa
            Erase CargaEmpresa
        
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
                Case 2
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                Case 3
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
                Case 4
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 5
                    CargaEmpresa(1, 1) = "0005"
                    CargaEmpresa(1, 2) = "Empresa05"
                Case 6
                    CargaEmpresa(1, 1) = "0006"
                    CargaEmpresa(1, 2) = "Empresa06"
                Case 7
                    CargaEmpresa(1, 1) = "0007"
                    CargaEmpresa(1, 2) = "Empresa07"
                Case 8
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 9
                    CargaEmpresa(1, 1) = "0009"
                    CargaEmpresa(1, 2) = "Empresa09"
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                Case 11
                    CargaEmpresa(1, 1) = "0011"
                    CargaEmpresa(1, 2) = "Empresa11"
                Case Else
            End Select
                
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    WCodigo = Codigo.Text
                    WDescripcion = Trim(Descripcion.Text)
                    WDescriEtiqueta = Trim(DescriEtiqueta.Text)
                    WLinea = Linea.Text
                    WUnidad = Unidad.Text
                    WInicial = "0"
                    WEntradas = "0"
                    WSalidas = "0"
                    WMinimo = Minimo.Text
                    WMinimo1 = Minimo1.Text
                    WDeposito = Deposito.Text
                    WPedido = "0"
                    WEnvase1 = Envase1.Text
                    WEnvase2 = Envase2.Text
                    WEnvase3 = Envase3.Text
                    WEnvase4 = Envase4.Text
                    WEnvase5 = Envase5.Text
                    WEnvase6 = Envase6.Text
                    WProceso = "0"
                    WCosto = ""
                    WFactor = ""
                    WDate = Date$
                    WImpreadi = Impreadi.Text
                    WClase = Trim(Clase.Text)
                    WSecundario = Trim(Secundario.Text)
                    WRiesgo = Trim(Riesgo.Text)
                    WIntervencion = Trim(Intervencion.Text)
                    WNaciones = Trim(Naciones.Text)
                    WEmbalaje = Trim(Embalaje.Text)
                    WControla = Str$(Controla.ListIndex)
                    WSedronar = Str$(Sedronar.ListIndex)
                    WCodSedronar = Trim(CodSedronar.Text)
                    WImpreVto = Str$(ImpreVto.ListIndex)
                    WMarca = Trim(Str$(Marca.ListIndex))
                    WObservaciones = Trim(Observaciones.Text)
                    WTipoeti = Trim(TipoEti.Text)
                    WEscrito = Str$(Escrito.ListIndex)
                    WFabrica = Fabrica.Text
                    WFabricaII = FabricaII.Text
                    WFabricaIII = FabricaIII.Text
                    If Val(FabricaII.Text) <> 0 And Val(FabricaIII.Text) <> 0 Then
                        WLoteAutorizado = Trim(FabricaII.Text) + " a " + Trim(FabricaIII.Text)
                            Else
                        WLoteAutorizado = ""
                    End If
                    WConservacion = Trim(Conservacion.Text)
                    WConservacionII = Trim(ConservacionII.Text)
                    WVida = Vida.Text
                    WSeguridad = Trim(Seguridad.Text)
            
                    WVersion = Version.Text
                    WVersionI = VersionI.Text
                    WVersionII = VersionII.Text
            
                    WFechaVersion = FechaVersion.Text
                    WFechaVersionI = FechaVersionI.Text
                    WFechaVersionII = FechaVersionII.Text
            
                    WEstado = Estado.Text
                    WEstadoI = EstadoI.Text
                    WEstadoII = EstadoII.Text
            
                    WObserva = Trim(Observa.Text)
                    WObservaI = Trim(ObservaI.Text)
                    WObservaII = Trim(ObservaII.Text)
            
                    WMetodo = Metodo.Text
                    WEfluentes = Efluentes.Text
                    
                    spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount = 0 Then
            
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO Terminado ("
                        ZSql = ZSql & "Codigo ,"
                        ZSql = ZSql & "Descripcion ,"
                        ZSql = ZSql & "DescriEtiqueta ,"
                        ZSql = ZSql & "Linea ,"
                        ZSql = ZSql & "Unidad ,"
                        ZSql = ZSql & "Inicial ,"
                        ZSql = ZSql & "Entradas ,"
                        ZSql = ZSql & "Salidas ,"
                        ZSql = ZSql & "Minimo ,"
                        ZSql = ZSql & "Minimo1 ,"
                        ZSql = ZSql & "Deposito ,"
                        ZSql = ZSql & "Pedido ,"
                        ZSql = ZSql & "Envase1 ,"
                        ZSql = ZSql & "Envase2 ,"
                        ZSql = ZSql & "Envase3 ,"
                        ZSql = ZSql & "Envase4 ,"
                        ZSql = ZSql & "Envase5 ,"
                        ZSql = ZSql & "Envase6 ,"
                        ZSql = ZSql & "Proceso ,"
                        ZSql = ZSql & "Costo ,"
                        ZSql = ZSql & "Factor ,"
                        ZSql = ZSql & "WDate ,"
                        ZSql = ZSql & "ImpreAdi ,"
                        ZSql = ZSql & "Clase ,"
                        ZSql = ZSql & "Secundario ,"
                        ZSql = ZSql & "Riesgo ,"
                        ZSql = ZSql & "Intervencion ,"
                        ZSql = ZSql & "Naciones ,"
                        ZSql = ZSql & "Embalaje ,"
                        ZSql = ZSql & "Controla ,"
                        ZSql = ZSql & "Sedronar ,"
                        ZSql = ZSql & "CodSedronar ,"
                        ZSql = ZSql & "ImpreVto ,"
                        ZSql = ZSql & "Marca ,"
                        ZSql = ZSql & "Observaciones ,"
                        ZSql = ZSql & "TipoEti ,"
                        ZSql = ZSql & "Escrito ,"
                        ZSql = ZSql & "Fabrica ,"
                        ZSql = ZSql & "FabricaII ,"
                        ZSql = ZSql & "FabricaIII ,"
                        ZSql = ZSql & "LoteAutorizado ,"
                        ZSql = ZSql & "Conservacion ,"
                        ZSql = ZSql & "ConservacionII ,"
                        ZSql = ZSql & "Vida ,"
                        ZSql = ZSql & "Seguridad ,"
                        ZSql = ZSql & "Version ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "FechaVersion ,"
                        ZSql = ZSql & "FechaVersionI ,"
                        ZSql = ZSql & "FechaVersionII ,"
                        ZSql = ZSql & "Estado ,"
                        ZSql = ZSql & "EstadoI ,"
                        ZSql = ZSql & "EstadoII ,"
                        ZSql = ZSql & "Observa ,"
                        ZSql = ZSql & "ObservaI ,"
                        ZSql = ZSql & "ObservaII ,"
                        ZSql = ZSql & "DescripcionIngles ,"
                        ZSql = ZSql & "DescriEtiquetaIngles ,"
                        ZSql = ZSql & "ConservacionIngles ,"
                        ZSql = ZSql & "ConservacionIIIngles ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WCodigo + "',"
                        ZSql = ZSql & "'" + WDescripcion + "',"
                        ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WUnidad + "',"
                        ZSql = ZSql & "'" + WInicial + "',"
                        ZSql = ZSql & "'" + WEntradas + "',"
                        ZSql = ZSql & "'" + WSalidas + "',"
                        ZSql = ZSql & "'" + WMinimo + "',"
                        ZSql = ZSql & "'" + WMinimo1 + "',"
                        ZSql = ZSql & "'" + WDeposito + "',"
                        ZSql = ZSql & "'" + WPedido + "',"
                        ZSql = ZSql & "'" + WEnvase1 + "',"
                        ZSql = ZSql & "'" + WEnvase2 + "',"
                        ZSql = ZSql & "'" + WEnvase3 + "',"
                        ZSql = ZSql & "'" + WEnvase4 + "',"
                        ZSql = ZSql & "'" + WEnvase5 + "',"
                        ZSql = ZSql & "'" + WEnvase6 + "',"
                        ZSql = ZSql & "'" + WProceso + "',"
                        ZSql = ZSql & "'" + WCosto + "',"
                        ZSql = ZSql & "'" + WFactor + "',"
                        ZSql = ZSql & "'" + WDate + "',"
                        ZSql = ZSql & "'" + WImpreadi + "',"
                        ZSql = ZSql & "'" + WClase + "',"
                        ZSql = ZSql & "'" + WSecundario + "',"
                        ZSql = ZSql & "'" + WRiesgo + "',"
                        ZSql = ZSql & "'" + WIntervencion + "',"
                        ZSql = ZSql & "'" + WNaciones + "',"
                        ZSql = ZSql & "'" + WEmbalaje + "',"
                        ZSql = ZSql & "'" + WControla + "',"
                        ZSql = ZSql & "'" + WSedronar + "',"
                        ZSql = ZSql & "'" + WCodSedronar + "',"
                        ZSql = ZSql & "'" + WImpreVto + "',"
                        ZSql = ZSql & "'" + WMarca + "',"
                        ZSql = ZSql & "'" + WObservaciones + "',"
                        ZSql = ZSql & "'" + WTipoeti + "',"
                        ZSql = ZSql & "'" + WEscrito + "',"
                        ZSql = ZSql & "'" + WFabrica + "',"
                        ZSql = ZSql & "'" + WFabricaII + "',"
                        ZSql = ZSql & "'" + WFabricaIII + "',"
                        ZSql = ZSql & "'" + WLoteAutorizado + "',"
                        ZSql = ZSql & "'" + WConservacion + "',"
                        ZSql = ZSql & "'" + WConservacionII + "',"
                        ZSql = ZSql & "'" + WVida + "',"
                        ZSql = ZSql & "'" + WSeguridad + "',"
                        ZSql = ZSql & "'" + WVersion + "',"
                        ZSql = ZSql & "'" + WVersionI + "',"
                        ZSql = ZSql & "'" + WVersionII + "',"
                        ZSql = ZSql & "'" + WFechaVersion + "',"
                        ZSql = ZSql & "'" + WFechaVersionI + "',"
                        ZSql = ZSql & "'" + WFechaVersionII + "',"
                        ZSql = ZSql & "'" + WEstado + "',"
                        ZSql = ZSql & "'" + WEstadoI + "',"
                        ZSql = ZSql & "'" + WEstadoII + "',"
                        ZSql = ZSql & "'" + WObserva + "',"
                        ZSql = ZSql & "'" + WObservaI + "',"
                        ZSql = ZSql & "'" + WObservaII + "',"
                        ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                        ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                        ZSql = ZSql & "'" + WMetodo + "',"
                        ZSql = ZSql & "'" + WEfluentes + "')"
      
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                                Else
                        
                        ZSql = ""
                        ZSql = ZSql & "UPDATE Terminado SET "
                        ZSql = ZSql & "Codigo = " + "'" + WCodigo + "',"
                        ZSql = ZSql & "Descripcion = " + "'" + WDescripcion + "',"
                        ZSql = ZSql & "DescriEtiqueta = " + "'" + WDescriEtiqueta + "',"
                        ZSql = ZSql & "Linea = " + "'" + WLinea + "',"
                        ZSql = ZSql & "Unidad = " + "'" + WUnidad + "',"
                        ZSql = ZSql & "Minimo = " + "'" + WMinimo + "',"
                        ZSql = ZSql & "Minimo1 = " + "'" + WMinimo1 + "',"
                        ZSql = ZSql & "Deposito = " + "'" + WDeposito + "',"
                        ZSql = ZSql & "Envase1 = " + "'" + WEnvase1 + "',"
                        ZSql = ZSql & "Envase2 = " + "'" + WEnvase2 + "',"
                        ZSql = ZSql & "Envase3 = " + "'" + WEnvase3 + "',"
                        ZSql = ZSql & "Envase4 = " + "'" + WEnvase4 + "',"
                        ZSql = ZSql & "Envase5 = " + "'" + WEnvase5 + "',"
                        ZSql = ZSql & "Envase6 = " + "'" + WEnvase6 + "',"
                        ZSql = ZSql & "Costo = " + "'" + WCosto + "',"
                        ZSql = ZSql & "Factor = " + "'" + WFactor + "',"
                        ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                        ZSql = ZSql & "ImpreAdi = " + "'" + WImpreadi + "',"
                        ZSql = ZSql & "Clase = " + "'" + WClase + "',"
                        ZSql = ZSql & "Secundario = " + "'" + WSecundario + "',"
                        ZSql = ZSql & "Riesgo = " + "'" + WRiesgo + "',"
                        ZSql = ZSql & "Intervencion = " + "'" + WIntervencion + "',"
                        ZSql = ZSql & "Naciones = " + "'" + WNaciones + "',"
                        ZSql = ZSql & "Embalaje = " + "'" + WEmbalaje + "',"
                        ZSql = ZSql & "Controla = " + "'" + WControla + "',"
                        ZSql = ZSql & "Sedronar = " + "'" + WSedronar + "',"
                        ZSql = ZSql & "CodSedronar = " + "'" + WCodSedronar + "',"
                        ZSql = ZSql & "ImpreVto = " + "'" + WImpreVto + "',"
                        ZSql = ZSql & "Marca = " + "'" + WMarca + "',"
                        ZSql = ZSql & "Observaciones = " + "'" + WObservaciones + "',"
                        ZSql = ZSql & "TipoEti = " + "'" + WTipoeti + "',"
                        ZSql = ZSql & "Escrito = " + "'" + WEscrito + "',"
                        ZSql = ZSql & "Fabrica = " + "'" + WFabrica + "',"
                        ZSql = ZSql & "FabricaII = " + "'" + WFabricaII + "',"
                        ZSql = ZSql & "FabricaIII = " + "'" + WFabricaIII + "',"
                        ZSql = ZSql & "LoteAutorizado = " + "'" + WLoteAutorizado + "',"
                        ZSql = ZSql & "Conservacion = " + "'" + WConservacion + "',"
                        ZSql = ZSql & "ConservacionII = " + "'" + WConservacionII + "',"
                        ZSql = ZSql & "Vida = " + "'" + WVida + "',"
                        ZSql = ZSql & "Seguridad = " + "'" + WSeguridad + "',"
                        ZSql = ZSql & "Version = " + "'" + WVersion + "',"
                        ZSql = ZSql & "VersionI = " + "'" + WVersionI + "',"
                        ZSql = ZSql & "VersionII = " + "'" + WVersionII + "',"
                        ZSql = ZSql & "FechaVersion = " + "'" + WFechaVersion + "',"
                        ZSql = ZSql & "FechaVersionI = " + "'" + WFechaVersionI + "',"
                        ZSql = ZSql & "FechaVersionII = " + "'" + WFechaVersionII + "',"
                        ZSql = ZSql & "Estado = " + "'" + WEstado + "',"
                        ZSql = ZSql & "EstadoI = " + "'" + WEstadoI + "',"
                        ZSql = ZSql & "EstadoII = " + "'" + WEstadoII + "',"
                        ZSql = ZSql & "Observa = " + "'" + WObserva + "',"
                        ZSql = ZSql & "ObservaI = " + "'" + WObservaI + "',"
                        ZSql = ZSql & "ObservaII = " + "'" + WObservaII + "',"
                        ZSql = ZSql & "DescripcionIngles = " + "'" + DescripcionIngles.Text + "',"
                        ZSql = ZSql & "DescriEtiquetaIngles = " + "'" + DescriEtiquetaIngles.Text + "',"
                        ZSql = ZSql & "ConservacionIngles = " + "'" + ConservacionIngles.Text + "',"
                        ZSql = ZSql & "ConservacionIIIngles = " + "'" + ConservacionIIIngles.Text + "',"
                        ZSql = ZSql & "Metodo = " + "'" + WMetodo + "',"
                        ZSql = ZSql & "Efluentes = " + "'" + WEfluentes + "'"
                        ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                    
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
                    ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
                    ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "',"
                    ZSql = ZSql & "Carga = " + "'" + Str$(Carga.ListIndex) + "',"
                    ZSql = ZSql & "EstadoProducto = " + "'" + Str$(EstadoProducto.ListIndex) + "',"
                    ZSql = ZSql & "ListaProducto = " + "'" + Str$(ListaProducto.ListIndex) + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                        
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Rem da de alta el nk
            
                    WNk = "NK" + Right$(Codigo.Text, 10)
        
                    spTerminado = "ConsultaTerminado " + "'" + WNk + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount = 0 Then
        
                        rstTerminado.Close
        
                        WCodigo = WNk
                        WDescripcion = Descripcion.Text
                        WDescriEtiqueta = DescriEtiqueta.Text
                        WLinea = Linea.Text
                        WUnidad = Unidad.Text
                        WInicial = ""
                        WEntradas = ""
                        WSalidas = ""
                        WMinimo = ""
                        WMinimo1 = ""
                        WDeposito = ""
                        WPedido = ""
                        WEnvase1 = Envase1.Text
                        WEnvase2 = Envase2.Text
                        WEnvase3 = Envase3.Text
                        WEnvase4 = Envase4.Text
                        WEnvase5 = Envase5.Text
                        WEnvase6 = Envase6.Text
                        WProceso = ""
                        WCosto = ""
                        WFactor = ""
                        WDate = Date$
                        WImpreadi = ""
                        WIntervencion = ""
                        WClase = ""
                        WSecundario = ""
                        WRiesgo = ""
                        WNaciones = ""
                        WEmbalaje = ""
                        WVersion = ""
                        WFechaVersion = "  /  /    "
                        WControla = "0"
                        WSedronar = "0"
                        WCodSedronar = ""
                        WImpreVto = "0"
                        WMarca = "0"
                        WObservaciones = ""
                        WEscrito = "0"
                        WFabrica = ""
                        WFabricaII = ""
                        WFabricaIII = ""
                        WLoteAutorizado = ""
                        WConservacion = ""
                        WConservacionII = ""
                        WVida = ""
                        WSeguridad = ""
                        
                        WVersion = ""
                        WVersionI = ""
                        WVersionII = ""
            
                        WFechaVersion = "  /  /    "
                        WFechaVersionI = "  /  /    "
                        WFechaVersionII = "  /  /    "
            
                        WEstado = ""
                        WEstadoI = ""
                        WEstadoII = ""
            
                        WObserva = ""
                        WObservaI = ""
                        WObservaII = ""
            
                        WMetodo = ""
                        WEfluentes = "0"
                
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO Terminado ("
                        ZSql = ZSql & "Codigo ,"
                        ZSql = ZSql & "Descripcion ,"
                        ZSql = ZSql & "DescriEtiqueta ,"
                        ZSql = ZSql & "Linea ,"
                        ZSql = ZSql & "Unidad ,"
                        ZSql = ZSql & "Inicial ,"
                        ZSql = ZSql & "Entradas ,"
                        ZSql = ZSql & "Salidas ,"
                        ZSql = ZSql & "Minimo ,"
                        ZSql = ZSql & "Minimo1 ,"
                        ZSql = ZSql & "Deposito ,"
                        ZSql = ZSql & "Pedido ,"
                        ZSql = ZSql & "Envase1 ,"
                        ZSql = ZSql & "Envase2 ,"
                        ZSql = ZSql & "Envase3 ,"
                        ZSql = ZSql & "Envase4 ,"
                        ZSql = ZSql & "Envase5 ,"
                        ZSql = ZSql & "Envase6 ,"
                        ZSql = ZSql & "Proceso ,"
                        ZSql = ZSql & "Costo ,"
                        ZSql = ZSql & "Factor ,"
                        ZSql = ZSql & "WDate ,"
                        ZSql = ZSql & "ImpreAdi ,"
                        ZSql = ZSql & "Clase ,"
                        ZSql = ZSql & "Secundario ,"
                        ZSql = ZSql & "Riesgo ,"
                        ZSql = ZSql & "Intervencion ,"
                        ZSql = ZSql & "Naciones ,"
                        ZSql = ZSql & "Embalaje ,"
                        ZSql = ZSql & "Controla ,"
                        ZSql = ZSql & "Sedronar ,"
                        ZSql = ZSql & "CodSedronar ,"
                        ZSql = ZSql & "ImpreVto ,"
                        ZSql = ZSql & "Marca ,"
                        ZSql = ZSql & "Observaciones ,"
                        ZSql = ZSql & "TipoEti ,"
                        ZSql = ZSql & "Escrito ,"
                        ZSql = ZSql & "Fabrica ,"
                        ZSql = ZSql & "FabricaII ,"
                        ZSql = ZSql & "FabricaIII ,"
                        ZSql = ZSql & "LoteAutorizado ,"
                        ZSql = ZSql & "Conservacion ,"
                        ZSql = ZSql & "ConservacionII ,"
                        ZSql = ZSql & "Vida ,"
                        ZSql = ZSql & "Seguridad ,"
                        ZSql = ZSql & "Version ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "FechaVersion ,"
                        ZSql = ZSql & "FechaVersionI ,"
                        ZSql = ZSql & "FechaVersionII ,"
                        ZSql = ZSql & "Estado ,"
                        ZSql = ZSql & "EstadoI ,"
                        ZSql = ZSql & "EstadoII ,"
                        ZSql = ZSql & "Observa ,"
                        ZSql = ZSql & "ObservaI ,"
                        ZSql = ZSql & "ObservaII ,"
                        ZSql = ZSql & "DescripcionIngles ,"
                        ZSql = ZSql & "DescriEtiquetaIngles ,"
                        ZSql = ZSql & "ConservacionIngles ,"
                        ZSql = ZSql & "ConservacionIIIngles ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WCodigo + "',"
                        ZSql = ZSql & "'" + WDescripcion + "',"
                        ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WUnidad + "',"
                        ZSql = ZSql & "'" + WInicial + "',"
                        ZSql = ZSql & "'" + WEntradas + "',"
                        ZSql = ZSql & "'" + WSalidas + "',"
                        ZSql = ZSql & "'" + WMinimo + "',"
                        ZSql = ZSql & "'" + WMinimo1 + "',"
                        ZSql = ZSql & "'" + WDeposito + "',"
                        ZSql = ZSql & "'" + WPedido + "',"
                        ZSql = ZSql & "'" + WEnvase1 + "',"
                        ZSql = ZSql & "'" + WEnvase2 + "',"
                        ZSql = ZSql & "'" + WEnvase3 + "',"
                        ZSql = ZSql & "'" + WEnvase4 + "',"
                        ZSql = ZSql & "'" + WEnvase5 + "',"
                        ZSql = ZSql & "'" + WEnvase6 + "',"
                        ZSql = ZSql & "'" + WProceso + "',"
                        ZSql = ZSql & "'" + WCosto + "',"
                        ZSql = ZSql & "'" + WFactor + "',"
                        ZSql = ZSql & "'" + WDate + "',"
                        ZSql = ZSql & "'" + WImpreadi + "',"
                        ZSql = ZSql & "'" + WClase + "',"
                        ZSql = ZSql & "'" + WSecundario + "',"
                        ZSql = ZSql & "'" + WRiesgo + "',"
                        ZSql = ZSql & "'" + WIntervencion + "',"
                        ZSql = ZSql & "'" + WNaciones + "',"
                        ZSql = ZSql & "'" + WEmbalaje + "',"
                        ZSql = ZSql & "'" + WControla + "',"
                        ZSql = ZSql & "'" + WSedronar + "',"
                        ZSql = ZSql & "'" + WCodSedronar + "',"
                        ZSql = ZSql & "'" + WImpreVto + "',"
                        ZSql = ZSql & "'" + WMarca + "',"
                        ZSql = ZSql & "'" + WObservaciones + "',"
                        ZSql = ZSql & "'" + WTipoeti + "',"
                        ZSql = ZSql & "'" + WEscrito + "',"
                        ZSql = ZSql & "'" + WFabrica + "',"
                        ZSql = ZSql & "'" + WFabricaII + "',"
                        ZSql = ZSql & "'" + WFabricaIII + "',"
                        ZSql = ZSql & "'" + WLoteAutorizado + "',"
                        ZSql = ZSql & "'" + WConservacion + "',"
                        ZSql = ZSql & "'" + WConservacionII + "',"
                        ZSql = ZSql & "'" + WVida + "',"
                        ZSql = ZSql & "'" + WSeguridad + "',"
                        ZSql = ZSql & "'" + WVersion + "',"
                        ZSql = ZSql & "'" + WVersionI + "',"
                        ZSql = ZSql & "'" + WVersionII + "',"
                        ZSql = ZSql & "'" + WFechaVersion + "',"
                        ZSql = ZSql & "'" + WFechaVersionI + "',"
                        ZSql = ZSql & "'" + WFechaVersionII + "',"
                        ZSql = ZSql & "'" + WEstado + "',"
                        ZSql = ZSql & "'" + WEstadoI + "',"
                        ZSql = ZSql & "'" + WEstadoII + "',"
                        ZSql = ZSql & "'" + WObserva + "',"
                        ZSql = ZSql & "'" + WObservaI + "',"
                        ZSql = ZSql & "'" + WObservaII + "',"
                        ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                        ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                        ZSql = ZSql & "'" + WMetodo + "',"
                        ZSql = ZSql & "'" + WEfluentes + "')"
      
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                    End If
        
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
                    ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
                    ZSql = ZSql & "ImpreAdi = " + "'" + ZZImpreAdi + "',"
                    ZSql = ZSql & "Clase = " + "'" + ZZClase + "',"
                    ZSql = ZSql & "Secundario = " + "'" + ZZSecundario + "',"
                    ZSql = ZSql & "Riesgo = " + "'" + ZZRiesgo + "',"
                    ZSql = ZSql & "Intervencion = " + "'" + ZZIntervencion + "',"
                    ZSql = ZSql & "Naciones = " + "'" + ZZNaciones + "',"
                    ZSql = ZSql & "Embalaje = " + "'" + ZZEmbalaje + "',"
                    ZSql = ZSql & "DescriOnu = " + "'" + ZZDescriOnu + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WNk + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
                    Rem da de alta el Re
        
                    WRe = "RE" + Right$(Codigo.Text, 10)
            
                    spTerminado = "ConsultaTerminado " + "'" + WRe + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount = 0 Then
            
                        rstTerminado.Close
        
                        WCodigo = WRe
                        WDescripcion = Descripcion.Text
                        WDescriEtiqueta = DescriEtiqueta.Text
                        WLinea = Linea.Text
                        WUnidad = Unidad.Text
                        WInicial = ""
                        WEntradas = ""
                        WSalidas = ""
                        WMinimo = ""
                        WMinimo1 = ""
                        WDeposito = ""
                        WPedido = ""
                        WEnvase1 = Envase1.Text
                        WEnvase2 = Envase2.Text
                        WEnvase3 = Envase3.Text
                        WEnvase4 = Envase4.Text
                        WEnvase5 = Envase5.Text
                        WEnvase6 = Envase6.Text
                        WProceso = ""
                        WCosto = ""
                        WFactor = ""
                        WDate = Date$
                        WImpreadi = ""
                        WIntervencion = ""
                        WClase = ""
                        WSecundario = ""
                        WRiesgo = ""
                        WNaciones = ""
                        WEmbalaje = ""
                        WVersion = ""
                        WFechaVersion = "  /  /    "
                        WControla = "0"
                        WSedronar = "0"
                        WCodSedronar = ""
                        WImpreVto = "0"
                        WMarca = "0"
                        WObservaciones = ""
                        WEscrito = "0"
                
                        WFabrica = ""
                        WFabricaII = ""
                        WFabricaIII = ""
                        WLoteAutorizado = ""
                        WConservacion = ""
                        WConservacionII = ""
                        WVida = ""
                        WSeguridad = ""
                
                        WVersion = ""
                        WVersionI = ""
                        WVersionII = ""
            
                        WFechaVersion = "  /  /    "
                        WFechaVersionI = "  /  /    "
                        WFechaVersionII = "  /  /    "
            
                        WEstado = ""
                        WEstadoI = ""
                        WEstadoII = ""
            
                        WObserva = ""
                        WObservaI = ""
                        WObservaII = ""
            
                        WMetodo = ""
                        WEfluentes = "0"
                
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO Terminado ("
                        ZSql = ZSql & "Codigo ,"
                        ZSql = ZSql & "Descripcion ,"
                        ZSql = ZSql & "DescriEtiqueta ,"
                        ZSql = ZSql & "Linea ,"
                        ZSql = ZSql & "Unidad ,"
                        ZSql = ZSql & "Inicial ,"
                        ZSql = ZSql & "Entradas ,"
                        ZSql = ZSql & "Salidas ,"
                        ZSql = ZSql & "Minimo ,"
                        ZSql = ZSql & "Minimo1 ,"
                        ZSql = ZSql & "Deposito ,"
                        ZSql = ZSql & "Pedido ,"
                        ZSql = ZSql & "Envase1 ,"
                        ZSql = ZSql & "Envase2 ,"
                        ZSql = ZSql & "Envase3 ,"
                        ZSql = ZSql & "Envase4 ,"
                        ZSql = ZSql & "Envase5 ,"
                        ZSql = ZSql & "Envase6 ,"
                        ZSql = ZSql & "Proceso ,"
                        ZSql = ZSql & "Costo ,"
                        ZSql = ZSql & "Factor ,"
                        ZSql = ZSql & "WDate ,"
                        ZSql = ZSql & "ImpreAdi ,"
                        ZSql = ZSql & "Clase ,"
                        ZSql = ZSql & "Secundario ,"
                        ZSql = ZSql & "Riesgo ,"
                        ZSql = ZSql & "Intervencion ,"
                        ZSql = ZSql & "Naciones ,"
                        ZSql = ZSql & "Embalaje ,"
                        ZSql = ZSql & "Controla ,"
                        ZSql = ZSql & "Sedronar ,"
                        ZSql = ZSql & "CodSedronar ,"
                        ZSql = ZSql & "ImpreVto ,"
                        ZSql = ZSql & "Marca ,"
                        ZSql = ZSql & "Observaciones ,"
                        ZSql = ZSql & "TipoEti ,"
                        ZSql = ZSql & "Escrito ,"
                        ZSql = ZSql & "Fabrica ,"
                        ZSql = ZSql & "FabricaII ,"
                        ZSql = ZSql & "FabricaIII ,"
                        ZSql = ZSql & "LoteAutorizado ,"
                        ZSql = ZSql & "Conservacion ,"
                        ZSql = ZSql & "ConservacionII ,"
                        ZSql = ZSql & "Vida ,"
                        ZSql = ZSql & "Seguridad ,"
                        ZSql = ZSql & "Version ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "FechaVersion ,"
                        ZSql = ZSql & "FechaVersionI ,"
                        ZSql = ZSql & "FechaVersionII ,"
                        ZSql = ZSql & "Estado ,"
                        ZSql = ZSql & "EstadoI ,"
                        ZSql = ZSql & "EstadoII ,"
                        ZSql = ZSql & "Observa ,"
                        ZSql = ZSql & "ObservaI ,"
                        ZSql = ZSql & "ObservaII ,"
                        ZSql = ZSql & "DescripcionIngles ,"
                        ZSql = ZSql & "DescriEtiquetaIngles ,"
                        ZSql = ZSql & "ConservacionIngles ,"
                        ZSql = ZSql & "ConservacionIIIngles ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WCodigo + "',"
                        ZSql = ZSql & "'" + WDescripcion + "',"
                        ZSql = ZSql & "'" + WDescriEtiqueta + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WUnidad + "',"
                        ZSql = ZSql & "'" + WInicial + "',"
                        ZSql = ZSql & "'" + WEntradas + "',"
                        ZSql = ZSql & "'" + WSalidas + "',"
                        ZSql = ZSql & "'" + WMinimo + "',"
                        ZSql = ZSql & "'" + WMinimo1 + "',"
                        ZSql = ZSql & "'" + WDeposito + "',"
                        ZSql = ZSql & "'" + WPedido + "',"
                        ZSql = ZSql & "'" + WEnvase1 + "',"
                        ZSql = ZSql & "'" + WEnvase2 + "',"
                        ZSql = ZSql & "'" + WEnvase3 + "',"
                        ZSql = ZSql & "'" + WEnvase4 + "',"
                        ZSql = ZSql & "'" + WEnvase5 + "',"
                        ZSql = ZSql & "'" + WEnvase6 + "',"
                        ZSql = ZSql & "'" + WProceso + "',"
                        ZSql = ZSql & "'" + WCosto + "',"
                        ZSql = ZSql & "'" + WFactor + "',"
                        ZSql = ZSql & "'" + WDate + "',"
                        ZSql = ZSql & "'" + WImpreadi + "',"
                        ZSql = ZSql & "'" + WClase + "',"
                        ZSql = ZSql & "'" + WSecundario + "',"
                        ZSql = ZSql & "'" + WRiesgo + "',"
                        ZSql = ZSql & "'" + WIntervencion + "',"
                        ZSql = ZSql & "'" + WNaciones + "',"
                        ZSql = ZSql & "'" + WEmbalaje + "',"
                        ZSql = ZSql & "'" + WControla + "',"
                        ZSql = ZSql & "'" + WSedronar + "',"
                        ZSql = ZSql & "'" + WCodSedronar + "',"
                        ZSql = ZSql & "'" + WImpreVto + "',"
                        ZSql = ZSql & "'" + WMarca + "',"
                        ZSql = ZSql & "'" + WObservaciones + "',"
                        ZSql = ZSql & "'" + WTipoeti + "',"
                        ZSql = ZSql & "'" + WEscrito + "',"
                        ZSql = ZSql & "'" + WFabrica + "',"
                        ZSql = ZSql & "'" + WFabricaII + "',"
                        ZSql = ZSql & "'" + WFabricaIII + "',"
                        ZSql = ZSql & "'" + WLoteAutorizado + "',"
                        ZSql = ZSql & "'" + WConservacion + "',"
                        ZSql = ZSql & "'" + WConservacionII + "',"
                        ZSql = ZSql & "'" + WVida + "',"
                        ZSql = ZSql & "'" + WSeguridad + "',"
                        ZSql = ZSql & "'" + WVersion + "',"
                        ZSql = ZSql & "'" + WVersionI + "',"
                        ZSql = ZSql & "'" + WVersionII + "',"
                        ZSql = ZSql & "'" + WFechaVersion + "',"
                        ZSql = ZSql & "'" + WFechaVersionI + "',"
                        ZSql = ZSql & "'" + WFechaVersionII + "',"
                        ZSql = ZSql & "'" + WEstado + "',"
                        ZSql = ZSql & "'" + WEstadoI + "',"
                        ZSql = ZSql & "'" + WEstadoII + "',"
                        ZSql = ZSql & "'" + WObserva + "',"
                        ZSql = ZSql & "'" + WObservaI + "',"
                        ZSql = ZSql & "'" + WObservaII + "',"
                        ZSql = ZSql & "'" + DescripcionIngles.Text + "',"
                        ZSql = ZSql & "'" + DescriEtiquetaIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIngles.Text + "',"
                        ZSql = ZSql & "'" + ConservacionIIIngles.Text + "',"
                        ZSql = ZSql & "'" + WMetodo + "',"
                        ZSql = ZSql & "'" + WEfluentes + "')"
      
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                    End If
        
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "Restriccion = " + "'" + WRestriccion + "',"
                    ZSql = ZSql & "Responsable = " + "'" + Responsable + "',"
                    ZSql = ZSql & "ImpreAdi = " + "'" + ZZImpreAdi + "',"
                    ZSql = ZSql & "Clase = " + "'" + ZZClase + "',"
                    ZSql = ZSql & "Secundario = " + "'" + ZZSecundario + "',"
                    ZSql = ZSql & "Riesgo = " + "'" + ZZRiesgo + "',"
                    ZSql = ZSql & "Intervencion = " + "'" + ZZIntervencion + "',"
                    ZSql = ZSql & "Naciones = " + "'" + ZZNaciones + "',"
                    ZSql = ZSql & "Embalaje = " + "'" + ZZEmbalaje + "',"
                    ZSql = ZSql & "DescriOnu = " + "'" + ZZDescriOnu + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WRe + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            Next Cicla
        
            Select Case Val(XEmpresa)
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

    WProceso = 1
    
    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
        WGraba = ""
    If Codigo.Text <> "" Then
    
        spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            rstTerminado.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spTerminado = "BorrarTerminado " + "'" + Codigo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
   
    End If
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = "  -     -   "
    Descripcion.Text = ""
    DescriEtiqueta.Text = ""
    Linea.Text = ""
    Unidad.Text = ""
    Inicial.Text = ""
    Entradas.Text = ""
    Salidas.Text = ""
    Minimo.Text = ""
    Minimo1.Text = ""
    Fabrica.Text = ""
    FabricaII.Text = ""
    FabricaIII.Text = ""
    Rem LoteAutorizado.Text = ""
    Deposito.Text = ""
    Rem Pedido.text = ""
    Rem Envase.text = ""
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    Envase6.Text = ""
    Proceso.Caption = ""
    Impreadi.Text = ""
    Clase.Text = ""
    Secundario.Text = ""
    Riesgo.Text = ""
    Intervencion.Text = ""
    Naciones.Text = ""
    Embalaje.Text = ""
    Stock.Caption = ""
    DescriLinea.Caption = ""
    Codigo.SetFocus
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    WGraba = ""
    Nk.Caption = ""
    RE.Caption = ""
    Version.Text = ""
    FechaVersion.Text = ""
    Observaciones.Text = ""
    TipoEti.Text = ""
    Conservacion.Text = ""
    ConservacionII.Text = ""
    Vida.Text = ""
    Seguridad.Text = ""
    Caracteristicas.Text = ""
    HojaTecnica.Text = ""
    CodSedronar.Text = ""
    
    Metodo.Text = ""
    Efluentes.Text = ""
    DesEfluentes.Text = ""
    DescripcionIngles.Text = ""
    DescriEtiquetaIngles.Text = ""
    ConservacionIngles.Text = ""
    ConservacionIIIngles.Text = ""
    
    Version.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    
    FechaVersion.Text = "  /  /    "
    FechaVersionI.Text = "  /  /    "
    FechaVersionII.Text = "  /  /    "
    
    Estado.Text = ""
    EstadoI.Text = ""
    EstadoII.Text = ""
    
    Observa.Text = ""
    ObservaI.Text = ""
    ObservaII.Text = ""
    
    PasaControla = "N"
    PasaEscrito = "N"
    PasaCarga = "N"
    PasaEstadoProducto = "N"
    PasaListaProducto = "N"
    
    Controla.ListIndex = 0
    Sedronar.ListIndex = 0
    ImpreVto.ListIndex = 0
    Marca.ListIndex = 0
    Escrito.ListIndex = 0
    Carga.ListIndex = 0
    EstadoProducto.ListIndex = 0
    ListaProducto.ListIndex = 0
    Restriccion.Value = 0
    
    PasaControla = "S"
    PasaEscrito = "S"
    PasaCarga = "S"
    PasaEstadoProducto = "S"
    PasaListaProducto = "S"
    
    LabelDescriEtiqueta.Visible = True
    DescriEtiqueta.Visible = True
    LabelDescriEtiquetaIngles.Visible = True
    DescriEtiquetaIngles.Visible = True
    
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
    ZCampo14 = "N"
    ZCampo15 = "N"
    ZCampo16 = "N"
    ZCampo17 = "N"
    ZCampo18 = "N"
    ZCampo19 = "N"
    ZCampo20 = "N"
    ZCampo21 = "N"
    ZCampo22 = "N"
    ZCampo23 = "N"
    ZCampo24 = "N"
    ZCampo25 = "N"
    ZCampo26 = "N"
    ZCampo27 = "N"
    ZCampo28 = "N"
    ZCampo29 = "N"
    ZCampo30 = "N"
    ZCampo31 = "N"
    ZCampo32 = "N"
    
    Codigo.SetFocus

End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    
    Codigo.SetFocus
    PrgTermi.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spTerminado = "AnteriorTerminado " + "'" + Codigo.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveLast
            Codigo.Text = rstTerminado!Codigo
        End With
        rstTerminado.Close
    End If
    Call Imprime_Datos
    Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Terminados", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If DBGrid1.Col < 3 Then
            DBGrid1.Col = DBGrid1.Col + 1
                Else
            If DBGrid1.Row < 99 Then
                DBGrid1.Row = DBGrid1.Row + 1
                DBGrid1.Col = 0
            End If
        End If
    End If
End Sub

Private Sub Command1_Click()

    Erase Vector
    Lugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecifUnifica"
    spEspecifUnifica = Sql1 + Sql2
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    Lugar = Lugar + 1
                    Vector(Lugar, 1) = rstEspecifUnifica!Producto
                    Vector(Lugar, 2) = IIf(IsNull(rstEspecifUnifica!Version), "0", rstEspecifUnifica!Version)
                    Vector(Lugar, 3) = IIf(IsNull(rstEspecifUnifica!Fecha), "", rstEspecifUnifica!Fecha)
                    Vector(Lugar, 4) = IIf(IsNull(rstEspecifUnifica!Estado), "S", rstEspecifUnifica!Estado)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    XEmpresa = Wempresa
    Erase CargaEmpresa
            
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
            
        Case 2, 4, 8, 9
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case Else
    End Select
            
    For Cicla = 1 To 7
            
        If CargaEmpresa(Cicla, 1) <> "" Then
            
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            For XCiclo = 1 To Lugar
            
                XXCodigo = Vector(XCiclo, 1)
                XXVersion = Vector(XCiclo, 2)
                XXFechaVersion = Vector(XCiclo, 3)
                XXEstado = Vector(XCiclo, 4)
                XXObservaciones = ""
            
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "VersionII = " + "'" + XXVersion + "',"
                ZSql = ZSql & "FechaVersionII = " + "'" + XXFechaVersion + "',"
                ZSql = ZSql & "EstadoII = " + "'" + XXEstado + "',"
                ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + XXCodigo + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            Next XCiclo
                    
        End If
                    
    Next Cicla
            
    Select Case Val(XEmpresa)
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
    
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgTermi.Caption = "Ingreso de Producto Terminado :  " + !Nombre
        End If
    End With
    
    PasaControla = "N"
    PasaEscrito = "N"
    PasaCarga = "N"
    PasaEstadoProducto = "N"
    PasaListaProducto = "N"
    
    Controla.Clear
    
    Controla.AddItem "Controla Lote"
    Controla.AddItem "No Controla Lote"
    Controla.AddItem "A Granel"
    
    Controla.ListIndex = 0
    
    Sedronar.Clear
    
    Sedronar.AddItem "No"
    Sedronar.AddItem "Si"
    
    Sedronar.ListIndex = 0
    
    ImpreVto.Clear
    
    ImpreVto.AddItem "No Imprime"
    ImpreVto.AddItem "Imprime"
    
    ImpreVto.ListIndex = 0
    
    Marca.Clear
    
    Marca.AddItem "Exige Pedido"
    Marca.AddItem "No Exige Pedido"
    
    Marca.ListIndex = 0
    
    Escrito.Clear
    
    Escrito.AddItem "No"
    Escrito.AddItem "Si"
    Escrito.AddItem "Nuevo"
    
    Escrito.ListIndex = 0
    
    Carga.Clear
    
    Carga.AddItem ""
    Carga.AddItem "Alcalino"
    Carga.AddItem "Acido"
    Carga.AddItem "No ionico/Alcalino"
    Carga.AddItem "Neutro"
    
    Carga.ListIndex = 0
    
    EstadoProducto.Clear
    
    EstadoProducto.AddItem ""
    EstadoProducto.AddItem "Polvo"
    EstadoProducto.AddItem "Liquido"
    EstadoProducto.AddItem "Metal"
    EstadoProducto.AddItem "Pasta"
    
    EstadoProducto.ListIndex = 0
    
    ListaProducto.Clear
    
    ListaProducto.AddItem ""
    ListaProducto.AddItem "Si"
    ListaProducto.AddItem "No"
    
    ListaProducto.ListIndex = 0
    
    Restriccion.Value = 0
    
    
    PasaControla = "S"
    PasaEscrito = "S"
    PasaCarga = "S"
    PasaEstadoProducto = "S"
    PasaListaProducto = "S"
    
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
    ZCampo14 = "N"
    ZCampo15 = "N"
    ZCampo16 = "N"
    ZCampo17 = "N"
    ZCampo18 = "N"
    ZCampo19 = "N"
    ZCampo20 = "N"
    ZCampo21 = "N"
    ZCampo22 = "N"
    ZCampo23 = "N"
    ZCampo24 = "N"
    ZCampo25 = "N"
    ZCampo26 = "N"
    ZCampo27 = "N"
    ZCampo28 = "N"
    ZCampo29 = "N"
    ZCampo30 = "N"
    ZCampo31 = "N"
    ZCampo32 = "N"

End Sub

Private Sub Lista_Click()
    Desdecodigo.Text = "  -     -   "
    HastaCodigo.Text = "  -     -   "
    DesdeLinea.Text = ""
    HastaLinea.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desdecodigo.SetFocus
End Sub

Private Sub DesdeCodigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdecodigo.Text = UCase(Desdecodigo.Text)
        HastaCodigo.SetFocus
    End If
End Sub

Private Sub HastaCodigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCodigo.Text = UCase(HastaCodigo.Text)
        DesdeLinea.SetFocus
    End If
End Sub

Private Sub DesdeLinea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaLinea.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaLinea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdecodigo.SetFocus
    End If
     Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Rejilla_KeyDown(KeyCode As Integer, Shift As Integer)
    Rem Rejilla.Row = Rejilla.Row + 1
End Sub

Private Sub Rejilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Rejilla.Col < 1 Then
            Rejilla.Col = Rejilla.Col + 1
        End If
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
      If XTipoPro = "FA" Then
       
       DescriEtiqueta.SetFocus
           Else
      Linea.SetFocus
      End If
    End If
End Sub

Private Sub DescriEtiqueta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Linea.SetFocus
    End If
End Sub

Private Sub Linea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo3 = "S"
        spLineas = "ConsultaLinea " + "'" + Linea.Text + "'"
        Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
        If rstLineas.RecordCount > 0 Then
            DescriLinea.Caption = rstLineas!Nombre
            rstLineas.Close
            Impreadi.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ImpreAdi_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Impreadi.Text = "" Or Impreadi.Text = "S" Or Impreadi.Text = "N" Then
            ZCampo4 = "S"
            Minimo.SetFocus
        End If
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
        Deposito.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Deposito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo7 = "S"
        Controla.SetFocus
    End If
End Sub

Private Sub Controla_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo8 = "S"
        Vida.SetFocus
    End If
End Sub

Private Sub Controla_Click()
    If PasaControla <> "N" Then
        ZCampo8 = "S"
        Vida.SetFocus
    End If
End Sub

Private Sub Marca_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo8 = "S"
        Vida.SetFocus
    End If
End Sub

Private Sub Marca_Click()
    If PasaControla <> "N" Then
        ZCampo8 = "S"
        Vida.SetFocus
    End If
End Sub

Private Sub Vida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo9 = "S"
        Escrito.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escrito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo10 = "S"
        Fabrica.SetFocus
    End If
End Sub

Private Sub Escrito_Click()
    If PasaEscrito <> "N" Then
        ZCampo10 = "S"
        Fabrica.SetFocus
    End If
End Sub

Private Sub Fabrica_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo11 = "S"
        FabricaII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FabricaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FabricaIII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FabricaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo11 = "S"
        Conservacion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub LoteAutorizado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Conservacion.SetFocus
    End If
End Sub

Private Sub Conservacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo12 = "S"
        ConservacionII.SetFocus
    End If
End Sub

Private Sub ConservacionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo13 = "S"
        Metodo.SetFocus
    End If
End Sub

Private Sub Metodo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo14 = "S"
        Efluentes.SetFocus
    End If
End Sub

Private Sub Efluentes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo15 = "S"
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM Efluentes"
        ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + Efluentes.Text + "'"
        spEfluentes = ZSql
        Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
        If rstEfluentes.RecordCount > 0 Then
            DesEfluentes.Text = rstEfluentes!Descripcion
            rstEfluentes.Close
            Envase1.SetFocus
        End If
    End If
End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo16 = "S"
        If Val(Envase1.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase1.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri1.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Envase2.SetFocus
            End If
                Else
            Descri1.Caption = ""
            Envase2.SetFocus
        End If
    End If
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo17 = "S"
        If Val(Envase2.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase2.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri2.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Envase3.SetFocus
            End If
                Else
            Descri2.Caption = ""
            Envase3.SetFocus
        End If
    End If
End Sub

Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo18 = "S"
        If Val(Envase3.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase3.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri3.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Envase4.SetFocus
            End If
                Else
            Descri3.Caption = ""
            Envase4.SetFocus
        End If
    End If
End Sub

Private Sub Envase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo19 = "S"
        If Val(Envase4.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase4.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri4.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Envase5.SetFocus
            End If
                Else
            Descri4.Caption = ""
            Envase5.SetFocus
        End If
    End If
End Sub

Private Sub Envase5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo20 = "S"
        If Val(Envase5.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase5.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri5.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Envase6.SetFocus
            End If
                Else
            Descri5.Caption = ""
            Envase6.SetFocus
        End If
    End If
End Sub

Private Sub Envase6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo21 = "S"
        If Val(Envase5.Text) <> 0 Then
            spEnvase = "ConsultaEnvases " + "'" + Envase6.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                Descri6.Caption = rstEnvase!Descripcion
                rstEnvase.Close
                Descripcion.SetFocus
            End If
                Else
            Descri6.Caption = ""
            Observaciones.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo22 = "S"
        TipoEti.SetFocus
    End If
End Sub

Private Sub TipoEti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo23 = "S"
        Naciones.SetFocus
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
                    Secundario.Text = Trim(ZPeligrosoIV)
                    Caracteristicas.Text = Left$(ZPeligrosoII, 100)
                    Intervencion.Text = Trim(ZPeligrosoI)
                    Riesgo.Text = Trim(ZPeligrosoV)
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
                Secundario.Text = ""
                Riesgo.Text = ""
                Embalaje.Text = ""
                Caracteristicas.Text = ""
                Intervencion.Text = ""
            
                m$ = "Nro de Naciones Unidas Inexistente"
                a% = MsgBox(m$, 0, "Archivo de Productos Terminados")
                Exit Sub
                
            End If
            
                Else
                
            Clase.Text = ""
            Secundario.Text = ""
            Riesgo.Text = ""
            Embalaje.Text = ""
            Caracteristicas.Text = ""
            Intervencion.Text = ""
            
        End If
    
        ZCampo24 = "S"
        ZCampo25 = "S"
        ZCampo26 = "S"
        ZCampo27 = "S"
        ZCampo29 = "S"
        Seguridad.SetFocus
        
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
        Secundario.Text = ""
        Riesgo.Text = ""
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
                Secundario.Text = Trim(ZPeligrosoIV)
                Riesgo.Text = Trim(ZPeligrosoV)
                Embalaje.Text = Trim(ZPeligrosoVI)
                Caracteristicas.Text = Left$(ZPeligrosoII, 100)
                Intervencion.Text = Trim(ZPeligrosoI)
                
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
        ZCampo24 = "S"
        Secundario.SetFocus
    End If
End Sub

Private Sub Secundario_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo24 = "S"
        Riesgo.SetFocus
    End If
End Sub

Private Sub Riesgo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo24 = "S"
        Intervencion.SetFocus
    End If
End Sub

Private Sub Embalaje_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo27 = "S"
        Seguridad.SetFocus
    End If
End Sub

Private Sub Caracteristicas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo29 = "S"
        Carga.SetFocus
    End If
End Sub

Private Sub Intervencion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo25 = "S"
        Seguridad.SetFocus
    End If
End Sub

Private Sub Seguridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo28 = "S"
        Carga.SetFocus
    End If
End Sub

Private Sub Carga_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo30 = "S"
        EstadoProducto.SetFocus
    End If
End Sub

Private Sub Carga_Click()
    If PasaCarga <> "N" Then
        ZCampo30 = "S"
        EstadoProducto.SetFocus
    End If
End Sub

Private Sub EstadoProducto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo31 = "S"
        ListaProducto.SetFocus
    End If
End Sub

Private Sub EstadoProducto_Click()
    If PasaEstadoProducto <> "N" Then
        ZCampo31 = "S"
        ListaProducto.SetFocus
    End If
End Sub

Private Sub ListaProducto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZCampo32 = "S"
        Descripcion.SetFocus
    End If
End Sub

Private Sub ListaProducto_Click()
    If PasaListaProducto <> "N" Then
        ZCampo32 = "S"
        Descripcion.SetFocus
    End If
End Sub

Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Codigo.Text = UCase(Codigo.Text)
        If Codigo.Text <> "" Then
            spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount <= 0 Then
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
                Descripcion.SetFocus
                        Else
                Codigo.Text = rstTerminado!Codigo
                rstTerminado.Close
                Call Imprime_Datos
            End If
        End If
    End If
End Sub

Private Sub Consulta_Click()

     
     Opcion.Clear

     Opcion.AddItem "Productos"
     Opcion.AddItem "Lineas"
     Opcion.AddItem "Envases"
     Opcion.AddItem "Efluentes de Lavado"
     Opcion.AddItem "Naciones Unidas"

     Opcion.Visible = True
     
 End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
            
        Case 1
            spLineas = "ListaLinea"
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
            
                With rstLineas
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstLineas!Linea) + " " + rstLineas!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstLineas!Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineas.Close
            End If
        
        Case 2
            spEnvase = "ListaEnvases"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
            
                With rstEnvase
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnvase!Envases) + " " + rstEnvase!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnvase!Envases
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnvase.Close
            End If
            
        Case 3
            Sql1 = "Select *"
            Sql2 = " FROM Efluentes"
            Sql3 = " Order by Efluentes.Codigo"
            spEfluentes = Sql1 + Sql2 + Sql3
            Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
            If rstEfluentes.RecordCount > 0 Then
                With rstEfluentes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEfluentes!Codigo) + " " + rstEfluentes!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEfluentes!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEfluentes.Close
            End If
            
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
                                                        
                            IngresaItem = rstPeligroso!Ficha + " " + rstPeligroso!Descripcion
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
            ZZEntra = "N"
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
                                                        
                            ZZEntra = "S"
                            IngresaItem = rstPeligroso!Ficha + " " + rstPeligroso!Descripcion
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
        
            
        Case Else
    End Select
    Opcion.Visible = False
    If XIndice = 5 And ZZEntra = "N" Then
        Pantalla.Visible = False
            Else
        Pantalla.Visible = True
    End If

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WTerminado = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Codigo.Text = rstTerminado!Codigo
                rstTerminado.Close
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Codigo.Text = WTerminado
            End If
            
            Ayuda.Visible = False
            Rem Codigo.SetFocus
        
        Case 1
            Indice = Pantalla.ListIndex
            WLineas = WIndice.List(Indice)
            spLineas = "ConsultaLinea " + "'" + WLineas + "'"
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
                DescriLinea.Caption = rstLineas!Nombre
                Linea.Text = rstLineas!Linea
                rstLineas.Close
            End If
            Inicial.SetFocus
        
        Case 2
            WPasa = ""
            If Val(Envase1.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase1.Text = rstEnvase!Envases
                    Descri1.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Envase2.SetFocus
                End If
            End If
        
            If WPasa <> "S" Then
            If Val(Envase2.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase2.Text = rstEnvase!Envases
                    Descri2.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Envase3.SetFocus
                End If
            End If
            End If
        
            If WPasa <> "S" Then
            If Val(Envase3.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase3.Text = rstEnvase!Envases
                    Descri3.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Envase4.SetFocus
                End If
            End If
            End If
        
            If WPasa <> "S" Then
            If Val(Envase4.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase4.Text = rstEnvase!Envases
                    Descri4.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Envase5.SetFocus
                End If
            End If
            End If
        
            If WPasa <> "S" Then
            If Val(Envase5.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase5.Text = rstEnvase!Envases
                    Descri5.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Envase6.SetFocus
                End If
            End If
            End If
        
            If WPasa <> "S" Then
            If Val(Envase6.Text) = 0 Then
                Indice = Pantalla.ListIndex
                WEnvases = WIndice.List(Indice)
                spEnvase = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    Envase6.Text = rstEnvase!Envases
                    Descri6.Caption = rstEnvase!Descripcion
                    rstEnvase.Close
                    Descripcion.SetFocus
                End If
            End If
            End If
            
        Case 3
            Indice = Pantalla.ListIndex
            Efluentes.Text = WIndice.List(Indice)
            Call Efluentes_Keypress(13)
            
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
                Clase.Text = Trim(rstPeligroso!Clase)
                Secundario.Text = Trim(rstPeligroso!Secundario)
                Riesgo.Text = Trim(rstPeligroso!Riesgo)
                Embalaje.Text = Trim(rstPeligroso!Embalaje)
                Caracteristicas.Text = Left$(rstPeligroso!Descripcion, 100)
                Intervencion.Text = Trim(rstPeligroso!Ficha)
                rstPeligroso.Close
            End If
        
        Case Else
    End Select
    
End Sub


Private Sub Primer_Click()

    On Error GoTo WError
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
            
        With rstTerminado
            .MoveFirst
            Codigo.Text = rstTerminado!Codigo
        End With
    
        rstTerminado.Close
    
    End If
    
    Call Imprime_Datos
    Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()

    On Error GoTo Error_ultimo
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
            
        With rstTerminado
            .MoveLast
            Codigo.Text = rstTerminado!Codigo
        End With
    
        rstTerminado.Close
        
    End If
        
    Call Imprime_Datos
    Codigo.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     If coderr = 91 Then Resume Next
     Call Errores(coderr, "Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Rem Terminado.SetFocus
 End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spTerminado = "PosteriorTerminado " + "'" + Codigo.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
            Codigo.Text = rstTerminado!Codigo
        End With
    
        rstTerminado.Close
        
    End If
    
    Call Imprime_Datos
    Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Terminado", "No existe registro en el archivo")
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
                Call cmdDelete_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Especificaciones de Materia Prima")
            WClave.SetFocus
        End If

    End If

End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(rstTerminado!Descripcion) - WEspacios
                
                For AAa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, AAa, WEspacios) Then
                    
                        IngresaItem = rstTerminado!Codigo + "    " + rstTerminado!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstTerminado!Codigo
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
    
    rstTerminado.Close
    
    End If
    
    End If

End Sub

Private Sub HojaPend_Click()
    
    Sql1 = "UPDATE Hoja SET "
    Sql2 = " Realant = 0"
    Sql3 = " Where Realant IS NULL"
    spHoja = Sql1 + Sql2 + Sql3
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Hoja de Produccion Pendirentes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Hoja.Marca} <> " + Chr$(34) + "X" + Chr$(34) + " and {Hoja.Real} = 0 and {Hoja.Teorico} <> 0 and {Hoja.Renglon} = 1 and {Hoja.Producto} in " + Chr$(34) + Codigo.Text + Chr$(34) + " to " + Chr$(34) + Codigo.Text + Chr$(34)
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Renglon, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, Hoja.Real, Hoja.FechaIngOrd, Hoja.Marca, " _
                        + "Terminado.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Hoja.Producto = Terminado.Codigo AND " _
                        + "Hoja.Renglon = 1 AND " _
                        + "Hoja.Producto >= 'AA-00000-000' AND " _
                        + "Hoja.Producto <= 'ZZ-99999-999' AND " _
                        + "Hoja.Teorico <> 0 AND " _
                        + "Hoja.Real = 0 AND " _
                        + "Hoja.RealAnt = 0 AND " _
                        + "Hoja.Marca <> 'X'"
                        
    Listado.DataFiles(2) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    WListado = Listado.ReportFileName
    Listado.ReportFileName = "Wlisthojapend.rpt"
    Listado.Action = 1
    Listado.ReportFileName = WListado

End Sub

Private Sub Autoriza_Click()
    WClaveAutoriza.Text = ""
    PantaAutoriza.Visible = True
    WClaveAutoriza.SetFocus
End Sub

Private Sub CancelaAutoriza_Click()
    PantaAutoriza.Visible = False
    Call Codigo_KeyPress(13)
End Sub

Private Sub WClaveAutoriza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(WClaveAutoriza.Text) = "VERDE" Then
        
            XEmpresa = Wempresa
            Erase CargaEmpresa
        
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
                Case 2
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                Case 3
                    CargaEmpresa(1, 1) = "0003"
                    CargaEmpresa(1, 2) = "Empresa03"
                Case 4
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 5
                    CargaEmpresa(1, 1) = "0005"
                    CargaEmpresa(1, 2) = "Empresa05"
                Case 6
                    CargaEmpresa(1, 1) = "0006"
                    CargaEmpresa(1, 2) = "Empresa06"
                Case 7
                    CargaEmpresa(1, 1) = "0007"
                    CargaEmpresa(1, 2) = "Empresa07"
                Case 8
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 9
                    CargaEmpresa(1, 1) = "0009"
                    CargaEmpresa(1, 2) = "Empresa09"
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                Case 11
                    CargaEmpresa(1, 1) = "0011"
                    CargaEmpresa(1, 2) = "Empresa11"
                Case Else
            End Select
                
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                    
                        WVersionI = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
                        WVersionII = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
                        WVersionIII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
                        
                        WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                        WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                        WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                        
                        rstTerminado.Close
                        
                        If WEstadoI = "N" Then
                        
                            XFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            XEstado = "S"
                            XObservaciones = ""
                            
                            Sql1 = "UPDATE Terminado SET "
                            Sql2 = " FechaVersion = " + "'" + XFechaVersion + "',"
                            Sql3 = " Estado = " + "'" + XEstado + "',"
                            Sql4 = " Observa = " + "'" + XObservaciones + "'"
                            Sql5 = " Where Codigo = " + "'" + Codigo.Text + "'"
                            
                            spTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                        
                        If WEstadoII = "N" Then
                        
                            XFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            XEstado = "S"
                            XObservaciones = ""
                            
                            Sql1 = "UPDATE Terminado SET "
                            Sql2 = " FechaVersionI = " + "'" + XFechaVersion + "',"
                            Sql3 = " EstadoI = " + "'" + XEstado + "',"
                            Sql4 = " ObservaI = " + "'" + XObservaciones + "'"
                            Sql5 = " Where Codigo = " + "'" + Codigo.Text + "'"
                            
                            spTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                        
                        If WEstadoIII = "N" Then
                        
                            XFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            XEstado = "S"
                            XObservaciones = ""
                            
                            Sql1 = "UPDATE Terminado SET "
                            Sql2 = " FechaVersionII = " + "'" + XFechaVersion + "',"
                            Sql3 = " EstadoII = " + "'" + XEstado + "',"
                            Sql4 = " ObservaII = " + "'" + XObservaciones + "'"
                            Sql5 = " Where Codigo = " + "'" + Codigo.Text + "'"
                            
                            spTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                        
                    End If
                    
                End If
            Next Cicla
        
            Select Case Val(XEmpresa)
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
        
            Call CancelaAutoriza_Click
            
                Else
                
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Autorizacion de Version de Producto Terminado")
            WClaveAutoriza.SetFocus
            
        End If
    End If
End Sub

Private Sub CancelaLiberaHoja_Click()
    PantaLiberaHoja.Visible = False
End Sub

Private Sub WClaveLiberaHoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(WClaveLiberaHoja.Text) = "GOL08" Then
        
            XEmpresa = Wempresa
            Erase CargaEmpresa
        
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
                Case 2
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                Case 3
                    CargaEmpresa(1, 1) = "0003"
                    CargaEmpresa(1, 2) = "Empresa03"
                Case 4
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 5
                    CargaEmpresa(1, 1) = "0005"
                    CargaEmpresa(1, 2) = "Empresa05"
                Case 6
                    CargaEmpresa(1, 1) = "0006"
                    CargaEmpresa(1, 2) = "Empresa06"
                Case 7
                    CargaEmpresa(1, 1) = "0007"
                    CargaEmpresa(1, 2) = "Empresa07"
                Case 8
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case 9
                    CargaEmpresa(1, 1) = "0009"
                    CargaEmpresa(1, 2) = "Empresa09"
                Case 10
                    CargaEmpresa(1, 1) = "0010"
                    CargaEmpresa(1, 2) = "Empresa10"
                Case 11
                    CargaEmpresa(1, 1) = "0011"
                    CargaEmpresa(1, 2) = "Empresa11"
                Case Else
            End Select
                
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    Sql1 = "UPDATE Terminado SET "
                    Sql2 = " EstadoHoja = " + "'" + "S" + "'"
                    Sql3 = " Where Codigo = " + "'" + Codigo.Text + "'"
                    
                    spTerminado = Sql1 + Sql2 + Sql3
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            Next Cicla
        
            Select Case Val(XEmpresa)
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
        
            Call CancelaLiberaHoja_Click
            
                Else
                
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Autorizacion de Hojas de Seguridad")
            WClaveLiberaHoja.SetFocus
            
        End If
    End If
End Sub

Private Sub AltaIngles_Click()
    CargaIngles.Height = 3655
    CargaIngles.Left = 1680
    CargaIngles.Top = 1080
    CargaIngles.Width = 8215
    CargaIngles.Visible = True
    DescripcionIngles.SetFocus
End Sub

Private Sub CierraIngles_Click()
    CargaIngles.Visible = False
End Sub

Private Sub DescripcionIngles_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriEtiquetaIngles.SetFocus
    End If
End Sub

Private Sub DescriEtiquetaIngles_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ConservacionIngles.SetFocus
    End If
End Sub

Private Sub ConservacionIngles_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ConservacionIIIngles.SetFocus
    End If
End Sub

Private Sub ConservacionIIIngles_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CierraIngles_Click
    End If
End Sub




        
Private Sub Verifica_Msds()
    
    If Left$(UCase(Codigo.Text), 2) = "PT" Then

        Es = Descripcion.Text
        x = ""
        For XX = 1 To Len(Es)
            Y = Mid$(Es, XX, 1)
            If Y <> " " And Y <> "/" Then
                x = x + Y
            End If
        Next
        ZZCodArt = x + Mid$(Codigo.Text, 4, 5) + Right$(Codigo.Text, 3)
        
        ZZRuta = "w:\MSDSSIS\MSDS" + ZZCodArt + ".PDF"
        ZZEstado = Dir(ZZRuta)
        ZZEstado = Trim(ZZEstado)
        If ZZEstado = "" Then
            m$ = "El MSDS  (" + ZZCodArt + ")  no se ha encontrado"
            AAAAA% = MsgBox(m$, 0, "Impresion de comprobantes varios")
        End If
        
    End If

End Sub
        

