VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaSacAccion 
   Caption         =   "Carga de SAC -- Plan de Accion"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11775
   Begin VB.CommandButton CargaCausas 
      Caption         =   "Actualizacion de Causas"
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
      Left            =   9480
      TabIndex        =   63
      Top             =   6480
      Width           =   2175
   End
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   61
      Text            =   " "
      Top             =   1200
      Width           =   10455
   End
   Begin VB.TextBox Tipo 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "cargasacaccion.frx":0000
      Left            =   360
      List            =   "cargasacaccion.frx":0007
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
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
      Left            =   1440
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Accion11 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   42
      Text            =   " "
      Top             =   2400
      Width           =   8200
   End
   Begin VB.TextBox Accion12 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   41
      Text            =   " "
      Top             =   2640
      Width           =   8200
   End
   Begin VB.TextBox Accion21 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   40
      Text            =   " "
      Top             =   3120
      Width           =   8200
   End
   Begin VB.TextBox Accion22 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   39
      Text            =   " "
      Top             =   3360
      Width           =   8200
   End
   Begin VB.TextBox Accion31 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   38
      Text            =   " "
      Top             =   3720
      Width           =   8200
   End
   Begin VB.TextBox Accion32 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   37
      Text            =   " "
      Top             =   3960
      Width           =   8200
   End
   Begin VB.TextBox Accion41 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   36
      Text            =   " "
      Top             =   4440
      Width           =   8200
   End
   Begin VB.TextBox Accion42 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   35
      Text            =   " "
      Top             =   4680
      Width           =   8200
   End
   Begin VB.TextBox Accion51 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   34
      Text            =   " "
      Top             =   5160
      Width           =   8200
   End
   Begin VB.TextBox Accion52 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   33
      Text            =   " "
      Top             =   5400
      Width           =   8200
   End
   Begin VB.TextBox Accion61 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   32
      Text            =   " "
      Top             =   5880
      Width           =   8200
   End
   Begin VB.TextBox Accion62 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   31
      Text            =   " "
      Top             =   6120
      Width           =   8200
   End
   Begin VB.TextBox Responsable1 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   30
      Text            =   " "
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Responsable2 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   29
      Text            =   " "
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Responsable3 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   28
      Text            =   " "
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Responsable4 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   27
      Text            =   " "
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Responsable5 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   26
      Text            =   " "
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Responsable6 
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
      Left            =   8600
      MaxLength       =   6
      TabIndex        =   25
      Text            =   " "
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Ano 
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
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Titulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   9
      Text            =   " "
      Top             =   1560
      Width           =   10455
   End
   Begin VB.TextBox ResponsableDestino 
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
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   8
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableEmisor 
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   7
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Origen 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2535
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
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Centro 
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
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   855
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
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
   Begin MSMask.MaskEdBox Plazo1 
      Height          =   285
      Left            =   10600
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Plazo2 
      Height          =   285
      Left            =   10600
      TabIndex        =   44
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Plazo3 
      Height          =   285
      Left            =   10600
      TabIndex        =   45
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Plazo4 
      Height          =   285
      Left            =   10600
      TabIndex        =   46
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Plazo5 
      Height          =   285
      Left            =   10600
      TabIndex        =   47
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Plazo6 
      Height          =   285
      Left            =   10600
      TabIndex        =   48
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label26 
      Caption         =   "1"
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
      Left            =   75
      TabIndex        =   69
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label27 
      Caption         =   "2"
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
      Left            =   75
      TabIndex        =   68
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label28 
      Caption         =   "3"
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
      Left            =   75
      TabIndex        =   67
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label29 
      Caption         =   "4"
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
      Left            =   75
      TabIndex        =   66
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label30 
      Caption         =   "5"
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
      Left            =   75
      TabIndex        =   65
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label31 
      Caption         =   "6"
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
      Left            =   75
      TabIndex        =   64
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label DesTipo 
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
      Left            =   2040
      TabIndex        =   60
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Numero"
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
      Left            =   5160
      TabIndex        =   58
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8600
      TabIndex        =   57
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label DesResponsable1 
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
      Left            =   9300
      TabIndex        =   56
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plazo"
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
      Left            =   10600
      TabIndex        =   55
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acciones Correctivas"
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
      Left            =   240
      TabIndex        =   54
      Top             =   2040
      Width           =   8205
   End
   Begin VB.Label DesResponsable2 
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
      Left            =   9300
      TabIndex        =   53
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label DesResponsable3 
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
      Left            =   9300
      TabIndex        =   52
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label DesResponsable4 
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
      Left            =   9300
      TabIndex        =   51
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label DesResponsable5 
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
      Left            =   9300
      TabIndex        =   50
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label DesResponsable6 
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
      Left            =   9300
      TabIndex        =   49
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Titulo"
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
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label DesResponsableDestino 
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
      Left            =   7440
      TabIndex        =   23
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableEmisor 
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
      Left            =   2160
      TabIndex        =   22
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "Resp. Inv."
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
      Left            =   5400
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
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
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Emisor"
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
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
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
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label DesCentro 
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
      Left            =   9240
      TabIndex        =   17
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Centro"
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
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6600
      MouseIcon       =   "cargasacaccion.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "cargasacaccion.frx":031F
      ToolTipText     =   "Salida"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "cargasacaccion.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "cargasacaccion.frx":0E6B
      ToolTipText     =   "Elimina el Registro"
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "cargasacaccion.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "cargasacaccion.frx":19B7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5640
      MouseIcon       =   "cargasacaccion.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "cargasacaccion.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6720
      Width           =   480
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
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "PrgCargaSacAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim rstCargaSacII As Recordset
Dim spCargaSacII As String

Dim ZResponsableDestino As Integer
Dim ZResponsableCentro As Integer

Dim XParam As String
Dim ZZLugar As Integer

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If

End Sub

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        Centro.Text = rstCargaSac!Centro
        Fecha.Text = rstCargaSac!Fecha
        Origen.ListIndex = rstCargaSac!Origen
        Estado.ListIndex = rstCargaSac!Estado
        ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSac!ResponsableDestino
        Referencia.Text = Trim(rstCargaSac!Referencia)
        Titulo.Text = Trim(rstCargaSac!Titulo)
        
        rstCargaSac.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
    
        Accion11.Text = Trim(rstCargaSacII!Accion11)
        Accion12.Text = Trim(rstCargaSacII!Accion12)
        Accion21.Text = Trim(rstCargaSacII!Accion21)
        Accion22.Text = Trim(rstCargaSacII!Accion22)
        Accion31.Text = Trim(rstCargaSacII!Accion31)
        Accion32.Text = Trim(rstCargaSacII!Accion32)
        Accion41.Text = Trim(rstCargaSacII!Accion41)
        Accion42.Text = Trim(rstCargaSacII!Accion42)
        Accion51.Text = Trim(rstCargaSacII!Accion51)
        Accion52.Text = Trim(rstCargaSacII!Accion52)
        Accion61.Text = Trim(rstCargaSacII!Accion61)
        Accion62.Text = Trim(rstCargaSacII!Accion62)
        
        Responsable1.Text = rstCargaSacII!Responsable1
        Responsable2.Text = rstCargaSacII!Responsable2
        Responsable3.Text = rstCargaSacII!Responsable3
        Responsable4.Text = rstCargaSacII!Responsable4
        Responsable5.Text = rstCargaSacII!Responsable5
        Responsable6.Text = rstCargaSacII!Responsable6
        
        Plazo1.Text = rstCargaSacII!Plazo1
        Plazo2.Text = rstCargaSacII!Plazo2
        Plazo3.Text = rstCargaSacII!Plazo3
        Plazo4.Text = rstCargaSacII!Plazo4
        Plazo5.Text = rstCargaSacII!Plazo5
        Plazo6.Text = rstCargaSacII!Plazo6
        
        rstCargaSacII.Close
    End If
    
    Call Imprime_Descripcion
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub CargaCausas_Click()

    WPasaTipo = Tipo.Text
    WPasaAno = Ano.Text
    WPasaNumero = Numero.Text
    
    PrgCargaSacOtro.Show

End Sub

Private Sub cmdAdd_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacII"
        ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacII.RecordCount > 0 Then
        
            rstCargaSacII.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSacII SET "
            ZSql = ZSql + " Accion11 = " + "'" + Accion11.Text + "',"
            ZSql = ZSql + " Accion12 = " + "'" + Accion12.Text + "',"
            ZSql = ZSql + " Accion21 = " + "'" + Accion21.Text + "',"
            ZSql = ZSql + " Accion22 = " + "'" + Accion22.Text + "',"
            ZSql = ZSql + " Accion31 = " + "'" + Accion31.Text + "',"
            ZSql = ZSql + " Accion32 = " + "'" + Accion32.Text + "',"
            ZSql = ZSql + " Accion41 = " + "'" + Accion41.Text + "',"
            ZSql = ZSql + " Accion42 = " + "'" + Accion42.Text + "',"
            ZSql = ZSql + " Accion51 = " + "'" + Accion51.Text + "',"
            ZSql = ZSql + " Accion52 = " + "'" + Accion52.Text + "',"
            ZSql = ZSql + " Accion61 = " + "'" + Accion61.Text + "',"
            ZSql = ZSql + " Accion62 = " + "'" + Accion62.Text + "',"
            ZSql = ZSql + " Responsable1 = " + "'" + Responsable1.Text + "',"
            ZSql = ZSql + " Responsable2 = " + "'" + Responsable2.Text + "',"
            ZSql = ZSql + " Responsable3 = " + "'" + Responsable3.Text + "',"
            ZSql = ZSql + " Responsable4 = " + "'" + Responsable4.Text + "',"
            ZSql = ZSql + " Responsable5 = " + "'" + Responsable5.Text + "',"
            ZSql = ZSql + " Responsable6 = " + "'" + Responsable6.Text + "',"
            ZSql = ZSql + " Plazo1 = " + "'" + Plazo1.Text + "',"
            ZSql = ZSql + " Plazo2 = " + "'" + Plazo2.Text + "',"
            ZSql = ZSql + " Plazo3 = " + "'" + Plazo3.Text + "',"
            ZSql = ZSql + " Plazo4 = " + "'" + Plazo4.Text + "',"
            ZSql = ZSql + " Plazo5 = " + "'" + Plazo5.Text + "',"
            ZSql = ZSql + " Plazo6 = " + "'" + Plazo6.Text + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSacII = ZSql
            Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSacII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Accion11 ,"
            ZSql = ZSql + "Accion12 ,"
            ZSql = ZSql + "Accion21 ,"
            ZSql = ZSql + "Accion22 ,"
            ZSql = ZSql + "Accion31 ,"
            ZSql = ZSql + "Accion32 ,"
            ZSql = ZSql + "Accion41 ,"
            ZSql = ZSql + "Accion42 ,"
            ZSql = ZSql + "Accion51 ,"
            ZSql = ZSql + "Accion52 ,"
            ZSql = ZSql + "Accion61 ,"
            ZSql = ZSql + "Accion62 ,"
            ZSql = ZSql + "Responsable1 ,"
            ZSql = ZSql + "Responsable2 ,"
            ZSql = ZSql + "Responsable3 ,"
            ZSql = ZSql + "Responsable4 ,"
            ZSql = ZSql + "Responsable5 ,"
            ZSql = ZSql + "Responsable6 ,"
            ZSql = ZSql + "Plazo1 ,"
            ZSql = ZSql + "Plazo2 ,"
            ZSql = ZSql + "Plazo3 ,"
            ZSql = ZSql + "Plazo4 ,"
            ZSql = ZSql + "Plazo5 ,"
            ZSql = ZSql + "Plazo6 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Tipo.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Accion11.Text + "',"
            ZSql = ZSql + "'" + Accion12.Text + "',"
            ZSql = ZSql + "'" + Accion21.Text + "',"
            ZSql = ZSql + "'" + Accion22.Text + "',"
            ZSql = ZSql + "'" + Accion31.Text + "',"
            ZSql = ZSql + "'" + Accion32.Text + "',"
            ZSql = ZSql + "'" + Accion41.Text + "',"
            ZSql = ZSql + "'" + Accion42.Text + "',"
            ZSql = ZSql + "'" + Accion51.Text + "',"
            ZSql = ZSql + "'" + Accion52.Text + "',"
            ZSql = ZSql + "'" + Accion61.Text + "',"
            ZSql = ZSql + "'" + Accion62.Text + "',"
            ZSql = ZSql + "'" + Responsable1.Text + "',"
            ZSql = ZSql + "'" + Responsable2.Text + "',"
            ZSql = ZSql + "'" + Responsable3.Text + "',"
            ZSql = ZSql + "'" + Responsable4.Text + "',"
            ZSql = ZSql + "'" + Responsable5.Text + "',"
            ZSql = ZSql + "'" + Responsable6.Text + "',"
            ZSql = ZSql + "'" + Plazo1.Text + "',"
            ZSql = ZSql + "'" + Plazo2.Text + "',"
            ZSql = ZSql + "'" + Plazo3.Text + "',"
            ZSql = ZSql + "'" + Plazo4.Text + "',"
            ZSql = ZSql + "'" + Plazo5.Text + "',"
            ZSql = ZSql + "'" + Plazo6.Text + "')"
            
            spCargaSacII = ZSql
            Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
        T$ = "Carga de Acciones Correctivas"
        m$ = "Desea enviar el aviso al Responsable de Calidad"
        ZRespuesta% = MsgBox(m$, 32 + 4, T$)
        If ZRespuesta% = 6 Then
        
            ZZEmail = "ebiglieri@surfactan.com.ar; calidad@surfactan.com.ar"
            
            sTo = ZZEmail
            sCC = ""
            sBCC = ""
            Select Case Val(Tipo.Text)
                Case 1
                    sSubject = "Acciones Correctivas a cumplir"
                    sBody = "Se ingresaron accciones correctivas de " + _
                            DesTipo.Caption + " : " + _
                            Ano.Text + "/" + Numero.Text + _
                            " para su implementacion. " + _
                            "Referencia : " + Referencia.Text + _
                            "Titulo : " + Titulo.Text
                Case 2
                    sSubject = "Acciones Preventivas a cumplir"
                    sBody = "Se ingresaron accciones preventivas de " + _
                            DesTipo.Caption + " : " + _
                            Ano.Text + "/" + Numero.Text + _
                            " para su implementacion. " + _
                            "Referencia : " + Referencia.Text + _
                            "Titulo : " + Titulo.Text
                Case Else
                    sSubject = "Acciones de " + DesTipo.Caption + " a cumplir"
                    sBody = "Se ingresaron accciones de " + _
                            DesTipo.Caption + " : " + _
                            Ano.Text + "/" + Numero.Text + _
                            " para su implementacion. " + _
                            "Referencia : " + Referencia.Text + _
                            "Titulo : " + Titulo.Text
            End Select

            ret = Shell("Start.exe " _
                        & "mailto:" & """" & sTo & """" _
                        & "?Subject=" & """" & sSubject & """" _
                        & "&cc=" & """" & sCC & """" _
                        & "&bcc=" & """" & sBCC & """" _
                        & "&Body=" & """" & sBody & """" _
                        & "&File=" & """" & "c:\autoexec.bat" & """" _
                        , 0)
        End If
        
        
        
        
        
        T$ = "Carga de Acciones Correctivas"
        m$ = "Desea enviar el aviso a los responsables de accion correctiva"
        ZRespuesta% = MsgBox(m$, 32 + 4, T$)
        If ZRespuesta% = 6 Then
        
            For Ciclo = 1 To 100
            
                If Ciclo = Val(Responsable1.Text) Or Ciclo = Val(Responsable2.Text) Or Ciclo = Val(Responsable3.Text) Or Ciclo = Val(Responsable4.Text) Or Ciclo = Val(Responsable5.Text) Or Ciclo = Val(Responsable6.Text) Then
            
                    ZZResponsable = Ciclo
        
                    ZZEmail = ""
                
                    Sql1 = "Select *"
                    Sql2 = " FROM ResponsableSac"
                    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsable) + "'"
                    spResponsableSac = Sql1 + Sql2 + Sql3
                    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                    If rstResponsableSac.RecordCount > 0 Then
                        ZZEmail = Trim(rstResponsableSac!Email)
                        rstResponsableSac.Close
                    End If
                
                    If ZZEmail <> "" Then
            
                        sTo = ZZEmail
                        sCC = ""
                        sBCC = ""
                        sSubject = "Aviso de Acciones Correctivas"
                        sBody = "Se ingresaron accciones corectivas del " + DesTipo.Caption + " : " + Ano.Text + "/" + Numero.Text + " para su implemenatacion    Referencia : " + Referencia.Text
    
                        ret = Shell("Start.exe " _
                                & "mailto:" & """" & sTo & """" _
                                & "?Subject=" & """" & sSubject & """" _
                                & "&cc=" & """" & sCC & """" _
                                & "&bcc=" & """" & sBCC & """" _
                                & "&Body=" & """" & sBody & """" _
                                & "&File=" & """" & "c:\autoexec.bat" & """" _
                                , 0)
                    End If
                    
                End If
                
            Next Ciclo
            
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            rstCargaSac.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSac SET "
            ZSql = ZSql + " Centro = " + "'" + Centro.Text + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Origen = " + "'" + Str$(Origen.ListIndex) + "',"
            ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
            ZSql = ZSql + " ResponsableEmisor = " + "'" + ResponsableEmisor.Text + "',"
            ZSql = ZSql + " ResponsableDestino = " + "'" + ResponsableDestino.Text + "',"
            ZSql = ZSql + " Referencia = " + "'" + Referencia.Text + "',"
            ZSql = ZSql + " Titulo = " + "'" + Titulo.Text + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            ZZEstado = rstCargaSac!Estado
            rstCargaSac.Close
            
            If ZZEstado <= 2 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE CargaSac SET "
                ZSql = ZSql + " Estado = " + "'" + "3" + "'"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        End If
        
        
        
        
            
        Call CmdLimpiar_Click
        Tipo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacII"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacII.RecordCount > 0 Then
        
            rstCargaSacII.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                Sql1 = ZSql + "DELETE CargaSacII"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSacII = Sql1 + Sql2
                Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    End If
    
    Tipo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    
    
    Origen.ListIndex = 0
    Estado.ListIndex = 0
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    
    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Tipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgCargaSacAccion.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    
    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "INICIADA"
    Estado.AddItem "INVESTIGACION"
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "IMPLEMENTACION A VERIFICAR"
    Estado.AddItem "IMPLEMENTACION VERIFICADA"
    Estado.AddItem "CERRADA"
    Estado.AddItem "ANULADA"
    
    Estado.ListIndex = 0
    
    Origen.Clear
    
    Origen.AddItem ""
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            DesTipo.Caption = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
            
                ZResponsableDestino = rstCargaSac!ResponsableDestino
                ZResponsableCentro = 9999
                ZCentro = rstCargaSac!Centro
                
                rstCargaSac.Close
                
                Sql1 = "Select *"
                Sql2 = " FROM CentroSac"
                Sql3 = " Where CentroSac.Codigo = " + "'" + Str$(ZCentro) + "'"
                spCentroSac = Sql1 + Sql2 + Sql3
                Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCentroSac.RecordCount > 0 Then
                    ZResponsableCentro = rstCentroSac!Responsable
                    rstCentroSac.Close
                End If
                
                If WOperador = ZResponsableDestino Or WOperador = ZResponsableCentro Or ZZCodigoResponsable = 1 Then
                    Call Imprime_Datos
                    Accion11.SetFocus
                        Else
                    m$ = "No posee autorizacion para ingresar a actualizar esta SAC"
                    A% = MsgBox(m$, 0, "Archivo de Carga de Accion")
                    Call CmdLimpiar_Click
                End If
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Tipo.Text = WTipo
                Ano.Text = WAno
                Numero.Text = WNumero
                Sql1 = "Select *"
                Sql2 = " FROM TipoSac"
                Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
                spTipoSac = Sql1 + Sql2 + Sql3
                Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstTipoSac.RecordCount > 0 Then
                    DesTipo.Caption = Trim(rstTipoSac!Descripcion)
                    rstTipoSac.Close
                    Ano.SetFocus
                End If
                
                Tipo.SetFocus
                
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub





Private Sub Accion11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion12.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion11.Text = ""
    End If
End Sub

Private Sub Accion12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable1.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion12.Text = ""
    End If
End Sub

Private Sub Responsable1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable1.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo1.SetFocus
            End If
                Else
            DesResponsable1.Caption = ""
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable1.Text = ""
        DesResponsable1.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo1.Text, Auxi)
        If Auxi = "S" Or Plazo1.Text = "  /  /    " Then
            Accion21.SetFocus
                Else
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo1.Text = "  /  /    "
    End If
End Sub








Private Sub Accion21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion22.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion21.Text = ""
    End If
End Sub

Private Sub Accion22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable2.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion22.Text = ""
    End If
End Sub

Private Sub Responsable2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable2.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo2.SetFocus
            End If
                Else
            DesResponsable2.Caption = ""
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable2.Text = ""
        DesResponsable2.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo2.Text, Auxi)
        If Auxi = "S" Or Plazo2.Text = "  /  /    " Then
            Accion31.SetFocus
                Else
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo2.Text = "  /  /    "
    End If
End Sub







Private Sub Accion31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion32.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion31.Text = ""
    End If
End Sub

Private Sub Accion32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable3.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion32.Text = ""
    End If
End Sub

Private Sub Responsable3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable3.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo3.SetFocus
            End If
                Else
            DesResponsable3.Caption = ""
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable3.Text = ""
        DesResponsable3.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo3.Text, Auxi)
        If Auxi = "S" Or Plazo3.Text = "  /  /    " Then
            Accion41.SetFocus
                Else
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo3.Text = "  /  /    "
    End If
End Sub







Private Sub Accion41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion42.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion41.Text = ""
    End If
End Sub

Private Sub Accion42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable4.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion42.Text = ""
    End If
End Sub

Private Sub Responsable4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable4.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo4.SetFocus
            End If
                Else
            DesResponsable4.Caption = ""
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable4.Text = ""
        DesResponsable4.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo4.Text, Auxi)
        If Auxi = "S" Or Plazo4.Text = "  /  /    " Then
            Accion51.SetFocus
                Else
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo4.Text = "  /  /    "
    End If
End Sub







Private Sub Accion51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion52.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion51.Text = ""
    End If
End Sub

Private Sub Accion52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable5.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion52.Text = ""
    End If
End Sub

Private Sub Responsable5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable5.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo5.SetFocus
            End If
                Else
            DesResponsable5.Caption = ""
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable5.Text = ""
        DesResponsable5.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo5.Text, Auxi)
        If Auxi = "S" Or Plazo5.Text = "  /  /    " Then
            Accion61.SetFocus
                Else
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo5.Text = "  /  /    "
    End If
End Sub







Private Sub Accion61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion62.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion61.Text = ""
    End If
End Sub

Private Sub Accion62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable6.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion62.Text = ""
    End If
End Sub

Private Sub Responsable6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable6.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo6.SetFocus
            End If
                Else
            DesResponsable6.Caption = ""
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable6.Text = ""
        DesResponsable6.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo6.Text, Auxi)
        If Auxi = "S" Or Plazo6.Text = "  /  /    " Then
            Accion11.SetFocus
                Else
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo6.Text = "  /  /    "
    End If
End Sub






Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Tipo"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"

     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM tiposac"
            Sql3 = " Order by tiposac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTipoSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(Indice)
            Call Tipo_Keypress(13)
            
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Select Case ZZLugar
                Case 1
                    Responsable1.Text = WIndice.List(Indice)
                    Call Responsable1_Keypress(13)
                Case 2
                    Responsable2.Text = WIndice.List(Indice)
                    Call Responsable2_Keypress(13)
                Case 3
                    Responsable3.Text = WIndice.List(Indice)
                    Call Responsable3_Keypress(13)
                Case 4
                    Responsable4.Text = WIndice.List(Indice)
                    Call Responsable4_Keypress(13)
                Case 5
                    Responsable5.Text = WIndice.List(Indice)
                    Call Responsable5_Keypress(13)
                Case 6
                    Responsable6.Text = WIndice.List(Indice)
                    Call Responsable6_Keypress(13)
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    LugarAyuda = 0
    WIndice.Clear
    Pantalla.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM TipoSac"
            Sql3 = " Order by TipoSac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTipoSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTipoSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTipoSac!Codigo
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
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstResponsableSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstResponsableSac!Codigo
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
                rstResponsableSac.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Tipo_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable1_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable2_DblClick()

    ZZLugar = 2

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable3_DblClick()

    ZZLugar = 3

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable4_DblClick()

    ZZLugar = 4

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable5_DblClick()

    ZZLugar = 5

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable6_DblClick()

    ZZLugar = 6

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

