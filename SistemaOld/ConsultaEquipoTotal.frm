VERSION 5.00
Begin VB.Form PrgConsultaEquipoTotal 
   AutoRedraw      =   -1  'True
   Caption         =   "CONSULTA DE ESTADO DE REACTORES"
   ClientHeight    =   7320
   ClientLeft      =   90
   ClientTop       =   690
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11850
   Begin VB.TextBox HojaI 
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
      Height          =   405
      Left            =   600
      TabIndex        =   127
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox HojaII 
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
      Height          =   405
      Left            =   600
      TabIndex        =   126
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox HojaIII 
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
      Height          =   405
      Left            =   600
      TabIndex        =   125
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox HojaIV 
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
      Height          =   405
      Left            =   600
      TabIndex        =   124
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox HojaV 
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
      Height          =   405
      Left            =   600
      TabIndex        =   123
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox HojaVI 
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
      Height          =   405
      Left            =   600
      TabIndex        =   122
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox HojaVII 
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
      Height          =   405
      Left            =   600
      TabIndex        =   121
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox HojaVIII 
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
      Height          =   405
      Left            =   600
      TabIndex        =   120
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox HojaIX 
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
      Height          =   405
      Left            =   600
      TabIndex        =   119
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox HojaX 
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
      Height          =   405
      Left            =   600
      TabIndex        =   118
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox EtapaVII 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   117
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox TempeV 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   116
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Alerta3III 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   115
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox EquipoVI 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   114
      Text            =   "VI"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Alerta3X 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   113
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox Alerta2X 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   112
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox Alerta1X 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   111
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox TempeX 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   110
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox InicioX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   109
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox EtapaX 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   108
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox HoraX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   107
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox CantidadX 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   106
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox ProductoX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   105
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox OperadorX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   104
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Alerta3IX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   103
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox Alerta2IX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   102
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox Alerta1IX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   101
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox TempeIX 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   100
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox InicioIX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   99
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox EtapaIX 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   98
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox HoraIX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   97
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox CantidadIX 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   96
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox ProductoIX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   95
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox OperadorIX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   94
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Alerta3VIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   93
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Alerta2VIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   92
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Alerta1VIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   91
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox TempeVIII 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   90
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox InicioVIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   89
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox EtapaVIII 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   88
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox HoraVIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   87
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox CantidadVIII 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   86
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox ProductoVIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   85
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox OperadorVIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   84
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Alerta3VII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   83
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Alerta2VII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   82
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Alerta1VII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   81
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox TempeVII 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   80
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox InicioVII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   79
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox HoraVII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   78
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox CantidadVII 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   77
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox ProductoVII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   76
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox OperadorVII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   75
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Alerta3VI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   74
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox Alerta2VI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   73
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox Alerta1VI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   72
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox TempeVI 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   71
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox InicioVI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   70
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox EtapaVI 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   69
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox HoraVI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   68
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox CantidadVI 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   67
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox ProductoVI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   66
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox OperadorVI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   65
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Alerta3V 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   64
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Alerta2V 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   63
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Alerta1V 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   62
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox InicioV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   61
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox EtapaV 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   60
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox HoraV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   59
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox CantidadV 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   58
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox ProductoV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   57
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox OperadorV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   56
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Alerta3IV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   55
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Alerta2IV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   54
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Alerta1IV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   53
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TempeIV 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   52
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox InicioIV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   51
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox EtapaIV 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   50
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox HoraIV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   49
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox CantidadIV 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   48
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox ProductoIV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   47
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox OperadorIV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   46
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Alerta2III 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   45
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Alerta1III 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   44
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox TempeIII 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   43
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox InicioIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   42
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox EtapaIII 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   41
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox HoraIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   40
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox CantidadIII 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   39
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox ProductoIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   38
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox OperadorIII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   37
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Alerta3II 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   36
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Alerta2II 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   35
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Alerta1II 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   34
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox TempeII 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   33
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox InicioII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   32
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox EtapaII 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   31
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox HoraII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   30
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox CantidadII 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   29
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox ProductoII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   28
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox OperadorII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   27
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Alerta3I 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   26
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Alerta2I 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10920
      TabIndex        =   25
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Alerta1I 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   24
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox TempeI 
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
      Height          =   405
      Left            =   9600
      TabIndex        =   23
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox InicioI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox EtapaI 
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
      Height          =   405
      Left            =   7320
      TabIndex        =   21
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox HoraI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox CantidadI 
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
      Height          =   405
      Left            =   5040
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox ProductoI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox EquipoX 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   17
      Text            =   "X"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox EquipoIX 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Text            =   "IX"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox EquipoVIII 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   15
      Text            =   "VIII"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox EquipoVII 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Text            =   "VII"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox EquipoV 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   13
      Text            =   "V"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox EquipoIV 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Text            =   "IV"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox EquipoIII 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Text            =   "III"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox EquipoII 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Text            =   "II"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox OperadorI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox EquipoI 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Text            =   "I"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HOJA"
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
      Left            =   600
      TabIndex        =   128
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALERTAS"
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
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TEMP."
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
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HORA INICIO"
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
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ETAPA"
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
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HORA INICIO"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANT."
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
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTO"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPERADOR"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "ConsultaEquipoTotal.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaEquipoTotal.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaEquipoTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim XParam As String
Dim WGraba As String
Dim ZVector(100, 8) As String
Dim ZOpera(1000) As String
Dim XEmpresa As String

Private Sub cmdClose_Click()
    PrgConsultaEquipo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    
    HojaI.Text = ""
    HojaII.Text = ""
    HojaIII.Text = ""
    HojaIV.Text = ""
    HojaV.Text = ""
    HojaVI.Text = ""
    HojaVII.Text = ""
    HojaVIII.Text = ""
    HojaIX.Text = ""
    HojaX.Text = ""
    
    OperadorI.Text = ""
    OperadorII.Text = ""
    OperadorIII.Text = ""
    OperadorIV.Text = ""
    OperadorV.Text = ""
    OperadorVI.Text = ""
    OperadorVII.Text = ""
    OperadorVIII.Text = ""
    OperadorIX.Text = ""
    OperadorX.Text = ""
    
    ProductoI.Text = ""
    ProductoII.Text = ""
    ProductoIII.Text = ""
    ProductoIV.Text = ""
    ProductoV.Text = ""
    ProductoVI.Text = ""
    ProductoVII.Text = ""
    ProductoVIII.Text = ""
    ProductoIX.Text = ""
    ProductoX.Text = ""
    
    CantidadI.Text = ""
    CantidadII.Text = ""
    CantidadIII.Text = ""
    CantidadIV.Text = ""
    CantidadV.Text = ""
    CantidadVI.Text = ""
    CantidadVII.Text = ""
    CantidadVIII.Text = ""
    CantidadIX.Text = ""
    CantidadX.Text = ""
    
    HoraI.Text = ""
    HoraII.Text = ""
    HoraIII.Text = ""
    HoraIV.Text = ""
    HoraV.Text = ""
    HoraVI.Text = ""
    HoraVII.Text = ""
    HoraVIII.Text = ""
    HoraIX.Text = ""
    HoraX.Text = ""
    
    
    EtapaI.Text = ""
    EtapaII.Text = ""
    EtapaIII.Text = ""
    EtapaIV.Text = ""
    EtapaV.Text = ""
    EtapaVI.Text = ""
    EtapaVII.Text = ""
    EtapaVIII.Text = ""
    EtapaIX.Text = ""
    EtapaX.Text = ""
    
    InicioI.Text = ""
    InicioII.Text = ""
    InicioIII.Text = ""
    InicioIV.Text = ""
    InicioV.Text = ""
    InicioVI.Text = ""
    InicioVII.Text = ""
    InicioVIII.Text = ""
    InicioIX.Text = ""
    InicioX.Text = ""
    
    TempeI.Text = ""
    TempeII.Text = ""
    TempeIII.Text = ""
    TempeIV.Text = ""
    TempeV.Text = ""
    TempeVI.Text = ""
    TempeVII.Text = ""
    TempeVIII.Text = ""
    TempeIX.Text = ""
    TempeX.Text = ""
    
    Alerta1I.Text = ""
    Alerta1II.Text = ""
    Alerta1III.Text = ""
    Alerta1IV.Text = ""
    Alerta1V.Text = ""
    Alerta1VI.Text = ""
    Alerta1VII.Text = ""
    Alerta1VIII.Text = ""
    Alerta1IX.Text = ""
    Alerta1X.Text = ""
    
    Alerta2I.Text = ""
    Alerta2II.Text = ""
    Alerta2III.Text = ""
    Alerta2IV.Text = ""
    Alerta2V.Text = ""
    Alerta2VI.Text = ""
    Alerta2VII.Text = ""
    Alerta2VIII.Text = ""
    Alerta2IX.Text = ""
    Alerta2X.Text = ""
    
    Alerta3I.Text = ""
    Alerta3II.Text = ""
    Alerta3III.Text = ""
    Alerta3IV.Text = ""
    Alerta3V.Text = ""
    Alerta3VI.Text = ""
    Alerta3VII.Text = ""
    Alerta3VIII.Text = ""
    Alerta3IX.Text = ""
    Alerta3X.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.EstadoHoja = 1 and Hoja.Renglon = 1"
    ZSql = ZSql + " Order by Hoja.Hoja"
            
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZEquipo = Trim(rstHoja!Equipo)
                
                    Select Case ZEquipo
                        Case "I"
                            HojaI.Text = Str$(rstHoja!Hoja)
                            OperadorI.Text = ""
                            ProductoI.Text = rstHoja!Producto
                            CantidadI.Text = Str$(rstHoja!Teorico)
                            HoraI.Text = ""
                            EtapaI.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioI.Text = rstHoja!HoraInicioEtapa
                            TempeI.Text = ""
                            Alerta1I.Text = rstHoja!alarma
                            Alerta2I.Text = ""
                            Alerta3I.Text = ""
                        
                            ZOpera(1) = rstHoja!Operario
                            
                        Case "II"
                            HojaII.Text = Str$(rstHoja!Hoja)
                            OperadorII.Text = ""
                            ProductoII.Text = rstHoja!Producto
                            CantidadII.Text = Str$(rstHoja!Teorico)
                            HoraII.Text = ""
                            EtapaII.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioII.Text = rstHoja!HoraInicioEtapa
                            TempeII.Text = ""
                            Alerta1II.Text = rstHoja!alarma
                            Alerta2II.Text = ""
                            Alerta3II.Text = ""
                        
                            ZOpera(2) = rstHoja!Operario
                            
                        Case "III"
                            HojaIII.Text = Str$(rstHoja!Hoja)
                            OperadorIII.Text = ""
                            ProductoIII.Text = rstHoja!Producto
                            CantidadIII.Text = Str$(rstHoja!Teorico)
                            HoraIII.Text = ""
                            EtapaIII.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioIII.Text = rstHoja!HoraInicioEtapa
                            TempeIII.Text = ""
                            Alerta1III.Text = rstHoja!alarma
                            Alerta2III.Text = ""
                            Alerta3III.Text = ""
                        
                            ZOpera(3) = rstHoja!Operario
                            
                        Case "IV"
                            HojaIV.Text = Str$(rstHoja!Hoja)
                            OperadorIV.Text = ""
                            ProductoIV.Text = rstHoja!Producto
                            CantidadIV.Text = Str$(rstHoja!Teorico)
                            HoraIV.Text = ""
                            EtapaIV.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioIV.Text = rstHoja!HoraInicioEtapa
                            TempeIV.Text = ""
                            Alerta1IV.Text = rstHoja!alarma
                            Alerta2IV.Text = ""
                            Alerta3IV.Text = ""
                        
                            ZOpera(4) = rstHoja!Operario
                            
                        Case "V"
                            HojaV.Text = Str$(rstHoja!Hoja)
                            OperadorV.Text = ""
                            ProductoV.Text = rstHoja!Producto
                            CantidadV.Text = Str$(rstHoja!Teorico)
                            HoraV.Text = ""
                            EtapaV.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioV.Text = rstHoja!HoraInicioEtapa
                            TempeV.Text = ""
                            Alerta1V.Text = rstHoja!alarma
                            Alerta2V.Text = ""
                            Alerta3V.Text = ""
                        
                            ZOpera(5) = rstHoja!Operario
                            
                        Case "VI"
                            HojaVI.Text = Str$(rstHoja!Hoja)
                            OperadorVI.Text = ""
                            ProductoVI.Text = rstHoja!Producto
                            CantidadVI.Text = Str$(rstHoja!Teorico)
                            HoraVI.Text = ""
                            EtapaVI.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioVI.Text = rstHoja!HoraInicioEtapa
                            TempeVI.Text = ""
                            Alerta1VI.Text = rstHoja!alarma
                            Alerta2VI.Text = ""
                            Alerta3VI.Text = ""
                        
                            ZOpera(6) = rstHoja!Operario
                            
                        Case "VII"
                            HojaVII.Text = Str$(rstHoja!Hoja)
                            OperadorVII.Text = ""
                            ProductoVII.Text = rstHoja!Producto
                            CantidadVII.Text = Str$(rstHoja!Teorico)
                            HoraVII.Text = ""
                            EtapaVII.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioVII.Text = rstHoja!HoraInicioEtapa
                            TempeVII.Text = ""
                            Alerta1VII.Text = rstHoja!alarma
                            Alerta2VII.Text = ""
                            Alerta3VII.Text = ""
                        
                            ZOpera(7) = rstHoja!Operario
                            
                        Case "VIII"
                            HojaVIII.Text = Str$(rstHoja!Hoja)
                            OperadorVIII.Text = ""
                            ProductoVIII.Text = rstHoja!Producto
                            CantidadVIII.Text = Str$(rstHoja!Teorico)
                            HoraVIII.Text = ""
                            EtapaVIII.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioVIII.Text = rstHoja!HoraInicioEtapa
                            TempeVIII.Text = ""
                            Alerta1VIII.Text = rstHoja!alarma
                            Alerta2VIII.Text = ""
                            Alerta3VIII.Text = ""
                        
                            ZOpera(8) = rstHoja!Operario
                            
                        Case "IX"
                            HojaIX.Text = Str$(rstHoja!Hoja)
                            OperadorIX.Text = ""
                            ProductoIX.Text = rstHoja!Producto
                            CantidadIX.Text = Str$(rstHoja!Teorico)
                            HoraIX.Text = ""
                            EtapaIX.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioIX.Text = rstHoja!HoraInicioEtapa
                            TempeIX.Text = ""
                            Alerta1IX.Text = rstHoja!alarma
                            Alerta2IX.Text = ""
                            Alerta3IX.Text = ""
                        
                            ZOpera(9) = rstHoja!Operario
                            
                        Case "X"
                            HojaX.Text = Str$(rstHoja!Hoja)
                            OperadorX.Text = ""
                            ProductoX.Text = rstHoja!Producto
                            CantidadX.Text = Str$(rstHoja!Teorico)
                            HoraX.Text = ""
                            EtapaX.Text = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                            InicioX.Text = rstHoja!HoraInicioEtapa
                            TempeX.Text = ""
                            Alerta1X.Text = rstHoja!alarma
                            Alerta2X.Text = ""
                            Alerta3X.Text = ""
                        
                            ZOpera(10) = rstHoja!Operario
                            
                        Case Else
                        
                    End Select
                            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To 10
    
        Sql1 = "Select *"
        Sql2 = " FROM Operarios"
        Sql3 = " Where Operarios.Codigo = " + "'" + ZOpera(Ciclo) + "'"
        spOperarios = Sql1 + Sql2 + Sql3
        Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperarios.RecordCount > 0 Then
            Select Case Ciclo
                Case 1
                    OperadorI.Text = rstOperarios!Descripcion
                Case 2
                    OperadorII.Text = rstOperarios!Descripcion
                Case 3
                    OperadorIII.Text = rstOperarios!Descripcion
                Case 4
                    OperadorIV.Text = rstOperarios!Descripcion
                Case 5
                    OperadorV.Text = rstOperarios!Descripcion
                Case 6
                    OperadorVI.Text = rstOperarios!Descripcion
                Case 7
                    OperadorVII.Text = rstOperarios!Descripcion
                Case 8
                    OperadorVIII.Text = rstOperarios!Descripcion
                Case 9
                    OperadorIX.Text = rstOperarios!Descripcion
                Case 10
                    OperadorX.Text = rstOperarios!Descripcion
                Case Else
            End Select
            rstOperarios.Close
        End If
        
    Next Ciclo
    
End Sub

