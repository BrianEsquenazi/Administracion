VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgInformeConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Informe de Recepcion"
   ClientHeight    =   8160
   ClientLeft      =   75
   ClientTop       =   540
   ClientWidth     =   11835
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10200
      TabIndex        =   107
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame PantaEnvase 
      Caption         =   "Ingreso de Estado de Envases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   480
      TabIndex        =   66
      Top             =   1200
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox observa 
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
         MaxLength       =   50
         TabIndex        =   105
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox ObservaIV 
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
         MaxLength       =   50
         TabIndex        =   92
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox ObservaIII 
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
         MaxLength       =   50
         TabIndex        =   91
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox ObservaII 
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
         MaxLength       =   50
         TabIndex        =   90
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox ObservaI 
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
         MaxLength       =   50
         TabIndex        =   89
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox CantidadEnv 
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
         Left            =   3840
         MaxLength       =   6
         TabIndex        =   83
         Text            =   " "
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CheckBox EstadoEnvX 
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
         Left            =   4800
         MaskColor       =   &H00FF0000&
         TabIndex        =   81
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox EstadoEnvIX 
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         TabIndex        =   80
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox EstadoEnvVIII 
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
         Left            =   4800
         MaskColor       =   &H00FF0000&
         TabIndex        =   78
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox EstadoEnvVII 
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         TabIndex        =   77
         Top             =   2520
         Width           =   495
      End
      Begin VB.CheckBox EstadoEnvVI 
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
         Left            =   4800
         MaskColor       =   &H00FF0000&
         TabIndex        =   75
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox EstadoEnvV 
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         TabIndex        =   74
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox EstadoEnvII 
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
         Left            =   4800
         MaskColor       =   &H00FF0000&
         TabIndex        =   71
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox EstadoEnvI 
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         TabIndex        =   70
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox EstadoEnvIV 
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
         Left            =   4800
         MaskColor       =   &H00FF0000&
         TabIndex        =   69
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox EstadoEnvIII 
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         TabIndex        =   68
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton ConfirmaPantaEnvase 
         Caption         =   "Acepta"
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
         Left            =   7560
         TabIndex        =   67
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "No Cumple"
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
         Left            =   4560
         TabIndex        =   88
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Cumple"
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
         Left            =   3600
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Cantidad Rechazada"
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
         TabIndex        =   84
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label DescriEnsayo5 
         Caption         =   "Estado Envases"
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
         Height          =   525
         Left            =   360
         TabIndex        =   82
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label DescriEnsayo4 
         Caption         =   "Estado Envases"
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
         Height          =   525
         Left            =   360
         TabIndex        =   79
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label DescriEnsayo3 
         Caption         =   "Estado Envases"
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
         Height          =   525
         Left            =   360
         TabIndex        =   76
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label DescriEnsayo1 
         Caption         =   "Certif.de Analisis"
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
         Height          =   525
         Left            =   360
         TabIndex        =   73
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label DescriEnsayo2 
         Caption         =   "Estado Envases"
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
         Height          =   525
         Left            =   360
         TabIndex        =   72
         Top             =   1320
         Width           =   3135
      End
   End
   Begin VB.Frame PantaCertificado 
      Caption         =   "Ingreso de Certificado de Analisis y Estado de Envases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   1920
      TabIndex        =   53
      Top             =   1440
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox PartidaProveedor 
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
         MaxLength       =   30
         TabIndex        =   103
         Top             =   2880
         Visible         =   0   'False
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   101
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox Procedencia 
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
         MaxLength       =   30
         TabIndex        =   97
         Top             =   1920
         Width           =   3975
      End
      Begin VB.CommandButton ConfirmaCertificado 
         Caption         =   "Acepta"
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
         TabIndex        =   62
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox EstadoSi 
         Caption         =   "Si"
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
         Left            =   1920
         MaskColor       =   &H00FF0000&
         TabIndex        =   60
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Estado2 
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
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   59
         Text            =   " "
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox EstadoNo 
         Caption         =   "No"
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
         Left            =   2520
         MaskColor       =   &H00FF0000&
         TabIndex        =   58
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox CertificadoSi 
         Caption         =   "Si"
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
         Left            =   1920
         MaskColor       =   &H00FF0000&
         TabIndex        =   56
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Certificado2 
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
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   55
         Text            =   " "
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox CertificadoNo 
         Caption         =   "No"
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
         Left            =   2520
         MaskColor       =   &H00FF0000&
         TabIndex        =   54
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox Vencimiento 
         Height          =   285
         Left            =   1920
         TabIndex        =   85
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label26 
         Caption         =   "Partida Proveedor"
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
         TabIndex        =   104
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label24 
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
         Left            =   240
         TabIndex        =   102
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Made in "
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
         TabIndex        =   98
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Vencimiento"
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
         TabIndex        =   86
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Estado Envases"
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
         TabIndex        =   61
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Certif.de Analisis"
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
         TabIndex        =   57
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Impre Control Recepcion"
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
      Left            =   2400
      TabIndex        =   106
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton DatosAdicio 
      Caption         =   "Datos Despacho"
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
      Left            =   1200
      TabIndex        =   96
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Frame PantaDatosAdicionales 
      Height          =   1695
      Left            =   1560
      TabIndex        =   94
      Top             =   2760
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox NroDespacho 
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
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   99
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton DatosAdicionalesCancela 
         Caption         =   "Acepta"
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
         Left            =   3120
         TabIndex        =   95
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Nro. Despacho"
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
         Left            =   840
         TabIndex        =   100
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CheckList 
      Caption         =   "Chek-List"
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
      Left            =   120
      TabIndex        =   93
      Top             =   7440
      Width           =   975
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
      Index           =   9
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede actualizar el informe de recepcion. Sistema sin Conexion con las otras plantas"
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
      Left            =   3600
      Picture         =   "informeconsulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
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
      Index           =   8
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2520
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
      Index           =   7
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   2880
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
      Index           =   6
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   2160
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2160
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
      Index           =   2
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   1800
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   1800
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
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2160
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
      Index           =   5
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Aviso 
      Height          =   2535
      Left            =   2640
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton CierreAviso 
         Caption         =   "CIERRE"
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
         Left            =   2400
         TabIndex        =   42
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Aviso3 
         Alignment       =   2  'Center
         Caption         =   "DEBERA CUMPLIR CON LA ENTREGA DE LOS MISMOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   6015
      End
      Begin VB.Label Aviso2 
         Alignment       =   2  'Center
         Caption         =   "LA CANTIDAD DE 10 KGS. Y QUE EL PROVEEDOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Aviso1 
         Alignment       =   2  'Center
         Caption         =   "A CONFIRMADO QUE DEJA PENDIENTE DE RECEPCION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label13 
         Caption         =   "ATENCION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox XOrden 
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
      Left            =   9120
      MaxLength       =   6
      TabIndex        =   34
      Top             =   120
      Width           =   1095
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
      Left            =   3840
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.TextBox Remito 
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
      Left            =   9120
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
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
      MaxLength       =   11
      TabIndex        =   20
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
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
      Left            =   2280
      TabIndex        =   18
      Top             =   6840
      Width           =   975
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
      Left            =   4440
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Informe 
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
      Left            =   120
      TabIndex        =   11
      Top             =   6240
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
      Left            =   1200
      TabIndex        =   10
      Top             =   6840
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
      Left            =   2280
      TabIndex        =   8
      Top             =   6240
      Width           =   975
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
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   11415
      Begin VB.TextBox WEtiqueta 
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
         Left            =   10200
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox WEnvase 
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
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox WResta 
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
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   1095
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
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   23
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WOrden 
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
         Left            =   360
         MaxLength       =   6
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   600
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   10200
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         Left            =   8400
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desc. O/C"
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
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo O/C"
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
         Left            =   6240
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant.Ing."
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
         Left            =   5160
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WSaldo 
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
         Height          =   300
         Left            =   6240
         TabIndex        =   24
         Top             =   600
         Width           =   1095
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
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   2535
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
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
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
      Height          =   1425
      ItemData        =   "informeconsulta.frx":0742
      Left            =   3840
      List            =   "informeconsulta.frx":0749
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   6735
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
      Height          =   500
      Left            =   1200
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WGrilla 
      Height          =   3615
      Left            =   120
      TabIndex        =   49
      Top             =   1320
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6376
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label11 
      Caption         =   "Orden Compra"
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
      Left            =   7560
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   7560
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label DesProveedor 
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
      Left            =   3360
      TabIndex        =   16
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   120
      TabIndex        =   15
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
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Informe"
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
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgInformeConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private WSaldoUs As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double
Dim rstRecibosProvi As Recordset
Dim spRecibosProvi As String
Dim rstCuit As Recordset
Dim spCuit As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstCambio As Recordset
Dim spCambio As String
Dim XParam As String
Dim XParidad As String
Dim WParidad As String
Dim Pari As Double
Dim WEntra(100, 120) As String
Dim ZPasa As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim ZBancos(1000) As String
Dim XTipo1 As String
Dim XNumero1 As String
Dim ZClaveCheque(100, 10) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String
Dim WControlII As String

Dim ZZDa As Integer

Private Sub Suma_Datos()

    Rem If Val(WProv) = 24 Then
    Rem     Paridad.Text = "1"
    Rem End If

    Debitos.Caption = ""
    Creditos.Caption = ""
    ZPasa = "S"
    
    Creditos.Caption = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text) + Val(RetSuss.Text))
    For iRow = 1 To 99
    
        If Val(WVector1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(WVector1.TextMatrix(iRow, 10)))
        End If
        
        ZTipo = Val(WVector1.TextMatrix(iRow, 6))
        ZFecha = WVector1.TextMatrix(iRow, 8)
        
        WDias = 0
        WFechaDesde = ZFecha
        WFechaHasta = Fecha.Text
        
        WOrdFechaDesde = Right$(WFechaDesde, 4) + Mid$(WFechaDesde, 4, 2) + Left$(WFechaDesde, 2)
        WOrdFechaHasta = Right$(WFechaHasta, 4) + Mid$(WFechaHasta, 4, 2) + Left$(WFechaHasta, 2)
        
        If ZTipo = 2 And WOrdFechaDesde < WOrdFechaHasta Then
        
            XFec1 = ZFecha
            Call Valida_fecha1(XFec1, Auxi)
            If Auxi <> "S" Then
                ZPasa = "N"
                Exit Sub
            End If
        
            Do
                WDias = WDias + 1
                XFec1 = WFechaDesde
                SumaDia = 2
                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                WFechaDesde = XFec2
                If WFechaDesde = WFechaHasta Then
                    Exit Do
                End If
            Loop
            
            If WDias > 30 Then
                ZPasa = "N"
                Exit Sub
            End If
            
        End If
        
    Next iRow
    
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    Rem WVector1.Col = 1
    Rem WVector1.Row = 1
    
    
End Sub

Private Sub Lee_Datos()
    
    Call Limpia_Vector

    Auxi1 = Recibo.Text
    Call Ceros(Auxi1, 8)
    
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        ClaveRecibo = Recibo.Text + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
        
            Select Case Val(rstRecibosProvi!Tiporeg)
                Case 1
                    Debito = Debito + 1
                    WVector1.TextMatrix(Debito, 1) = rstRecibosProvi!Tipo1
                    WVector1.TextMatrix(Debito, 2) = rstRecibosProvi!Letra1
                    WVector1.TextMatrix(Debito, 3) = rstRecibosProvi!Punto1
                    WVector1.TextMatrix(Debito, 4) = rstRecibosProvi!Numero1
                    WVector1.TextMatrix(Debito, 5) = Str$(rstRecibosProvi!Importe1)
                    WVector1.TextMatrix(Debito, 5) = Alinea("###,###.##", WVector1.TextMatrix(Debito, 5))
                Case 2
                    Credito = Credito + 1
                    WVector1.TextMatrix(Credito, 6) = rstRecibosProvi!Tipo2
                    WVector1.TextMatrix(Credito, 7) = rstRecibosProvi!Numero2
                    WVector1.TextMatrix(Credito, 8) = rstRecibosProvi!Fecha2
                    WVector1.TextMatrix(Credito, 9) = rstRecibosProvi!Banco2
                    WVector1.TextMatrix(Credito, 10) = Str$(rstRecibosProvi!Importe2)
                    WVector1.TextMatrix(Credito, 10) = Alinea("###,###.##", WVector1.TextMatrix(Credito, 10))
                    
                    WCuenta(Credito) = rstRecibosProvi!Cuenta
                    
                    ZClaveCheque(Credito, 1) = IIf(IsNull(rstRecibosProvi!ClaveCheque), "", rstRecibosProvi!ClaveCheque)
                    ZClaveCheque(Credito, 2) = IIf(IsNull(rstRecibosProvi!BancoCheque), "", rstRecibosProvi!BancoCheque)
                    ZClaveCheque(Credito, 3) = IIf(IsNull(rstRecibosProvi!SucursalCheque), "", rstRecibosProvi!SucursalCheque)
                    ZClaveCheque(Credito, 4) = IIf(IsNull(rstRecibosProvi!ChequeCheque), "", rstRecibosProvi!ChequeCheque)
                    ZClaveCheque(Credito, 5) = IIf(IsNull(rstRecibosProvi!CuentaCheque), "", rstRecibosProvi!CuentaCheque)
                    ZClaveCheque(Credito, 6) = IIf(IsNull(rstRecibosProvi!Cuit), "", rstRecibosProvi!Cuit)
                    
                    ZReciboDefinitivo = IIf(IsNull(rstRecibosProvi!ReciboDefinitivo), "0", rstRecibosProvi!ReciboDefinitivo)
                    If Val(rstRecibosProvi!Tipo2) = "02" And rstRecibosProvi!Estado2 = "X" Then
                         cmdAdd.Enabled = False
                    End If
                    If ZReciboDefinitivo <> 0 Then
                         cmdAdd.Enabled = False
                    End If
                    
                Case Else
            End Select
            rstRecibosProvi.Close
            
                Else
            Exit Do
        End If
    Loop
End Sub

Sub Verifica_datos()
    If Val(Retganancias.Text) = 0 Then
        Retganancias.Text = "0"
    End If
    If Val(RetIva.Text) = 0 Then
        RetIva.Text = "0"
    End If
    If Val(RetOtra.Text) = 0 Then
        RetOtra.Text = "0"
    End If
    If Val(RetSuss.Text) = 0 Then
        RetSuss.Text = "0"
    End If
End Sub

Sub Format_datos()
    Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
    RetIva.Text = Alinea("###,###.##", RetIva.Text)
    RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
    RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
End Sub

Sub Imprime_Datos()
    spClientes = "ConsultaClientes " + "'" + Clientes.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        Clientes.Text = rstClientes!Cliente
        DesClientes.Caption = rstClientes!Razon
        WRazon = rstClientes!Razon
        WDireccion = rstClientes!Direccion
        WLocalidad = rstClientes!Localidad
        WPostal = rstClientes!Postal
        WProvincia = Provincia(rstClientes!Provincia)
        WProv = rstClientes!Provincia
        rstClientes.Close
        Call Format_datos
    End If
End Sub

Private Sub cmdAdd_Click()

    If Recibo.Text <> "" Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        Existe = "N"
    
        ClaveRecibo = Recibo.Text + "01"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM RecibosProvi"
        ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            Existe = "S"
            rstRecibosProvi.Close
        End If
        
        If Existe = "S" Then
        
            Existe = "N"
            
            Do
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                ClaveRecibo = Recibo.Text + Auxi1
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM RecibosProvi"
                ZSql = ZSql + " Where RecibosProvi.Clave = " + "'" + ClaveRecibo + "'"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibosProvi.RecordCount > 0 Then
                    If Val(rstRecibosProvi!Tiporeg) = 2 Then
                        If Val(rstRecibosProvi!Tipo2) = 2 Then
                            If rstRecibosProvi!Estado2 <> "P" Or rstRecibosProvi!ReciboProvisorio <> 0 Then
                                Existe = "S"
                            End If
                        End If
                    End If
                    rstRecibosProvi.Close
                        Else
                    Exit Do
                End If
            Loop
            
        End If
        
    End If
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        If ZPasa = "N" Then
            m1$ = "Error en la carga de fecha de cheques"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            Exit Sub
        End If
        
        If Val(Creditos.Caption) <> Val(TotalRecibo.Text) Then
            m1$ = "El total de los valores ingresados no coinciden con el total del recibo informado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            Exit Sub
        End If
        
        ZSql = ""
        ZSql = ZSql + "DELETE RecibosProvi"
        ZSql = ZSql + " Where Recibo = " + "'" + Recibo.Text + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        
        Renglon = 0
            
        For iRow = 1 To 99
        
            If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                        
                XRecibo = Recibo.Text
                XRenglon = Auxi1
                XClientes = Clientes.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    XTipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    XTipoRec = "2"
                End If
                If Tipo3.Value = True Then
                    XTipoRec = "3"
                End If
                XRetganancias = Str$(Val(Retganancias.Text))
                XRetIva = Str$(Val(RetIva.Text))
                XRetotra = Str$(Val(RetOtra.Text))
                XRetencion = ""
                XTiporeg = "1"
                XTipo1 = WVector1.TextMatrix(iRow, 1)
                XLetra1 = WVector1.TextMatrix(iRow, 2)
                XPunto1 = WVector1.TextMatrix(iRow, 3)
                XNumero1 = WVector1.TextMatrix(iRow, 4)
                XImporte1 = WVector1.TextMatrix(iRow, 5)
                XImporteBaja = WVector1.TextMatrix(iRow, 5)
                XTipo2 = ""
                XNumero2 = ""
                XFecha2 = ""
                XFechaOrd2 = ""
                XBanco2 = ""
                XImporte2 = ""
                XEstado2 = ""
                XObservaciones = Observaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Credito)
                XCuenta = ""
                XDestino = ""
                XImpolist = ""
                XImpo1list = ""
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""
                XReciboDefinitivo = "0"
                
                XClaveCheque = ""
                XBancoCheque = ""
                XSucursalCheque = ""
                XChequeCheque = ""
                XCuentaCheque = ""
                XCuit = ""
                
                ZSql = "INSERT INTO RecibosProvi ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Recibo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Fechaord ,"
                ZSql = ZSql + "TipoRec ,"
                ZSql = ZSql + "RetGanancias ,"
                ZSql = ZSql + "RetIva ,"
                ZSql = ZSql + "RetOtra ,"
                ZSql = ZSql + "Retencion ,"
                ZSql = ZSql + "TipoReg ,"
                ZSql = ZSql + "Tipo1 ,"
                ZSql = ZSql + "Letra1 ,"
                ZSql = ZSql + "Punto1 ,"
                ZSql = ZSql + "Numero1 ,"
                ZSql = ZSql + "Importe1 ,"
                ZSql = ZSql + "Tipo2 ,"
                ZSql = ZSql + "Numero2 ,"
                ZSql = ZSql + "Fecha2 ,"
                ZSql = ZSql + "banco2 ,"
                ZSql = ZSql + "Importe2 ,"
                ZSql = ZSql + "Estado2 ,"
                ZSql = ZSql + "Empresa ,"
                ZSql = ZSql + "FechaOrd2 ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Impolist ,"
                ZSql = ZSql + "Impo1list ,"
                ZSql = ZSql + "Destino ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Marca ,"
                ZSql = ZSql + "FechaDepo ,"
                ZSql = ZSql + "FechaDepoOrd ,"
                ZSql = ZSql + "ClaveCheque ,"
                ZSql = ZSql + "BancoCheque ,"
                ZSql = ZSql + "SucursalCheque ,"
                ZSql = ZSql + "ChequeCheque ,"
                ZSql = ZSql + "CuentaCheque ,"
                ZSql = ZSql + "ReciboDefinitivo ,"
                ZSql = ZSql + "Cuit)"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + XClave + "',"
                ZSql = ZSql + "'" + XRecibo + "',"
                ZSql = ZSql + "'" + XRenglon + "',"
                ZSql = ZSql + "'" + XClientes + "',"
                ZSql = ZSql + "'" + XFecha + "',"
                ZSql = ZSql + "'" + XFechaOrd + "',"
                ZSql = ZSql + "'" + XTipoRec + "',"
                ZSql = ZSql + "'" + XRetganancias + "',"
                ZSql = ZSql + "'" + XRetIva + "',"
                ZSql = ZSql + "'" + XRetotra + "',"
                ZSql = ZSql + "'" + XRetencion + "',"
                ZSql = ZSql + "'" + XTiporeg + "',"
                ZSql = ZSql + "'" + XTipo1 + "',"
                ZSql = ZSql + "'" + XLetra1 + "',"
                ZSql = ZSql + "'" + XPunto1 + "',"
                ZSql = ZSql + "'" + XNumero1 + "',"
                ZSql = ZSql + "'" + XImporte1 + "',"
                ZSql = ZSql + "'" + XTipo2 + "',"
                ZSql = ZSql + "'" + XNumero2 + "',"
                ZSql = ZSql + "'" + XFecha2 + "',"
                ZSql = ZSql + "'" + XBanco2 + "',"
                ZSql = ZSql + "'" + XImporte2 + "',"
                ZSql = ZSql + "'" + XEstado2 + "',"
                ZSql = ZSql + "'" + XEmpresa + "',"
                ZSql = ZSql + "'" + XFechaOrd2 + "',"
                ZSql = ZSql + "'" + XImporte + "',"
                ZSql = ZSql + "'" + XObservaciones + "',"
                ZSql = ZSql + "'" + XImpolist + "',"
                ZSql = ZSql + "'" + XImpo1list + "',"
                ZSql = ZSql + "'" + XDestino + "',"
                ZSql = ZSql + "'" + XCuenta + "',"
                ZSql = ZSql + "'" + XMarca + "',"
                ZSql = ZSql + "'" + XFechaDepo + "',"
                ZSql = ZSql + "'" + XFechaDepoOrd + "',"
                ZSql = ZSql + "'" + XClaveCheque + "',"
                ZSql = ZSql + "'" + XBancoCheque + "',"
                ZSql = ZSql + "'" + XSucursalCheque + "',"
                ZSql = ZSql + "'" + XChequeCheque + "',"
                ZSql = ZSql + "'" + XCuentaCheque + "',"
                ZSql = ZSql + "'" + XReciboDefinitivo + "',"
                ZSql = ZSql + "'" + XCuit + "')"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
                
        Next iRow
            
            
        For iRow = 1 To 99
        
            If Val(WVector1.TextMatrix(iRow, 10)) <> 0 Then
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                    
                XRecibo = Recibo.Text
                XRenglon = Auxi1
                XClientes = Clientes.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    XTipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    XTipoRec = "2"
                End If
                If Tipo3.Value = True Then
                    XTipoRec = "3"
                End If
                XRetganancias = Str$(Val(Retganancias.Text))
                XRetIva = Str$(Val(RetIva.Text))
                XRetotra = Str$(Val(RetOtra.Text))
                XRetencion = ""
                XTiporeg = "2"
                XTipo1 = ""
                XLetra1 = ""
                XPunto1 = ""
                XNumero1 = ""
                XImporte1 = ""
                XTipo2 = WVector1.TextMatrix(iRow, 6)
                XNumero2 = WVector1.TextMatrix(iRow, 7)
                XFecha2 = WVector1.TextMatrix(iRow, 8)
                XFechaOrd2 = Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                XBanco2 = WVector1.TextMatrix(iRow, 9)
                XImporte2 = WVector1.TextMatrix(iRow, 10)
                XEstado2 = "P"
                XObservaciones = Observaciones.Text
                XEmpresa = "1"
                XClave = XRecibo + XRenglon
                XImporte = Str$(Credito)
                If XTipo2 = 4 Then
                    XCuenta = WCuenta(iRow)
                        Else
                    XCuenta = ""
                End If
                XMarca = ""
                XFechaDepo = ""
                XFechaDepoOrd = ""
                If Val(XTipo2) = 1 Or Val(XTipo2) = 4 Then
                    XEstado2 = "X"
                End If
                
                XClaveCheque = ZClaveCheque(iRow, 1)
                XBancoCheque = ZClaveCheque(iRow, 2)
                XSucursalCheque = ZClaveCheque(iRow, 3)
                XChequeCheque = ZClaveCheque(iRow, 4)
                XCuentaCheque = ZClaveCheque(iRow, 5)
                XCuit = ZClaveCheque(iRow, 6)
                XReciboDefinitivo = "0"
                    
                ZSql = "INSERT INTO RecibosProvi ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Recibo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Fechaord ,"
                ZSql = ZSql + "TipoRec ,"
                ZSql = ZSql + "RetGanancias ,"
                ZSql = ZSql + "RetIva ,"
                ZSql = ZSql + "RetOtra ,"
                ZSql = ZSql + "Retencion ,"
                ZSql = ZSql + "TipoReg ,"
                ZSql = ZSql + "Tipo1 ,"
                ZSql = ZSql + "Letra1 ,"
                ZSql = ZSql + "Punto1 ,"
                ZSql = ZSql + "Numero1 ,"
                ZSql = ZSql + "Importe1 ,"
                ZSql = ZSql + "Tipo2 ,"
                ZSql = ZSql + "Numero2 ,"
                ZSql = ZSql + "Fecha2 ,"
                ZSql = ZSql + "banco2 ,"
                ZSql = ZSql + "Importe2 ,"
                ZSql = ZSql + "Estado2 ,"
                ZSql = ZSql + "Empresa ,"
                ZSql = ZSql + "FechaOrd2 ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Impolist ,"
                ZSql = ZSql + "Impo1list ,"
                ZSql = ZSql + "Destino ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Marca ,"
                ZSql = ZSql + "FechaDepo ,"
                ZSql = ZSql + "FechaDepoOrd ,"
                ZSql = ZSql + "ClaveCheque ,"
                ZSql = ZSql + "BancoCheque ,"
                ZSql = ZSql + "SucursalCheque ,"
                ZSql = ZSql + "ChequeCheque ,"
                ZSql = ZSql + "CuentaCheque ,"
                ZSql = ZSql + "ReciboDefinitivo ,"
                ZSql = ZSql + "Cuit)"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + XClave + "',"
                ZSql = ZSql + "'" + XRecibo + "',"
                ZSql = ZSql + "'" + XRenglon + "',"
                ZSql = ZSql + "'" + XClientes + "',"
                ZSql = ZSql + "'" + XFecha + "',"
                ZSql = ZSql + "'" + XFechaOrd + "',"
                ZSql = ZSql + "'" + XTipoRec + "',"
                ZSql = ZSql + "'" + XRetganancias + "',"
                ZSql = ZSql + "'" + XRetIva + "',"
                ZSql = ZSql + "'" + XRetotra + "',"
                ZSql = ZSql + "'" + XRetencion + "',"
                ZSql = ZSql + "'" + XTiporeg + "',"
                ZSql = ZSql + "'" + XTipo1 + "',"
                ZSql = ZSql + "'" + XLetra1 + "',"
                ZSql = ZSql + "'" + XPunto1 + "',"
                ZSql = ZSql + "'" + XNumero1 + "',"
                ZSql = ZSql + "'" + XImporte1 + "',"
                ZSql = ZSql + "'" + XTipo2 + "',"
                ZSql = ZSql + "'" + XNumero2 + "',"
                ZSql = ZSql + "'" + XFecha2 + "',"
                ZSql = ZSql + "'" + XBanco2 + "',"
                ZSql = ZSql + "'" + XImporte2 + "',"
                ZSql = ZSql + "'" + XEstado2 + "',"
                ZSql = ZSql + "'" + XEmpresa + "',"
                ZSql = ZSql + "'" + XFechaOrd2 + "',"
                ZSql = ZSql + "'" + XImporte + "',"
                ZSql = ZSql + "'" + XObservaciones + "',"
                ZSql = ZSql + "'" + XImpolist + "',"
                ZSql = ZSql + "'" + XImpo1list + "',"
                ZSql = ZSql + "'" + XDestino + "',"
                ZSql = ZSql + "'" + XCuenta + "',"
                ZSql = ZSql + "'" + XMarca + "',"
                ZSql = ZSql + "'" + XFechaDepo + "',"
                ZSql = ZSql + "'" + XFechaDepoOrd + "',"
                ZSql = ZSql + "'" + XClaveCheque + "',"
                ZSql = ZSql + "'" + XBancoCheque + "',"
                ZSql = ZSql + "'" + XSucursalCheque + "',"
                ZSql = ZSql + "'" + XChequeCheque + "',"
                ZSql = ZSql + "'" + XCuentaCheque + "',"
                ZSql = ZSql + "'" + XReciboDefinitivo + "',"
                ZSql = ZSql + "'" + XCuit + "')"
                spRecibosProvi = ZSql
                Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                
                If Trim(XCuit) <> "" Then
                
                    XClaveCuit = XBancoCheque + XSucursalCheque + XCuentaCheque
            
                    ZSql = "Select *"
                    ZSql = ZSql + " FROM Cuit"
                    ZSql = ZSql + " Where Cuit.Clave = " + "'" + XClaveCuit + "'"
                    spCuit = ZSql
                    Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuit.RecordCount > 0 Then
                        rstCuit.Close
                            Else
                        ZSql = "INSERT INTO Cuit ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Banco ,"
                        ZSql = ZSql + "Sucursal ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Cuit)"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + XClaveCuit + "',"
                        ZSql = ZSql + "'" + XBancoCheque + "',"
                        ZSql = ZSql + "'" + XSucursalCheque + "',"
                        ZSql = ZSql + "'" + XCuentaCheque + "',"
                        ZSql = ZSql + "'" + XCuit + "')"
                        spCuit = ZSql
                        Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                End If
                    
            End If
               
        Next iRow
            
        ZSql = ""
        ZSql = ZSql + "UPDATE RecibosProvi SET "
        ZSql = ZSql + " RetSuss = " + "'" + RetSuss.Text + "',"
        ZSql = ZSql + " ComproGanan = " + "'" + ComproGanan.Text + "',"
        ZSql = ZSql + " ComproIva = " + "'" + ComproIva.Text + "',"
        ZSql = ZSql + " ComproIb = " + "'" + ComproIB.Text + "',"
        ZSql = ZSql + " ComproSuss = " + "'" + ComproSuss.Text + "'"
        ZSql = ZSql + " Where Recibo = " + "'" + Recibo.Text + "'"
        spRecibosProvi = ZSql
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            

        Call CmdLimpiar_Click
        Recibo.SetFocus
        
    End If
End Sub



Private Sub cmdDelete_Click()
    If Recibo.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Recibo.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
    End If
    Clientes.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    
    Call Limpia_Vector
        
    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    ComproGanan.Text = ""
    ComproIva.Text = ""
    ComproIB.Text = ""
    ComproSuss.Text = ""
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Cuenta.Text = ""
    Paridad.Text = ""
    TotalRecibo.Text = ""
    
    cmdAdd.Enabled = True
    
    Erase ZClaveCheque
    
    IngreCuenta.Visible = False
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
    
    Recibo.Text = ""
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM RecibosProvi"
    Rem spRecibosProvi = ZSql
    Rem Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibosProvi.RecordCount > 0 Then
    Rem     rstRecibosProvi.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibosProvi!ReciboMayor), "0", rstRecibosProvi!ReciboMayor)
    Rem     Recibo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
    Rem     rstRecibosProvi.Close
    Rem         Else
    Rem     Recibo.Text = "1"
    Rem End If
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstImpreRec
        .Close
    End With
    Recibo.SetFocus
    PrgRecibosProviNuevo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImpreRec
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(Recibo.Text) <> 0 Then
    
            Auxi1 = Recibo.Text
            Call Ceros(Auxi1, 6)
            Recibo.Text = Auxi1
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM RecibosProvi"
            ZSql = ZSql + " Where RecibosProvi.Recibo = " + "'" + Recibo.Text + "'"
            spRecibosProvi = ZSql
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibosProvi.RecordCount > 0 Then
                Existe = "S"
                Clientes.Text = rstRecibosProvi!Cliente
                Observaciones.Text = rstRecibosProvi!Observaciones
                Fecha.Text = rstRecibosProvi!Fecha
                Retganancias.Text = rstRecibosProvi!Retganancias
                RetIva.Text = rstRecibosProvi!RetIva
                RetOtra.Text = rstRecibosProvi!RetOtra
                RetSuss.Text = IIf(IsNull(rstRecibosProvi!RetSuss), "", rstRecibosProvi!RetSuss)
                ComproGanan.Text = IIf(IsNull(rstRecibosProvi!ComproGanan), "", rstRecibosProvi!ComproGanan)
                ComproIva.Text = IIf(IsNull(rstRecibosProvi!ComproIva), "", rstRecibosProvi!ComproIva)
                ComproIB.Text = IIf(IsNull(rstRecibosProvi!ComproIB), "", rstRecibosProvi!ComproIB)
                ComproSuss.Text = IIf(IsNull(rstRecibosProvi!ComproSuss), "", rstRecibosProvi!ComproSuss)
                Tipo1.Value = True
                Tipo2.Value = False
                Select Case Val(rstRecibosProvi!TipoRec)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case Else
                End Select
                rstRecibosProvi.Close
            End If
            
            If Existe = "S" Then
                Call Imprime_Datos
                Call Lee_Datos
                Call Suma_Datos
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                    Else
                Fecha.SetFocus
            End If
        
        End If
        
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Clientes.SetFocus
                Else
            G$ = "Formato de fecha invalido"
            A% = MsgBox(G$, 0, "Emision de Recibos")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Clientes.Text <> "" Then
            With rstClientes
                spClientes = "ConsultaClientes " + "'" + Clientes.Text + "'"
                Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                If rstClientes.RecordCount > 0 Then
                    Clientes.Text = rstClientes!Cliente
                    DesClientes.Caption = rstClientes!Razon
                    WRazon = rstClientes!Razon
                    WDireccion = rstClientes!Direccion
                    WLocalidad = rstClientes!Localidad
                    WPostal = rstClientes!Postal
                    WProvincia = Provincia(rstClientes!Provincia)
                    WProv = rstClientes!Provincia
                    Rem Call Imprime_Datos
                    Observaciones.SetFocus
                    rstClientes.Close
                        Else
                    Clientes.SetFocus
                End If
            End With
        End If
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.SetFocus
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
        Call Suma_Datos
        If Val(Retganancias.Text) <> 0 Then
            EntraComproGanan.Visible = True
            ComproGanan.SetFocus
                Else
            RetIva.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproGanan_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproGanan.Visible = False
        RetIva.SetFocus
    End If
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.Text = Alinea("###,###.##", RetIva.Text)
        Call Suma_Datos
        If Val(RetIva.Text) <> 0 Then
            EntraComproIva.Visible = True
            ComproIva.SetFocus
                Else
            RetOtra.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproIva.Visible = False
        RetOtra.SetFocus
    End If
End Sub

Private Sub RetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
        Call Suma_Datos
        If Val(RetOtra.Text) <> 0 Then
            EntraComproIb.Visible = True
            ComproIB.SetFocus
                Else
            RetSuss.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ComproIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproIb.Visible = False
        RetSuss.SetFocus
    End If
End Sub

Private Sub RetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
        Call Suma_Datos
        If Val(RetSuss.Text) <> 0 Then
            EntraComproSuss.Visible = True
            ComproSuss.SetFocus
                Else
            TotalRecibo.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub totalrecibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TotalRecibo.Text = Alinea("###,###.##", TotalRecibo.Text)
        If Val(WEmpresa) = 1 Then
            WVector1.Col = 6
                Else
            WVector1.Col = 1
        End If
        WVector1.Row = 1
        Call StartEdit
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub ComproSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EntraComproSuss.Visible = False
        If Val(WEmpresa) = 1 Then
            WVector1.Col = 6
                Else
            WVector1.Col = 1
        End If
        WVector1.Row = 1
        Call StartEdit
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IngresaCuit.Visible = False
        ZClaveCheque(WVector1.Row, 6) = Cuit.Text
        If WVector1.Row < WVector1.Rows - 1 Then
            WVector1.Row = WVector1.Row + 1
        End If
        WVector1.Col = 6
        Call StartEdit
    End If
End Sub

Private Sub Consulta_Click()

    XRow = WVector1.Row
    XCol = WVector1.Col

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Cuenta Corrientes"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
    Ayuda.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spCliente = "ListaCliente"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCliente.Close
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 1
        
            XParam = "'" + Clientes.Text + "','" _
                        + Clientes.Text + "'"
            spCtaCte = "ListaCtacteDesdeHasta" + XParam
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
            
            With rstCtaCte
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstCtaCte!Saldo <> 0 Then
                            Auxi = Str$(rstCtaCte!Saldo)
                            Auxi = Mascara("###,###.##", Auxi$)
                            Auxi1 = Str$(rstCtaCte!Numero)
                            Call Ceros(Auxi1, 6)
                            IngresaItem = rstCtaCte!Impre + " " + Auxi1 + " " + rstCtaCte!Fecha + " " + Auxi
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCtaCte!Clave
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtaCte.Close
            
            End If
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        If Ayuda.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Razon"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
        End If
    
    End If

End Sub


Private Sub Pantalla_Click()
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Clientes.Text = WCliente
                DesClientes.Caption = rstCliente!Razon
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WProvincia = Provincia(rstCliente!Provincia)
                WProv = rstCliente!Provincia
                                Else
                Clientes.Text = ""
            End If
            
            Pantalla.Visible = False
            Clientes.SetFocus
            
        Case 1
        
            If Tipo1.Value = True Then
        
                Entra = "S"
                Indice = Pantalla.ListIndex
                Compara1 = WIndice.List(Indice)
            
                For iRow = 1 To 99
                    Compara2 = WVector1.TextMatrix(iRow, 1)
                    Compara2 = Compara2 + WVector1.TextMatrix(iRow, 4) + "01"
                    If Compara1 = Compara2 Then
                        Entra = "N"
                        Exit For
                    End If
                Next iRow
                
                If Entra = "S" Then
                
                    For iRow = 1 To 99
                        If WVector1.TextMatrix(iRow, 1) = "" Then
                            XRow = iRow
                            Exit For
                        End If
                    Next iRow
                    
                    Indice = Pantalla.ListIndex
                    ClaveCtacte = WIndice.List(Indice)
                    spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCtaCte.RecordCount > 0 Then
                        
                        Auxi = rstCtaCte!Tipo
                        Call Ceros(Auxi, 2)
                        WVector1.TextMatrix(XRow, 1) = Auxi
                        
                        WVector1.TextMatrix(XRow, 2) = ""
                        
                        WVector1.TextMatrix(XRow, 3) = ""
                        
                        Auxi = rstCtaCte!Numero
                        Call Ceros(Auxi, 8)
                        WVector1.TextMatrix(XRow, 4) = Auxi
                        
                        WVector1.TextMatrix(XRow, 5) = Str$(rstCtaCte!Saldo)
                        WVector1.TextMatrix(XRow, 5) = Alinea("###,###.##", WVector1.TextMatrix(XRow, 5))
                        
                        Call Suma_Datos
                        
                        rstCtaCte.Close
                        
                    End If
                
                End If
                    
                WVector1.Row = XRow
                WVector1.Col = 1
                Call StartEdit
            
            End If
                
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()
    
    Call Limpia_Vector
 
 
    Provincia$(0) = "Capital Federal"
    Provincia$(1) = "Buenos Aires"
    Provincia$(2) = "Catamarca"
    Provincia$(3) = "Cordoba"
    Provincia$(4) = "Corrientes"
    Provincia$(5) = "Chaco"
    Provincia$(6) = "Chubut"
    Provincia$(7) = "Entre Rios"
    Provincia$(8) = "Formosa"
    Provincia$(9) = "Jujuy"
    Provincia$(10) = "La Pampa"
    Provincia$(11) = "La Rioja"
    Provincia$(12) = "Mendoza"
    Provincia$(13) = "Misiones"
    Provincia$(14) = "Neuquen"
    Provincia$(15) = "Rio Negro"
    Provincia$(16) = "Salta"
    Provincia$(17) = "San Juan"
    Provincia$(18) = "San Luis"
    Provincia$(19) = "Santa Cruz"
    Provincia$(20) = "Santa Fe"
    Provincia$(21) = "Santiago del Estero"
    Provincia$(22) = "Tucuman"
    Provincia$(23) = "Tierra del Fuego"
    Provincia$(24) = "Exterior"
    Provincia$(25) = ""
    
    ZBancos(3) = "BEAL"
    ZBancos(5) = "AMRO BANK"
    ZBancos(7) = "GALICIA"
    ZBancos(10) = "LLOYDS BANK"
    ZBancos(11) = "NACION"
    ZBancos(14) = "PROVINCIA"
    ZBancos(15) = "BANKBOSTON"
    ZBancos(16) = "CITIBANK"
    ZBancos(17) = "FRANCES"
    ZBancos(18) = "TOKYO"
    ZBancos(20) = "CORDOBA"
    ZBancos(27) = "SUPERVIELLE"
    ZBancos(29) = "CIUDAD"
    ZBancos(30) = "CENTRAL"
    ZBancos(34) = "PATAGONIA"
    ZBancos(44) = "HIPOTECARIO"
    ZBancos(45) = "SAN JUAN"
    ZBancos(46) = "BRASIL"
    ZBancos(60) = "TUCUMAN"
    ZBancos(65) = "ROSARIO"
    ZBancos(72) = "RIO"
    ZBancos(79) = "CUYO"
    ZBancos(83) = "CHUBUT"
    ZBancos(86) = "SANTA CRUZ"
    ZBancos(93) = "LA PAMPA"
    ZBancos(94) = "CORRIENTES "
    ZBancos(97) = "NEUQUEN"
    ZBancos(137) = "EMP.TUCUMAN"
    ZBancos(147) = "B.I.CRED."
    ZBancos(148) = "LA PLATA"
    ZBancos(150) = "HSBC"
    ZBancos(165) = "JPMORGAN"
    ZBancos(191) = "CREDICOOP"
    ZBancos(198) = "VALORES"
    ZBancos(247) = "ROELA"
    ZBancos(254) = "MARIVA"
    ZBancos(259) = "ITAU"
    ZBancos(265) = "HSBC"
    ZBancos(262) = "OF AMERICA"
    ZBancos(266) = "BNP PARIBAS"
    ZBancos(268) = "T.FUEGO"
    ZBancos(269) = "URUGUAY"
    ZBancos(277) = "SAENZ"
    ZBancos(281) = "MERIDIAN"
    ZBancos(285) = "MACRO"
    ZBancos(293) = "MERCURIO"
    ZBancos(294) = "ING.BANK"
    ZBancos(295) = "AMERICAN"
    ZBancos(297) = "BANEX"
    ZBancos(299) = "COMAFI"
    ZBancos(300) = "INVERSION"
    ZBancos(301) = "PIANO"
    ZBancos(303) = "FINANSUR"
    ZBancos(305) = "JULIO"
    ZBancos(306) = "P.INVERSIONES"
    ZBancos(309) = "LA RIOJA"
    ZBancos(310) = "DEL SOL"
    ZBancos(311) = "CHACO"
    ZBancos(312) = "DE INVERSIONES"
    ZBancos(315) = "FORMOSA"
    ZBancos(319) = "CMF"
    ZBancos(320) = "BANEX"
    ZBancos(321) = "S.ESTERO"
    ZBancos(322) = "IND.AZUL"
    ZBancos(325) = "DEUTSCHE BANK"
    ZBancos(330) = "SANTA FE"
    ZBancos(331) = "CETELEM"
    ZBancos(332) = "SERV.FINAN."
    ZBancos(335) = "COFIDIS"
    ZBancos(336) = "BRADESCO"
    ZBancos(338) = "SERV.Y TRANS."
    ZBancos(339) = "RCI BANQUE"
    ZBancos(340) = "DE CREDITO"
    ZBancos(386) = "ENTRE RIOS"
    ZBancos(387) = "SUQUIA"
    ZBancos(388) = "BISEL"
    ZBancos(389) = "COLUMBIA"
     
    ImpreTipo$(1) = "FC"
     
    Tipo1.Value = True
    Tipo2.Value = False
    
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    
    ComproGanan.Text = ""
    ComproIva.Text = ""
    ComproIB.Text = ""
    ComproSuss.Text = ""

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    Debitos.Caption = ""
    Creditos.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    Paridad.Text = ""
    TotalRecibo.Text = ""
    
    cmdAdd.Enabled = True
    
    Erase ZClaveCheque
    
    Recibo.Text = ""
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM RecibosProvi"
    Rem spRecibosProvi = ZSql
    Rem Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibosProvi.RecordCount > 0 Then
    Rem     rstRecibosProvi.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibosProvi!ReciboMayor), "0", rstRecibosProvi!ReciboMayor)
    Rem     Recibo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
    Rem     rstRecibosProvi.Close
    Rem         Else
    Rem     Recibo.Text = "1"
    Rem End If
    
End Sub

Private Sub Cuenta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta1.Text <> "" Then
            spCuenta = "ConsultaCuentas " + "'" + Cuenta1.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                rstCuenta.Close
                WCuenta(WVector1.Row) = Cuenta1.Text
                IngreCuenta.Visible = False
                WVector1.Col = 6
                Call StartEdit
                    Else
                Cuenta.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Lectora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Lectora.Text) = 31 Then
        
            ZEntra = "S"
        
            Sql1 = "Select *"
            Sql2 = " FROM Recibos"
            Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
            spRecibos = Sql1 + Sql2 + Sql3
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                m1$ = "Cheque ya cargado"
                A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                ZEntra = "N"
                rstRecibos.Close
            End If
        
            Sql1 = "Select *"
            Sql2 = " FROM RecibosProvi"
            Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
            spRecibosProvi = Sql1 + Sql2 + Sql3
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibosProvi.RecordCount > 0 Then
                m1$ = "Cheque ya cargado"
                A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                ZEntra = "N"
                rstRecibosProvi.Close
            End If
            
            If ZEntra = "S" Then
                For ZZCiclo = 1 To 100
                    If ZClaveCheque(ZZCiclo, 1) = Lectora.Text Then
                        m1$ = "Cheque ya cargado"
                        A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                        ZEntra = "N"
                    End If
                Next ZZCiclo
            End If
            
            If ZEntra = "S" Then
        
                ZNombreBanco = ZBancos(Val(Mid$(Lectora, 2, 3)))
                ZNroCuenta = Mid$(Lectora, 12, 8)
            
                ZZBanco = Mid$(Lectora, 2, 3)
                ZZSucursal = Mid$(Lectora, 5, 3)
                ZZNroCheque = Mid$(Lectora, 12, 8)
                ZZNroCuenta = Mid$(Lectora, 20, 11)

                ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
                
                WVector1.TextMatrix(WVector1.Row, 6) = "02"
                WVector1.TextMatrix(WVector1.Row, 9) = ZNombreBanco + "/" + Left$(Clientes.Text, 1) + ZSuma
                WVector1.TextMatrix(WVector1.Row, 7) = ZNroCuenta
                WVector1.TextMatrix(WVector1.Row, 8).Text = ""
            
                ZClaveCheque(WVector1.Row, 1) = Lectora.Text
                ZClaveCheque(WVector1.Row, 2) = ZZBanco
                ZClaveCheque(WVector1.Row, 3) = ZZSucursal
                ZClaveCheque(WVector1.Row, 4) = ZZNroCheque
                ZClaveCheque(WVector1.Row, 5) = ZZNroCuenta
                ZClaveCheque(WVector1.Row, 6) = ""
            
                ZZClave = ZZBanco + ZZSucursal + ZZNroCuenta
                ZZCuit = ""
            
                ZSql = "Select *"
                ZSql = ZSql + " FROM Cuit"
                ZSql = ZSql + " Where Cuit.Clave = " + "'" + ZZClave + "'"
                spCuit = ZSql
                Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuit.RecordCount > 0 Then
                    ZZCuit = Trim(rstCuit!Cuit)
                    rstCuit.Close
                End If
                
                ZClaveCheque(WVector1.Row, 6) = ZZCuit
                
                Lectora.Visible = False
                
                WVector1.Col = 6
                Call StartEdit
                
                    Else
                    
                WVector1.Col = 6
                WVector1.Text = ""
                WVector1.Col = 5
                Lectora.Visible = False
                Call StartEdit
                
            End If
                    
            
                Else
                
            WVector1.Col = 6
            WVector1.Text = ""
            Call StartEdit
            
            Lectora.Visible = False
            
        End If
    End If
End Sub

Private Sub Toto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 69 Or KeyAscii = 101 Then
        WVector1.SetFocus
    End If
End Sub






Rem
Rem Controles de la grilla
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
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
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
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
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
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
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

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
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

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

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

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case 10
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 6
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
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 2)
                WVector1.Text = Auxi$
                WVector1.Col = 4
            End If
        
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
            With rstCtaCte
            
                WVector1.Col = 1
                XTipo = WVector1.Text
                
                WVector1.Col = 4
                XNumero = WVector1.Text
                
                ClaveCtacte = XTipo + XNumero + "01"
                
                spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                
                    WVector1.Col = 5
                    XRow = WVector1.Row
                    If Val(WVector1.Text) = 0 Then
                        WVector1.Text = Str$(!Saldo)
                        Call Suma_Datos
                    End If
                    rstCtaCte.Close
                    
                        Else
                        
                    WControl = "N"
                    
                End If
            End With
                
        Case 5
            With rstCtaCte
                WVector1.Col = 1
                XTipo = WVector1.Text
                WVector1.Col = 4
                XNumero = WVector1.Text
                
                ClaveCtacte = XTipo + XNumero + "01"
                
                spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    Saldo = Alinea("###,###.##", Str$(rstCtaCte!Saldo))
                    rstCtaCte.Close
                        Else
                    Saldo = 0
                End If
            
            End With
        
            WVector1.Col = 5
            If Abs(Val(WVector1.Text)) > Abs(Val(Saldo)) Then
                WVector1.Text = ""
                WControl = "N"
                    Else
                WVector1.Text = Alinea("###,###.##", WVector1.Text)
                Call Suma_Datos
            End If
            
        Case 6
            If Len(WVector1.Text) = 31 Then
                Lectora.Text = WVector1.Text
                Call Lectora_Keypress(13)
                    Else
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Or Val(WVector1.Text) = 99 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                    Select Case Val(WVector1.Text)
                        Case 1, 4
                            WVector1.Col = 7
                            WVector1.Text = ""
                            WVector1.Col = 8
                            WVector1.Text = ""
                            WVector1.Col = 9
                            WVector1.Text = ""
                    End Select
                        Else
                    WControl = "N"
                End If
            End If

        Case 7
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
        
        Case 8
            If Len(WVector1.Text) = 5 Then
                If Val(Right$(WVector1.Text, 2)) > 6 Then
                    WVector1.Text = WVector1.Text + "/2013"
                        Else
                    WVector1.Text = WVector1.Text + "/2014"
                End If
            End If
            WVector1.Col = 8
            Call Valida_fecha1(WVector1.Text, Auxi)
            
            If Auxi <> "S" Then
            
                WControl = "N"
                
                        Else
                        
                ZPasa = ""
                ZFecha = WVector1.Text
                WVector1.Col = 6
                ZTipo = Val(WVector1.Text)

                WDias = 0
                WFechaDesde = ZFecha
                WFechaHasta = Fecha.Text

                WOrdFechaDesde = Right$(WFechaDesde, 4) + Mid$(WFechaDesde, 4, 2) + Left$(WFechaDesde, 2)
                WOrdFechaHasta = Right$(WFechaHasta, 4) + Mid$(WFechaHasta, 4, 2) + Left$(WFechaHasta, 2)

                If ZTipo = 2 And WOrdFechaDesde < WOrdFechaHasta Then

                    Do
                        WDias = WDias + 1
                        XFec1 = WFechaDesde
                        SumaDia = 2
                        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                        WFechaDesde = XFec2
                        If WFechaDesde = WFechaHasta Then
                            Exit Do
                        End If
                    Loop
    
                    If WDias > 30 Then
                        ZPasa = "N"
                    End If
    
                End If
                
                If ZPasa = "N" Then
                    m1$ = "Error en la carga de fecha de cheque"
                    A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                    WControl = "N"
                        Else
                    WVector1.Col = 9
                    If Trim(WVector1.Text) = "" Then
                        WVector1.Col = 8
                    End If
                End If
            
            End If
    
    
        
        Case 9
            ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
            ZAgrega = Left$(Clientes.Text, 1) + ZSuma
            ZLong = Len(ZAgrega)
            If Right$(WVector1.Text, ZLong) <> ZAgrega Then
                WVector1.Text = WVector1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
            End If
            
        Case 10
            iRow = WVector1.Row
            WVector1.Col = 6
            XTipo = WVector1.Text
            WVector1.Col = 10
            WVector1.Text = Alinea("###,###.##", WVector1.Text)
            Call Suma_Datos
            WVector1.Row = iRow
            
            If Val(XTipo) = 4 Then
                Cuenta1.Text = WCuenta(WVector1.Row)
                IngreCuenta.Visible = True
                Cuenta1.SetFocus
            End If
            
            ZZCuit = ZClaveCheque(WVector1.Row, 6)
            If Val(XTipo) = 2 And ZZCuit = "" Then
                WControlII = "N"
                WControl = "N"
                Cuit.Text = ""
                IngresaCuit.Visible = True
                Cuit.SetFocus
            End If
            
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
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

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 11
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
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
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 6
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Numero/Cta"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 40
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
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

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


