VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgInforme 
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
      Left            =   120
      TabIndex        =   79
      Top             =   960
      Visible         =   0   'False
      Width           =   11415
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
         TabIndex        =   105
         Top             =   1800
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
         TabIndex        =   104
         Top             =   1440
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
         TabIndex        =   103
         Top             =   1080
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
         TabIndex        =   102
         Top             =   720
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
         TabIndex        =   96
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   97
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
         TabIndex        =   95
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
         TabIndex        =   92
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
         TabIndex        =   89
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   1320
         Width           =   3135
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
      Index           =   9
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
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
      Height          =   2535
      Left            =   1680
      TabIndex        =   66
      Top             =   1920
      Visible         =   0   'False
      Width           =   6015
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
         TabIndex        =   75
         Top             =   2040
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox Vencimiento 
         Height          =   285
         Left            =   1920
         TabIndex        =   98
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
         TabIndex        =   99
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
         TabIndex        =   74
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
         TabIndex        =   70
         Top             =   480
         Width           =   1575
      End
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
      Picture         =   "informelabo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   65
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
      TabIndex        =   64
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
      TabIndex        =   63
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
      TabIndex        =   61
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
      TabIndex        =   60
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
      TabIndex        =   59
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   56
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3600
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox WLote1 
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
         TabIndex        =   53
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox WLote2 
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
         TabIndex        =   52
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Wlote3 
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
         TabIndex        =   51
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox WCanti1 
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
         TabIndex        =   50
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WCanti2 
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
         TabIndex        =   49
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCanti3 
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
         TabIndex        =   48
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WLote4 
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
         TabIndex        =   47
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox WLote5 
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
         TabIndex        =   46
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox WCanti4 
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
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WCanti5 
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
         TabIndex        =   44
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
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
      Left            =   3360
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
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
         TabIndex        =   76
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
         TabIndex        =   77
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
      ItemData        =   "informelabo.frx":0742
      Left            =   3360
      List            =   "informelabo.frx":0749
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   7215
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
      TabIndex        =   62
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
Attribute VB_Name = "PrgInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private Verifica(100, 2) As String
Private Entra As String
Private Auxiliar(100, 10) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim XParam As String
Private XLote(100, 15) As String
Dim XLote1 As String
Dim XCantiLote1 As String
Dim XLote2 As String
Dim XCantiLote2 As String
Dim XLote3 As String
Dim XCantiLote3 As String
Dim XLote4 As String
Dim XCantiLote4 As String
Dim XLote5 As String
Dim XCantiLote5 As String
Dim XCertificado As Integer
Dim WCertificado As String
Dim XEstado As Integer
Dim WEstado As String
Dim CargaEmpresa(12, 2) As String
Dim DatosCertificado(100, 30) As String
Dim ZLugar As Integer
Dim WCantiEti As Integer
Dim ZZCantidad As Double
Dim WProducto As String
Dim WKilosEnvase As Integer
Dim ZVencimiento As String
Dim ZCodigo As String
Dim ZRenglon As String
Dim XObservaEnvases As String
Dim EmailAddress As String
Dim WEmail(100) As String
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Sub Verifica_datos()
    If Val(Remito.Text) = 0 Then
        Remito.Text = "0"
    End If
End Sub

Private Sub AvisoError_Click()
    AvisoError.Visible = False
End Sub

Private Sub Borra_Click()

    WGrilla.Col = 1
    WGrilla.Text = ""
    
    WGrilla.Col = 2
    WGrilla.Text = ""

    WGrilla.Col = 3
    WGrilla.Text = ""
    
    WGrilla.Col = 4
    WGrilla.Text = ""
    
    WGrilla.Col = 5
    WGrilla.Text = ""
    
    WGrilla.Col = 6
    WGrilla.Text = ""
    
    WGrilla.Col = 7
    WGrilla.Text = ""
    
    WGrilla.Col = 8
    WGrilla.Text = ""
    
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WSaldo.Caption = ""
    WResta.Text = ""
    WLinea.Text = ""
    WEnvase.Text = ""
    WEtiqueta.Text = ""
    
    WLugar = WGrilla.Row
    
    Verifica(WLugar, 1) = WArticulo.Text
    Verifica(WLugar, 2) = WOrden.Text
    
    XLote(WLugar, 1) = ""
    XLote(WLugar, 2) = ""
    XLote(WLugar, 3) = ""
    XLote(WLugar, 4) = ""
    XLote(WLugar, 5) = ""
    XLote(WLugar, 6) = ""
    XLote(WLugar, 7) = ""
    XLote(WLugar, 8) = ""
    XLote(WLugar, 9) = ""
    XLote(WLugar, 10) = ""
    
    WOrden.SetFocus
    
End Sub


Private Sub CierreAviso_Click()
    Aviso.Visible = False
    WOrden.SetFocus
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    Rem  With rstProveedor
    Rem     .Close
    Rem End With
    Rem With rstArticulo
    Rem     .Close
    Rem End With
    Rem With rstOrden
    Rem     .Close
    Rem End With
    Rem With rstInforme
    Rem     .Close
    Rem End With
    
    Rem DbsVentas.Close
    Rem DbsAdminis.Close
    Rem DbsCotiza.Close
    
    PrgInforme.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub ConfirmaCertificado_Click()
    
    DatosCertificado(ZLugar, 1) = Str$(CertificadoSi.Value)
    DatosCertificado(ZLugar, 2) = Str$(CertificadoNo.Value)
    DatosCertificado(ZLugar, 3) = Certificado2.Text
            
    DatosCertificado(ZLugar, 4) = Str$(EstadoSi.Value)
    DatosCertificado(ZLugar, 5) = Str$(EstadoNo.Value)
    DatosCertificado(ZLugar, 6) = Estado2.Text
    
    DatosCertificado(ZLugar, 7) = Vencimiento.Text
    
    PantaCertificado.Visible = False
    
    If Left$(WArticulo.Text, 2) = "DY" Or Left$(WArticulo.Text, 2) = "DW" Or Left$(WArticulo.Text, 2) = "DS" Or Left$(WArticulo.Text, 2) = "DQ" Then
        CargaLote.Visible = True
        WLote1.SetFocus
            Else
        Call Alta_Vector
        Call Ingresa_Click
        WOrden.SetFocus
    End If
    
End Sub

Private Sub ConfirmaPantaEnvase_Click()
    
    DatosCertificado(ZLugar, 11) = Str$(EstadoEnvI.Value)
    DatosCertificado(ZLugar, 12) = Str$(EstadoEnvII.Value)
    DatosCertificado(ZLugar, 13) = Str$(EstadoEnvIII.Value)
    DatosCertificado(ZLugar, 14) = Str$(EstadoEnvIV.Value)
    DatosCertificado(ZLugar, 15) = Str$(EstadoEnvV.Value)
    DatosCertificado(ZLugar, 16) = Str$(EstadoEnvVI.Value)
    DatosCertificado(ZLugar, 17) = Str$(EstadoEnvVII.Value)
    DatosCertificado(ZLugar, 18) = Str$(EstadoEnvVIII.Value)
    DatosCertificado(ZLugar, 19) = Str$(EstadoEnvIX.Value)
    DatosCertificado(ZLugar, 20) = Str$(EstadoEnvX.Value)
    DatosCertificado(ZLugar, 21) = CantidadEnv.Text
    DatosCertificado(ZLugar, 22) = ObservaI.Text
    DatosCertificado(ZLugar, 23) = ObservaII.Text
    DatosCertificado(ZLugar, 24) = ObservaIII.Text
    DatosCertificado(ZLugar, 25) = ObservaIV.Text
    
    PantaEnvase.Visible = False
    
    Call Alta_Vector
    Call Ingresa_Click
    WOrden.SetFocus
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Orden de Compra"
     Opcion.AddItem "Envases"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    Rem OPEN_FILE_Informe
    Rem OPEN_FILE_Orden
    Rem OPEN_FILE_Proveedor
    Rem OPEN_FILE_Articulo
End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            spProveedor = "ListaProveedoresOrd"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = RstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = RstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
            
        Case 1
            XParam = "'" + Proveedor.Text + "'"
            spOrden = "ListaOrdenProveedor " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
            
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        Saldo = rstOrden!Cantidad - rstOrden!Recibida
                        If Saldo > 0 Then
                            IngresaItem = Str$(rstOrden!Orden) + " " + rstOrden!Articulo
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstOrden!Clave
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
            
            End If
            
        Case 2
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub WGrilla_GotFocus()

    ZLugar = WGrilla.Row

    WGrilla.Col = 1
    WOrden.Text = WGrilla.Text
    
    WGrilla.Col = 2
    If Len(WGrilla.Text) = 10 Then
        WLinea.Text = WGrilla.Row
        WArticulo.Text = WGrilla.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    WGrilla.Col = 3
    WDescripcion.Caption = WGrilla.Text

    WGrilla.Col = 4
    If Val(WGrilla.Text) <> 0 Then
        WCantidad.Text = WGrilla.Text
            Else
        WCantidad.Text = ""
    End If

    WGrilla.Col = 5
    WSaldo.Caption = WGrilla.Text
    
    WGrilla.Col = 6
    If Val(WGrilla.Text) <> 0 Then
        WResta.Text = WGrilla.Text
            Else
        WResta.Text = ""
    End If
    
    WGrilla.Col = 7
    WEnvase.Text = WGrilla.Text
    
    WGrilla.Col = 9
    WEtiqueta.Text = WGrilla.Text
    
    Entra = "N"
    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        Entra = "S"
        rstInforme.Close
    End If
    
    If Entra = "N" Then
        WOrden.SetFocus
            Else
        WEnvase.SetFocus
    End If

End Sub

Private Sub Graba_Click()

    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha del informe de recepcion es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
    
    If Trim(Remito.Text) = "" Then
        m$ = "Es obligatorio informar el numero de remito"
        A% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
        Exit Sub
    End If
    
    If Val(XOrden.Text) < 800000 Then
    
        For Ciclo = 1 To 100
        
            If Val(WGrilla.TextMatrix(Ciclo, 1)) <> 0 Then
            
                ZCertificadoSi = Val(DatosCertificado(Ciclo, 1))
                ZCertificadoNo = Val(DatosCertificado(Ciclo, 2))
                ZCertificado2 = DatosCertificado(Ciclo, 3)
            
                ZEstadoSi = Val(DatosCertificado(Ciclo, 4))
                ZEstadoNo = Val(DatosCertificado(Ciclo, 5))
                ZEstado2 = DatosCertificado(Ciclo, 6)
                
                ZVencimiento = DatosCertificado(Ciclo, 7)
            
                If ZCertificadoSi = 1 And ZCertificadoNo = 1 Then
                    m$ = "Datos del Certificado de Analisis incorrectos"
                    d% = MsgBox(m$, 0, "Informe de Recepcion")
                    Exit Sub
                End If
    
                If ZCertificadoSi = 0 And ZCertificadoNo = 0 Then
                    m$ = "Datos del Certificado de Analisis incorrectos"
                    d% = MsgBox(m$, 0, "Informe de Recepcion")
                    Exit Sub
                End If
    
                If ZEstadoSi = 1 And ZEstadoNo = 1 Then
                    m$ = "Datos del Estado de Envases incorrectos"
                    d% = MsgBox(m$, 0, "Informe de Recepcion")
                    Exit Sub
                End If
    
                If ZEstadoSi = 0 And ZEstadoNo = 0 Then
                    m$ = "Datos del Estado de Envases incorrectos"
                    d% = MsgBox(m$, 0, "Informe de Recepcion")
                    Exit Sub
                End If
            
            End If
    
        Next Ciclo
    
    End If
    
    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        rstInforme.Close
        Exit Sub
    End If
    
    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
        
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
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    If WSalidaError = "N" Then Exit Sub
    
    On Error GoTo WError
    
    If Val(XOrden.Text) < 800000 Then
    
        Renglon = 0
        Erase Auxiliar
    
        For A = 1 To 100
        
            WRow = A
            WGrilla.Row = WRow
                    
            WGrilla.Col = 1
            Orden = WGrilla.Text
                   
            WGrilla.Col = 2
            Articulo = UCase(WGrilla.Text)
                    
            WGrilla.Col = 4
            Cantidad = Val(WGrilla.Text)
        
            WGrilla.Col = 7
            Envase = Val(WGrilla.Text)
                                
            If Articulo <> "" Or Cantidad <> 0 Then
                If Envase = 0 Then
                    m$ = "No se ha informado el tipo de envase recibido"
                    d% = MsgBox(m$, 0, "Informe de Recepcion")
                    WGrilla.Col = 1
                    WGrilla.Row = 1
                    Exit Sub
                End If
            End If
                                        
        Next A
        
    End If
    
    Call Verifica_datos

    Renglon = 0
    Erase Auxiliar

    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)

    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Auxiliar(Renglon, 1) = rstInforme!Orden
                    Auxiliar(Renglon, 2) = rstInforme!Resta
                    Auxiliar(Renglon, 3) = rstInforme!Articulo
                    Auxiliar(Renglon, 4) = rstInforme!Cantidad
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    For Da = 1 To Renglon

        Orden = Auxiliar(Da, 1)
        Resta = Val(Auxiliar(Da, 2))
        Articulo = Auxiliar(Da, 3)
        Cantidad = Val(Auxiliar(Da, 4))
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCodigo = Articulo
            WPedido = rstArticulo!Pedido + Resta
            WLaboratorio = rstArticulo!Laboratorio - Cantidad
            WDate = Date$
            WEnvase = ""
            WProveedor = ""
            WEnvase = rstArticulo!Proveedor
            WProveedor = rstArticulo!Proveedor
                
            XParam = "'" + WCodigo + "','" _
                    + WPedido + "','" _
                    + WLaboratorio + "','" _
                    + WDate + "','" _
                    + WEnvase + "','" _
                    + WProveedor + "'"
                                           
            spArticulo = "ModificaArticuloInforme " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
                
        spOrden = "ListaOrden " + "'" + Orden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    If Articulo = rstOrden!Articulo Then
                        WClave = rstOrden!Clave
                        WRecibida = rstOrden!Recibida - Resta
                        WDate = Date$
                
                        XParam = "'" + WClave + "','" _
                                + WRecibida + "','" _
                                + WDate + "'"
                                           
                        spOrden = "ModificaOrdenInforme " + XParam
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
        
    Next Da
                
    spInforme = "BorrarInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenDynaset, dbSQLPassThrough)

    Renglon = 0
    Erase Auxiliar
    
    XObservaEnvases = ""
    
    For A = 1 To 100
        
        WRow = A
        WGrilla.Row = WRow
                    
        WGrilla.Col = 1
        Orden = WGrilla.Text
                   
        WGrilla.Col = 2
        Articulo = UCase(WGrilla.Text)
        
        WGrilla.Col = 3
        ZDesArticulo = UCase(WGrilla.Text)
                    
        WGrilla.Col = 4
        Cantidad = Val(WGrilla.Text)
                    
        WGrilla.Col = 5
        Saldo = Val(WGrilla.Text)
                    
        WGrilla.Col = 6
        Resta = Val(WGrilla.Text)
            
        WGrilla.Col = 7
        Envase = Val(WGrilla.Text)
        
        WGrilla.Col = 9
        WCantiEti = Val(WGrilla.Text)
            
        WLugar = WGrilla.Row
                    
        XLote1 = XLote(WLugar, 1)
        XLote2 = XLote(WLugar, 3)
        XLote3 = XLote(WLugar, 5)
        XLote4 = XLote(WLugar, 7)
        XLote5 = XLote(WLugar, 9)
        XCantiLote1 = XLote(WLugar, 2)
        XCantiLote2 = XLote(WLugar, 4)
        XCantiLote3 = XLote(WLugar, 6)
        XCantiLote4 = XLote(WLugar, 8)
        XCantiLote5 = XLote(WLugar, 10)
        
        ZCertificadoSi = Val(DatosCertificado(WLugar, 1))
        ZCertificadoNo = Val(DatosCertificado(WLugar, 2))
        ZCertificado2 = DatosCertificado(WLugar, 3)
            
        ZEstadoSi = Val(DatosCertificado(WLugar, 4))
        ZEstadoNo = Val(DatosCertificado(WLugar, 5))
        ZEstado2 = DatosCertificado(WLugar, 6)
        
        ZVencimiento = DatosCertificado(WLugar, 7)
        ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
        
        If ZCertificadoNo = 1 Then
            WCertificado1 = "0"
        End If
    
        If ZCertificadoSi = 1 Then
            WCertificado1 = "1"
        End If
    
        If ZEstadoNo = 1 Then
            WEstado1 = "0"
        End If
    
        If ZEstadoSi = 1 Then
            WEstado1 = "1"
        End If
        
        ZEstadoEnv1 = DatosCertificado(WLugar, 11)
        ZEstadoEnv2 = DatosCertificado(WLugar, 12)
        ZEstadoEnv3 = DatosCertificado(WLugar, 13)
        ZEstadoEnv4 = DatosCertificado(WLugar, 14)
        ZEstadoEnv5 = DatosCertificado(WLugar, 15)
        ZEstadoEnv6 = DatosCertificado(WLugar, 16)
        ZEstadoEnv7 = DatosCertificado(WLugar, 17)
        ZEstadoEnv8 = DatosCertificado(WLugar, 18)
        ZEstadoEnv9 = DatosCertificado(WLugar, 19)
        ZEstadoEnv10 = DatosCertificado(WLugar, 20)
        ZCantidadEnv = DatosCertificado(WLugar, 21)
        ZObservaI = DatosCertificado(WLugar, 22)
        ZObservaII = DatosCertificado(WLugar, 23)
        ZObservaIII = DatosCertificado(WLugar, 24)
        ZObservaIV = DatosCertificado(WLugar, 25)
        
        If Val(ZEstadoEnv2) = 1 Or Val(ZEstadoEnv4) = 1 Or Val(ZEstadoEnv6) = 1 Or Val(ZEstadoEnv8) = 1 Or Val(ZEstadoEnv10) = 1 Then
            XObservaEnvases = XObservaEnvases + "Envase : " + Articulo + "  " + Trim(ZDesArticulo) + "   "
        End If
        
        WCertificado2 = ZCertificado2
        WEstado2 = ZEstado2
                   
        If Articulo <> "" Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Informe.Text)
            Call Ceros(Auxi1, 6)
                
            WClave = Auxi1 + Auxi
            WInforme = Informe.Text
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WProveedor = Proveedor.Text
            WRemito = Remito.Text
            WOrden = Orden
            WArticulo = Articulo
            WCantidad = Cantidad
            WResta = Resta
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WDate = Date$
            WEnvase = Envase
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Informe ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Informe ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Resta ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "Envase ,"
            ZSql = ZSql + "Lote1 ,"
            ZSql = ZSql + "Canti1 ,"
            ZSql = ZSql + "Lote2 ,"
            ZSql = ZSql + "Canti2 ,"
            ZSql = ZSql + "Lote3 ,"
            ZSql = ZSql + "Canti3 ,"
            ZSql = ZSql + "Lote4 ,"
            ZSql = ZSql + "Canti4 ,"
            ZSql = ZSql + "Lote5 ,"
            ZSql = ZSql + "Canti5 ,"
            ZSql = ZSql + "Certificado1 ,"
            ZSql = ZSql + "Certificado2 ,"
            ZSql = ZSql + "Estado1 ,"
            ZSql = ZSql + "Estado2 ,"
            ZSql = ZSql + "EstadoEnvI ,"
            ZSql = ZSql + "EstadoEnvII ,"
            ZSql = ZSql + "EstadoEnvIII ,"
            ZSql = ZSql + "EstadoEnvIV ,"
            ZSql = ZSql + "EstadoEnvV ,"
            ZSql = ZSql + "EstadoEnvVI ,"
            ZSql = ZSql + "EstadoEnvVII ,"
            ZSql = ZSql + "EstadoEnvVIII ,"
            ZSql = ZSql + "EstadoEnvIX ,"
            ZSql = ZSql + "EstadoEnvX ,"
            ZSql = ZSql + "CantidadEnv ,"
            ZSql = ZSql + "ObservaI ,"
            ZSql = ZSql + "ObservaII ,"
            ZSql = ZSql + "ObservaIII ,"
            ZSql = ZSql + "ObservaIV ,"
            ZSql = ZSql + "FechaVencimiento ,"
            ZSql = ZSql + "OrdFechaVencimiento )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WInforme + "',"
            ZSql = ZSql + "'" + WRenglon + "',"
            ZSql = ZSql + "'" + WFecha + "',"
            ZSql = ZSql + "'" + WRemito + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WOrden + "',"
            ZSql = ZSql + "'" + WArticulo + "',"
            ZSql = ZSql + "'" + WCantidad + "',"
            ZSql = ZSql + "'" + WResta + "',"
            ZSql = ZSql + "'" + WFechaord + "',"
            ZSql = ZSql + "'" + WEnvase + "',"
            ZSql = ZSql + "'" + XLote1 + "',"
            ZSql = ZSql + "'" + XCantiLote1 + "',"
            ZSql = ZSql + "'" + XLote2 + "',"
            ZSql = ZSql + "'" + XCantiLote2 + "',"
            ZSql = ZSql + "'" + XLote3 + "',"
            ZSql = ZSql + "'" + XCantiLote3 + "',"
            ZSql = ZSql + "'" + XLote4 + "',"
            ZSql = ZSql + "'" + XCantiLote4 + "',"
            ZSql = ZSql + "'" + XLote5 + "',"
            ZSql = ZSql + "'" + XCantiLote5 + "',"
            ZSql = ZSql + "'" + WCertificado1 + "',"
            ZSql = ZSql + "'" + WCertificado2 + "',"
            ZSql = ZSql + "'" + WEstado1 + "',"
            ZSql = ZSql + "'" + WEstado2 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv1 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv2 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv3 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv4 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv5 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv6 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv7 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv8 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv9 + "',"
            ZSql = ZSql + "'" + ZEstadoEnv10 + "',"
            ZSql = ZSql + "'" + ZCantidadEnv + "',"
            ZSql = ZSql + "'" + ZObservaI + "',"
            ZSql = ZSql + "'" + ZObservaII + "',"
            ZSql = ZSql + "'" + ZObservaIII + "',"
            ZSql = ZSql + "'" + ZObservaIV + "',"
            ZSql = ZSql + "'" + ZVencimiento + "',"
            ZSql = ZSql + "'" + ZOrdVencimiento + "')"
        
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
            Auxiliar(Renglon, 1) = WOrden
            Auxiliar(Renglon, 2) = WResta
            Auxiliar(Renglon, 3) = WArticulo
            Auxiliar(Renglon, 4) = WCantidad
            Auxiliar(Renglon, 5) = WEnvase
            Auxiliar(Renglon, 6) = WCantidad
            
            If WCantiEti <> 0 Then
            
                WProducto = WArticulo
                WKilosEnvase = 0
                ZZCantidad = Cantidad
                
                spEnvase = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    WKilosEnvase = rstEnvase!Kilos
                    rstEnvase.Close
                End If
                
                Call ImpreEtiqueta
            
            End If
                
        End If
            
    Next A
    
    WInforme = Informe.Text
    
    For Da = 1 To Renglon

        Orden = Auxiliar(Da, 1)
        Resta = Val(Auxiliar(Da, 2))
        Articulo = Auxiliar(Da, 3)
        Cantidad = Val(Auxiliar(Da, 4))
        Envase = Auxiliar(Da, 5)
        Cantidad = Auxiliar(Da, 6)
        WTipoOrden = 0
        
        spOrden = "ListaOrden " + "'" + Orden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WTipoOrden = rstOrden!Tipo
            rstOrden.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstArticulo.RecordCount > 0 Then
            
            If WTipoOrden = 2 Then
            
                WCodigo = Articulo
                WPedido = Str$(Resta * -1)
                WLaboratorio = Str$(Cantidad)
                WDate = Date$
                WEnvase = Envase
                WProveedor = Proveedor.Text
                        
                XParam = "'" + WCodigo + "','" _
                    + WPedido + "'"
                                           
                spArticulo = "ModificaArticuloPedido " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                XParam = "'" + WCodigo + "','" _
                    + WLaboratorio + "'"
                                           
                spArticulo = "ModificaArticuloLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                    Else
                    
                WCodigo = Articulo
                WPedido = Str$(rstArticulo!Pedido - Resta)
                WLaboratorio = Str$(rstArticulo!Laboratorio + Cantidad)
                WDate = Date$
                WEnvase = Envase
                WProveedor = Proveedor.Text
                                
                XParam = "'" + WCodigo + "','" _
                    + WPedido + "','" _
                    + WLaboratorio + "','" _
                    + WDate + "','" _
                    + WEnvase + "','" _
                    + WProveedor + "'"
                                           
                spArticulo = "ModificaArticuloInforme " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        End If
        
        If WTipoOrden <> 2 Then
        
            Rem Actualiza costos de todas las empresas
            
            XEmpresa = WEmpresa
        
            XParam = "'" + WCodigo + "','" _
                    + WEnvase + "','" _
                    + WProveedor + "'"
            
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
          
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
               
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spArticulo = "ModificaArticuloVarios1 " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
   
            Call Conecta_Empresa
        
        End If
        
        spOrden = "ListaOrdenArticulo " + "'" + Orden + "','" + Articulo + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
        If rstOrden.RecordCount > 0 Then
            WClave = rstOrden!Clave
            Rem If Cantidad > Resta Then
            Rem     WRecibida = Str$(rstOrden!Recibida + Cantidad)
            Rem         Else
            Rem     WRecibida = Str$(rstOrden!Recibida + Resta)
            Rem End If
            WRecibida = Str$(rstOrden!Recibida + Resta)
            WDate = Date$
            rstOrden.Close
                
            XParam = "'" + WClave + "','" _
                         + WRecibida + "','" _
                         + WDate + "'"
                                           
            spOrden = "ModificaOrdenInforme " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Da
    
    
    If Val(XOrden.Text) >= 800000 Then
    If WTipoOrden = 3 Then
    
        XCodigo = ""
        spMovvar = "ListaMovvarNumero"
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
            With rstMovvar
                .MoveLast
                XCodigo = Str$(rstMovvar!Codigo + 1)
            End With
            rstMovvar.Close
                Else
            XCodigo = "1"
        End If
    
        For Da = 1 To Renglon

            Orden = Auxiliar(Da, 1)
            Resta = Val(Auxiliar(Da, 2))
            Articulo = Auxiliar(Da, 3)
            Cantidad = Auxiliar(Da, 4)
            Envase = Auxiliar(Da, 5)
            Cantidad = Auxiliar(Da, 6)
            Tipo = "M"
            Terminado = ""
            Movi = "E"
            Lote = ""
            
            If Val(Cantidad) <> 0 Then
            
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = Articulo
                    WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Cantidad))
                    WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                    WCosto1 = Str$(rstArticulo!Costo1)
                    WCosto3 = Str$(IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3))
                    WDate = Date$
                    rstArticulo.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WLaboratorio + "','" _
                            + WEntradas + "','" _
                            + WDate + "','" _
                            + WCosto1 + "','" _
                            + WCosto3 + "'"
                                           
                    spArticulo = "ModificaArticuloLaudo " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
            
                ZCodigo = XCodigo
                Call Ceros(ZCodigo, 6)
                ZRenglon = Str$(Da)
                Call Ceros(ZRenglon, 2)
                ZFecha = Fecha.Text
                ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZTipo = "M"
                ZArticulo = Left$(Articulo, 3) + Right$(Articulo, 7)
                ZTerminado = "  -     -   "
                ZCantidad = Cantidad
                zMovi = "E"
                ZTipomov = "0"
                ZObservaciones = "Informe " + Informe.Text + " O/C " + XOrden.Text
                ZClave = ZCodigo + ZRenglon
                ZDate = Date$
                ZMarca = ""
                ZLote = ""
                
                XParam = "'" + ZClave + "','" _
                            + ZCodigo + "','" _
                            + ZRenglon + "','" _
                            + ZFecha + "','" _
                            + ZTipo + "','" _
                            + ZArticulo + "','" _
                            + ZTerminado + "','" _
                            + ZCantidad + "','" _
                            + ZFechaOrd + "','" _
                            + zMovi + "','" _
                            + ZTipomov + "','" _
                            + ZObservaciones + "','" _
                            + ZDate + "','" _
                            + ZMarca + "','" _
                            + ZLote + "'"
                           
                spMovvar = "AltaMovvar " + XParam
                Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                        
            End If
            
        Next Da
        
    End If
    End If
    
    If XObservaEnvases <> "" Then
    
        sTo = "Envases"
        sCC = ""
        sBCC = ""
        sSubject = "Informe de Recepcion de Envases con Observaciones"
        sBody = "Orden :" + XOrden.Text + " - " + _
                "Informe :" + Informe.Text + " - " + _
                "Proveedor :" + DesProveedor.Caption + " - " + _
                XObservaEnvases
    
        ret = Shell("Start.exe " _
            & "mailto:" & """" & sTo & """" _
            & "?Subject=" & """" & sSubject & """" _
            & "&cc=" & """" & sCC & """" _
            & "&bcc=" & """" & sBCC & """" _
            & "&Body=" & """" & sBody & """" _
            & "&File=" & """" & "c:\autoexec.bat" & """" _
            , 0)
                        
    End If
        
    Call Limpia_Click

    WGrilla.Col = 1
    WGrilla.Row = 1
    
    Informe.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
        
End Sub

Private Sub ImpreEtiqueta()

    On Error GoTo WError
    
    OPEN_FILE_Etiqueta
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
        
        WClase = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        WDesProducto = ""
        
        spArticulo = "ConsultaArticulo " + "'" + WProducto + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDesProducto = IIf(IsNull(rstArticulo!Descripcion), "", rstArticulo!Descripcion)
            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            WIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
            WNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
            WEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
            rstArticulo.Close
        End If
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
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
        
        ZCantidad = Int(WCantiEti / 2)
        If ZCantidad * 2 <> WCantiEti Then
            ZCantidad = ZCantidad + 1
        End If
        
        If Val(WEmpresa) <> 5 Then
        
            With rstEtiqueta
                For Da = 1 To ZCantidad
                    .Index = "Codigo"
                    .AddNew
                
                    ZLote = ""
                
                    ZDa = Int((Da - 1) / 2)
                
                    !Codigo = Da
                    !Terminado = WProducto
                    !Lote = 0
                    !Cliente = ""
                    !Cantidad = WKilosEnvase
                    !Nombre = "Fec.Inf.: " + Fecha.Text
                    If ZVencimiento <> "00/00/0000" Then
                        !Impre1 = "Fec.Ven.:" + ZVencimiento
                            Else
                        !Impre1 = ""
                    End If
                    !razon = "Informe : " + Informe.Text
                    !DirEntrega = Str$(WKilosEnvase) + " Kgs."
                    !Clase = WClase
                    !Intervencion = WIntervencion
                    !Naciones = WNaciones
                    !Embalaje = WEmbalaje
                    !Bruto = 0
                    !Neto = ZDa
                    !Observaciones = "CUARENTENA"
                    .Update
                Next Da
            End With

            Listado.WindowTitle = "Emision de Etiquetas"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
        
            Select Case Mid$(WClase, 1, 1)
                Case "3"
                    Listado.ReportFileName = "WEtiVerde3.rpt"
                Case "5"
                    Listado.ReportFileName = "WEtiVerde5.rpt"
                Case "6"
                    Listado.ReportFileName = "WEtiVerde6.rpt"
                Case "8"
                    Listado.ReportFileName = "WEtiVerde8.rpt"
                Case "9"
                    Listado.ReportFileName = "WEtiVerde9.rpt"
                Case Else
                    Listado.ReportFileName = "WEtiVerde.rpt"
            End Select
                
            Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
            Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
            Rem Listado.Connect = Connect()
    
            Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
            Listado.Destination = 1
            Listado.PrinterCopies = 1
            Listado.Action = 1
            
            
                Else
                
            ZBultos = ZZCantidad / WKilosEnvase
            ZBultos = Abs(Int(ZBultos * -1))
        
            With rstEtiqueta
                For Da = 1 To ZCantidad
                    .Index = "Codigo"
                    .AddNew
                
                    ZLote = ""
                
                    ZDa = Int((Da - 1) / 2)
                
                    !Codigo = Da
                    !Neto = ZDa
                    !Lote = 0
                    !Bruto = 0
                    !Cliente = ""
                    !Naciones = WNaciones
                    !Embalaje = WEmbalaje
                    !Intervencion = WIntervencion
                    !Cantidad = 0
                    
                    !Observaciones = "CUARENTENA"
                    !Nombre = Left$(WDesProducto, 30)
                    !Terminado = WProducto
                    !Clase = "Fecha Recepcion : " + Fecha.Text
                    !razon = "Informe : " + Informe.Text
                    
                    !Impre1 = "Cant.Bulto:" + Str$(WKilosEnvase) + " Kg"
                    !DirEntrega = Trim(Str$(ZBultos)) + " bultos"
                    .Update
                Next Da
            End With

            Listado.WindowTitle = "Emision de Etiquetas"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
        
            Listado.ReportFileName = "WEtiVerdeFarma.rpt"
                
            Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
            Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
            Rem Listado.Connect = Connect()
    
            Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
            Listado.Destination = 1
            Listado.PrinterCopies = 1
            Listado.Action = 1
            
        End If
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
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
    
    End If
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WSaldo.Caption = ""
    WResta.Text = ""
    WEnvase.Text = ""
    WEtiqueta.Text = ""
    
    WOrden.SetFocus
    
End Sub

Private Sub Limpia_Click()

    CargaLote.Visible = False
    
    Erase DatosCertificado
    Erase XLote
    
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WSaldo.Caption = ""
    WResta.Text = ""
    WEnvase.Text = ""
    WEtiqueta.Text = ""
    XOrden.Text = ""

    Informe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    
    CertificadoSi.Value = 0
    CertificadoNo.Value = 0
    Certificado2.Text = ""
    EstadoSi.Value = 0
    EstadoNo.Value = 0
    Estado2.Text = ""
    Vencimiento.Text = "  /  /    "
    
    Call Limpia_Grilla
    
    Rem With rstInforme
    Rem .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Informe.Text = !Informe + 1
    Rem             Else
    Rem         Informe.Text = ""
    Rem     End If
    Rem End With
    
    spInforme = "ListaInformeNumero"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveLast
            Informe.Text = rstInforme!Informe + 1
        End With
        rstInforme.Close
            Else
        Informe.Text = "1"
    End If
    
    Erase Verifica
    
    Renglon = 0
    Informe.SetFocus

End Sub

Private Sub WOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            rstOrden.Close
            WArticulo.SetFocus
                Else
            WOrden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        Pasa = "N"
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        If WArticulo.Text = rstOrden!Articulo Then
                            Pasa = "S"
                            Saldo = rstOrden!Cantidad - rstOrden!Recibida
                            If Saldo > 0 Then
                                WSaldo.Caption = Pusing("###,###.##", Str$(Saldo))
                                    Else
                                WSaldo.Caption = ""
                                WArticulo.Text = "  -   -   "
                                Pasa = "N"
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
        
        If Pasa = "S" Then
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                WCantidad.SetFocus
            End If
                        Else
            WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WResta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WResta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WResta.Text) > Val(WSaldo.Caption) Then
            m$ = "La cantidad a descontar supera el saldo de la orden de compra"
            A% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
            WResta.Text = ""
            WResta.SetFocus
              Else
            WResta.Text = Pusing("###,###.##", WResta.Text)
            WEnvase.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEnvase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WTipoOrden = 0
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WTipoOrden = rstOrden!Tipo
            rstOrden.Close
        End If
        
        If WTipoOrden <> 3 And WTipoOrden <> 4 Then
        
            spEnvase = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                rstEnvase.Close
                WEtiqueta.SetFocus
                    Else
                WEnvase.SetFocus
             End If
            
                Else
                
            If WTipoOrden = 3 Or WTipoOrden = 4 Then
            
                Entra = "N"
                spInforme = "ListaInforme " + "'" + Informe.Text + "'"
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    Entra = "S"
                    rstInforme.Close
                End If
                
                WEnvase.Text = "0"
                WEtiqueta.Text = "0"
            
                If Entra = "N" Then
                    If Val(WResta.Text) > Val(WSaldo.Caption) Then
                        m$ = "La cantidad a descontar supera el saldo de la orden de compra"
                        A% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
                        WResta.Text = ""
                        WResta.SetFocus
                        Exit Sub
                            Else
                        If Val(WResta.Text) <> Val(WSaldo.Caption) Then
                            Dife = Str$(Val(WSaldo.Caption) - Val(WResta.Text))
                            T$ = "Ingreso de Informe de recepcion"
                            m$ = "La orden de compra del " + WArticulo.Text + " quedara con un saldo pendiente de entrega de " + Dife + " Kgs" + Chr$(13) + "Confirma este procedimiento"
                            Respuesta% = MsgBox(m$, 32 + 4, T$)
                            If Respuesta% <> 6 Then
                                Exit Sub
                            End If
                            Aviso2.Caption = "LA CANTIDAD DE " + Str$(Dife) + " KGS. Y QUE EL PROVEEDOR"
                            Aviso.Visible = True
                        End If
                    End If
                End If
            
                EstadoEnvI.Value = Val(DatosCertificado(ZLugar, 11))
                EstadoEnvII.Value = Val(DatosCertificado(ZLugar, 12))
                EstadoEnvIII.Value = Val(DatosCertificado(ZLugar, 13))
                EstadoEnvIV.Value = Val(DatosCertificado(ZLugar, 14))
                EstadoEnvV.Value = Val(DatosCertificado(ZLugar, 15))
                EstadoEnvVI.Value = Val(DatosCertificado(ZLugar, 16))
                EstadoEnvVII.Value = Val(DatosCertificado(ZLugar, 17))
                EstadoEnvVIII.Value = Val(DatosCertificado(ZLugar, 18))
                EstadoEnvIX.Value = Val(DatosCertificado(ZLugar, 19))
                EstadoEnvX.Value = Val(DatosCertificado(ZLugar, 20))
            
                CantidadEnv.Text = DatosCertificado(ZLugar, 21)
            
                ObservaI.Text = DatosCertificado(ZLugar, 22)
                ObservaII.Text = DatosCertificado(ZLugar, 23)
                ObservaIII.Text = DatosCertificado(ZLugar, 24)
                ObservaIV.Text = DatosCertificado(ZLugar, 25)
            
                ZEnsayo1 = "0"
                ZEnsayo2 = "0"
                ZEnsayo3 = "0"
                ZEnsayo4 = "0"
                ZEnsayo5 = "0"
            
                DescriEnsayo1.Caption = ""
                DescriEnsayo2.Caption = ""
                DescriEnsayo3.Caption = ""
                DescriEnsayo4.Caption = ""
                DescriEnsayo5.Caption = ""
            
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            
                Sql1 = "Select *"
                Sql2 = " FROM EspecificacionesUnifica"
                Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + WArticulo.Text + "'"
                spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnifica.RecordCount > 0 Then
                    ZEnsayo1 = Str$(rstEspecificacionesUnifica!Ensayo1)
                    ZEnsayo2 = Str$(rstEspecificacionesUnifica!Ensayo2)
                    ZEnsayo3 = Str$(rstEspecificacionesUnifica!Ensayo3)
                    ZEnsayo4 = Str$(rstEspecificacionesUnifica!Ensayo4)
                    ZEnsayo5 = Str$(rstEspecificacionesUnifica!Ensayo5)
                    rstEspecificacionesUnifica.Close
                End If
    
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo1 + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    DescriEnsayo1.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
        
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo2 + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    DescriEnsayo2.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
        
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo3 + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    DescriEnsayo3.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
        
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo4 + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    DescriEnsayo4.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
        
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo5 + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    DescriEnsayo5.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
            
                Call Conecta_Empresa
        
                PantaEnvase.Visible = True
            
                Exit Sub
                
                    Else
                    
                Call Alta_Vector
                Call Ingresa_Click
                WOrden.SetFocus
                    
            End If
            
        End If
    End If
End Sub

Private Sub WEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WTipoOrden = 0
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WTipoOrden = rstOrden!Tipo
            rstOrden.Close
        End If
        
        spInforme = "ListaInforme " + "'" + Informe.Text + "'"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            CertificadoSi.Value = Val(DatosCertificado(ZLugar, 1))
            CertificadoNo.Value = Val(DatosCertificado(ZLugar, 2))
            Certificado2.Text = DatosCertificado(ZLugar, 3)
            
            EstadoSi.Value = Val(DatosCertificado(ZLugar, 4))
            EstadoNo.Value = Val(DatosCertificado(ZLugar, 5))
            Estado2.Text = DatosCertificado(ZLugar, 6)
            
            If DatosCertificado(ZLugar, 7) <> "" Then
                Vencimiento.Text = DatosCertificado(ZLugar, 7)
                    Else
                Vencimiento.Text = "  /  /    "
            End If
            
            PantaCertificado.Visible = True
            Certificado2.SetFocus
            rstInforme.Close
            Exit Sub
        End If
        
        If WTipoOrden <> 2 Then
    
            spEnvase = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvase.RecordCount > 0 Then
                rstEnvase.Close
                If Val(WResta.Text) > Val(WSaldo.Caption) Then
                    m$ = "La cantidad a descontar supera el saldo de la orden de compra"
                    A% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
                    WResta.Text = ""
                    WResta.SetFocus
                    Exit Sub
                        Else
                    If Val(WResta.Text) <> Val(WSaldo.Caption) Then
                        Dife = Str$(Val(WSaldo.Caption) - Val(WResta.Text))
                        T$ = "Ingreso de Informe de recepcion"
                        m$ = "La orden de compra del " + WArticulo.Text + " quedara con un saldo pendiente de entrega de " + Dife + " Kgs" + Chr$(13) + "Confirma este procedimiento"
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% <> 6 Then
                            Exit Sub
                        End If
                        Aviso2.Caption = "LA CANTIDAD DE " + Str$(Dife) + " KGS. Y QUE EL PROVEEDOR"
                        Aviso.Visible = True
                    End If
                End If
                
                If Left$(WArticulo.Text, 2) = "DY" Or Left$(WArticulo.Text, 2) = "DW" Or Left$(WArticulo.Text, 2) = "DS" Or Left$(WArticulo.Text, 2) = "DQ" Then
                
                    WLugar = WGrilla.Row
                            
                    If Val(XLote(WLugar, 2)) <> 0 Then
                        WLote1.Text = XLote(WLugar, 1)
                        WCanti1.Text = XLote(WLugar, 2)
                            Else
                        WLote1.Text = ""
                        WCanti1.Text = ""
                    End If
                    If Val(XLote(WLugar, 4)) <> 0 Then
                        WLote2.Text = XLote(WLugar, 3)
                        WCanti2.Text = XLote(WLugar, 4)
                            Else
                        WLote2.Text = ""
                        WCanti2.Text = ""
                    End If
                    If Val(XLote(WLugar, 6)) <> 0 Then
                        WLote3.Text = XLote(WLugar, 5)
                        WCanti3.Text = XLote(WLugar, 6)
                            Else
                        WLote3.Text = ""
                        WCanti3.Text = ""
                    End If
                    If Val(XLote(WLugar, 8)) <> 0 Then
                        WLote4.Text = XLote(WLugar, 7)
                        WCanti4.Text = XLote(WLugar, 8)
                            Else
                        WLote4.Text = ""
                        WCanti4.Text = ""
                    End If
                    If Val(XLote(WLugar, 10)) <> 0 Then
                        WLote5.Text = XLote(WLugar, 9)
                        WCanti5.Text = XLote(WLugar, 10)
                            Else
                        WLote5.Text = ""
                        WCanti5.Text = ""
                    End If
                
                    Rem CargaLote.Visible = True
                    Rem WLote1.SetFocus
                    
                    CertificadoSi.Value = Val(DatosCertificado(ZLugar, 1))
                    CertificadoNo.Value = Val(DatosCertificado(ZLugar, 2))
                    Certificado2.Text = DatosCertificado(ZLugar, 3)
            
                    EstadoSi.Value = Val(DatosCertificado(ZLugar, 4))
                    EstadoNo.Value = Val(DatosCertificado(ZLugar, 5))
                    Estado2.Text = DatosCertificado(ZLugar, 6)
                    
                    If DatosCertificado(ZLugar, 7) <> "" Then
                        Vencimiento.Text = DatosCertificado(ZLugar, 7)
                            Else
                        Vencimiento.Text = "  /  /    "
                    End If
                    
                    PantaCertificado.Visible = True
                    Certificado2.SetFocus
                    
                    
                        Else
                        
                    Rem Call Alta_Vector
                    Rem Call Ingresa_Click
                    Rem WOrden.SetFocus
                    
                    CertificadoSi.Value = Val(DatosCertificado(ZLugar, 1))
                    CertificadoNo.Value = Val(DatosCertificado(ZLugar, 2))
                    Certificado2.Text = DatosCertificado(ZLugar, 3)
            
                    EstadoSi.Value = Val(DatosCertificado(ZLugar, 4))
                    EstadoNo.Value = Val(DatosCertificado(ZLugar, 5))
                    Estado2.Text = DatosCertificado(ZLugar, 6)
                    
                    If DatosCertificado(ZLugar, 7) <> "" Then
                        Vencimiento.Text = DatosCertificado(ZLugar, 7)
                            Else
                        Vencimiento.Text = "  /  /    "
                    End If
                    
                    PantaCertificado.Visible = True
                    Certificado2.SetFocus
                    
                End If
                
                    Else
                    
                WEtiqueta.SetFocus
                
            End If
            
                Else
                
            Rem Call Alta_Vector
            Rem Call Ingresa_Click
            Rem WOrden.SetFocus
            
            CertificadoSi.Value = Val(DatosCertificado(ZLugar, 1))
            CertificadoNo.Value = Val(DatosCertificado(ZLugar, 2))
            Certificado2.Text = DatosCertificado(ZLugar, 3)
            
            EstadoSi.Value = Val(DatosCertificado(ZLugar, 4))
            EstadoNo.Value = Val(DatosCertificado(ZLugar, 5))
            Estado2.Text = DatosCertificado(ZLugar, 6)
            
            If DatosCertificado(ZLugar, 7) <> "" Then
                Vencimiento.Text = DatosCertificado(ZLugar, 7)
                    Else
                Vencimiento.Text = "  /  /    "
            End If
            
            PantaCertificado.Visible = True
            Certificado2.SetFocus
            
        End If
    End If
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti1.SetFocus
    End If
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = Val(WCanti1.Text) + Val(WCanti2.Text) + Val(WCanti3.Text) + Val(WCanti4.Text) + Val(WCanti5.Text)
        If Compara = Val(WCantidad.Text) And Val(WCanti1.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti2.SetFocus
    End If
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = Val(WCanti1.Text) + Val(WCanti2.Text) + Val(WCanti3.Text) + Val(WCanti4.Text) + Val(WCanti5.Text)
        If Compara = Val(WCantidad.Text) And Val(WCanti2.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti3.SetFocus
    End If
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = Val(WCanti1.Text) + Val(WCanti2.Text) + Val(WCanti3.Text) + Val(WCanti4.Text) + Val(WCanti5.Text)
        If Compara = Val(WCantidad.Text) And Val(WCanti3.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote4.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti4.SetFocus
    End If
End Sub

Private Sub WCanti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = Val(WCanti1.Text) + Val(WCanti2.Text) + Val(WCanti3.Text) + Val(WCanti4.Text) + Val(WCanti5.Text)
        If Compara = Val(WCantidad.Text) And Val(WCanti4.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti5.SetFocus
    End If
End Sub

Private Sub WCanti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = Val(WCanti1.Text) + Val(WCanti2.Text) + Val(WCanti3.Text) + Val(WCanti4.Text) + Val(WCanti5.Text)
        If Compara = Val(WCantidad.Text) Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = WProveedor
                DesProveedor.Caption = RstProveedor!Nombre
                RstProveedor.Close
                    Else
                WPasa = "N"
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            XOrden.Text = WIndice.List(Indice)
            Call Proceso
            XOrden.SetFocus
            
        Case 2
            Indice = Pantalla.ListIndex
            WEnvase.Text = WIndice.List(Indice)
            Call WEnvase_KeyPress(13)
            
        Case Else
        
    End Select
    
End Sub


Private Sub Form_Load()
    
    Call Limpia_Grilla
    
    Erase DatosCertificado
    Erase XLote

    CargaLote.Visible = False
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WSaldo.Caption = ""
    WResta.Text = ""
    WEnvase.Text = ""
    WEtiqueta.Text = ""
    XOrden.Text = ""

    Informe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    
    CertificadoSi.Value = 0
    CertificadoNo.Value = 0
    Certificado2.Text = ""
    EstadoSi.Value = 0
    EstadoNo.Value = 0
    Estado2.Text = ""
    Vencimiento.Text = "  /  /    "

    spInforme = "ListaInformeNumero"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveLast
            Informe.Text = rstInforme!Informe + 1
        End With
        rstInforme.Close
            Else
        Informe.Text = "1"
    End If

    Rem With rstInforme
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Informe.Text = !Informe + 1
    Rem             Else
    Rem         Informe.Text = ""
    Rem     End If
    Rem End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgInforme.Caption = "Ingreso de Informe de Recepcion :  " + !Nombre
        End If
    End With

    WGrilla.Col = 1
    WGrilla.Row = 1
    
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    Call Limpia_Grilla
    
    Renglon = 0
    Erase Auxiliar
    
    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    WGrilla.Row = Renglon
                
                    WGrilla.Col = 1
                    WGrilla.Text = rstInforme!Orden
                
                    WGrilla.Col = 2
                    WGrilla.Text = rstInforme!Articulo
                    Auxi1 = rstInforme!Articulo
                
                    WGrilla.Col = 4
                    WGrilla.Text = Pusing("###,###.##", Val(rstInforme!Cantidad))
                
                    WGrilla.Col = 6
                    WGrilla.Text = Pusing("###,###.##", Val(rstInforme!Resta))
                    
                    WGrilla.Col = 7
                    WGrilla.Text = ""
                    WGrilla.Text = rstInforme!Envase
                    
                    WCertificado1 = IIf(IsNull(rstInforme!Certificado1), "0", rstInforme!Certificado1)
                    If WCertificado1 = 0 Then
                        ZCertificadoNo = 1
                        ZCertificadoSi = 0
                            Else
                        ZCertificadoNo = 0
                        ZCertificadoSi = 1
                    End If
                    ZCertificado2 = IIf(IsNull(rstInforme!Certificado2), "", rstInforme!Certificado2)
            
                    WEstado1 = IIf(IsNull(rstInforme!Estado1), "0", rstInforme!Estado1)
                    If WEstado1 = 0 Then
                        ZEstadoNo = 1
                        ZEstadoSi = 0
                            Else
                        ZEstadoNo = 0
                        ZEstadoSi = 1
                    End If
                    
                    ZEstado2 = IIf(IsNull(rstInforme!Estado2), "", rstInforme!Estado2)
                    ZVencimiento = IIf(IsNull(rstInforme!fechavencimiento), "  /  /    ", rstInforme!fechavencimiento)
                            
                    DatosCertificado(Renglon, 1) = Str$(ZCertificadoSi)
                    DatosCertificado(Renglon, 2) = Str$(ZCertificadoNo)
                    DatosCertificado(Renglon, 3) = ZCertificado2
            
                    DatosCertificado(Renglon, 4) = Str$(ZEstadoSi)
                    DatosCertificado(Renglon, 5) = Str$(ZEstadoNo)
                    DatosCertificado(Renglon, 6) = ZEstado2
                    
                    DatosCertificado(Renglon, 7) = ZVencimiento
                    
                    ZEstadoEnv1 = IIf(IsNull(rstInforme!EstadoEnvI), "0", rstInforme!EstadoEnvI)
                    ZEstadoEnv2 = IIf(IsNull(rstInforme!EstadoEnvII), "0", rstInforme!EstadoEnvII)
                    ZEstadoEnv3 = IIf(IsNull(rstInforme!EstadoEnvIII), "0", rstInforme!EstadoEnvIII)
                    ZEstadoEnv4 = IIf(IsNull(rstInforme!EstadoEnvIV), "0", rstInforme!EstadoEnvIV)
                    ZEstadoEnv5 = IIf(IsNull(rstInforme!EstadoEnvV), "0", rstInforme!EstadoEnvV)
                    ZEstadoEnv6 = IIf(IsNull(rstInforme!EstadoEnvVI), "0", rstInforme!EstadoEnvVI)
                    ZEstadoEnv7 = IIf(IsNull(rstInforme!EstadoEnvVII), "0", rstInforme!EstadoEnvVII)
                    ZEstadoEnv8 = IIf(IsNull(rstInforme!EstadoEnvVIII), "0", rstInforme!EstadoEnvVIII)
                    ZEstadoEnv9 = IIf(IsNull(rstInforme!EstadoEnvIX), "0", rstInforme!EstadoEnvIX)
                    ZEstadoEnv10 = IIf(IsNull(rstInforme!EstadoEnvX), "0", rstInforme!EstadoEnvX)
                    ZCantidadEnv = IIf(IsNull(rstInforme!CantidadEnv), "0", rstInforme!CantidadEnv)
                    ZObservaI = IIf(IsNull(rstInforme!ObservaI), "", rstInforme!ObservaI)
                    ZObservaII = IIf(IsNull(rstInforme!ObservaII), "", rstInforme!ObservaII)
                    ZObservaIII = IIf(IsNull(rstInforme!ObservaIII), "", rstInforme!ObservaIII)
                    ZObservaIV = IIf(IsNull(rstInforme!ObservaIV), "", rstInforme!ObservaIV)
                    
                    DatosCertificado(Renglon, 11) = ZEstadoEnv1
                    DatosCertificado(Renglon, 12) = ZEstadoEnv2
                    DatosCertificado(Renglon, 13) = ZEstadoEnv3
                    DatosCertificado(Renglon, 14) = ZEstadoEnv4
                    DatosCertificado(Renglon, 15) = ZEstadoEnv5
                    DatosCertificado(Renglon, 16) = ZEstadoEnv6
                    DatosCertificado(Renglon, 17) = ZEstadoEnv7
                    DatosCertificado(Renglon, 18) = ZEstadoEnv8
                    DatosCertificado(Renglon, 19) = ZEstadoEnv9
                    DatosCertificado(Renglon, 20) = ZEstadoEnv10
                    DatosCertificado(Renglon, 21) = ZCantidadEnv
                    DatosCertificado(Renglon, 22) = ZObservaI
                    DatosCertificado(Renglon, 23) = ZObservaII
                    DatosCertificado(Renglon, 24) = ZObservaIII
                    DatosCertificado(Renglon, 25) = ZObservaIV
                    
                    spEnvases = "ConsultaEnvases " + "'" + Str$(rstInforme!Envase) + "'"
                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvases.RecordCount > 0 Then
                        WGrilla.Col = 8
                        WGrilla.Text = rstEnvases!Descripcion
                            Else
                        WGrilla.Col = 8
                        WGrilla.Text = ""
                    End If
                    
                    Auxiliar(Renglon, 1) = rstInforme!Articulo
                    Auxiliar(Renglon, 2) = rstInforme!Envase
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        WGrilla.Row = Renglon
                
        spArticulo = "ConsultaArticulo " + "'" + Auxiliar(Renglon, 1) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WGrilla.Col = 3
            WGrilla.Text = rstArticulo!Descripcion
            WOrden.SetFocus
            rstArticulo.Close
        End If
    Next Da

    WOrden.SetFocus
    Exit Sub

WError:

    Resume Next
    

End Sub

Private Sub Alta_Vector()

    Entra = "S"
    
    If Val(WLinea.Text) = 0 Then
        For Da = 1 To 100
            If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text Then
                Entra = "N"
                Exit For
            End If
        Next Da
            Else
        Lugar = WGrilla.Row
        For Da = 1 To 100
            If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text And Da <> Lugar Then
                Entra = "N"
                Exit For
            End If
        Next Da
    End If
    
    If Entra = "N" Then
        m$ = "El articulo ya se encuentra dado de alta en el informe de recepcion"
        A% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
    End If
                
    If Entra = "S" Then

    If Val(WLinea.Text) = 0 Then
    
            Renglon = Renglon + 1
            WGrilla.Row = Renglon
                
            WAnterior = WGrilla.Row
            
            WGrilla.Col = 1
            WGrilla.Text = WOrden.Text
            
            WGrilla.Col = 2
            WGrilla.Text = WArticulo.Text
            
            WGrilla.Col = 3
            WGrilla.Text = WDescripcion.Caption
                
            WGrilla.Col = 4
            WGrilla.Text = Pusing("###,###.##", WCantidad.Text)
                
            WGrilla.Col = 5
            WGrilla.Text = Pusing("###,###.##", WSaldo.Caption)
                
            WGrilla.Col = 6
            WGrilla.Text = Pusing("###,###.##", WResta.Text)
            
            WGrilla.Col = 7
            WGrilla.Text = WEnvase.Text
            
            spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WGrilla.Col = 8
                WGrilla.Text = rstEnvases!Descripcion
                rstEnvases.Close
                    Else
                WGrilla.Col = 8
                WGrilla.Text = ""
            End If
            
            WGrilla.Col = 9
            WGrilla.Text = WEtiqueta.Text
            
            Verifica(Renglon, 1) = WArticulo.Text
            Verifica(Renglon, 2) = WOrden.Text
            
            XLote(Renglon, 1) = WLote1.Text
            XLote(Renglon, 2) = WCanti1.Text
            XLote(Renglon, 3) = WLote2.Text
            XLote(Renglon, 4) = WCanti2.Text
            XLote(Renglon, 5) = WLote3.Text
            XLote(Renglon, 6) = WCanti3.Text
            XLote(Renglon, 7) = WLote4.Text
            XLote(Renglon, 8) = WCanti4.Text
            XLote(Renglon, 9) = WLote5.Text
            XLote(Renglon, 10) = WCanti5.Text
            XLote(Renglon, 15) = WEnvase.Text
            
                Else
                
            WGrilla.Row = Val(WLinea.Text)
                
            WAnterior = WGrilla.Row
            
            WGrilla.Col = 1
            WGrilla.Text = WOrden.Text
            
            WGrilla.Col = 2
            WGrilla.Text = WArticulo.Text
            
            WGrilla.Col = 3
            WGrilla.Text = WDescripcion.Caption
                
            WGrilla.Col = 4
            WGrilla.Text = Pusing("###,###.##", WCantidad.Text)
            
            WGrilla.Col = 5
            WGrilla.Text = Pusing("###,###.##", WSaldo.Caption)
            
            WGrilla.Col = 6
            WGrilla.Text = Pusing("###,###.##", WResta.Text)
            
            WGrilla.Col = 7
            WGrilla.Text = WEnvase.Text
            
            spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WGrilla.Col = 8
                WGrilla.Text = rstEnvases!Descripcion
                rstEnvases.Close
                    Else
                WGrilla.Col = 8
                WGrilla.Text = ""
            End If
            
            WGrilla.Col = 9
            WGrilla.Text = WEtiqueta.Text
            
            Lugar = WGrilla.Row
            Verifica(Lugar, 1) = WArticulo.Text
            Verifica(Lugar, 2) = WOrden.Text
            
            XLote(Lugar, 1) = WLote1.Text
            XLote(Lugar, 2) = WCanti1.Text
            XLote(Lugar, 3) = WLote2.Text
            XLote(Lugar, 4) = WCanti2.Text
            XLote(Lugar, 5) = WLote3.Text
            XLote(Lugar, 6) = WCanti3.Text
            XLote(Lugar, 7) = WLote4.Text
            XLote(Lugar, 8) = WCanti4.Text
            XLote(Lugar, 9) = WLote5.Text
            XLote(Lugar, 10) = WCanti5.Text
            XLote(Lugar, 15) = WEnvase.Text
            
    End If
    
    End If

End Sub

Private Sub Informe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Entra = "N"
        
        spInforme = "ListaInforme " + "'" + Informe.Text + "'"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            Entra = "S"
            Fecha.Text = rstInforme!Fecha
            Proveedor.Text = rstInforme!Proveedor
            Remito.Text = rstInforme!Remito
            XOrden.Text = rstInforme!Orden
            rstInforme.Close
        End If
        
        If Entra = "S" Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                DesProveedor.Caption = RstProveedor!Nombre
                RstProveedor.Close
            End If
            Call Proceso_Click
                Else
            WInforme = Informe.Text
            Call Limpia_Click
            Informe.Text = WInforme
            Fecha.SetFocus
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            XOrden.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                    Proveedor.Text = RstProveedor!Proveedor
                    DesProveedor.Caption = RstProveedor!Nombre
                    RstProveedor.Close
                    Remito.SetFocus
                        Else
                    Proveedor.Text = Claveven$
                    Proveedor.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WOrden.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(RstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(RstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + RstProveedor!Nombre
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

End Sub


Private Sub XOrden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XOrden) <> 0 Then
        
            WTipoOrden = 0
            spOrden = "ListaOrden " + "'" + XOrden.Text + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WTipoOrden = rstOrden!Tipo
                rstOrden.Close
            End If
        
            If WTipoOrden <> 4 Then Exit Sub
        
            Call Proceso
            WOrden.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Proceso()

    If Val(XOrden) <> 0 Then
        
        Call Limpia_Grilla
    
        Renglon = 0
        Erase Auxiliar
    
        spOrden = "ListaOrden " + "'" + XOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Saldo = rstOrden!Cantidad - rstOrden!Recibida
                        
                        If Saldo > 0 Then
                    
                            Renglon = Renglon + 1
                            WGrilla.Row = Renglon
                
                            WGrilla.Col = 1
                            WGrilla.Text = XOrden.Text
                
                            WGrilla.Col = 2
                            WGrilla.Text = rstOrden!Articulo
                            Auxi1 = rstOrden!Articulo
                    
                            Saldo = rstOrden!Cantidad - rstOrden!Recibida
                
                            WGrilla.Col = 4
                            WGrilla.Text = Pusing("###,###.##", Str$(Saldo))
                            
                            WGrilla.Col = 5
                            WGrilla.Text = Pusing("###,###.##", Str$(Saldo))
                
                            WGrilla.Col = 6
                            WGrilla.Text = Pusing("###,###.##", Str$(Saldo))
                            
                    
                            Auxiliar(Renglon, 1) = rstOrden!Articulo
                        
                        End If
                    
                        Proveedor.Text = rstOrden!Proveedor
                        
                        Verifica(Renglon, 1) = rstOrden!Articulo
                        Verifica(Renglon, 2) = rstOrden!Articulo
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
    
        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
        End If
    
    
        WRenglon = Renglon
        Renglon = 0
    
        For Da = 1 To WRenglon
    
            Renglon = Renglon + 1
            WGrilla.Row = Renglon
        
            spArticulo = "ConsultaArticulo " + "'" + Auxiliar(Renglon, 1) + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WGrilla.Col = 3
                WGrilla.Text = rstArticulo!Descripcion
                rstArticulo.Close
                WOrden.SetFocus
            End If
        Next Da

    End If
    
End Sub

Private Sub Limpia_Grilla()

    WGrilla.Clear
    WGrilla.Font.Bold = True
    
    WGrilla.FixedCols = 1
    WGrilla.Cols = 10
    WGrilla.FixedRows = 1
    WGrilla.Rows = 101
    
    WGrilla.ColWidth(0) = 200
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        Select Case Ciclo
            Case 1
                WGrilla.Text = "Orden"
                WGrilla.ColWidth(Ciclo) = 1000
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WGrilla.Text = "Producto"
                WGrilla.ColWidth(Ciclo) = 1300
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WGrilla.Text = "Descripcion"
                WGrilla.ColWidth(Ciclo) = 2500
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WGrilla.Text = "Cant.Ing."
                WGrilla.ColWidth(Ciclo) = 1100
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WGrilla.Text = "Saldo O/C"
                WGrilla.ColWidth(Ciclo) = 1100
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WGrilla.Text = "Descon.OC"
                WGrilla.ColWidth(Ciclo) = 1100
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WGrilla.Text = "Envase"
                WGrilla.ColWidth(Ciclo) = 800
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WGrilla.Text = "Desc."
                WGrilla.ColWidth(Ciclo) = 900
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                WGrilla.Text = "Etiq"
                WGrilla.ColWidth(Ciclo) = 900
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        WTitulo(Ciclo).Text = WGrilla.Text
        WTitulo(Ciclo).Left = WGrilla.CellLeft + WGrilla.Left
        WTitulo(Ciclo).Top = WGrilla.CellTop + WGrilla.Top
        WTitulo(Ciclo).Width = WGrilla.CellWidth
        WTitulo(Ciclo).Height = WGrilla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WGrilla.Cols - 1
        WAncho = WAncho + WGrilla.ColWidth(Ciclo)
    Next Ciclo
    WGrilla.Width = WAncho

    ' Size the columns.
    Font.Name = WGrilla.Font.Name
    Font.Size = WGrilla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WGrilla.AllowUserResizing = flexResizeBoth
    
    WGrilla.Visible = True
    
    WGrilla.Col = 1
    WGrilla.Row = 1
    
End Sub


Private Sub Command1_Click()

    For A = 1 To 100
        
        WRow = A
        WGrilla.Row = WRow
                    
        WGrilla.Col = 1
        Orden = WGrilla.Text
                   
        WGrilla.Col = 2
        Articulo = UCase(WGrilla.Text)
                    
        WGrilla.Col = 4
        Cantidad = Val(WGrilla.Text)
                    
        WGrilla.Col = 5
        Saldo = Val(WGrilla.Text)
                    
        WGrilla.Col = 6
        Resta = Val(WGrilla.Text)
            
        WGrilla.Col = 7
        Envase = Val(WGrilla.Text)
        
        WGrilla.Col = 9
        WCantiEti = Val(WGrilla.Text)
            
        WLugar = WGrilla.Row
                    
        XLote1 = XLote(WLugar, 1)
        XLote2 = XLote(WLugar, 3)
        XLote3 = XLote(WLugar, 5)
        XLote4 = XLote(WLugar, 7)
        XLote5 = XLote(WLugar, 9)
        XCantiLote1 = XLote(WLugar, 2)
        XCantiLote2 = XLote(WLugar, 4)
        XCantiLote3 = XLote(WLugar, 6)
        XCantiLote4 = XLote(WLugar, 8)
        XCantiLote5 = XLote(WLugar, 10)
        
        ZCertificadoSi = Val(DatosCertificado(WLugar, 1))
        ZCertificadoNo = Val(DatosCertificado(WLugar, 2))
        ZCertificado2 = DatosCertificado(WLugar, 3)
            
        ZEstadoSi = Val(DatosCertificado(WLugar, 4))
        ZEstadoNo = Val(DatosCertificado(WLugar, 5))
        ZEstado2 = DatosCertificado(WLugar, 6)
        
        ZVencimiento = DatosCertificado(WLugar, 7)
        ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
        
        If ZCertificadoNo = 1 Then
            WCertificado1 = "0"
        End If
    
        If ZCertificadoSi = 1 Then
            WCertificado1 = "1"
        End If
    
        If ZEstadoNo = 1 Then
            WEstado1 = "0"
        End If
    
        If ZEstadoSi = 1 Then
            WEstado1 = "1"
        End If
        
        ZEstadoEnv1 = DatosCertificado(WLugar, 11)
        ZEstadoEnv2 = DatosCertificado(WLugar, 12)
        ZEstadoEnv3 = DatosCertificado(WLugar, 13)
        ZEstadoEnv4 = DatosCertificado(WLugar, 14)
        ZEstadoEnv5 = DatosCertificado(WLugar, 15)
        ZEstadoEnv6 = DatosCertificado(WLugar, 16)
        ZEstadoEnv7 = DatosCertificado(WLugar, 17)
        ZEstadoEnv8 = DatosCertificado(WLugar, 18)
        ZEstadoEnv9 = DatosCertificado(WLugar, 19)
        ZEstadoEnv10 = DatosCertificado(WLugar, 20)
        ZCantidadEnv = DatosCertificado(WLugar, 21)
        ZObservaI = DatosCertificado(WLugar, 22)
        ZObservaII = DatosCertificado(WLugar, 23)
        ZObservaIII = DatosCertificado(WLugar, 24)
        ZObservaIV = DatosCertificado(WLugar, 25)
        
        WCertificado2 = ZCertificado2
        WEstado2 = ZEstado2
                   
        If Articulo <> "" Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Informe.Text)
            Call Ceros(Auxi1, 6)
                
            WClave = Auxi1 + Auxi
            WInforme = Informe.Text
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WProveedor = Proveedor.Text
            WRemito = Remito.Text
            WOrden = Orden
            WArticulo = Articulo
            WCantidad = Cantidad
            WResta = Resta
            WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WDate = Date$
            WEnvase = Envase
        
            WCantiEti = 3
        
            If WCantiEti <> 0 Then
            
                WProducto = WArticulo
                WKilosEnvase = 0
                ZZCantidad = Cantidad
                
                spEnvase = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    WKilosEnvase = rstEnvase!Kilos
                    rstEnvase.Close
                End If
                
                Call ImpreEtiqueta
            
            End If
                
        End If
            
    Next A
    
End Sub

Private Sub ObservaI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI.Text = ""
    End If
End Sub

Private Sub ObservaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII.Text = ""
    End If
End Sub

Private Sub ObservaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIII.Text = ""
    End If
End Sub

Private Sub ObservaIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIV.Text = ""
    End If
End Sub


