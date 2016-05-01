VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCalificaGrillaEnvases 
   AutoRedraw      =   -1  'True
   Caption         =   "Actualizacion de  evaluacion  Semestral de Proveedores"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   11880
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
      Height          =   4335
      Left            =   720
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   11295
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
         TabIndex        =   43
         Top             =   3600
         Width           =   1095
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
         TabIndex        =   42
         Top             =   1320
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
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   40
         Top             =   720
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
         TabIndex        =   39
         Top             =   720
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
         TabIndex        =   38
         Top             =   1920
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
         TabIndex        =   37
         Top             =   1920
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
         TabIndex        =   36
         Top             =   2520
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
         TabIndex        =   35
         Top             =   2520
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
         TabIndex        =   34
         Top             =   3120
         Width           =   495
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
         TabIndex        =   33
         Top             =   3120
         Width           =   615
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
         TabIndex        =   32
         Text            =   " "
         Top             =   3720
         Width           =   1095
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
         TabIndex        =   31
         Top             =   1320
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
         TabIndex        =   30
         Top             =   1920
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
         TabIndex        =   29
         Top             =   2520
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
         TabIndex        =   28
         Top             =   3120
         Width           =   5415
      End
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
         TabIndex        =   27
         Top             =   720
         Width           =   5415
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
         TabIndex        =   51
         Top             =   1320
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
         TabIndex        =   50
         Top             =   720
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
         TabIndex        =   49
         Top             =   1920
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
         TabIndex        =   48
         Top             =   2520
         Width           =   3135
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
         TabIndex        =   47
         Top             =   3120
         Width           =   3135
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
         TabIndex        =   46
         Top             =   3720
         Width           =   2655
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
         TabIndex        =   45
         Top             =   360
         Width           =   615
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
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame PantaInforme 
      Height          =   5175
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   11415
      Begin MSFlexGridLib.MSFlexGrid WGrilla 
         Height          =   3615
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6376
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Image CancelaInforme 
         Height          =   480
         Left            =   5040
         MouseIcon       =   "calificagrillaEnvases.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "calificagrillaEnvases.frx":030A
         ToolTipText     =   "Salida"
         Top             =   4440
         Width           =   480
      End
   End
   Begin VB.ComboBox TipoActualiza 
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
      Left            =   5880
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame PantaAnaliza 
      Height          =   6735
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   11535
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   4095
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7223
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.CommandButton BotonCerrar 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   6120
         Width           =   1095
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
      Index           =   11
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1920
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
      Index           =   9
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   375
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
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Top             =   3240
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
      Left            =   2160
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   2640
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
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
      Index           =   3
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
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
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   8
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
      Index           =   6
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   7
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
      Index           =   7
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   6
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
      Index           =   8
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6855
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12091
      _Version        =   327680
      BackColor       =   16316587
   End
   Begin MSMask.MaskEdBox Hasta 
      Height          =   300
      Left            =   4320
      TabIndex        =   1
      Top             =   120
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
      Left            =   1440
      TabIndex        =   0
      Top             =   120
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
      Index           =   10
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Cancela 
      Height          =   480
      Left            =   6000
      MouseIcon       =   "calificagrillaEnvases.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "calificagrillaEnvases.frx":0E56
      ToolTipText     =   "Salida"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Actualiza 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "calificagrillaEnvases.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "calificagrillaEnvases.frx":19A2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCalificaGrillaEnvases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim WVector(1000, 20) As String
Dim WDevuelta As String
Dim WLiberada As String
Dim WPartida1 As String
Dim WPartida2 As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim ZProveedor(5000, 10) As String
Dim CargaEmpresa(10, 2) As String
Dim ZImpre7 As String
Dim ZImpre8 As String
Dim ZImpre9 As String
Dim XProveedor As String
Dim DatosCertificado(100, 30) As String
Dim Auxiliar(100, 10) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Procesa()

    If Desde.Text = "  /  /    " Or Hasta.Text = "  /  /    " Then
        Exit Sub
    End If

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

    Erase ZProveedor
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where TipoProv = 2"
    ZSql = ZSql + " Order by Nombre"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Rem If !Nombre <= "ATAN" Then
                    
                    ZLugar = ZLugar + 1
                    ZProveedor(ZLugar, 1) = !Proveedor
                    
                    Rem End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProveedor.Close
    End If
    
   Rem  ZLugar = 50
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    XEmpresa = WEmpresa
    
    For ZCiclo = 1 To 9
    
        WEmpresa = CargaEmpresa(ZCiclo, 1)
        txtOdbc = CargaEmpresa(ZCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        For WCiclo = 1 To ZLugar
        
            Erase WVector
            Lugar = 0
                    
            XParam = "'" + WDesde + "','" _
                         + WHasta + "','" _
                         + ZProveedor(WCiclo, 1) + "','" _
                         + ZProveedor(WCiclo, 1) + "'"

            spInforme = "ListaInformeListado" + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
            
                With rstInforme
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                            Lugar = Lugar + 1
                
                            WCertificado = IIf(IsNull(rstInforme!Certificado1), "1", rstInforme!Certificado1)
                            WEstado = IIf(IsNull(rstInforme!Estado1), "1", rstInforme!Estado1)
                            
                            WVector(Lugar, 1) = rstInforme!Articulo
                            WVector(Lugar, 2) = rstInforme!Orden
                            WVector(Lugar, 3) = rstInforme!FechaOrd
                            WVector(Lugar, 4) = rstInforme!Clave
                            WVector(Lugar, 5) = WCertificado
                            WVector(Lugar, 6) = WEstado
                            WVector(Lugar, 7) = rstInforme!Informe
                            WVector(Lugar, 8) = rstInforme!Cantidad
                            WVector(Lugar, 9) = rstInforme!Fecha
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                    End If
                End With
        
                rstInforme.Close
        
            End If
    
            For Ciclo = 1 To Lugar
            
                Rem Calcula las diferencias de fecha entre la
                Rem Orden de compra y el informe de recepcion
    
                WArticulo = WVector(Ciclo, 1)
                WOrden = WVector(Ciclo, 2)
                WFecha = WVector(Ciclo, 3)
                WClave = WVector(Ciclo, 4)
                WCertificado = WVector(Ciclo, 5)
                WEstado = WVector(Ciclo, 6)
                WInforme = WVector(Ciclo, 7)
                WCantidad = WVector(Ciclo, 8)
                
                
                
                XFecha = "  /  /    "
                XOrdFecha = "00000000"
        
                XParam = "'" + WOrden + "','" _
                             + WArticulo + "'"

                spOrden = "ListaOrdenArticulo" + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
        
                    XOrdFecha = Right$(rstOrden!Fecha2, 4) + Mid$(rstOrden!Fecha2, 4, 2) + Left$(rstOrden!Fecha2, 2)
                    XFecha = rstOrden!Fecha2
            
                    rstOrden.Close
            
                End If
        
                BAse1 = (Val(Left$(XOrdFecha, 4)) * 365) + (Val(Mid$(XOrdFecha, 5, 2)) * 30) + (Val(Right$(XOrdFecha, 2)) * 1)
                Base2 = (Val(Left$(WFecha, 4)) * 365) + (Val(Mid$(WFecha, 5, 2)) * 30) + (Val(Right$(WFecha, 2)) * 1)
        
                Dife = Base2 - BAse1
                If Dife < 0 Then
                    Dife = 0
                End If
                
                ZProveedor(WCiclo, 2) = Str$(Val(ZProveedor(WCiclo, 2)) + 1)
                
                If Val(WCertificado) = 1 Then
                    ZProveedor(WCiclo, 6) = Str$(Val(ZProveedor(WCiclo, 6)) + 1)
                End If
                
                If Val(WEstado) = 1 Then
                    ZProveedor(WCiclo, 7) = Str$(Val(ZProveedor(WCiclo, 7)) + 1)
                End If
                
                aa = Dife
                If Dife > 100 Then Dife = 0
                ZProveedor(WCiclo, 8) = Str$(Val(ZProveedor(WCiclo, 8)) + Dife)
                
                
                
                Rem Calcula las diferencias de fecha entre la
                Rem Orden de compra y el informe de recepcion
                
                WArticulo = WVector(Ciclo, 1)
                WOrden = WVector(Ciclo, 2)
                WFecha = WVector(Ciclo, 3)
                WClave = WVector(Ciclo, 4)
                WCertificado = WVector(Ciclo, 5)
                WEstado = WVector(Ciclo, 6)
                WInforme = WVector(Ciclo, 7)
                WCantidad = Val(WVector(Ciclo, 8))
                WFecha = WVector(Ciclo, 9)
                
                WLiberada = ""
                WDevuelta = ""
                WPartida1 = ""
                WPartida2 = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Where Informe.Clave = " + "'" + WClave + "'"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                
                    ZZCanti = rstInforme!Cantidad
                    ZZCantienv = rstInforme!CantidadEnv
                    ZZLiberada = ZZCanti - ZZCantienv
                    ZZDevuelta = ZZCantienv
                    ZZDesvio = 0
                    
                    If rstInforme!EstadoEnvII = 1 Or rstInforme!EstadoEnvIV = 1 Or rstInforme!EstadoEnvVI = 1 Or rstInforme!EstadoEnvVIII = 1 Or rstInforme!EstadoEnvX = 1 Then
                        ZZDesvio = ZZLiberada
                        ZZLiberada = 0
                    End If
                
                    rstInforme.Close
                End If
                
                If ZZDevuelta > 0 Then
                    ZProveedor(WCiclo, 5) = Str$(Val(ZProveedor(WCiclo, 5)) + 1)
                End If
                
                If ZZDesvio > 0 Then
                    ZProveedor(WCiclo, 4) = Str$(Val(ZProveedor(WCiclo, 4)) + 1)
                End If
                
                If ZZLiberada > 0 Then
                    ZProveedor(WCiclo, 3) = Str$(Val(ZProveedor(WCiclo, 3)) + 1)
                End If
                
            Next Ciclo
        
        Next WCiclo
        
    Next ZCiclo
    
    Call Conecta_Empresa
    
    LugarVector = 0
    Call Limpia_Vector
    
    For WCiclo = 1 To ZLugar
        
        ZProve = ZProveedor(WCiclo, 1)
        
        Rem total de movimientos
        ZImpre1 = ZProveedor(WCiclo, 2)
        
        Rem item aprobados
        ZImpre2 = ZProveedor(WCiclo, 3)
        
        Rem item desvios
        ZImpre3 = ZProveedor(WCiclo, 4)
        
        Rem item rechazados
        ZImpre4 = ZProveedor(WCiclo, 5)
        
        Rem cantidad de certificado ok
        ZImpre5 = ZProveedor(WCiclo, 6)
        
        Rem cantidad de estados de envases ok
        ZImpre6 = ZProveedor(WCiclo, 7)
        
        Rem cantidad de estados de envases ok
        ZRetrazo = ZProveedor(WCiclo, 8)
        
        If Val(ZImpre1) <> 0 Then
            ZImpre7 = Str$((Val(ZImpre5) / Val(ZImpre1)) * 100)
                Else
            ZImpre7 = ""
        End If
        
        If Val(ZImpre1) <> 0 Then
            ZImpre8 = Str$((Val(ZImpre6) / Val(ZImpre1)) * 100)
                Else
            ZImpre8 = ""
        End If
        
        If Val(ZImpre1) <> 0 Then
            ZImpre9 = Str$(((Val(ZImpre5) + Val(ZImpre6)) / (Val(ZImpre1) * 2)) * 100)
                Else
            ZImpre9 = ""
        End If
            
        ZImpre10 = ""
        If Val(ZImpre1) <> 0 Then
            ZImpre10 = Str$(Val(ZRetrazo) / Val(ZImpre1))
        End If
        ZImpre10 = Str$(Int(Val(ZImpre10)))
                         
        If Val(ZImpre10) <= 1 Then
            ZImpre11 = "Muy Bueno"
                Else
            If Val(ZImpre10) <= 2 Then
                ZImpre11 = "Bueno"
                    Else
                If Val(ZImpre10) <= 7 Then
                    ZImpre11 = "Regular"
                        Else
                    ZImpre11 = "Malo"
                End If
            End If
        End If
        
        If Val(ZImpre4) = 0 Then
            ZImpre12 = "A"
                Else
            If Val(ZImpre4) = 1 Then
                ZImpre12 = "B"
                    Else
                ZImpre12 = "C"
            End If
        End If
        
        LugarVector = LugarVector + 1
        
        ZImpre11 = ""
        ZImpre12 = ""
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZProve + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WRazon = rstProveedor!Nombre
            
            ZCalificaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
            ZCalificaII = IIf(IsNull(rstProveedor!CategoriaII), "0", rstProveedor!CategoriaII)
            
            ZFechaCategoria = IIf(IsNull(rstProveedor!fechaCategoria), "00/00/0000", rstProveedor!fechaCategoria)
            ZFechaCategoria = Mid$(ZFechaCategoria, 1, 6) + Right$(ZFechaCategoria, 2)
            
            Select Case ZCalificaI
                Case 1
                    ZImpre12 = "A"
                Case 2
                    ZImpre12 = "B"
                Case 3
                    ZImpre12 = "C"
                Case 4
                    ZImpre12 = "E"
                Case Else
                    ZImpre12 = ""
            End Select
            
            Select Case ZCalificaII
                Case 1
                    ZImpre11 = "Muy Bueno"
                Case 2
                    ZImpre11 = "Bueno"
                Case 3
                    ZImpre11 = "Regular"
                Case 4
                    ZImpre11 = "Malo"
                Case Else
                    ZImpre11 = "Sin Calificar"
            End Select
            
            rstProveedor.Close
                Else
            WRazon = ""
        End If
                
                
        WVector1.TextMatrix(LugarVector, 1) = WRazon
        WVector1.TextMatrix(LugarVector, 2) = ZImpre1
        WVector1.TextMatrix(LugarVector, 3) = ZImpre2
        WVector1.TextMatrix(LugarVector, 4) = ZImpre3
        WVector1.TextMatrix(LugarVector, 5) = ZImpre4
        Rem WVector1.TextMatrix(LugarVector, 6) = ZImpre5
        Rem WVector1.TextMatrix(LugarVector, 7) = ZImpre6
        Rem WVector1.TextMatrix(LugarVector, 8) = Pusing("###.##", ZImpre7)
        Rem WVector1.TextMatrix(LugarVector, 9) = Pusing("###.##", ZImpre8)
        Rem WVector1.TextMatrix(LugarVector, 10) = Pusing("###.##", ZImpre9)
        WVector1.TextMatrix(LugarVector, 6) = ZImpre10
        WVector1.TextMatrix(LugarVector, 7) = ZImpre12
        WVector1.TextMatrix(LugarVector, 8) = ZImpre11
        WVector1.TextMatrix(LugarVector, 9) = ""
        WVector1.TextMatrix(LugarVector, 10) = ZFechaCategoria
        WVector1.TextMatrix(LugarVector, 11) = ZProve
    
    Next WCiclo
    
    WVector1.Row = 1
    WVector1.Col = 1
    WVector1.TopRow = 1
    
    If TipoActualiza.ListIndex = 1 Then
        WVector1.Row = 1
        WVector1.Col = 7
        WVector1.TopRow = 1
        Call StartEdit
            Else
        WVector1.Row = 1
        WVector1.Col = 8
        WVector1.TopRow = 1
        Call StartEdit
    End If
    
    
End Sub

Private Sub BotonCerrar_Click()
    PantaAnaliza.Visible = False
End Sub

Private Sub ConfirmaPantaEnvase_Click()
    PantaEnvase.Visible = False
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "

    TipoActualiza.Clear
    
    TipoActualiza.AddItem ""
    TipoActualiza.AddItem "Calidad"
    TipoActualiza.AddItem "Entrega"
    
    TipoActualiza.ListIndex = 0
    
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Procesa
    End If
End Sub

Private Sub Cancela_click()
    PrgCalificaGrillaEnvases.Hide
    Unload Me
    Menu.Show
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

Private Sub TipoActualiza_Click()
    Call Procesa
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

Private Sub WVector1_DblClick()
    XProveedor = WVector1.TextMatrix(WVector1.Row, 11)
    Call ProcesaProveedor
End Sub

Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            If WParametros(2, WVector1.Col) = 0 Then
                Rem Carga los datos en el caso que el campo a editar sea un combo
                WCombo1.Clear
                WCombo1.AddItem "Sin Calificar"
                WCombo1.AddItem "Muy Bueno"
                WCombo1.AddItem "Bueno"
                WCombo1.AddItem "Regular"
                WCombo1.AddItem "Malo"
                On Error Resume Next
                WCombo1.Text = WVector1.Text
                On Error GoTo 0
                GridEditCombo
            End If
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 9
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 7
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
    WVector1.Cols = 12
    WVector1.FixedRows = 1
    WVector1.Rows = 5001
    
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
    
    WVector1.ColWidth(0) = 10
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Proveedor"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Items"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Aprob"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Desv."
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Rech."
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Atraso"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Calidad"
                WVector1.ColWidth(Ciclo) = "800"
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                If TipoActualiza.ListIndex = 1 Then
                    WParametros(2, Ciclo) = 0
                        Else
                    WParametros(2, Ciclo) = 1
                End If
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Entrega"
                WVector1.ColWidth(Ciclo) = "1200"
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                If TipoActualiza.ListIndex = 2 Then
                    WParametros(2, Ciclo) = 0
                        Else
                    WParametros(2, Ciclo) = 1
                End If
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 1
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Ok"
                WVector1.ColWidth(Ciclo) = "400"
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = "1000"
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = "50"
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 1
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

Private Sub Actualiza_Click()

    For Ciclo = 1 To 5000
    
        If WVector1.TextMatrix(Ciclo, 9) = "S" Then
        
            ZProve = ZProveedor(Ciclo, 1)
            
            Select Case TipoActualiza.ListIndex
                Case 1
                    WCalificaI = WVector1.TextMatrix(Ciclo, 7)
                    WCalificaII = WVector1.TextMatrix(Ciclo, 8)
                    
                    Select Case WCalificaI
                        Case "A"
                            ZCalificaI = "1"
                        Case "B"
                            ZCalificaI = "2"
                        Case "C"
                            ZCalificaI = "3"
                        Case "E"
                            ZCalificaI = "4"
                        Case Else
                            ZCalificaI = "0"
                    End Select
                    
                    Select Case WCalificaII
                        Case "Muy Bueno"
                            ZCalificaII = "1"
                        Case "Bueno"
                            ZCalificaII = "2"
                        Case "Regular"
                            ZCalificaII = "3"
                        Case "Malo"
                            ZCalificaII = "4"
                        Case Else
                            ZCalificaII = "0"
                    End Select
                    ZFechaCategoria = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZOrdFechaCategoria = Right$(ZFechaCalifica, 4) + Mid$(ZFechaCalifica, 4, 2) + Left$(ZFechaCalifica, 2)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Proveedor SET "
                    ZSql = ZSql + " FechaCategoria = " + "'" + ZFechaCategoria + "',"
                    ZSql = ZSql + " OrdFechaCategoria = " + "'" + ZOrdFechaCategoria + "',"
                    ZSql = ZSql + " CategoriaI = " + "'" + ZCalificaI + "'"
                    ZSql = ZSql + " Where Proveedor = " + "'" + ZProve + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
                Case 2
                    WCalificaI = WVector1.TextMatrix(Ciclo, 7)
                    WCalificaII = WVector1.TextMatrix(Ciclo, 8)
                    
                    Select Case WCalificaI
                        Case "A"
                            ZCalificaI = "1"
                        Case "B"
                            ZCalificaI = "2"
                        Case "C"
                            ZCalificaI = "3"
                        Case "E"
                            ZCalificaI = "4"
                        Case Else
                            ZCalificaI = "0"
                    End Select
                    
                    Select Case WCalificaII
                        Case "Muy Bueno"
                            ZCalificaII = "1"
                        Case "Bueno"
                            ZCalificaII = "2"
                        Case "Regular"
                            ZCalificaII = "3"
                        Case "Malo"
                            ZCalificaII = "4"
                        Case Else
                            ZCalificaII = "0"
                    End Select
                    ZFechaCategoria = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZOrdFechaCategoria = Right$(ZFechaCalifica, 4) + Mid$(ZFechaCalifica, 4, 2) + Left$(ZFechaCalifica, 2)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Proveedor SET "
                    ZSql = ZSql + " FechaCategoria = " + "'" + ZFechaCategoria + "',"
                    ZSql = ZSql + " OrdFechaCategoria = " + "'" + ZOrdFechaCategoria + "',"
                    ZSql = ZSql + " CategoriaII = " + "'" + ZCalificaII + "'"
                    ZSql = ZSql + " Where Proveedor = " + "'" + ZProve + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    
            End Select
            
            
        End If
        
    Next Ciclo
    
    Call Cancela_click

End Sub

Private Sub ProcesaProveedor()

    Call Limpia_Vector2
    
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

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    XEmpresa = WEmpresa
    Xlugar = 0
    
    For ZCiclo = 1 To 9
    
        WEmpresa = CargaEmpresa(ZCiclo, 1)
        txtOdbc = CargaEmpresa(ZCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
          
        Erase WVector
        Lugar = 0
        
        XParam = "'" + WDesde + "','" _
                     + WHasta + "','" _
                     + XProveedor + "','" _
                     + XProveedor + "'"

        spInforme = "ListaInformeListado" + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            
            With rstInforme
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        Lugar = Lugar + 1
                
                        WCertificado = IIf(IsNull(rstInforme!Certificado1), "1", rstInforme!Certificado1)
                        WEstado = IIf(IsNull(rstInforme!Estado1), "1", rstInforme!Estado1)
                            
                        WVector(Lugar, 1) = rstInforme!Articulo
                        WVector(Lugar, 2) = rstInforme!Orden
                        WVector(Lugar, 3) = rstInforme!FechaOrd
                        WVector(Lugar, 4) = rstInforme!Clave
                        WVector(Lugar, 5) = WCertificado
                        WVector(Lugar, 6) = WEstado
                        WVector(Lugar, 7) = rstInforme!Informe
                        WVector(Lugar, 8) = rstInforme!Cantidad
                        WVector(Lugar, 9) = rstInforme!Fecha
                        WVector(Lugar, 10) = ZCiclo

                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
            
                    Loop
                End If
            End With
        
            rstInforme.Close
        
        End If
    
        For Ciclo = 1 To Lugar
            
            Rem Calcula las diferencias de fecha entre la
            Rem Orden de compra y el informe de recepcion
    
            WArticulo = WVector(Ciclo, 1)
            WOrden = WVector(Ciclo, 2)
            WFecha = WVector(Ciclo, 3)
            WClave = WVector(Ciclo, 4)
            WCertificado = WVector(Ciclo, 5)
            WEstado = WVector(Ciclo, 6)
            WInforme = WVector(Ciclo, 7)
            WCantidad = WVector(Ciclo, 8)
            Rem WFecha = WVector(Ciclo, 9)
            WCiclo = WVector(Ciclo, 10)
            
            Campo2 = ""
            Campo3 = ""
            Campo4 = ""
            Campo5 = ""
            Campo6 = ""
            Campo7 = ""
            Campo8 = 0
                
            XFecha = "  /  /    "
            XOrdFecha = "00000000"
        
            XParam = "'" + WOrden + "','" _
                         + WArticulo + "'"

            spOrden = "ListaOrdenArticulo" + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
        
                XOrdFecha = Right$(rstOrden!Fecha2, 4) + Mid$(rstOrden!Fecha2, 4, 2) + Left$(rstOrden!Fecha2, 2)
                XFecha = rstOrden!Fecha2
            
                rstOrden.Close
            
            End If
        
            BAse1 = (Val(Left$(XOrdFecha, 4)) * 365) + (Val(Mid$(XOrdFecha, 5, 2)) * 30) + (Val(Right$(XOrdFecha, 2)) * 1)
            Base2 = (Val(Left$(WFecha, 4)) * 365) + (Val(Mid$(WFecha, 5, 2)) * 30) + (Val(Right$(WFecha, 2)) * 1)
        
            Dife = Base2 - BAse1
            If Dife < 0 Then
                Dife = 0
            End If
                
            Campo2 = "X"
                
            If Val(WCertificado) = 1 Then
                Campo3 = "X"
            End If
                
            If Val(WEstado) = 1 Then
                Campo7 = "X"
            End If
                
            aa = Dife
            If Dife > 100 Then Dife = 0
            Campo8 = Dife
                
            Rem Calcula las diferencias de fecha entre la
            Rem Orden de compra y el informe de recepcion
                
            WLiberada = ""
            WDevuelta = ""
            WPartida1 = ""
            WPartida2 = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Clave = " + "'" + WClave + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
            
                ZZCanti = rstInforme!Cantidad
                ZZCantienv = rstInforme!CantidadEnv
                ZZLiberada = ZZCanti - ZZCantienv
                ZZDevuelta = ZZCantienv
                ZZDesvio = 0
                
                If rstInforme!EstadoEnvII = 1 Or rstInforme!EstadoEnvIV = 1 Or rstInforme!EstadoEnvVI = 1 Or rstInforme!EstadoEnvVIII = 1 Or rstInforme!EstadoEnvX = 1 Then
                    ZZDesvio = ZZLiberada
                    ZZLiberada = 0
                End If
            
                rstInforme.Close
            End If
            
            If ZZDevuelta > 0 Then
                Campo5 = "X"
            End If
            
            If ZZDesvio > 0 Then
                Campo4 = "X"
            End If
            
            If ZZLiberada > 0 Then
                Campo3 = "X"
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDesArticulo = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                WDesArticulo = ""
            End If
            
            WFecha = WVector(Ciclo, 9)
            
            Xlugar = Xlugar + 1
            
            WVector2.TextMatrix(Xlugar, 1) = WInforme
            WVector2.TextMatrix(Xlugar, 2) = WOrden
            WVector2.TextMatrix(Xlugar, 3) = WFecha
            WVector2.TextMatrix(Xlugar, 4) = WArticulo
            WVector2.TextMatrix(Xlugar, 5) = WDesArticulo
            WVector2.TextMatrix(Xlugar, 6) = Campo3
            WVector2.TextMatrix(Xlugar, 7) = Campo4
            WVector2.TextMatrix(Xlugar, 8) = Campo5
            WVector2.TextMatrix(Xlugar, 9) = Campo8
            If WPartida1 <> "" Then
                WVector2.TextMatrix(Xlugar, 10) = WPartida1
                    Else
                WVector2.TextMatrix(Xlugar, 10) = WPartida2
            End If
            WVector2.TextMatrix(Xlugar, 11) = WCiclo
                
        Next Ciclo
            
    Next ZCiclo
    
    WVector2.TopRow = 1
    WVector2.Row = 1
    WVector2.Col = 1
    
    Call Conecta_Empresa
    
    PantaAnaliza.Visible = True
    
End Sub

Private Sub Limpia_Vector2()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
    WVector2.FixedRows = 1
    WVector2.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros2(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros2(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros2(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros2(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Informe"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector2.Text = "Articulo"
                WVector2.ColWidth(Ciclo) = 1300
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 2000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                WVector2.Text = "Aprob."
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                WVector2.Text = "Desvio"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Rechaz"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                WVector2.Text = "Atraso"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 10
                WVector2.Text = ""
                WVector2.ColWidth(Ciclo) = 50
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                WVector2.Text = ""
                WVector2.ColWidth(Ciclo) = 50
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo2(Ciclo).Text = WVector2.Text
        Rem WTitulo2(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo2(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo2(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo2(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo2(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
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
    
    WGrilla.TopRow = 1
    WGrilla.Col = 1
    WGrilla.Row = 1
    
End Sub

Private Sub WVector2_dblClick()

    XInforme = WVector2.TextMatrix(WVector2.Row, 1)
    XEmpresa = WVector2.TextMatrix(WVector2.Row, 11)
    
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    
    XXEmpresa = WEmpresa
    Call Conecta_Empresa

    Call Limpia_Grilla
    Renglon = 0
    Erase DatosCertificado
    Erase Auxiliar
    
    spInforme = "ListaInforme " + "'" + XInforme + "'"
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
                    ZVencimiento = IIf(IsNull(rstInforme!FechaVencimiento), "  /  /    ", rstInforme!FechaVencimiento)
                    ZProcedencia = IIf(IsNull(rstInforme!Procedencia), "", rstInforme!Procedencia)
                    ZObserva = IIf(IsNull(rstInforme!observa), "", rstInforme!observa)
                            
                    DatosCertificado(Renglon, 1) = Str$(ZCertificadoSi)
                    DatosCertificado(Renglon, 2) = Str$(ZCertificadoNo)
                    DatosCertificado(Renglon, 3) = ZCertificado2
            
                    DatosCertificado(Renglon, 4) = Str$(ZEstadoSi)
                    DatosCertificado(Renglon, 5) = Str$(ZEstadoNo)
                    DatosCertificado(Renglon, 6) = ZEstado2
                    
                    DatosCertificado(Renglon, 7) = ZVencimiento
                    DatosCertificado(Renglon, 8) = ZProcedencia
                    DatosCertificado(Renglon, 9) = ZObserva
                    
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
                    
                    
                    Auxiliar(Renglon, 3) = rstInforme!Articulo
                    Auxiliar(Renglon, 5) = rstInforme!Envase
                    Auxiliar(Renglon, 7) = rstInforme!Clave
                    
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
    
    For da = 1 To WRenglon
    
        Renglon = Renglon + 1
        WGrilla.Row = Renglon
                
        spArticulo = "ConsultaArticulo " + "'" + Auxiliar(Renglon, 3) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WGrilla.Col = 3
            WGrilla.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        spEnvases = "ConsultaEnvases " + "'" + Auxiliar(Renglon, 5) + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            WGrilla.Col = 8
            WGrilla.Text = rstEnvases!Descripcion
            rstEnvases.Close
                Else
            WGrilla.Col = 8
            WGrilla.Text = ""
        End If
        
        
    Next da
    
    XEmpresa = XXEmpresa
    Call Conecta_Empresa
    
    PantaInforme.Visible = True
    
End Sub

Private Sub CancelaInforme_Click()
    PantaInforme.Visible = False
End Sub

Private Sub WGrilla_DblClick()

    ZLugar = WGrilla.Row
    ZArticulo = WGrilla.TextMatrix(ZLugar, 2)

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
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + ZArticulo + "'"
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

End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub


