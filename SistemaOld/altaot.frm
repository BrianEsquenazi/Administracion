VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaOt 
   Caption         =   "Solicitud de Muestras para Clientes"
   ClientHeight    =   8415
   ClientLeft      =   1050
   ClientTop       =   390
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   9960
   Begin VB.CommandButton Config 
      Caption         =   "Configuracion"
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
      Left            =   8400
      TabIndex        =   105
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame AltaMaquina 
      Caption         =   "Maquina"
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
      Left            =   8880
      TabIndex        =   46
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
      Begin VB.CheckBox Maquina1 
         Caption         =   "Baño Maria"
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
         TabIndex        =   62
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Maquina2 
         Caption         =   "HTx 12"
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
         TabIndex        =   61
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox Maquina3 
         Caption         =   "Rotadyer Mathis"
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
         TabIndex        =   60
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Maquina4 
         Caption         =   "Madejera"
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
         TabIndex        =   59
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox Maquina5 
         Caption         =   "Recubridor"
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
         TabIndex        =   58
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox Maquina6 
         Caption         =   "Jigger"
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
         TabIndex        =   57
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox Maquina7 
         Caption         =   "Vaporizador"
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
         TabIndex        =   56
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Maquina8 
         Caption         =   "Foulard"
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
         Left            =   3240
         TabIndex        =   55
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Maquina9 
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
         Left            =   3240
         TabIndex        =   54
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox Maquina10 
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
         Left            =   3240
         TabIndex        =   53
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox Maquina11 
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
         Left            =   3240
         TabIndex        =   52
         Top             =   1560
         Width           =   3135
      End
      Begin VB.CheckBox Maquina12 
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
         Left            =   3240
         TabIndex        =   51
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CheckBox Maquina13 
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
         Left            =   3240
         TabIndex        =   50
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Maqui 
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
         TabIndex        =   49
         Top             =   3000
         Width           =   5295
      End
      Begin VB.CheckBox Maquina14 
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
         Left            =   3240
         TabIndex        =   48
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton FinMaquina 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   47
         Top             =   3480
         Width           =   1455
      End
   End
   Begin VB.Frame AltaColor 
      Caption         =   "Ingreso de Colorantes a Utilizar"
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
      Left            =   8760
      TabIndex        =   63
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
      Begin VB.CheckBox Color21 
         Caption         =   "Procion XL+"
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
         Left            =   5880
         TabIndex        =   104
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Color20 
         Caption         =   "Procion XL+"
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
         Left            =   5880
         TabIndex        =   103
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Color19 
         Caption         =   "Procion H-E"
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
         Left            =   5880
         TabIndex        =   102
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox Color18 
         Caption         =   "B.Optico"
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
         Left            =   5880
         TabIndex        =   101
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Color17 
         Caption         =   "Astrazon"
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
         Left            =   5880
         TabIndex        =   100
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox Color16 
         Caption         =   "Indanthrene"
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
         Left            =   5880
         TabIndex        =   99
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox Color15 
         Caption         =   "Isolan"
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
         Left            =   5880
         TabIndex        =   98
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Color1 
         Caption         =   "Sirius"
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
         TabIndex        =   79
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox Color2 
         Caption         =   "Remazol"
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
         TabIndex        =   78
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox Color3 
         Caption         =   "Levafix"
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
         TabIndex        =   77
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Color4 
         Caption         =   "Dianix"
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
         TabIndex        =   76
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox Color5 
         Caption         =   "Dispersol"
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
         TabIndex        =   75
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox Color6 
         Caption         =   "Telon"
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
         TabIndex        =   74
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox Color7 
         Caption         =   "Nylomine"
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
         TabIndex        =   73
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Color8 
         Caption         =   "Isolan"
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
         Left            =   3240
         TabIndex        =   72
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox Color9 
         Caption         =   "Indanthrene"
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
         Left            =   3240
         TabIndex        =   71
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox Color10 
         Caption         =   "Astrazon"
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
         Left            =   3240
         TabIndex        =   70
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox Color11 
         Caption         =   "B.Optico"
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
         Left            =   3240
         TabIndex        =   69
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox Color12 
         Caption         =   "Procion H-E"
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
         Left            =   3240
         TabIndex        =   68
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox Color13 
         Caption         =   "Procion XL+"
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
         Left            =   3240
         TabIndex        =   67
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Color 
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
         TabIndex        =   66
         Top             =   3120
         Width           =   5295
      End
      Begin VB.CheckBox Color14 
         Caption         =   "Procion XL+"
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
         Left            =   3240
         TabIndex        =   65
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton FinColor 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   64
         Top             =   3600
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3495
      Left            =   120
      TabIndex        =   97
      Top             =   4680
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin Crystal.CrystalReport Listagrilla 
      Left            =   9120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreot.rpt"
   End
   Begin VB.Frame AltaTrabajo 
      Caption         =   "Trabajos a Realizar"
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
      Left            =   9120
      TabIndex        =   80
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
      Begin VB.CheckBox Trabajo1 
         Caption         =   "Tintura"
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
         TabIndex        =   96
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox Trabajo2 
         Caption         =   "Imitacion Tono"
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
         TabIndex        =   95
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox Trabajo3 
         Caption         =   "Descrude"
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
         TabIndex        =   94
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Trabajo4 
         Caption         =   "Descrude/Blanqueo"
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
         TabIndex        =   93
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox Trabajo5 
         Caption         =   "Recubrimiento"
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
         TabIndex        =   92
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox Trabajo6 
         Caption         =   "Reconocimiento Fibra"
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
         TabIndex        =   91
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CheckBox Trabajo7 
         Caption         =   "Ensayo y Solidedes"
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
         TabIndex        =   90
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Trabajo8 
         Caption         =   "Comparacion Competencia"
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
         Left            =   3240
         TabIndex        =   89
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox Trabajo9 
         Caption         =   "Terminacion"
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
         Left            =   3240
         TabIndex        =   88
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox Trabajo10 
         Caption         =   "Ilustracion Colorante"
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
         Left            =   3240
         TabIndex        =   87
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox Trabajo11 
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
         Left            =   3240
         TabIndex        =   86
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox Trabajo12 
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
         Left            =   3240
         TabIndex        =   85
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox Trabajo13 
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
         Left            =   3240
         TabIndex        =   84
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Traba 
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
         TabIndex        =   83
         Top             =   3000
         Width           =   5295
      End
      Begin VB.CheckBox Trabajo14 
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
         Left            =   3240
         TabIndex        =   82
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton FinTrabajo 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   81
         Top             =   3480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Trabajo 
      Caption         =   " Trabajos a Realizar  (F5)"
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
      Left            =   8400
      TabIndex        =   45
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Colorante 
      Caption         =   "Colorantes a Usar  (F6)"
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
      Left            =   8400
      TabIndex        =   44
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Maquina 
      Caption         =   "     Maquina              (F7)"
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
      Left            =   8400
      TabIndex        =   43
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Composicion 
      Caption         =   "  Composicion           (F4)"
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
      Left            =   8400
      TabIndex        =   42
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame AltaCompo 
      Caption         =   "Ingreso de Composicion"
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
      Left            =   8640
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton FinCompo 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox Compo14 
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
         Left            =   3240
         TabIndex        =   40
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox Compo 
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
         TabIndex        =   35
         Top             =   3000
         Width           =   5295
      End
      Begin VB.CheckBox Compo13 
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
         Left            =   3240
         TabIndex        =   34
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox Compo12 
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
         Left            =   3240
         TabIndex        =   33
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CheckBox Compo11 
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
         Left            =   3240
         TabIndex        =   32
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CheckBox Compo10 
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
         Left            =   3240
         TabIndex        =   31
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox Compo9 
         Caption         =   "Mezclas"
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
         Left            =   3240
         TabIndex        =   30
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox Compo8 
         Caption         =   "Viscosa Rayon"
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
         Left            =   3240
         TabIndex        =   29
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox Compo7 
         Caption         =   "Lana"
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
         TabIndex        =   28
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Compo6 
         Caption         =   "Acetato"
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
         TabIndex        =   27
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Compo5 
         Caption         =   "Acrilico"
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
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox Compo4 
         Caption         =   "Poliester"
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
         TabIndex        =   25
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Compo3 
         Caption         =   "Poliamida"
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
         TabIndex        =   24
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Compo2 
         Caption         =   "Algodon mercerizado"
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
         TabIndex        =   23
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox Compo1 
         Caption         =   "Algodon"
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
         TabIndex        =   22
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.TextBox Solicitante 
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
      MaxLength       =   15
      TabIndex        =   38
      Text            =   " "
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Observaciones3 
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
      MaxLength       =   50
      TabIndex        =   37
      Text            =   " "
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox Observaciones2 
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
      MaxLength       =   50
      TabIndex        =   36
      Text            =   " "
      Top             =   1560
      Width           =   6015
   End
   Begin VB.TextBox Observaciones1 
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
      MaxLength       =   50
      TabIndex        =   15
      Text            =   " "
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox Solidez 
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
      MaxLength       =   50
      TabIndex        =   14
      Text            =   " "
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox Preparacion 
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
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox Razon 
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
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   11
      Top             =   480
      Width           =   4575
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
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   120
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "   Consulta        Datos           (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "    Limpia         Pantalla          (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "    Fin de         Ingreso          (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6000
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "    Graba            (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox FechaCompro 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   3000
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
   Begin MSMask.MaskEdBox FechaSalida 
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   3000
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
   Begin VB.Label Label1 
      Caption         =   "Solicitado"
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
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha de Salida"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha Comprometida"
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
      TabIndex        =   18
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha de Solicitud"
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
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgAltaOt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstOt As Recordset
Dim spOt As String
Dim rstConfigOt As Recordset
Dim spConfigOt As String
Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer


Private Sub cmdGraba_Click()

    If Val(WOt) <> 0 Then
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdFechaCompro = Right$(FechaCompro.Text, 4) + Mid$(FechaCompro.Text, 4, 2) + Left$(FechaCompro.Text, 2)
        WOrdFechaSalida = Right$(FechaSalida.Text, 4) + Mid$(FechaSalida.Text, 4, 2) + Left$(FechaSalida.Text, 2)
        
        WCompo1 = Str$(Compo1.Value)
        WCompo2 = Str$(Compo2.Value)
        WCompo3 = Str$(Compo3.Value)
        WCompo4 = Str$(Compo4.Value)
        WCompo5 = Str$(Compo5.Value)
        WCompo6 = Str$(Compo6.Value)
        WCompo7 = Str$(Compo7.Value)
        WCompo8 = Str$(Compo8.Value)
        WCompo9 = Str$(Compo9.Value)
        WCompo10 = Str$(Compo10.Value)
        WCompo11 = Str$(Compo11.Value)
        WCompo12 = Str$(Compo12.Value)
        WCompo13 = Str$(Compo13.Value)
        WCompo14 = Str$(Compo14.Value)
        
        WTrabajo1 = Str$(Trabajo1.Value)
        WTrabajo2 = Str$(Trabajo2.Value)
        WTrabajo3 = Str$(Trabajo3.Value)
        WTrabajo4 = Str$(Trabajo4.Value)
        WTrabajo5 = Str$(Trabajo5.Value)
        WTrabajo6 = Str$(Trabajo6.Value)
        WTrabajo7 = Str$(Trabajo7.Value)
        WTrabajo8 = Str$(Trabajo8.Value)
        WTrabajo9 = Str$(Trabajo9.Value)
        WTrabajo10 = Str$(Trabajo10.Value)
        WTrabajo11 = Str$(Trabajo11.Value)
        WTrabajo12 = Str$(Trabajo12.Value)
        WTrabajo13 = Str$(Trabajo13.Value)
        WTrabajo14 = Str$(Trabajo14.Value)
        
        WColor1 = Str$(Color1.Value)
        WColor2 = Str$(Color2.Value)
        WColor3 = Str$(Color3.Value)
        WColor4 = Str$(Color4.Value)
        WColor5 = Str$(Color5.Value)
        WColor6 = Str$(Color6.Value)
        WColor7 = Str$(Color7.Value)
        WColor8 = Str$(Color8.Value)
        WColor9 = Str$(Color9.Value)
        WColor10 = Str$(Color10.Value)
        WColor11 = Str$(Color11.Value)
        WColor12 = Str$(Color12.Value)
        WColor13 = Str$(Color13.Value)
        WColor14 = Str$(Color14.Value)
        WColor15 = Str$(Color15.Value)
        WColor16 = Str$(Color16.Value)
        WColor17 = Str$(Color17.Value)
        WColor18 = Str$(Color18.Value)
        WColor19 = Str$(Color19.Value)
        WColor20 = Str$(Color20.Value)
        WColor21 = Str$(Color21.Value)
        
        WMaquina1 = Str$(Maquina1.Value)
        WMaquina2 = Str$(Maquina2.Value)
        WMaquina3 = Str$(Maquina3.Value)
        WMaquina4 = Str$(Maquina4.Value)
        WMaquina5 = Str$(Maquina5.Value)
        WMaquina6 = Str$(Maquina6.Value)
        WMaquina7 = Str$(Maquina7.Value)
        WMaquina8 = Str$(Maquina8.Value)
        WMaquina9 = Str$(Maquina9.Value)
        WMaquina10 = Str$(Maquina10.Value)
        WMaquina11 = Str$(Maquina11.Value)
        WMaquina12 = Str$(Maquina12.Value)
        WMaquina13 = Str$(Maquina13.Value)
        WMaquina14 = Str$(Maquina14.Value)
        
        WClave = "1"
        
        XParam = "'" + WOt + "','" _
                 + Fecha.Text + "','" _
                 + Cliente.Text + "','" _
                 + Razon.Text + "','" _
                 + Preparacion.Text + "','" _
                 + Solidez.Text + "','" _
                 + Observaciones1.Text + "','" _
                 + Observaciones2.Text + "','" _
                 + Observaciones3.Text + "','" _
                 + Solicitante.Text + "','" _
                 + Compo.Text + "','" + WCompo1 + "','" + WCompo2 + "','" + WCompo3 + "','" + WCompo4 + "','" + WCompo5 + "','" + WCompo6 + "','" + WCompo7 + "','" + WCompo8 + "','" _
                 + WCompo9 + "','" + WCompo10 + "','" + WCompo11 + "','" + WCompo12 + "','" + WCompo13 + "','" + WCompo14 + "','" _
                 + Traba.Text + "','" + WTrabajo1 + "','" + WTrabajo2 + "','" + WTrabajo3 + "','" + WTrabajo4 + "','" + WTrabajo5 + "','" + WTrabajo6 + "','" + WTrabajo7 + "','" + WTrabajo8 + "','" _
                 + WTrabajo9 + "','" + WTrabajo10 + "','" + WTrabajo11 + "','" + WTrabajo12 + "','" + WTrabajo13 + "','" + WTrabajo14 + "','" _
                 + Color.Text + "','" + WColor1 + "','" + WColor2 + "','" + WColor3 + "','" + WColor4 + "','" + WColor5 + "','" + WColor6 + "','" + WColor7 + "','" + WColor8 + "','" _
                 + WColor9 + "','" + WColor10 + "','" + WColor11 + "','" + WColor12 + "','" + WColor13 + "','" + WColor14 + "','" _
                 + WColor15 + "','" + WColor16 + "','" + WColor17 + "','" + WColor18 + "','" + WColor19 + "','" + WColor20 + "','" + WColor21 + "','" _
                 + Maqui.Text + "','" + WMaquina1 + "','" + WMaquina2 + "','" + WMaquina3 + "','" + WMaquina4 + "','" + WMaquina5 + "','" + WMaquina6 + "','" + WMaquina7 + "','" + WMaquina8 + "','" _
                 + WMaquina9 + "','" + WMaquina10 + "','" + WMaquina11 + "','" + WMaquina12 + "','" + WMaquina13 + "','" + WMaquina14 + "','" _
                 + FechaCompro.Text + "','" + FechaSalida.Text + "','" _
                 + WOrdFecha + "','" + WOrdFechaCompro + "','" + WOrdFechaSalida + "','" + WClave + "'"
                 
        Set rstOt = db.OpenRecordset("ModificaOt " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        T$ = "Impresion de Orden de Trabajo"
        m$ = "Desea Imprimir la Orden de Trabajo"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Listagrilla.GroupSelectionFormula = "{Ot.Codigo} in " + WOt + " to " + WOt
            Listagrilla.Destination = 1
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            Listagrilla.SQLQuery = "SELECT  Ot.Codigo, Ot.Fecha, Ot.Razon, Ot.Preparacion, Ot.Solidez, Ot.Observaciones1, Ot.Observaciones2, " _
                        + "Ot.Observaciones3, Ot.Solicitante, Ot.FechaCompro, " _
                        + "Ot.Compo, Ot.Compo1, Ot.Compo2, Ot.Compo3, Ot.Compo4, Ot.Compo5, Ot.Compo6, Ot.Compo7, Ot.Compo8, Ot.Compo9, Ot.Compo10, Ot.Compo11, Ot.Compo12, Ot.Compo13, Ot.Compo14, " _
                        + "Ot.Traba, Ot.Trabajo1, Ot.Trabajo2, Ot.Trabajo3, Ot.Trabajo4, Ot.Trabajo5, Ot.Trabajo6, Ot.Trabajo7, Ot.Trabajo8, Ot.Trabajo9, Ot.Trabajo10, Ot.Trabajo11, Ot.Trabajo12, Ot.Trabajo13, Ot.Trabajo14, " _
                        + "Ot.Color, Ot.Color1, Ot.Color2, Ot.Color3, Ot.Color4, Ot.Color5, Ot.Color6, Ot.Color7, Ot.Color8, Ot.Color9, Ot.Color10, Ot.Color11, Ot.Color12, Ot.Color13, Ot.Color14, Ot.Color15, Ot.Color16, Ot.Color17, Ot.Color18, Ot.Color19, Ot.Color20, Ot.Color21, " _
                        + "Ot.Maqui, Ot.Maquina1, Ot.Maquina2, Ot.Maquina3, Ot.Maquina4, Ot.Maquina5, Ot.Maquina6, Ot.Maquina7, Ot.Maquina8, Ot.Maquina9, Ot.Maquina10, Ot.Maquina11, Ot.Maquina12, Ot.Maquina13, Ot.Maquina14, " _
                        + "OtConfig.Compo1, OtConfig.Compo2, OtConfig.Compo3, OtConfig.Compo4, OtConfig.Compo5, OtConfig.Compo6, OtConfig.Compo7, OtConfig.Compo8, OtConfig.Compo9, OtConfig.Compo10, OtConfig.Compo11, OtConfig.Compo12, OtConfig.Compo13, OtConfig.Compo14, " _
                        + "OtConfig.Trabajo1, OtConfig.Trabajo2, OtConfig.Trabajo3, OtConfig.Trabajo4, OtConfig.Trabajo5, OtConfig.Trabajo6, OtConfig.Trabajo7, OtConfig.Trabajo8, OtConfig.Trabajo9, OtConfig.Trabajo10, OtConfig.Trabajo11, OtConfig.Trabajo12, OtConfig.Trabajo13, OtConfig.Trabajo14, " _
                        + "OtConfig.Color1, OtConfig.Color2, OtConfig.Color3, OtConfig.Color4, OtConfig.Color5, OtConfig.Color6, OtConfig.Color7, OtConfig.Color8, OtConfig.Color9, OtConfig.Color10, OtConfig.Color11, OtConfig.Color12, OtConfig.Color13, OtConfig.Color14, " _
                        + "OtConfig.Color15, OtConfig.Color16, OtConfig.Color17, OtConfig.Color18, OtConfig.Color19, OtConfig.Color20, OtConfig.Color21, " _
                        + "OtConfig.Maquina1, OtConfig.Maquina2, OtConfig.Maquina3, OtConfig.Maquina4, OtConfig.Maquina5, OtConfig.Maquina6, OtConfig.Maquina7, OtConfig.Maquina8, OtConfig.Maquina9, OtConfig.Maquina10, OtConfig.Maquina11, OtConfig.Maquina12, OtConfig.Maquina13, OtConfig.Maquina14 " _
                        + "From " _
                        + DSQ + ".dbo.Ot Ot, " _
                        + DSQ + ".dbo.OtConfig OtConfig " _
                        + "Where " _
                        + "Ot.Clave = OtConfig.Clave AND " _
                        + "Ot.Codigo >= " + WOt + " AND Ot.Codigo <= " + WOt
            Listagrilla.Connect = Connect()
            Listagrilla.Action = 1
        End If
        
        Call cmdClose_Click

            Else

        WCodigo = 1
        spOt = "ListaOtNumero"
        Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
        If rstOt.RecordCount > 0 Then
            With rstOt
                .MoveLast
                WCodigo = rstOt!Codigo + 1
            End With
            rstOt.Close
        End If
    
        XCodigo = Str$(WCodigo)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdFechaCompro = Right$(FechaCompro.Text, 4) + Mid$(FechaCompro.Text, 4, 2) + Left$(FechaCompro.Text, 2)
        WOrdFechaSalida = Right$(FechaSalida.Text, 4) + Mid$(FechaSalida.Text, 4, 2) + Left$(FechaSalida.Text, 2)
        
        WCompo1 = Str$(Compo1.Value)
        WCompo2 = Str$(Compo2.Value)
        WCompo3 = Str$(Compo3.Value)
        WCompo4 = Str$(Compo4.Value)
        WCompo5 = Str$(Compo5.Value)
        WCompo6 = Str$(Compo6.Value)
        WCompo7 = Str$(Compo7.Value)
        WCompo8 = Str$(Compo8.Value)
        WCompo9 = Str$(Compo9.Value)
        WCompo10 = Str$(Compo10.Value)
        WCompo11 = Str$(Compo11.Value)
        WCompo12 = Str$(Compo12.Value)
        WCompo13 = Str$(Compo13.Value)
        WCompo14 = Str$(Compo14.Value)
        
        WTrabajo1 = Str$(Trabajo1.Value)
        WTrabajo2 = Str$(Trabajo2.Value)
        WTrabajo3 = Str$(Trabajo3.Value)
        WTrabajo4 = Str$(Trabajo4.Value)
        WTrabajo5 = Str$(Trabajo5.Value)
        WTrabajo6 = Str$(Trabajo6.Value)
        WTrabajo7 = Str$(Trabajo7.Value)
        WTrabajo8 = Str$(Trabajo8.Value)
        WTrabajo9 = Str$(Trabajo9.Value)
        WTrabajo10 = Str$(Trabajo10.Value)
        WTrabajo11 = Str$(Trabajo11.Value)
        WTrabajo12 = Str$(Trabajo12.Value)
        WTrabajo13 = Str$(Trabajo13.Value)
        WTrabajo14 = Str$(Trabajo14.Value)
        
        WColor1 = Str$(Color1.Value)
        WColor2 = Str$(Color2.Value)
        WColor3 = Str$(Color3.Value)
        WColor4 = Str$(Color4.Value)
        WColor5 = Str$(Color5.Value)
        WColor6 = Str$(Color6.Value)
        WColor7 = Str$(Color7.Value)
        WColor8 = Str$(Color8.Value)
        WColor9 = Str$(Color9.Value)
        WColor10 = Str$(Color10.Value)
        WColor11 = Str$(Color11.Value)
        WColor12 = Str$(Color12.Value)
        WColor13 = Str$(Color13.Value)
        WColor14 = Str$(Color14.Value)
        WColor15 = Str$(Color15.Value)
        WColor16 = Str$(Color16.Value)
        WColor17 = Str$(Color17.Value)
        WColor18 = Str$(Color18.Value)
        WColor19 = Str$(Color19.Value)
        WColor20 = Str$(Color20.Value)
        WColor21 = Str$(Color21.Value)
        
        WMaquina1 = Str$(Maquina1.Value)
        WMaquina2 = Str$(Maquina2.Value)
        WMaquina3 = Str$(Maquina3.Value)
        WMaquina4 = Str$(Maquina4.Value)
        WMaquina5 = Str$(Maquina5.Value)
        WMaquina6 = Str$(Maquina6.Value)
        WMaquina7 = Str$(Maquina7.Value)
        WMaquina8 = Str$(Maquina8.Value)
        WMaquina9 = Str$(Maquina9.Value)
        WMaquina10 = Str$(Maquina10.Value)
        WMaquina11 = Str$(Maquina11.Value)
        WMaquina12 = Str$(Maquina12.Value)
        WMaquina13 = Str$(Maquina13.Value)
        WMaquina14 = Str$(Maquina14.Value)
        
        WClave = "1"
        
        XParam = "'" + XCodigo + "','" _
                 + Fecha.Text + "','" _
                 + Cliente.Text + "','" _
                 + Razon.Text + "','" _
                 + Preparacion.Text + "','" _
                 + Solidez.Text + "','" _
                 + Observaciones1.Text + "','" _
                 + Observaciones2.Text + "','" _
                 + Observaciones3.Text + "','" _
                 + Solicitante.Text + "','" _
                 + Compo.Text + "','" + WCompo1 + "','" + WCompo2 + "','" + WCompo3 + "','" + WCompo4 + "','" + WCompo5 + "','" + WCompo6 + "','" + WCompo7 + "','" + WCompo8 + "','" _
                 + WCompo9 + "','" + WCompo10 + "','" + WCompo11 + "','" + WCompo12 + "','" + WCompo13 + "','" + WCompo14 + "','" _
                 + Traba.Text + "','" + WTrabajo1 + "','" + WTrabajo2 + "','" + WTrabajo3 + "','" + WTrabajo4 + "','" + WTrabajo5 + "','" + WTrabajo6 + "','" + WTrabajo7 + "','" + WTrabajo8 + "','" _
                 + WTrabajo9 + "','" + WTrabajo10 + "','" + WTrabajo11 + "','" + WTrabajo12 + "','" + WTrabajo13 + "','" + WTrabajo14 + "','" _
                 + Color.Text + "','" + WColor1 + "','" + WColor2 + "','" + WColor3 + "','" + WColor4 + "','" + WColor5 + "','" + WColor6 + "','" + WColor7 + "','" + WColor8 + "','" _
                 + WColor9 + "','" + WColor10 + "','" + WColor11 + "','" + WColor12 + "','" + WColor13 + "','" + WColor14 + "','" _
                 + WColor15 + "','" + WColor16 + "','" + WColor17 + "','" + WColor18 + "','" + WColor19 + "','" + WColor20 + "','" + WColor21 + "','" _
                 + Maqui.Text + "','" + WMaquina1 + "','" + WMaquina2 + "','" + WMaquina3 + "','" + WMaquina4 + "','" + WMaquina5 + "','" + WMaquina6 + "','" + WMaquina7 + "','" + WMaquina8 + "','" _
                 + WMaquina9 + "','" + WMaquina10 + "','" + WMaquina11 + "','" + WMaquina12 + "','" + WMaquina13 + "','" + WMaquina14 + "','" _
                 + FechaCompro.Text + "','" + FechaSalida.Text + "','" _
                 + WOrdFecha + "','" + WOrdFechaCompro + "','" + WOrdFechaSalida + "','" + WClave + "'"

        Set rstOt = db.OpenRecordset("AltaOt " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        T$ = "Impresion de Orden de Trabajo"
        m$ = "Desea Imprimir la Orden de Trabajo"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Listagrilla.GroupSelectionFormula = "{Ot.Codigo} in " + XCodigo + " to " + XCodigo
            Listagrilla.Destination = 1
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            Listagrilla.SQLQuery = "SELECT  Ot.Codigo, Ot.Fecha, Ot.Razon, Ot.Preparacion, Ot.Solidez, Ot.Observaciones1, Ot.Observaciones2, " _
                        + "Ot.Observaciones3, Ot.Solicitante, Ot.FechaCompro, " _
                        + "Ot.Compo, Ot.Compo1, Ot.Compo2, Ot.Compo3, Ot.Compo4, Ot.Compo5, Ot.Compo6, Ot.Compo7, Ot.Compo8, Ot.Compo9, Ot.Compo10, Ot.Compo11, Ot.Compo12, Ot.Compo13, Ot.Compo14, " _
                        + "Ot.Traba, Ot.Trabajo1, Ot.Trabajo2, Ot.Trabajo3, Ot.Trabajo4, Ot.Trabajo5, Ot.Trabajo6, Ot.Trabajo7, Ot.Trabajo8, Ot.Trabajo9, Ot.Trabajo10, Ot.Trabajo11, Ot.Trabajo12, Ot.Trabajo13, Ot.Trabajo14, " _
                        + "Ot.Color, Ot.Color1, Ot.Color2, Ot.Color3, Ot.Color4, Ot.Color5, Ot.Color6, Ot.Color7, Ot.Color8, Ot.Color9, Ot.Color10, Ot.Color11, Ot.Color12, Ot.Color13, Ot.Color14, Ot.Color15, Ot.Color16, Ot.Color17, Ot.Color18, Ot.Color19, Ot.Color20, Ot.Color21, " _
                        + "Ot.Maqui, Ot.Maquina1, Ot.Maquina2, Ot.Maquina3, Ot.Maquina4, Ot.Maquina5, Ot.Maquina6, Ot.Maquina7, Ot.Maquina8, Ot.Maquina9, Ot.Maquina10, Ot.Maquina11, Ot.Maquina12, Ot.Maquina13, Ot.Maquina14, " _
                        + "OtConfig.Compo1, OtConfig.Compo2, OtConfig.Compo3, OtConfig.Compo4, OtConfig.Compo5, OtConfig.Compo6, OtConfig.Compo7, OtConfig.Compo8, OtConfig.Compo9, OtConfig.Compo10, OtConfig.Compo11, OtConfig.Compo12, OtConfig.Compo13, OtConfig.Compo14, " _
                        + "OtConfig.Trabajo1, OtConfig.Trabajo2, OtConfig.Trabajo3, OtConfig.Trabajo4, OtConfig.Trabajo5, OtConfig.Trabajo6, OtConfig.Trabajo7, OtConfig.Trabajo8, OtConfig.Trabajo9, OtConfig.Trabajo10, OtConfig.Trabajo11, OtConfig.Trabajo12, OtConfig.Trabajo13, OtConfig.Trabajo14, " _
                        + "OtConfig.Color1, OtConfig.Color2, OtConfig.Color3, OtConfig.Color4, OtConfig.Color5, OtConfig.Color6, OtConfig.Color7, OtConfig.Color8, OtConfig.Color9, OtConfig.Color10, OtConfig.Color11, OtConfig.Color12, OtConfig.Color13, OtConfig.Color14, " _
                        + "OtConfig.Color15, OtConfig.Color16, OtConfig.Color17, OtConfig.Color18, OtConfig.Color19, OtConfig.Color20, OtConfig.Color21, " _
                        + "OtConfig.Maquina1, OtConfig.Maquina2, OtConfig.Maquina3, OtConfig.Maquina4, OtConfig.Maquina5, OtConfig.Maquina6, OtConfig.Maquina7, OtConfig.Maquina8, OtConfig.Maquina9, OtConfig.Maquina10, OtConfig.Maquina11, OtConfig.Maquina12, OtConfig.Maquina13, OtConfig.Maquina14 " _
                        + "From " _
                        + DSQ + ".dbo.Ot Ot, " _
                        + DSQ + ".dbo.OtConfig OtConfig " _
                        + "Where " _
                        + "Ot.Clave = OtConfig.Clave AND " _
                        + "Ot.Codigo >= " + XCodigo + " AND Ot.Codigo <= " + XCodigo
            Listagrilla.Connect = Connect()
            Listagrilla.Action = 1
        End If
    
        Call CmdLimpiar_Click
        Fecha.SetFocus
        
    End If
        
End Sub

Private Sub CmdLimpiar_Click()

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    Razon.Text = ""
    Preparacion.Text = ""
    Solidez.Text = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Observaciones3.Text = ""
    Solicitante.Text = ""
    
    Compo1.Value = 0
    Compo2.Value = 0
    Compo3.Value = 0
    Compo4.Value = 0
    Compo5.Value = 0
    Compo6.Value = 0
    Compo7.Value = 0
    Compo8.Value = 0
    Compo9.Value = 0
    Compo10.Value = 0
    Compo11.Value = 0
    Compo12.Value = 0
    Compo13.Value = 0
    Compo14.Value = 0
    Compo.Text = ""
    
    Trabajo1.Value = 0
    Trabajo2.Value = 0
    Trabajo3.Value = 0
    Trabajo4.Value = 0
    Trabajo5.Value = 0
    Trabajo6.Value = 0
    Trabajo7.Value = 0
    Trabajo8.Value = 0
    Trabajo9.Value = 0
    Trabajo10.Value = 0
    Trabajo11.Value = 0
    Trabajo12.Value = 0
    Trabajo13.Value = 0
    Trabajo14.Value = 0
    Traba.Text = ""
    
    Color1.Value = 0
    Color2.Value = 0
    Color3.Value = 0
    Color4.Value = 0
    Color5.Value = 0
    Color6.Value = 0
    Color7.Value = 0
    Color8.Value = 0
    Color9.Value = 0
    Color10.Value = 0
    Color11.Value = 0
    Color12.Value = 0
    Color13.Value = 0
    Color14.Value = 0
    Color15.Value = 0
    Color16.Value = 0
    Color17.Value = 0
    Color18.Value = 0
    Color19.Value = 0
    Color20.Value = 0
    Color21.Value = 0
    Color.Text = ""
    
    Maquina1.Value = 0
    Maquina2.Value = 0
    Maquina3.Value = 0
    Maquina4.Value = 0
    Maquina5.Value = 0
    Maquina6.Value = 0
    Maquina7.Value = 0
    Maquina8.Value = 0
    Maquina9.Value = 0
    Maquina10.Value = 0
    Maquina11.Value = 0
    Maquina12.Value = 0
    Maquina13.Value = 0
    Maquina14.Value = 0
    Maqui.Text = ""
    
    FechaCompro.Text = "  /  /    "
    FechaSalida.Text = "  /  /    "
    
    Fecha.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgAltaOt.Hide
    Unload Me
    PrgOt.Show
End Sub

Private Sub Command1_Click()

    WCero = "0"

    Sql1 = "UPDATE Ot SET "
    Sql2 = "Color15 = " + "'" + WCero + "',"
    Sql3 = "Color16 = " + "'" + WCero + "',"
    Sql4 = "Color17 = " + "'" + WCero + "',"
    Sql5 = "Color18 = " + "'" + WCero + "',"
    Sql6 = "Color19 = " + "'" + WCero + "',"
    Sql7 = "Color20 = " + "'" + WCero + "',"
    sql8 = "Color21 = " + "'" + WCero + "'"
                     
    spOt = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + sql8
    Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Config_Click()
    PrgConfigOt.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" And Cliente.Text <> Space$(6) Then
            Cliente.Text = UCase(Cliente.Text)
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Razon.Text = rstCliente!Razon
                rstCliente.Close
                Razon.SetFocus
                    Else
                Cliente.SetFocus
            End If
                Else
            Razon.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    Sql1 = "Select *"
    Sql2 = " FROM OtConfig"
    Sql3 = " Where Clave = 1"
    spOtConfig = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstOtConfig = db.OpenRecordset(spOtConfig, dbOpenSnapshot, dbSQLPassThrough)
    If rstOtConfig.RecordCount > 0 Then
        Color1.Caption = rstOtConfig!Color1
        Color2.Caption = rstOtConfig!Color2
        Color3.Caption = rstOtConfig!Color3
        Color4.Caption = rstOtConfig!Color4
        Color5.Caption = rstOtConfig!Color5
        Color6.Caption = rstOtConfig!Color6
        Color7.Caption = rstOtConfig!Color7
        Color8.Caption = rstOtConfig!Color8
        Color9.Caption = rstOtConfig!Color9
        Color10.Caption = rstOtConfig!Color10
        Color11.Caption = rstOtConfig!Color11
        Color12.Caption = rstOtConfig!Color12
        Color13.Caption = rstOtConfig!Color13
        Color14.Caption = rstOtConfig!Color14
        Color15.Caption = rstOtConfig!Color15
        Color16.Caption = rstOtConfig!Color16
        Color17.Caption = rstOtConfig!Color17
        Color18.Caption = rstOtConfig!Color18
        Color19.Caption = rstOtConfig!Color19
        Color20.Caption = rstOtConfig!Color20
        Color21.Caption = rstOtConfig!Color21
        Compo1.Caption = rstOtConfig!Compo1
        Compo2.Caption = rstOtConfig!Compo2
        Compo3.Caption = rstOtConfig!Compo3
        Compo4.Caption = rstOtConfig!Compo4
        Compo5.Caption = rstOtConfig!Compo5
        Compo6.Caption = rstOtConfig!Compo6
        Compo7.Caption = rstOtConfig!Compo7
        Compo8.Caption = rstOtConfig!Compo8
        Compo9.Caption = rstOtConfig!Compo9
        Compo10.Caption = rstOtConfig!Compo10
        Compo11.Caption = rstOtConfig!Compo11
        Compo12.Caption = rstOtConfig!Compo12
        Compo13.Caption = rstOtConfig!Compo13
        Compo14.Caption = rstOtConfig!Compo14
        Trabajo1.Caption = rstOtConfig!Trabajo1
        Trabajo2.Caption = rstOtConfig!Trabajo2
        Trabajo3.Caption = rstOtConfig!Trabajo3
        Trabajo4.Caption = rstOtConfig!Trabajo4
        Trabajo5.Caption = rstOtConfig!Trabajo5
        Trabajo6.Caption = rstOtConfig!Trabajo6
        Trabajo7.Caption = rstOtConfig!Trabajo7
        Trabajo8.Caption = rstOtConfig!Trabajo8
        Trabajo9.Caption = rstOtConfig!Trabajo9
        Trabajo10.Caption = rstOtConfig!Trabajo10
        Trabajo11.Caption = rstOtConfig!Trabajo11
        Trabajo12.Caption = rstOtConfig!Trabajo12
        Trabajo13.Caption = rstOtConfig!Trabajo13
        Trabajo14.Caption = rstOtConfig!Trabajo14
        Maquina1.Caption = rstOtConfig!Maquina1
        Maquina2.Caption = rstOtConfig!Maquina2
        Maquina3.Caption = rstOtConfig!Maquina3
        Maquina4.Caption = rstOtConfig!Maquina4
        Maquina5.Caption = rstOtConfig!Maquina5
        Maquina6.Caption = rstOtConfig!Maquina6
        Maquina7.Caption = rstOtConfig!Maquina7
        Maquina8.Caption = rstOtConfig!Maquina8
        Maquina9.Caption = rstOtConfig!Maquina9
        Maquina10.Caption = rstOtConfig!Maquina10
        Maquina11.Caption = rstOtConfig!Maquina11
        Maquina12.Caption = rstOtConfig!Maquina12
        Maquina13.Caption = rstOtConfig!Maquina13
        Maquina14.Caption = rstOtConfig!Maquina14
        rstOtConfig.Close
    End If
End Sub

Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Preparacion.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub Preparacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solidez.SetFocus
    End If
    If KeyAscii = 27 Then
        Preparacion.Text = ""
    End If
End Sub

Private Sub Solidez_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones1.SetFocus
    End If
    If KeyAscii = 27 Then
        Solidez.Text = ""
    End If
End Sub

Private Sub Observaciones1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones2.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones1.Text = ""
    End If
End Sub

Private Sub Observaciones2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones3.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones2.Text = ""
    End If
End Sub

Private Sub Observaciones3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solicitante.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones3.Text = ""
    End If
End Sub

Private Sub Solicitante_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaCompro.SetFocus
    End If
    If KeyAscii = 27 Then
        Solicitante.Text = ""
    End If
End Sub

Private Sub FechaCompro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaCompro.Text, Auxi)
        If Auxi = "S" Or FechaCompro.Text = "  /  /    " Then
            FechaSalida.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaCompro.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaSalida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaSalida.Text, Auxi)
        If Auxi = "S" Or FechaSalida.Text = "  /  /    " Then
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaSalida.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Call Opcion_Click
End Sub

Private Sub Opcion_Click()

    Call Limpia_Vector
    XIndice = 0
    Lugar = 0
    
    Select Case XIndice
        Case 0
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugar = Lugar + 1
                            Pantalla.Row = Lugar
                            Pantalla.Col = 1
                            Pantalla.Text = rstClientes!Cliente
                            Pantalla.Col = 2
                            Pantalla.Text = rstClientes!Razon
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Cliente.Text = Pantalla.TextMatrix(Pantalla.Row, 1)
            Call Cliente_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    Call Limpia_Vector
    XIndice = 0
    Lugar = 0
    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 0
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstClientes!Razon) - WEspacios
                            For Aaa = 1 To da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstClientes!Razon, Aaa, WEspacios) Then
                                    Lugar = Lugar + 1
                                    Pantalla.Row = Lugar
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstClientes!Cliente
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstClientes!Razon
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
    
        Case Else
    End Select
    
    End If

End Sub

Private Sub Form_Load()

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    Razon.Text = ""
    Preparacion.Text = ""
    Solidez.Text = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Observaciones3.Text = ""
    Solicitante.Text = ""
    
    Compo1.Value = 0
    Compo2.Value = 0
    Compo3.Value = 0
    Compo4.Value = 0
    Compo5.Value = 0
    Compo6.Value = 0
    Compo7.Value = 0
    Compo8.Value = 0
    Compo9.Value = 0
    Compo10.Value = 0
    Compo11.Value = 0
    Compo12.Value = 0
    Compo13.Value = 0
    Compo14.Value = 0
    Compo.Text = ""
    
    Trabajo1.Value = 0
    Trabajo2.Value = 0
    Trabajo3.Value = 0
    Trabajo4.Value = 0
    Trabajo5.Value = 0
    Trabajo6.Value = 0
    Trabajo7.Value = 0
    Trabajo8.Value = 0
    Trabajo9.Value = 0
    Trabajo10.Value = 0
    Trabajo11.Value = 0
    Trabajo12.Value = 0
    Trabajo13.Value = 0
    Trabajo14.Value = 0
    Traba.Text = ""
    
    Color1.Value = 0
    Color2.Value = 0
    Color3.Value = 0
    Color4.Value = 0
    Color5.Value = 0
    Color6.Value = 0
    Color7.Value = 0
    Color8.Value = 0
    Color9.Value = 0
    Color10.Value = 0
    Color11.Value = 0
    Color12.Value = 0
    Color13.Value = 0
    Color14.Value = 0
    Color15.Value = 0
    Color16.Value = 0
    Color17.Value = 0
    Color18.Value = 0
    Color19.Value = 0
    Color20.Value = 0
    Color21.Value = 0
    Color.Text = ""
    
    Maquina1.Value = 0
    Maquina2.Value = 0
    Maquina3.Value = 0
    Maquina4.Value = 0
    Maquina5.Value = 0
    Maquina6.Value = 0
    Maquina7.Value = 0
    Maquina8.Value = 0
    Maquina9.Value = 0
    Maquina10.Value = 0
    Maquina11.Value = 0
    Maquina12.Value = 0
    Maquina13.Value = 0
    Maquina14.Value = 0
    Maqui.Text = ""
    
    FechaCompro.Text = "  /  /    "
    FechaSalida.Text = "  /  /    "
    
    If Val(WOt) <> 0 Then
        spOt = "ConsultaOt " + "'" + WOt + "'"
        Set rstOt = db.OpenRecordset(spOt, dbOpenSnapshot, dbSQLPassThrough)
        If rstOt.RecordCount > 0 Then
            Fecha.Text = rstOt!Fecha
            Cliente.Text = rstOt!Cliente
            Razon.Text = rstOt!Razon
            Preparacion.Text = rstOt!Preparacion
            Solidez.Text = rstOt!Solidez
            Observaciones1.Text = rstOt!Observaciones1
            Observaciones2.Text = rstOt!Observaciones2
            Observaciones3.Text = rstOt!Observaciones3
            Solicitante.Text = rstOt!Solicitante
            FechaCompro.Text = rstOt!FechaCompro
            FechaSalida.Text = rstOt!FechaSalida
            Compo.Text = rstOt!Compo
            Compo1.Value = rstOt!Compo1
            Compo2.Value = rstOt!Compo2
            Compo3.Value = rstOt!Compo3
            Compo4.Value = rstOt!Compo4
            Compo5.Value = rstOt!Compo5
            Compo6.Value = rstOt!Compo6
            Compo7.Value = rstOt!Compo7
            Compo8.Value = rstOt!Compo8
            Compo9.Value = rstOt!Compo9
            Compo10.Value = rstOt!Compo10
            Compo11.Value = rstOt!Compo11
            Compo12.Value = rstOt!Compo12
            Compo13.Value = rstOt!Compo13
            Compo14.Value = rstOt!Compo14
            Traba.Text = rstOt!Traba
            Trabajo1.Value = rstOt!Trabajo1
            Trabajo2.Value = rstOt!Trabajo2
            Trabajo3.Value = rstOt!Trabajo3
            Trabajo4.Value = rstOt!Trabajo4
            Trabajo5.Value = rstOt!Trabajo5
            Trabajo6.Value = rstOt!Trabajo6
            Trabajo7.Value = rstOt!Trabajo7
            Trabajo8.Value = rstOt!Trabajo8
            Trabajo9.Value = rstOt!Trabajo9
            Trabajo10.Value = rstOt!Trabajo10
            Trabajo11.Value = rstOt!Trabajo11
            Trabajo12.Value = rstOt!Trabajo12
            Trabajo13.Value = rstOt!Trabajo13
            Trabajo14.Value = rstOt!Trabajo14
            Color.Text = rstOt!Color
            Color1.Value = rstOt!Color1
            Color2.Value = rstOt!Color2
            Color3.Value = rstOt!Color3
            Color4.Value = rstOt!Color4
            Color5.Value = rstOt!Color5
            Color6.Value = rstOt!Color6
            Color7.Value = rstOt!Color7
            Color8.Value = rstOt!Color8
            Color9.Value = rstOt!Color9
            Color10.Value = rstOt!Color10
            Color11.Value = rstOt!Color11
            Color12.Value = rstOt!Color12
            Color13.Value = rstOt!Color13
            Color14.Value = rstOt!Color14
            Color15.Value = rstOt!Color15
            Color16.Value = rstOt!Color16
            Color17.Value = rstOt!Color17
            Color18.Value = rstOt!Color18
            Color19.Value = rstOt!Color19
            Color20.Value = rstOt!Color20
            Color21.Value = rstOt!Color21
            Maqui.Text = rstOt!Maqui
            Maquina1.Value = rstOt!Maquina1
            Maquina2.Value = rstOt!Maquina2
            Maquina3.Value = rstOt!Maquina3
            Maquina4.Value = rstOt!Maquina4
            Maquina5.Value = rstOt!Maquina5
            Maquina6.Value = rstOt!Maquina6
            Maquina7.Value = rstOt!Maquina7
            Maquina8.Value = rstOt!Maquina8
            Maquina9.Value = rstOt!Maquina9
            Maquina10.Value = rstOt!Maquina10
            Maquina11.Value = rstOt!Maquina11
            Maquina12.Value = rstOt!Maquina12
            Maquina13.Value = rstOt!Maquina13
            Maquina14.Value = rstOt!Maquina14
            rstOt.Close
        End If
        
    End If
        
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Preparacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Solidez_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Solicitante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fechacompro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub FechaSalida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub
Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdGraba_Click
        Case 113
            Call CmdLimpiar_Click
        Case 114
            Call Consulta_Click
        Case 115
            Call Composicion_Click
        Case 116
            Call Trabajo_Click
        Case 117
            Call Colorante_Click
        Case 118
            Call Maquina_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Trabajo_Click()
    AltaTrabajo.Height = 4095
    AltaTrabajo.Left = 1680
    AltaTrabajo.Top = 600
    AltaTrabajo.Width = 5895
    AltaTrabajo.Visible = True
End Sub

Private Sub Maquina_Click()
    AltaMaquina.Height = 4095
    AltaMaquina.Left = 1680
    AltaMaquina.Top = 600
    AltaMaquina.Width = 5895
    AltaMaquina.Visible = True
End Sub

Private Sub Colorante_Click()
    AltaColor.Height = 4095
    AltaColor.Left = 360
    AltaColor.Top = 600
    AltaColor.Width = 8775
    AltaColor.Visible = True
End Sub

Private Sub Composicion_Click()
    AltaCompo.Height = 4095
    AltaCompo.Left = 1680
    AltaCompo.Top = 600
    AltaCompo.Width = 5895
    AltaCompo.Visible = True
End Sub

Private Sub FinTrabajo_Click()
    AltaTrabajo.Visible = False
End Sub

Private Sub FinMaquina_Click()
    AltaMaquina.Visible = False
End Sub

Private Sub FinColor_Click()
    AltaColor.Visible = False
End Sub

Private Sub FinCompo_Click()
    AltaCompo.Visible = False
End Sub

Private Sub Limpia_Vector()

    Pantalla.Clear
    Pantalla.Font.Bold = True

    Pantalla.FixedCols = 1
    Pantalla.Cols = 3
    Pantalla.FixedRows = 1
    Pantalla.Rows = 10001
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Select Case Ciclo
            Case 1
                Pantalla.Text = "Codigo"
                Pantalla.ColWidth(Ciclo) = 1800
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Pantalla.Text = "Razon Social"
                Pantalla.ColWidth(Ciclo) = 5800
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    WAncho = 340
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub

