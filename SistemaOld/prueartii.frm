VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPruartii 
   Caption         =   "Ingreso de Ensayos de Materia Prima"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8145
   ScaleWidth      =   11880
   Begin VB.Frame IngresaEstado 
      Caption         =   " "
      Height          =   840
      Left            =   9120
      TabIndex        =   169
      Top             =   5400
      Visible         =   0   'False
      Width           =   1290
      Begin VB.CommandButton ConfirmaEstado 
         Caption         =   "Confirma "
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
         Left            =   2640
         TabIndex        =   176
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Certificado1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   173
         Top             =   1320
         Width           =   135
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
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   172
         Text            =   " "
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CheckBox Estado1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   171
         Top             =   1680
         Width           =   135
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
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   170
         Text            =   " "
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "ACTUALIZACION DE DATOS DE CERTIFICADO DE ANALISIS Y ESTADO DE ENVASES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   177
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label70 
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
         TabIndex        =   175
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label69 
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
         TabIndex        =   174
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Frame NroLote 
      Height          =   495
      Left            =   480
      TabIndex        =   124
      Top             =   4680
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton FinNroLote 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   168
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label68 
         Caption         =   "72.000   a 72.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   167
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label67 
         Caption         =   "71.000   a 71.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   166
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label66 
         Caption         =   "76.000   a 76.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   165
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label65 
         Caption         =   "73.000   a 73.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   164
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label64 
         Caption         =   "75.000   a 75.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   163
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label63 
         Caption         =   "70.000   a 70.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   162
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label62 
         Caption         =   "995.000 a 999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   161
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label61 
         Caption         =   "74.000   a 74.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   160
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label60 
         Caption         =   "78.000   a 78.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   159
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label59 
         Caption         =   "590.000 a 594.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   158
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "690.000 a 694.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   157
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label57 
         Caption         =   "790.000 a 794.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   156
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label56 
         Caption         =   "890.000 a 894.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   155
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label55 
         Caption         =   "490.000 a 494.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   154
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label54 
         Caption         =   "190.000 a 194.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   153
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label53 
         Caption         =   "990.000 a 994.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   152
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label52 
         Caption         =   "290.000 a 294.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   151
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label51 
         Caption         =   "390.000 a 394.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   150
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label50 
         Caption         =   "500.000 a 589.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   149
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label49 
         Caption         =   "600.000 a 689.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   148
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label48 
         Caption         =   "700.000 a 789.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   147
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "800.000 a 889.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   146
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "400.000 a 489.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   145
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label45 
         Caption         =   "300.000 a 389.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   144
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "200.000 a 289.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   143
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "900.000 a 949.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   142
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label42 
         Caption         =   "100.000 a 189.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   141
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "PELLITAL III"
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
         TabIndex        =   140
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label40 
         Caption         =   "PELLITAL II"
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
         TabIndex        =   139
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label39 
         Caption         =   "PELLITAL I"
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
         TabIndex        =   138
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label38 
         Caption         =   "SURFACTAN  V"
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
         TabIndex        =   137
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "SURFACTAN  IV"
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
         TabIndex        =   136
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "SURFACTAN  III"
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
         TabIndex        =   135
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label35 
         Caption         =   "SURFACTAN  II"
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
         TabIndex        =   134
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label34 
         Caption         =   "SURFACTAN  I (Colorantes)"
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
         TabIndex        =   133
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label33 
         Caption         =   "SURFACTAN  I"
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
         TabIndex        =   132
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "RECHAZADO"
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
         Left            =   6840
         TabIndex        =   131
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "ROJO (NO OK)"
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
         Left            =   6840
         TabIndex        =   130
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Aprob. por Desvio"
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
         Left            =   4680
         TabIndex        =   129
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "AMARILLO (CR)"
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
         Left            =   4680
         TabIndex        =   128
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "APROBADO"
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
         Left            =   2520
         TabIndex        =   127
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "VERDE (OK)"
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
         Left            =   2520
         TabIndex        =   126
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numeros de Lotes Reservados para cada Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   125
         Top             =   240
         Width           =   8295
      End
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
      Left            =   240
      TabIndex        =   123
      Top             =   5400
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "Modificacion de  Prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   459
      Left            =   9600
      TabIndex        =   118
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Pass 
      Height          =   1815
      Left            =   3600
      TabIndex        =   114
      Top             =   1440
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton WCancela 
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
         Left            =   600
         TabIndex        =   116
         Top             =   1200
         Width           =   2055
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
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   115
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label23 
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
         TabIndex        =   117
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Modif 
      Caption         =   "Modificacion de Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   720
      TabIndex        =   105
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      Begin MSMask.MaskEdBox Modif_Recibido 
         Height          =   285
         Left            =   2400
         TabIndex        =   113
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Modif_Solicitado 
         Height          =   285
         Left            =   2400
         TabIndex        =   112
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox Modif_Orden 
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   111
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Modif_Cancela 
         Caption         =   "Cancela"
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
         Left            =   2280
         TabIndex        =   110
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Modif_Confirma 
         Caption         =   "Confirma "
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
         Left            =   480
         TabIndex        =   109
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Producto recibido"
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
         TabIndex        =   108
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Producto Solicitatado"
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
         TabIndex        =   107
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label20 
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
         Left            =   240
         TabIndex        =   106
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "  Modificacion          de O/C"
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
      Left            =   6240
      TabIndex        =   104
      Top             =   4800
      Width           =   1455
   End
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
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   97
      Text            =   " "
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Impensayo 
      Caption         =   "Impresion Prueba"
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
      Left            =   7920
      TabIndex        =   93
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5760
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   5895
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   300
         Left            =   4440
         TabIndex        =   103
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   300
         Left            =   4440
         TabIndex        =   102
         Top             =   240
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
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1560
         TabIndex        =   99
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1560
         TabIndex        =   98
         Top             =   240
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
      Begin VB.Frame Frame3 
         Caption         =   "Destino"
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
         Left            =   240
         TabIndex        =   90
         Top             =   1200
         Width           =   1695
         Begin VB.OptionButton ImprePantalla 
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
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton ImpreListado 
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   975
         Left            =   2160
         TabIndex        =   87
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Rechazo 
            Caption         =   "Rechazados"
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
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Aprobado 
            Caption         =   "Aprobados"
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
            TabIndex        =   88
            Top             =   240
            Width           =   1335
         End
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
         Height          =   375
         Left            =   4320
         TabIndex        =   40
         Top             =   1200
         Width           =   1095
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
         Height          =   375
         Left            =   4320
         TabIndex        =   39
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Hasta fecha"
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
         Left            =   3120
         TabIndex        =   101
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   3120
         TabIndex        =   100
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Codigo"
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
         TabIndex        =   38
         Top             =   720
         Width           =   1215
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox orden 
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
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   86
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame panLote 
      Caption         =   "Grabacion de Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox OrigenMercaderia 
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
         Left            =   8760
         MaxLength       =   30
         TabIndex        =   121
         Text            =   " "
         Top             =   480
         Width           =   2535
      End
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   119
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox NroRechazo 
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
         TabIndex        =   95
         Text            =   " "
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CancelaLote 
         Caption         =   "Cancela Operacion"
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
         Left            =   6120
         TabIndex        =   61
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton GrabaLote 
         Caption         =   "Graba Prueba"
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
         Left            =   3600
         TabIndex        =   60
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Nueva 
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
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   59
         Text            =   " "
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Devuelta 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   58
         Text            =   " "
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Liberada 
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
         MaxLength       =   10
         TabIndex        =   57
         Text            =   " "
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
         Left            =   360
         MaxLength       =   6
         TabIndex        =   56
         Text            =   " "
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Origen Mercaderia"
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
         Left            =   9240
         TabIndex        =   122
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Nro Partida Proveedor"
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
         TabIndex        =   120
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Nro Rechazo"
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
         TabIndex        =   94
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Nueva O/C"
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
         Left            =   5520
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Canti.Devuelta"
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
         Left            =   2640
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Cant.Liberada"
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
         Left            =   1200
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Prueba"
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
         TabIndex        =   52
         Top             =   240
         Width           =   615
      End
   End
   Begin MSMask.MaskEdBox fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   63
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton CmdAddRechazo 
      Caption         =   "Graba Rechazo"
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
      Left            =   6240
      TabIndex        =   50
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   49
      Text            =   " "
      Top             =   4320
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
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   4080
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
      TabIndex        =   47
      Text            =   " "
      Top             =   3840
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
      TabIndex        =   46
      Text            =   " "
      Top             =   3600
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   10800
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wpruart.rpt"
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
      Left            =   1440
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   5880
      TabIndex        =   34
      Top             =   3720
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   32
      Text            =   " "
      Top             =   3240
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   3000
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   2760
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   2520
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   2280
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   2040
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   1800
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   1560
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   1320
      Width           =   3855
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
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
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
      Height          =   2220
      ItemData        =   "prueartii.frx":0000
      Left            =   240
      List            =   "prueartii.frx":0007
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Listado 
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
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
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
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Pantalla"
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
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
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
      Left            =   7920
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddlote 
      Caption         =   "Graba  Prueba"
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
      Left            =   7920
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Nro Prueba"
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
      Left            =   5400
      TabIndex        =   96
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label17 
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
      Left            =   5400
      TabIndex        =   85
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Ensayo10 
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
      Left            =   120
      TabIndex        =   84
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Ensayo9 
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
      Left            =   120
      TabIndex        =   83
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Ensayo8 
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
      Left            =   120
      TabIndex        =   82
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Ensayo7 
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
      Left            =   120
      TabIndex        =   81
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Ensayo6 
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
      Left            =   120
      TabIndex        =   80
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Ensayo5 
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
      Left            =   120
      TabIndex        =   79
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Ensayo4 
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
      Height          =   375
      Left            =   120
      TabIndex        =   78
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Ensayo3 
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
      Left            =   120
      TabIndex        =   77
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Ensayo2 
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
      Left            =   120
      TabIndex        =   76
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Ensayo1 
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
      Left            =   120
      TabIndex        =   75
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Std10 
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
      Left            =   4200
      TabIndex        =   74
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Std9 
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
      Left            =   4200
      TabIndex        =   73
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Std8 
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
      Left            =   4200
      TabIndex        =   72
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Std7 
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
      Left            =   4200
      TabIndex        =   71
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Std6 
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
      Left            =   4200
      TabIndex        =   70
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Std5 
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
      Left            =   4200
      TabIndex        =   69
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Std4 
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
      Left            =   4200
      TabIndex        =   68
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Std3 
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
      Left            =   4200
      TabIndex        =   67
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Std2 
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
      Left            =   4200
      TabIndex        =   66
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Std1 
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
      Left            =   4200
      TabIndex        =   65
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Obtenido"
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
      Left            =   7920
      TabIndex        =   64
      Top             =   720
      Width           =   3855
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
      TabIndex        =   62
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   4320
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   4080
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3840
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3600
      Width           =   1455
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
      TabIndex        =   41
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
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
      Left            =   1320
      TabIndex        =   33
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Descri10 
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
      Left            =   1080
      TabIndex        =   22
      Top             =   3240
      Width           =   3060
   End
   Begin VB.Label Descri9 
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
      Left            =   1080
      TabIndex        =   21
      Top             =   3000
      Width           =   3060
   End
   Begin VB.Label Descri8 
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
      Left            =   1080
      TabIndex        =   20
      Top             =   2760
      Width           =   3060
   End
   Begin VB.Label Descri7 
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
      Left            =   1080
      TabIndex        =   19
      Top             =   2520
      Width           =   3060
   End
   Begin VB.Label Descri6 
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
      Left            =   1080
      TabIndex        =   18
      Top             =   2280
      Width           =   3060
   End
   Begin VB.Label Descri5 
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
      Left            =   1080
      TabIndex        =   17
      Top             =   2040
      Width           =   3060
   End
   Begin VB.Label Descri4 
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
      Left            =   1080
      TabIndex        =   16
      Top             =   1800
      Width           =   3060
   End
   Begin VB.Label Descri3 
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
      Left            =   1080
      TabIndex        =   15
      Top             =   1560
      Width           =   3060
   End
   Begin VB.Label descri2 
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
      Left            =   1080
      TabIndex        =   14
      Top             =   1320
      Width           =   3060
   End
   Begin VB.Label Descri1 
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
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Width           =   3060
   End
   Begin VB.Label lblresultado 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Standard"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblDescri 
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
      Left            =   1080
      TabIndex        =   11
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblensayo 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ensayo"
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
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   375
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "PrgPruartii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInforme As String
Private Pasa As String
Private Auxi3 As String
Private Auxi4 As String
Private WLote As String
Private SaldoOrden As Double
Dim rstPrueart As Recordset
Dim spPrueart As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim XParam As String
Dim WCosto1 As String
Dim WCosto3 As String
Dim WPrecio As Double
Dim XStock As Double
Dim XCosto As Double
Dim XCostoTotal As Double
Dim XStock1 As Double
Dim XCosto1 As Double
Dim XCostoTotal1 As Double
Dim XStock2 As Double
Dim XCosto2 As Double
Dim XCostoTotal2 As Double
Dim XCosto3 As Double
Dim WTipoOrden As Single
Dim WOrigen  As String
Dim EmpresaActual As String
Dim WRecibida As Double
Dim WLaudada As Double
Dim WCantidad1 As Double
Dim WCantidad2 As Double
Dim XCertificado1 As Integer
Dim XCertificado2 As String
Dim XEstado1 As Integer
Dim XEstado2 As String

Private Sub Acepta_Click()

    If Aprobado.Value = True Then
        Desdepru = "100000"
        HastaPru = "199999"
            Else
        Desdepru = "200000"
        HastaPru = "299999"
    End If
    
    WAno = Right$(Desdefec.Text, 4)
    WMes = Mid$(Desdefec.Text, 4, 2)
    WDia = Left$(Desdefec.Text, 2)
    FDesde = WAno + WMes + WDia
    WAno = Right$(Hastafec.Text, 4)
    WMes = Mid$(Hastafec.Text, 4, 2)
    WDia = Left$(Hastafec.Text, 2)
    Fhasta = WAno + WMes + WDia

    
    lista.ReportFileName = "WPruArt.rpt"
    
    lista.WindowTitle = "Listado de Controles de Materias Primas"
    lista.WindowTop = -3
    lista.WindowLeft = -3
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Uno = "{Prueart.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Dos = " and {Prueart.Fechaord} in " + Chr$(34) + FDesde + Chr$(34) + " to " + Chr$(34) + Fhasta + Chr$(34)
    Tres = " and {Prueart.Prueba} in " + Chr$(34) + Desdepru + Chr$(34) + " to " + Chr$(34) + HastaPru + Chr$(34)
    
    lista.GroupSelectionFormula = Uno + Dos + Tres
    
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    lista.SQLQuery = "SELECT prueart.Prueba, prueart.Producto, prueart.Fecha, prueart.Orden, prueart.Lote, Articulo.Descripcion " _
                     + "From " + DSQ + ".dbo.prueart prueart, " _
                     + DSQ + ".dbo.Articulo Articulo " _
                     + "Where prueart.Producto = Articulo.Codigo AND prueart.Prueba >= ' ' AND prueart.Prueba <= '9999999'"
    
    lista.DataFiles(2) = WEmpresa + "auxi.mdb"
    lista.Connect = Connect()
    
    lista.Action = 1
    Frame2.Visible = False
End Sub


Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()
    Rem With rstprueart
    Rem     .Index = "Prueba"
    Rem     ClavePrue$ = "1999999"
    Rem     .Seek "<", ClavePrue$
    Rem     If .NoMatch Then
    Rem         Lote.Text = "1"
    Rem             Else
    Rem         Lote.Text = Str$(Val(!Prueba) + 1)
    Rem     End If
    Rem
    Rem     Auxi1 = Lote.Text
    Rem     Call Ceros(Auxi1, 6)
    Rem     Lote.Text = Auxi1
    Rem
    Rem     Auxi = "1"
    Rem
    Rem     Liberada.Text = ""
    Rem     Devuelta.Text = ""
    Rem     NroRechazo.Text = ""
    Rem     Nueva.Text = ""
    Rem
    Rem     panLote.Visible = True
    Rem
    Rem     Liberada.SetFocus
    Rem
    Rem End With
    
    Auxi = "1"
    
    Lote.Text = ""
    Liberada.Text = ""
    Devuelta.Text = ""
    NroRechazo.Text = ""
    Nueva.Text = ""
    PartidaProveedor.Text = ""
    OrigenMercaderia.Text = WOrigen
        
    panLote.Visible = True
    
    Lote.SetFocus
        
End Sub

Private Sub cmdAddRechazo_Click()

    spPrueart = "ConsultaPruebaMenor " + "'" + "2999999" + "'"
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        If rstPrueart!Tipo = "2" Then
            Lote.Text = Str$(Val(rstPrueart!Prueba) + 1)
                Else
            Lote.Text = "2000"
        End If
        rstPrueart.Close
            Else
        Lote.Text = "1"
    End If
    
    Auxi1 = Lote.Text
    Call Ceros(Auxi1, 6)
    Lote.Text = Auxi1
        
    Auxi = "2"
            
    Liberada.Text = ""
    Devuelta.Text = ""
    NroRechazo.Text = ""
    Nueva.Text = ""
        
    panLote.Visible = True
        
    Liberada.SetFocus
        
End Sub

Private Sub FinNroLote_Click()
    NroLote.Visible = False
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
    OPEN_FILE_Empresa
End Sub

Private Sub GrabaLote_Click()

    On Error GoTo WError

    Entra = "N"

    Select Case Val(WEmpresa)
        Case 1
            If Val(Lote.Text) >= 100000 And Val(Lote.Text) <= 189999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 190000 And Val(Lote.Text) <= 194999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 70000 And Val(Lote.Text) <= 70999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 900000 And Val(Lote.Text) <= 949999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 990000 And Val(Lote.Text) <= 994999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 995000 And Val(Lote.Text) <= 999999 Then
                Entra = "S"
            End If
        Case 2
            If Val(Lote.Text) >= 600000 And Val(Lote.Text) <= 689999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 690000 And Val(Lote.Text) <= 694999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 71000 And Val(Lote.Text) <= 71999 Then
                Entra = "S"
            End If
        Case 3
            If Val(Lote.Text) >= 200000 And Val(Lote.Text) <= 289999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 290000 And Val(Lote.Text) <= 294999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 74000 And Val(Lote.Text) <= 74999 Then
                Entra = "S"
            End If
        Case 4
            If Val(Lote.Text) >= 700000 And Val(Lote.Text) <= 789999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 790000 And Val(Lote.Text) <= 794999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 76000 And Val(Lote.Text) <= 76999 Then
                Entra = "S"
            End If
        Case 5
            If Val(Lote.Text) >= 300000 And Val(Lote.Text) <= 389999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 390000 And Val(Lote.Text) <= 394999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 78000 And Val(Lote.Text) <= 78999 Then
                Entra = "S"
            End If
        Case 6
            If Val(Lote.Text) >= 400000 And Val(Lote.Text) <= 489999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 490000 And Val(Lote.Text) <= 494999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 75000 And Val(Lote.Text) <= 75999 Then
                Entra = "S"
            End If
        Case 7
            If Val(Lote.Text) >= 500000 And Val(Lote.Text) <= 589999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 590000 And Val(Lote.Text) <= 594999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 72000 And Val(Lote.Text) <= 72999 Then
                Entra = "S"
            End If
        Case 8
            If Val(Lote.Text) >= 800000 And Val(Lote.Text) <= 889999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 890000 And Val(Lote.Text) <= 894999 Then
                Entra = "S"
            End If
            If Val(Lote.Text) >= 73000 And Val(Lote.Text) <= 73999 Then
                Entra = "S"
            End If
        Case Else
            Entra = "S"
    End Select
    
    If Entra = "N" Then
        NroLote.Height = 5775
        NroLote.Left = 1320
        NroLote.Top = 360
        NroLote.Width = 8775
        NroLote.Visible = True
        Exit Sub
    End If

    If Val(Liberada.Text) = 0 Then
        Liberada.Text = "0"
    End If
    
    If Val(Devuelta.Text) = 0 Then
        Devuelta.Text = "0"
    End If
    
    If Val(NroRechazo.Text) = 0 Then
        NroRechazo.Text = "0"
    End If
    
    Cantidad = Val(Liberada.Text) + Val(Devuelta.Text)

    Call Calcula_SaldoOrden

    If Cantidad > SaldoOrden Then
        m$ = "La Cantidad supera al saldo del informe de recepcion" + Chr$(13) _
             + "Cantidad recibida (Informe de recepcion) : " + Str$(WRecibida) + Chr$(13) _
             + "Cantidad Laudada (Laudos Anteriores) : " + Str$(WLaudada) + Chr$(13) _
             + "Saldo Disponible para laudar : " + Str$(SaldoOrden)
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Exit Sub
            Else
        If Cantidad < SaldoOrden Then
            m$ = "Atencion : la cantidad laudada es menor al saldo del informe de recepcion" + Chr$(13) _
                + "Cantidad recibida (Informe de recepcion) : " + Str$(WRecibida) + Chr$(13) _
                + "Cantidad Laudada (Laudos Anteriores) : " + Str$(WLaudada) + Chr$(13) _
                + "Saldo Disponible para laudar : " + Str$(SaldoOrden) + Chr$(13) _
                + "Saldo Pendiente para futuros laudos : " + Str$(SaldoOrden - Cantidad) + Chr$(13) _
                + "Confirma la grabacion del LAUDO"
                Respuesta% = MsgBox(m$, 32 + 4, "Ingreso de Pruebas")
                If Respuesta% = 7 Then
                    Exit Sub
                End If
        End If
    End If
    
    XEnvase = 0
    
    spInforme = "ConsultaInformeOrden " + "'" + orden.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
    
        With rstInforme
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If Producto.Text = rstInforme!Articulo Then
                    Informe = rstInforme!Informe
                    XEnvase = rstInforme!envase
                    XCertificado1 = IIf(IsNull(rstInforme!Certificado1), "0", rstInforme!Certificado1)
                    XCertificado2 = IIf(IsNull(rstInforme!Certificado2), "", rstInforme!Certificado2)
                    XEstado1 = IIf(IsNull(rstInforme!Estado1), "0", rstInforme!Estado1)
                    XEstado2 = IIf(IsNull(rstInforme!Estado2), "", rstInforme!Estado2)
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                                                                        
            Loop
            End If
            
        End With
        rstInforme.Close
    End If
    
    Call Busca_Informe(orden.Text, WInforme, Producto.Text, Pasa)
    
    If Val(Liberada.Text) > 0 Then
    
        WLaudo = Lote.Text
        WRenglon = "1"
        WFecha = Fecha.Text
        WOrden = orden.Text
        WArticulo = Producto.Text
        WLiberada = Liberada.Text
        WDevuelta = "0"
        WLote = Lote.Text
        WRechazo = ""
        WActualiza = "N"
        WMarca = ""
        WInforme = WInforme
        WSaldo = Liberada.Text
        WOrigen = OrigenMercaderia.Text
        WPartiOri = PartidaProveedor.Text
        WEnvase = Str$(XEnvase)
            
        Auxi1 = Str$(WLaudo)
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(WRenglon)
        Call Ceros(Auxi2, 2)
            
        WClave = Auxi1 + Auxi2
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                + WLaudo + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WArticulo + "','" _
                + WLiberada + "','" _
                + WDevuelta + "','" _
                + WOrden + "','" _
                + WMarca + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WInforme + "','" _
                + WActualiza + "','" _
                + WDate + "','" _
                + WSaldo + "','" _
                + WOrigen + "','" _
                + WPartiOri + "','" _
                + WEnvase + "'"
                
        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
                     
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    If Val(Devuelta.Text) > 0 Then
    
        WLaudo = NroRechazo.Text
        WRenglon = "1"
        WFecha = Fecha.Text
        WOrden = orden.Text
        WArticulo = Producto.Text
        WLiberada = ""
        WDevuelta = Devuelta.Text
        WLote = NroRechazo.Text
        WRechazo = NroRechazo.Text
        WActualiza = Nueva.Text
        WMarca = ""
        WInforme = WInforme
        WSaldo = "0"
        WOrigen = OrigenMercaderia.Text
        WPartiOri = PartidaProveedor.Text
        WEnvase = Str$(XEnvase)
            
        Auxi1 = Str$(WLaudo)
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(WRenglon)
        Call Ceros(Auxi2, 2)
            
        WClave = Auxi1 + Auxi2
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                + WLaudo + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WArticulo + "','" _
                + WLiberada + "','" _
                + WDevuelta + "','" _
                + WOrden + "','" _
                + WMarca + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WInforme + "','" _
                + WActualiza + "','" _
                + WDate + "','" _
                + WSaldo + "','" _
                + WOrigen + "','" _
                + WPartiOri + "','" _
                + WEnvase + "'"
                
        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
                     
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    WPrecio = 0
    
    For WDa% = 1 To 40
        Auxi3 = orden.Text
        Call Ceros(Auxi3, 6)
        Auxi1 = WDa%
        Call Ceros(Auxi1, 2)
        WClave = Auxi3 + Auxi1
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WMoneda = rstOrden!Moneda
            WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            If Producto.Text = rstOrden!Articulo Then
                WPrecio = rstOrden!Precio
                WLiberada = Str$(rstOrden!Liberada + Val(Liberada.Text))
                WDevuelta = Str$(rstOrden!Devuelta + Val(Devuelta.Text))
                WFechaentrega = Fecha.Text
                WDate = Date$
                rstOrden.Close
                XParam = "'" + WClave + "','" _
                    + WLiberada + "','" _
                    + WDevuelta + "','" _
                    + WFechaentrega + "','" _
                    + WDate + "'"
                spOrden = "ModificaOrdenPrueba " + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                rstOrden.Close
            End If
        End If
    Next WDa%
    
    spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
        WProducto = Producto.Text
        
        If Nueva.Text = "S" Then
            WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Liberada) - Val(Devuelta))
                Else
            WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Liberada))
        End If
        
        Select Case WTipoOrden
            Case 1, 2
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo1), "0", rstArticulo!Costo1)
                    XCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                    XCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                End If
            
            Case Else
                XStock1 = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                End If
                XCostoTotal1 = XStock1 * XCosto1
                
                XStock2 = Val(Liberada)
                XCosto2 = WPrecio
                XCostoTotal2 = XStock2 * XCosto2
                
                XCosto = 0
                XStock = XStock1 + XStock2
                XCostoTotal = XCostoTotal1 + XCostoTotal2
                If XStock <> 0 Then
                    XCosto = XCostoTotal / XStock
                End If
                
                Call Redondeo(XCosto)
                    
                WCosto1 = Str$(WPrecio)
                WCosto3 = Str$(XCosto)
                
        End Select
        
        WEntradas = Str$(rstArticulo!Entradas + Val(Liberada))
        WDate = Date$
        rstArticulo.Close
        
        If WMoneda = 0 Then
            XParam = "'" + WProducto + "','" _
                + WLaboratorio + "','" _
                + WEntradas + "','" _
                + WDate + "','" _
                + WCosto1 + "','" _
                + WCosto3 + "'"
            spArticulo = "ModificaArticuloLaudoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + WProducto + "','" _
                + WLaboratorio + "','" _
                + WEntradas + "','" _
                + WDate + "','" _
                + WCosto1 + "','" _
                + WCosto3 + "'"
            spArticulo = "ModificaArticuloLaudoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem actualiza los datos de la empresa
        
        XEmpresa = WEmpresa
        
        WCodigo = WProducto
        XParam = "'" + WCodigo + "','" _
                     + WCosto1 + "'"
            
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WMoneda = 0 Then
            spArticulo = "ModificaArticuloCostoDolares " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            spArticulo = "ModificaArticuloCostoPesos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
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
        
    End If
    
    If XCertificado1 = 0 Or XEstado1 = 0 Then
        Certificado1.Value = XCertificado1
        Estado1.Value = XEstado1
        Certificado2.Text = XCertificado2
        Estado2.Text = XEstado2
        IngresaEstado.Height = 3000
        IngresaEstado.Left = 2640
        IngresaEstado.Top = 1500
        IngresaEstado.Width = 6335
        IngresaEstado.Visible = True
            Else
        Call CmdLimpiar_Click
        panLote.Visible = False
        Producto.SetFocus
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Confirmaestado_Click()

    WCertificado1 = Str$(Certificado1.Value)
    WEstado1 = Str$(Estado1.Value)
    
    XParam = "'" + WInforme + "','" _
                + WCertificado1 + "','" _
                + Certificado2.Text + "','" _
                + WEstado1 + "','" _
                + Estado2.Text + "'"
                         
    spInforme = "ModificaInformeDatos " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    IngresaEstado.Visible = False
    Call CmdLimpiar_Click
    panLote.Visible = False
    
    Producto.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -   -   "
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    orden.Text = ""
    Ensayo1.Caption = ""
    Valor1.Text = ""
    Ensayo2.Caption = ""
    valor2.Text = ""
    Ensayo3.Caption = ""
    Valor3.Text = ""
    Ensayo4.Caption = ""
    valor4.Text = ""
    Ensayo5.Caption = ""
    valor5.Text = ""
    Ensayo6.Caption = ""
    valor6.Text = ""
    Ensayo7.Caption = ""
    valor7.Text = ""
    Ensayo8.Caption = ""
    valor8.Text = ""
    Ensayo9.Caption = ""
    valor9.Text = ""
    Ensayo10.Caption = ""
    valor10.Text = ""
    Descriprod.Caption = ""
    Descri1.Caption = ""
    descri2.Caption = ""
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
    Partida.Text = ""
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgPruart.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi4)
        If Auxi4 = "S" Then
            orden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Impensayo_Click()

    If Val(Auxi) = 0 Then
        Auxi = "0"
    End If
    
    If Val(Lote.Text) = 0 Then
        Lote.Text = "000000"
    End If

    Rem lista.ReportFileName = "Ensayo.rpt"
    Rem lista.GroupSelectionFormula = "{Prueart.Prueba} = " + Chr$(34) + Auxi + Lote.Text + Chr$(34)
    Rem lista.Destination = 1
    Rem lista.Action = 1
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Rem dada
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = "12"
    Printer.Print Tab(1); ""
    Printer.FontSize = "10"
    
    Printer.Print Tab(1); "Empresa : " + WAuxiliar
    Printer.Print Tab(1); ""
    Printer.Print Tab(20); "ENSAYO DE MATERIA PRIMA"
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Prueba"; Tab(15); Lote.Text
    Printer.Print Tab(1); "Producto"; Tab(15); Producto.Text; Tab(40); Descriprod.Caption
    Printer.Print Tab(1); "Fecha"; Tab(15); Fecha.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo1.Caption; Tab(25); Descri1.Caption; Tab(80); Std1.Caption; Tab(105); Valor1.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo2.Caption; Tab(25); descri2.Caption; Tab(80); Std2.Caption; Tab(105); valor2.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo3.Caption; Tab(25); Descri3.Caption; Tab(80); Std3.Caption; Tab(105); Valor3.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo4.Caption; Tab(25); Descri4.Caption; Tab(80); Std4.Caption; Tab(105); valor4.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo5.Caption; Tab(25); Descri5.Caption; Tab(80); Std5.Caption; Tab(105); valor5.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo6.Caption; Tab(25); Descri6.Caption; Tab(80); Std6.Caption; Tab(105); valor6.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo7.Caption; Tab(25); Descri7.Caption; Tab(80); Std7.Caption; Tab(105); valor7.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo8.Caption; Tab(25); Descri8.Caption; Tab(80); Std8.Caption; Tab(105); valor8.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo9.Caption; Tab(25); Descri9.Caption; Tab(80); Std9.Caption; Tab(105); valor9.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); Ensayo10.Caption; Tab(25); Descri10.Caption; Tab(80); Std10.Caption; Tab(105); valor10.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Observaciones"; Tab(20); Ensayo.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Aspecto.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Observaciones.Text
    Printer.Print Tab(1); "Confecciono"; Tab(20); Confecciono.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Liberada"; Tab(30); Pusing("###,###", Val(Liberada.Text))
    Printer.Print Tab(1); "Devuelta"; Tab(30); Pusing("###,###", Val(Devuelta.Text))
    Printer.Print Tab(1); "Nro Rechazo"; Tab(30); Pusing("######", Val(NroRechazo.Text))
    Printer.Print Tab(1); ""
    
    Printer.EndDoc
    

End Sub

Private Sub Modifica_Click()
        XParam = "'" + "1" + Partida.Text + "','" _
                + Ensayo.Text + "','" _
                + Aspecto.Text + "','" _
                + Observaciones.Text + "','" _
                + Confecciono.Text + "'"
        Set rstPrueart = db.OpenRecordset("ModificaPrueartObservaciones " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        Call CmdLimpiar_Click
        Producto.SetFocus
End Sub



Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + orden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            rstOrden.Close
            Llave = "N"
            For WDa% = 1 To 99
                Auxi3 = orden
                Call Ceros(Auxi3, 6)
                Auxi1 = WDa%
                Call Ceros(Auxi1, 2)
                WClave = Auxi3 + Auxi1
                spOrden = "ConsultaOrden " + "'" + WClave + "'"
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    If UCase(Producto.Text) = UCase(rstOrden!Articulo) Then
                        Llave = "S"
                        If rstOrden!Recibida = 0 Then
                            Llave = "X"
                                Else
                            WOrigen = rstOrden!origen
                        End If
                    End If
                    rstOrden.Close
                End If
            Next WDa%
    
            Select Case Llave
                Case "S"
                    Valor1.SetFocus
                Case "N"
                    m$ = "No existe el articulo en la orden de compra especificada"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    orden.SetFocus
                Case "X"
                    m$ = "Orden de compra sin la actualizacion de Informe de receocion"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    orden.SetFocus
                Case Else
            End Select
                Else
            m$ = "Orden de Compra Inexistente"
            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
            orden.SetFocus
        End If
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Listado_Click()
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Aprobado.Value = True
    Rechazo.Value = False
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub
Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub
Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastafec.SetFocus
    End If
End Sub
Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor2.SetFocus
    End If
End Sub
Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor3.SetFocus
    End If
End Sub
Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor4.SetFocus
    End If
End Sub
Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor5.SetFocus
    End If
End Sub
Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor6.SetFocus
    End If
End Sub
Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor7.SetFocus
    End If
End Sub
Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor8.SetFocus
    End If
End Sub
Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor9.SetFocus
    End If
End Sub
Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        valor10.SetFocus
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

Private Sub Lote_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spLaudo = "ListaLaudo " + "'" + Lote.Text + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            m$ = "Numero de lote ya existente"
            A% = MsgBox(m$, 0, "Pruebas de Materias Primas")
            rstLaudo.Close
                Else
            Liberada.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Liberada_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Liberada.Text = Pusing("###,###.##", Liberada.Text)
        Devuelta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Devuelta_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Devuelta.Text = Pusing("###,###.##", Devuelta.Text)
        If Val(Devuelta.Text) = 0 Then
            NroRechazo.Text = ""
            Nueva.Text = "N"
            PartidaProveedor.SetFocus
                Else
            NroRechazo.SetFocus
        End If
    End If
End Sub

Private Sub NroRechazo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spLaudo = "ListaLaudo " + "'" + NroRechazo.Text + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            m$ = "Numero de lote ya existente"
            A% = MsgBox(m$, 0, "Pruebas de Materias Primas")
            rstLaudo.Close
                Else
            Nueva.SetFocus
        End If
    End If
End Sub

Private Sub Nueva_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Verifica_datos(Nueva.Text, "SN", Auxi4)
        If Auxi4 = "S" Then
            PartidaProveedor.SetFocus
        End If
    End If
End Sub

Private Sub PartidaProveedor_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OrigenMercaderia.SetFocus
    End If
End Sub

Private Sub OrigenMercaderia_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Liberada.SetFocus
    End If
End Sub

Private Sub imprime_Click()

    spEspecificaciones = "ConsultaEspecificaciones " + "'" + Producto.Text + "'"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        Producto.Text = rstEspecificaciones!Producto
        Ensayo1.Caption = rstEspecificaciones!Ensayo1
        Ensayo2.Caption = rstEspecificaciones!Ensayo2
        Ensayo3.Caption = rstEspecificaciones!Ensayo3
        Ensayo4.Caption = rstEspecificaciones!Ensayo4
        Ensayo5.Caption = rstEspecificaciones!Ensayo5
        Ensayo6.Caption = rstEspecificaciones!Ensayo6
        Ensayo7.Caption = rstEspecificaciones!Ensayo7
        Ensayo8.Caption = rstEspecificaciones!Ensayo8
        Ensayo9.Caption = rstEspecificaciones!Ensayo9
        Ensayo10.Caption = rstEspecificaciones!Ensayo10
        Std1.Caption = rstEspecificaciones!Valor1
        Std2.Caption = rstEspecificaciones!valor2
        Std3.Caption = rstEspecificaciones!Valor3
        Std4.Caption = rstEspecificaciones!valor4
        Std5.Caption = rstEspecificaciones!valor5
        Std6.Caption = rstEspecificaciones!valor6
        Std7.Caption = rstEspecificaciones!valor7
        Std8.Caption = rstEspecificaciones!valor8
        Std9.Caption = rstEspecificaciones!valor9
        Std10.Caption = rstEspecificaciones!valor10
        
        rstEspecificaciones.Close
                        
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
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            descri2.Caption = ""
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
    End If
        
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            Producto.Text = UCase(Producto.Text)
            spEspecificaciones = "ConsultaEspecificaciones " + "'" + Producto.Text + "'"
            Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificaciones.RecordCount > 0 Then
                rstEspecificaciones.Close
                Call imprime_Click
                    Else
                WProducto = Producto.Text
                CmdLimpiar_Click
                Producto.Text = WProducto
            End If
            spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                Producto.SetFocus
                Exit Sub
            End If
        End If
        Fecha.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Pruebas"
    
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
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
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
            
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 1
            spPrueart = "ListaPruebaConsulta"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
            
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = "Tipo:" + Left$(rstPrueart!Prueba, 1) + " Prueba:" + Mid$(rstPrueart!Prueba, 2, 6) + " Articulo:" + rstPrueart!Producto + " Fecha : " + rstPrueart!Fecha
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstPrueart!Prueba
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Clavepro$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Clavepro$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Producto.Text = rstArticulo!Codigo
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                Call imprime_Click
                    Else
                CmdLimpiar_Click
                Producto.Text = ""
                Descriprod.Caption = ""
            End If
            Producto.SetFocus
            
        Case 1
            Indice = Pantalla.ListIndex
            ClavePrue$ = WIndice.List(Indice)
            spPrueart = "ConsultaPrueart" + "'" + ClavePrue$ + "'"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
                Partida.Text = Mid$(ClavePrue$, 2, 6)
                Producto.Text = rstPrueart!Producto
                Fecha.Text = rstPrueart!Fecha
                orden.Text = rstPrueart!orden
                Rem Ensayo1.Caption = ""
                Valor1.Text = rstPrueart!Valor1
                valor2.Text = rstPrueart!valor2
                Valor3.Text = rstPrueart!Valor3
                valor4.Text = rstPrueart!valor4
                valor5.Text = rstPrueart!valor5
                valor6.Text = rstPrueart!valor6
                valor7.Text = rstPrueart!valor7
                valor8.Text = rstPrueart!valor8
                valor9.Text = rstPrueart!valor9
                valor10.Text = rstPrueart!valor10
                Rem Descriprod.Caption = ""
                Rem Descri1.Caption = ""
                Ensayo.Text = rstPrueart!Ensayo
                Aspecto.Text = rstPrueart!Aspecto
                Observaciones.Text = rstPrueart!Observaciones
                Confecciono.Text = rstPrueart!Confecciono
                Rem Std1.Caption = ""
                Auxi = Left$(rstPrueart!Prueba, 1)
                Lote.Text = Right$(rstPrueart!Prueba, 6)
                Liberada.Text = rstPrueart!Liberada
                Devuelta.Text = rstPrueart!Devuelta
                NroRechazo.Text = rstPrueart!Rechazo
                Nueva.Text = rstPrueart!Nueva
                
                rstPrueart.Close
                
                spEspecificaciones = "ConsultaEspecificaciones " + "'" + Producto.Text + "'"
                Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificaciones.RecordCount > 0 Then
                    rstEspecificaciones.Close
                    Call imprime_Click
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Descriprod.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                    
                    Else
                    
                Call CmdLimpiar_Click
                
            End If
            Producto.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPruart.Caption = "Ingreso de Ensayos de Materia Prima :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
End Sub

Sub Busca_Informe(orden As String, Informe As String, Articulo As String, Pasa As String)

    Informe = ""
    Pasa = "N"
    
    spInforme = "ConsultaInformeOrden " + "'" + orden + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
    
        With rstInforme
            .MoveFirst
            If .NoMatch = False Then
            
                Do
            
                    If Articulo = rstInforme!Articulo Then
                        Informe = rstInforme!Informe
                        Pasa = "S"
                    End If
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                        
                Loop
        
            End If
        
        End With
        rstInforme.Close
    End If

End Sub

Private Sub Cambio_Click()
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub Modif_Cancela_Click()
    Modif.Visible = False
End Sub

Private Sub Modif_Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + Modif_Orden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            rstOrden.Close
            Modif_Solicitado.SetFocus
                Else
            m$ = "Orden de Compra Inexistente"
            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
            Modif_Orden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Modif_Solicitado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Modif_Solicitado.Text <> "" Then
            Modif_Solicitado.Text = UCase(Modif_Solicitado.Text)
            spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
            
                Llave = "N"
                For WDa% = 1 To 40
                    Auxi3 = Modif_Orden.Text
                    Call Ceros(Auxi3, 6)
                    Auxi1 = WDa%
                    Call Ceros(Auxi1, 2)
                    WClave = Auxi3 + Auxi1
                    spOrden = "ConsultaOrden " + "'" + WClave + "'"
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If Modif_Solicitado.Text = rstOrden!Articulo Then
                            Llave = "S"
                        End If
                        rstOrden.Close
                    End If
                Next WDa%
    
                Select Case Llave
                    Case "S"
                        Modif_Recibido.SetFocus
                    Case "N"
                        m$ = "No existe el articulo en la orden de compra especificada"
                        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Modif_Solicitado.SetFocus
                    Case Else
                End Select
                    Else
                m$ = "No existe el articulo especificado"
                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                Modif_Solicitado.SetFocus
            End If
        End If
    End If
End Sub
    
Sub Modif_Recibido_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Modif_Recibido.Text <> "" Then
            Modif_Recibido.Text = UCase(Modif_Recibido.Text)
            spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido, 6) Then
                    m$ = "El articulo recibido no es igual al solicitado"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    Modif_Recibido.SetFocus
                        Else
                    Modif_Orden.SetFocus
                End If
                rstArticulo.Close
                    Else
                If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido, 6) Then
                    m$ = "El articulo recibido no es igual al solicitado"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    Modif_Solicitado.SetFocus
                        Else
                    T$ = "Ingreso de Pruebas"
                    m$ = "No existe el articulo especificado, Desea darlo de alta"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                    
                        spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                        
                            WCodigo = Modif_Recibido.Text
                            WDescripcion = rstArticulo!Descripcion
                            WUnidad = rstArticulo!Unidad
                            WDeposito = rstArticulo!Deposito
                            WInicial = ""
                            WEntradas = ""
                            WSalidas = ""
                            WMinimo = ""
                            WLaboratorio = ""
                            WPedido = ""
                            WEnvase = Str$(rstArticulo!envase)
                            WCosto1 = Str$(rstArticulo!Costo1)
                            WCosto2 = Str$(rstArticulo!Costo2)
                            WRs = rstArticulo!Rs
                            WFlete = Str$(rstArticulo!Flete)
                            WMoneda = rstArticulo!Moneda
                            WControla = Str$(IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla))
                            WDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                            WProveedor = IIf(IsNull(rstArticulo!Proveedor), "", rstArticulo!Proveedor)
                            WDate = Date$
                            WFecha = IIf(IsNull(rstArticulo!Fecha), "", rstArticulo!Fecha)
                            WOrden = IIf(IsNull(rstArticulo!orden), "", Str$(rstArticulo!orden))
                            WDife = ""
                            WCosto3 = ""
                            
                            rstArticulo.Close
                            
                            XParam = "'" + WCodigo + "','" _
                                + WDescripcion + "','" _
                                + WCosto1 + "','" _
                                + WCosto2 + "','" _
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
                                + WFlete + "','" _
                                + WMoneda + "','" _
                                + WControla + "','" _
                                + WDensidad + "','" _
                                + WCosto3 + "'"
                         
                            Set rstArticulo = db.OpenRecordset("AltaArticulo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Modif_Confirma_Click()

    Trabajo = "S"
    
    spOrden = "ListaOrden " + "'" + Modif_Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        rstOrden.Close
            Else
        Trabajo = "N"
        m$ = "Orden de Compra Inexistente"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Trabajo = "N"
        m$ = "No existe el articulo especificado en Articulo Pedido"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Trabajo = "N"
        m$ = "No existe el articulo especificado en Articulo Recibido"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    Llave = "N"
    For WDa% = 1 To 40
        Auxi3 = Modif_Orden.Text
        Call Ceros(Auxi3, 6)
        Auxi1 = WDa%
        Call Ceros(Auxi1, 2)
        WClave = Auxi3 + Auxi1
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            If Modif_Solicitado.Text = rstOrden!Articulo Then
                Llave = "S"
            End If
            rstOrden.Close
        End If
    Next WDa%
    
    Select Case Llave
        Case "N"
            Trabajo = "N"
            m$ = "No existe el articulo en la orden de compra especificada"
            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Case Else
    End Select
    
    If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido.Text, 6) Then
        Trabajo = "N"
        m$ = "El articulo recibido no es igual al solicitado"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    If Trabajo = "S" Then
    
        Modif.Visible = False
        
        For WDa% = 1 To 40
            Auxi3 = Modif_Orden.Text
            Call Ceros(Auxi3, 6)
            Auxi1 = WDa%
            Call Ceros(Auxi1, 2)
            WClave = Auxi3 + Auxi1
            spOrden = "ConsultaOrden " + "'" + WClave + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                If Modif_Solicitado.Text = rstOrden!Articulo Then
                
                    WCantidad = rstOrden!Cantidad
                    WClave = rstOrden!Clave
                    WArticulo = Modif_Recibido.Text
                    WDate = Date$
                    rstOrden.Close
                    XParam = "'" + WClave + "','" _
                                + WArticulo + "','" _
                                + WDate + "'"
                    Set rstOrden = db.OpenRecordset("ModificaOrdenLaboratorio " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                    spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = Modif_Solicitado.Text
                        WPedido = Str$(rstArticulo!Pedido - WCantidad)
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WPedido + "','" _
                                + WDate + "'"
                        spArticulo = "ModificaArticuloOrdenLaboratorio " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                    spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = Modif_Recibido.Text
                        WPedido = Str$(rstArticulo!Pedido + WCantidad)
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WPedido + "','" _
                                + WDate + "'"
                        spArticulo = "ModificaArticuloOrdenLaboratorio " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                        Else
                        
                    rstOrden.Close
                    
                End If
            End If
        Next WDa%
        
        XParam = "'" + Modif_Orden.Text + "','" _
                    + Modif_Solicitado.Text + "'"
        spInforme = "ListaInformeOrdenArticulo " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            WCantidad = rstInforme!Cantidad
            WResta = rstInforme!Resta
            WClave = rstInforme!Clave
            WArticulo = Modif_Recibido.Text
            WDate = Date$
            rstInforme.Close
            XParam = "'" + WClave + "','" _
                        + WArticulo + "','" _
                        + WDate + "'"
            Set rstInforme = db.OpenRecordset("ModificaInformeLaboratorio " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
            spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = Modif_Solicitado.Text
                WPedido = Str$(rstArticulo!Pedido + WResta)
                WLaboratorio = Str$(rstArticulo!Laboratorio - WCantidad)
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WLaboratorio + "','" _
                            + WDate + "'"
                spArticulo = "ModificaArticuloInformeLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = Modif_Recibido.Text
                WPedido = Str$(rstArticulo!Pedido - WResta)
                WLaboratorio = Str$(rstArticulo!Laboratorio + WCantidad)
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WLaboratorio + "','" _
                            + WDate + "'"
                spArticulo = "ModificaArticuloInformeLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
        End If
        
        Producto.Text = Modif_Recibido.Text
        Producto.SetFocus
        
    End If
    
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WClave.Text = "MATIZ" Then
            Pass.Visible = False
            Modif_Orden.Text = ""
            Modif_Solicitado.Text = "  -   -   "
            Modif_Recibido.Text = "  -   -   "
            Modif.Visible = True
            Modif_Orden.SetFocus
        End If
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub Calcula_SaldoOrden()

    WRecibida = 0
    WLaudada = 0

    spInforme = "ListaInformeOrden " + "'" + orden.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If Producto.Text = rstInforme!Articulo Then
                    WRecibida = WRecibida + rstInforme!Cantidad
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    spLaudo = "ListaLaudoOrden " + "'" + orden.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If Producto.Text = rstLaudo!Articulo Then
                    WCantidad1 = rstLaudo!Liberada
                    WCantidad2 = rstLaudo!Devuelta
                    WLaudada = WLaudada + WCantidad1 + WCantidad2
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    SaldoOrden = WRecibida - WLaudada

End Sub
