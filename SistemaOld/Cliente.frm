VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgcliente 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Clientes"
   ClientHeight    =   8190
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   11910
   Begin VB.CommandButton Command211 
      Caption         =   "Command2"
      Height          =   855
      Left            =   11400
      TabIndex        =   202
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox RequiereCertificado 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   200
      Top             =   5760
      Width           =   255
   End
   Begin VB.Frame IngresaIb 
      Height          =   2895
      Left            =   0
      TabIndex        =   101
      Top             =   1320
      Visible         =   0   'False
      Width           =   11775
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
         Left            =   1680
         TabIndex        =   199
         Top             =   1560
         Width           =   4575
      End
      Begin VB.TextBox PorceCm05Tucu 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   174
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox PorceIbCaba 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   167
         Text            =   " "
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox PorceIb 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   115
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox NroIbCiudad 
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
         Left            =   6960
         MaxLength       =   13
         TabIndex        =   113
         Text            =   " "
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox NroIb 
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
         Left            =   6960
         MaxLength       =   13
         TabIndex        =   111
         Text            =   " "
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CierraIb 
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
         Height          =   420
         Left            =   4920
         TabIndex        =   110
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox IbCiudad 
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
         Left            =   1680
         TabIndex        =   109
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ComboBox IbTucu 
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
         Left            =   1680
         TabIndex        =   106
         Top             =   720
         Width           =   4575
      End
      Begin VB.ComboBox Ib 
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
         Left            =   1680
         TabIndex        =   103
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox NroIbTucu 
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
         Left            =   6960
         MaxLength       =   13
         TabIndex        =   102
         Text            =   " "
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "Coeficiente CM05 para  Tucuman    "
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
         Height          =   495
         Left            =   8880
         TabIndex        =   175
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "Alicuota I.B.  "
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
         Left            =   8880
         TabIndex        =   169
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label43 
         Caption         =   "Porcel IB  CABA"
         Height          =   255
         Left            =   1680
         TabIndex        =   168
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Alicuota I.B."
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
         Left            =   8880
         TabIndex        =   116
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "Nro."
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
         Left            =   6360
         TabIndex        =   114
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Nro."
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
         Left            =   6360
         TabIndex        =   112
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "I.B.  C.A.Bs.As."
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
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label24 
         Caption         =   "I.B.  Tucuman"
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
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "I.B. Bs.As"
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
         TabIndex        =   105
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Nro. "
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
         Left            =   6360
         TabIndex        =   104
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton emailenvio 
      Caption         =   "email"
      Height          =   375
      Left            =   9960
      TabIndex        =   197
      Top             =   0
      Width           =   1815
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
      Left            =   9720
      TabIndex        =   196
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Frame PantaCufe 
      Height          =   3135
      Left            =   1680
      TabIndex        =   183
      Top             =   1560
      Visible         =   0   'False
      Width           =   7575
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
         TabIndex        =   190
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Cufe 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   189
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox DirCufe 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   188
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox CufeII 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   187
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox DirCufeII 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   186
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox CufeIII 
         Height          =   285
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   185
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox DirCufeIII 
         Height          =   285
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   184
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label54 
         Caption         =   "CUFE"
         Height          =   255
         Left            =   240
         TabIndex        =   195
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label53 
         Caption         =   "CUFE II"
         Height          =   255
         Left            =   240
         TabIndex        =   194
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label52 
         Caption         =   "CUFE III"
         Height          =   255
         Left            =   240
         TabIndex        =   193
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label51 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   1800
         TabIndex        =   192
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label50 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   3480
         TabIndex        =   191
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CheckBox DolarEspecial 
      Caption         =   "Dolar Especial"
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
      Left            =   9720
      TabIndex        =   180
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CheckBox EtiII 
      Caption         =   "Direccion en Etiqueta"
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
      Left            =   9720
      TabIndex        =   179
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CheckBox EtiI 
      Caption         =   "O/C en Etiqueta"
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
      Left            =   9720
      TabIndex        =   178
      Top             =   5400
      Width           =   2175
   End
   Begin VB.ComboBox Idioma 
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
      Left            =   7800
      TabIndex        =   176
      Text            =   " "
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox emailenv 
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
      TabIndex        =   173
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox CuitII 
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
      MaxLength       =   13
      TabIndex        =   165
      Text            =   " "
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame IngresaEspecif 
      Height          =   855
      Left            =   5160
      TabIndex        =   83
      Top             =   6240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame13 
         Height          =   2415
         Left            =   120
         TabIndex        =   150
         Top             =   4080
         Width           =   7455
         Begin VB.TextBox Especif4 
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
            MaxLength       =   50
            TabIndex        =   156
            Top             =   1560
            Width           =   7095
         End
         Begin VB.TextBox Especif5 
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
            MaxLength       =   50
            TabIndex        =   155
            Top             =   1920
            Width           =   7095
         End
         Begin VB.TextBox Especif1 
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
            MaxLength       =   50
            TabIndex        =   154
            Top             =   480
            Width           =   7095
         End
         Begin VB.TextBox Especif2 
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
            MaxLength       =   50
            TabIndex        =   153
            Top             =   840
            Width           =   7095
         End
         Begin VB.TextBox Especif3 
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
            MaxLength       =   50
            TabIndex        =   152
            Top             =   1200
            Width           =   7095
         End
         Begin VB.Label Label37 
            Caption         =   "Otras"
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
            TabIndex        =   151
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame Frame12 
         Height          =   855
         Left            =   6480
         TabIndex        =   146
         Top             =   1800
         Width           =   5175
         Begin VB.TextBox EtiquetaII 
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
            MaxLength       =   50
            TabIndex        =   148
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox EtiquetaI 
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   147
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label Label36 
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
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1215
         Left            =   6480
         TabIndex        =   140
         Top             =   600
         Width           =   5175
         Begin VB.TextBox EnvasesIII 
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
            MaxLength       =   50
            TabIndex        =   145
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox EnvasesI 
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   142
            Top             =   120
            Width           =   4095
         End
         Begin VB.TextBox EnvasesII 
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
            MaxLength       =   50
            TabIndex        =   141
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label34 
            Caption         =   "Envases"
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
            TabIndex        =   143
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.TextBox CantidadPartidas 
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
         Left            =   10320
         MaxLength       =   6
         TabIndex        =   139
         Text            =   " "
         Top             =   240
         Width           =   495
      End
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   120
         TabIndex        =   133
         Top             =   720
         Width           =   6255
         Begin VB.CheckBox RequiereMsds 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   136
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox EmailMsds 
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
            MaxLength       =   50
            TabIndex        =   134
            Top             =   120
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label35 
            Caption         =   "MSDS 1ra vez"
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
            TabIndex        =   135
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame Frame10 
         Height          =   495
         Left            =   6480
         TabIndex        =   131
         Top             =   120
         Width           =   5175
         Begin VB.CheckBox PartidasVarias 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2040
            TabIndex        =   159
            Top             =   195
            Width           =   255
         End
         Begin VB.Label Label23 
            Caption         =   "Limite"
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
            TabIndex        =   160
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label40 
            Caption         =   "Varias Partidas"
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
            TabIndex        =   132
            Top             =   195
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1335
         Left            =   120
         TabIndex        =   127
         Top             =   2640
         Width           =   6855
         Begin VB.TextBox DiasIII 
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
            MaxLength       =   50
            TabIndex        =   144
            Top             =   840
            Width           =   6615
         End
         Begin VB.TextBox DiasII 
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
            MaxLength       =   50
            TabIndex        =   130
            Top             =   480
            Width           =   6615
         End
         Begin VB.TextBox DiasI 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   129
            Top             =   120
            Width           =   5055
         End
         Begin VB.Label Label39 
            Caption         =   "Dias y Horario"
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
            TabIndex        =   128
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   495
         Left            =   120
         TabIndex        =   124
         Top             =   2040
         Width           =   6255
         Begin VB.CheckBox PermiteParcial 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   158
            Top             =   150
            Width           =   255
         End
         Begin VB.Label Label33 
            Caption         =   "Entrega Parcial"
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
            TabIndex        =   125
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Height          =   495
         Left            =   120
         TabIndex        =   122
         Top             =   1560
         Width           =   6255
         Begin VB.CheckBox RequiereHoja 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   138
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox EmailHoja 
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
            MaxLength       =   50
            TabIndex        =   126
            Top             =   120
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label32 
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
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   195
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   120
         TabIndex        =   120
         Top             =   1080
         Width           =   6255
         Begin VB.CheckBox RequiereMsdsCada 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   137
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label31 
            Caption         =   "MSDS cada entrega"
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
            TabIndex        =   121
            Top             =   195
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   118
         Top             =   120
         Width           =   6255
         Begin VB.TextBox EmailCertificado 
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
            MaxLength       =   50
            TabIndex        =   157
            Top             =   120
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label30 
            Caption         =   "Certificado de Analisis"
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
            TabIndex        =   119
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.CommandButton FinEspecificaciones 
         Caption         =   "Fin de Ingreso"
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
         Left            =   4440
         TabIndex        =   89
         Top             =   6720
         Width           =   2175
      End
      Begin VB.TextBox Especif10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         MaxLength       =   50
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Especif9 
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
         MaxLength       =   50
         TabIndex        =   87
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Especif8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         MaxLength       =   50
         TabIndex        =   86
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Especif7 
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
         MaxLength       =   50
         TabIndex        =   85
         Top             =   -120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Especif6 
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
         MaxLength       =   50
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.TextBox NroSedronar 
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
      MaxLength       =   15
      TabIndex        =   161
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton BotonIb 
      Caption         =   "Ingresos Brutos"
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
      Left            =   10200
      TabIndex        =   117
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   2880
      TabIndex        =   78
      Top             =   2760
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   80
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   79
         Top             =   1440
         Width           =   2415
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
         TabIndex        =   81
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   8160
      TabIndex        =   100
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox ImpreVto 
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
      Left            =   1920
      TabIndex        =   98
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton BotonDirEntrega 
      Caption         =   "Direciones  de  Entrega"
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
      Left            =   10200
      TabIndex        =   97
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame PantaDirEntrega 
      Height          =   1575
      Left            =   5280
      TabIndex        =   90
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox DirEntregaV 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   96
         Text            =   " "
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox DirEntregaIV 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   95
         Text            =   " "
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox DirEntregaIII 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   94
         Text            =   " "
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox DirEntregaII 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   93
         Text            =   " "
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox DirEntrega 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   91
         Text            =   " "
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label16 
         Caption         =   "Direccion de Entrega"
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
         TabIndex        =   92
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1815
      Left            =   360
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Hasta 
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
         MaxLength       =   6
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         MaxLength       =   6
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   855
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   1200
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
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
         Left            =   2880
         TabIndex        =   21
         Top             =   360
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
         Left            =   2880
         TabIndex        =   20
         Top             =   960
         Width           =   1095
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
         Left            =   120
         TabIndex        =   19
         Top             =   600
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   2175
      Left            =   4800
      TabIndex        =   72
      Top             =   5880
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Cliente.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Especificaciones 
      Caption         =   "Especificaciones  de Entrega"
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
      Left            =   10200
      TabIndex        =   82
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Precio 
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
      Left            =   10320
      MaxLength       =   2
      TabIndex        =   77
      Top             =   1800
      Width           =   975
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
      TabIndex        =   75
      Top             =   6120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Block2 
      Caption         =   "   Cerrar       Block de      Notas"
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
      Left            =   10200
      TabIndex        =   74
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Block1 
      Caption         =   "Ver Block de Notas"
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
      Left            =   10200
      TabIndex        =   73
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox MInimo 
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
      TabIndex        =   69
      Text            =   " "
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Limite 
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
      TabIndex        =   68
      Text            =   " "
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox pago2 
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
      MaxLength       =   4
      TabIndex        =   65
      Text            =   " "
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Pago1 
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
      MaxLength       =   4
      TabIndex        =   64
      Text            =   " "
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Horario 
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
      MaxLength       =   20
      TabIndex        =   63
      Text            =   " "
      Top             =   3600
      Width           =   2535
   End
   Begin VB.ComboBox Provincia 
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
      TabIndex        =   59
      Text            =   " "
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox fax 
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
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   57
      Text            =   " "
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox email 
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
      MaxLength       =   80
      TabIndex        =   56
      Text            =   " "
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Rubro 
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
      MaxLength       =   4
      TabIndex        =   51
      Text            =   " "
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
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
      Left            =   8640
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      Begin VB.CommandButton Anterior 
         Caption         =   "Anterior"
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Siguiente"
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
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Vendedor 
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
      MaxLength       =   4
      TabIndex        =   43
      Text            =   " "
      Top             =   2160
      Width           =   855
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   42
      Text            =   " "
      Top             =   2880
      Width           =   7575
   End
   Begin VB.TextBox Contacto 
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
      TabIndex        =   38
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condicion de Iva"
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
      Height          =   1215
      Left            =   5640
      TabIndex        =   36
      Top             =   480
      Width           =   4095
      Begin VB.OptionButton Iva6 
         Caption         =   "No catalogado"
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
         Left            =   2160
         TabIndex        =   53
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Iva5 
         Caption         =   "Monotributo"
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
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Iva4 
         Caption         =   "Exento"
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
         Left            =   2160
         TabIndex        =   48
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Iva3 
         Caption         =   "Cons. Final"
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
         Left            =   2160
         TabIndex        =   47
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Iva2 
         Caption         =   "No Inscripto"
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
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Iva1 
         Caption         =   "Inscripto"
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
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Cuit 
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
      Left            =   4080
      MaxLength       =   13
      TabIndex        =   35
      Text            =   " "
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Telefono 
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
      MaxLength       =   40
      TabIndex        =   34
      Text            =   " "
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Postal 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   33
      Text            =   " "
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Localidad 
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
      TabIndex        =   32
      Text            =   " "
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Direccion 
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
      TabIndex        =   27
      Text            =   " "
      Top             =   720
      Width           =   3135
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
      Top             =   0
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wcliente.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Clientes"
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
      Left            =   7320
      TabIndex        =   16
      Top             =   0
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
      Left            =   7560
      TabIndex        =   14
      Top             =   3600
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
      Left            =   6480
      TabIndex        =   13
      Top             =   4320
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
      Left            =   7560
      TabIndex        =   1
      Top             =   4320
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
      Left            =   7560
      TabIndex        =   7
      Top             =   3960
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
      Left            =   6480
      TabIndex        =   6
      Top             =   3600
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
      Left            =   6480
      TabIndex        =   5
      Top             =   3960
      Width           =   975
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
      Left            =   5640
      MaxLength       =   50
      TabIndex        =   4
      Top             =   0
      Width           =   4215
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
      Height          =   1620
      ItemData        =   "Cliente.frx":007A
      Left            =   120
      List            =   "Cliente.frx":0081
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   2160
      TabIndex        =   39
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Pais 
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
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   163
      Text            =   " "
      Top             =   1440
      Width           =   615
   End
   Begin MSMask.MaskEdBox fechsedro 
      Height          =   285
      Left            =   10320
      TabIndex        =   171
      Top             =   5040
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
      Left            =   7800
      TabIndex        =   198
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox EmailFactura 
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
      Left            =   4080
      TabIndex        =   181
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label55 
      Caption         =   "Email Certificado"
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
      Left            =   2400
      TabIndex        =   201
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label49 
      Caption         =   "Requiere Certificado"
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
      TabIndex        =   182
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label48 
      Caption         =   "Idioma Etiquetas/Certif."
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
      Left            =   5160
      TabIndex        =   177
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label46 
      Caption         =   "EMail Envases"
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
      TabIndex        =   172
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label45 
      Caption         =   "F.Venc  SEDRONAR"
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
      Left            =   8160
      TabIndex        =   170
      Top             =   5080
      Width           =   2535
   End
   Begin VB.Label Label42 
      Caption         =   "Nro Impositivo"
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
      Left            =   5760
      TabIndex        =   166
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label41 
      Caption         =   "Pais"
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
      TabIndex        =   164
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label38 
      Caption         =   "SEDRONAR Nro. Insc. "
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
      TabIndex        =   162
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label25 
      Caption         =   "Fecha Vencimiento"
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
      TabIndex        =   99
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Mod. Precio"
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
      Left            =   9000
      TabIndex        =   76
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Despago2 
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
      TabIndex        =   71
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Despago1 
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
      TabIndex        =   70
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label15 
      Caption         =   "Minimo a Facturar"
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
      TabIndex        =   67
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Limite de Credito"
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
      TabIndex        =   66
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Condicion de Proyeccion"
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
      TabIndex        =   62
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Condicion de Pago"
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
      TabIndex        =   61
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Horario"
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
      TabIndex        =   60
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Provincia"
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
      TabIndex        =   58
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Fax"
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
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail"
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
      TabIndex        =   54
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label DesRubro 
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
      Left            =   3240
      TabIndex        =   52
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label17 
      Caption         =   "Rubro"
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
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label DesVendedor 
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
      TabIndex        =   44
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Vendedor"
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
      TabIndex        =   41
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
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
      TabIndex        =   40
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Provi 
      Caption         =   "Contacto"
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
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Cuit"
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
      Left            =   3360
      TabIndex        =   31
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono"
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
      TabIndex        =   30
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "C.Postal"
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
      TabIndex        =   29
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Poblaci 
      Caption         =   "Localidad"
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
      TabIndex        =   28
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
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
      TabIndex        =   26
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Razon Social"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Cliente"
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
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "prgcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstClienteEspecif As Recordset
Dim spClienteEspecif As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstRubro As Recordset
Dim spRubro As String
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String
Private WIva As String
Private WProvincia As String
Private WIb As String
Private WIbTucu As String
Private WIbCiudad As String
Private WIbCiudadII As String
Private WImpreVto As String
Private WIdioma As String
Private WImporte1 As String
Private WImporte2 As String
Private WImporte3 As String
Private WImporte4 As String
Private WImporte5 As String
Private WImporte6 As String
Private WDate As String
Private WGraba As String
Private WPorceI As String
Private WPorceII As String

Dim ZEtiI As Integer
Dim ZEtiIi As Integer
Dim ZDolarEspecial As Integer

Dim ZZPorceIbCaba As Double
Dim ZZRestriccion As Integer
Dim WRestriccion As String

Dim CargaEmpresa(12, 2) As String

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

Private Sub Command211_Click()

    
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
    ruta = "C:\david\Rubrosfarma.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
    
            
            
            ZZCodigo = appExcel.cells(LugarPlanilla, 1).Value
            ZZRubro = appExcel.cells(LugarPlanilla, 4).Value
                    
            If Trim(ZZCodigo) = "" Then Exit Do
            
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Cliente SET "
                    ZSql = ZSql & "Rubro = " + "'" + Str$(ZZRubro) + "'"
                    ZSql = ZSql & " Where Cliente = " + "'" + Trim(UCase(ZZCodigo)) + "'"
                    spCliente = ZSql
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    
                            
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Estadistica SET "
                    ZSql = ZSql & "Rubro = " + "'" + Str$(ZZRubro) + "'"
                    ZSql = ZSql & " Where Cliente = " + "'" + Trim(UCase(ZZCodigo)) + "'"
                    ZSql = ZSql & " and OrdFecha >= " + "'" + "20160101" + "'"
                    spEstadistica = ZSql
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                    
            
                End If
                
            Next Cicla
            
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    
Stop


End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Sub Imprime_Descripcion()

    Rem lee rubro

    WRubro = Rubro.Text
    spRubro = "ConsultaRubro " + "'" + Rubro.Text + "'"
    Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
    If rstRubro.RecordCount > 0 Then
        DesRubro.Caption = rstRubro!Nombre
        rstRubro.Close
            Else
        DesRubro.Caption = ""
    End If
    
    Rem lee vendedor
    WVendedor = vendedor.Text
    spVendedor = "ConsultaVendedor " + "'" + WVendedor + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        DesVendedor.Caption = rstVendedor!Nombre
        rstVendedor.Close
            Else
        DesVendedor.Caption = ""
    End If

    
    Rem lee condicion de pago 1
    
    WPago1 = Pago1.Text
    spPago = "ConsultaPago " + "'" + Pago1.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        Despago1.Caption = rstPago!Nombre
        rstPago.Close
            Else
        Despago1.Caption = ""
    End If

    Rem lee condicion de pago 2
    
    WPago2 = Pago2.Text
    spPago = "ConsultaPago " + "'" + Pago2.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        Despago2.Caption = rstPago!Nombre
        rstPago.Close
            Else
        Despago2.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(vendedor.Text) = 0 Then
        vendedor.Text = "0"
    End If
    If Val(Rubro.Text) = 0 Then
        Rubro.Text = "0"
    End If
    If Val(Pago1.Text) = 0 Then
        Pago1.Text = "0"
    End If
    If Val(Pago2.Text) = 0 Then
        Pago2.Text = "0"
    End If
    If Val(Limite.Text) = 0 Then
        Limite.Text = "0"
    End If
    If Val(MInimo.Text) = 0 Then
        MInimo.Text = "0"
    End If
    If Val(PorceIb.Text) = 0 Then
        PorceIb.Text = "0"
    End If
    If Val(PorceIbCaba.Text) = 0 Then
        PorceIbCaba.Text = "0"
    End If
End Sub

Sub Format_datos()
    Limite.Text = Pusing("###,###.##", Limite.Text)
    MInimo.Text = Pusing("###,###.##", MInimo.Text)
    PorceIb.Text = Pusing("###,###.##", PorceIb.Text)
    PorceIbCaba.Text = Pusing("###,###.##", PorceIbCaba.Text)
End Sub

Sub Imprime_Datos()


   Rem BY NAN 18-10-2012
    
    WCliente = Cliente.Text
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = IIf(IsNull(rstCliente!Cliente), "", rstCliente!Cliente)
        Razon.Text = IIf(IsNull(rstCliente!Razon), "", rstCliente!Razon)
        Direccion.Text = IIf(IsNull(rstCliente!Direccion), "", rstCliente!Direccion)
        Localidad.Text = IIf(IsNull(rstCliente!Localidad), "", rstCliente!Localidad)
        Postal.Text = IIf(IsNull(rstCliente!Postal), "", rstCliente!Postal)
        Telefono.Text = IIf(IsNull(rstCliente!Telefono), "", rstCliente!Telefono)
        Contacto.Text = IIf(IsNull(rstCliente!Contacto), "", rstCliente!Contacto)
        Observaciones.Text = IIf(IsNull(rstCliente!Observaciones), "", rstCliente!Observaciones)
        Cuit.Text = IIf(IsNull(rstCliente!Cuit), "", rstCliente!Cuit)
        vendedor.Text = rstCliente!vendedor
        email.Text = rstCliente!email
        fax.Text = rstCliente!fax
        Rubro.Text = rstCliente!Rubro
        Horario.Text = rstCliente!Horario
        NroSedronar.Text = Trim(IIf(IsNull(rstCliente!NroSedronar), "", rstCliente!NroSedronar))
        Pago1.Text = rstCliente!Pago1
        Pago2.Text = rstCliente!Pago2
        Limite.Text = rstCliente!Limite
        MInimo.Text = rstCliente!MInimo
        DirEntrega.Text = rstCliente!DirEntrega
        Precio.Text = IIf(IsNull(rstCliente!Precio), "", rstCliente!Precio)
        fechsedro = IIf(IsNull(rstCliente!fechsedro), "  /  /    ", rstCliente!fechsedro)
        Rem Especif1.Text = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
        Rem Especif2.Text = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
        Rem Especif3.Text = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
        Rem Especif4.Text = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
        Rem Especif5.Text = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
        Rem Especif6.Text = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
        Rem Especif7.Text = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
        Rem Especif8.Text = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
        Rem Especif9.Text = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
        Rem Especif10.Text = IIf(IsNull(rstCliente!Especif10), "", rstCliente!Especif10)
        Rem Especif1.Text = RTrim(Especif1.Text)
        Rem Especif2.Text = RTrim(Especif2.Text)
        Rem Especif3.Text = RTrim(Especif3.Text)
        Rem Especif4.Text = RTrim(Especif4.Text)
        Rem Especif5.Text = RTrim(Especif5.Text)
        Rem Especif6.Text = RTrim(Especif6.Text)
        Rem Especif7.Text = RTrim(Especif7.Text)
        Rem Especif8.Text = RTrim(Especif8.Text)
        Rem Especif9.Text = RTrim(Especif9.Text)
        Rem Especif10.Text = RTrim(Especif10.Text)
        emailenv.Text = Trim(IIf(IsNull(rstCliente!emailenv), "", rstCliente!emailenv))
        EmailFactura.Text = Trim(IIf(IsNull(rstCliente!EmailFactura), "", rstCliente!EmailFactura))
        DirEntregaII.Text = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        DirEntregaIII.Text = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        DirEntregaIV.Text = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        DirEntregaV.Text = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        Iva1.Value = False
        Iva2.Value = False
        Iva3.Value = False
        Iva4.Value = False
        Iva5.Value = False
        Iva6.Value = False
        Provincia.ListIndex = rstCliente!Provincia
        
        Ib.ListIndex = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
        IbTucu.ListIndex = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
        IbCiudad.ListIndex = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
        IbCiudadII.ListIndex = IIf(IsNull(rstCliente!IbCiudadII), "0", rstCliente!IbCiudadII)
        
        ZZPorceIbCaba = IIf(IsNull(rstCliente!PorceIbCaba), "0", rstCliente!PorceIbCaba)
        PorceIbCaba.Text = Str$(ZZPorceIbCaba)
        
        PorceIb.Text = IIf(IsNull(rstCliente!PorceIb), "0", rstCliente!PorceIb)
        
        NroIb.Text = IIf(IsNull(rstCliente!NroIb), "", rstCliente!NroIb)
        NroIbTucu.Text = IIf(IsNull(rstCliente!NroIbTucu), "", rstCliente!NroIbTucu)
        ZZPorceCm05Tucu = IIf(IsNull(rstCliente!PorceCm05Tucu), "0", rstCliente!PorceCm05Tucu)
        PorceCm05Tucu.Text = Str$(ZZPorceCm05Tucu)
        NroIbCiudad.Text = IIf(IsNull(rstCliente!NroIbCiudad), "", rstCliente!NroIbCiudad)
        
        NroIbCiudad.Text = Trim(NroIbCiudad.Text)
        NroIbTucu.Text = Trim(NroIbTucu.Text)
        
        ImpreVto.ListIndex = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
        Idioma.ListIndex = IIf(IsNull(rstCliente!Idioma), "0", rstCliente!Idioma)
        
        Select Case Val(rstCliente!Iva)
            Case 1
                Iva1.Value = True
            Case 2
                Iva2.Value = True
            Case 3
                Iva3.Value = True
            Case 4
                Iva4.Value = True
            Case 5
                Iva5.Value = True
            Case 6
                Iva6.Value = True
            Case Else
        End Select
        
        Pais.Text = Trim(IIf(IsNull(rstCliente!Pais), "", rstCliente!Pais))
        CuitII.Text = Trim(IIf(IsNull(rstCliente!CuitII), "", rstCliente!CuitII))
        
        ZEtiI = Trim(IIf(IsNull(rstCliente!EtiI), "0", rstCliente!EtiI))
        ZEtiIi = Trim(IIf(IsNull(rstCliente!EtiII), "0", rstCliente!EtiII))
        ZDolarEspecial = Trim(IIf(IsNull(rstCliente!DolarEspecial), "0", rstCliente!DolarEspecial))
        
        EtiI.Value = ZEtiI
        EtiII.Value = ZEtiIi
        DolarEspecial.Value = ZDolarEspecial
        
        Cufe.Text = IIf(IsNull(rstCliente!Cufe), "", rstCliente!Cufe)
        CufeII.Text = IIf(IsNull(rstCliente!CufeII), "", rstCliente!CufeII)
        CufeIII.Text = IIf(IsNull(rstCliente!CufeIII), "", rstCliente!CufeIII)
        DirCufe.Text = IIf(IsNull(rstCliente!DirCufe), "", rstCliente!DirCufe)
        DirCufeII.Text = IIf(IsNull(rstCliente!DirCufeII), "", rstCliente!DirCufeII)
        DirCufeIII.Text = IIf(IsNull(rstCliente!DirCufeIII), "", rstCliente!DirCufeIII)
        
        ZZRestriccion = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
        Restriccion.Value = ZZRestriccion
        
        rstCliente.Close
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteEspecif"
        ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + Cliente.Text + "'"
        spClienteEspecif = ZSql
        Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteEspecif.RecordCount > 0 Then
        
            RequiereCertificado.Value = IIf(IsNull(rstClienteEspecif!RequiereCertificado), "0", rstClienteEspecif!RequiereCertificado)
            RequiereMsds.Value = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
            RequiereMsdsCada.Value = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
            RequiereHoja.Value = IIf(IsNull(rstClienteEspecif!RequiereHoja), "0", rstClienteEspecif!RequiereHoja)
            PermiteParcial.Value = IIf(IsNull(rstClienteEspecif!PermiteParcial), "0", rstClienteEspecif!PermiteParcial)
            PartidasVarias.Value = IIf(IsNull(rstClienteEspecif!PartidaVarias), "0", rstClienteEspecif!PartidaVarias)
            CantidadPartidas.Text = IIf(IsNull(rstClienteEspecif!CantidadPartidas), "", rstClienteEspecif!CantidadPartidas)
            
            EmailCertificado.Text = IIf(IsNull(rstClienteEspecif!EmailCertificado), "", rstClienteEspecif!EmailCertificado)
            EmailMsds.Text = IIf(IsNull(rstClienteEspecif!EmailMsds), "", rstClienteEspecif!EmailMsds)
            EmailHoja.Text = IIf(IsNull(rstClienteEspecif!EmailHoja), "", rstClienteEspecif!EmailHoja)
            DiasI.Text = IIf(IsNull(rstClienteEspecif!DiasI), "", rstClienteEspecif!DiasI)
            DiasII.Text = IIf(IsNull(rstClienteEspecif!DiasII), "", rstClienteEspecif!DiasII)
            DiasIII.Text = IIf(IsNull(rstClienteEspecif!DiasIII), "", rstClienteEspecif!DiasIII)
            EnvasesI.Text = IIf(IsNull(rstClienteEspecif!EnvasesI), "", rstClienteEspecif!EnvasesI)
            EnvasesII.Text = IIf(IsNull(rstClienteEspecif!EnvasesII), "", rstClienteEspecif!EnvasesII)
            EnvasesIII.Text = IIf(IsNull(rstClienteEspecif!EnvasesIII), "", rstClienteEspecif!EnvasesIII)
            EtiquetaI.Text = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
            EtiquetaII.Text = IIf(IsNull(rstClienteEspecif!EtiquetaII), "", rstClienteEspecif!EtiquetaII)
            Especif1.Text = IIf(IsNull(rstClienteEspecif!Especif1), "", rstClienteEspecif!Especif1)
            Especif2.Text = IIf(IsNull(rstClienteEspecif!Especif2), "", rstClienteEspecif!Especif2)
            Especif3.Text = IIf(IsNull(rstClienteEspecif!Especif3), "", rstClienteEspecif!Especif3)
            Especif4.Text = IIf(IsNull(rstClienteEspecif!Especif4), "", rstClienteEspecif!Especif4)
            Especif5.Text = IIf(IsNull(rstClienteEspecif!Especif5), "", rstClienteEspecif!Especif5)
            
            rstClienteEspecif.Close
            
        End If
        
        Call Format_datos
        Call Imprime_Descripcion
    End If

End Sub

Private Sub Acepta_Click()
    
    Listado.WindowTitle = "Listado de Clientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Cliente.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Cliente.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.Direccion, Cliente.Localidad, Cliente.Provincia, Cliente.Postal " _
                        + "From " _
                        + DSQ + ".dbo.Cliente Cliente " _
                        + "Where " _
                        + "Cliente.Cliente >= '" + Desde.Text + "' AND " _
                        + "Cliente.Cliente <= '" + Hasta.Text + "'"
    
    Listado.DataFiles(1) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
      
    Listado.Action = 1
    Frame2.Visible = False
End Sub



Private Sub AltaCufe_Click()
    PantaCufe.Visible = True
    Cufe.SetFocus
End Sub

Private Sub Block1_Click()

    On Error GoTo WError

    Agenda.LoadFile "blanco.rtf", 0
    Agenda.LoadFile Cliente.Text + ".rtf", 0
    Agenda.Visible = True
    Block1.Visible = False
    Block2.Visible = True
    Agenda.Height = 7215
    Agenda.Left = 840
    Agenda.Top = 720
    Agenda.Width = 9375
    Agenda.SetFocus
    
WError:
    Resume Next
    
End Sub

Private Sub Block2_Click()
    Agenda.SaveFile Cliente.Text + ".rtf", 0
    Agenda.Visible = False
    Block1.Visible = True
    Block2.Visible = False
    Agenda.Height = 2175
    Agenda.Left = 5280
    Agenda.Top = 5760
    Agenda.Width = 4695
End Sub

Private Sub BotonDirEntrega_Click()
    PantaDirEntrega.Height = 2295
    PantaDirEntrega.Left = 1320
    PantaDirEntrega.Top = 2640
    PantaDirEntrega.Width = 8535
    PantaDirEntrega.Visible = True
    DirEntrega.SetFocus
End Sub

Private Sub BotonIb_Click()
    IngresaIb.Height = 2535
    IngresaIb.Left = 120
    IngresaIb.Top = 1080
    IngresaIb.Width = 11655
    IngresaIb.Visible = True
    Ib.SetFocus
End Sub

Private Sub CierraIb_Click()
    Razon.SetFocus
    IngresaIb.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub CierraPantaCufe_Click()
    PantaCufe.Visible = False
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo WError
    
    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else

        Cliente.Text = UCase(Cliente.Text)
    
        If Cliente.Text <> "" Then

            Call Verifica_datos
            XPasa = "S"
    
            If Val(Rubro.Text) <> 0 Then
                WRubro = Rubro.Text
                spRubro = "ConsultaRubro " + "'" + Rubro.Text + "'"
                Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
                If rstRubro.RecordCount <= 0 Then
                    XPasa = "N"
                    m$ = "Codigo de Rubro Incorrecto"
                    a% = MsgBox(m$, 0, "Archivo de Clientes")
                        Else
                    rstRubro.Close
                End If
            End If
    
            If Val(vendedor.Text) <> 0 Then
                WVendedor = vendedor.Text
                spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
                Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstVendedor.RecordCount <= 0 Then
                    XPasa = "N"
                    m$ = "Codigo de Vendedor Incorrecto"
                    a% = MsgBox(m$, 0, "Archivo de Clientes")
                        Else
                    rstVendedor.Close
                End If
            End If
        
            If Provincia.ListIndex > 24 Then
                XPasa = "N"
                m$ = "Codigo de Provincia Incorrecto"
                a% = MsgBox(m$, 0, "Archivo de Clientes")
            End If
        
            If XPasa = "S" Then
        
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    rstCliente.Close
                    WPasa = "S"
                        Else
                    WPasa = "N"
                End If
        
                Call Verifica_datos
                If Iva1.Value = True Then
                    WIva = "1"
                End If
                If Iva2.Value = True Then
                    WIva = "2"
                End If
                If Iva3.Value = True Then
                    WIva = "3"
                End If
                If Iva4.Value = True Then
                    WIva = "4"
                End If
                If Iva5.Value = True Then
                    WIva = "5"
                End If
                If Iva6.Value = True Then
                    WIva = "6"
                End If
            
                WImporte1 = 0
                WImporte2 = 0
                WImporte3 = 0
                WImporte4 = 0
                WImporte5 = 0
                WImporte6 = 0
                WDate = Date$
                WProvincia = Provincia.ListIndex
                
                WIb = Ib.ListIndex
                WIbTucu = IbTucu.ListIndex
                WIbCiudad = IbCiudad.ListIndex
                WIbCiudadII = IbCiudadII.ListIndex
                
                WImpreVto = ImpreVto.ListIndex
                If WImpreVto = -1 Then
                    WImpreVto = 0
                End If
                
                WIdioma = Idioma.ListIndex
                If WIdioma = -1 Then
                    WIdioma = 0
                End If
                
                If EtiI.Value = 1 Then
                    WEtiI = "1"
                        Else
                    WEtiI = "0"
                End If
                
                If EtiII.Value = 1 Then
                    WEtiII = "1"
                        Else
                    WEtiII = "0"
                End If
                    
                If DolarEspecial.Value = 1 Then
                    WDolarEspecial = "1"
                        Else
                    WDolarEspecial = "0"
                End If
                
                
                
                
                WRubro = Val(Rubro.Text)
                WVendedor = Val(vendedor.Text)
                WPago1 = Val(Pago1.Text)
                WPago2 = Val(Pago2.Text)
                WLImite = Val(Limite.Text)
                WMinimo = Val(MInimo.Text)
                WAdicional = ""
                WRestriccion = Restriccion.Value
        
                If WPasa = "N" Then
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Cliente ("
                    ZSql = ZSql + "Cliente ,"
                    ZSql = ZSql + "Razon ,"
                    ZSql = ZSql + "Direccion ,"
                    ZSql = ZSql + "Localidad ,"
                    ZSql = ZSql + "Pais ,"
                    ZSql = ZSql + "CuitII ,"
                    ZSql = ZSql + "Provincia ,"
                    ZSql = ZSql + "Postal ,"
                    ZSql = ZSql + "Email ,"
                    ZSql = ZSql + "Fax ,"
                    ZSql = ZSql + "Telefono ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "Contacto ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Vendedor ,"
                    ZSql = ZSql + "Iva ,"
                    ZSql = ZSql + "Rubro ,"
                    ZSql = ZSql + "Horario ,"
                    ZSql = ZSql + "NroSedronar ,"
                    ZSql = ZSql + "Pago1 ,"
                    ZSql = ZSql + "Pago2 ,"
                    ZSql = ZSql + "Limite ,"
                    ZSql = ZSql + "Minimo ,"
                    ZSql = ZSql + "DirEntrega ,"
                    ZSql = ZSql + "Importe1 ,"
                    ZSql = ZSql + "Importe2 ,"
                    ZSql = ZSql + "Importe3 ,"
                    ZSql = ZSql + "Importe4 ,"
                    ZSql = ZSql + "Importe5 ,"
                    ZSql = ZSql + "Importe6 ,"
                    ZSql = ZSql + "WDate ,"
                    ZSql = ZSql + "Precio ,"
                    ZSql = ZSql + "NroIb ,"
                    ZSql = ZSql + "NroIbTucu ,"
                    ZSql = ZSql + "PorceCm05Tucu ,"
                    ZSql = ZSql + "NroIbCiudad ,"
                    ZSql = ZSql + "Ib ,"
                    ZSql = ZSql + "IbTucu ,"
                    ZSql = ZSql + "IbCiudad ,"
                    ZSql = ZSql + "IbCiudadII ,"
                    ZSql = ZSql + "PorceIb ,"
                    ZSql = ZSql + "PorceIbCaba ,"
                    ZSql = ZSql + "ImpreVto ,"
                    ZSql = ZSql + "Idioma ,"
                    ZSql = ZSql + "EtiI ,"
                    ZSql = ZSql + "EtiII ,"
                    ZSql = ZSql + "DolarEspecial ,"
                    ZSql = ZSql + "Especif1 ,"
                    ZSql = ZSql + "Especif2 ,"
                    ZSql = ZSql + "Especif3 ,"
                    ZSql = ZSql + "Especif4 ,"
                    ZSql = ZSql + "Especif5 ,"
                    ZSql = ZSql + "Especif6 ,"
                    ZSql = ZSql + "Especif7 ,"
                    ZSql = ZSql + "Especif8 ,"
                    ZSql = ZSql + "Especif9 ,"
                    ZSql = ZSql + "Especif10 ,"
                    ZSql = ZSql + "DirEntregaII , "
                    ZSql = ZSql + "DirEntregaIII ,"
                    ZSql = ZSql + "DirEntregaIV ,"
                    ZSql = ZSql + "fechsedro ,"
                    ZSql = ZSql + "Restriccion ,"
                    ZSql = ZSql + "Emailenv ,"
                    ZSql = ZSql + "EmailFactura ,"
                    ZSql = ZSql + "DirEntregaV )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + Cliente.Text + "',"
                    ZSql = ZSql + "'" + Razon.Text + "',"
                    ZSql = ZSql + "'" + Direccion.Text + "',"
                    ZSql = ZSql + "'" + Localidad.Text + "',"
                    ZSql = ZSql + "'" + Pais.Text + "',"
                    ZSql = ZSql + "'" + CuitII.Text + "',"
                    ZSql = ZSql + "'" + WProvincia + "',"
                    ZSql = ZSql + "'" + Postal.Text + "',"
                    ZSql = ZSql + "'" + email.Text + "',"
                    ZSql = ZSql + "'" + fax.Text + "',"
                    ZSql = ZSql + "'" + Telefono.Text + "',"
                    ZSql = ZSql + "'" + Cuit.Text + "',"
                    ZSql = ZSql + "'" + Contacto.Text + "',"
                    ZSql = ZSql + "'" + Observaciones.Text + "',"
                    ZSql = ZSql + "'" + vendedor.Text + "',"
                    ZSql = ZSql + "'" + WIva + "',"
                    ZSql = ZSql + "'" + Rubro.Text + "',"
                    ZSql = ZSql + "'" + Horario.Text + "',"
                    ZSql = ZSql + "'" + NroSedronar.Text + "',"
                    ZSql = ZSql + "'" + Pago1.Text + "',"
                    ZSql = ZSql + "'" + Pago2.Text + "',"
                    ZSql = ZSql + "'" + Limite.Text + "',"
                    ZSql = ZSql + "'" + MInimo.Text + "',"
                    ZSql = ZSql + "'" + DirEntrega.Text + "',"
                    ZSql = ZSql + "'" + WImporte1 + "',"
                    ZSql = ZSql + "'" + WImporte2 + "',"
                    ZSql = ZSql + "'" + WImporte3 + "',"
                    ZSql = ZSql + "'" + WImporte4 + "',"
                    ZSql = ZSql + "'" + WImporte5 + "',"
                    ZSql = ZSql + "'" + WImporte6 + "',"
                    ZSql = ZSql + "'" + WDate + "',"
                    ZSql = ZSql + "'" + Precio.Text + "',"
                    ZSql = ZSql + "'" + NroIb.Text + "',"
                    ZSql = ZSql + "'" + NroIbTucu.Text + "',"
                    ZSql = ZSql + "'" + PorceCm05Tucu.Text + "',"
                    ZSql = ZSql + "'" + NroIbCiudad.Text + "',"
                    ZSql = ZSql + "'" + WIb + "',"
                    ZSql = ZSql + "'" + WIbTucu + "',"
                    ZSql = ZSql + "'" + WIbCiudad + "',"
                    ZSql = ZSql + "'" + WIbCiudadII + "',"
                    ZSql = ZSql + "'" + PorceIb.Text + "',"
                    ZSql = ZSql + "'" + PorceIbCaba.Text + "',"
                    ZSql = ZSql + "'" + WImpreVto + "',"
                    ZSql = ZSql + "'" + WIdioma + "',"
                    ZSql = ZSql + "'" + WEtiI + "',"
                    ZSql = ZSql + "'" + WEtiII + "',"
                    ZSql = ZSql + "'" + WDolarEspecial + "',"
                    ZSql = ZSql + "'" + Especif1.Text + "',"
                    ZSql = ZSql + "'" + Especif2.Text + "',"
                    ZSql = ZSql + "'" + Especif3.Text + "',"
                    ZSql = ZSql + "'" + Especif4.Text + "',"
                    ZSql = ZSql + "'" + Especif5.Text + "',"
                    ZSql = ZSql + "'" + Especif6.Text + "',"
                    ZSql = ZSql + "'" + Especif7.Text + "',"
                    ZSql = ZSql + "'" + Especif8.Text + "',"
                    ZSql = ZSql + "'" + Especif9.Text + "',"
                    ZSql = ZSql + "'" + Especif10.Text + "',"
                    ZSql = ZSql + "'" + DirEntregaII.Text + "',"
                    ZSql = ZSql + "'" + DirEntregaIII.Text + "',"
                    ZSql = ZSql + "'" + DirEntregaIV.Text + "',"
                    ZSql = ZSql + "'" + fechsedro + "',"
                    ZSql = ZSql + "'" + emailenv.Text + "',"
                    ZSql = ZSql + "'" + WRestriccion + "',"
                    ZSql = ZSql + "'" + EmailFactura.Text + "',"
                    ZSql = ZSql + "'" + DirEntregaV.Text + "')"
       
                    spCliente = ZSql
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Cliente SET "
                    ZSql = ZSql + "Cufe = " + "'" + Cufe.Text + "',"
                    ZSql = ZSql + "CufeII = " + "'" + CufeII.Text + "',"
                    ZSql = ZSql + "CufeIII = " + "'" + CufeIII.Text + "',"
                    ZSql = ZSql + "DirCufe = " + "'" + DirCufe.Text + "',"
                    ZSql = ZSql + "DirCufeII = " + "'" + DirCufeII.Text + "',"
                    ZSql = ZSql + "DirCufeIII = " + "'" + DirCufeIII.Text + "'"
                    ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                     
                    spCliente = ZSql
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                        Else
                            
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Cliente SET "
                    ZSql = ZSql + "Razon = " + "'" + Trim(Razon.Text) + "',"
                    ZSql = ZSql + "Direccion = " + "'" + Trim(Direccion.Text) + "',"
                    ZSql = ZSql + "Localidad = " + "'" + Trim(Localidad.Text) + "',"
                    ZSql = ZSql + "Pais = " + "'" + Pais.Text + "',"
                    ZSql = ZSql + "CuitII = " + "'" + CuitII.Text + "',"
                    ZSql = ZSql + "Provincia = " + "'" + WProvincia + "',"
                    ZSql = ZSql + "Postal = " + "'" + Trim(Postal.Text) + "',"
                    ZSql = ZSql + "Email = " + "'" + Trim(email.Text) + "',"
                    ZSql = ZSql + "Fax = " + "'" + Trim(fax.Text) + "',"
                    ZSql = ZSql + "Telefono = " + "'" + Telefono.Text + "',"
                    ZSql = ZSql + "Cuit = " + "'" + Cuit.Text + "',"
                    ZSql = ZSql + "Contacto = " + "'" + Contacto.Text + "',"
                    ZSql = ZSql + "Observaciones = " + "'" + Trim(Observaciones.Text) + "',"
                    ZSql = ZSql + "Vendedor = " + "'" + vendedor.Text + "',"
                    ZSql = ZSql + "Iva = " + "'" + WIva + "',"
                    ZSql = ZSql + "Rubro = " + "'" + Rubro.Text + "',"
                    ZSql = ZSql + "Horario = " + "'" + Trim(Horario.Text) + "',"
                    ZSql = ZSql + "NroSedronar = " + "'" + NroSedronar.Text + "',"
                    ZSql = ZSql + "Pago1 = " + "'" + Pago1.Text + "',"
                    ZSql = ZSql + "Pago2 = " + "'" + Pago2.Text + "',"
                    ZSql = ZSql + "Limite = " + "'" + Limite.Text + "',"
                    ZSql = ZSql + "Minimo = " + "'" + MInimo.Text + "',"
                    ZSql = ZSql + "DirEntrega = " + "'" + Trim(DirEntrega.Text) + "',"
                    ZSql = ZSql + "Importe1 = " + "'" + WImporte1 + "',"
                    ZSql = ZSql + "Importe2 = " + "'" + WImporte2 + "',"
                    ZSql = ZSql + "Importe3 = " + "'" + WImporte3 + "',"
                    ZSql = ZSql + "Importe4 = " + "'" + WImporte4 + "',"
                    ZSql = ZSql + "Importe5 = " + "'" + WImporte5 + "',"
                    ZSql = ZSql + "Importe6 = " + "'" + WImporte6 + "',"
                    ZSql = ZSql + "WDate = " + "'" + WDate + "',"
                    ZSql = ZSql + "Precio = " + "'" + Precio.Text + "',"
                    ZSql = ZSql + "NroIb = " + "'" + NroIb.Text + "',"
                    ZSql = ZSql + "NroIbTucu = " + "'" + NroIbTucu.Text + "',"
                    ZSql = ZSql + "PorceCm05Tucu = " + "'" + PorceCm05Tucu.Text + "',"
                    ZSql = ZSql + "NroIbCiudad = " + "'" + NroIbCiudad.Text + "',"
                    ZSql = ZSql + "Ib = " + "'" + WIb + "',"
                    ZSql = ZSql + "IbTucu = " + "'" + WIbTucu + "',"
                    ZSql = ZSql + "IbCiudad = " + "'" + WIbCiudad + "',"
                    ZSql = ZSql + "IbCiudadII = " + "'" + WIbCiudadII + "',"
                    ZSql = ZSql + "PorceIb = " + "'" + PorceIb.Text + "',"
                    ZSql = ZSql + "PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                    ZSql = ZSql + "ImpreVto = " + "'" + WImpreVto + "',"
                    ZSql = ZSql + "Idioma = " + "'" + WIdioma + "',"
                    ZSql = ZSql + "EtiI = " + "'" + WEtiI + "',"
                    ZSql = ZSql + "EtiII = " + "'" + WEtiII + "',"
                    ZSql = ZSql + "DolarEspecial = " + "'" + WDolarEspecial + "',"
                    ZSql = ZSql + "Especif1 = " + "'" + Especif1.Text + "',"
                    ZSql = ZSql + "Especif2 = " + "'" + Especif2.Text + "',"
                    ZSql = ZSql + "Especif3 = " + "'" + Especif3.Text + "',"
                    ZSql = ZSql + "Especif4 = " + "'" + Especif4.Text + "',"
                    ZSql = ZSql + "Especif5 = " + "'" + Especif5.Text + "',"
                    ZSql = ZSql + "Especif6 = " + "'" + Especif6.Text + "',"
                    ZSql = ZSql + "Especif7 = " + "'" + Especif7.Text + "',"
                    ZSql = ZSql + "Especif8 = " + "'" + Especif8.Text + "',"
                    ZSql = ZSql + "Especif9 = " + "'" + Especif9.Text + "',"
                    ZSql = ZSql + "Especif10 = " + "'" + Especif10.Text + "',"
                    ZSql = ZSql + "DirEntregaII = " + "'" + DirEntregaII.Text + "',"
                    ZSql = ZSql + "DirEntregaIII = " + "'" + DirEntregaIII.Text + "',"
                    ZSql = ZSql + "DirEntregaIV = " + "'" + DirEntregaIV.Text + "',"
                    ZSql = ZSql + "fechsedro = " + "'" + fechsedro + "',"
                    ZSql = ZSql + "Restriccion = " + "'" + WRestriccion + "',"
                    ZSql = ZSql + "Cufe = " + "'" + Cufe.Text + "',"
                    ZSql = ZSql + "CufeII = " + "'" + CufeII.Text + "',"
                    ZSql = ZSql + "CufeIII = " + "'" + CufeIII.Text + "',"
                    ZSql = ZSql + "DirCufe = " + "'" + DirCufe.Text + "',"
                    ZSql = ZSql + "DirCufeII = " + "'" + DirCufeII.Text + "',"
                    ZSql = ZSql + "DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                    ZSql = ZSql + "Emailenv = " + "'" + emailenv.Text + "',"
                    ZSql = ZSql + "EmailFactura = " + "'" + EmailFactura.Text + "',"
                    ZSql = ZSql + "DirEntregaV = " + "'" + DirEntregaV.Text + "'"
                    ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                     
                    spCliente = ZSql
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM ClienteEspecif"
                ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + Cliente.Text + "'"
                spClienteEspecif = ZSql
                Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
                If rstClienteEspecif.RecordCount > 0 Then
                    rstClienteEspecif.Close
            
                    ZSql = ""
                    ZSql = ZSql + "UPDATE ClienteEspecif SET "
                    ZSql = ZSql + "RequiereCertificado = " + "'" + Str$(RequiereCertificado.Value) + "',"
                    ZSql = ZSql + "RequiereMsds = " + "'" + Str$(RequiereMsds.Value) + "',"
                    ZSql = ZSql + "RequiereMsdsCada = " + "'" + Str$(RequiereMsdsCada.Value) + "',"
                    ZSql = ZSql + "RequiereHoja = " + "'" + Str$(RequiereHoja.Value) + "',"
                    ZSql = ZSql + "PermiteParcial = " + "'" + Str$(PermiteParcial.Value) + "',"
                    ZSql = ZSql + "PartidaVarias = " + "'" + Str$(PartidasVarias.Value) + "',"
                    ZSql = ZSql + "CantidadPartidas = " + "'" + CantidadPartidas.Text + "',"
                    ZSql = ZSql + "EmailCertificado = " + "'" + EmailCertificado.Text + "',"
                    ZSql = ZSql + "EmailMsds = " + "'" + EmailMsds.Text + "',"
                    ZSql = ZSql + "EmailHoja = " + "'" + EmailHoja.Text + "',"
                    ZSql = ZSql + "Especif1 = " + "'" + Especif1.Text + "',"
                    ZSql = ZSql + "Especif2 = " + "'" + Especif2.Text + "',"
                    ZSql = ZSql + "Especif3 = " + "'" + Especif3.Text + "',"
                    ZSql = ZSql + "Especif4 = " + "'" + Especif4.Text + "',"
                    ZSql = ZSql + "Especif5 = " + "'" + Especif5.Text + "',"
                    ZSql = ZSql + "DiasI = " + "'" + DiasI.Text + "',"
                    ZSql = ZSql + "DiasII = " + "'" + DiasII.Text + "',"
                    ZSql = ZSql + "DiasIII = " + "'" + DiasIII.Text + "',"
                    ZSql = ZSql + "EnvasesI = " + "'" + EnvasesI.Text + "',"
                    ZSql = ZSql + "EnvasesII = " + "'" + EnvasesII.Text + "',"
                    ZSql = ZSql + "EnvasesIII = " + "'" + EnvasesIII.Text + "',"
                    ZSql = ZSql + "EtiquetaI = " + "'" + EtiquetaI.Text + "',"
                    ZSql = ZSql + "EtiquetaII = " + "'" + EtiquetaII.Text + "'"
                    ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                     
                    spClienteEspecif = ZSql
                    Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO ClienteEspecif ("
                    ZSql = ZSql + "Cliente ,"
                    ZSql = ZSql + "RequiereCertificado ,"
                    ZSql = ZSql + "RequiereMsDs ,"
                    ZSql = ZSql + "RequiereMsDsCada ,"
                    ZSql = ZSql + "RequiereHoja ,"
                    ZSql = ZSql + "PermiteParcial ,"
                    ZSql = ZSql + "PartidaVarias ,"
                    ZSql = ZSql + "CantidadPartidas ,"
                    ZSql = ZSql + "EmailCertificado ,"
                    ZSql = ZSql + "EmailMsds ,"
                    ZSql = ZSql + "EmailHoja ,"
                    ZSql = ZSql + "Especif1 ,"
                    ZSql = ZSql + "Especif2 ,"
                    ZSql = ZSql + "Especif3 ,"
                    ZSql = ZSql + "Especif4 ,"
                    ZSql = ZSql + "Especif5 ,"
                    ZSql = ZSql + "DiasI ,"
                    ZSql = ZSql + "DiasII ,"
                    ZSql = ZSql + "DiasIII ,"
                    ZSql = ZSql + "EnvasesI ,"
                    ZSql = ZSql + "EnvasesII ,"
                    ZSql = ZSql + "EnvasesIII ,"
                    ZSql = ZSql + "EtiquetaI ,"
                    ZSql = ZSql + "EtiquetaII )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + Cliente.Text + "',"
                    ZSql = ZSql + "'" + Str$(RequiereCertificado.Value) + "',"
                    ZSql = ZSql + "'" + Str$(RequiereMsds.Value) + "',"
                    ZSql = ZSql + "'" + Str$(RequiereMsdsCada.Value) + "',"
                    ZSql = ZSql + "'" + Str$(RequiereHoja.Value) + "',"
                    ZSql = ZSql + "'" + Str$(PermiteParcial.Value) + "',"
                    ZSql = ZSql + "'" + Str$(PartidasVarias.Value) + "',"
                    ZSql = ZSql + "'" + CantidadPartidas.Text + "',"
                    ZSql = ZSql + "'" + EmailCertificado.Text + "',"
                    ZSql = ZSql + "'" + EmailMsds.Text + "',"
                    ZSql = ZSql + "'" + EmailHoja.Text + "',"
                    ZSql = ZSql + "'" + Especif1.Text + "',"
                    ZSql = ZSql + "'" + Especif2.Text + "',"
                    ZSql = ZSql + "'" + Especif3.Text + "',"
                    ZSql = ZSql + "'" + Especif4.Text + "',"
                    ZSql = ZSql + "'" + Especif5.Text + "',"
                    ZSql = ZSql + "'" + DiasI.Text + "',"
                    ZSql = ZSql + "'" + DiasII.Text + "',"
                    ZSql = ZSql + "'" + DiasIII.Text + "',"
                    ZSql = ZSql + "'" + EnvasesI.Text + "',"
                    ZSql = ZSql + "'" + EnvasesII.Text + "',"
                    ZSql = ZSql + "'" + EnvasesIII.Text + "',"
                    ZSql = ZSql + "'" + EtiquetaI.Text + "',"
                    ZSql = ZSql + "'" + EtiquetaII.Text + "')"
       
                    spClienteEspecif = ZSql
                    Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
                XEmpresa = Wempresa
                Erase CargaEmpresa
            
                If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
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
                        Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    CargaEmpresa(5, 1) = ""
                    CargaEmpresa(5, 2) = ""
                    CargaEmpresa(6, 1) = ""
                    CargaEmpresa(6, 2) = ""
                    CargaEmpresa(7, 1) = ""
                    CargaEmpresa(7, 2) = ""
                End If
                
                XEmpresa = Wempresa
                
                For Cicla = 1 To 7
                
                    If CargaEmpresa(Cicla, 1) <> "" Then
                
                        Wempresa = CargaEmpresa(Cicla, 1)
                        txtOdbc = CargaEmpresa(Cicla, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCliente.RecordCount > 0 Then
                            rstCliente.Close
                                
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Cliente SET "
                            ZSql = ZSql + "Razon = " + "'" + Trim(Razon.Text) + "',"
                            ZSql = ZSql + "Direccion = " + "'" + Trim(Direccion.Text) + "',"
                            ZSql = ZSql + "Localidad = " + "'" + Trim(Localidad.Text) + "',"
                            ZSql = ZSql + "Pais = " + "'" + Pais.Text + "',"
                            ZSql = ZSql + "CuitII = " + "'" + CuitII.Text + "',"
                            ZSql = ZSql + "Provincia = " + "'" + WProvincia + "',"
                            ZSql = ZSql + "Postal = " + "'" + Trim(Postal.Text) + "',"
                            ZSql = ZSql + "Email = " + "'" + Trim(email.Text) + "',"
                            ZSql = ZSql + "Fax = " + "'" + Trim(fax.Text) + "',"
                            ZSql = ZSql + "Telefono = " + "'" + Telefono.Text + "',"
                            ZSql = ZSql + "Cuit = " + "'" + Cuit.Text + "',"
                            ZSql = ZSql + "Contacto = " + "'" + Contacto.Text + "',"
                            ZSql = ZSql + "Observaciones = " + "'" + Trim(Observaciones.Text) + "',"
                            ZSql = ZSql + "Vendedor = " + "'" + vendedor.Text + "',"
                            ZSql = ZSql + "Iva = " + "'" + WIva + "',"
                            ZSql = ZSql + "Rubro = " + "'" + Rubro.Text + "',"
                            ZSql = ZSql + "Horario = " + "'" + Trim(Horario.Text) + "',"
                            ZSql = ZSql + "NroSedronar = " + "'" + NroSedronar.Text + "',"
                            ZSql = ZSql + "Pago1 = " + "'" + Pago1.Text + "',"
                            ZSql = ZSql + "Pago2 = " + "'" + Pago2.Text + "',"
                            ZSql = ZSql + "Limite = " + "'" + Limite.Text + "',"
                            ZSql = ZSql + "Minimo = " + "'" + MInimo.Text + "',"
                            ZSql = ZSql + "DirEntrega = " + "'" + Trim(DirEntrega.Text) + "',"
                            ZSql = ZSql + "Importe1 = " + "'" + WImporte1 + "',"
                            ZSql = ZSql + "Importe2 = " + "'" + WImporte2 + "',"
                            ZSql = ZSql + "Importe3 = " + "'" + WImporte3 + "',"
                            ZSql = ZSql + "Importe4 = " + "'" + WImporte4 + "',"
                            ZSql = ZSql + "Importe5 = " + "'" + WImporte5 + "',"
                            ZSql = ZSql + "Importe6 = " + "'" + WImporte6 + "',"
                            ZSql = ZSql + "WDate = " + "'" + WDate + "',"
                            ZSql = ZSql + "Precio = " + "'" + Precio.Text + "',"
                            ZSql = ZSql + "NroIb = " + "'" + NroIb.Text + "',"
                            ZSql = ZSql + "NroIbTucu = " + "'" + NroIbTucu.Text + "',"
                            ZSql = ZSql + "PorceCm05Tucu = " + "'" + PorceCm05Tucu.Text + "',"
                            ZSql = ZSql + "NroIbCiudad = " + "'" + NroIbCiudad.Text + "',"
                            ZSql = ZSql + "Ib = " + "'" + WIb + "',"
                            ZSql = ZSql + "IbTucu = " + "'" + WIbTucu + "',"
                            ZSql = ZSql + "IbCiudad = " + "'" + WIbCiudad + "',"
                            ZSql = ZSql + "IbCiudadII = " + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + "PorceIb = " + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + "PorceIbCaba = " + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + "ImpreVto = " + "'" + WImpreVto + "',"
                            ZSql = ZSql + "Idioma = " + "'" + WIdioma + "',"
                            ZSql = ZSql + "EtiI = " + "'" + WEtiI + "',"
                            ZSql = ZSql + "EtiII = " + "'" + WEtiII + "',"
                            ZSql = ZSql + "DolarEspecial = " + "'" + WDolarEspecial + "',"
                            ZSql = ZSql + "Especif1 = " + "'" + Especif1.Text + "',"
                            ZSql = ZSql + "Especif2 = " + "'" + Especif2.Text + "',"
                            ZSql = ZSql + "Especif3 = " + "'" + Especif3.Text + "',"
                            ZSql = ZSql + "Especif4 = " + "'" + Especif4.Text + "',"
                            ZSql = ZSql + "Especif5 = " + "'" + Especif5.Text + "',"
                            ZSql = ZSql + "Especif6 = " + "'" + Especif6.Text + "',"
                            ZSql = ZSql + "Especif7 = " + "'" + Especif7.Text + "',"
                            ZSql = ZSql + "Especif8 = " + "'" + Especif8.Text + "',"
                            ZSql = ZSql + "Especif9 = " + "'" + Especif9.Text + "',"
                            ZSql = ZSql + "Especif10 = " + "'" + Especif10.Text + "',"
                            ZSql = ZSql + "DirEntregaII = " + "'" + DirEntregaII.Text + "',"
                            ZSql = ZSql + "DirEntregaIII = " + "'" + DirEntregaIII.Text + "',"
                            ZSql = ZSql + "DirEntregaIV = " + "'" + DirEntregaIV.Text + "',"
                            ZSql = ZSql + "fechsedro = " + "'" + fechsedro + "',"
                            ZSql = ZSql + "Restriccion = " + "'" + WRestriccion + "',"
                            ZSql = ZSql + "Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + "CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + "CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + "DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + "DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + "DirCufeIII = " + "'" + DirCufeIII.Text + "',"
                            ZSql = ZSql + "Emailenv = " + "'" + emailenv.Text + "',"
                            ZSql = ZSql + "EmailFactura = " + "'" + EmailFactura.Text + "',"
                            ZSql = ZSql + "DirEntregaV = " + "'" + DirEntregaV.Text + "'"
                            ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                             
                            spCliente = ZSql
                            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                     
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Cliente ("
                            ZSql = ZSql + "Cliente ,"
                            ZSql = ZSql + "Razon ,"
                            ZSql = ZSql + "Direccion ,"
                            ZSql = ZSql + "Localidad ,"
                            ZSql = ZSql + "Pais ,"
                            ZSql = ZSql + "CuitII ,"
                            ZSql = ZSql + "Provincia ,"
                            ZSql = ZSql + "Postal ,"
                            ZSql = ZSql + "Email ,"
                            ZSql = ZSql + "Fax ,"
                            ZSql = ZSql + "Telefono ,"
                            ZSql = ZSql + "Cuit ,"
                            ZSql = ZSql + "Contacto ,"
                            ZSql = ZSql + "Observaciones ,"
                            ZSql = ZSql + "Vendedor ,"
                            ZSql = ZSql + "Iva ,"
                            ZSql = ZSql + "Rubro ,"
                            ZSql = ZSql + "Horario ,"
                            ZSql = ZSql + "NroSedronar ,"
                            ZSql = ZSql + "Pago1 ,"
                            ZSql = ZSql + "Pago2 ,"
                            ZSql = ZSql + "Limite ,"
                            ZSql = ZSql + "Minimo ,"
                            ZSql = ZSql + "DirEntrega ,"
                            ZSql = ZSql + "Importe1 ,"
                            ZSql = ZSql + "Importe2 ,"
                            ZSql = ZSql + "Importe3 ,"
                            ZSql = ZSql + "Importe4 ,"
                            ZSql = ZSql + "Importe5 ,"
                            ZSql = ZSql + "Importe6 ,"
                            ZSql = ZSql + "WDate ,"
                            ZSql = ZSql + "Precio ,"
                            ZSql = ZSql + "NroIb ,"
                            ZSql = ZSql + "NroIbTucu ,"
                            ZSql = ZSql + "PorceCm05Tucu ,"
                            ZSql = ZSql + "NroIbCiudad ,"
                            ZSql = ZSql + "Ib ,"
                            ZSql = ZSql + "IbTucu ,"
                            ZSql = ZSql + "IbCiudad ,"
                            ZSql = ZSql + "IbCiudadII ,"
                            ZSql = ZSql + "PorceIb ,"
                            ZSql = ZSql + "PorceIbCaba ,"
                            ZSql = ZSql + "ImpreVto ,"
                            ZSql = ZSql + "Idioma ,"
                            ZSql = ZSql + "EtiI ,"
                            ZSql = ZSql + "EtiII ,"
                            ZSql = ZSql + "DolarEspecial ,"
                            ZSql = ZSql + "Especif1 ,"
                            ZSql = ZSql + "Especif2 ,"
                            ZSql = ZSql + "Especif3 ,"
                            ZSql = ZSql + "Especif4 ,"
                            ZSql = ZSql + "Especif5 ,"
                            ZSql = ZSql + "Especif6 ,"
                            ZSql = ZSql + "Especif7 ,"
                            ZSql = ZSql + "Especif8 ,"
                            ZSql = ZSql + "Especif9 ,"
                            ZSql = ZSql + "Especif10 ,"
                            ZSql = ZSql + "DirEntregaII , "
                            ZSql = ZSql + "DirEntregaIII ,"
                            ZSql = ZSql + "DirEntregaIV ,"
                            ZSql = ZSql + "fechsedro ,"
                            ZSql = ZSql + "Restriccion ,"
                            ZSql = ZSql + "Emailenv ,"
                            ZSql = ZSql + "EmailFactura ,"
                            ZSql = ZSql + "DirEntregaV )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + Cliente.Text + "',"
                            ZSql = ZSql + "'" + Razon.Text + "',"
                            ZSql = ZSql + "'" + Direccion.Text + "',"
                            ZSql = ZSql + "'" + Localidad.Text + "',"
                            ZSql = ZSql + "'" + Pais.Text + "',"
                            ZSql = ZSql + "'" + CuitII.Text + "',"
                            ZSql = ZSql + "'" + WProvincia + "',"
                            ZSql = ZSql + "'" + Postal.Text + "',"
                            ZSql = ZSql + "'" + email.Text + "',"
                            ZSql = ZSql + "'" + fax.Text + "',"
                            ZSql = ZSql + "'" + Telefono.Text + "',"
                            ZSql = ZSql + "'" + Cuit.Text + "',"
                            ZSql = ZSql + "'" + Contacto.Text + "',"
                            ZSql = ZSql + "'" + Observaciones.Text + "',"
                            ZSql = ZSql + "'" + vendedor.Text + "',"
                            ZSql = ZSql + "'" + WIva + "',"
                            ZSql = ZSql + "'" + Rubro.Text + "',"
                            ZSql = ZSql + "'" + Horario.Text + "',"
                            ZSql = ZSql + "'" + NroSedronar.Text + "',"
                            ZSql = ZSql + "'" + Pago1.Text + "',"
                            ZSql = ZSql + "'" + Pago2.Text + "',"
                            ZSql = ZSql + "'" + Limite.Text + "',"
                            ZSql = ZSql + "'" + MInimo.Text + "',"
                            ZSql = ZSql + "'" + DirEntrega.Text + "',"
                            ZSql = ZSql + "'" + WImporte1 + "',"
                            ZSql = ZSql + "'" + WImporte2 + "',"
                            ZSql = ZSql + "'" + WImporte3 + "',"
                            ZSql = ZSql + "'" + WImporte4 + "',"
                            ZSql = ZSql + "'" + WImporte5 + "',"
                            ZSql = ZSql + "'" + WImporte6 + "',"
                            ZSql = ZSql + "'" + WDate + "',"
                            ZSql = ZSql + "'" + Precio.Text + "',"
                            ZSql = ZSql + "'" + NroIb.Text + "',"
                            ZSql = ZSql + "'" + NroIbTucu.Text + "',"
                            ZSql = ZSql + "'" + PorceCm05Tucu.Text + "',"
                            ZSql = ZSql + "'" + NroIbCiudad.Text + "',"
                            ZSql = ZSql + "'" + WIb + "',"
                            ZSql = ZSql + "'" + WIbTucu + "',"
                            ZSql = ZSql + "'" + WIbCiudad + "',"
                            ZSql = ZSql + "'" + WIbCiudadII + "',"
                            ZSql = ZSql + "'" + PorceIb.Text + "',"
                            ZSql = ZSql + "'" + PorceIbCaba.Text + "',"
                            ZSql = ZSql + "'" + WImpreVto + "',"
                            ZSql = ZSql + "'" + WIdioma + "',"
                            ZSql = ZSql + "'" + WEtiI + "',"
                            ZSql = ZSql + "'" + WEtiII + "',"
                            ZSql = ZSql + "'" + WDolarEspecial + "',"
                            ZSql = ZSql + "'" + Especif1.Text + "',"
                            ZSql = ZSql + "'" + Especif2.Text + "',"
                            ZSql = ZSql + "'" + Especif3.Text + "',"
                            ZSql = ZSql + "'" + Especif4.Text + "',"
                            ZSql = ZSql + "'" + Especif5.Text + "',"
                            ZSql = ZSql + "'" + Especif6.Text + "',"
                            ZSql = ZSql + "'" + Especif7.Text + "',"
                            ZSql = ZSql + "'" + Especif8.Text + "',"
                            ZSql = ZSql + "'" + Especif9.Text + "',"
                            ZSql = ZSql + "'" + Especif10.Text + "',"
                            ZSql = ZSql + "'" + DirEntregaII.Text + "',"
                            ZSql = ZSql + "'" + DirEntregaIII.Text + "',"
                            ZSql = ZSql + "'" + DirEntregaIV.Text + "',"
                            ZSql = ZSql + "'" + fechsedro + "',"
                            ZSql = ZSql + "'" + emailenv.Text + "',"
                            ZSql = ZSql + "'" + WRestriccion + "',"
                            ZSql = ZSql + "'" + EmailFactura.Text + "',"
                            ZSql = ZSql + "'" + DirEntregaV.Text + "')"
            
                            spCliente = ZSql
                            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                         
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Cliente SET "
                            ZSql = ZSql + "Cufe = " + "'" + Cufe.Text + "',"
                            ZSql = ZSql + "CufeII = " + "'" + CufeII.Text + "',"
                            ZSql = ZSql + "CufeIII = " + "'" + CufeIII.Text + "',"
                            ZSql = ZSql + "DirCufe = " + "'" + DirCufe.Text + "',"
                            ZSql = ZSql + "DirCufeII = " + "'" + DirCufeII.Text + "',"
                            ZSql = ZSql + "DirCufeIII = " + "'" + DirCufeIII.Text + "'"
                            ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                          
                            spCliente = ZSql
                            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                    End If
                        
                Next Cicla
                
                Call Conecta_Empresa
        
                Call CmdLimpiar_Click
        
            End If
        
            Cliente.SetFocus
        
        End If
        
    End If
    
    Exit Sub

WError:
     Resume Next
        
End Sub

Private Sub cmdDelete_Click()
    If Cliente.Text <> "" Then
        
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstClientes = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spCliente = "BorrarCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Cliente.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    fechsedro = "  /  /    "
    Restriccion.Value = 0
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Pais.Text = ""
    CuitII.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    vendedor.Text = ""
    DesVendedor.Caption = ""
    email.Text = ""
    emailenv.Text = ""
    EmailFactura.Text = ""
    fax.Text = ""
    Rubro.Text = ""
    DesRubro.Caption = ""
    Horario.Text = ""
    NroSedronar.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    Limite.Text = ""
    MInimo.Text = ""
    DirEntrega.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    DesRubro.Caption = ""
    Despago1.Caption = ""
    Despago2.Caption = ""
    Precio.Text = ""
    NroIb.Text = ""
    NroIbTucu.Text = ""
    PorceCm05Tucu.Text = ""
    NroIbCiudad.Text = ""
    DirEntregaII.Text = ""
    DirEntregaIII.Text = ""
    DirEntregaIV.Text = ""
    DirEntregaV.Text = ""
    Provincia.ListIndex = 25
    Ib.ListIndex = 0
    IbTucu.ListIndex = 0
    IbCiudad.ListIndex = 1
    IbCiudadII.ListIndex = 0
    PorceIbCaba.Text = ""
    PorceIb.Text = ""
    ImpreVto.ListIndex = 0
    Idioma.ListIndex = 0
    WGraba = ""
    
    RequiereCertificado.Value = 0
    RequiereMsds.Value = 0
    RequiereMsdsCada.Value = 0
    RequiereHoja.Value = 0
    PermiteParcial.Value = 0
    PartidasVarias.Value = 0
    CantidadPartidas.Text = ""
    
    EmailCertificado.Text = ""
    EmailMsds.Text = ""
    EmailHoja.Text = ""
    DiasI.Text = ""
    DiasII.Text = ""
    DiasIII.Text = ""
    EnvasesI.Text = ""
    EnvasesII.Text = ""
    EnvasesIII.Text = ""
    EtiquetaI.Text = ""
    EtiquetaII.Text = ""
    Especif1.Text = ""
    Especif2.Text = ""
    Especif3.Text = ""
    Especif4.Text = ""
    Especif5.Text = ""
    
    Cufe.Text = ""
    CufeII.Text = ""
    CufeIII.Text = ""
    DirCufe.Text = ""
    DirCufeII.Text = ""
    DirCufeIII.Text = ""
    
    EtiI.Value = False
    EtiII.Value = False
    DolarEspecial.Value = False
    
    Cliente.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Rem With rstPago
    Rem     .Close
    Rem End With
    Rem With rstClientes
    Rem     .Close
    Rem End With
    Rem With rstRubros
    Rem     .Close
    Rem End With
    Rem With rstVendedores
    Rem     .Close
    Rem End With
    Rem With rstEmpresa
    Rem     .Close
    Rem End With
    Rem With rstAuxiliar
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    Cliente.SetFocus
    prgcliente.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Command1_Click()

    Open "c:\padron\padron.txt" For Input As #10
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + "PorceIb = " + "'" + "3" + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Proveedor SET "
    ZSql = ZSql + "PorceIb = " + "'" + "1.75" + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSuma = 0
    aa = Time
    
    Do
        Line Input #10, WDatos
        If EOF(10) Then Exit Do
        
        WCuit = Mid$(WDatos, 28, 11)
        WCuitBusqueda = Mid$(WDatos, 30, 8)
        WPorceI = Mid$(WDatos, 46, 1) + "." + Mid$(WDatos, 48, 2)
        WPorceII = Mid$(WDatos, 51, 1) + "." + Mid$(WDatos, 53, 2)
        
        WCliente = ""
        WProveedor = ""
        
        ZSuma = ZSuma + 1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cuit LIKE " + "'" + "%" + WCuitBusqueda + "%" + "'"
        ZSql = ZSql + " Order by Cuit"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        WCliente = rstCliente!Cliente
                         .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCliente.Close
        End If
        
        If WCliente <> "" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Cuit LIKE " + "'" + "%" + WCuitBusqueda + "%" + "'"
        ZSql = ZSql + " Order by Cuit"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        WProveedor = rstProveedor!Proveedor
                         .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstProveedor.Close
        End If
        
        If WProveedor <> "" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceII + "'"
            ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Loop Until EOF(10)
    
    aaII = Time

End Sub

Private Sub Command2_Click()


    Dim ZGraba(5000, 15) As String
    
    Erase ZGraba
    ZLugar = 0

    spCliente = "ListaCliente"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    Rem dada
                    ZZCodigo = rstCliente!Cliente
                    ZZEspecif1 = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
                    ZZEspecif2 = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
                    ZZEspecif3 = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
                    ZZEspecif4 = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
                    ZZEspecif5 = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
                    ZZEspecif1 = RTrim(ZZEspecif1)
                    ZZEspecif2 = RTrim(ZZEspecif2)
                    ZZEspecif3 = RTrim(ZZEspecif3)
                    ZZEspecif4 = RTrim(ZZEspecif4)
                    ZZEspecif5 = RTrim(ZZEspecif5)
                    
                    ZLugar = ZLugar + 1
                    ZGraba(ZLugar, 1) = ZZCodigo
                    ZGraba(ZLugar, 2) = ZZEspecif1
                    ZGraba(ZLugar, 3) = ZZEspecif2
                    ZGraba(ZLugar, 4) = ZZEspecif3
                    ZGraba(ZLugar, 5) = ZZEspecif4
                    ZGraba(ZLugar, 6) = ZZEspecif5
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZZCodigo = ZGraba(Ciclo, 1)
        ZZEspecif1 = ZGraba(Ciclo, 2)
        ZZEspecif2 = ZGraba(Ciclo, 3)
        ZZEspecif3 = ZGraba(Ciclo, 4)
        ZZEspecif4 = ZGraba(Ciclo, 5)
        ZZEspecif5 = ZGraba(Ciclo, 6)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ClienteEspecif ("
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "RequiereCertificado ,"
        ZSql = ZSql + "RequiereMsDs ,"
        ZSql = ZSql + "RequiereMsDsCada ,"
        ZSql = ZSql + "RequiereHoja ,"
        ZSql = ZSql + "PermiteParcial ,"
        ZSql = ZSql + "PartidaVarias ,"
        ZSql = ZSql + "CantidadPartidas ,"
        ZSql = ZSql + "EmailCertificado ,"
        ZSql = ZSql + "EmailMsds ,"
        ZSql = ZSql + "EmailHoja ,"
        ZSql = ZSql + "Especif1 ,"
        ZSql = ZSql + "Especif2 ,"
        ZSql = ZSql + "Especif3 ,"
        ZSql = ZSql + "Especif4 ,"
        ZSql = ZSql + "Especif5 ,"
        ZSql = ZSql + "DiasI ,"
        ZSql = ZSql + "DiasII ,"
        ZSql = ZSql + "DiasIII ,"
        ZSql = ZSql + "EnvasesI ,"
        ZSql = ZSql + "EnvasesII ,"
        ZSql = ZSql + "EnvasesIII ,"
        ZSql = ZSql + "EtiquetaI ,"
        ZSql = ZSql + "EtiquetaII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCodigo + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + ZZEspecif1 + "',"
        ZSql = ZSql + "'" + ZZEspecif2 + "',"
        ZSql = ZSql + "'" + ZZEspecif3 + "',"
        ZSql = ZSql + "'" + ZZEspecif4 + "',"
        ZSql = ZSql + "'" + ZZEspecif5 + "',"
        ZSql = ZSql + "'" + "" + " ',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "')"

        spClienteEspecif = ZSql
        Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo





End Sub

Private Sub emailenvio_Click()

    OPEN_FILE_Email
    
    ZZSuma = 0
    Suma = 0
    ZZDirecion = ""
    
    With rstEmail
        .Index = "Codigo"
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                ZZRazon = rstEmail!Razon
                ZZEmail = IIf(IsNull(!email), "", !email)
                ZZEmail = Trim(ZZEmail)
                
                Rem ZZEmail = "d_esquenazi@yahoo.com"
                
                If ZZEmail <> "" Then
                
                    sTo = ZZEmail
                    sCC = ""
                    sBCC = ""
                    sSubject = "Aviso de Certificado de exclusion de IVA"
                    sBody = "Estimados Clientes" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "Por medio de la presente les adjuntamos el certificado de exclusion de IVA" + Chr$(13) _
                            + "" + Chr$(13) _
                            + "Saludamos cordialmente," + Chr$(13) _
                            + "" + Chr$(13) _
                            + "SURFACTAN S.A." + Chr$(13) _
                            + "4714-4097" + Chr$(13)
                            
                    SFile = "c:\email\Iva.pdf"
                
                    EmailAddress = sTo
                    CopiaAddress = sCC
                    MSubject = sSubject
                    MBody = sBody
                    MAttach = SFile
                    MAttachI = SFile
                    MAttachII = ""
                    MAttachIII = ""
                    MAttachIV = ""
                    MAttachV = ""
                
                    SendEmail
                
                End If
            
            
            Rem  Stop
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    

    
    
    m$ = "Proceso Finalizado"
    a% = MsgBox(m$, 0, "Envio de Avisio de Embargo")


End Sub

Private Sub Especificaciones_Click()

    IngresaEspecif.Height = 7575
    IngresaEspecif.Left = 120
    IngresaEspecif.Top = 360
    IngresaEspecif.Width = 11775
    
    IngresaEspecif.Visible = True
    
    Rem Especif1.SetFocus

End Sub


Private Sub Lista_Click()
    Desde.Text = "A00000"
    Hasta.Text = "Z99999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub NroSedronar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        fechsedro.SetFocus
    End If
End Sub

Private Sub Razon_Keypress(KeyAscii As Integer)
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
        Pais.SetFocus
    End If
End Sub

Private Sub Pais_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Localidad.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Postal.Text) < 1000 Then
            m$ = "El Codigo Postal debe ser mayor a 1000"
            a% = MsgBox(m$, 0, "Archivo de Cliente")
            Postal.SetFocus
                Else
            Cuit.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CuitII.SetFocus
    End If
End Sub

Private Sub CuitII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Telefono.SetFocus
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vendedor.SetFocus
    End If
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVendedor = vendedor.Text
        spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            DesVendedor.Caption = rstVendedor!Nombre
            Contacto.SetFocus
                Else
            vendedor.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Contacto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        email.SetFocus
    End If
End Sub

Private Sub EMail_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fax.SetFocus
    End If
End Sub

Private Sub Fax_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rubro.SetFocus
    End If
End Sub

Private Sub Rubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WRubro = Rubro.Text
        spRubro = "ConsultaRubro " + "'" + Rubro.Text + "'"
        Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
        If rstRubro.RecordCount > 0 Then
            DesRubro.Caption = rstRubro!Nombre
            Horario.SetFocus
                    Else
            Rubro.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Horario_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pago1.SetFocus
    End If
End Sub

Private Sub Pago1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPago = Pago1.Text
        spPago = "ConsultaPago " + "'" + Pago1.Text + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            Despago1.Caption = rstPago!Nombre
            Pago2.SetFocus
                Else
            Pago1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pago2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPago = Pago2.Text
        spPago = "ConsultaPago " + "'" + Pago2.Text + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            Despago2.Caption = rstPago!Nombre
            Limite.SetFocus
                Else
            Pago2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Limite_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Limite.Text = Pusing("###,###.##", Limite.Text)
        MInimo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MInimo.Text = Pusing("###,###.##", MInimo.Text)
        Razon.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DirEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirEntregaII.SetFocus
    End If
End Sub

Private Sub DirEntregaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirEntregaIII.SetFocus
    End If
End Sub

Private Sub DirEntregaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirEntregaIV.SetFocus
    End If
End Sub

Private Sub DirEntregaIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DirEntregaV.SetFocus
    End If
End Sub

Private Sub DirEntregaV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Razon.SetFocus
        PantaDirEntrega.Visible = False
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            If Len(Cliente.Text) < 6 Then
                m$ = "El codigo de cliente debe tener 6 digitos"
                a% = MsgBox(m$, 0, "Archivo de Cliente")
                Cliente.SetFocus
                    Else
                WCliente = Cliente.Text
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WPasa = "S"
                        Else
                    WPasa = "N"
                End If
                
                If WPasa = "S" Then
                    Cliente.Text = rstCliente!Cliente
                    Call Imprime_Datos
                        Else
                    WCliente = Cliente.Text
                    CmdLimpiar_Click
                    Cliente.Text = WCliente
                End If
                    
                Razon.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    ''Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    ''Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Vendedores"
     Opcion.AddItem "Rubro"

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
            Ayuda.Text = ""
            Ayuda.Visible = True
            spCliente = "ListaClienteConsulta"
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
            Ayuda.SetFocus
        
        Case 1
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstVendedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstVendedor!vendedor) + " " + rstVendedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstVendedor!vendedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstVendedor.Close
            
        Case 2
            spRubro = "ListaRubro"
            Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstRubro
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstRubro!Rubro) + " " + rstRubro!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstRubro!Rubro
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstRubro.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                Call Imprime_Datos
                    Else
                CmdLimpiar_Click
                Cliente.Text = Claveven$
            End If
            
            Cliente.SetFocus
            
        Case 1
            Indice = Pantalla.ListIndex
            WVendedor = WIndice.List(Indice)
            spVendedor = "ConsultaVendedor " + "'" + Str$(WVendedor) + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                vendedor.Text = rstVendedor!vendedor
                Call Imprime_Descripcion
                    Else
                CmdLimpiar_Click
                vendedor.Text = Claveven$
            End If
            
            vendedor.SetFocus
            
        Case 2
            Indice = Pantalla.ListIndex
            WRubro = WIndice.List(Indice)
            spRubro = "ConsultaRubro " + "'" + Str$(WRubro) + "'"
            Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
            If rstRubro.RecordCount > 0 Then
                Rubro.Text = rstRubro!Rubro
                Call Imprime_Descripcion
                    Else
                CmdLimpiar_Click
                Rubro.Text = Claveven$
            End If
            
            Rubro.SetFocus
                    
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Pais.Text = ""
    CuitII.Text = ""
    
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    vendedor.Text = ""
    DesVendedor.Caption = ""
    email.Text = ""
    fax.Text = ""
    Rubro.Text = ""
    DesRubro.Caption = ""
    Horario.Text = ""
    NroSedronar.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    Limite.Text = ""
    MInimo.Text = ""
    DirEntrega.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    Despago1.Caption = ""
    Despago2.Caption = ""
    Precio.Text = ""
    NroIb.Text = ""
    NroIbTucu.Text = ""
    PorceCm05Tucu.Text = ""
    NroIbCiudad.Text = ""
    DirEntregaII.Text = ""
    DirEntregaIII.Text = ""
    DirEntregaIV.Text = ""
    DirEntregaV.Text = ""
    PorceIb.Text = ""
    PorceIbCaba.Text = ""
    Restriccion.Value = 0
    
    Cufe.Text = ""
    CufeII.Text = ""
    CufeIII.Text = ""
    DirCufe.Text = ""
    DirCufeII.Text = ""
    DirCufeIII.Text = ""
    
    
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
    
    Provincia.ListIndex = 25
    
    Ib.Clear
    
    Ib.AddItem "Inscriptos"
    Ib.AddItem "Convenio Multilateral"
    Ib.AddItem "Exentos"
    
    Ib.ListIndex = 0
    
    IbTucu.Clear
    
    IbTucu.AddItem "Exentos"
    IbTucu.AddItem "Convenio Multilateral S/Sede en Tucuman"
    IbTucu.AddItem "Convenio Multilateral C/sede en Tucuman"
    IbTucu.AddItem "Convenio Multilateral con Alta en Tucuman"
    IbTucu.AddItem "Contribuyente Local de Tucuman"
    IbTucu.AddItem "Contribuyente C/dom. en Tucuman y s/alta en Tucuman"
    
    IbTucu.ListIndex = 0
    
    IbCiudad.Clear
    
    IbCiudad.AddItem "Exentos Certificado"
    IbCiudad.AddItem "Retiene Normal"
    IbCiudad.AddItem "Retiene Riesgo"
    IbCiudad.AddItem "Exentos Lugar Ent."
    IbCiudad.AddItem "Exentos Agente"
    
    IbCiudad.ListIndex = 1
    
    
    IbCiudadII.Clear
    
    IbCiudadII.AddItem ""
    IbCiudadII.AddItem "Local"
    IbCiudadII.AddItem "Conv. Multilateral"
    IbCiudadII.AddItem "No Inscripto"
    IbCiudadII.AddItem "Reg. Simplificado"
    
    IbCiudadII.ListIndex = 0
    
    
    ImpreVto.Clear
    
    ImpreVto.AddItem ""
    ImpreVto.AddItem "Exige"
    
    ImpreVto.ListIndex = 0
    
    
    Idioma.Clear
    
    Idioma.AddItem "Castellano"
    Idioma.AddItem "Ingles"
    
    Idioma.ListIndex = 0
    
    
    RequiereCertificado.Value = 0
    RequiereMsds.Value = 0
    RequiereMsdsCada.Value = 0
    RequiereHoja.Value = 0
    PermiteParcial.Value = 0
    PartidasVarias.Value = 0
    CantidadPartidas.Text = ""
    
    EmailCertificado.Text = ""
    EmailMsds.Text = ""
    EmailHoja.Text = ""
    DiasI.Text = ""
    DiasII.Text = ""
    DiasIII.Text = ""
    EnvasesI.Text = ""
    EnvasesII.Text = ""
    EnvasesIII.Text = ""
    EtiquetaI.Text = ""
    EtiquetaII.Text = ""
    Especif1.Text = ""
    Especif2.Text = ""
    Especif3.Text = ""
    Especif4.Text = ""
    Especif5.Text = ""
    
    EtiI.Value = False
    EtiII.Value = False
    DolarEspecial.Value = False
    
End Sub

Private Sub Primer_Click()

     On Error GoTo WError
    
    spCliente = "ListaClientes"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCliente
        .MoveFirst
        Cliente.Text = rstCliente!Cliente
    End With
    
    rstCliente.Close
    Call Imprime_Datos
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cliente", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spCliente = "ListaClientes"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCliente
        .MoveLast
        Cliente.Text = rstCliente!Cliente
    End With
    Cliente.SetFocus
    rstCliente.Close
    Call Imprime_Datos
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Cliente", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spCliente = "AnteriorCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCliente
        .MoveLast
        Cliente.Text = rstCliente!Cliente
    End With
    
    rstCliente.Close
    Cliente.SetFocus
    Call Imprime_Datos
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cliente", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spCliente = "PosteriorCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCliente
        .MoveFirst
        Cliente.Text = rstCliente!Cliente
    End With
    
    rstCliente.Close
    Cliente.SetFocus
    Call Imprime_Datos
   Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cliente", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
    
End Sub


Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
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
        rstCliente.Close
    End If
    End If

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
        If WClave.Text = "NORMA" Then
            WGraba = "S"
            XClave.Visible = False
            Call cmdAdd_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
    End If
End Sub

Private Sub FinEspecificaciones_Click()
    IngresaEspecif.Visible = False
End Sub

Private Sub Especif1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif2.SetFocus
    End If
End Sub

Private Sub Especif2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif3.SetFocus
    End If
End Sub

Private Sub Especif3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif4.SetFocus
    End If
End Sub

Private Sub Especif4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif5.SetFocus
    End If
End Sub

Private Sub Especif5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif6.SetFocus
    End If
End Sub

Private Sub Especif6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif7.SetFocus
    End If
End Sub

Private Sub Especif7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif8.SetFocus
    End If
End Sub

Private Sub Especif8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif9.SetFocus
    End If
End Sub

Private Sub Especif9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif10.SetFocus
    End If
End Sub

Private Sub Especif10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Especif1.SetFocus
    End If
End Sub



Private Sub Ib_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroIb.SetFocus
    End If
End Sub

Private Sub NroIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceIb.SetFocus
    End If
End Sub

Private Sub PorceIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IbTucu.SetFocus
    End If
End Sub

Private Sub IbTucu_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroIbTucu.SetFocus
    End If
End Sub

Private Sub NroIbTucu_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceCm05Tucu.SetFocus
    End If
End Sub

Private Sub PorceCm05Tucu_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IbCiudad.SetFocus
    End If
End Sub

Private Sub IbCiudad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IbCiudadII.SetFocus
    End If
End Sub

Private Sub IbCiudadII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroIbCiudad.SetFocus
    End If
End Sub

Private Sub NroIbCiudad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ib.SetFocus
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
    
    
    
    
    







