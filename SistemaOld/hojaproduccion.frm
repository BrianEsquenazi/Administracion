VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHojaProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Hoja de Produccion"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.Frame PantaDesvio 
      Caption         =   "Informe de Desvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   720
      TabIndex        =   139
      Top             =   2160
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox ObservaDesvio 
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
         MaxLength       =   50
         TabIndex        =   143
         Text            =   " "
         Top             =   1560
         Width           =   7455
      End
      Begin VB.ComboBox MotivoDesvio 
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
         Left            =   240
         TabIndex        =   140
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label26 
         Caption         =   "Rendimiento teorico"
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
         TabIndex        =   142
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   "Rendimiento teorico"
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
         TabIndex        =   141
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton LoteColo 
      Caption         =   "Lote Colorante"
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
      TabIndex        =   150
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame CargaPartida 
      Height          =   3495
      Left            =   2040
      TabIndex        =   146
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaLote 
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
         Height          =   500
         Left            =   1920
         TabIndex        =   152
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton AceptaLote 
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
         Left            =   720
         TabIndex        =   151
         Top             =   2760
         Width           =   975
      End
      Begin VB.ListBox PantaLote 
         Height          =   1815
         Left            =   360
         TabIndex        =   149
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox LoteColorante 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   147
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label27 
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
         TabIndex        =   148
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame XClave 
      Caption         =   "Ingrese de Clave de Seguridad"
      Height          =   1815
      Left            =   3960
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Grabacion"
         Height          =   375
         Left            =   600
         TabIndex        =   49
         Top             =   1200
         Width           =   2535
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
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   48
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame IngresaEnsayo 
      Height          =   2775
      Left            =   0
      TabIndex        =   52
      Top             =   1560
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton CierraIngresaEnsayo 
         Caption         =   "Cierra Pantalla"
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
         Left            =   7680
         TabIndex        =   104
         Top             =   5880
         Width           =   1935
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
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   103
         Text            =   " "
         Top             =   6480
         Width           =   3975
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   102
         Text            =   " "
         Top             =   6840
         Width           =   3975
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox valor2 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox valor4 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   61
         Text            =   " "
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox valor5 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   60
         Text            =   " "
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox valor6 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   59
         Text            =   " "
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox valor7 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   58
         Text            =   " "
         Top             =   3600
         Width           =   3975
      End
      Begin VB.TextBox valor8 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   57
         Text            =   " "
         Top             =   4080
         Width           =   3975
      End
      Begin VB.TextBox valor9 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   56
         Text            =   " "
         Top             =   4560
         Width           =   3975
      End
      Begin VB.TextBox valor10 
         Height          =   285
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   55
         Text            =   " "
         Top             =   5040
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   54
         Text            =   " "
         Top             =   5760
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   6120
         Width           =   3975
      End
      Begin VB.Label Std1010 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   122
         Top             =   5280
         Width           =   4455
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   121
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Std99 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   120
         Top             =   4800
         Width           =   4455
      End
      Begin VB.Label Std88 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   119
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Std77 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   118
         Top             =   3840
         Width           =   4455
      End
      Begin VB.Label Std66 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   117
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Std55 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   116
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Std44 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   115
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Std33 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   114
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   113
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   112
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblensayo 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         Left            =   240
         TabIndex        =   101
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDescri 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         TabIndex        =   100
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         Left            =   3360
         TabIndex        =   99
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   98
         Top             =   720
         Width           =   2220
      End
      Begin VB.Label descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   97
         Top             =   1200
         Width           =   2220
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   96
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   95
         Top             =   2160
         Width           =   2220
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   94
         Top             =   2640
         Width           =   2220
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   93
         Top             =   3120
         Width           =   2220
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   92
         Top             =   3600
         Width           =   2220
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   91
         Top             =   4080
         Width           =   2220
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   90
         Top             =   4560
         Width           =   2220
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   89
         Top             =   5040
         Width           =   2220
      End
      Begin VB.Label Label18 
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
         TabIndex        =   88
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label17 
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
         TabIndex        =   87
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label15 
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
         TabIndex        =   86
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   85
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         TabIndex        =   84
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   83
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   82
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   81
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   80
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Std6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   79
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   78
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   77
         Top             =   4080
         Width           =   4455
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   76
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   75
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label Ensayo1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   74
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Ensayo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   73
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Ensayo3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   72
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Ensayo4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   71
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Ensayo5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   70
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Ensayo6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   69
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Ensayo7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   68
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Ensayo8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   67
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Ensayo9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   66
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Ensayo10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   240
         TabIndex        =   65
         Top             =   5040
         Width           =   735
      End
   End
   Begin VB.TextBox NroPedido 
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
      Left            =   10680
      MaxLength       =   6
      TabIndex        =   144
      Text            =   " "
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox HoraFinal 
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
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   138
      Text            =   " "
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox HoraInicio 
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
      MaxLength       =   10
      TabIndex        =   135
      Text            =   " "
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Equipo 
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
      Left            =   9240
      MaxLength       =   2
      TabIndex        =   131
      Text            =   " "
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox VersionI 
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   127
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox VersionII 
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
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   126
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox VersionIII 
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
      Left            =   11280
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   125
      Text            =   " "
      Top             =   120
      Width           =   495
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
      Height          =   495
      Left            =   1200
      TabIndex        =   124
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Block2 
      Caption         =   "Cerrar Block"
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
      TabIndex        =   123
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   11160
      TabIndex        =   51
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"hojaproduccion.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   106
      Top             =   1800
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   107
      Top             =   1800
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   108
      Top             =   1800
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   109
      Top             =   1800
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   110
      Top             =   1800
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   111
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Anula 
      Caption         =   "Anula Confirm."
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   50
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   9120
      TabIndex        =   34
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WControl3 
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
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox WControl2 
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
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox WControl1 
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
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   720
         Width           =   375
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
         Left            =   1200
         TabIndex        =   42
         Top             =   1440
         Width           =   855
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
         Left            =   1200
         TabIndex        =   41
         Top             =   1080
         Width           =   855
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
         Left            =   1200
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox WLote3 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox WLote2 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   38
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox WLote1 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label13 
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
         Left            =   1200
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label dada 
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
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox fechaIng 
      Height          =   285
      Left            =   8040
      TabIndex        =   5
      Top             =   840
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
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   480
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
   Begin VB.TextBox Real 
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
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Teorico 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   4200
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
      TabIndex        =   22
      Top             =   7080
      Width           =   975
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
      Left            =   3360
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   120
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
   Begin VB.TextBox Hoja 
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
      Left            =   2040
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
      TabIndex        =   16
      Top             =   6480
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
      TabIndex        =   15
      Top             =   7080
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
      Left            =   3360
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   8895
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
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   28
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.TextBox WTipo 
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
         MaxLength       =   1
         TabIndex        =   26
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   14
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   12
         Top             =   600
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
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
         Left            =   7440
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   3840
         TabIndex        =   32
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label9 
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
         Left            =   2400
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prodcuto Terminado"
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
         Left            =   840
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label WDescripcion 
         BackColor       =   &H00FFFF00&
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
         Left            =   3840
         TabIndex        =   11
         Top             =   600
         Width           =   3615
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
      TabIndex        =   9
      Top             =   7080
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10080
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "hojaproduccion.frx":007C
      Left            =   3360
      List            =   "hojaproduccion.frx":0083
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   8415
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
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Vector 
      Height          =   3735
      Left            =   120
      TabIndex        =   105
      Top             =   1560
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
      _Version        =   327680
      BackColor       =   16777088
   End
   Begin MSMask.MaskEdBox FechaInicio 
      Height          =   285
      Left            =   3360
      TabIndex        =   133
      Top             =   1200
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
   Begin MSMask.MaskEdBox FechaFinal 
      Height          =   285
      Left            =   9240
      TabIndex        =   136
      Top             =   1200
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
   Begin VB.Label LabelPedido 
      Caption         =   "Pedido"
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
      Left            =   9480
      TabIndex        =   145
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label24 
      Caption         =   "Fecha y Hora de Envasamiento"
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
      Left            =   6240
      TabIndex        =   137
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label23 
      Caption         =   "Fecha y Hora Inicio  Produccion"
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
      TabIndex        =   134
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label22 
      Caption         =   "Equipo"
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
      Left            =   8400
      TabIndex        =   132
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "Version Formula"
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
      Left            =   6000
      TabIndex        =   130
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "Procedim. "
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
      Left            =   8400
      TabIndex        =   129
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Especif."
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
      Left            =   10080
      TabIndex        =   128
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha Ingreso"
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
      Left            =   6600
      TabIndex        =   25
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Rendimiento Real"
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
      Left            =   3600
      TabIndex        =   24
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Rendimiento teorico"
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
      TabIndex        =   23
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label DesProducto 
      BackColor       =   &H00FFFF00&
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
      Left            =   3600
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Producto"
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
      TabIndex        =   19
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
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Hoja de Produccion"
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
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgHojaProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Tipo As String
Private Existe  As String
Private Auxi1 As String
Private Auxi2 As String
Private XIndice As Integer
Private WImpre As String
Private Cantidad As String
Private Auxiliar(100, 20) As String
Private ZAuxiliar(100, 7) As String
Private XLote(1000, 7) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHoja As Recordset
Dim ZZLoteColo(1000, 2) As String

Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPrecio As Recordset
Dim spPrecio As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String

Dim XParam As String
Dim LeeHoja As String
Dim Ultimo As Integer
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim WEstado As String
Private BajaLote(3, 2) As String
Private WControla As String
Private WSaldoant As Double
Private ZCantidad As Double
Private WExiste As String
Dim XCosto1 As Double
Dim XCosto2 As Double
Dim XCosto3 As Double
Dim WCosto1 As String
Dim WCosto2 As String
Dim WCosto3 As String
Dim WE As String
Dim Verifica(100, 10) As String
Dim WLugar As Integer
Dim WEntraLote As Integer
Dim WCicloLote As Integer
Dim XCicloLote As Integer
Dim Entra As String
Dim WCompara As String

Dim ZZArticulo As String
Dim ZZLote As String
Dim ZZMarcaVencida As String
Dim Empe(12, 10) As String
Dim ZFechaVto As String
Dim XMes As String
Dim XAno As String
Dim ZVto As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer

Private Sub AceptaLote_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.LoteColorante = " + "'" + LoteColorante.Text + "'"
    ZSql = ZSql + " and Hoja.Producto <> " + "'" + Producto.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.Close
        Exit Sub
    End If
    
    CargaPartida.Visible = False
    
End Sub

Private Sub Block1_Click()

    On Error GoTo WError

    Agenda.LoadFile "blanco.rtf", 0
    Agenda.LoadFile "H" + Hoja.Text + ".rtf", 0
    Agenda.Visible = True
    Block1.Visible = False
    Block2.Visible = True
    Agenda.Height = 6700
    Agenda.Left = 840
    Agenda.Top = 720
    Agenda.Width = 9375
    Agenda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Block2_Click()
    Agenda.SaveFile "H" + Hoja.Text + ".rtf", 0
    Agenda.Visible = False
    Block1.Visible = True
    Block2.Visible = False
End Sub

Private Sub Borra_Click()

    Vector.Col = 1
    Vector.Text = ""
    
    Vector.Col = 2
    Vector.Text = ""

    Vector.Col = 3
    Vector.Text = ""
    
    Vector.Col = 4
    Vector.Text = ""
    
    Vector.Col = 5
    Vector.Text = ""
    
    Vector.Col = 6
    Vector.Text = ""
    
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True

    CargaLote.Visible = False
    WLinea.Text = ""
    WArticulo.SetFocus
    
End Sub

Private Sub CancelaLote_Click()
    CargaPartida.Visible = False
End Sub

Private Sub cmdClose_Click()
    PrgHojaProduccion.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Materia Prima"
     Opcion.AddItem "Productos Terminados"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
End Sub

Private Sub LoteColo_Click()

    Erase ZZLoteColo
    ZZLugar = 0
    
    ZZLugar = ZZLugar + 1
    ZZLoteColo(ZZLugar, 1) = Producto.Text
    ZZLoteColo(ZZLugar, 2) = Hoja.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Producto = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " and Hoja.LoteColorante <> " + "'" + "" + "'"
    ZSql = ZSql + " Order by Hoja.LoteColorante"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZLoteColorante = Trim(rstHoja!LoteColorante)
                
                    ZZEntra = "S"
                    For ZZCiclo = 1 To ZZLugar
                        If ZZLoteColo(ZZCiclo, 2) = ZZLoteColorante Then
                            ZZEntra = "N"
                        End If
                    Next ZZCiclo
                    
                    If ZZEntra = "S" Then
                        
                        ZZLugar = ZZLugar + 1
                        ZZLoteColo(ZZLugar, 1) = Producto.Text
                        ZZLoteColo(ZZLugar, 2) = ZZLoteColorante
                        
                    End If
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstHoja.Close
    
    End If
    
    For iRow = 1 To 40
        
        WRow = iRow
        WLugar = iRow
            
        Tipo = Vector.TextMatrix(WRow, 1)
        Terminado = UCase(Vector.TextMatrix(WRow, 2))
        Articulo = UCase(Vector.TextMatrix(WRow, 3))
        Cantidad = Vector.TextMatrix(WRow, 5)
        Estado = Vector.TextMatrix(WRow, 6)
        
        If Trim(Articulo) <> "" Then
        
            If Left$(Articulo, 2) = "DY" Then
                
                For WCicloLote = 1 To 6 Step 2
                    
                    Lote = XLote(WLugar, WCicloLote)
                    ZZLoteColorante = ""
                        
                    If Lote <> "" Then
                    
                        XParam = "'" + Lote + "','" _
                                + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZLoteColorante = rstLaudo!PartiOri
                            rstLaudo.Close
                            
                            ZZEntra = "S"
                            For ZZCiclo = 1 To ZZLugar
                                If ZZLoteColo(ZZCiclo, 2) = ZZLocaColorante Then
                                    ZZEntra = "N"
                                End If
                            Next ZZCiclo
                            
                            If ZZEntra = "S" Then
                                
                                ZZLugar = ZZLugar + 1
                                ZZLoteColo(ZZLugar, 1) = Articulo
                                ZZLoteColo(ZZLugar, 2) = ZZLoteColorante
                                
                            End If
                            
                        End If
                    
                    End If
                        
                Next WCicloLote
                
            End If
        
        End If
                        
    Next iRow
    
    
    
    
    
    
    PantaLote.Clear
    For ZZCiclo = 1 To ZZLugar
        PantaLote.AddItem ZZLoteColo(ZZCiclo, 1) + "     " + ZZLoteColo(ZZCiclo, 2)
    Next ZZCiclo

    CargaPartida.Visible = True
    LoteColorante.Text = ""
    LoteColorante.SetFocus
    
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
            
        Case 1
        
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Rem IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            IngresaItem = rstTerminado!Codigo
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub OTRO_Click()
    Call Anula_Click
End Sub

Private Sub PantaLote_Click()
    LoteColorante.Text = ZZLoteColo(PantaLote.ListIndex + 1, 2)
    LoteColorante.SetFocus
End Sub

Private Sub Vector_GotFocus()

    Vector.Col = 1
    If Len(Vector.Text) = 1 Then
        WLinea.Text = Vector.Row
        WTipo.Text = Vector.Text
            Else
        WTipo.Text = ""
        WLinea.Text = ""
    End If

    Vector.Col = 2
    If Len(Vector.Text) = 12 Then
        WTerminado.Text = Vector.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    Vector.Col = 3
    If Len(Vector.Text) = 10 Then
        WArticulo.Text = Vector.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    Vector.Col = 4
    WDescripcion.Caption = Vector.Text

    Vector.Col = 5
    WCantidad.Text = Vector.Text
    
    WCompara = WCantidad.Text
    
    If Val(Teorico.Text) = 0 Then
        Teorico.SetFocus
            Else
        WCantidad.SetFocus
    End If
        

End Sub

Private Sub Graba_Click()

    If Val(Wempresa) = 1 Then

        WTerminado = Producto.Text
        XCodigo = Val(Mid$(WTerminado, 4, 5))
        XTipoPro = ""
        If Left$(WTerminado, 2) = "PT" Then
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 12999 Then
                    XTipoPro = "CO"
                End If
            End If
        End If
        
        If XTipoPro = "CO" Then
            If Trim(LoteColorante.Text) = "" Then
                ca% = MsgBox("Se debe informar la partida original", 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        End If
        
    End If
    
    If Val(Real.Text) = 0 Or fechaIng.Text = "  /  /    " Then
        ca% = MsgBox("Se debe informar el Rendimiento Real", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If

    Call Valida_fecha(FechaInicio.Text, Auxi)
    If Auxi <> "S" Or FechaInicio.Text = "  /  /    " Then
        ca% = MsgBox("Se debe informar la fecha y hora de inicio de produccion", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If

    If Val(HoraInicio.Text) <= 0 Or Val(HoraInicio.Text) > 24 Then
        ca% = MsgBox("Se debe informar la fecha y hora de inicio de produccion", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If
    
    Call Valida_fecha(FechaFinal.Text, Auxi)
    If Auxi <> "S" Or FechaFinal.Text = "  /  /    " Then
        ca% = MsgBox("Se debe informar la fecha y hora de inicio envasamiento", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
        Rem BY NAN 20-11-2012 verif fecha final
            Else
        ZZZFechaInicio = Right$(FechaInicio, 4) + Mid$(FechaInicio, 4, 2) + Left$(FechaInicio, 2)
        ZZZFechaFinal = Right$(FechaFinal, 4) + Mid$(FechaFinal, 4, 2) + Left$(FechaFinal, 2)
        If ZZZFechaFinal < ZZZFechaInicio Then
            ca% = MsgBox("La fecha final de Envasamiento no puede ser menor que la de inicio de Produccion", 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        Rem fin BY NAN
    End If

    If Val(HoraFinal.Text) <= 0 Or Val(HoraFinal.Text) > 24 Then
        ca% = MsgBox("Se debe informar la fecha y hora de inicio envasamiento", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If

    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Hoja"
    ZSql = ZSql & " Where Hoja.Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WMarcaLabora = IIf(IsNull(rstHoja!MarcaLabora), "", rstHoja!MarcaLabora)
        WProducto = rstHoja!Producto
        WTeorico = rstHoja!Teorico
        WReal = rstHoja!Real
        rstHoja.Close
      Rem by nan
       If WMarcaLabora <> "S" Then
            ca% = MsgBox("La Hoja de Produccion NO fue actualizada por laboratorio", 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        If WReal <> 0 Then
            ca% = MsgBox("La Hoja ya actualizada por Cotiza", 0, "Ingreso de Hoja de Produccion")
           Exit Sub
        End If
            Else
        ca% = MsgBox("Hoja de Produccion Inexistente", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
   End If
    
    If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
    
    If Val(NroPedido.Text) <> 0 Then
        
        XEmpresa = Wempresa
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZEntra = "N"
        If Producto.Text <> "YH" Then
            ZZProducto = "PT-5" + Mid$(Producto.Text, 5, 8)
                Else
            ZZProducto = "PT-03001-001"
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
        ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            WSaldo = rstPedido!Cantidad - rstPedido!Facturado
            If WSaldo > 0 Then
                ZEntra = "S"
            End If
            rstPedido.Close
        End If
            
        Call Conecta_Empresa
            
        If ZEntra = "N" Then
            m$ = "Nro de Pedido Incorrecto"
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
            Else
            
        WMarca = 0
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WMarca = IIf(IsNull(rstTerminado!Marca), "0", rstTerminado!Marca)
            rstTerminado.Close
        End If
        If WMarca = 0 Then
            m$ = "Se debe informar Nro de Pedido"
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Rem Exit Sub
        End If
        
    End If
    
    End If
    




    General = "S"
    XSuma = 0
    WLugar = 0
    WEntraLote = 0
    Erase Verifica
        
    For iRow = 1 To 40
        
        WRow = iRow
            
        Tipo = Vector.TextMatrix(WRow, 1)
        Terminado = UCase(Vector.TextMatrix(WRow, 2))
        Articulo = UCase(Vector.TextMatrix(WRow, 3))
        Cantidad = Vector.TextMatrix(WRow, 5)
        Estado = Vector.TextMatrix(WRow, 6)
        
        XSuma = XSuma + Val(Cantidad)
        WLugar = WLugar + 1
            
        For WCicloLote = 1 To 6 Step 2
            
            WLote1 = XLote(WLugar, WCicloLote)
            WCanti = XLote(WLugar, WCicloLote + 1)
                
            If WLote1 <> "" Or WCanti <> "" Then
                
            Entra = "S"
            For XCicloLote = 1 To WEntraLote
                If Val(Verifica(XCicloLote, 1)) = Val(WLote1) Then
                    Verifica(XCicloLote, 2) = Str$(Val(Verifica(XCicloLote, 2)) + Val(WCanti))
                    Entra = "N"
                    Exit For
                End If
            Next XCicloLote
                        
            If Entra = "S" Then
                WEntraLote = WEntraLote + 1
                Verifica(WEntraLote, 1) = WLote1
                Verifica(WEntraLote, 2) = WCanti
                Verifica(WEntraLote, 3) = Tipo
                Verifica(WEntraLote, 4) = Terminado
                Verifica(WEntraLote, 5) = Articulo
            End If
            
            End If
                
        Next WCicloLote
                    
        If Articulo <> "" Then
                        
            If Tipo = "M" Then
        
                WEntra = "N"
        
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    rstArticulo.Close
                End If
            
                If WControla = 0 Then
                    If Estado = "S" Then
                        WEntra = "S"
                    End If
                End If
                    
                If WControla <> 0 Then
                    WEntra = "S"
                End If
               
                If WEntra <> "S" Then
                    m$ = Articulo + " Articulo inexistente o Lote nro. " + Lote + " inexistente"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    General = "N"
                End If
            
                    Else
        
                WEntra = "N"
                    
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
            
                If WControla = 0 Then
                    If Estado = "S" Then
                        WEntra = "S"
                    End If
                End If
                    
                If WControla <> 0 Then
                    WEntra = "S"
                End If
                    
                If WEntra <> "S" Then
                    m$ = Terminado + " Producto inexistente o Lote nro. " + Lote + " inexistente"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    General = "N"
                End If
            
            End If
        End If
                        
    Next iRow
    
    For Ciclo = 1 To WEntraLote
    
        WSaldo1 = 0
        
        WLote1 = Verifica(Ciclo, 1)
        WCanti = Verifica(Ciclo, 2)
        Tipo = Verifica(Ciclo, 3)
        Terminado = Verifica(Ciclo, 4)
        Articulo = Verifica(Ciclo, 5)
        
        If Tipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            Articulo = UCase(Articulo)
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1 + "','" _
                            + Articulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + Articulo + "','" _
                            + WLote1 + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
            
            If WEntra <> "S" Then
                m$ = Articulo + " Articulo inexistente o Lote nro. " + WLote1 + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1 + "','" _
                        + Terminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + Terminado + "','" _
                            + WLote1 + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                General = "N"
                m$ = Terminado + " Producto inexistente o Lote nro. " + WLote1 + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo1 < Val(WCanti) And WControla = 0 Then
            General = "N"
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            If Tipo = "M" Then
                m$ = Articulo + " Cantidad Insuficiente Stock : " + XSaldo1 + " Lote Nro. " + WLote1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = Terminado + " Cantidad Insuficiente Stock : " + XSaldo1 + " Lote Nro. " + WLote1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
        End If
        
    Next Ciclo
    
     If Val(Real.Text) = 0 Then
        m$ = "Cantidad real en 0"
        G% = MsgBox(m$, 0, "Actualizacion de Hoja de Produccion")
        General = "N"
    End If
    
    If General = "S" Then
    
        Dife = Abs(Val(Real.Text) - XSuma)
        Porce = Abs(XSuma * 0.07)
        If Dife > Porce Then
            T$ = "Grabacion de Hoja de Produccion"
            m$ = "El rendimiento real difiere de los componentes en un +-7%. Desea continuar con la grabacion"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 7 Then
                Exit Sub
            End If
        End If
        
        
        
        Dife = Abs(Val(Real.Text) - Val(Teorico.Text))
        Porce = Abs(Val(Teorico.Text) * 0.07)
        If Dife > Porce Then
            T$ = "Grabacion de Hoja de Produccion"
            m$ = "El rendimiento real difiere de la cantidad teorica en un +-7%. Desea continuar con la grabacion"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 7 Then
                Exit Sub
            End If
        End If
        
        
        
        Dife = Abs(Val(Teorico.Text) - XSuma)
        If Val(Teorico.Text) <> XSuma Then
            T$ = "Grabacion de Hoja de Produccion"
            m$ = "La suma de material ingresado (" + Str$(XSuma) + ") es distinto a la cantidad teorica (" + Teorico.Text + ")" + Chr$(13) + _
                "Si ratifica los datos ingresados pulse [SI] " + Chr$(13) + _
                "Si desea cambiar la cantidad teorica pulse [NO], haga los cambios corespondientes y vuela a grabar la hoja"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 7 Then
                Exit Sub
            End If
        End If
        
    

        
        
        
        
        
        
        
        WHoja = Hoja.Text
        WFecha = Fecha.Text
        WProducto = Producto.Text
        WTeorico = Teorico.Text
        WReal = Real.Text
        WFechaing = fechaIng.Text

        PLote = Hoja.Text
        PTerminado = Producto.Text
        
        Teorico.Text = Pusing("###,###.##", Teorico.Text)
        Real.Text = Pusing("###,###.##", Real.Text)
        
        ZZHoja = Hoja.Text
        ZZFecha = Fecha.Text
        ZZProducto = Producto.Text
        ZZTeorico = Teorico.Text
        ZZReal = Real.Text
        ZZFechaing = fechaIng.Text
        

        Renglon = 0
        Erase Auxiliar
        
        spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                        Renglon = Renglon + 1
                        Auxiliar(Renglon, 1) = rstHoja!Producto
                        Auxiliar(Renglon, 2) = rstHoja!Terminado
                        Auxiliar(Renglon, 3) = rstHoja!Articulo
                        Auxiliar(Renglon, 4) = rstHoja!Cantidad
                        Auxiliar(Renglon, 5) = rstHoja!Real
                        Auxiliar(Renglon, 6) = rstHoja!Teorico
                        Auxiliar(Renglon, 7) = rstHoja!Tipo
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
            
        End If
    
        For Da = 1 To Renglon
    
            WWProducto = Auxiliar(Da, 1)
            Terminado = Auxiliar(Da, 2)
            Articulo = Auxiliar(Da, 3)
            Cantidad = Auxiliar(Da, 4)
            Real = Auxiliar(Da, 5)
            Teorico = Auxiliar(Da, 6)
            Tipo = Auxiliar(Da, 7)
        
            If Da = 1 Then
        
                spTerminado = "ConsultaTerminado " + "'" + WWProducto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WProceso = Str$(rstTerminado!Proceso - Teorico)
                    WEntradas = Str$(rstTerminado!Entradas)
                    WDate = Date$
                    rstTerminado.Close
                    
                    XParam = "'" + WCodigo + "','" _
                                + WEntradas + "','" _
                                + WProceso + "','" _
                                + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoHoja " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            End If
                
            Select Case Tipo
                Case "M"
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = rstArticulo!Codigo
                        WSalidas = Str$(rstArticulo!Salidas - Cantidad)
                        WDate = Date$
                        XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                        rstArticulo.Close
                                            
                        spArticulo = "ModificaArticuloSalidas " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                                        
                Case "T"
            
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = rstTerminado!Codigo
                        WSalidas = Str$(rstTerminado!Salidas - Cantidad)
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                                            
                        spTerminado = "ModificaTerminadoSalidas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
            
                Case Else
            End Select
        
        Next Da
    
        spHoja = "BorrarHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)
        
        Hoja.Text = ZZHoja
        Fecha.Text = ZZFecha
        Producto.Text = ZZProducto
        Teorico.Text = ZZTeorico
        Real.Text = ZZReal
        fechaIng.Text = ZZFechaing
        
    
        Renglon = 0
        Erase Auxiliar
        
        Suma = 0
        
        For iRow = 1 To 40
        
            Suma = Suma + 1
            WRow = iRow
            
            Tipo = Vector.TextMatrix(WRow, 1)
            Terminado = UCase(Vector.TextMatrix(WRow, 2))
            Articulo = UCase(Vector.TextMatrix(WRow, 3))
            Cantidad = Vector.TextMatrix(WRow, 5)
            Lote = Vector.TextMatrix(WRow, 6)
                    
            If Articulo <> "" Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Hoja.Text)
                Call Ceros(Auxi1, 6)
                    
                WClave = Auxi1 + Auxi
                WRenglon = Str$(Renglon)
                WHoja = Hoja.Text
                WFecha = Fecha.Text
                WProducto = Producto.Text
                WTeorico = Teorico.Text
                WReal = Real.Text
                WFechaing = fechaIng.Text
                WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Cantidad
                WLote = "0"
                WWDate = Date$
                WImporte = ""
                WMarca = ""
                WSaldo = Str$(Val(Real.Text))
                
                WLote1 = XLote(Suma, 1)
                WLote2 = XLote(Suma, 3)
                WLote3 = XLote(Suma, 5)
                WCanti1 = XLote(Suma, 2)
                WCanti2 = XLote(Suma, 4)
                WCanti3 = XLote(Suma, 6)
                WCosto1 = "0"
                WCosto2 = "0"
                WCosto3 = "0"
                    
                XCosto1 = 0
                XCosto2 = 0
                XCosto3 = 0
                
                Select Case Tipo
                    Case "T"
                        ZZProducto = Producto.Text
                        Producto = Terminado
                        Call Calcula_Costo_Produccion(Producto, XCosto1, XCosto2, XCosto3)
                        Producto.Text = ZZProducto
                        
                
                    Case "M"
                        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            XCosto1 = rstArticulo!Costo1
                            XCosto2 = rstArticulo!Costo2
                            XCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                            rstArticulo.Close
                        End If
                        
                    Case Else
                End Select
        
                WCosto1 = Str$(XCosto1)
                WCosto2 = Str$(XCosto2)
                WCosto3 = Str$(XCosto3)

                XParam = "'" + WClave + "','" _
                            + WHoja + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WProducto + "','" _
                            + WCantidad + "','" _
                            + WTipo + "','" _
                            + WLote + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WTeorico + "','" _
                            + WReal + "','" _
                            + WFechaing + "','" _
                            + WFechaingord + "','" _
                            + WDate + "','" _
                            + WImporte + "','" _
                            + WMarca + "','" _
                            + WSaldo + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + WLote3 + "','" + WCanti3 + "','" _
                            + WCosto1 + "','" _
                            + WCosto2 + "','" _
                            + WCosto3 + "'"
                                           
                spHoja = "AltaHoja " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
                ZImpresionI = "N"
                ZImpresionII = "N"
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " LoteColorante = " + "'" + LoteColorante.Text + "',"
                ZSql = ZSql + " NroPedido = " + "'" + NroPedido.Text + "',"
                ZSql = ZSql + " Equipo = " + "'" + Equipo.Text + "',"
                ZSql = ZSql + " VersionI = " + "'" + VersionI.Text + "',"
                ZSql = ZSql + " VersionII = " + "'" + VersionII.Text + "',"
                ZSql = ZSql + " VersionIII = " + "'" + VersionIII.Text + "',"
                ZSql = ZSql + " ImpresionI = " + "'" + ZImpresionI + "',"
                ZSql = ZSql + " ImpresionII = " + "'" + ZImpresionII + "',"
                ZSql = ZSql + " FechaInicio = " + "'" + FechaInicio.Text + "',"
                ZSql = ZSql + " HoraInicio = " + "'" + HoraInicio.Text + "',"
                ZSql = ZSql + " FechaFinal = " + "'" + FechaFinal.Text + "',"
                ZSql = ZSql + " HoraFinal = " + "'" + HoraFinal.Text + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                Auxiliar(Renglon, 1) = WProducto
                Auxiliar(Renglon, 2) = WTerminado
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WCantidad
                Auxiliar(Renglon, 5) = WReal
                Auxiliar(Renglon, 6) = WTeorico
                Auxiliar(Renglon, 7) = WTipo
                Auxiliar(Renglon, 8) = WLote1
                Auxiliar(Renglon, 9) = WCanti1
                Auxiliar(Renglon, 10) = WLote2
                Auxiliar(Renglon, 11) = WCanti2
                Auxiliar(Renglon, 12) = WLote3
                Auxiliar(Renglon, 13) = WCanti3
                
            End If
                        
        Next iRow
    
        WHoja = Hoja.Text
        WFecha = Fecha.Text
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WHoja + "','" _
                     + WFechaord + "'"
        Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)

        For Da = 1 To Renglon

            WWProducto = Auxiliar(Da, 1)
            Terminado = Auxiliar(Da, 2)
            Articulo = Auxiliar(Da, 3)
            Cantidad = Auxiliar(Da, 4)
            Real = Auxiliar(Da, 5)
            Teorico = Auxiliar(Da, 6)
            Tipo = Auxiliar(Da, 7)
            BajaLote(1, 1) = Auxiliar(Da, 8)
            BajaLote(1, 2) = Auxiliar(Da, 9)
            BajaLote(2, 1) = Auxiliar(Da, 10)
            BajaLote(2, 2) = Auxiliar(Da, 11)
            BajaLote(3, 1) = Auxiliar(Da, 12)
            BajaLote(3, 2) = Auxiliar(Da, 13)
        
            If Da = 1 Then
        
                spTerminado = "ConsultaTerminado " + "'" + WWProducto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WEntradas = Str$(rstTerminado!Entradas + Val(Real))
                    WProceso = Str$(rstTerminado!Proceso)
                    WDate = Date$
                    rstTerminado.Close
                        
                    XParam = "'" + WCodigo + "','" _
                        + WEntradas + "','" _
                        + WProceso + "','" _
                        + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoHoja " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
                
            Select Case Tipo
                Case "M"
            
                    For xda = 1 To 3
                
                        Lote = BajaLote(xda, 1)
                        Cantidad = BajaLote(xda, 2)
            
                        WControla = 0
                        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                            WCodigo = rstArticulo!Codigo
                            WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                            WDate = Date$
                            XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                            rstArticulo.Close
                                            
                            spArticulo = "ModificaArticuloSalidas " + XParam
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Lote = BajaLote(xda, 1)
                            Cantidad = BajaLote(xda, 2)
                    
                            If WControla = 0 And Val(Lote) <> 0 Then
                        
                                XParam = "'" + Lote + "','" _
                                        + Articulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + WClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                        Else
                                
                                    XParam = "'" + Articulo + "','" _
                                            + Lote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + WClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                            
                                End If
                            End If
                    
                        End If
                    Next xda
                                            
                Case "T"
            
                    For xda = 1 To 3
                
                        Lote = BajaLote(xda, 1)
                        Cantidad = BajaLote(xda, 2)
            
                        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            WCodigo = rstTerminado!Codigo
                            WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                            WDate = Date$
                            rstTerminado.Close
                        
                            XParam = "'" + WCodigo + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                                            
                            spTerminado = "ModificaTerminadoSalidas " + XParam
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        
                            Lote = BajaLote(xda, 1)
                            Cantidad = BajaLote(xda, 2)
                        
                            If WControla = 0 And Val(Lote) <> 0 Then
                                XParam = "'" + Lote + "','" _
                                            + Terminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                                    WDate = Date$
                                    rstHoja.Close
                                
                                    XParam = "'" + WClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                
                                        Else
                                
                                    XParam = "'" + Terminado + "','" _
                                                + Lote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + WClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                            
                                End If
                            End If
                    
                        End If
                    
                    Next xda
                    
                Case Else
            End Select
        
        Next Da
        
        Rem If Left$(Producto.Text, 2) = "DW" Then
        Rem     Call Actualiza_Hoja
        Rem End If
                   
        If Val(NroPedido.Text) <> 0 Then

            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZSaldo = 0
            ZEntra = "N"
            If Producto.Text <> "YH" Then
                ZZProducto = "PT-5" + Mid$(Producto.Text, 5, 8)
                    Else
                ZZProducto = "PT-03001-001"
            End If
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
            ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
            
                ZSaldo = rstPedido!Cantidad - rstPedido!Facturado
                rstPedido.Close
                
                If ZSaldo > Val(Real.Text) Then
                
                    T$ = "Grabacion de Hoja de Produccion"
                    m$ = "El rendimiento real de la hoja de produccion  (" + Real.Text + ") es menor a la cantidad solicitada en el pedido (" + Str$(ZSaldo) + ")" + Chr$(13) + _
                         "Si desea dar por cumplido el pedido pulse [SI] " + Chr$(13) + _
                         "Si desea dejar pendiente los " + Str$(ZSaldo - Val(Real.Text)) + " pulse [NO]"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    
                    If Respuesta% = 6 Then
                    
                        ZResta = Str$(ZSaldo - Val(Real.Text))
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Pedido SET "
                        ZSql = ZSql + " Cantidad = Cantidad - " + "'" + ZResta + "'"
                        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
                        ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
                        spPedido = ZSql
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                
                XCantidad1 = Real.Text
                xCantidad2 = Real.Text
                XLote1 = Hoja.Text
                XCantiLote1 = Real.Text
                XLote2 = ""
                XCantiLote2 = ""
                XLote3 = ""
                XCantiLote3 = ""
                XLote4 = ""
                XCantiLote4 = ""
                XLote5 = ""
                XCantiLote5 = ""
                XEnv1 = ""
                XCantiEnv1 = ""
                XEnv2 = ""
                XCantiEnv2 = ""
                XEnv3 = ""
                XCantiEnv3 = ""
                XEnv4 = ""
                XCantiEnv4 = ""
                XEnv5 = ""
                XCantiEnv5 = ""
                XBultos1 = ""
                XBultos2 = ""
                XBultos3 = ""
                XBultos4 = ""
                XBultos5 = ""
                ZFechaActualiza = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZOrdFechaActualiza = Right$(ZFechaActualiza, 4) + Mid$(ZFechaActualiza, 4, 2) + Left$(ZFechaActualiza, 2)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + "Cantidad1 = " + "'" + XCantidad1 + "',"
                ZSql = ZSql + "Cantidad2 = " + "'" + xCantidad2 + "',"
                ZSql = ZSql + "Lote1 = " + "'" + XLote1 + "',"
                ZSql = ZSql + "CantiLote1 = " + "'" + XCantiLote1 + "',"
                ZSql = ZSql + "Lote2 = " + "'" + XLote2 + "',"
                ZSql = ZSql + "CantiLote2 = " + "'" + XCantiLote2 + "',"
                ZSql = ZSql + "Lote3 = " + "'" + XLote3 + "',"
                ZSql = ZSql + "CantiLote3 = " + "'" + XCantiLote3 + "',"
                ZSql = ZSql + "Lote4 = " + "'" + XLote4 + "',"
                ZSql = ZSql + "CantiLote4 = " + "'" + XCantiLote4 + "',"
                ZSql = ZSql + "Lote5 = " + "'" + XLote5 + "',"
                ZSql = ZSql + "CantiLote5 = " + "'" + XCantiLote5 + "',"
                ZSql = ZSql + "Env1 = " + "'" + XEnv1 + "',"
                ZSql = ZSql + "CantiEnv1 = " + "'" + XCantiEnv1 + "',"
                ZSql = ZSql + "Env2 = " + "'" + XEnv2 + "',"
                ZSql = ZSql + "CantiEnv2 = " + "'" + XCantiEnv2 + "',"
                ZSql = ZSql + "Env3 = " + "'" + XEnv3 + "',"
                ZSql = ZSql + "CantiEnv3 = " + "'" + XCantiEnv3 + "',"
                ZSql = ZSql + "Env4 = " + "'" + XEnv4 + "',"
                ZSql = ZSql + "CantiEnv4 = " + "'" + XCantiEnv4 + "',"
                ZSql = ZSql + "Env5 = " + "'" + XEnv5 + "',"
                ZSql = ZSql + "CantiEnv5 = " + "'" + XCantiEnv5 + "',"
                ZSql = ZSql + "CantidadFac = " + "'" + "0" + "',"
                ZSql = ZSql + "Bultos1 = " + "'" + XBultos1 + "',"
                ZSql = ZSql + "Bultos2 = " + "'" + XBultos2 + "',"
                ZSql = ZSql + "Bultos3 = " + "'" + XBultos3 + "',"
                ZSql = ZSql + "Bultos4 = " + "'" + XBultos4 + "',"
                ZSql = ZSql + "Bultos5 = " + "'" + XBultos5 + "',"
                ZSql = ZSql + "FechaActualizacion = " + "'" + ZFechaActualiza + "',"
                ZSql = ZSql + "OrdFechaActualizacion = " + "'" + ZOrdFechaActualiza + "'"
                ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
                ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
                
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Pedido SET "
                ZSql = ZSql & " MarcaFactura = " + "'" + "1" + "'"
                ZSql = ZSql & " Where Pedido = " + "'" + NroPedido.Text + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            Call Conecta_Empresa
            
        End If
        
        PorceDesvio = 0
        If Val(Real.Text) <> 0 And Val(Teorico.Text) <> 0 Then
            PorceDesvio = Abs((Val(Real.Text) - Val(Teorico.Text)) / (Val(Teorico.Text) / 100))
        End If
        
        If PorceDesvio >= 3 And Val(Wempresa) <> 5 Then
        
            MotivoDesvio.ListIndex = 0
            ObservaDesvio.Text = ""
            PantaDesvio.Visible = True
            MotivoDesvio.SetFocus
            
                Else
                
            Call Limpia_Click
            
            Vector.TopRow = 1
            Vector.Col = 1
            Vector.Row = 1
    
            Hoja.SetFocus
            
        End If
    
    End If
        
        
End Sub

Private Sub MotivoDesvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If MotivoDesvio.ListIndex > 0 Then
            ObservaDesvio.SetFocus
        End If
    End If
End Sub

Private Sub ObservaDesvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If MotivoDesvio.ListIndex <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " MotivoDesvio = " + "'" + Str$(MotivoDesvio.ListIndex) + "',"
            ZSql = ZSql + " ObservaDesvio = " + "'" + ObservaDesvio.Text + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            PantaDesvio.Visible = False
            
            Call Limpia_Click
            
            Vector.TopRow = 1
            Vector.Col = 1
            Vector.Row = 1
            
            Hoja.SetFocus
            
                Else
                
            m$ = "Se debe informar un motivo de desvio"
            G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            
        End If
    End If
End Sub

Private Sub Anula_Click()

    If WExiste = "S" Then
        Call Ingresa_clave
        Exit Sub
    End If
    
    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WProducto = Producto.Text
    WTeorico = Teorico.Text
    WReal = Real.Text
    WFechaing = fechaIng.Text

    PLote = Hoja.Text
    PTerminado = Producto.Text
    
    Renglon = 0
    Erase Auxiliar
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then

    With rstHoja
        .MoveFirst
        Do
            If .EOF = False Then
                Renglon = Renglon + 1
                Auxiliar(Renglon, 1) = rstHoja!Producto
                Auxiliar(Renglon, 2) = rstHoja!Terminado
                Auxiliar(Renglon, 3) = rstHoja!Articulo
                Auxiliar(Renglon, 4) = rstHoja!Cantidad
                Auxiliar(Renglon, 5) = rstHoja!Real
                Auxiliar(Renglon, 6) = rstHoja!Teorico
                Auxiliar(Renglon, 7) = rstHoja!Tipo
                Auxiliar(Renglon, 8) = rstHoja!lote1
                Auxiliar(Renglon, 9) = rstHoja!Canti1
                Auxiliar(Renglon, 10) = rstHoja!lote2
                Auxiliar(Renglon, 11) = rstHoja!Canti2
                Auxiliar(Renglon, 12) = rstHoja!lote3
                Auxiliar(Renglon, 13) = rstHoja!Canti3
                
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstHoja.Close
    
    End If
    
    
    For Da = 1 To Renglon
    
        WWProducto = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Real = Auxiliar(Da, 5)
        Teorico = Auxiliar(Da, 6)
        Tipo = Auxiliar(Da, 7)
        BajaLote(1, 1) = Auxiliar(Da, 8)
        BajaLote(1, 2) = Auxiliar(Da, 9)
        BajaLote(2, 1) = Auxiliar(Da, 10)
        BajaLote(2, 2) = Auxiliar(Da, 11)
        BajaLote(3, 1) = Auxiliar(Da, 12)
        BajaLote(3, 2) = Auxiliar(Da, 13)
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + WWProducto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                If Real <> 0 Then
                    WEntradas = Str$(rstTerminado!Entradas - Real)
                    WProceso = Str$(rstTerminado!Proceso)
                        Else
                    WProceso = Str$(rstTerminado!Proceso - Teorico)
                    WEntradas = Str$(rstTerminado!Entradas)
                End If
                WDate = Date$
                rstTerminado.Close
                
                XParam = "'" + WCodigo + "','" _
                    + WEntradas + "','" _
                    + WProceso + "','" _
                    + WDate + "'"
                                           
                spTerminado = "ModificaTerminadoHoja " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
                
        Select Case Tipo
            Case "M"
            
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WSalidas = Str$(rstArticulo!Salidas - Cantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                    rstArticulo.Close
                                            
                    spArticulo = "ModificaArticuloSalidas " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                For xda = 1 To 3
                
                    Lote = BajaLote(xda, 1)
                    ZCantidad = BajaLote(xda, 2)
            
                    WControla = 0
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        WCodigo = rstArticulo!Codigo
                        rstArticulo.Close
                    
                        Lote = BajaLote(xda, 1)
                        ZCantidad = BajaLote(xda, 2)
                    
                        If WControla = 0 And Val(Lote) <> 0 Then
                        
                            XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                            spLaudo = "ListaLaudoArticulo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WClave = rstLaudo!Clave
                                WSaldo = Str$(rstLaudo!Saldo + ZCantidad)
                                WDate = Date$
                                rstLaudo.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spLaudo = "ModificaLaudoSaldo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                
                                XParam = "'" + Articulo + "','" _
                                        + Lote + "'"
                                spMovguia = "ListaMovguiaLote " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + ZCantidad)
                                    WDate = Date$
                                    rstMovguia.Close
                            
                                    XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                                    spMovguia = "ModificaMovguiaSaldo " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                            
                            End If
                        End If
                    
                    End If
                Next xda
                                        
            Case "T"
            
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WSalidas = Str$(rstTerminado!Salidas - Cantidad)
                    WDate = Date$
                    rstTerminado.Close
                        
                    XParam = "'" + WCodigo + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                            
                    spTerminado = "ModificaTerminadoSalidas " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                For xda = 1 To 3
                
                    Lote = BajaLote(xda, 1)
                    ZCantidad = BajaLote(xda, 2)
            
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        WCodigo = rstTerminado!Codigo
                        rstTerminado.Close
                        
                        Lote = BajaLote(xda, 1)
                        ZCantidad = BajaLote(xda, 2)
                        
                        If WControla = 0 And Val(Lote) <> 0 Then
                            XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WClave = rstHoja!Clave
                                WSaldo = Str$(rstHoja!Saldo + ZCantidad)
                                WDate = Date$
                                rstHoja.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spHoja = "ModificaHojaSaldo " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                
                                XParam = "'" + Terminado + "','" _
                                        + Lote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + ZCantidad)
                                    WDate = Date$
                                    rstMovguia.Close
                            
                                    XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                                    spMovguia = "ModificaMovguiaSaldo " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                            
                            End If
                        End If
                    
                    End If
                    
                Next xda
                    
            Case Else
        End Select
        
    Next Da
    
    spHoja = "BorrarHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)
    
    Renglon = 0
    Erase Auxiliar
        
    Hoja.Text = WHoja
    Fecha.Text = WFecha
    Producto.Text = WProducto
    Teorico.Text = WTeorico
    WTeorico = Pusing("###,###.##", Teorico.Text)
    Real.Text = "0"
    fechaIng.Text = "  /  /    "
    WReal = "0"
    WFechaing = "  /  /    "
    
    For iRow = 1 To 40
        
        Suma = Suma + 1
                
        WRow = iRow
        Vector.Row = WRow
            
        Tipo = Vector.TextMatrix(WRow, 1)
        Terminado = UCase(Vector.TextMatrix(WRow, 2))
        Articulo = UCase(Vector.TextMatrix(WRow, 3))
        Cantidad = Vector.TextMatrix(WRow, 5)
        Lote = Vector.TextMatrix(WRow, 6)
                    
        If Articulo <> "" Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Hoja.Text)
            Call Ceros(Auxi1, 6)
                    
            WClave = Auxi1 + Auxi
            WHoja = WHoja
            WRenglon = Str$(Renglon)
            WFecha = WFecha
            WProducto = WProducto
            WTeorico = WTeorico
            WReal = WReal
            WFechaing = WFechaing
            WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
            WTipo = Tipo
            WArticulo = Articulo
            WTerminado = Terminado
            WCantidad = Cantidad
            WLote = "0"
            WWDate = Date$
            WImporte = ""
            WMarca = ""
            WSaldo = "0"
            WLote1 = "0"
            WLote2 = "0"
            WLote3 = "0"
            WCanti1 = "0"
            WCanti2 = "0"
            WCanti3 = "0"
            WCosto1 = "0"
            WCosto2 = "0"
            WCosto3 = "0"
                
            XParam = "'" + WClave + "','" _
                            + WHoja + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WProducto + "','" _
                            + WCantidad + "','" _
                            + WTipo + "','" _
                            + WLote + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WTeorico + "','" _
                            + WReal + "','" _
                            + WFechaing + "','" _
                            + WFechaingord + "','" _
                            + WDate + "','" _
                            + WImporte + "','" _
                            + WMarca + "','" _
                            + WSaldo + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + WLote3 + "','" + WLote3 + "','" _
                            + WCosto1 + "','" _
                            + WCosto2 + "','" _
                            + WCosto3 + "'"
                                           
            spHoja = "AltaHoja " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
            Auxiliar(Renglon, 1) = WProducto
            Auxiliar(Renglon, 2) = WTerminado
            Auxiliar(Renglon, 3) = WArticulo
            Auxiliar(Renglon, 4) = WCantidad
            Auxiliar(Renglon, 5) = WReal
            Auxiliar(Renglon, 6) = WTeorico
            Auxiliar(Renglon, 7) = WTipo
            Auxiliar(Renglon, 8) = WLote1
            Auxiliar(Renglon, 9) = WCanti1
            Auxiliar(Renglon, 10) = WLote2
            Auxiliar(Renglon, 11) = WCanti2
            Auxiliar(Renglon, 12) = WLote3
            Auxiliar(Renglon, 13) = WCanti3
                
        End If
                        
    Next iRow
    
    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    XParam = "'" + WHoja + "','" _
                 + WFechaord + "'"
    Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    For Da = 1 To Renglon
    
        WWProducto = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        ZCantidad = Val(Auxiliar(Da, 4))
        Real = Auxiliar(Da, 5)
        Teorico = Auxiliar(Da, 6)
        Tipo = Auxiliar(Da, 7)
        BajaLote(1, 1) = Auxiliar(Da, 8)
        BajaLote(1, 2) = Auxiliar(Da, 9)
        BajaLote(2, 1) = Auxiliar(Da, 10)
        BajaLote(2, 2) = Auxiliar(Da, 11)
        BajaLote(3, 1) = Auxiliar(Da, 12)
        BajaLote(3, 2) = Auxiliar(Da, 13)
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + WWProducto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                If Real <> 0 Then
                    WEntradas = Str$(rstTerminado!Entradas + Val(Real))
                    WProceso = Str$(rstTerminado!Proceso)
                        Else
                    WProceso = Str$(rstTerminado!Proceso + Val(Teorico))
                    WEntradas = Str$(rstTerminado!Entradas)
                End If
                WDate = Date$
                rstTerminado.Close
                    
                XParam = "'" + WCodigo + "','" _
                        + WEntradas + "','" _
                        + WProceso + "','" _
                        + WDate + "'"
                                           
                spTerminado = "ModificaTerminadoHoja " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
                
        Select Case Tipo
            Case "M"
            
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WSalidas = Str$(rstArticulo!Salidas + ZCantidad)
                    WDate = Date$
                    XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                    rstArticulo.Close
                                            
                    spArticulo = "ModificaArticuloSalidas " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                                            
            Case "T"
            
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WSalidas = Str$(rstTerminado!Salidas + ZCantidad)
                    WDate = Date$
                    rstTerminado.Close
                        
                    XParam = "'" + WCodigo + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                            
                    spTerminado = "ModificaTerminadoSalidas " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
            Case Else
        End Select
        
    Next Da
    
    Sql1 = "DELETE Prueter"
    Sql2 = " Where Lote = " + "'" + Hoja.Text + "'"
    spPrueter = Sql1 + Sql2
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Limpia_Click
    Hoja.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False
    
    WTipo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    MotivoDesvio.ListIndex = 0
    
    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False
    Erase XLote

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    FechaInicio.Text = "  /  /    "
    HoraInicio.Text = ""
    FechaFinal.Text = "  /  /    "
    HoraFinal.Text = ""
    NroPedido.Text = ""
    LoteColorante.Text = ""
    
    Call Limpia_Vector
    
    If LeeHoja <> "N" Then
        spHoja = "ListaHojaNumero"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveLast
                Hoja.Text = rstHoja!Hoja + 1
            End With
            rstHoja.Close
        End If
    End If
    
    Renglon = 0
    Rem by nan
    Graba.Enabled = False
  Rem  Anula.Enabled = False

    Hoja.SetFocus

End Sub

Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo.SetFocus
                    Else
                WTerminado.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTerminado.Text = UCase(WTerminado.Text)
        spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Rem WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            WCantidad.SetFocus
                Else
            WTerminado.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            WCantidad.SetFocus
                Else
            WArticulo.SetFocus
        End If
    End If
End Sub


Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WTipo.Text = "O" Then
            Call Alta_Vector
            Call Ingresa_Click
            WTipo.SetFocus
            Exit Sub
        End If
    
        Dife = Abs(Val(WCompara) - Val(WCantidad.Text))
        Porce = Abs(Val(WCompara) * 0.1)
        If Dife > Porce And Val(WCompara) <> 0 Then
            T$ = "Actualizacion de Hoja de Produccion"
            m$ = "Las cantidades utilizadas difieren un +-10% de las utulizadas originalmente. Desea continuar con la carga"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 7 Then
                WCantidad.SetFocus
                Exit Sub
            End If
        End If
    
        WCantidad.Text = Pusing("###,###.###", WCantidad.Text)
        CargaLote.Visible = True
        If WTipo.Text = "M" Then
            CargaLote.Caption = "Ingreso de Lote"
            Dada.Caption = "Lote"
                Else
            CargaLote.Caption = "Ingreso de Partida"
            Dada.Caption = "Partida"
        End If
        WLote1.Text = ""
        WCanti1.Text = ""
        WLote2.Text = ""
        WCanti2.Text = ""
        WLote3.Text = ""
        WCanti3.Text = ""
        WControl1.Locked = False
        WControl2.Locked = False
        WControl3.Locked = False
        WControl1.Text = ""
        WControl2.Text = ""
        WControl3.Text = ""
        WControl1.Locked = True
        WControl2.Locked = True
        WControl3.Locked = True
        
        If Val(XLote(Val(WLinea.Text), 1)) <> 0 Then
            WLote1.Text = XLote(Val(WLinea.Text), 1)
            WCanti1.Text = XLote(Val(WLinea.Text), 2)
            WControl1.Locked = False
            WControl1.Text = ""
            WControl1.Locked = True
        End If
        If Val(XLote(Val(WLinea.Text), 3)) <> 0 Then
            WLote2.Text = XLote(Val(WLinea.Text), 3)
            WCanti2.Text = XLote(Val(WLinea.Text), 4)
            WControl2.Locked = False
            WControl2.Text = ""
            WControl2.Locked = True
        End If
        If Val(XLote(Val(WLinea.Text), 5)) <> 0 Then
            WLote3.Text = XLote(Val(WLinea.Text), 5)
            WCanti3.Text = XLote(Val(WLinea.Text), 6)
            WControl3.Locked = False
            WControl3.Text = ""
            WControl3.Locked = True
        End If
        WLote1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WTipo.Text = "O" Then
            Call Alta_Vector
            Call Ingresa_Click
            WTipo.SetFocus
            Exit Sub
        End If
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + Str$(Int(Val(WLote1.Text))) + "','" _
                             + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + Str$(Int(Val(WLote1.Text))) + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Verifica_Lote
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        WSaldo1 = 0
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo1 >= Val(WCanti1.Text) Or WControla > 0 Then
            WCanti1.Text = Pusing("###,###.###", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            If WTipo.Text = "M" Then
                m$ = WArticulo.Text + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WTerminado.Text + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo2 = 0
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo2 >= Val(WCanti2.Text) Or WControla > 0 Then
            WCanti2.Text = Pusing("###,###.###", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
            WLote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            If WTipo.Text = "M" Then
                m$ = WArticulo.Text + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WTerminado.Text + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo3 = 0
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo3 >= Val(WCanti3.Text) Or WControla > 0 Then
            WCanti3.Text = Pusing("###,###.###", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
            Call Verifica_Lote
            If WEstado = "S" Then
                Call Alta_Vector
                Call Ingresa_Click
                WTipo.SetFocus
            End If
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            If WTipo.Text = "M" Then
                m$ = WArticulo.Text + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WTerminado.Text + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote3.SetFocus
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
            Claveven$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            Call Alta_Vector
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WTerminado.Text = rstTerminado!Codigo
                Rem WDescripcion.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            Call Alta_Vector
            
        Case Else
    End Select
    
    Call Indica
    
End Sub

Sub Indica()

    Select Case XIndice
        Case 0
            Producto.SetFocus
        Case 1, 2
        Case Else
    End Select

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    MotivoDesvio.Clear
    
    MotivoDesvio.AddItem ""
    MotivoDesvio.AddItem "Formula incorrecta"
    MotivoDesvio.AddItem "Cambio de Proceso"
    MotivoDesvio.AddItem "Varios"
    
    MotivoDesvio.ListIndex = 0

    Erase XLote
    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    LoteColorante.Text = ""
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    FechaInicio.Text = "  /  /    "
    HoraInicio.Text = ""
    FechaFinal.Text = "  /  /    "
    HoraFinal.Text = ""
    NroPedido.Text = ""
    
    If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
        LabelPedido.Visible = True
        NroPedido.Visible = True
    End If
    
    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            Hoja.Text = rstHoja!Hoja + 1
        End With
        rstHoja.Close
    End If
    
    WE = Wempresa
    
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgHojaProduccion.Caption = "Ingreso de Hoja de Produccion :  " + !Nombre
        End If
    End With
    
    Graba.Enabled = True
    Anula.Enabled = False
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    Erase Auxiliar
    Erase XLote
    WSaldoant = 0
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
                    Vector.Row = Renglon
                
                    Vector.Col = 1
                    Vector.Text = rstHoja!Tipo
                    
                    Vector.Col = 2
                    Vector.Text = rstHoja!Terminado
                    Auxi1 = rstHoja!Terminado
                
                    Vector.Col = 3
                    Vector.Text = rstHoja!Articulo
                    Auxi2 = rstHoja!Articulo
                
                    Vector.Col = 5
                    Vector.Text = Pusing("###,###.###", rstHoja!Cantidad)
                
                    Auxiliar(Renglon, 1) = rstHoja!Tipo
                    Auxiliar(Renglon, 2) = Auxi1
                    Auxiliar(Renglon, 3) = Auxi2
                    
                    XLote(Renglon, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(Renglon, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    XLote(Renglon, 3) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(Renglon, 4) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    XLote(Renglon, 5) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(Renglon, 6) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    XLote(Renglon, 7) = ""
                    
                    If Val(Real.Text) <> 0 Then
                        If Val(XLote(Renglon, 1)) = 0 And rstHoja!Lote <> 0 Then
                            XLote(Renglon, 1) = rstHoja!Lote
                            XLote(Renglon, 2) = rstHoja!Cantidad
                        End If
                    End If
                    
                    Rem If Val(XLote(Renglon, 2)) <> 0 Then
                    Rem     XLote(Renglon, 2) = Pusing("###,###.##", XLote(Renglon, 2))
                    Rem End If
                    
                    Rem If Val(XLote(Renglon, 4)) <> 0 Then
                    Rem     XLote(Renglon, 4) = Pusing("###,###.##", XLote(Renglon, 4))
                    Rem End If
                    
                    Rem If Val(XLote(Renglon, 6)) <> 0 Then
                    Rem     XLote(Renglon, 6) = Pusing("###,###.##", XLote(Renglon, 6))
                    Rem End If
                    
                    If rstHoja!Renglon = 1 Then
                        WSaldoant = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        Vector.Row = Renglon
        
        Tipo = Auxiliar(Renglon, 1)
        Auxi1 = Auxiliar(Renglon, 2)
        Auxi2 = Auxiliar(Renglon, 3)
                
        Select Case Tipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Vector.Col = 4
                    Rem Vector.Text = rstTerminado!Descripcion
                    Vector.Text = ""
                    rstTerminado.Close
                    WArticulo.SetFocus
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Vector.Col = 4
                    Vector.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                    WArticulo.SetFocus
                End If
            Case Else
        End Select
    Next Da
    
    Vector.TopRow = 1
    Vector.Row = 1
    Vector.Col = 1
    
    Real.Text = Pusing("###,###.##", Real.Text)

    If Val(Real.Text) <> 0 And Val(Real.Text) <> WSaldoant Then
      Rem by nan
       Graba.Enabled = False
        Anula.Enabled = False
            Else
        If Val(Real.Text) = 0 Then
            Graba.Enabled = True
            Anula.Enabled = False
            WExiste = "N"
                Else
            Graba.Enabled = False
            Anula.Enabled = False
            Rem Anula.Enabled = True
            WExiste = "S"
        End If
    End If
    
    Rem Renglon = Renglon - 1
    
    WTipo.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

        Renglon = Renglon + 1
        Ultimo = Renglon + 1
            
        Vector.Row = Renglon
            
        Vector.Col = 1
        Vector.Text = WTipo.Text
            
        Vector.Col = 2
        Vector.Text = WTerminado.Text
            
        Vector.Col = 3
        Vector.Text = WArticulo.Text
            
        Vector.Col = 4
        Vector.Text = WDescripcion.Caption
                
        Vector.Col = 5
        Vector.Text = Pusing("###,###.###", WCantidad.Text)
            
        Vector.Col = 6
        Vector.Text = "S"
            
        XLote(Renglon, 1) = WLote1.Text
        XLote(Renglon, 2) = WCanti1.Text
        XLote(Renglon, 3) = WLote2.Text
        XLote(Renglon, 4) = WCanti2.Text
        XLote(Renglon, 5) = WLote3.Text
        XLote(Renglon, 6) = WCanti3.Text
            
                Else
                
        WRen = Val(WLinea.Text)
        Vector.Row = WRen
            
        Vector.Col = 1
        Vector.Text = WTipo.Text
            
        Vector.Col = 2
        Vector.Text = WTerminado.Text
            
        Vector.Col = 3
        Vector.Text = WArticulo.Text
            
        Vector.Col = 4
        Vector.Text = WDescripcion.Caption
                
        Vector.Col = 5
        Vector.Text = Pusing("###,###.###", WCantidad.Text)
            
        Vector.Col = 6
        Vector.Text = "S"
            
        XLote(WRen, 1) = WLote1.Text
        XLote(WRen, 2) = WCanti1.Text
        XLote(WRen, 3) = WLote2.Text
        XLote(WRen, 4) = WCanti2.Text
        XLote(WRen, 5) = WLote3.Text
        XLote(WRen, 6) = WCanti3.Text
            
    End If
    
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entra = "N"
        spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            Entra = "S"
            Fecha.Text = rstHoja!Fecha
            Real.Text = Str$(rstHoja!Real)
            Teorico.Text = Str$(rstHoja!Teorico)
            fechaIng.Text = rstHoja!fechaIng
            Producto.Text = rstHoja!Producto
            Equipo.Text = IIf(IsNull(rstHoja!Equipo), "", rstHoja!Equipo)
            NroPedido.Text = IIf(IsNull(rstHoja!NroPedido), "", rstHoja!NroPedido)
            VersionI.Text = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
            VersionII.Text = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
            VersionIII.Text = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
            FechaFinal.Text = IIf(IsNull(rstHoja!FechaFinal), "  /  /    ", rstHoja!FechaFinal)
            HoraFinal.Text = IIf(IsNull(rstHoja!HoraFinal), "", rstHoja!HoraFinal)
            FechaInicio.Text = IIf(IsNull(rstHoja!FechaInicio), "  /  /    ", rstHoja!FechaInicio)
            HoraInicio.Text = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
            HoraInicio.Text = Trim(HoraInicio.Text)
            HoraFinal.Text = Trim(HoraFinal.Text)
            rstHoja.Close
        End If
        
        If Entra = "S" Then
        
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                Rem DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            Call Proceso_Click
                
                Else
                    
            Existe = "N"
            Hoja.SetFocus
                
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Teorico_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Teorico.Text = Pusing("###,###.##", Teorico.Text)
        Real.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Real_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Real.Text = Pusing("###,###.##", Real.Text)
        fechaIng.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fechaing_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(fechaIng.Text, Auxi)
        If Auxi = "S" Or fechaIng.Text = "  /  /    " Then
            If Existe = "N" Then
                Call Lee_Composicion
            End If
            FechaInicio.SetFocus
                Else
            fechaIng.SetFocus
        End If
    End If
End Sub

Private Sub fechaIng_GotFocus()
    If fechaIng.Text = "  /  /    " Then
        fechaIng.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    End If
End Sub

Private Sub FechaInicio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaInicio.Text, Auxi)
        If Auxi = "S" Then
            HoraInicio.SetFocus
                Else
            FechaInicio.SetFocus
        End If
    End If
End Sub

Private Sub HoraInicio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(HoraInicio.Text) > 0 And Val(HoraInicio.Text) <= 24 Then
            FechaFinal.SetFocus
                Else
            HoraInicio.SetFocus
        End If
    End If
End Sub

Private Sub FechaFinal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaFinal.Text, Auxi)
        If Auxi = "S" Then
            HoraFinal.SetFocus
                Else
            FechaFinal.SetFocus
        End If
    End If
End Sub

Private Sub HoraFinal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(HoraFinal.Text) > 0 And Val(HoraFinal.Text) <= 24 Then
            WTipo.SetFocus
                Else
            HoraFinal.SetFocus
        End If
    End If
End Sub

Private Sub Lee_Composicion()

    Erase Auxiliar
    Renglon = 0
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Producto.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
        
                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
        
                        Renglon = Renglon + 1
                        Vector.Row = Renglon
                    
                        Vector.Col = 1
                        Vector.Text = rstComposicion!Tipo
                    
                        If rstComposicion!Articulo1 = "  -   -  " Then
                            Vector.Col = 3
                            Vector.Text = "  -   -   "
                            Auxi1 = "  -   -   "
                                Else
                            Vector.Col = 3
                            Vector.Text = rstComposicion!Articulo1
                            Auxi1 = rstComposicion!Articulo1
                        End If
                    
                        Vector.Col = 2
                        Vector.Text = rstComposicion!Articulo2
                        Auxi2 = rstComposicion!Articulo2
                    
                        Cantidad = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                    
                        Vector.Col = 5
                        Vector.Text = Pusing("###,###.###", Cantidad)
                    
                        Vector.Col = 6
                        Vector.Text = ""
                        
                        Auxiliar(Renglon, 1) = rstComposicion!Tipo
                        Auxiliar(Renglon, 2) = Auxi1
                        Auxiliar(Renglon, 3) = Auxi2
                        Auxiliar(Renglon, 4) = Cantidad
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
                    
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        Vector.Row = Renglon
        
        Tipo = Auxiliar(Renglon, 1)
        Auxi2 = Auxiliar(Renglon, 2)
        Auxi1 = Auxiliar(Renglon, 3)
        XCantidad = Val(Auxiliar(Renglon, 4))
        
        WStock = 0
                
        Select Case Tipo
            Case "T"
                WImpre1 = Auxi1
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Vector.Col = 4
                    Rem Vector.Text = rstTerminado!Descripcion
                    Vector.Text = ""
                    WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                    WArticulo.SetFocus
                End If
            Case "M"
                WImpre1 = Auxi2
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Vector.Col = 4
                    Vector.Text = rstArticulo!Descripcion
                    WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                    WArticulo.SetFocus
                End If
            Case Else
        End Select
        
        Vector.Col = 5
        Vector.Text = Pusing("###,###.###", Str$(XCantidad))
        
    Next Da
    
End Sub

Sub Impresion()

        Open "lpt1" For Output As #1

        Print #1, Chr$(27) + Chr$(71)
        Print #1,
        Print #1, Chr$(18)

        Print #1, Tab(15); Left$(Producto.Text, 2)
        Select Case Val(Wempresa)
            Case 1
                Print #1, Tab(70); "SI"
            Case 2
                Print #1, Tab(70); "PI"
            Case 3
                Print #1, Tab(70); "SII"
            Case 4
                Print #1, Tab(70); "PII"
            Case 5
                Print #1, Tab(70); "SIII"
            Case 6
                Print #1, Tab(70); "SIV"
            Case 7
                Print #1, Tab(70); "SV"
            Case 8
                Print #1, Tab(70); "PV"
            Case 9
                Print #1, Tab(70); "PVI"
            Case 10
                Print #1, Tab(70); "SVI"
            Case 11
                Print #1, Tab(70); "SVII"
            Case Else
        End Select

        Print #1, Tab(1); Fecha.Text;
        Print #1, Tab(13); Alinea("#####", Mid$(Producto.Text, 4, 5));
        Print #1, "/"; Right$(Producto.Text, 3);
        Print #1, Tab(26); Chr$(14); Alinea("######", Hoja.Text)

        Print #1,
        Print #1,
        Print #1,

        Linea = 0
        
        For iRow = 1 To 40
                
            WRow = iRow
            Tipo = Vector.TextMatrix(WRow, 1)
            Terminado = UCase(Vector.TextMatrix(WRow, 2))
            Articulo = UCase(Vector.TextMatrix(WRow, 3))
            Cantidad = Vector.TextMatrix(WRow, 5)
            Lote = Vector.TextMatrix(WRow, 6)
                    
            If Tipo = "M" Then

                Linea = Linea + 1

                Print #1, Tab(6); Left$(Articulo, 2);
                Print #1, Tab(11); Mid$(Articulo, 4, 3);
                Print #1, "-";
                Print #1, Right$(Articulo, 3);
                Print #1, Tab(33); Alinea("####.#", Str$(Cantidad))
                Print #1,

            End If

            If Tipo = "T" Then

                Linea = Linea + 1

                Print #1, Tab(6); Left$(Terminado, 2);
                Print #1, Tab(11); Mid$(Terminado, 4, 5);
                Print #1, "-";
                Print #1, Right$(Terminado, 3);
                Print #1, Tab(33); Alinea("####.#", Str$(Cantidad))
                Print #1,

            End If
                    
        Next iRow

        For Ciclo = Linea To 14

                Print #1,
                Print #1,

        Next Ciclo

        Print #1, Tab(33); Alinea("####.#", Teorico.Text)

        Print #1,
        Print #1, Chr$(27) + Chr$(72)
        Print #1, Chr$(12)
        
        Close #1

End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True

    
    WSaldo1 = 0
    WSaldo2 = 0
    WSaldo3 = 0
    
    If Val(WLote1.Text) <> 0 Then
        If WTipo.Text = "M" Then
        
            WEntra = "N"
            WMarcaEstado = ""
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + WLote1.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                ZZArticulo = WArticulo.Text
                ZZLote = WLote1.Text
                Call Verifica_Vencido
                If ZZMarcaVencida = "S" Then
                    m$ = WArticulo.Text + " Partida Vencida : " + WLote1.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo1 = 0
                End If
                If WMarcaEstado = "N" Then
                    m$ = WArticulo.Text + " Partida Bloqueada : " + WLote1.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo1 = 0
                End If
            End If
            
                Else
        
            WEntra = "N"
            WMarcaEstado = ""
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                If WMarcaEstado = "N" Then
                    m$ = WTerminado.Text + " Partida Bloqueada : " + WLote1.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo1 = 0
                End If
            End If
            
            
        End If
        
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.###", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
        End If
        
    End If
    
    If Val(WLote2.Text) <> 0 Then
        If WTipo.Text = "M" Then
        
            WEntra = "N"
            WMarcaEstado = ""
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                ZZArticulo = WArticulo.Text
                ZZLote = WLote2.Text
                Call Verifica_Vencido
                If WMarcaVencida = "S" Then
                    m$ = WArticulo.Text + " Partida Vencida : " + WLote2.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo2 = 0
                End If
                If WMarcaEstado = "N" Then
                    m$ = WArticulo.Text + " Partida Bloqueada : " + WLote2.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo2 = 0
                End If
            End If
            
                Else
        
            WEntra = "N"
            WMarcaEstado = ""
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                If WMarcaEstado = "N" Then
                    m$ = WTerminado.Text + " Partida Bloqueada : " + WLote2.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo2 = 0
                End If
            End If
            
        End If
            
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.###", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
        End If
        
    End If
    
    
    If Val(WLote3.Text) <> 0 Then
        If WTipo.Text = "M" Then
        
            WEntra = "N"
            WMarcaEstado = ""
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                ZZArticulo = WArticulo.Text
                ZZLote = WLote3.Text
                Call Verifica_Vencido
                If WMarcaVencida = "S" Then
                    m$ = WArticulo.Text + " Partida Vencida : " + WLote3.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo3 = 0
                End If
                If WMarcaEstado = "N" Then
                    m$ = WArticulo.Text + " Partida Bloqueada : " + WLote3.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo3 = 0
                End If
            End If
            
                Else
        
            WEntra = "N"
            WMarcaEstado = ""
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WMarcaEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If Left$(Producto.Text, 2) <> "RE" And Left$(Producto.Text, 2) <> "NK" Then
                If WMarcaEstado = "N" Then
                    m$ = WTerminado.Text + " Partida Bloqueada : " + WLote3.Text + Chr$(13) + "No se puede informar en una hoja de produccion " + Chr$(13) + "Solo es posible informarla en un producto RE o NK"
                    G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    WSaldo3 = 0
                End If
            End If
            
        End If
        
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.###", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
        End If
        
    End If
    
    If Val(WLote1.Text) <> 0 And WControl1.Text = "X" Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If Val(WLote2.Text) <> 0 And WControl2.Text = "X" Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Val(WLote3.Text) <> 0 And WControl3.Text = "X" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    
    If Suma = Val(WCantidad.Text) Then
        WEstado = "S"
    End If
    
    If WControla <> 0 Then
        WEstado = "S"
    End If
    
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    XClave.Visible = False
    Hoja.SetFocus

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave.Text = "V4589" Then
            WExiste = "N"
            XClave.Visible = False
            Call Anula_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            WClave.SetFocus
        End If
    End If

End Sub

Private Sub Calcula_Costo_Produccion(ZProducto As String, ZCosto1 As Double, ZCosto2 As Double, ZCosto3 As Double)

    Dim ZVector(100, 2) As String
    Erase ZAuxiliar
    ZRenglon = 0
    
    ZVector(1, 1) = ZProducto
    ZVector(1, 2) = "1"
    ZCosto1 = 0
    ZCosto2 = 0
    ZCosto3 = 0
    ZLugar = 1
    ZCicla = 0
    
    Do
        ZCicla = ZCicla + 1
        If ZVector(ZCicla, 1) <> "" Then
    
            ZEntra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + ZVector(ZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZEntra = "N"
                        
                        ZTipo = rstComposicion!Tipo
                        ZArticulo1 = rstComposicion!Articulo1
                        ZArticulo2 = rstComposicion!Articulo2
                        ZCantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(ZArticulo1, 2) = "DW" Then
                        Rem     ZTipo = "T"
                        Rem     ZArticulo2 = Left$(ZArticulo1, 3) + "00" + Right$(ZArticulo1, 7)
                        Rem End If
                        
                        Select Case ZTipo
                            Case "T"
                                If ZProducto <> ZArticulo2 Then
                                    ZLugar = ZLugar + 1
                                    ZVector(ZLugar, 1) = ZArticulo2
                                    ZVector(ZLugar, 2) = Str$(ZCantidad * Val(ZVector(ZCicla, 2)))
                                End If
                            Case "M"
                                ZRenglon = ZRenglon + 1
                                ZAuxiliar(ZRenglon, 1) = ZArticulo1
                                ZAuxiliar(ZRenglon, 2) = ZCantidad
                                ZAuxiliar(ZRenglon, 3) = ZVector(ZCicla, 2)
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
            
            Rem If ZEntra = "S" And Left$(ZVector(ZCicla, 1), 2) = "DW" Then
            Rem     ZRenglon = ZRenglon + 1
            Rem     ZAuxiliar(ZRenglon, 1) = Left$(ZVector(ZCicla, 1), 3) + Right$(ZVector(ZCicla, 1), 7)
            Rem     ZAuxiliar(ZRenglon, 2) = 1
            Rem     ZAuxiliar(ZRenglon, 3) = ZVector(ZCicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
    
    ZCosto1 = 0
    ZCosto2 = 0
    ZCosto3 = 0
                    
    For ZDa = 1 To ZRenglon
        ZArticulo = ZAuxiliar(ZDa, 1)
        ZCantidad = ZAuxiliar(ZDa, 2)
        ZWVector = ZAuxiliar(ZDa, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCos1 = (ZCantidad * rstArticulo!Costo2 * Val(ZWVector))
            ZCosto1 = ZCosto1 + WCos1
            WCos2 = (ZCantidad * rstArticulo!Costo1 * Val(ZWVector))
            ZCosto2 = ZCosto2 + WCos2
            WCos3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
            WCos3 = (ZCantidad * WCos3 * Val(ZWVector))
            ZCosto3 = ZCosto3 + WCos3
            rstArticulo.Close
        End If
    Next ZDa
    
    Call Redondeo(XCosto1)
    Call Redondeo(XCosto2)
    Call Redondeo(XCosto3)
    
End Sub

Private Sub Limpia_Vector()

    Vector.Clear
    Vector.Font.Bold = True

    Vector.FixedCols = 1
    Vector.Cols = 7
    Vector.FixedRows = 1
    Vector.Rows = 41
    
    Vector.ColWidth(0) = 200
    Vector.Row = 0
    For Ciclo = 1 To Vector.Cols - 1
        Vector.Col = Ciclo
        Select Case Ciclo
            Case 1
                Vector.Text = "Tipo"
                Vector.ColWidth(Ciclo) = 550
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Vector.Text = "Prod.Terminado"
                Vector.ColWidth(Ciclo) = 1600
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Vector.Text = "Materia Prima"
                Vector.ColWidth(Ciclo) = 1400
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Vector.Text = "Descripcion"
                Vector.ColWidth(Ciclo) = 3600
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Vector.Text = "Cantidad"
                Vector.ColWidth(Ciclo) = 1100
                Vector.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                Vector.Text = "OK"
                Vector.ColWidth(Ciclo) = 500
                Vector.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    Vector.Row = 0
    For Ciclo = 1 To Vector.Cols - 1
        Vector.Col = Ciclo
        WTitulo(Ciclo).Text = Vector.Text
        WTitulo(Ciclo).Left = Vector.CellLeft + Vector.Left
        WTitulo(Ciclo).Top = Vector.CellTop + Vector.Top
        WTitulo(Ciclo).Width = Vector.CellWidth
        WTitulo(Ciclo).Height = Vector.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To Vector.Cols - 1
        WAncho = WAncho + Vector.ColWidth(Ciclo)
    Next Ciclo
    Vector.Width = WAncho

    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Vector.AllowUserResizing = flexResizeBoth
    
    Vector.Col = 1
    Vector.Row = 1
    
End Sub

Private Sub Actualiza_Hoja()

    WTipomov = "9"
    WDestino = "9"
    
    spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveLast
            Do
                WCodigo = Str$(rstMovguia!Codigo + 1)
                If Val(WCodigo) > 900000 Then
                    .MovePrevious
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
            Else
        WCodigo = "1"
    End If
    
    
    Tipo = "T"
    Terminado = Producto.Text
    Articulo = "  -   -   "
    Cantidad = Real.Text
    Movi = "S"
    Lote = Hoja.Text
    Transito = ""
    Orden = ""
    Descontar = ""
    
    Auxi1 = WCodigo
    Call Ceros(Auxi1, 6)
    Auxi = "01"
    
    WTipomov = WTipomov
    WDestino = WDestino
    WCodigo = WCodigo
    WRenglon = "1"
    WFecha = Fecha.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WTipo = Tipo
    WArticulo = Articulo
    WTerminado = Terminado
    WCantidad = Cantidad
    WMovi = Movi
    WObservaciones = "Hoja de Produccion Nro`. " + ZLote
    WClave = WTipomov + Auxi1 + Auxi
    WDate = Date$
    WMarca = ""
    WPartida = Lote
    WLote = ""
    WSaldo = "0"
    WPartiOri = Lote
    WTransito = Transito
    WOrden = Orden
    WDescontar = Descontar
    
    Sql1 = "INSERT INTO Guia ("
    Sql2 = "Clave ,"
    Sql3 = "TipoMov ,"
    Sql4 = "Codigo ,"
    Sql5 = "Renglon ,"
    Sql6 = "Fecha ,"
    Sql7 = "Tipo ,"
    Sql8 = "Articulo ,"
    Sql9 = "Terminado ,"
    Sql10 = "Cantidad ,"
    Sql11 = "FechaOrd ,"
    Sql12 = "Movi,"
    Sql13 = "Observaciones,"
    Sql14 = "Marca,"
    Sql15 = "Destino,"
    Sql16 = "Lote,"
    Sql17 = "Saldo,"
    Sql18 = "Partida,"
    Sql19 = "PartiOri,"
    Sql20 = "Transito,"
    Sql21 = "Orden,"
    Sql22 = "Descontar )"
    Sql23 = "Values ("
    Sql24 = "'" + WClave + "',"
    Sql25 = "'" + WTipomov + "',"
    Sql26 = "'" + WCodigo + "',"
    Sql27 = "'" + WRenglon + "',"
    Sql28 = "'" + WFecha + "',"
    Sql29 = "'" + WTipo + "',"
    Sql30 = "'" + WArticulo + "',"
    Sql31 = "'" + WTerminado + "',"
    Sql32 = "'" + WCantidad + "',"
    Sql33 = "'" + WFechaord + "',"
    Sql34 = "'" + WMovi + "',"
    Sql35 = "'" + WObservaciones + "',"
    Sql36 = "'" + WMarca + "',"
    Sql37 = "'" + WDestino + "',"
    Sql38 = "'" + WLote + "',"
    Sql39 = "'" + WSaldo + "',"
    Sql40 = "'" + WPartida + "',"
    Sql41 = "'" + WPartiOri + "',"
    Sql42 = "'" + WTransito + "',"
    Sql43 = "'" + WOrden + "',"
    Sql44 = "'" + WDescontar + "')"
        
    spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                Sql41 + Sql42 + Sql43 + Sql44
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                
    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        If Movi = "E" Then
            WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad))
            WSalidas = Str$(rstTerminado!Salidas)
                Else
            WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
            WEntradas = Str$(rstTerminado!Entradas)
        End If
        WDate = Date$
        rstTerminado.Close
                
        XParam = "'" + Terminado + "','" _
                     + WEntradas + "','" _
                     + WSalidas + "','" _
                     + WDate + "'"
                                           
        spTerminado = "ModificaTerminadoMovimientos " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
        XParam = "'" + Lote + "','" _
                     + Terminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WClave = rstHoja!Clave
            If Movi = "E" Then
                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                    Else
                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
            End If
            WDate = Date$
            rstHoja.Close
                            
            XParam = "'" + WClave + "','" _
                         + WDate + "','" _
                         + WSaldo + "'"
            spHoja = "ModificahojaSaldo " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        End If
                    
    End If
                
    Tipo = "M"
    Terminado = "  -     -   "
    Articulo = Left$(Producto.Text, 2) + "-" + Right$(Producto.Text, 7)
    Cantidad = Real.Text
    Movi = "E"
    Lote = Hoja.Text
    Transito = ""
    Orden = ""
    Descontar = ""
    
    Auxi1 = WCodigo
    Call Ceros(Auxi1, 6)
    Auxi = "02"
    
    WTipomov = WTipomov
    WDestino = WDestino
    WCodigo = WCodigo
    WRenglon = "2"
    WFecha = Fecha.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WTipo = Tipo
    WArticulo = Articulo
    WTerminado = Terminado
    WCantidad = Cantidad
    WMovi = Movi
    WObservaciones = "Hoja de Produccion Nro. " + Lote
    WClave = WTipomov + Auxi1 + Auxi
    WDate = Date$
    WMarca = ""
    WPartida = ""
    WLote = Lote
    WSaldo = Real.Text
    WPartiOri = Lote
    WTransito = Transito
    WOrden = Orden
    WDescontar = Descontar
    
    Sql1 = "INSERT INTO Guia ("
    Sql2 = "Clave ,"
    Sql3 = "TipoMov ,"
    Sql4 = "Codigo ,"
    Sql5 = "Renglon ,"
    Sql6 = "Fecha ,"
    Sql7 = "Tipo ,"
    Sql8 = "Articulo ,"
    Sql9 = "Terminado ,"
    Sql10 = "Cantidad ,"
    Sql11 = "FechaOrd ,"
    Sql12 = "Movi,"
    Sql13 = "Observaciones,"
    Sql14 = "Marca,"
    Sql15 = "Destino,"
    Sql16 = "Lote,"
    Sql17 = "Saldo,"
    Sql18 = "Partida,"
    Sql19 = "PartiOri,"
    Sql20 = "Transito,"
    Sql21 = "Orden,"
    Sql22 = "Descontar )"
    Sql23 = "Values ("
    Sql24 = "'" + WClave + "',"
    Sql25 = "'" + WTipomov + "',"
    Sql26 = "'" + WCodigo + "',"
    Sql27 = "'" + WRenglon + "',"
    Sql28 = "'" + WFecha + "',"
    Sql29 = "'" + WTipo + "',"
    Sql30 = "'" + WArticulo + "',"
    Sql31 = "'" + WTerminado + "',"
    Sql32 = "'" + WCantidad + "',"
    Sql33 = "'" + WFechaord + "',"
    Sql34 = "'" + WMovi + "',"
    Sql35 = "'" + WObservaciones + "',"
    Sql36 = "'" + WMarca + "',"
    Sql37 = "'" + WDestino + "',"
    Sql38 = "'" + WLote + "',"
    Sql39 = "'" + WSaldo + "',"
    Sql40 = "'" + WPartida + "',"
    Sql41 = "'" + WPartiOri + "',"
    Sql42 = "'" + WTransito + "',"
    Sql43 = "'" + WOrden + "',"
    Sql44 = "'" + WDescontar + "')"
        
    spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                Sql41 + Sql42 + Sql43 + Sql44
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        If WMovi = "E" Then
            WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
            WSalidas = Str$(rstArticulo!Salidas)
                Else
            WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
            WEntradas = Str$(rstArticulo!Entradas)
        End If
        WDate = Date$
                    
        XParam = "'" + WArticulo + "','" _
                     + WEntradas + "','" _
                     + WSalidas + "','" _
                     + WDate + "'"
                                           
        spArticulo = "ModificaArticuloMovimientos " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
End Sub


Private Sub Verifica_Vencido()
        
    ZLaudo = ZZLote
    ZArticulo = ZZArticulo
        
    XEmpresa = Wempresa
    
    Empe(1, 1) = "0003"
    Empe(1, 2) = "Empresa03"
    Empe(2, 1) = "0008"
    Empe(2, 2) = "Empresa08"
    Empe(3, 1) = "0007"
    Empe(3, 2) = "Empresa07"
    Empe(4, 1) = "0004"
    Empe(4, 2) = "Empresa04"
    Empe(5, 1) = "0005"
    Empe(5, 2) = "Empresa05"
    Empe(6, 1) = "0001"
    Empe(6, 2) = "Empresa01"
    Empe(7, 1) = "0002"
    Empe(7, 2) = "Empresa02"
    Empe(8, 1) = "0006"
    Empe(8, 2) = "Empresa06"
    Empe(9, 1) = "0009"
    Empe(9, 2) = "Empresa09"
    Empe(10, 1) = "0010"
    Empe(10, 2) = "Empresa10"
    Empe(11, 1) = "0011"
    Empe(11, 2) = "Empresa11"
    
    XHasta = 11
    
    For Ciclo2 = 1 To XHasta
    
        Wempresa = Empe(Ciclo2, 1)
        txtOdbc = Empe(Ciclo2, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
        ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            ZFecha = rstLaudo!Fecha
            ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
            rstLaudo.Close
            Exit For
        End If
            
    Next Ciclo2
            
    Call Conecta_Empresa
    
    ZVto = ""
    ZZMarcaVencida = ""
            
    ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
    If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
        Call Valida_fecha(ZFechaVto, Auxi)
        If Auxi = "S" Then
            ZVto = ZFechaVto
        End If
    End If
            
    If ZVto = "" Then
            
        ZMeses = 0
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZMeses = rstArticulo!Meses
            rstArticulo.Close
        End If
            
        WMes = Val(Mid$(ZFecha, 4, 2))
        WAno = Val(Right$(ZFecha, 4))
        For ZCiclo = 1 To ZMeses
            WMes = WMes + 1
            If WMes > 12 Then
                WAno = WAno + 1
                WMes = 1
            End If
        Next ZCiclo
            
        XMes = Str$(WMes)
        XAno = Str$(WAno)
        Call Ceros(XMes, 2)
        Call Ceros(XAno, 4)
        If Val(Left$(ZFecha, 2)) <= 30 Then
            If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
                ZVto = "28/" + XMes + "/" + XAno
                    Else
                ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
            End If
                Else
            If Val(XMes) = 2 Then
                ZVto = "28/" + XMes + "/" + XAno
                    Else
                ZVto = "30/" + XMes + "/" + XAno
            End If
        End If
           
    End If
        
    If ZFecha <> "" Then
        
        Do
            Call Valida_fecha(ZVto, Auxi)
            If Auxi = "S" Then
                Exit Do
                    Else
                XFec1 = ZVto
                SumaDia = 1
                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                ZVto = XFec2
            End If
        Loop
        
        Rem WFechaActual = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem WFechaVto = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        ZZComparaI = Fecha.Text
        Rem ZZComparaII = ZVto
        If Left$(ZVto, 2) > "28" Then
            ZZComparaII = "28" + Mid$(ZVto, 3, 8)
                Else
            ZZComparaII = ZVto
        End If
        
        ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
        
        If Val(ZDias) < 0 Then
            ZZMarcaVencida = "S"
        End If
        
    End If
    
End Sub
                


