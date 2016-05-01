VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgmodhojaOLD 
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
   Begin VB.CommandButton VerControl 
      Caption         =   "Muestra Ensayos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   109
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame IngresaEnsayo 
      Height          =   2055
      Left            =   9480
      TabIndex        =   55
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
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
         Left            =   7800
         TabIndex        =   108
         Top             =   3600
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
         TabIndex        =   107
         Text            =   " "
         Top             =   3960
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
         TabIndex        =   106
         Text            =   " "
         Top             =   4320
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   67
         Text            =   " "
         Top             =   720
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   66
         Text            =   " "
         Top             =   960
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   65
         Text            =   " "
         Top             =   1200
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   1440
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   1680
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   1920
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   61
         Text            =   " "
         Top             =   2160
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   60
         Text            =   " "
         Top             =   2400
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   59
         Text            =   " "
         Top             =   2640
         Width           =   3975
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
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   58
         Text            =   " "
         Top             =   2880
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
         TabIndex        =   57
         Text            =   " "
         Top             =   3240
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
         TabIndex        =   56
         Text            =   " "
         Top             =   3600
         Width           =   3975
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
         TabIndex        =   105
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
         TabIndex        =   104
         Top             =   360
         Width           =   2655
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
         Left            =   3840
         TabIndex        =   103
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Descri1 
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
         Left            =   1080
         TabIndex        =   102
         Top             =   720
         Width           =   2700
      End
      Begin VB.Label descri2 
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
         Left            =   1080
         TabIndex        =   101
         Top             =   960
         Width           =   2700
      End
      Begin VB.Label Descri3 
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
         Left            =   1080
         TabIndex        =   100
         Top             =   1200
         Width           =   2700
      End
      Begin VB.Label Descri4 
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
         Left            =   1080
         TabIndex        =   99
         Top             =   1440
         Width           =   2700
      End
      Begin VB.Label Descri5 
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
         Left            =   1080
         TabIndex        =   98
         Top             =   1680
         Width           =   2700
      End
      Begin VB.Label Descri6 
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
         Left            =   1080
         TabIndex        =   97
         Top             =   1920
         Width           =   2700
      End
      Begin VB.Label Descri7 
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
         Left            =   1080
         TabIndex        =   96
         Top             =   2160
         Width           =   2700
      End
      Begin VB.Label Descri8 
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
         Left            =   1080
         TabIndex        =   95
         Top             =   2400
         Width           =   2700
      End
      Begin VB.Label Descri9 
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
         Left            =   1080
         TabIndex        =   94
         Top             =   2640
         Width           =   2700
      End
      Begin VB.Label Descri10 
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
         Left            =   1080
         TabIndex        =   93
         Top             =   2880
         Width           =   2700
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
         TabIndex        =   92
         Top             =   3240
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
         TabIndex        =   91
         Top             =   3600
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
         TabIndex        =   90
         Top             =   3960
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
         TabIndex        =   89
         Top             =   4320
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
         Left            =   7200
         TabIndex        =   88
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Std1 
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
         Left            =   3840
         TabIndex        =   87
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Std2 
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
         Left            =   3840
         TabIndex        =   86
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Std3 
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
         Left            =   3840
         TabIndex        =   85
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Std4 
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
         Left            =   3840
         TabIndex        =   84
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Std5 
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
         Left            =   3840
         TabIndex        =   83
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Std6 
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
         Left            =   3840
         TabIndex        =   82
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Std7 
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
         Left            =   3840
         TabIndex        =   81
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Std8 
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
         Left            =   3840
         TabIndex        =   80
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Std9 
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
         Left            =   3840
         TabIndex        =   79
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Std10 
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
         Left            =   3840
         TabIndex        =   78
         Top             =   2880
         Width           =   3255
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
         TabIndex        =   77
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
         TabIndex        =   76
         Top             =   960
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
         TabIndex        =   75
         Top             =   1200
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
         TabIndex        =   74
         Top             =   1440
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
         TabIndex        =   73
         Top             =   1680
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
         TabIndex        =   72
         Top             =   1920
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
         TabIndex        =   71
         Top             =   2160
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
         TabIndex        =   70
         Top             =   2400
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
         TabIndex        =   69
         Top             =   2640
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
         TabIndex        =   68
         Top             =   2880
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   11160
      TabIndex        =   54
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"modhojaOLD.frx":0000
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
      TabIndex        =   53
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   52
      Top             =   7680
      Width           =   975
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
      TabIndex        =   51
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame XClave 
      Caption         =   "Ingrese de Clave de Seguridad"
      Height          =   1815
      Left            =   2760
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Grabacion"
         Height          =   375
         Left            =   600
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   9120
      TabIndex        =   35
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox fechaIng 
      Height          =   285
      Left            =   8760
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
      Left            =   9480
      Top             =   1560
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4560
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      Left            =   2280
      TabIndex        =   14
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
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
         TabIndex        =   29
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   28
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
         TabIndex        =   27
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   15
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   13
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   12
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
      TabIndex        =   10
      Top             =   7080
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "modhojaOLD.frx":00F7
      TabIndex        =   9
      Top             =   1320
      Width           =   9135
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10080
      TabIndex        =   8
      Top             =   1680
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
      ItemData        =   "modhojaOLD.frx":0ADD
      Left            =   3360
      List            =   "modhojaOLD.frx":0AE4
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
      Left            =   6960
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   21
      Top             =   480
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
      TabIndex        =   20
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
      Left            =   3600
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgmodhojaOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Tipo As String
Private Existe  As String
Private Auxi1 As String
Private Auxi2 As String
Private XIndice As Integer
Private WImpre As String
Private Cantidad As String
Private Auxiliar(100, 20) As String
Private ZAuxiliar(100, 7) As String
Private xLote(100, 7) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHoja As Recordset
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
Dim EmpresaActual As String
Dim Verifica(100, 10) As String
Dim WLugar As Integer
Dim WEntraLote As Integer
Dim WCicloLote As Integer
Dim XCicloLote As Integer
Dim Entra As String

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

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
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
    
    CargaLote.Visible = False
    
    WArticulo.SetFocus
    
End Sub


Private Sub cmdClose_Click()

    LeeHoja = "N"
    Call Limpia_Click
    LeeHoja = "S"
    
    With rstEmpresa
        .Close
    End With
    With rstEtiqueta
        .Close
    End With
    
    Prgmodhoja.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Materia Prima"
     Opcion.AddItem "Productos Terminados"

     Opcion.Visible = True
     
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
    OPEN_FILE_Etiqueta
End Sub


Private Sub Opcion_Click()

    Dim IngresaItem As String
    pantalla.Clear
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
                            pantalla.AddItem IngresaItem
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
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            pantalla.AddItem IngresaItem
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
            
    pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.Row + 1
        WTipo.Text = DBGrid1.Text
            Else
        WTipo.Text = ""
        WLinea.Text = ""
    End If

    DBGrid1.Col = 1
    If Len(DBGrid1.Text) = 12 Then
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    DBGrid1.Col = 2
    If Len(DBGrid1.Text) = 10 Then
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    DBGrid1.Col = 3
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 4
    WCantidad.Text = DBGrid1.Text
    
    If Val(Teorico.Text) = 0 Then
        Teorico.SetFocus
            Else
        WCantidad.SetFocus
    End If
        

End Sub

Private Sub Graba_Click()

    If WExiste = "S" Then
        Call Ingresa_clave
        Exit Sub
    End If
    
    General = "S"
    XSuma = 0
    WLugar = 0
    WEntraLote = 0
    Erase Verifica
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                
            DBGrid1.Col = 1
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Articulo = UCase(DBGrid1.Text)
                                
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
            XSuma = XSuma + Val(Cantidad)
            
            DBGrid1.Col = 5
            Estado = DBGrid1.Text
            
            WLugar = WLugar + 1
            
            For WCicloLote = 1 To 6 Step 2
            
                WLote1 = xLote(WLugar, WCicloLote)
                WCanti = xLote(WLugar, WCicloLote + 1)
                
                Entra = "S"
                For XCicloLote = 1 To WEntraLote
                    If Val(Verifica(XCicloLote, 1)) = Val(WLote1) Then
                        Verifica(XCicloLote, 2) = Str$(Val(Verifica(XCicloLote, 2)) + Val(WCanti1))
                        Entra = "N"
                        Exit For
                    End If
                Next XCicloLote
                        
                If Entra = "S" Then
                    WEntraLote = WEntraLote + 1
                    Verifica(WEntraLote, 1) = WLote1
                    Verifica(WEntraLote, 2) = WCanti1
                    Verifica(WEntraLote, 3) = Tipo
                    Verifica(WEntraLote, 4) = Terminado
                    Verifica(WEntraLote, 5) = Articulo
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
            
    Next A
    
    If Val(Real.Text) = 0 Then
        m$ = "Cantidad real en 0"
        G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
        General = "N"
    End If
    
    If General = "S" Then
    
    Dife = Abs(Val(Real.Text) - XSuma)
    Porce = Abs(XSuma * 0.15)
    If Dife > Porce Then
    
        DBGrid1.FirstRow = 0
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                
            DBGrid1.Col = 1
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Articulo = UCase(DBGrid1.Text)
            
            DBGrid1.Col = 3
            WDescri = DBGrid1.Text
                                
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
            
            DBGrid1.Col = 5
            Estado = DBGrid1.Text
                        
        Next iRow
        
        T$ = "Grabacion de Hoja de Produccion"
        m$ = "El rendimiento real difiere de los componentes en un +-15%. Desea continuar con la grabacion"
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

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
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
    
        Producto = Auxiliar(Da, 1)
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
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
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
        
    DBGrid1.Refresh
        
    Hoja.Text = WHoja
    Fecha.Text = WFecha
    Producto.Text = WProducto
    Teorico.Text = WTeorico
    Real.Text = WReal
    fechaIng.Text = WFechaing
    
    Suma = 0
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
        
            Suma = Suma + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                
            DBGrid1.Col = 1
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Articulo = UCase(DBGrid1.Text)
                                
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Lote = DBGrid1.Text
                    
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
                If Val(Real.Text) <> 0 Then
                    WSaldo = Str$(Val(Real.Text))
                        Else
                    WSaldo = "0"
                End If
                WLote1 = xLote(Suma, 1)
                WLote2 = xLote(Suma, 3)
                WLote3 = xLote(Suma, 5)
                WCanti1 = xLote(Suma, 2)
                WCanti2 = xLote(Suma, 4)
                WCanti3 = xLote(Suma, 6)
                WCosto1 = "0"
                WCosto2 = "0"
                WCosto3 = "0"
                
                XCosto1 = 0
                XCosto2 = 0
                XCosto3 = 0
                
                Select Case Tipo
                    Case "T"
                        Producto = Terminado
                        Call Calcula_Costo_Produccion(Producto, XCosto1, XCosto2, XCosto3)
                
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
                
                Rem WCosto1 = "0"
                Rem WCosto2 = "0"
                Rem WCosto3 = "0"

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
            
    Next A
    
    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    XParam = "'" + WHoja + "','" _
                 + WFechaord + "'"
    Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)

    For Da = 1 To Renglon

        Producto = Auxiliar(Da, 1)
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
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
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
    
    If Valor1.Text <> "" Or valor2.Text <> "" Or Valor3.Text <> "" Or valor4.Text <> "" Or valor5.Text <> "" Or valor6.Text <> "" Or valor7.Text <> "" Or valor8.Text <> "" Or valor9.Text <> "" Or valor10.Text <> "" Or Ensayo.Text <> "" Or Aspecto.Text <> "" Or Observaciones.Text <> "" Or Confecciono.Text <> "" Then
        Call GrabaPrueba
    End If
    
    Call Limpia_Click
    
    End If
        
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Hoja.SetFocus
        
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

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
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
    
        Producto = Auxiliar(Da, 1)
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
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
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
        
    DBGrid1.Refresh
        
    Hoja.Text = WHoja
    Fecha.Text = WFecha
    Producto.Text = WProducto
    Teorico.Text = WTeorico
    Real.Text = "0"
    fechaIng.Text = "  /  /    "
    WReal = "0"
    WFechaing = "  /  /    "
    
    Suma = 0
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
        
            Suma = Suma + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                
            DBGrid1.Col = 1
            Terminado = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Articulo = UCase(DBGrid1.Text)
                                
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Lote = DBGrid1.Text
                    
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
            
    Next A
    
    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    XParam = "'" + WHoja + "','" _
                 + WFechaord + "'"
    Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    For Da = 1 To Renglon
    
        Producto = Auxiliar(Da, 1)
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
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
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
        
    Call Limpia_Click
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
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

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    Valor1.Text = ""
    valor2.Text = ""
    Valor3.Text = ""
    valor4.Text = ""
    valor5.Text = ""
    valor6.Text = ""
    valor7.Text = ""
    valor8.Text = ""
    valor9.Text = ""
    valor10.Text = ""
    
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    
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
    Erase xLote

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    
    salgo = "N"
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                If iCol = 0 Then
                    If DBGrid1.Text = "" Then
                        salgo = "S"
                            Else
                        DBGrid1.Text = ""
                    End If
                        Else
                    DBGrid1.Text = ""
                End If
                If salgo = "S" Then Exit For
            Next iCol
            If salgo = "S" Then Exit For
        Next iRow
        If salgo = "S" Then Exit For
    Next A
    
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
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    Renglon = 0
    Graba.Enabled = True
    Anula.Enabled = False

    Hoja.SetFocus

End Sub

Private Sub VerControl_Click()
    Call ImprimeEnsayo
    IngresaEnsayo.Height = 5055
    IngresaEnsayo.Left = 360
    IngresaEnsayo.Top = 1200
    IngresaEnsayo.Width = 11295
    IngresaEnsayo.Visible = True
    Valor1.SetFocus
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
            WDescripcion.Caption = rstTerminado!Descripcion
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
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        CargaLote.Visible = True
        If WTipo.Text = "M" Then
            CargaLote.Caption = "Ingreso de Lote"
            dada.Caption = "Lote"
                Else
            CargaLote.Caption = "Ingreso de Partida"
            dada.Caption = "Partida"
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
        
        If Val(xLote(Val(WLinea.Text), 1)) <> 0 Then
            WLote1.Text = xLote(Val(WLinea.Text), 1)
            WCanti1.Text = xLote(Val(WLinea.Text), 2)
            WControl1.Locked = False
            WControl1.Text = ""
            WControl1.Locked = True
        End If
        If Val(xLote(Val(WLinea.Text), 3)) <> 0 Then
            WLote2.Text = xLote(Val(WLinea.Text), 3)
            WCanti2.Text = xLote(Val(WLinea.Text), 4)
            WControl2.Locked = False
            WControl2.Text = ""
            WControl2.Locked = True
        End If
        If Val(xLote(Val(WLinea.Text), 5)) <> 0 Then
            WLote3.Text = xLote(Val(WLinea.Text), 5)
            WCanti3.Text = xLote(Val(WLinea.Text), 6)
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
                XParam = "'" + WLote1.Text + "','" _
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
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
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
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
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
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
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
    pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    
                Rem DBGrid1.Col = 0
                Rem DBGrid1.Text = "M"
                Rem DBGrid1.Col = 1
                Rem DBGrid1.Text = "  -     -   "
                Rem DBGrid1.Col = 2
                Rem DBGrid1.Text = !Codigo
                Rem DBGrid1.Col = 3
                Rem DBGrid1.Text = !Descripcion
                Rem
                Rem Call Alta_Vector
                Rem WLinea.Text = WAnterior + 1
                Rem If ValF(WLinea.Text) > 0 Then
                Rem     DBGrid1.Row = Val(WLinea.Text) - 1
                Rem End If
                Rem
                Rem Call DBGrid1.SetFocus
                Rem WCantidad.SetFocus
                    
            End If
            Call Alta_Vector
            
        Case 1
            Indice = pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WTerminado.Text = rstTerminado!Codigo
                WDescripcion.Caption = rstTerminado!Descripcion
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

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 5, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 550
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1

         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i

    Erase xLote
    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    Valor1.Text = ""
    valor2.Text = ""
    Valor3.Text = ""
    valor4.Text = ""
    valor5.Text = ""
    valor6.Text = ""
    valor7.Text = ""
    valor8.Text = ""
    valor9.Text = ""
    valor10.Text = ""
    
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    
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
    
    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            Hoja.Text = rstHoja!Hoja + 1
        End With
        rstHoja.Close
    End If
    
    WE = WEmpresa
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            Prgmodhoja.Caption = "Ingreso de Hoja de Produccion :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
    
    Graba.Enabled = True
    Anula.Enabled = False
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Hoja.SetFocus
    
End Sub

Private Sub Proceso_Click()

    salgo = "N"
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                If iCol = 0 Then
                    If DBGrid1.Text = "" Then
                        salgo = "S"
                            Else
                        DBGrid1.Text = ""
                    End If
                        Else
                    DBGrid1.Text = ""
                End If
                If salgo = "S" Then Exit For
            Next iCol
            If salgo = "S" Then Exit For
        Next iRow
        If salgo = "S" Then Exit For
    Next A

    Renglon = 0
    Erase Auxiliar
    Erase xLote
    WSaldoant = 0
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstHoja!Tipo
                    
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstHoja!Terminado
                    Auxi1 = rstHoja!Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = rstHoja!Articulo
                    Auxi2 = rstHoja!Articulo
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", rstHoja!Cantidad)
                
                    Auxiliar(Renglon, 1) = rstHoja!Tipo
                    Auxiliar(Renglon, 2) = Auxi1
                    Auxiliar(Renglon, 3) = Auxi2
                    
                    xLote(Renglon, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    xLote(Renglon, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    xLote(Renglon, 3) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    xLote(Renglon, 4) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    xLote(Renglon, 5) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    xLote(Renglon, 6) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    xLote(Renglon, 7) = ""
                    
                    If Val(Real.Text) <> 0 Then
                        If Val(xLote(Renglon, 1)) = 0 And rstHoja!Lote <> 0 Then
                            xLote(Renglon, 1) = rstHoja!Lote
                            xLote(Renglon, 2) = rstHoja!Cantidad
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
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        Tipo = Auxiliar(Renglon, 1)
        Auxi1 = Auxiliar(Renglon, 2)
        Auxi2 = Auxiliar(Renglon, 3)
                
        Select Case Tipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstTerminado!Descripcion
                    rstTerminado.Close
                    WArticulo.SetFocus
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                    WArticulo.SetFocus
                End If
            Case Else
        End Select
    Next Da

    Renglon = Renglon + 1
    Ultimo = Renglon
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    Rem Renglon = Renglon - 2
    Rem Lugar1 = Int((Renglon - 1) / 10) * 10
    Rem Lugar2 = Renglon - Lugar1
    Rem DBGrid1.FirstRow = Lugar1
    Rem DBGrid1.Row = Lugar2 - 1
    
    If Val(Real.Text) <> 0 And Val(Real.Text) <> WSaldoant Then
        Graba.Enabled = False
        Anula.Enabled = False
            Else
        If Val(Real.Text) = 0 Then
            Graba.Enabled = True
            Anula.Enabled = False
            WExiste = "N"
                Else
            Graba.Enabled = False
            Anula.Enabled = True
            WExiste = "S"
        End If
    End If
    
    Renglon = Renglon - 1
    
    WTipo.SetFocus

End Sub

Private Sub Alta_Vector()

    Lugar1 = Int((Ultimo - 1) / 10) * 10
    Lugar2 = Ultimo - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    DBGrid1.Col = 4
    DBGrid1.Text = ""

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            Ultimo = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = "S"
            
            xLote(Renglon, 1) = WLote1.Text
            xLote(Renglon, 2) = WCanti1.Text
            xLote(Renglon, 3) = WLote2.Text
            xLote(Renglon, 4) = WCanti2.Text
            xLote(Renglon, 5) = WLote3.Text
            xLote(Renglon, 6) = WCanti3.Text
            
            Rem DBGrid1.Row = Renglon
            Rem DBGrid1.Col = 0
            
                Else
                
            WRen = Val(WLinea.Text)
            
            Lugar1 = Int((WRen - 1) / 10) * 10
            Lugar2 = WRen - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = "S"
            
            xLote(WRen, 1) = WLote1.Text
            xLote(WRen, 2) = WCanti1.Text
            xLote(WRen, 3) = WLote2.Text
            xLote(WRen, 4) = WCanti2.Text
            xLote(WRen, 5) = WLote3.Text
            xLote(WRen, 6) = WCanti3.Text
            
            Rem DBGrid1.Row = Anterior
            Rem DBGrid1.Col = 0
            
    End If
    
    Lugar1 = Int((Ultimo - 1) / 10) * 10
    Lugar2 = Ultimo - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entra = "N"
        spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            Entra = "S"
            Fecha.Text = rstHoja!Fecha
            Real.Text = rstHoja!Real
            Teorico.Text = rstHoja!Teorico
            fechaIng.Text = rstHoja!fechaIng
            Producto.Text = rstHoja!Producto
            rstHoja.Close
        End If
        
        If Entra = "S" Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            Call Proceso_Click
            If Val(Real.Text) = 0 Then
                Call ImprimeEnsayo
                IngresaEnsayo.Height = 5055
                IngresaEnsayo.Left = 360
                IngresaEnsayo.Top = 1200
                IngresaEnsayo.Width = 11295
                IngresaEnsayo.Visible = True
                Valor1.SetFocus
                    Else
                Auxi = Hoja.Text
                Call Ceros(Auxi, 5)
                ClavePrue$ = "1" + Auxi
                spPrueter = "ConsultaPrueter " + "'" + ClavePrue$ + "'"
                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrueter.RecordCount > 0 Then
                    Valor1.Text = rstPrueter!Valor1
                    valor2.Text = rstPrueter!valor2
                    Valor3.Text = rstPrueter!Valor3
                    valor4.Text = rstPrueter!valor4
                    valor5.Text = rstPrueter!valor5
                    valor6.Text = rstPrueter!valor6
                    valor7.Text = rstPrueter!valor7
                    valor8.Text = rstPrueter!valor8
                    valor9.Text = rstPrueter!valor9
                    valor10.Text = rstPrueter!valor10
                    Ensayo.Text = rstPrueter!Ensayo
                    Aspecto.Text = rstPrueter!Aspecto
                    Observaciones.Text = rstPrueter!Observaciones
                    Confecciono.Text = rstPrueter!Confecciono
                    rstPrueter.Close
                        Else
                    Valor1.Text = ""
                    valor2.Text = ""
                    Valor3.Text = ""
                    valor4.Text = ""
                    valor5.Text = ""
                    valor6.Text = ""
                    valor7.Text = ""
                    valor8.Text = ""
                    valor9.Text = ""
                    valor10.Text = ""
                    Ensayo.Text = ""
                    Aspecto.Text = ""
                    Observaciones.Text = ""
                    Confecciono.Text = ""
                End If
                
            End If
                
                Else
                    
            Existe = "N"
                    
            WHoja = Hoja.Text
            LeeHoja = "N"
            Call Limpia_Click
            LeeHoja = "S"
            Hoja.Text = WHoja
            Producto.SetFocus
                
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    Rem If KeyAscii = 13 Then
    Rem     Call Valida_fecha(Fecha.Text, Auxi)
    Rem     If Auxi = "S" Then
    Rem         Producto.SetFocus
    Rem             Else
    Rem         Fecha.SetFocus
    Rem     End If
    Rem End If
End Sub

Private Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                Call ImprimeEnsayo
                IngresaEnsayo.Height = 5055
                IngresaEnsayo.Left = 360
                IngresaEnsayo.Top = 1200
                IngresaEnsayo.Width = 11295
                IngresaEnsayo.Visible = True
                Valor1.SetFocus
                    Else
                Producto.SetFocus
            End If
        End If
    End If
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
            WTipo.SetFocus
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
    
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstComposicion!Tipo
                
                    If rstComposicion!Articulo1 = "  -   -  " Then
                        DBGrid1.Col = 2
                        DBGrid1.Text = "  -   -   "
                        Auxi1 = "  -   -   "
                            Else
                        DBGrid1.Col = 2
                        DBGrid1.Text = rstComposicion!Articulo1
                        Auxi1 = rstComposicion!Articulo1
                    End If
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstComposicion!Articulo2
                    Auxi2 = rstComposicion!Articulo2
                
                    Cantidad = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Cantidad)
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = ""
                    
                    Auxiliar(Renglon, 1) = rstComposicion!Tipo
                    Auxiliar(Renglon, 2) = Auxi1
                    Auxiliar(Renglon, 3) = Auxi2
                    Auxiliar(Renglon, 4) = Cantidad
                
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
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
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
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstTerminado!Descripcion
                    WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                    WArticulo.SetFocus
                End If
            Case "M"
                WImpre1 = Auxi2
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstArticulo!Descripcion
                    WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                    WArticulo.SetFocus
                End If
            Case Else
        End Select
        
        DBGrid1.Col = 4
        DBGrid1.Text = Pusing("###,###.##", Str$(XCantidad))
        
    Next Da
    
    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1

End Sub

Sub Impresion()

        Open "lpt1" For Output As #1

        Print #1, Chr$(27) + Chr$(71)
        Print #1,
        Print #1, Chr$(18)

        Print #1, Tab(15); Left$(Producto.Text, 2)
        Select Case Val(WEmpresa)
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
        
        For A = 0 To 3
        
                Suma = A * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Tipo = DBGrid1.Text
                    
                    DBGrid1.Col = 1
                    Terminado = UCase(DBGrid1.Text)
                    
                    DBGrid1.Col = 2
                    Articulo = UCase(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Cantidad = DBGrid1.Text
                    
                    DBGrid1.Col = 5
                    Lote = DBGrid1.Text
                    
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
            
        Next A

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

    Rem caca
    
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
        
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
        End If
        
    End If
    
    If Val(WLote2.Text) <> 0 Then
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
            
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
        End If
        
    End If
    
    
    If Val(WLote3.Text) <> 0 Then
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
        
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
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
    
            spComposicion = "ConsultaComposicionProducto " + "'" + ZVector(ZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZTipo = rstComposicion!Tipo
                        ZArticulo1 = rstComposicion!Articulo1
                        ZArticulo2 = rstComposicion!Articulo2
                        ZCantidad = rstComposicion!Cantidad
                        
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
        Call CierraIngresaEnsayo_Click
    End If
End Sub

Private Sub CierraIngresaEnsayo_Click()
    IngresaEnsayo.Visible = False
    Teorico.SetFocus
End Sub

Private Sub ImprimeEnsayo()

    WProducto = "PT" + Mid$(Producto.Text, 3, 10)
    
    Ensayo1.Caption = ""
    Ensayo2.Caption = ""
    Ensayo3.Caption = ""
    Ensayo4.Caption = ""
    Ensayo5.Caption = ""
    Ensayo6.Caption = ""
    Ensayo7.Caption = ""
    Ensayo8.Caption = ""
    Ensayo9.Caption = ""
    Ensayo10.Caption = ""
    
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

    spEspecif = "ConsultaEspecif " + "'" + WProducto + "'"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        Ensayo1.Caption = rstEspecif!Ensayo1
        Ensayo2.Caption = rstEspecif!Ensayo2
        Ensayo3.Caption = rstEspecif!Ensayo3
        Ensayo4.Caption = rstEspecif!Ensayo4
        Ensayo5.Caption = rstEspecif!Ensayo5
        Ensayo6.Caption = rstEspecif!Ensayo6
        Ensayo7.Caption = rstEspecif!Ensayo7
        Ensayo8.Caption = rstEspecif!Ensayo8
        Ensayo9.Caption = rstEspecif!Ensayo9
        Ensayo10.Caption = rstEspecif!Ensayo10
        Std1.Caption = rstEspecif!Valor1
        Std2.Caption = rstEspecif!valor2
        Std3.Caption = rstEspecif!Valor3
        Std4.Caption = rstEspecif!valor4
        Std5.Caption = rstEspecif!valor5
        Std6.Caption = rstEspecif!valor6
        Std7.Caption = rstEspecif!valor7
        Std8.Caption = rstEspecif!valor8
        Std9.Caption = rstEspecif!valor9
        Std10.Caption = rstEspecif!valor10
        
        rstEspecif.Close
                        
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
    
    If Val(Ensayo1.Caption) = 0 Then
        Ensayo1.Caption = ""
    End If
    If Val(Ensayo2.Caption) = 0 Then
        Ensayo2.Caption = ""
    End If
    If Val(Ensayo3.Caption) = 0 Then
        Ensayo3.Caption = ""
    End If
    If Val(Ensayo4.Caption) = 0 Then
        Ensayo4.Caption = ""
    End If
    If Val(Ensayo5.Caption) = 0 Then
        Ensayo5.Caption = ""
    End If
    If Val(Ensayo6.Caption) = 0 Then
        Ensayo6.Caption = ""
    End If
    If Val(Ensayo7.Caption) = 0 Then
        Ensayo7.Caption = ""
    End If
    If Val(Ensayo8.Caption) = 0 Then
        Ensayo8.Caption = ""
    End If
    If Val(Ensayo9.Caption) = 0 Then
        Ensayo9.Caption = ""
    End If
    If Val(Ensayo10.Caption) = 0 Then
        Ensayo10.Caption = ""
    End If

End Sub

Private Sub GrabaPrueba()

    WPasa = "S"
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Hoja.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If WPasa = "S" Then
        Auxi = Hoja.Text
        Call Ceros(Auxi, 5)
        ClavePrue$ = "1" + Auxi
        spPrueter = "ConsultaPrueter " + "'" + ClavePrue$ + "'"
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            m$ = "Prueba ya ingresada"
            A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
            rstPrueter.Close
                    Else
            ClavePrue$ = "2" + Auxi
            spPrueter = "ConsultaPrueter " + "'" + ClavePrue$ + "'"
            Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
                m$ = "Prueba ya ingresada"
                A% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
                WPasa = "N"
                rstPrueter.Close
            End If
        End If
    End If

    If WPasa = "S" Then

        Auxi1 = Hoja.Text
        Call Ceros(Auxi1, 5)
        Lote = Auxi1
        
        Auxi = "1"
        
        WPrueba = Auxi + Lote
        WProducto = Producto.Text
        WFecha = fechaIng.Text
        WValor1 = Valor1.Text
        Wvalor2 = valor2.Text
        WValor3 = Valor3.Text
        Wvalor4 = valor4.Text
        Wvalor5 = valor5.Text
        Wvalor6 = valor6.Text
        Wvalor7 = valor7.Text
        Wvalor8 = valor8.Text
        Wvalor9 = valor9.Text
        Wvalor10 = valor10.Text
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = ""
        WLote = Lote
        WRechazo = Lote
        WDate = Date$
        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        XParam = "'" + WPrueba + "','" _
                + WProducto + "','" _
                + WFecha + "','" _
                + WValor1 + "','" _
                + Wvalor2 + "','" _
                + WValor3 + "','" _
                + Wvalor4 + "','" _
                + Wvalor5 + "','" _
                + Wvalor6 + "','" _
                + Wvalor7 + "','" _
                + Wvalor8 + "','" _
                + Wvalor9 + "','" _
                + Wvalor10 + "','" _
                + WEnsayo + "','" _
                + WAspecto + "','" _
                + WObservaciones + "','" _
                + WConfecciono + "','" _
                + WLiberada + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WFechaord + "','" _
                + WDate + "'"
        Set rstPrueter = db.OpenRecordset("AltaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
        
End Sub
