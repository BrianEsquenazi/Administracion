VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgmodhoja 
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
   Begin VB.Frame IngresaEnsayo 
      Height          =   1185
      Left            =   600
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   11160
      Begin VB.TextBox ValorNumero10 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   145
         Top             =   5040
         Width           =   800
      End
      Begin VB.TextBox ValorNumero9 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   144
         Top             =   4560
         Width           =   800
      End
      Begin VB.TextBox ValorNumero8 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   143
         Top             =   4080
         Width           =   800
      End
      Begin VB.TextBox ValorNumero7 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   142
         Top             =   3600
         Width           =   800
      End
      Begin VB.TextBox ValorNumero6 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   141
         Top             =   3120
         Width           =   800
      End
      Begin VB.TextBox ValorNumero5 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   140
         Top             =   2640
         Width           =   800
      End
      Begin VB.TextBox ValorNumero4 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   139
         Top             =   2160
         Width           =   800
      End
      Begin VB.TextBox ValorNumero3 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   138
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox ValorNumero2 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   137
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox ValorNumero1 
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
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   136
         Top             =   720
         Width           =   800
      End
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
         Text            =   " "
         Top             =   6840
         Width           =   3975
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Text            =   " "
         Top             =   720
         Width           =   2900
      End
      Begin VB.TextBox valor2 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Text            =   " "
         Top             =   1200
         Width           =   2900
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   1680
         Width           =   2900
      End
      Begin VB.TextBox valor4 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   2160
         Width           =   2900
      End
      Begin VB.TextBox valor5 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   2640
         Width           =   2900
      End
      Begin VB.TextBox valor6 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   3120
         Width           =   2900
      End
      Begin VB.TextBox valor7 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   3600
         Width           =   2900
      End
      Begin VB.TextBox valor8 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Text            =   " "
         Top             =   4080
         Width           =   2900
      End
      Begin VB.TextBox valor9 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   4560
         Width           =   2900
      End
      Begin VB.TextBox valor10 
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   " "
         Top             =   5040
         Width           =   2900
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
         TabIndex        =   41
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
         TabIndex        =   40
         Text            =   " "
         Top             =   6120
         Width           =   3975
      End
      Begin MSMask.MaskEdBox FechaPrueba 
         Height          =   285
         Left            =   1800
         TabIndex        =   118
         Top             =   7320
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label25 
         Alignment       =   2  'Center
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
         Left            =   10920
         TabIndex        =   146
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Prueba"
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
         TabIndex        =   117
         Top             =   7320
         Width           =   1455
      End
      Begin VB.Label Std1010 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   109
         Top             =   5280
         Width           =   4455
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   108
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Std99 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   107
         Top             =   4800
         Width           =   4455
      End
      Begin VB.Label Std88 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   106
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Std77 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   105
         Top             =   3840
         Width           =   4455
      End
      Begin VB.Label Std66 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   104
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Std55 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   103
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Std44 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   102
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Std33 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   101
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   100
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   99
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblensayo 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   88
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDescri 
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
         Left            =   1080
         TabIndex        =   87
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
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
         Left            =   3360
         TabIndex        =   86
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   85
         Top             =   720
         Width           =   2145
      End
      Begin VB.Label descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   84
         Top             =   1200
         Width           =   2220
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   83
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   82
         Top             =   2160
         Width           =   2220
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   81
         Top             =   2640
         Width           =   2220
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   80
         Top             =   3120
         Width           =   2220
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   79
         Top             =   3600
         Width           =   2220
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   78
         Top             =   4080
         Width           =   2220
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   77
         Top             =   4560
         Width           =   2220
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
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
         TabIndex        =   71
         Top             =   360
         Width           =   2900
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   70
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   69
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   67
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Std6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   66
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   65
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   64
         Top             =   4080
         Width           =   4455
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   63
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   5040
         Width           =   735
      End
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
      Left            =   8040
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   130
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
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   129
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
      Left            =   11400
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   128
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   9720
      TabIndex        =   119
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
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
         TabIndex        =   125
         Top             =   720
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
         TabIndex        =   124
         Top             =   1080
         Width           =   975
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
         TabIndex        =   123
         Top             =   1440
         Width           =   975
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
         TabIndex        =   122
         Top             =   720
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
         TabIndex        =   121
         Top             =   1080
         Width           =   855
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
         TabIndex        =   120
         Top             =   1440
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
         TabIndex        =   127
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label20 
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
         TabIndex        =   126
         Top             =   360
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   11280
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"modhoja.frx":0000
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   94
      Top             =   1560
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   1560
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   96
      Top             =   1560
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   1560
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton VerControl 
      Caption         =   "Muestra   Ensayos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   92
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Frame XClave 
      Caption         =   "Ingrese de Clave de Seguridad"
      Height          =   1815
      Left            =   3120
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Grabacion"
         Height          =   375
         Left            =   600
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   240
         Width           =   2895
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
      Left            =   11400
      Top             =   720
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
      Left            =   2280
      TabIndex        =   13
      Top             =   6480
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
      Left            =   10680
      TabIndex        =   8
      Top             =   720
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
      ItemData        =   "modhoja.frx":007C
      Left            =   3360
      List            =   "modhoja.frx":0083
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
      Height          =   4095
      Left            =   120
      TabIndex        =   93
      Top             =   1320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7223
      _Version        =   327680
      BackColor       =   16777088
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   9720
      TabIndex        =   112
      Top             =   1320
      Width           =   1935
      Begin VB.TextBox Resultado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   115
         Text            =   " "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton ConfirmaResultado 
         Caption         =   "Confirma"
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
         Left            =   360
         TabIndex        =   114
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Producto"
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
         TabIndex        =   113
         Top             =   240
         Width           =   1695
      End
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
      TabIndex        =   110
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
      TabIndex        =   111
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Anula 
      Caption         =   "Baja  Hoja"
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
      TabIndex        =   116
      Top             =   7680
      Width           =   975
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
      Left            =   9480
      MaxLength       =   2
      TabIndex        =   134
      Text            =   " "
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label24 
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
      Left            =   8520
      TabIndex        =   135
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label23 
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
      Left            =   6360
      TabIndex        =   133
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label22 
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
      Left            =   8640
      TabIndex        =   132
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label21 
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
      Left            =   10320
      TabIndex        =   131
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
      Left            =   6960
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
      Left            =   3600
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
Attribute VB_Name = "Prgmodhoja"
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
Private XLote(100, 7) As String


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
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String


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
Private bajaLote(3, 2) As String
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
Dim Verifica(100, 10) As String
Dim WLugar As Integer
Dim WEntraLote As Integer
Dim WCicloLote As Integer
Dim XCicloLote As Integer
Dim Entra As String
Dim ZProducto As String
Dim WProducto As String
Dim WProductoNuevo As String
Dim ZZBloqueo As Double
Dim ZZTipoBloqueo As String
Dim ZZCodigoBloqueo As String

Dim ZArti(150, 10) As String
Dim Empe(12, 10) As String

Dim WWArticulo As String
Dim WWSaldo As Double
Dim WWVencido As Double

Dim XXSaldo As Double
Dim ZFechaVto As String
Dim XMes As String
Dim XAno As String
Dim ZVto As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim QSaldo As Double

Dim EmpresaActual As String

Dim ZEnsayo(10) As String
Dim ZDesde(10) As String
Dim ZHasta(10) As String
Dim ZUnidad(10) As String
Dim ZValorNumero(10) As String

Dim XCantidad As Double
Dim XSaldo As Double
Dim ZSaldo As Double
Dim WEscrito As Integer
Dim WTeorico As String
Dim ZZSaldo As Double

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
    
        Producto = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Real = Auxiliar(Da, 5)
        Teorico = Auxiliar(Da, 6)
        Tipo = Auxiliar(Da, 7)
        bajaLote(1, 1) = Auxiliar(Da, 8)
        bajaLote(1, 2) = Auxiliar(Da, 9)
        bajaLote(2, 1) = Auxiliar(Da, 10)
        bajaLote(2, 2) = Auxiliar(Da, 11)
        bajaLote(3, 1) = Auxiliar(Da, 12)
        bajaLote(3, 2) = Auxiliar(Da, 13)
        
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
                
                    Lote = bajaLote(xda, 1)
                    ZCantidad = bajaLote(xda, 2)
            
                    WControla = 0
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        WCodigo = rstArticulo!Codigo
                        rstArticulo.Close
                    
                        Lote = bajaLote(xda, 1)
                        ZCantidad = bajaLote(xda, 2)
                    
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
                
                    Lote = bajaLote(xda, 1)
                    ZCantidad = bajaLote(xda, 2)
            
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        WCodigo = rstTerminado!Codigo
                        rstTerminado.Close
                        
                        Lote = bajaLote(xda, 1)
                        ZCantidad = bajaLote(xda, 2)
                        
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
    
    ZSql = ""
    ZSql = ZSql + "DELETE Prueter"
    ZSql = ZSql + " Where Lote = " + "'" + Hoja.Text + "'"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Limpia_Click
    Hoja.SetFocus
        
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
    
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WLinea.Text = ""
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

Private Sub ConfirmaResultado_Click()

    Resultado.Text = UCase(Resultado.Text)

    If Left$(Producto.Text, 2) = "DW" Or Left$(Producto.Text, 2) = "NW" Then

        If Resultado.Text <> "DW" Then
            If Resultado.Text <> "NW" Then
                ca% = MsgBox("El resultado debe ser PT, NK, RE o DW o NW", 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        End If
        
            Else
        
        If Left$(Producto.Text, 2) = "PT" Then
    
            If Resultado.Text <> "PT" Then
                If Resultado.Text <> "NK" Then
                    If Resultado.Text <> "RE" Then
                        ca% = MsgBox("El resultado debe ser PT, SE, NK, RE", 0, "Ingreso de Hoja de Produccion")
                        Exit Sub
                    End If
                End If
            End If
            
                Else
            
            If Left$(Producto.Text, 2) = "SE" Then
            
                If Resultado.Text <> "NK" Then
                    If Resultado.Text <> "RE" Then
                        If Resultado.Text <> "SE" Then
                            ca% = MsgBox("El resultado debe ser PT, SE, NK, RE", 0, "Ingreso de Hoja de Produccion")
                            Exit Sub
                        End If
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
    
    
    Rem If Valor1.Text <> "" Or valor2.Text <> "" Or Valor3.Text <> "" Or valor4.Text <> "" Or valor5.Text <> "" Or valor6.Text <> "" Or valor7.Text <> "" Or valor8.Text <> "" Or valor9.Text <> "" Or valor10.Text <> "" Or Ensayo.Text <> "" Or Aspecto.Text <> "" Or Observaciones.Text <> "" Or Confecciono.Text <> "" Then
    
        XEmpresa = Wempresa
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
    
        Sql1 = "Select Desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,Desde9,Desde10"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
        
            ZDesde(1) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZDesde(2) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZDesde(3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZDesde(4) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZDesde(5) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZDesde(6) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZDesde(7) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZDesde(8) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZDesde(9) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZDesde(10) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            rstEspecifUnifica.Close
            End If
           
        Sql1 = "Select Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Ensayo1,Ensayo2,Ensayo3,Ensayo4,Ensayo5,Ensayo6,ensayo7,Ensayo8,Ensayo9,Ensayo10"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstEspecifUnifica.RecordCount > 0 Then
            ZHasta(1) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZHasta(2) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZHasta(3) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZHasta(4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZHasta(5) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZHasta(6) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZHasta(7) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZHasta(8) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZHasta(9) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZHasta(10) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
            rstEspecifUnifica.Close
        End If
        
        Call Conecta_Empresa
    
        ZValorNumero(1) = ValorNumero1.Text
        ZValorNumero(2) = ValorNumero2.Text
        ZValorNumero(3) = ValorNumero3.Text
        ZValorNumero(4) = ValorNumero4.Text
        ZValorNumero(5) = ValorNumero5.Text
        ZValorNumero(6) = ValorNumero6.Text
        ZValorNumero(7) = ValorNumero7.Text
        ZValorNumero(8) = ValorNumero8.Text
        ZValorNumero(9) = ValorNumero9.Text
        ZValorNumero(10) = ValorNumero10.Text
        
        Rem For WWCiclo = 1 To 10
        Rem     If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
        Rem         If Val(ZValorNumero(WWCiclo)) = 0 Then
        Rem             m$ = "No se informado valor de control en una de las ensayos que requiere validacion"
        Rem             A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Rem             Exit Sub
        Rem         End If
        Rem     End If
        Rem Next WWCiclo
        
        
        If Resultado.Text = "PT" Or Resultado.Text = "SE" Then
        
            For WWCiclo = 1 To 10
            
                If ZEnsayo(WWCiclo) <> 0 Then
            
                    If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
                    
                        If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                            aa = Val(ZValorNumero(WWCiclo))
                            If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Or Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                                    
                        If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) = 0 Then
                            If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                                
                        If Val(ZDesde(WWCiclo)) = 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                            If Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                        
                            Else
                            
                        If Trim(UCase(ZValorNumero(WWCiclo))) <> "S" Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                        
                    End If
                    
                End If
            
            Next WWCiclo
        
        End If
    
    Rem End If
    
    If Resultado.Text = "PT" Then
    
    
        WTerminado = Producto.Text
        XCodigo = Val(Mid$(WTerminado, 4, 5))
        XTipoPro = ""
        If Left$(WTerminado, 2) = "PT" Then
            If XCodigo >= 0 And XCodigo <= 999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 11000 And XCodigo <= 11999 Then
                    XTipoPro = "CO"
                End If
            End If
        End If
        
        If XTipoPro <> "CO" Then
            Rem BY NAN 24-09-2013 SE AGREGA PARA PELLITAL
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql & "Select *"
                    ZSql = ZSql & " FROM AltaCertificado"
                    ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
                    ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                    spAltaCertificado = ZSql
                    Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstAltaCertificado.RecordCount > 0 Then
                        rstAltaCertificado.Close
                            Else
                        Call Conecta_Empresa
                        m$ = "No se ha definido el certificado de analiss de este producto para el cliente generico S00102"
                        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Exit Sub
                    End If
                Case Else
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
      
   
                    ZSql = ""
                    ZSql = ZSql & "Select *"
                    ZSql = ZSql & " FROM AltaCertificado"
                    ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
                    ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "P99999" + "'"
                    spAltaCertificado = ZSql
                    Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstAltaCertificado.RecordCount > 0 Then
                        rstAltaCertificado.Close
                            Else
                        Call Conecta_Empresa
                        m$ = "No se ha definido el certificado de analiss de este producto para el cliente generico P99999"
                        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Exit Sub
                    End If
            End Select
        
        End If
        
        Call Conecta_Empresa
        
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
        rstHoja.Close
        If WMarcaLabora = "S" Then
            ca% = MsgBox("La Hoja de Produccion ya fue actualizada", 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
            Else
        ca% = MsgBox("Hoja de Produccion Inexistente", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If
    
    WProductoNuevo = Resultado.Text + Mid$(WProducto, 3, 10)
    
    If WProducto <> WProductoNuevo Then
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + "Producto = " + "'" + WProductoNuevo + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        spTerminado = "ConsultaTerminado " + "'" + WProductoNuevo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = rstTerminado!Codigo
            WProceso = Str$(rstTerminado!Proceso + WTeorico)
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
        
        
        spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WCodigo = rstTerminado!Codigo
            WProceso = Str$(rstTerminado!Proceso - WTeorico)
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
    
    WMarcaLabora = "S"
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + "MarcaLabora = " + "'" + WMarcaLabora + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    WProducto = WProductoNuevo
    Producto.Text = WProductoNuevo
    
    If Valor1.Text <> "" Or valor2.Text <> "" Or Valor3.Text <> "" Or valor4.Text <> "" Or valor5.Text <> "" Or valor6.Text <> "" Or valor7.Text <> "" Or valor8.Text <> "" Or valor9.Text <> "" Or valor10.Text <> "" Or Ensayo.Text <> "" Or Aspecto.Text <> "" Or Observaciones.Text <> "" Or Confecciono.Text <> "" Then
        Call GrabaPrueba
    End If
    
    Call Limpia_Click
    
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
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
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
    
    If Val(Teorico.Text) = 0 Then
        Teorico.SetFocus
            Else
        WCantidad.SetFocus
    End If
    
    CargaLote.Visible = False

End Sub

Private Sub Graba_Click()

    Select Case Val(Wempresa)
        Case 1
            Rem If Val(Hoja.Text) > 69999 Or Val(Hoja.Text) < 57600 Then
            If Val(Hoja.Text) > 199999 Or Val(Hoja.Text) < 100000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 100000 y 199999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 2
            If Val(Hoja.Text) > 55999 Or Val(Hoja.Text) < 55300 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 55300 y 55999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 3
            Rem If Val(Hoja.Text) > 99999 Or Val(Hoja.Text) < 82000 Then
            If Val(Hoja.Text) > 299999 Or Val(Hoja.Text) < 200000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 200000 y 299999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 4
            If Val(Hoja.Text) > 19999 Or Val(Hoja.Text) < 11100 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 11100 y 19999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 5
            Rem If Val(Hoja.Text) > 9999 Or Val(Hoja.Text) < 4600 Then
            If Val(Hoja.Text) > 399999 Or Val(Hoja.Text) < 300000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 300000 y 399999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 6
            Rem If Val(Hoja.Text) > 1999 Or Val(Hoja.Text) < 1740 Then
            If Val(Hoja.Text) > 499999 Or Val(Hoja.Text) < 400000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 400000 y 499999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 7
            Rem If Val(Hoja.Text) > 999 Or Val(Hoja.Text) < 7 Then
            If Val(Hoja.Text) > 519999 Or Val(Hoja.Text) < 500000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 500000 y 519999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 8
            If Val(Hoja.Text) > 29999 Or Val(Hoja.Text) < 20800 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 20800 y 29999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 9
            If Val(Hoja.Text) > 30999 Or Val(Hoja.Text) < 30000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 30000 y 30999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 10
            If Val(Hoja.Text) > 539999 Or Val(Hoja.Text) < 520000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 520000 y 539999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 11
            If Val(Hoja.Text) > 559999 Or Val(Hoja.Text) < 540000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 540000 y 559999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case Else
    End Select

    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.Close
        ca% = MsgBox("Hoja de Produccion ya ingresada", 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If

    If Left$(Producto.Text, 2) = "PT" Or Left$(Producto.Text, 2) = "SE" Or Left$(Producto.Text, 2) = "DW" Then
        If Val(Wempresa) = 1 Or Val(Wempresa) = 2 Or Val(Wempresa) = 3 Or Val(Wempresa) = 4 Then
    
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                    m$ = "El Producto Terminado no se encuentra autorizado para la Produccion"
                    If WEstadoI = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                    End If
                    If WEstadoII = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                    End If
                    If WEstadoIII = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                    End If
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                    Exit Sub
                End If
            End If
            
        End If
    End If
    
    Real.Text = ""
    fechaIng.Text = "  /  /    "
    
    Suma = 0
    Renglon = 0
        
    For iRow = 1 To 40
        
        Suma = Suma + 1
        WRow = iRow
            
        Tipo = Vector.TextMatrix(WRow, 1)
        Terminado = UCase(Vector.TextMatrix(WRow, 2))
        Articulo = UCase(Vector.TextMatrix(WRow, 3))
        Cantidad = Vector.TextMatrix(WRow, 5)
                    
        If Articulo <> "" Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Hoja.Text)
            Call Ceros(Auxi1, 6)
                    
            WClave = Auxi1 + Auxi
            WHoja = Hoja.Text
            WRenglon = Str$(Renglon)
            WFecha = Fecha.Text
            WProducto = Producto.Text
            WTeorico = Teorico.Text
            WReal = Real.Text
            WFechaing = fechaIng.Text
            WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
            WTipo = Tipo
            WTerminado = Terminado
            WArticulo = Articulo
            WCantidad = Cantidad
            WLote = "0"
            WWDate = Date$
            WImporte = ""
            WMarca = ""
            WSaldo = "0"
            
            WLote1 = ""
            WLote2 = ""
            WLote3 = ""
            WCanti1 = "0"
            WCanti2 = "0"
            WCanti3 = "0"
            WCosto1 = "0"
            WCosto2 = "0"
            WCosto3 = "0"
                    
            XCosto1 = 0
            XCosto2 = 0
            XCosto3 = 0
                
            Select Case Tipo
                Case "T"
                    ZProducto = Terminado
                    Call Calcula_Costo_Produccion(ZProducto, XCosto1, XCosto2, XCosto3)
                
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
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " VersionI = " + "'" + VersionI.Text + "',"
            ZSql = ZSql + " VersionII = " + "'" + VersionII.Text + "',"
            ZSql = ZSql + " VersionIII = " + "'" + VersionIII.Text + "'"
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
                
        End If
                        
    Next iRow
    
    
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
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
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
                
        Select Case Tipo
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
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
                    WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
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
    
    Rem If Valor1.Text <> "" Or valor2.Text <> "" Or Valor3.Text <> "" Or valor4.Text <> "" Or valor5.Text <> "" Or valor6.Text <> "" Or valor7.Text <> "" Or valor8.Text <> "" Or valor9.Text <> "" Or valor10.Text <> "" Or Ensayo.Text <> "" Or Aspecto.Text <> "" Or Observaciones.Text <> "" Or Confecciono.Text <> "" Then
    Rem     Call GrabaPrueba
    Rem End If
    
    Call Limpia_Click
        
    Vector.TopRow = 1
    Vector.Col = 1
    Vector.Row = 1
    
    Hoja.SetFocus
        
End Sub

Private Sub GrabaPrueba()

    WPasa = "S"
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        rstTerminado.Close
                    Else
        m$ = "Codigo de Producto invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If Val(Hoja.Text) = 0 Then
        m$ = "Codigo de Partida invalido"
        a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
        WPasa = "N"
    End If
    
    If WPasa = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Prueter"
        ZSql = ZSql + " Where Prueter.Lote = " + "'" + Hoja.Text + "'"
        rsPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrueter.RecordCount > 0 Then
            m$ = "Prueba ya ingresada"
            a% = MsgBox(m$, 0, "Grabacion de Pruebas de Prodcuto Terminado")
            WPasa = "N"
            rstPrueter.Close
        End If
    End If

    If WPasa = "S" Then

        Auxi1 = Hoja.Text
        Call Ceros(Auxi1, 6)
        Lote = Auxi1
        
        Auxi = "1"
        
        WPrueba = Auxi + Lote
        WProducto = Producto.Text
        WFecha = FechaPrueba.Text
        WValor1 = Valor1.Text
        WValor2 = valor2.Text
        WValor3 = Valor3.Text
        WValor4 = valor4.Text
        WValor5 = valor5.Text
        WValor6 = valor6.Text
        WValor7 = valor7.Text
        WValor8 = valor8.Text
        WValor9 = valor9.Text
        WValor10 = valor10.Text
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
                + WValor2 + "','" _
                + WValor3 + "','" _
                + WValor4 + "','" _
                + WValor5 + "','" _
                + WValor6 + "','" _
                + WValor7 + "','" _
                + WValor8 + "','" _
                + WValor9 + "','" _
                + WValor10 + "','" _
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
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueter SET "
        ZSql = ZSql + " ValorNumero1 = " + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + " ValorNumero2 = " + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + " ValorNumero3 = " + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + " ValorNumero4 = " + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + " ValorNumero5 = " + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + " ValorNumero6 = " + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + " ValorNumero7 = " + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + " ValorNumero8 = " + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + " ValorNumero9 = " + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + " ValorNumero10 = " + "'" + ValorNumero10.Text + "'"
        ZSql = ZSql + " Where Prueba = " + "'" + WPrueba + "'"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueter SET "
        ZSql = ZSql + " ValorOriginal1 = " + "'" + WValor1 + "',"
        ZSql = ZSql + " ValorOriginal2 = " + "'" + WValor2 + "',"
        ZSql = ZSql + " ValorOriginal3 = " + "'" + WValor3 + "',"
        ZSql = ZSql + " ValorOriginal4 = " + "'" + WValor4 + "',"
        ZSql = ZSql + " ValorOriginal5 = " + "'" + WValor5 + "',"
        ZSql = ZSql + " ValorOriginal6 = " + "'" + WValor6 + "',"
        ZSql = ZSql + " ValorOriginal7 = " + "'" + WValor7 + "',"
        ZSql = ZSql + " ValorOriginal8 = " + "'" + WValor8 + "',"
        ZSql = ZSql + " ValorOriginal9 = " + "'" + WValor9 + "',"
        ZSql = ZSql + " ValorOriginal10 = " + "'" + WValor10 + "',"
        ZSql = ZSql + " ValorNumeroOriginal1 = " + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal2 = " + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal3 = " + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal4 = " + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal5 = " + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal6 = " + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal7 = " + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal8 = " + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal9 = " + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + " ValorNumeroOriginal10 = " + "'" + ValorNumero10.Text + "'"
        ZSql = ZSql + " Where Prueba = " + "'" + WPrueba + "'"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    
    
    End If
        
End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WTipo.SetFocus
    
End Sub

Private Sub Limpia_Click()
    Graba.Enabled = True
    CargaLote.Visible = False
    
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
    ValorNumero1.Text = ""
    ValorNumero2.Text = ""
    ValorNumero3.Text = ""
    ValorNumero4.Text = ""
    ValorNumero5.Text = ""
    ValorNumero6.Text = ""
    ValorNumero7.Text = ""
    ValorNumero8.Text = ""
    ValorNumero9.Text = ""
    ValorNumero10.Text = ""
    
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    Resultado.Text = ""
    
    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    FechaPrueba.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    
    WExiste = "S"
    
    Call Limpia_Vector
    Anula.Enabled = True
    
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
    Hoja.SetFocus

End Sub

Private Sub VerControl_Click()

    If Val(Wempresa) = 5 Then
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
            rstTerminado.Close
        End If
        If WEscrito = 1 Then
            Exit Sub
        End If
    End If

    Call ImprimeEnsayo
    IngresaEnsayo.Height = 7800
    IngresaEnsayo.Left = 10
    IngresaEnsayo.Top = 400
    IngresaEnsayo.Width = 12000
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
    
        If Val(WCantidad.Text) > 0 Then
        
            WCantidad.Text = Pusing("###,###.###", WCantidad.Text)
            
            Tipo = WTipo.Text
            Auxi2 = WArticulo.Text
            Auxi1 = WTerminado.Text
            XCantidad = Val(WCantidad.Text)
            
            Select Case Tipo
                Case "T"
                    spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                Case "M"
                    spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        rstArticulo.Close
                    End If
                Case Else
            End Select
            
            If Val(Real.Text) = 0 Then
            
                If XCantidad <= WStock Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                        Else
                    m$ = "No existe stock suficiente"
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                End If
                
                    Else
                    
                CargaLote.Visible = True
            
                WLote1.Text = ""
                WCanti1.Text = ""
                WLote2.Text = ""
                WCanti2.Text = ""
                WLote3.Text = ""
                WCanti3.Text = ""
            
                If Val(XLote(Val(WLinea.Text), 1)) <> 0 Then
                    WLote1.Text = XLote(Val(WLinea.Text), 1)
                    WCanti1.Text = XLote(Val(WLinea.Text), 2)
                End If
                If Val(XLote(Val(WLinea.Text), 3)) <> 0 Then
                    WLote2.Text = XLote(Val(WLinea.Text), 3)
                    WCanti2.Text = XLote(Val(WLinea.Text), 4)
                End If
                If Val(XLote(Val(WLinea.Text), 5)) <> 0 Then
                    WLote3.Text = XLote(Val(WLinea.Text), 5)
                    WCanti3.Text = XLote(Val(WLinea.Text), 6)
                End If
            
            End If
        
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

    CargaLote.Visible = False
    Call Limpia_Vector

    Erase XLote
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
    ValorNumero1.Text = ""
    ValorNumero2.Text = ""
    ValorNumero3.Text = ""
    ValorNumero4.Text = ""
    ValorNumero5.Text = ""
    ValorNumero6.Text = ""
    ValorNumero7.Text = ""
    ValorNumero8.Text = ""
    ValorNumero9.Text = ""
    ValorNumero10.Text = ""
    
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    
    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    Resultado.Text = ""
    FechaPrueba.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WExiste = "S"
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""

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
            Prgmodhoja.Caption = "Ingreso de Hoja de Produccion :  " + !Nombre
        End If
    End With
    EmpresaActual = Wempresa
    
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
                    
                    If rstHoja!Renglon = 1 Then
                        WSaldoant = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    End If
                
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
    
    If Val(Real.Text) <> 0 And Val(Real.Text) <> WSaldoant Then
        Anula.Enabled = False
            Else
        Anula.Enabled = True
    End If
    
    If Left$(Producto.Text, 2) = "YQ" Then
        Anula.Enabled = False
    End If
    If Left$(Producto.Text, 2) = "YF" Then
        Anula.Enabled = False
    End If
    If Left$(Producto.Text, 2) = "YP" Then
        Anula.Enabled = False
    End If
    
    Vector.TopRow = 1
    Vector.Row = 1
    Vector.Col = 1
    
    Real.Text = Pusing("###,###.##", Real.Text)
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
            
    End If
    
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Anula.Enabled = True
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
            VersionI.Text = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
            VersionII.Text = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
            VersionIII.Text = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
            ZMarca = IIf(IsNull(rstHoja!Marca), "", rstHoja!Marca)
            rstHoja.Close
        End If
        
        If Entra = "S" Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                Rem DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            Call Proceso_Click
            
            If Val(Real.Text) = 0 And ZMarca <> "X" Then
            
                If WEscrito = 0 Then
            
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
            
                    Call ImprimeEnsayo
                    IngresaEnsayo.Height = 7800
                    IngresaEnsayo.Left = 10
                    IngresaEnsayo.Top = 400
                    IngresaEnsayo.Width = 12000
                    IngresaEnsayo.Visible = True
                    ValorNumero1.SetFocus
                    
                        Else
                        
                    Teorico.SetFocus
            
                End If
                
                    Else
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Prueter"
                ZSql = ZSql + " Where Prueter.Lote = " + "'" + Hoja.Text + "'"
                rsPrueter = ZSql
                Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
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
    
    
    
    
    
    
    
    
    
    
     Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            
                Producto.Text = UCase(Producto.Text)
                
                If Left$(Producto.Text, 2) = "PT" Or Left$(Producto.Text, 2) = "SE" Or Left$(Producto.Text, 2) = "DW" Then
                    If Val(Wempresa) = 1 Or Val(Wempresa) = 2 Or Val(Wempresa) = 3 Or Val(Wempresa) = 4 Then
                        WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                        WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                        WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                        wwlinea = IIf(IsNull(rstTerminado!Linea), "", rstTerminado!Linea)
                        
                        Rem ************** 7-1-2015 la linea 8 no puede actalizar******************
                        If wwlinea = 8 Then
                           Rem ******* gerente produccion puede grabar 27-2-2015
                                If UCase(WClaveOperador) <> "POLOK" Then
                                    m$ = m$ + Chr$(13) + "(No se encuentra habilitada esta opcion para la linea 8)"
                                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                                     Graba.Enabled = False
                                       rstTerminado.Close
                                       Exit Sub
                                     Else
                                       Graba.Enabled = True
                                End If
                             Else
                              Graba.Enabled = True
                        
                        End If
                        Rem fin
                        
                        
                        
                        If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                            m$ = "El Producto Terminado no se encuentra autorizado para la Produccion"
                            If WEstadoI = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                            End If
                            If WEstadoII = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                            End If
                            If WEstadoIII = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                            End If
                            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                            rstTerminado.Close
                            Exit Sub
                        End If
                    End If
                    
                    VersionI.Text = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
                    VersionII.Text = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
                    VersionIII.Text = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
                    
                End If
                
                Rem DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                
                Rem Call ImprimeEnsayo
                Rem IngresaEnsayo.Height = 7800
                Rem IngresaEnsayo.Left = 10
                Rem IngresaEnsayo.Top = 400
                Rem IngresaEnsayo.Width = 12000
                Rem IngresaEnsayo.Visible = True
                Rem Valor1.SetFocus
                
                Teorico.SetFocus
                
                    Else
                Producto.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Teorico_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Teorico.Text = Pusing("###,###.##", Teorico.Text)
        If Existe = "N" Then
            Call Lee_Composicion
        End If
        WTipo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Resultado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ConfirmaResultado_Click
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
                
                If Left$(Producto.Text, 2) = "PT" Then
                    ZZBloqueo = 0
                    ZZCodigoBloqueo = Auxi1
                    ZZTipoBloqueo = "T"
                    Call Calcula_Bloqueo
                    If ZZBloqueo > 0 Then
                        m$ = "Existe el producto " + Auxi1 + " la cantidad de : " + Str$(ZZBloqueo) + " Kgs. Bloqueados" + Chr$(13) + "Comuniquese con el laboratorio para su liberacion"
                        ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        WStock = WStock - ZZBloqueo
                    End If
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
                
                WWArticulo = Auxi2
                WWVencido = 0
                Call Calcula_Stock_Vencido
                If WWVencido > 0 Then
                    m$ = "Existe la materia prima " + Auxi2 + " la cantidad de : " + Str$(WWVencido) + " Kgs. vencidos." + Chr$(13) + "Comuniquese con el laboratorio para su revalida"
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                    WStock = WStock - WWVencido
                End If
                
                If Left$(Producto.Text, 2) = "PT" Then
                    ZZBloqueo = 0
                    ZZCodigoBloqueo = Auxi2
                    ZZTipoBloqueo = "M"
                    Call Calcula_Bloqueo
                    If ZZBloqueo > 0 Then
                        m$ = "Existe la materia prima " + Auxi2 + " la cantidad de : " + Str$(ZZBloqueo) + " Kgs. Bloqueados" + Chr$(13) + "Comuniquese con el laboratorio para su liberacion"
                        ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        WStock = WStock - ZZBloqueo
                    End If
                End If
                
            Case Else
        End Select
        
        If XCantidad <= WStock Then
            Vector.Col = 5
            Vector.Text = Pusing("###,###.###", Str$(XCantidad))
                Else
            WImpre = Str$(WStock)
            WImpre = Pusing("###,###.##", WImpre)
            m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Vector.Col = 5
            Vector.Text = ""
        End If
        
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
                Print #1, Tab(70); "PF"
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

Private Sub ValorNumero1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(1)), ".")
            ZNumeII = Len(Trim(ZDesde(1)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero1.Text = Pusing("###,###.#", ValorNumero1.Text)
                Case 2
                    ValorNumero1.Text = Pusing("###,###.##", ValorNumero1.Text)
                Case 3
                    ValorNumero1.Text = Pusing("###,###.###", ValorNumero1.Text)
                Case 4
                    ValorNumero1.Text = Pusing("###,###.####", ValorNumero1.Text)
                Case 5
                    ValorNumero1.Text = Pusing("###,###.#####", ValorNumero1.Text)
                Case 6
                    ValorNumero1.Text = Pusing("###,###.######", ValorNumero1.Text)
                Case Else
                    ValorNumero1.Text = Pusing("###,###", ValorNumero1.Text)
            End Select
            
            Valor1.Text = ValorNumero1.Text + " " + ZUnidad(1)
            
            ValorNumero2.SetFocus
            
                Else
                
            If ValorNumero1.Text = "S" Or ValorNumero1.Text = "N" Then
                If ValorNumero1.Text = "S" Then
                    Valor1.Text = "Cumple"
                        Else
                    Valor1.Text = "No Cumple"
                End If
                ValorNumero2.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero1.Text = ""
    End If
 Rem BY NAN PARA PELLI 2-7-2015
    If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
      Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
           KeyAscii = 0
        End If
    End If
 Rem BY NAN
End Sub



Private Sub ValorNumero2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(2)), ".")
            ZNumeII = Len(Trim(ZDesde(2)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero2.Text = Pusing("###,###.#", ValorNumero2.Text)
                Case 2
                    ValorNumero2.Text = Pusing("###,###.##", ValorNumero2.Text)
                Case 3
                    ValorNumero2.Text = Pusing("###,###.###", ValorNumero2.Text)
                Case 4
                    ValorNumero2.Text = Pusing("###,###.####", ValorNumero2.Text)
                Case 5
                    ValorNumero2.Text = Pusing("###,###.#####", ValorNumero2.Text)
                Case 6
                    ValorNumero2.Text = Pusing("###,###.######", ValorNumero2.Text)
                Case Else
                    ValorNumero2.Text = Pusing("###,###", ValorNumero2.Text)
            End Select
            
            valor2.Text = ValorNumero2.Text + " " + ZUnidad(2)
            
            ValorNumero3.SetFocus
            
                Else
                
            If ValorNumero2.Text = "S" Or ValorNumero2.Text = "N" Then
                If ValorNumero2.Text = "S" Then
                    valor2.Text = "Cumple"
                        Else
                    valor2.Text = "No Cumple"
                End If
                ValorNumero3.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero2.Text = ""
    End If
    
    If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(3)), ".")
            ZNumeII = Len(Trim(ZDesde(3)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero3.Text = Pusing("###,###.#", ValorNumero3.Text)
                Case 2
                    ValorNumero3.Text = Pusing("###,###.##", ValorNumero3.Text)
                Case 3
                    ValorNumero3.Text = Pusing("###,###.###", ValorNumero3.Text)
                Case 4
                    ValorNumero3.Text = Pusing("###,###.####", ValorNumero3.Text)
                Case 5
                    ValorNumero3.Text = Pusing("###,###.#####", ValorNumero3.Text)
                Case 6
                    ValorNumero3.Text = Pusing("###,###.######", ValorNumero3.Text)
                Case Else
                    ValorNumero3.Text = Pusing("###,###", ValorNumero3.Text)
            End Select
            
            Valor3.Text = ValorNumero3.Text + " " + ZUnidad(3)
            
            ValorNumero4.SetFocus
            
                Else
                
            If ValorNumero3.Text = "S" Or ValorNumero3.Text = "N" Then
                If ValorNumero3.Text = "S" Then
                    Valor3.Text = "Cumple"
                        Else
                    Valor3.Text = "No Cumple"
                End If
                ValorNumero4.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero3.Text = ""
    End If
    
    If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(4)), ".")
            ZNumeII = Len(Trim(ZDesde(4)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero4.Text = Pusing("###,###.#", ValorNumero4.Text)
                Case 2
                    ValorNumero4.Text = Pusing("###,###.##", ValorNumero4.Text)
                Case 3
                    ValorNumero4.Text = Pusing("###,###.###", ValorNumero4.Text)
                Case 4
                    ValorNumero4.Text = Pusing("###,###.####", ValorNumero4.Text)
                Case 5
                    ValorNumero4.Text = Pusing("###,###.#####", ValorNumero4.Text)
                Case 6
                    ValorNumero4.Text = Pusing("###,###.######", ValorNumero4.Text)
                Case Else
                    ValorNumero4.Text = Pusing("###,###", ValorNumero4.Text)
            End Select
            
            valor4.Text = ValorNumero4.Text + " " + ZUnidad(4)
            
            ValorNumero5.SetFocus
            
                Else
                
            If ValorNumero4.Text = "S" Or ValorNumero4.Text = "N" Then
                If ValorNumero4.Text = "S" Then
                    valor4.Text = "Cumple"
                        Else
                    valor4.Text = "No Cumple"
                End If
                ValorNumero5.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero4.Text = ""
    End If
    
    If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(5)), ".")
            ZNumeII = Len(Trim(ZDesde(5)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero5.Text = Pusing("###,###.#", ValorNumero5.Text)
                Case 2
                    ValorNumero5.Text = Pusing("###,###.##", ValorNumero5.Text)
                Case 3
                    ValorNumero5.Text = Pusing("###,###.###", ValorNumero5.Text)
                Case 4
                    ValorNumero5.Text = Pusing("###,###.####", ValorNumero5.Text)
                Case 5
                    ValorNumero5.Text = Pusing("###,###.#####", ValorNumero5.Text)
                Case 6
                    ValorNumero5.Text = Pusing("###,###.######", ValorNumero5.Text)
                Case Else
                    ValorNumero5.Text = Pusing("###,###", ValorNumero5.Text)
            End Select
            
            valor5.Text = ValorNumero5.Text + " " + ZUnidad(5)
            
            ValorNumero6.SetFocus
            
                Else
                
            If ValorNumero5.Text = "S" Or ValorNumero5.Text = "N" Then
                If ValorNumero5.Text = "S" Then
                    valor5.Text = "Cumple"
                        Else
                    valor5.Text = "No Cumple"
                End If
                ValorNumero6.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero5.Text = ""
    End If
    
    If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(6)), ".")
            ZNumeII = Len(Trim(ZDesde(6)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero6.Text = Pusing("###,###.#", ValorNumero6.Text)
                Case 2
                    ValorNumero6.Text = Pusing("###,###.##", ValorNumero6.Text)
                Case 3
                    ValorNumero6.Text = Pusing("###,###.###", ValorNumero6.Text)
                Case 4
                    ValorNumero6.Text = Pusing("###,###.####", ValorNumero6.Text)
                Case 5
                    ValorNumero6.Text = Pusing("###,###.#####", ValorNumero6.Text)
                Case 6
                    ValorNumero6.Text = Pusing("###,###.######", ValorNumero6.Text)
                Case Else
                    ValorNumero6.Text = Pusing("###,###", ValorNumero6.Text)
            End Select
            
            valor6.Text = ValorNumero6.Text + " " + ZUnidad(6)
            
            ValorNumero7.SetFocus
            
                Else
                
            If ValorNumero6.Text = "S" Or ValorNumero6.Text = "N" Then
                If ValorNumero6.Text = "S" Then
                    valor6.Text = "Cumple"
                        Else
                    valor6.Text = "No Cumple"
                End If
                ValorNumero7.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero6.Text = ""
    End If
    
    If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(7)), ".")
            ZNumeII = Len(Trim(ZDesde(7)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero7.Text = Pusing("###,###.#", ValorNumero7.Text)
                Case 2
                    ValorNumero7.Text = Pusing("###,###.##", ValorNumero7.Text)
                Case 3
                    ValorNumero7.Text = Pusing("###,###.###", ValorNumero7.Text)
                Case 4
                    ValorNumero7.Text = Pusing("###,###.####", ValorNumero7.Text)
                Case 5
                    ValorNumero7.Text = Pusing("###,###.#####", ValorNumero7.Text)
                Case 6
                    ValorNumero7.Text = Pusing("###,###.######", ValorNumero7.Text)
                Case Else
                    ValorNumero7.Text = Pusing("###,###", ValorNumero7.Text)
            End Select
            
            valor7.Text = ValorNumero7.Text + " " + ZUnidad(7)
            
            ValorNumero8.SetFocus
            
                Else
                
            If ValorNumero7.Text = "S" Or ValorNumero7.Text = "N" Then
                If ValorNumero7.Text = "S" Then
                    valor7.Text = "Cumple"
                        Else
                    valor7.Text = "No Cumple"
                End If
                ValorNumero8.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero7.Text = ""
    End If
    
    If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(8)), ".")
            ZNumeII = Len(Trim(ZDesde(8)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero8.Text = Pusing("###,###.#", ValorNumero8.Text)
                Case 2
                    ValorNumero8.Text = Pusing("###,###.##", ValorNumero8.Text)
                Case 3
                    ValorNumero8.Text = Pusing("###,###.###", ValorNumero8.Text)
                Case 4
                    ValorNumero8.Text = Pusing("###,###.####", ValorNumero8.Text)
                Case 5
                    ValorNumero8.Text = Pusing("###,###.#####", ValorNumero8.Text)
                Case 6
                    ValorNumero8.Text = Pusing("###,###.######", ValorNumero8.Text)
                Case Else
                    ValorNumero8.Text = Pusing("###,###", ValorNumero8.Text)
            End Select
            
            valor8.Text = ValorNumero8.Text + " " + ZUnidad(8)
            
            ValorNumero9.SetFocus
            
                Else
                
            If ValorNumero8.Text = "S" Or ValorNumero8.Text = "N" Then
                If ValorNumero8.Text = "S" Then
                    valor8.Text = "Cumple"
                        Else
                    valor8.Text = "No Cumple"
                End If
                ValorNumero9.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero8.Text = ""
    End If
    
    If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(9)), ".")
            ZNumeII = Len(Trim(ZDesde(9)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero9.Text = Pusing("###,###.#", ValorNumero9.Text)
                Case 2
                    ValorNumero9.Text = Pusing("###,###.##", ValorNumero9.Text)
                Case 3
                    ValorNumero9.Text = Pusing("###,###.###", ValorNumero9.Text)
                Case 4
                    ValorNumero9.Text = Pusing("###,###.####", ValorNumero9.Text)
                Case 5
                    ValorNumero9.Text = Pusing("###,###.#####", ValorNumero9.Text)
                Case 6
                    ValorNumero9.Text = Pusing("###,###.######", ValorNumero9.Text)
                Case Else
                    ValorNumero9.Text = Pusing("###,###", ValorNumero9.Text)
            End Select
            
            valor9.Text = ValorNumero9.Text + " " + ZUnidad(9)
            
            ValorNumero10.SetFocus
            
                Else
                
            If ValorNumero9.Text = "S" Or ValorNumero9.Text = "N" Then
                If ValorNumero9.Text = "S" Then
                    valor9.Text = "Cumple"
                        Else
                    valor9.Text = "No Cumple"
                End If
                ValorNumero10.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero9.Text = ""
    End If
    
    If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(10)), ".")
            ZNumeII = Len(Trim(ZDesde(10)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero10.Text = Pusing("###,###.#", ValorNumero10.Text)
                Case 2
                    ValorNumero10.Text = Pusing("###,###.##", ValorNumero10.Text)
                Case 3
                    ValorNumero10.Text = Pusing("###,###.###", ValorNumero10.Text)
                Case 4
                    ValorNumero10.Text = Pusing("###,###.####", ValorNumero10.Text)
                Case 5
                    ValorNumero10.Text = Pusing("###,###.#####", ValorNumero10.Text)
                Case 6
                    ValorNumero10.Text = Pusing("###,###.######", ValorNumero10.Text)
                Case Else
                    ValorNumero10.Text = Pusing("###,###", ValorNumero10.Text)
            End Select
            
            valor10.Text = ValorNumero10.Text + " " + ZUnidad(10)
            
            ValorNumero1.SetFocus
            
                Else
                
            If ValorNumero10.Text = "S" Or ValorNumero10.Text = "N" Then
                If ValorNumero10.Text = "S" Then
                    valor10.Text = "Cumple"
                        Else
                    valor10.Text = "No Cumple"
                End If
                ValorNumero1.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero10.Text = ""
    End If
    
    If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
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

    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    If Left$(Producto.Text, 2) = "DW" Then
        WProducto = "DW" + Mid$(Producto.Text, 3, 10)
            Else
        If Left$(Producto.Text, 2) = "SE" Then
            WProducto = "SE" + Mid$(Producto.Text, 3, 10)
                Else
            WProducto = "PT" + Mid$(Producto.Text, 3, 10)
        End If
    End If
    
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    llamaImprime = "N"
    
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
    Std11.Caption = ""
    Std22.Caption = ""
    Std33.Caption = ""
    Std44.Caption = ""
    Std55.Caption = ""
    Std66.Caption = ""
    Std77.Caption = ""
    Std88.Caption = ""
    Std99.Caption = ""
    Std1010.Caption = ""
    
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
    
    
    
    
    
    
    
    
    
    
    
    
    
                
    ZEnsayo1 = ""
    ZEnsayo2 = ""
    ZEnsayo3 = ""
    ZEnsayo4 = ""
    ZEnsayo5 = ""
    ZEnsayo6 = ""
    ZEnsayo7 = ""
    ZEnsayo8 = ""
    ZEnsayo9 = ""
    ZEnsayo10 = ""
    ZStd1 = ""
    ZStd2 = ""
    ZStd3 = ""
    ZStd4 = ""
    ZStd5 = ""
    ZStd6 = ""
    ZStd7 = ""
    ZStd8 = ""
    ZStd9 = ""
    ZStd10 = ""
    ZStd11 = ""
    ZStd22 = ""
    ZStd33 = ""
    ZStd44 = ""
    ZStd55 = ""
    ZStd66 = ""
    ZStd77 = ""
    ZStd88 = ""
    ZStd99 = ""
    ZStd1010 = ""
    ZVersion = 0
                
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecifUnificaVersion"
    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                
    spEspecifUnificaVersion = ZSql
    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnificaVersion.RecordCount > 0 Then
        With rstEspecifUnificaVersion
            .MoveFirst
            Do
                If .EOF = False Then
                            
                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                        ZEnsayo1 = rstEspecifUnificaVersion!Ensayo1
                        ZEnsayo2 = rstEspecifUnificaVersion!Ensayo2
                        ZEnsayo3 = rstEspecifUnificaVersion!Ensayo3
                        ZEnsayo4 = rstEspecifUnificaVersion!Ensayo4
                        ZEnsayo5 = rstEspecifUnificaVersion!Ensayo5
                        ZEnsayo6 = rstEspecifUnificaVersion!Ensayo6
                        ZEnsayo7 = rstEspecifUnificaVersion!Ensayo7
                        ZEnsayo8 = rstEspecifUnificaVersion!Ensayo8
                        ZEnsayo9 = rstEspecifUnificaVersion!Ensayo9
                        ZEnsayo10 = rstEspecifUnificaVersion!Ensayo10
                        ZStd1 = rstEspecifUnificaVersion!Valor1
                        ZStd2 = rstEspecifUnificaVersion!valor2
                        ZStd3 = rstEspecifUnificaVersion!Valor3
                        ZStd4 = rstEspecifUnificaVersion!valor4
                        ZStd5 = rstEspecifUnificaVersion!valor5
                        ZStd6 = rstEspecifUnificaVersion!valor6
                        ZStd7 = rstEspecifUnificaVersion!valor7
                        ZStd8 = rstEspecifUnificaVersion!valor8
                        ZStd9 = rstEspecifUnificaVersion!valor9
                        ZStd10 = rstEspecifUnificaVersion!valor10
                        ZStd11 = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                        ZStd22 = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                        ZStd33 = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                        ZStd44 = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                        ZStd55 = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                        ZStd66 = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                        ZStd77 = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                        ZStd88 = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                        ZStd99 = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                        ZStd1010 = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                        ZVersion = rstEspecifUnificaVersion!Version
                        llamaImprime = "S"
                    End If
                                
                    If WDesde > WFechaord And llamaImprime = "N" Then
                        ZEnsayo1 = rstEspecifUnificaVersion!Ensayo1
                        ZEnsayo2 = rstEspecifUnificaVersion!Ensayo2
                        ZEnsayo3 = rstEspecifUnificaVersion!Ensayo3
                        ZEnsayo4 = rstEspecifUnificaVersion!Ensayo4
                        ZEnsayo5 = rstEspecifUnificaVersion!Ensayo5
                        ZEnsayo6 = rstEspecifUnificaVersion!Ensayo6
                        ZEnsayo7 = rstEspecifUnificaVersion!Ensayo7
                        ZEnsayo8 = rstEspecifUnificaVersion!Ensayo8
                        ZEnsayo9 = rstEspecifUnificaVersion!Ensayo9
                        ZEnsayo10 = rstEspecifUnificaVersion!Ensayo10
                        ZStd1 = rstEspecifUnificaVersion!Valor1
                        ZStd2 = rstEspecifUnificaVersion!valor2
                        ZStd3 = rstEspecifUnificaVersion!Valor3
                        ZStd4 = rstEspecifUnificaVersion!valor4
                        ZStd5 = rstEspecifUnificaVersion!valor5
                        ZStd6 = rstEspecifUnificaVersion!valor6
                        ZStd7 = rstEspecifUnificaVersion!valor7
                        ZStd8 = rstEspecifUnificaVersion!valor8
                        ZStd9 = rstEspecifUnificaVersion!valor9
                        ZStd10 = rstEspecifUnificaVersion!valor10
                        ZStd11 = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                        ZStd22 = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                        ZStd33 = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                        ZStd44 = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                        ZStd55 = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                        ZStd66 = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                        ZStd77 = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                        ZStd88 = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                        ZStd99 = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                        ZStd1010 = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                        ZVersion = rstEspecifUnificaVersion!Version
                        llamaImprime = "S"
                    End If
                                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnificaVersion.Close
    End If
    
    Rem BY NAN 28-6-2013
    llamaImprime = "N"
    Rem ??????????????
    
    If llamaImprime = "N" Then
                
        Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,Ensayo4,Ensayo5,Ensayo6,Ensayo7,Ensayo8,Ensayo9,Ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            ZEnsayo1 = rstEspecifUnifica!Ensayo1
            ZEnsayo2 = rstEspecifUnifica!Ensayo2
            ZEnsayo3 = rstEspecifUnifica!Ensayo3
            ZEnsayo4 = rstEspecifUnifica!Ensayo4
            ZEnsayo5 = rstEspecifUnifica!Ensayo5
            ZEnsayo6 = rstEspecifUnifica!Ensayo6
            ZEnsayo7 = rstEspecifUnifica!Ensayo7
            ZEnsayo8 = rstEspecifUnifica!Ensayo8
            ZEnsayo9 = rstEspecifUnifica!Ensayo9
            ZEnsayo10 = rstEspecifUnifica!Ensayo10
            ZStd1 = rstEspecifUnifica!Valor1
            ZStd2 = rstEspecifUnifica!valor2
            ZStd3 = rstEspecifUnifica!Valor3
            ZStd4 = rstEspecifUnifica!valor4
            ZStd5 = rstEspecifUnifica!valor5
            ZStd6 = rstEspecifUnifica!valor6
            ZStd7 = rstEspecifUnifica!valor7
            ZStd8 = rstEspecifUnifica!valor8
            ZStd9 = rstEspecifUnifica!valor9
            ZStd10 = rstEspecifUnifica!valor10
            ZStd11 = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
            ZStd22 = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
            ZStd33 = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
            ZStd44 = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
            ZStd55 = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
            ZStd66 = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
            ZStd77 = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
            ZStd88 = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
            ZStd99 = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
            ZStd1010 = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
            rstEspecifUnifica.Close
        End If
        
        Rem 2 agregado
            
                    
        Sql1 = "Select Desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,desde8,Desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Version"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + WProducto + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            Rem fin
            ZDesde(1) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZDesde(2) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZDesde(3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZDesde(4) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZDesde(5) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZDesde(6) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZDesde(7) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZDesde(8) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZDesde(9) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZDesde(10) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            
            ZHasta(1) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZHasta(2) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZHasta(3) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZHasta(4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZHasta(5) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZHasta(6) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZHasta(7) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZHasta(8) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZHasta(9) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZHasta(10) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            ZDesde(1) = Trim(ZDesde(1))
            ZDesde(2) = Trim(ZDesde(2))
            ZDesde(3) = Trim(ZDesde(3))
            ZDesde(4) = Trim(ZDesde(4))
            ZDesde(5) = Trim(ZDesde(5))
            ZDesde(6) = Trim(ZDesde(6))
            ZDesde(7) = Trim(ZDesde(7))
            ZDesde(8) = Trim(ZDesde(8))
            ZDesde(9) = Trim(ZDesde(9))
            ZDesde(10) = Trim(ZDesde(10))
            
            ZHasta(1) = Trim(ZHasta(1))
            ZHasta(2) = Trim(ZHasta(2))
            ZHasta(3) = Trim(ZHasta(3))
            ZHasta(4) = Trim(ZHasta(4))
            ZHasta(5) = Trim(ZHasta(5))
            ZHasta(6) = Trim(ZHasta(6))
            ZHasta(7) = Trim(ZHasta(7))
            ZHasta(8) = Trim(ZHasta(8))
            ZHasta(9) = Trim(ZHasta(9))
            ZHasta(10) = Trim(ZHasta(10))
            
            ZVersion = rstEspecifUnifica!Version
            rstEspecifUnifica.Close
            llamaImprime = "S"
        End If
                
    End If
                
    If llamaImprime = "S" Then
        Ensayo1.Caption = ZEnsayo1
        Ensayo2.Caption = ZEnsayo2
        Ensayo3.Caption = ZEnsayo3
        Ensayo4.Caption = ZEnsayo4
        Ensayo5.Caption = ZEnsayo5
        Ensayo6.Caption = ZEnsayo6
        Ensayo7.Caption = ZEnsayo7
        Ensayo8.Caption = ZEnsayo8
        Ensayo9.Caption = ZEnsayo9
        Ensayo10.Caption = ZEnsayo10
        
        Call ImprimeII_Click
        
        If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
            ZStd1 = Trim(ZDesde(1)) + " - " + Trim(ZHasta(1)) + " " + Trim(ZUnidad(1)) + " " + Left$(ZStd1, 50)
        End If
        If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
            ZStd2 = Trim(ZDesde(2)) + " - " + Trim(ZHasta(2)) + " " + Trim(ZUnidad(2)) + " " + Left$(ZStd2, 50)
        End If
        If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
            ZStd3 = Trim(ZDesde(3)) + " - " + Trim(ZHasta(3)) + " " + Trim(ZUnidad(3)) + " " + Left$(ZStd3, 50)
        End If
        If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
            ZStd4 = Trim(ZDesde(4)) + " - " + Trim(ZHasta(4)) + " " + Trim(ZUnidad(4)) + " " + Left$(ZStd4, 50)
        End If
        If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
            ZStd5 = Trim(ZDesde(5)) + " - " + Trim(ZHasta(5)) + " " + Trim(ZUnidad(5)) + " " + Left$(ZStd5, 50)
        End If
        If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
            ZStd6 = Trim(ZDesde(6)) + " - " + Trim(ZHasta(6)) + " " + Trim(ZUnidad(6)) + " " + Left$(ZStd6, 50)
        End If
        If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
            ZStd7 = Trim(ZDesde(7)) + " - " + Trim(ZHasta(7)) + " " + Trim(ZUnidad(7)) + " " + Left$(ZStd7, 50)
        End If
        If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
            ZStd8 = Trim(ZDesde(8)) + " - " + Trim(ZHasta(8)) + " " + Trim(ZUnidad(8)) + " " + Left$(ZStd8, 50)
        End If
        If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
            ZStd9 = Trim(ZDesde(9)) + " - " + Trim(ZHasta(9)) + " " + Trim(ZUnidad(9)) + " " + Left$(ZStd9, 50)
        End If
        If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
            ZStd10 = Trim(ZDesde(10)) + " - " + Trim(ZHasta(10)) + " " + Trim(ZUnidad(10)) + " " + Left$(ZStd10, 50)
        End If
        
        Std1.Caption = ZStd1
        Std2.Caption = ZStd2
        Std3.Caption = ZStd3
        Std4.Caption = ZStd4
        Std5.Caption = ZStd5
        Std6.Caption = ZStd6
        Std7.Caption = ZStd7
        Std8.Caption = ZStd8
        Std9.Caption = ZStd9
        Std10.Caption = ZStd10
        
        Std11.Caption = ZStd11
        Std22.Caption = ZStd22
        Std33.Caption = ZStd33
        Std44.Caption = ZStd44
        Std55.Caption = ZStd55
        Std66.Caption = ZStd66
        Std77.Caption = ZStd77
        Std88.Caption = ZStd88
        Std99.Caption = ZStd99
        Std1010.Caption = ZStd1010
    End If
                        
    
    
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
    
    Call Conecta_Empresa
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM PrueTer"
    ZSql = ZSql & " Where PrueTer.Lote = " + "'" + Hoja.Text + "'"
    spPrueter = ZSql
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
        
        ValorNumero1.Text = IIf(IsNull(rstPrueter!ValorNumero1), "", rstPrueter!ValorNumero1)
        ValorNumero2.Text = IIf(IsNull(rstPrueter!ValorNumero2), "", rstPrueter!ValorNumero2)
        ValorNumero3.Text = IIf(IsNull(rstPrueter!ValorNumero3), "", rstPrueter!ValorNumero3)
        ValorNumero4.Text = IIf(IsNull(rstPrueter!ValorNumero4), "", rstPrueter!ValorNumero4)
        ValorNumero5.Text = IIf(IsNull(rstPrueter!ValorNumero5), "", rstPrueter!ValorNumero5)
        ValorNumero6.Text = IIf(IsNull(rstPrueter!ValorNumero6), "", rstPrueter!ValorNumero6)
        ValorNumero7.Text = IIf(IsNull(rstPrueter!ValorNumero7), "", rstPrueter!ValorNumero7)
        ValorNumero8.Text = IIf(IsNull(rstPrueter!ValorNumero8), "", rstPrueter!ValorNumero8)
        ValorNumero9.Text = IIf(IsNull(rstPrueter!ValorNumero9), "", rstPrueter!ValorNumero9)
        ValorNumero10.Text = IIf(IsNull(rstPrueter!ValorNumero10), "", rstPrueter!ValorNumero10)
        
        ValorNumero1.Text = Trim(ValorNumero1.Text)
        ValorNumero2.Text = Trim(ValorNumero2.Text)
        ValorNumero3.Text = Trim(ValorNumero3.Text)
        ValorNumero4.Text = Trim(ValorNumero4.Text)
        ValorNumero5.Text = Trim(ValorNumero5.Text)
        ValorNumero6.Text = Trim(ValorNumero6.Text)
        ValorNumero7.Text = Trim(ValorNumero7.Text)
        ValorNumero8.Text = Trim(ValorNumero8.Text)
        ValorNumero9.Text = Trim(ValorNumero9.Text)
        ValorNumero10.Text = Trim(ValorNumero10.Text)
        
        Ensayo.Text = rstPrueter!Ensayo
        Aspecto.Text = rstPrueter!Aspecto
        Observaciones.Text = rstPrueter!Observaciones
        Confecciono.Text = rstPrueter!Confecciono
        
        rstPrueter.Close
        
    End If

End Sub


Private Sub Limpia_Vector()

    Vector.Clear
    Vector.Font.Bold = True

    Vector.FixedCols = 1
    Vector.Cols = 6
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
                Vector.ColWidth(Ciclo) = 1450
                Vector.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    WAncho = 340
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
            a% = MsgBox(m$, 0, "Anulacion de Hoja de Produccion")
            WClave.SetFocus
        End If
    End If

End Sub

Private Sub ImprimeII_Click()

    Erase ZUnidad
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        ZUnidad(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        descri2.Caption = rstEnsayo!Descripcion
        ZUnidad(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        ZUnidad(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        ZUnidad(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        ZUnidad(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        ZUnidad(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        ZUnidad(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        ZUnidad(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        ZUnidad(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        ZUnidad(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If

End Sub


Private Sub Calcula_Bloqueo()
                
    ZZBloqueo = 0
    Rem dada
    
    If ZZTipoBloqueo = "M" Then
    
        Rem PROCESA LOS LAUDOS
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            With rstLaudo
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstLaudo.Close
        End If
    
        Rem PROCESA LOS GUIAS DE TRASLADO INTERN
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstMovguia.Close
        End If
        
            Else
    
        Rem PROCESA LAS HOJAS
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Producto = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Renglon = 1"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            With rstHoja
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstHoja.Close
        End If
    
        Rem PROCESA LOS GUIAS DE TRASLADO INTERN
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstMovguia.Close
        End If
            
    End If
    
End Sub

Private Sub Calcula_Stock_Vencido()

    Erase ZArti
    ZLugar = 0
    WWVencido = 0
    If WWArticulo = "AA-100-100" Then
        Exit Sub
    End If
    
    Rem PROCESA LOS LAUDOS
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Articulo = " + "'" + WWArticulo + "'"
    ZSql = ZSql + " and Saldo <> 0"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                        Else
                    
                    XXLaudo = rstLaudo!Laudo
                    XXSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    Call Redondeo(XXSaldo)
                    XXArticulo = rstLaudo!Articulo
                    XXCantidad = rstLaudo!Liberada
                    XXFecha = rstLaudo!Fecha
                    XXClave = rstLaudo!Clave
                    
                    If XXSaldo <> 0 Then
                        ZLugar = ZLugar + 1
                        ZArti(ZLugar, 1) = XXLaudo
                        ZArti(ZLugar, 2) = XXArticulo
                        ZArti(ZLugar, 3) = Str$(XXCantidad)
                        ZArti(ZLugar, 4) = Str$(XXSaldo)
                        ZArti(ZLugar, 5) = XXFecha
                        ZArti(ZLugar, 6) = XXClave
                        ZArti(ZLugar, 7) = "L"
                    End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Articulo = " + "'" + WWArticulo + "'"
    ZSql = ZSql + " and Saldo <> 0"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 And rstMovguia!Codigo > 900000 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" Then
                    
                        XXLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        XXArticulo = rstMovguia!Articulo
                        XXCantidad = rstMovguia!Cantidad
                        XXSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(XXSaldo)
                        XXFecha = rstMovguia!Fecha
                        XXClave = rstMovguia!Clave
                        
                        If XXSaldo <> 0 Then
                            ZLugar = ZLugar + 1
                            ZArti(ZLugar, 1) = XXLaudo
                            ZArti(ZLugar, 2) = XXArticulo
                            ZArti(ZLugar, 3) = Str$(XXCantidad)
                            ZArti(ZLugar, 4) = Str$(XXSaldo)
                            ZArti(ZLugar, 5) = XXFecha
                            ZArti(ZLugar, 6) = XXClave
                            ZArti(ZLugar, 7) = "G"
                        End If
                        
                    End If
                End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZVto = ""
        
        ZLaudo = ZArti(Ciclo, 1)
        ZArticulo = ZArti(Ciclo, 2)
        ZCantidad = Val(ZArti(Ciclo, 3))
        ZSaldo = Val(ZArti(Ciclo, 4))
        ZFecha = ZArti(Ciclo, 5)
        ZClave = ZArti(Ciclo, 6)
        ZTipo = ZArti(Ciclo, 7)
        ZMarcaVencida = ""
        
        XEmpresa = Wempresa
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
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
                ZFechaVto = IIf(IsNull(rstLaudo!fechavencimiento), "", rstLaudo!fechavencimiento)
                rstLaudo.Close
                Exit For
            End If
            
        Next Ciclo2
            
        Call Conecta_Empresa
    
        ZVto = ""
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
            If Left$(ZVto, 2) > "28" Then
                ZZComparaII = "28" + Mid$(ZVto, 3, 8)
                    Else
                ZZComparaII = ZVto
            End If
        
            ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
        
            If Val(ZDias) < 0 Then
                ZMarcaVencida = "S"
                WWVencido = WWVencido + ZSaldo
            End If
        
        End If
        
        If ZTipo = "L" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Laudo SET "
            ZSql = ZSql + " MarcaVencida = " + "'" + ZMarcaVencida + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "UPDATE Guia SET "
            ZSql = ZSql + " MarcaVencida = " + "'" + ZMarcaVencida + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Ciclo
    
End Sub
                



