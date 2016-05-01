VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspeUnifica 
   Caption         =   "Ingreso de Especificaciones de Productos Terminados (Unificado)"
   ClientHeight    =   9060
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   11730
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   11730
   Begin VB.CommandButton CaratulaII 
      Caption         =   "Caratula Total"
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
      Left            =   7800
      TabIndex        =   148
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command111 
      Caption         =   "Prueba seleccion"
      Height          =   375
      Left            =   2400
      TabIndex        =   147
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame PantaIngles 
      Height          =   2895
      Left            =   4680
      TabIndex        =   113
      Top             =   0
      Visible         =   0   'False
      Width           =   2460
      Begin VB.CommandButton AceptaIngles 
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
         Left            =   2280
         TabIndex        =   135
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox Valor1Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   133
         Text            =   " "
         Top             =   600
         Width           =   5040
      End
      Begin VB.TextBox Valor2Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   132
         Text            =   " "
         Top             =   1125
         Width           =   5040
      End
      Begin VB.TextBox Valor3Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   131
         Text            =   " "
         Top             =   1800
         Width           =   5040
      End
      Begin VB.TextBox Valor4Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   130
         Text            =   " "
         Top             =   2400
         Width           =   5040
      End
      Begin VB.TextBox Valor5Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   129
         Text            =   " "
         Top             =   3000
         Width           =   5040
      End
      Begin VB.TextBox Valor6Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   128
         Text            =   " "
         Top             =   3600
         Width           =   5040
      End
      Begin VB.TextBox Valor7Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   127
         Text            =   " "
         Top             =   4200
         Width           =   5040
      End
      Begin VB.TextBox Valor8Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   126
         Text            =   " "
         Top             =   4800
         Width           =   5040
      End
      Begin VB.TextBox Valor9Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   125
         Text            =   " "
         Top             =   5400
         Width           =   5040
      End
      Begin VB.TextBox Valor10Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   124
         Text            =   " "
         Top             =   6000
         Width           =   5040
      End
      Begin VB.TextBox Valor11Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   123
         Text            =   " "
         Top             =   840
         Width           =   5040
      End
      Begin VB.TextBox Valor22Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   122
         Text            =   " "
         Top             =   1440
         Width           =   5040
      End
      Begin VB.TextBox Valor33Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   121
         Text            =   " "
         Top             =   2040
         Width           =   5040
      End
      Begin VB.TextBox Valor44Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   120
         Text            =   " "
         Top             =   2640
         Width           =   5040
      End
      Begin VB.TextBox Valor55Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   119
         Text            =   " "
         Top             =   3240
         Width           =   5040
      End
      Begin VB.TextBox Valor66Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   118
         Text            =   " "
         Top             =   3840
         Width           =   5040
      End
      Begin VB.TextBox Valor77Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   117
         Text            =   " "
         Top             =   4440
         Width           =   5040
      End
      Begin VB.TextBox Valor88Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   116
         Text            =   " "
         Top             =   5040
         Width           =   5040
      End
      Begin VB.TextBox Valor99Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   115
         Text            =   " "
         Top             =   5640
         Width           =   5040
      End
      Begin VB.TextBox Valor1010Ing 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   20
         MaxLength       =   50
         TabIndex        =   114
         Text            =   " "
         Top             =   6240
         Width           =   5040
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard en Ingles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   20
         TabIndex        =   134
         Top             =   240
         Width           =   5040
      End
   End
   Begin VB.CommandButton Command22 
      Caption         =   "listado farmacopea"
      Height          =   615
      Left            =   5880
      TabIndex        =   146
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   2280
      TabIndex        =   50
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1920
         TabIndex        =   61
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1920
         TabIndex        =   60
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.OptionButton ImpreListado 
         Caption         =   "Option2"
         Height          =   195
         Left            =   2640
         TabIndex        =   56
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton ImprePantalla 
         Caption         =   "Option1"
         Height          =   195
         Left            =   2640
         TabIndex        =   55
         Top             =   960
         Width           =   255
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
         Left            =   1320
         TabIndex        =   54
         Top             =   960
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
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   3000
         TabIndex        =   58
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   3000
         TabIndex        =   57
         Top             =   960
         Width           =   735
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
         Left            =   480
         TabIndex        =   52
         Top             =   600
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
         Left            =   480
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
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
      Height          =   1260
      Left            =   120
      TabIndex        =   49
      Top             =   6720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   145
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3840
      TabIndex        =   142
      Top             =   7200
      Width           =   255
   End
   Begin VB.Frame XClaveIII 
      Height          =   1935
      Left            =   3840
      TabIndex        =   137
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGrabaIII 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   139
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClaveIII 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   138
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
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
         TabIndex        =   140
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox ControlCambio 
      BeginProperty Font 
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
      MaxLength       =   100
      TabIndex        =   112
      Text            =   " "
      Top             =   6720
      Width           =   5760
   End
   Begin VB.TextBox Desde1 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   108
      Text            =   " "
      Top             =   720
      Width           =   840
   End
   Begin VB.TextBox Hasta1 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   107
      Text            =   " "
      Top             =   720
      Width           =   840
   End
   Begin VB.TextBox Desde2 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   106
      Text            =   " "
      Top             =   1250
      Width           =   840
   End
   Begin VB.TextBox Hasta2 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   105
      Text            =   " "
      Top             =   1250
      Width           =   840
   End
   Begin VB.TextBox Desde3 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   104
      Text            =   " "
      Top             =   1920
      Width           =   840
   End
   Begin VB.TextBox Hasta3 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   103
      Text            =   " "
      Top             =   1920
      Width           =   840
   End
   Begin VB.TextBox Desde4 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   102
      Text            =   " "
      Top             =   2520
      Width           =   840
   End
   Begin VB.TextBox Hasta4 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   101
      Text            =   " "
      Top             =   2520
      Width           =   840
   End
   Begin VB.TextBox Desde5 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   100
      Text            =   " "
      Top             =   3120
      Width           =   840
   End
   Begin VB.TextBox Hasta5 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   99
      Text            =   " "
      Top             =   3120
      Width           =   840
   End
   Begin VB.TextBox Desde6 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   98
      Text            =   " "
      Top             =   3720
      Width           =   840
   End
   Begin VB.TextBox Hasta6 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   97
      Text            =   " "
      Top             =   3720
      Width           =   840
   End
   Begin VB.TextBox Desde7 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   96
      Text            =   " "
      Top             =   4320
      Width           =   840
   End
   Begin VB.TextBox Hasta7 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   95
      Text            =   " "
      Top             =   4320
      Width           =   840
   End
   Begin VB.TextBox Desde8 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   94
      Text            =   " "
      Top             =   4920
      Width           =   840
   End
   Begin VB.TextBox Hasta8 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   93
      Text            =   " "
      Top             =   4920
      Width           =   840
   End
   Begin VB.TextBox Desde9 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   92
      Text            =   " "
      Top             =   5520
      Width           =   840
   End
   Begin VB.TextBox Hasta9 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   91
      Text            =   " "
      Top             =   5520
      Width           =   840
   End
   Begin VB.TextBox Desde10 
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   90
      Text            =   " "
      Top             =   6120
      Width           =   840
   End
   Begin VB.TextBox Hasta10 
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
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   89
      Text            =   " "
      Top             =   6120
      Width           =   840
   End
   Begin VB.Frame XClaveII 
      Height          =   1935
      Left            =   3480
      TabIndex        =   83
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClaveII 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   85
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGrabaII 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   84
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         TabIndex        =   86
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
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
      Left            =   7800
      TabIndex        =   82
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Estado 
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
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   81
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3360
      TabIndex        =   76
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   78
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   77
         Top             =   960
         Width           =   1815
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
         TabIndex        =   79
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox Version 
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
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   73
      Text            =   " "
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Fecha 
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
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   72
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Valor1010 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   71
      Text            =   " "
      Top             =   6360
      Width           =   5040
   End
   Begin VB.TextBox Valor99 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   70
      Text            =   " "
      Top             =   5760
      Width           =   5040
   End
   Begin VB.TextBox Valor88 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   69
      Text            =   " "
      Top             =   5160
      Width           =   5040
   End
   Begin VB.TextBox Valor77 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   68
      Text            =   " "
      Top             =   4560
      Width           =   5040
   End
   Begin VB.TextBox Valor66 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   67
      Text            =   " "
      Top             =   3960
      Width           =   5040
   End
   Begin VB.TextBox Valor55 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   66
      Text            =   " "
      Top             =   3360
      Width           =   5040
   End
   Begin VB.TextBox Valor44 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   65
      Text            =   " "
      Top             =   2760
      Width           =   5040
   End
   Begin VB.TextBox Valor33 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   64
      Text            =   " "
      Top             =   2160
      Width           =   5040
   End
   Begin VB.TextBox Valor22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   63
      Text            =   " "
      Top             =   1560
      Width           =   5040
   End
   Begin VB.TextBox Valor11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   62
      Text            =   " "
      Top             =   960
      Width           =   5040
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   0
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
   Begin Crystal.CrystalReport Lista 
      Left            =   10680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEspefUnifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   11160
      TabIndex        =   48
      Top             =   0
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   6120
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   5520
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   4920
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   4320
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   3720
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   3120
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   2520
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   1920
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   39
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   1250
      Width           =   5040
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
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      ToolTipText     =   "En caso del resultado es un resultado numerico el sistema tomara automaticamente ese parametro y le agregara el valor a informar"
      Top             =   720
      Width           =   5040
   End
   Begin VB.TextBox Ensayo10 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   37
      Text            =   " "
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox Ensayo9 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   36
      Text            =   " "
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Ensayo8 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   35
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Ensayo7 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   34
      Text            =   " "
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Ensayo6 
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
      Height          =   315
      Left            =   0
      MaxLength       =   4
      TabIndex        =   33
      Text            =   " "
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Ensayo5 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Ensayo4 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   31
      Text            =   " "
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Ensayo3 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Ensayo2 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   29
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Ensayo1 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   28
      Text            =   " "
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
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
      Height          =   300
      Left            =   8880
      TabIndex        =   11
      Top             =   7200
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
      Left            =   8880
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   9960
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
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
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo "
         BeginProperty Font 
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
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer "
         BeginProperty Font 
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
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   8520
      TabIndex        =   4
      Top             =   5520
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
      Left            =   8880
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
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
      Height          =   300
      Left            =   6720
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
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
      Left            =   7800
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton EntraIngles 
      Caption         =   "Valor Standard en Ingles"
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
      Left            =   0
      TabIndex        =   136
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton CaratulaI 
      Caption         =   "Caratula Cliente"
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
      Left            =   7800
      TabIndex        =   141
      Top             =   7560
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   1095
      Left            =   4800
      TabIndex        =   144
      Top             =   7800
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1931
      _Version        =   327680
      Rows            =   10
      Cols            =   6
      BackColor       =   16777152
      ScrollTrack     =   -1  'True
   End
   Begin Crystal.CrystalReport lista2 
      Left            =   1080
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\listafarmacopea.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
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
      Height          =   1260
      ItemData        =   "espeunifica.frx":0000
      Left            =   1200
      List            =   "espeunifica.frx":0007
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Farmacop 
      Caption         =   "Farmacopea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   143
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Control de Cambios"
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
      Index           =   3
      Left            =   0
      TabIndex        =   111
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   110
      Top             =   360
      Width           =   840
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   109
      Top             =   360
      Width           =   840
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   88
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   87
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Estado"
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
      Left            =   6360
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Version"
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
      Left            =   2760
      TabIndex        =   75
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   4200
      TabIndex        =   74
      Top             =   0
      Width           =   855
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
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   855
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
      Left            =   840
      TabIndex        =   27
      Top             =   6120
      Width           =   4000
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
      Left            =   840
      TabIndex        =   26
      Top             =   5520
      Width           =   4000
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
      Left            =   840
      TabIndex        =   25
      Top             =   4920
      Width           =   4000
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
      Left            =   840
      TabIndex        =   24
      Top             =   4320
      Width           =   4000
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
      Left            =   840
      TabIndex        =   23
      Top             =   3720
      Width           =   4000
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
      Left            =   840
      TabIndex        =   22
      Top             =   3120
      Width           =   4000
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
      Left            =   840
      TabIndex        =   21
      Top             =   2520
      Width           =   4000
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
      Left            =   840
      TabIndex        =   20
      Top             =   1920
      Width           =   4000
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
      Left            =   840
      TabIndex        =   19
      Top             =   1320
      Width           =   4000
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
      Left            =   840
      TabIndex        =   18
      Top             =   720
      Width           =   4000
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
      Left            =   4845
      TabIndex        =   17
      Top             =   360
      Width           =   5040
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
      Left            =   840
      TabIndex        =   16
      Top             =   360
      Width           =   3840
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
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgEspeUnifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstOperador As Recordset
Dim spOperador As String

Dim XParam As String
Dim EmpresaActual As String
Dim ZFecha As String
Dim ZVersion As String
Private WGraba As String
Private WGrabaII As String
Private WGrabaIII As String
Dim CargaEmpresa(12, 2) As String
Dim ZOperador As String
Dim ZZOperador As String
Dim ZZVersion As String
Dim ZZFecha As String
Dim ZZProcesoImpre As Integer

Dim ZVector(10000) As String
Dim ZEnsayo(10) As String
Dim ZValor(20) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String


Dim ZZOpcion(10) As Integer
Dim ZZValor(10) As String
Dim ZZEnsayo(10) As String
Dim ZZStd(10, 6) As String
Dim ZZDescri(10) As String
Dim ZZDescriII(10) As String

Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String

Dim Graba(1000) As String
Dim GrabaII(1000, 40) As String
Dim XPaso As String

Dim ZZProceso As Integer

Private Sub Acepta_Click()
    
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
            
            
            
    Erase ZVector
    ZLugar = 0
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select Producto"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstEspecifUnifica!Producto >= Desde.Text And rstEspecifUnifica!Producto <= Hasta.Text Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = rstEspecifUnifica!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
        
        Erase ZEnsayo
        Erase ZValor
        Erase ZDescri
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecifUnifica"
        ZSql = ZSql + " Where EspecifUnifica.Producto = " + "'" + ZCodigo + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
    
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
        
        For Cicla = 1 To 10
            If Val(ZEnsayo(Cicla)) <> 0 Then
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    ZDescriII(Cicla) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                    ZDescriII(Cicla) = Trim(ZDescriII(Cicla))
                    rstEnsayo.Close
                End If
            End If
        Next Cicla
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecifUnifica"
        ZSql = ZSql + " Where EspecifUnifica.Producto = " + "'" + ZCodigo + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
    
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
            
            ZZDesde1 = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
            ZZDesde2 = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
            ZZDesde3 = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
            ZZDesde4 = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
            ZZDesde5 = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
            ZZDesde6 = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
            ZZDesde7 = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
            ZZDesde8 = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
            ZZDesde9 = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
            ZZDesde10 = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
            
            ZZHasta1 = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
            ZZHasta2 = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
            ZZHasta3 = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
            ZZHasta4 = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
            ZZHasta5 = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
            ZZHasta6 = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
            ZZHasta7 = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
            ZZHasta8 = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
            ZZHasta9 = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
            ZZHasta10 = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
            
            If Val(ZZDesde1) <> 0 Or Val(ZZHasta1) <> 0 Then
                ZValor(1) = Trim(ZZDesde1) + " - " + " " + Trim(ZZHasta1) + " " + ZDescriII(1) + " " + rstEspecifUnifica!Valor1
                    Else
                ZValor(1) = rstEspecifUnifica!Valor1
            End If
            If Val(ZZDesde2) <> 0 Or Val(ZZHasta2) <> 0 Then
                ZValor(2) = Trim(ZZDesde2) + " - " + " " + Trim(ZZHasta2) + " " + ZDescriII(2) + " " + rstEspecifUnifica!valor2
                    Else
                ZValor(2) = rstEspecifUnifica!valor2
            End If
            If Val(ZZDesde3) <> 0 Or Val(ZZHasta3) <> 0 Then
                ZValor(3) = Trim(ZZDesde3) + " - " + " " + Trim(ZZHasta3) + " " + ZDescriII(3) + " " + rstEspecifUnifica!Valor3
                    Else
                ZValor(3) = rstEspecifUnifica!Valor3
            End If
            If Val(ZZDesde4) <> 0 Or Val(ZZHasta4) <> 0 Then
                ZValor(4) = Trim(ZZDesde4) + " - " + " " + Trim(ZZHasta4) + " " + ZDescriII(4) + " " + rstEspecifUnifica!valor4
                    Else
                ZValor(4) = rstEspecifUnifica!valor4
            End If
            If Val(ZZDesde5) <> 0 Or Val(ZZHasta5) <> 0 Then
                ZValor(5) = Trim(ZZDesde5) + " - " + " " + Trim(ZZHasta5) + " " + ZDescriII(5) + " " + rstEspecifUnifica!valor5
                    Else
                ZValor(5) = rstEspecifUnifica!valor5
            End If
            If Val(ZZDesde6) <> 0 Or Val(ZZHasta6) <> 0 Then
                ZValor(6) = Trim(ZZDesde6) + " - " + " " + Trim(ZZHasta6) + " " + ZDescriII(6) + " " + rstEspecifUnifica!valor6
                    Else
                ZValor(6) = rstEspecifUnifica!valor6
            End If
            If Val(ZZDesde7) <> 0 Or Val(ZZHasta7) <> 0 Then
                ZValor(7) = Trim(ZZDesde7) + " - " + " " + Trim(ZZHasta7) + " " + ZDescriII(7) + " " + rstEspecifUnifica!valor7
                    Else
                ZValor(7) = rstEspecifUnifica!valor7
            End If
            If Val(ZZDesde8) <> 0 Or Val(ZZHasta8) <> 0 Then
                ZValor(8) = Trim(ZZDesde8) + " - " + " " + Trim(ZZHasta8) + " " + ZDescriII(8) + " " + rstEspecifUnifica!valor8
                    Else
                ZValor(8) = rstEspecifUnifica!valor8
            End If
            If Val(ZZDesde9) <> 0 Or Val(ZZHasta9) <> 0 Then
                ZValor(9) = Trim(ZZDesde9) + " - " + " " + Trim(ZZHasta9) + " " + ZDescriII(9) + " " + rstEspecifUnifica!valor9
                    Else
                ZValor(9) = rstEspecifUnifica!valor9
            End If
            If Val(ZZDesde10) <> 0 Or Val(ZZHasta10) <> 0 Then
                ZValor(10) = Trim(ZZDesde10) + " - " + " " + Trim(ZZHasta10) + " " + ZDescriII(10) + " " + rstEspecifUnifica!valor10
                    Else
                ZValor(10) = rstEspecifUnifica!valor10
            End If
            
            ZValor(11) = rstEspecifUnifica!Valor11
            ZValor(12) = rstEspecifUnifica!Valor22
            ZValor(13) = rstEspecifUnifica!Valor33
            ZValor(14) = rstEspecifUnifica!Valor44
            ZValor(15) = rstEspecifUnifica!Valor55
            ZValor(16) = rstEspecifUnifica!Valor66
            ZValor(17) = rstEspecifUnifica!Valor77
            ZValor(18) = rstEspecifUnifica!Valor88
            ZValor(19) = rstEspecifUnifica!Valor99
            ZValor(20) = rstEspecifUnifica!Valor1010
            
            ZZOperador = IIf(IsNull(rstEspecifUnifica!Operador), "O", rstEspecifUnifica!Operador)
            ZZVersion = rstEspecifUnifica!Version
            ZZFecha = rstEspecifUnifica!Fecha
        
            rstEspecifUnifica.Close
                        
        End If
    
        For Cicla = 1 To 10
            If Val(ZEnsayo(Cicla)) <> 0 Then
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    ZDescri(Cicla) = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
            End If
        Next Cicla
    
        spTerminado = "ConsultaTerminado " + "'" + ZCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZDescripcion = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        ZZDesOperador = ""
        If Val(ZZOperador) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.Operador = " + "'" + ZZOperador + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZZDesOperador = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
                rstOperador.Close
            End If
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaEspe ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion,"
        ZSql = ZSql + "Codigo1,"
        ZSql = ZSql + "Codigo2,"
        ZSql = ZSql + "Codigo3,"
        ZSql = ZSql + "Codigo4,"
        ZSql = ZSql + "Codigo5,"
        ZSql = ZSql + "Codigo6,"
        ZSql = ZSql + "Codigo7,"
        ZSql = ZSql + "Codigo8,"
        ZSql = ZSql + "Codigo9,"
        ZSql = ZSql + "Codigo10,"
        ZSql = ZSql + "Descri1,"
        ZSql = ZSql + "Descri2,"
        ZSql = ZSql + "Descri3,"
        ZSql = ZSql + "Descri4,"
        ZSql = ZSql + "Descri5,"
        ZSql = ZSql + "Descri6,"
        ZSql = ZSql + "Descri7,"
        ZSql = ZSql + "Descri8,"
        ZSql = ZSql + "Descri9,"
        ZSql = ZSql + "Descri10,"
        ZSql = ZSql + "Valor1,"
        ZSql = ZSql + "Valor2,"
        ZSql = ZSql + "Valor3,"
        ZSql = ZSql + "Valor4,"
        ZSql = ZSql + "Valor5,"
        ZSql = ZSql + "Valor6,"
        ZSql = ZSql + "Valor7,"
        ZSql = ZSql + "Valor8,"
        ZSql = ZSql + "Valor9,"
        ZSql = ZSql + "Valor10,"
        ZSql = ZSql + "Valor11,"
        ZSql = ZSql + "Valor22,"
        ZSql = ZSql + "Valor33,"
        ZSql = ZSql + "Valor44,"
        ZSql = ZSql + "Valor55,"
        ZSql = ZSql + "Valor66,"
        ZSql = ZSql + "Valor77,"
        ZSql = ZSql + "Valor88,"
        ZSql = ZSql + "Valor99,"
        ZSql = ZSql + "Valor1010,"
        ZSql = ZSql + "Version ,"
        ZSql = ZSql + "Responsable,"
        ZSql = ZSql + "Fecha )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZEnsayo(1) + "',"
        ZSql = ZSql + "'" + ZEnsayo(2) + "',"
        ZSql = ZSql + "'" + ZEnsayo(3) + "',"
        ZSql = ZSql + "'" + ZEnsayo(4) + "',"
        ZSql = ZSql + "'" + ZEnsayo(5) + "',"
        ZSql = ZSql + "'" + ZEnsayo(6) + "',"
        ZSql = ZSql + "'" + ZEnsayo(7) + "',"
        ZSql = ZSql + "'" + ZEnsayo(8) + "',"
        ZSql = ZSql + "'" + ZEnsayo(9) + "',"
        ZSql = ZSql + "'" + ZEnsayo(10) + "',"
        ZSql = ZSql + "'" + ZDescri(1) + "',"
        ZSql = ZSql + "'" + ZDescri(2) + "',"
        ZSql = ZSql + "'" + ZDescri(3) + "',"
        ZSql = ZSql + "'" + ZDescri(4) + "',"
        ZSql = ZSql + "'" + ZDescri(5) + "',"
        ZSql = ZSql + "'" + ZDescri(6) + "',"
        ZSql = ZSql + "'" + ZDescri(7) + "',"
        ZSql = ZSql + "'" + ZDescri(8) + "',"
        ZSql = ZSql + "'" + ZDescri(9) + "',"
        ZSql = ZSql + "'" + ZDescri(10) + "',"
        ZSql = ZSql + "'" + ZValor(1) + "',"
        ZSql = ZSql + "'" + ZValor(2) + "',"
        ZSql = ZSql + "'" + ZValor(3) + "',"
        ZSql = ZSql + "'" + ZValor(4) + "',"
        ZSql = ZSql + "'" + ZValor(5) + "',"
        ZSql = ZSql + "'" + ZValor(6) + "',"
        ZSql = ZSql + "'" + ZValor(7) + "',"
        ZSql = ZSql + "'" + ZValor(8) + "',"
        ZSql = ZSql + "'" + ZValor(9) + "',"
        ZSql = ZSql + "'" + ZValor(10) + "',"
        ZSql = ZSql + "'" + ZValor(11) + "',"
        ZSql = ZSql + "'" + ZValor(12) + "',"
        ZSql = ZSql + "'" + ZValor(13) + "',"
        ZSql = ZSql + "'" + ZValor(14) + "',"
        ZSql = ZSql + "'" + ZValor(15) + "',"
        ZSql = ZSql + "'" + ZValor(16) + "',"
        ZSql = ZSql + "'" + ZValor(17) + "',"
        ZSql = ZSql + "'" + ZValor(18) + "',"
        ZSql = ZSql + "'" + ZValor(19) + "',"
        ZSql = ZSql + "'" + ZValor(20) + "',"
        ZSql = ZSql + "'" + ZZVersion + "',"
        ZSql = ZSql + "'" + ZZDesOperador + "',"
        ZSql = ZSql + "'" + ZZFecha + "')"
        
        spListaEspe = ZSql
        Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ListaEspe SET "
    ZSql = ZSql + "ControlCambio = " + "'" + ControlCambio.Text + "'"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    Lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height

    Rem lista.GroupSelectionFormula = "{EspecifUnifica.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If ImpreListado.Value = True Then
        Lista.Destination = 1
            Else
        Lista.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(WEmpresa) = 3 Then
        Lista.ReportFileName = "ListaEspePt.rpt"
            Else
        Lista.ReportFileName = "Listaesptnuevopell.rpt"
    End If
    
    Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, ListaEspe.Valor11, ListaEspe.Valor22, ListaEspe.Valor33, ListaEspe.Valor44, ListaEspe.Valor55, ListaEspe.Valor66, ListaEspe.Valor77, ListaEspe.Valor88, ListaEspe.Valor99, ListaEspe.Valor1010, ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha, ListaEspe.ControlCambio " _
                    + "From " _
                    + DSQ + ".dbo.ListaEspe ListaEspe " _
                    + "Where " _
                    + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
                    + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    Lista.Connect = Connect()
    
    ZZImpreAnterior = Printer.DeviceName
    
    If ZZProceso = 1 Then
        Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "CutePDF Writer" + Chr$(34)
    End If
    
    Lista.Action = 1
    
    If ZZProceso = 1 Then
        Rem Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "HP Color LaserJet 2600n" + Chr$(34)
        Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + ZZImpreAnterior + Chr$(34)
    End If
    
    Frame2.Visible = False
    
    Call Conecta_Empresa
    Command22.Visible = True
End Sub


Private Sub ImprimeAutomatico()
    
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
            
            
            
    Erase ZVector
    ZLugar = 0
    
    Desde.Text = Producto.Text
    Hasta.Text = Producto.Text
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select producto"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstEspecifUnifica!Producto >= Desde.Text And rstEspecifUnifica!Producto <= Hasta.Text Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = rstEspecifUnifica!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
        
        Erase ZEnsayo
        Erase ZValor
        Erase ZDescri
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecifUnifica"
        ZSql = ZSql + " Where EspecifUnifica.Producto = " + "'" + ZCodigo + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
    
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
            
            ZValor(1) = rstEspecifUnifica!Valor1
            ZValor(2) = rstEspecifUnifica!valor2
            ZValor(3) = rstEspecifUnifica!Valor3
            ZValor(4) = rstEspecifUnifica!valor4
            ZValor(5) = rstEspecifUnifica!valor5
            ZValor(6) = rstEspecifUnifica!valor6
            ZValor(7) = rstEspecifUnifica!valor7
            ZValor(8) = rstEspecifUnifica!valor8
            ZValor(9) = rstEspecifUnifica!valor9
            ZValor(10) = rstEspecifUnifica!valor10
            
            ZValor(11) = rstEspecifUnifica!Valor11
            ZValor(12) = rstEspecifUnifica!Valor22
            ZValor(13) = rstEspecifUnifica!Valor33
            ZValor(14) = rstEspecifUnifica!Valor44
            ZValor(15) = rstEspecifUnifica!Valor55
            ZValor(16) = rstEspecifUnifica!Valor66
            ZValor(17) = rstEspecifUnifica!Valor77
            ZValor(18) = rstEspecifUnifica!Valor88
            ZValor(19) = rstEspecifUnifica!Valor99
            ZValor(20) = rstEspecifUnifica!Valor1010
            
            ZZOperador = IIf(IsNull(rstEspecifUnifica!Operador), "O", rstEspecifUnifica!Operador)
            ZZVersion = rstEspecifUnifica!Version
            ZZFecha = rstEspecifUnifica!Fecha
        
            rstEspecifUnifica.Close
                        
        End If
    
        For Cicla = 1 To 10
            If Val(ZEnsayo(Cicla)) <> 0 Then
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    ZDescri(Cicla) = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
            End If
        Next Cicla
    
        spTerminado = "ConsultaTerminado " + "'" + ZCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZDescripcion = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        ZZDesOperador = ""
        If Val(ZZOperador) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.Operador = " + "'" + ZZOperador + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZZDesOperador = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
                rstOperador.Close
            End If
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaEspe ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion,"
        ZSql = ZSql + "Codigo1,"
        ZSql = ZSql + "Codigo2,"
        ZSql = ZSql + "Codigo3,"
        ZSql = ZSql + "Codigo4,"
        ZSql = ZSql + "Codigo5,"
        ZSql = ZSql + "Codigo6,"
        ZSql = ZSql + "Codigo7,"
        ZSql = ZSql + "Codigo8,"
        ZSql = ZSql + "Codigo9,"
        ZSql = ZSql + "Codigo10,"
        ZSql = ZSql + "Descri1,"
        ZSql = ZSql + "Descri2,"
        ZSql = ZSql + "Descri3,"
        ZSql = ZSql + "Descri4,"
        ZSql = ZSql + "Descri5,"
        ZSql = ZSql + "Descri6,"
        ZSql = ZSql + "Descri7,"
        ZSql = ZSql + "Descri8,"
        ZSql = ZSql + "Descri9,"
        ZSql = ZSql + "Descri10,"
        ZSql = ZSql + "Valor1,"
        ZSql = ZSql + "Valor2,"
        ZSql = ZSql + "Valor3,"
        ZSql = ZSql + "Valor4,"
        ZSql = ZSql + "Valor5,"
        ZSql = ZSql + "Valor6,"
        ZSql = ZSql + "Valor7,"
        ZSql = ZSql + "Valor8,"
        ZSql = ZSql + "Valor9,"
        ZSql = ZSql + "Valor10,"
        ZSql = ZSql + "Valor11,"
        ZSql = ZSql + "Valor22,"
        ZSql = ZSql + "Valor33,"
        ZSql = ZSql + "Valor44,"
        ZSql = ZSql + "Valor55,"
        ZSql = ZSql + "Valor66,"
        ZSql = ZSql + "Valor77,"
        ZSql = ZSql + "Valor88,"
        ZSql = ZSql + "Valor99,"
        ZSql = ZSql + "Valor1010,"
        ZSql = ZSql + "Version ,"
        ZSql = ZSql + "Responsable,"
        ZSql = ZSql + "Fecha )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZEnsayo(1) + "',"
        ZSql = ZSql + "'" + ZEnsayo(2) + "',"
        ZSql = ZSql + "'" + ZEnsayo(3) + "',"
        ZSql = ZSql + "'" + ZEnsayo(4) + "',"
        ZSql = ZSql + "'" + ZEnsayo(5) + "',"
        ZSql = ZSql + "'" + ZEnsayo(6) + "',"
        ZSql = ZSql + "'" + ZEnsayo(7) + "',"
        ZSql = ZSql + "'" + ZEnsayo(8) + "',"
        ZSql = ZSql + "'" + ZEnsayo(9) + "',"
        ZSql = ZSql + "'" + ZEnsayo(10) + "',"
        ZSql = ZSql + "'" + ZDescri(1) + "',"
        ZSql = ZSql + "'" + ZDescri(2) + "',"
        ZSql = ZSql + "'" + ZDescri(3) + "',"
        ZSql = ZSql + "'" + ZDescri(4) + "',"
        ZSql = ZSql + "'" + ZDescri(5) + "',"
        ZSql = ZSql + "'" + ZDescri(6) + "',"
        ZSql = ZSql + "'" + ZDescri(7) + "',"
        ZSql = ZSql + "'" + ZDescri(8) + "',"
        ZSql = ZSql + "'" + ZDescri(9) + "',"
        ZSql = ZSql + "'" + ZDescri(10) + "',"
        ZSql = ZSql + "'" + ZValor(1) + "',"
        ZSql = ZSql + "'" + ZValor(2) + "',"
        ZSql = ZSql + "'" + ZValor(3) + "',"
        ZSql = ZSql + "'" + ZValor(4) + "',"
        ZSql = ZSql + "'" + ZValor(5) + "',"
        ZSql = ZSql + "'" + ZValor(6) + "',"
        ZSql = ZSql + "'" + ZValor(7) + "',"
        ZSql = ZSql + "'" + ZValor(8) + "',"
        ZSql = ZSql + "'" + ZValor(9) + "',"
        ZSql = ZSql + "'" + ZValor(10) + "',"
        ZSql = ZSql + "'" + ZValor(11) + "',"
        ZSql = ZSql + "'" + ZValor(12) + "',"
        ZSql = ZSql + "'" + ZValor(13) + "',"
        ZSql = ZSql + "'" + ZValor(14) + "',"
        ZSql = ZSql + "'" + ZValor(15) + "',"
        ZSql = ZSql + "'" + ZValor(16) + "',"
        ZSql = ZSql + "'" + ZValor(17) + "',"
        ZSql = ZSql + "'" + ZValor(18) + "',"
        ZSql = ZSql + "'" + ZValor(19) + "',"
        ZSql = ZSql + "'" + ZValor(20) + "',"
        ZSql = ZSql + "'" + ZZVersion + "',"
        ZSql = ZSql + "'" + ZZDesOperador + "',"
        ZSql = ZSql + "'" + ZZFecha + "')"
        
        spListaEspe = ZSql
        Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ListaEspe SET "
    ZSql = ZSql + "ControlCambio = " + "'" + ControlCambio.Text + "'"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    Lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height

    Rem lista.GroupSelectionFormula = "{EspecifUnifica.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Lista.Destination = 1
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(WEmpresa) = 3 Then
        Lista.ReportFileName = "ListaEspePt.rpt"
            Else
        Lista.ReportFileName = "ListaEspePtpelli.rpt"
    End If
    
    Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, ListaEspe.Valor11, ListaEspe.Valor22, ListaEspe.Valor33, ListaEspe.Valor44, ListaEspe.Valor55, ListaEspe.Valor66, ListaEspe.Valor77, ListaEspe.Valor88, ListaEspe.Valor99, ListaEspe.Valor1010, ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha, ListaEspe.ControlCambio " _
                    + "From " _
                    + DSQ + ".dbo.ListaEspe ListaEspe " _
                    + "Where " _
                    + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
                    + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    Lista.Connect = Connect()
    
    Lista.Action = 1
    Frame2.Visible = False
    
    Call Conecta_Empresa
    
End Sub

Private Sub AceptaIngles_Click()
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

    ZSql = ""
    ZSql = ZSql & "UPDATE EspecifUnifica SET "
    ZSql = ZSql & "Valor1Ing = " + "'" + Valor1Ing.Text + "',"
    ZSql = ZSql & "Valor2Ing = " + "'" + Valor2Ing.Text + "',"
    ZSql = ZSql & "Valor3Ing = " + "'" + Valor3Ing.Text + "',"
    ZSql = ZSql & "Valor4Ing = " + "'" + Valor4Ing.Text + "',"
    ZSql = ZSql & "Valor5Ing = " + "'" + Valor5Ing.Text + "',"
    ZSql = ZSql & "Valor6Ing = " + "'" + Valor6Ing.Text + "',"
    ZSql = ZSql & "Valor7Ing = " + "'" + Valor7Ing.Text + "',"
    ZSql = ZSql & "Valor8Ing = " + "'" + Valor8Ing.Text + "',"
    ZSql = ZSql & "Valor9Ing = " + "'" + Valor9Ing.Text + "',"
    ZSql = ZSql & "Valor10Ing = " + "'" + Valor10Ing.Text + "',"
    ZSql = ZSql & "Valor11Ing = " + "'" + Valor11Ing.Text + "',"
    ZSql = ZSql & "Valor22Ing = " + "'" + Valor22Ing.Text + "',"
    ZSql = ZSql & "Valor33Ing = " + "'" + Valor33Ing.Text + "',"
    ZSql = ZSql & "Valor44Ing = " + "'" + Valor44Ing.Text + "',"
    ZSql = ZSql & "Valor55Ing = " + "'" + Valor55Ing.Text + "',"
    ZSql = ZSql & "Valor66Ing = " + "'" + Valor66Ing.Text + "',"
    ZSql = ZSql & "Valor77Ing = " + "'" + Valor77Ing.Text + "',"
    ZSql = ZSql & "Valor88Ing = " + "'" + Valor88Ing.Text + "',"
    ZSql = ZSql & "Valor99Ing = " + "'" + Valor99Ing.Text + "',"
    ZSql = ZSql & "Valor1010Ing = " + "'" + Valor1010Ing.Text + "'"
    ZSql = ZSql & " Where Producto = " + "'" + Producto.Text + "'"
            
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Conecta_Empresa
    

    PantaIngles.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
Command22.Visible = True

End Sub


Private Sub Command111_Click()
            ZZPasaTerminado = Producto.Text
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    ZZPasaCliente = "S00102"
                Case Else
                    ZZPasaCliente = "P99999"
            End Select
        
            Call CmdLimpiar_Click
            
            PrgAltaCertificadoAuto.Show

End Sub

Private Sub Command22_Click()
DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
      DSQ = "surfactan_II"
 
 lista2.ReportFileName = "Listafarmacopea.rpt"
 
 lista2.SQLQuery = "SELECT * " _
                + "From " _
                + DSQ + ".dbo.especifunifica especifunifica " _
                + "Where " _
                + " EspecifUnifica.farmacopea = 1 "
                
    
   
    
    
    
    
    lista2.Connect = Connect()
    lista2.Destination = 0
    lista2.Action = 1
End Sub

Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub EntraIngles_Click()
    
    If WGrabaIII <> "S" Then
    
        Call Ingresa_ClaveIII

               Else
    
        WGrabaIII = ""
        PantaIngles.Height = 7335
        PantaIngles.Left = 4850
        PantaIngles.Top = 120
        PantaIngles.Width = 5100
        
        PantaIngles.Visible = True
        
        Valor1Ing.SetFocus
        
    End If
    
End Sub

Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Private Sub Imprime_Datos()

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
    
    ZSql = ""
    ZSql = ZSql & "Select EspecifUnifica.Producto, EspecifUnifica.Ensayo1, EspecifUnifica.Ensayo2, EspecifUnifica.Ensayo3, EspecifUnifica.Ensayo4, EspecifUnifica.Ensayo5, EspecifUnifica.Ensayo6, EspecifUnifica.Ensayo7, EspecifUnifica.Ensayo8, EspecifUnifica.Ensayo9, EspecifUnifica.Ensayo10, " _
                + "EspecifUnifica.valor1, EspecifUnifica.valor2, EspecifUnifica.valor3, EspecifUnifica.valor4, EspecifUnifica.valor5, EspecifUnifica.valor6, EspecifUnifica.valor7, EspecifUnifica.valor8, EspecifUnifica.valor9, EspecifUnifica.valor10, " _
                + "EspecifUnifica.valor11, EspecifUnifica.valor22, EspecifUnifica.valor33, EspecifUnifica.valor44, EspecifUnifica.valor55, EspecifUnifica.valor66, EspecifUnifica.valor77, EspecifUnifica.valor88, EspecifUnifica.valor99, EspecifUnifica.valor1010, " _
                + "EspecifUnifica.desde1, EspecifUnifica.desde2, EspecifUnifica.desde3, EspecifUnifica.desde4, EspecifUnifica.desde5, EspecifUnifica.desde6, EspecifUnifica.desde7, EspecifUnifica.desde8, EspecifUnifica.desde9, EspecifUnifica.desde10, " _
                + "EspecifUnifica.hasta1, EspecifUnifica.hasta2, EspecifUnifica.hasta3, EspecifUnifica.hasta4, EspecifUnifica.hasta5, EspecifUnifica.hasta6, EspecifUnifica.hasta7, EspecifUnifica.hasta8, EspecifUnifica.hasta9, EspecifUnifica.hasta10"
    ZSql = ZSql & " FROM EspecifUnifica"
    ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        Producto.Text = rstEspecifUnifica!Producto
        
        Ensayo1.Text = rstEspecifUnifica!Ensayo1
        Ensayo2.Text = rstEspecifUnifica!Ensayo2
        Ensayo3.Text = rstEspecifUnifica!Ensayo3
        Ensayo4.Text = rstEspecifUnifica!Ensayo4
        Ensayo5.Text = rstEspecifUnifica!Ensayo5
        Ensayo6.Text = rstEspecifUnifica!Ensayo6
        Ensayo7.Text = rstEspecifUnifica!Ensayo7
        Ensayo8.Text = rstEspecifUnifica!Ensayo8
        Ensayo9.Text = rstEspecifUnifica!Ensayo9
        Ensayo10.Text = rstEspecifUnifica!Ensayo10
        
        Valor1.Text = rstEspecifUnifica!Valor1
        valor2.Text = rstEspecifUnifica!valor2
        Valor3.Text = rstEspecifUnifica!Valor3
        valor4.Text = rstEspecifUnifica!valor4
        valor5.Text = rstEspecifUnifica!valor5
        valor6.Text = rstEspecifUnifica!valor6
        valor7.Text = rstEspecifUnifica!valor7
        valor8.Text = rstEspecifUnifica!valor8
        valor9.Text = rstEspecifUnifica!valor9
        valor10.Text = rstEspecifUnifica!valor10
        Valor11.Text = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
        Valor22.Text = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
        Valor33.Text = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
        Valor44.Text = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
        Valor55.Text = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
        Valor66.Text = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
        Valor77.Text = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
        Valor88.Text = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
        Valor99.Text = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
        Valor1010.Text = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
        
        
        
        Desde1.Text = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
        Desde2.Text = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
        Desde3.Text = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
        Desde4.Text = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
        Desde5.Text = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
        Desde6.Text = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
        Desde7.Text = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
        Desde8.Text = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
        Desde9.Text = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
        Desde10.Text = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
        
        Hasta1.Text = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
        Hasta2.Text = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
        Hasta3.Text = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
        Hasta4.Text = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
        Hasta5.Text = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
        Hasta6.Text = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
        Hasta7.Text = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
        Hasta8.Text = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
        Hasta9.Text = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
        Hasta10.Text = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
        
        Desde1.Text = Trim(Desde1.Text)
        Desde2.Text = Trim(Desde2.Text)
        Desde3.Text = Trim(Desde3.Text)
        Desde4.Text = Trim(Desde4.Text)
        Desde5.Text = Trim(Desde5.Text)
        Desde6.Text = Trim(Desde6.Text)
        Desde7.Text = Trim(Desde7.Text)
        Desde8.Text = Trim(Desde8.Text)
        Desde9.Text = Trim(Desde9.Text)
        Desde10.Text = Trim(Desde10.Text)
        
        Hasta1.Text = Trim(Hasta1.Text)
        Hasta2.Text = Trim(Hasta2.Text)
        Hasta3.Text = Trim(Hasta3.Text)
        Hasta4.Text = Trim(Hasta4.Text)
        Hasta5.Text = Trim(Hasta5.Text)
        Hasta6.Text = Trim(Hasta6.Text)
        Hasta7.Text = Trim(Hasta7.Text)
        Hasta8.Text = Trim(Hasta8.Text)
        Hasta9.Text = Trim(Hasta9.Text)
        Hasta10.Text = Trim(Hasta10.Text)
        
        rstEspecifUnifica.Close
        
    End If
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql & "Select EspecifUnifica.Valor1Ing, EspecifUnifica.Valor2Ing, EspecifUnifica.Valor3Ing, EspecifUnifica.Valor4Ing, EspecifUnifica.Valor5Ing, EspecifUnifica.Valor6Ing, EspecifUnifica.Valor7Ing, EspecifUnifica.Valor8Ing, EspecifUnifica.Valor9Ing, EspecifUnifica.Valor10Ing, " _
                  + "EspecifUnifica.Valor11Ing, EspecifUnifica.Valor22Ing, EspecifUnifica.Valor33Ing, EspecifUnifica.Valor44Ing, EspecifUnifica.Valor55Ing, EspecifUnifica.Valor66Ing, EspecifUnifica.Valor77Ing, EspecifUnifica.Valor88Ing, EspecifUnifica.Valor99Ing, EspecifUnifica.Valor1010Ing, " _
                  + "EspecifUnifica.Version, EspecifUnifica.fecha, EspecifUnifica.Estado, EspecifUnifica.Operador, EspecifUnifica.farmacopea,EspecifUnifica.ControlCambio"
    ZSql = ZSql & " FROM EspecifUnifica"
    ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        
        Valor1Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor1Ing), "", rstEspecifUnifica!Valor1Ing)
        Valor2Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor2Ing), "", rstEspecifUnifica!Valor2Ing)
        Valor3Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor3Ing), "", rstEspecifUnifica!Valor3Ing)
        Valor4Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor4Ing), "", rstEspecifUnifica!Valor4Ing)
        Valor5Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor5Ing), "", rstEspecifUnifica!Valor5Ing)
        Valor6Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor6Ing), "", rstEspecifUnifica!Valor6Ing)
        Valor7Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor7Ing), "", rstEspecifUnifica!Valor7Ing)
        Valor8Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor8Ing), "", rstEspecifUnifica!Valor8Ing)
        Valor9Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor9Ing), "", rstEspecifUnifica!Valor9Ing)
        Valor10Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor10Ing), "", rstEspecifUnifica!Valor10Ing)
        Valor11Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor11Ing), "", rstEspecifUnifica!Valor11Ing)
        Valor22Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor22Ing), "", rstEspecifUnifica!Valor22Ing)
        Valor33Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor33Ing), "", rstEspecifUnifica!Valor33Ing)
        Valor44Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor44Ing), "", rstEspecifUnifica!Valor44Ing)
        Valor55Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor55Ing), "", rstEspecifUnifica!Valor55Ing)
        Valor66Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor66Ing), "", rstEspecifUnifica!Valor66Ing)
        Valor77Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor77Ing), "", rstEspecifUnifica!Valor77Ing)
        Valor88Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor88Ing), "", rstEspecifUnifica!Valor88Ing)
        Valor99Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor99Ing), "", rstEspecifUnifica!Valor99Ing)
        Valor1010Ing.Text = IIf(IsNull(rstEspecifUnifica!Valor1010Ing), "", rstEspecifUnifica!Valor1010Ing)
        
        
        Valor1Ing.Text = Trim(Valor1Ing.Text)
        Valor2Ing.Text = Trim(Valor2Ing.Text)
        Valor3Ing.Text = Trim(Valor3Ing.Text)
        Valor4Ing.Text = Trim(Valor4Ing.Text)
        Valor5Ing.Text = Trim(Valor5Ing.Text)
        Valor6Ing.Text = Trim(Valor6Ing.Text)
        Valor7Ing.Text = Trim(Valor7Ing.Text)
        Valor8Ing.Text = Trim(Valor8Ing.Text)
        Valor9Ing.Text = Trim(Valor9Ing.Text)
        Valor10Ing.Text = Trim(Valor10Ing.Text)
        Valor11Ing.Text = Trim(Valor11Ing.Text)
        Valor22Ing.Text = Trim(Valor22Ing.Text)
        Valor33Ing.Text = Trim(Valor33Ing.Text)
        Valor44Ing.Text = Trim(Valor44Ing.Text)
        Valor55Ing.Text = Trim(Valor55Ing.Text)
        Valor66Ing.Text = Trim(Valor66Ing.Text)
        Valor77Ing.Text = Trim(Valor77Ing.Text)
        Valor88Ing.Text = Trim(Valor88Ing.Text)
        Valor99Ing.Text = Trim(Valor99Ing.Text)
        Valor1010Ing.Text = Trim(Valor1010Ing.Text)
        
        
        
        Version.Text = rstEspecifUnifica!Version
        Fecha.Text = rstEspecifUnifica!Fecha
        Estado.Text = IIf(IsNull(rstEspecifUnifica!Estado), "", rstEspecifUnifica!Estado)
        ZOperador = IIf(IsNull(rstEspecifUnifica!Operador), "O", rstEspecifUnifica!Operador)
        
        ControlCambio.Text = IIf(IsNull(rstEspecifUnifica!ControlCambio), "", rstEspecifUnifica!ControlCambio)
        farmacopea = IIf(IsNull(rstEspecifUnifica!farmacopea), "0", rstEspecifUnifica!farmacopea)
        
        
        rstEspecifUnifica.Close
        
    End If
    Rem by nan
    If farmacopea = 1 Then
    Check1.Value = 1
    
    Else
    Check1.Value = 0
    End If
    Rem by nan
    
    
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
    
    DesOperador.Caption = ""
    If Val(ZOperador) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Operador = " + "'" + ZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            DesOperador.Caption = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
            rstOperador.Close
        End If
    End If
    
    Call Conecta_Empresa
        
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo WError
    
    If Val(WEmpresa) = 5 Then
        m$ = "No se puede actualizar los datos por ser el proucto terminado de farma"
        A% = MsgBox(m$, 0, "Exportacion de Homologacion de Muestras")
        Exit Sub
    End If
    
    If Trim(ControlCambio.Text) = "" Then
        m$ = "Se debe informar el campo Control de Cambio"
        A% = MsgBox(m$, 0, "Especificaciones de Producto Terminado")
        Exit Sub
    End If
    
    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else

        If Producto.Text <> "" Then
        
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZZCliente = "S00102"
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZZCliente = "P99999"
            End Select
        
            WProducto = Producto.Text
            
            WEnsayo1 = Ensayo1.Text
            WEnsayo2 = Ensayo2.Text
            WEnsayo3 = Ensayo3.Text
            WEnsayo4 = Ensayo4.Text
            WEnsayo5 = Ensayo5.Text
            WEnsayo6 = Ensayo6.Text
            WEnsayo7 = Ensayo7.Text
            WEnsayo8 = Ensayo8.Text
            WEnsayo9 = Ensayo9.Text
            WEnsayo10 = Ensayo10.Text
            
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
            WValor11 = Valor11.Text
            WValor22 = Valor22.Text
            WValor33 = Valor33.Text
            WValor44 = Valor44.Text
            WValor55 = Valor55.Text
            WValor66 = Valor66.Text
            WValor77 = Valor77.Text
            WValor88 = Valor88.Text
            WValor99 = Valor99.Text
            WValor1010 = Valor1010.Text
            
            WDesde1 = Desde1.Text
            WDesde2 = Desde2.Text
            WDesde3 = Desde3.Text
            WDesde4 = Desde4.Text
            WDesde5 = Desde5.Text
            WDesde6 = Desde6.Text
            WDesde7 = Desde7.Text
            WDesde8 = Desde8.Text
            WDesde9 = Desde9.Text
            WDesde10 = Desde10.Text
            
            WHasta1 = Hasta1.Text
            WHasta2 = Hasta2.Text
            WHasta3 = Hasta3.Text
            WHasta4 = Hasta4.Text
            WHasta5 = Hasta5.Text
            WHasta6 = Hasta6.Text
            WHasta7 = Hasta7.Text
            WHasta8 = Hasta8.Text
            WHasta9 = Hasta9.Text
            WHasta10 = Hasta10.Text
            
            WDate = Date$
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
            
                ZProducto = rstEspecifUnifica!Producto
                
                ZEnsayo1 = Str$(rstEspecifUnifica!Ensayo1)
                ZEnsayo2 = Str$(rstEspecifUnifica!Ensayo2)
                ZEnsayo3 = Str$(rstEspecifUnifica!Ensayo3)
                ZEnsayo4 = Str$(rstEspecifUnifica!Ensayo4)
                ZEnsayo5 = Str$(rstEspecifUnifica!Ensayo5)
                ZEnsayo6 = Str$(rstEspecifUnifica!Ensayo6)
                ZEnsayo7 = Str$(rstEspecifUnifica!Ensayo7)
                ZEnsayo8 = Str$(rstEspecifUnifica!Ensayo8)
                ZEnsayo9 = Str$(rstEspecifUnifica!Ensayo9)
                ZEnsayo10 = Str$(rstEspecifUnifica!Ensayo10)
                
                ZValor1 = rstEspecifUnifica!Valor1
                ZValor2 = rstEspecifUnifica!valor2
                ZValor3 = rstEspecifUnifica!Valor3
                ZValor4 = rstEspecifUnifica!valor4
                ZValor5 = rstEspecifUnifica!valor5
                ZValor6 = rstEspecifUnifica!valor6
                ZValor7 = rstEspecifUnifica!valor7
                ZValor8 = rstEspecifUnifica!valor8
                ZValor9 = rstEspecifUnifica!valor9
                ZValor10 = rstEspecifUnifica!valor10
                
                ZValor11 = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                ZValor22 = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                ZValor33 = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                ZValor44 = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                ZValor55 = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                ZValor66 = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                ZValor77 = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                ZValor88 = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                ZValor99 = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                ZValor1010 = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                
                ZDesde1 = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                ZDesde2 = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                ZDesde3 = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                ZDesde4 = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                ZDesde5 = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                ZDesde6 = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                ZDesde7 = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                ZDesde8 = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                ZDesde9 = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                ZDesde10 = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                
                ZHasta1 = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                ZHasta2 = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
                ZHasta3 = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                ZHasta4 = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                ZHasta5 = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                ZHasta6 = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                ZHasta7 = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                ZHasta8 = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                ZHasta9 = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                ZHasta10 = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                
                ZVersion = Version.Text
                ZFechaInicio = Fecha.Text
                ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
                ZControlCambio = ControlCambio.Text
            
                rstEspecifUnifica.Close
            
                Call Ceros(ZVersion, 4)
                ZClave = ZVersion + ZProducto
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnificaVersion ("
                ZSql = ZSql & "Clave, "
                ZSql = ZSql & "Version, "
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "FechaInicio , "
                ZSql = ZSql & "FechaFinal , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + ZClave + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZProducto + "',"
                ZSql = ZSql & "'" + ZEnsayo1 + "'," + "'" + ZValor1 + "',"
                ZSql = ZSql & "'" + ZEnsayo2 + "'," + "'" + ZValor2 + "',"
                ZSql = ZSql & "'" + ZEnsayo3 + "'," + "'" + ZValor3 + "',"
                ZSql = ZSql & "'" + ZEnsayo4 + "'," + "'" + ZValor4 + "',"
                ZSql = ZSql & "'" + ZEnsayo5 + "'," + "'" + ZValor5 + "',"
                ZSql = ZSql & "'" + ZEnsayo6 + "'," + "'" + ZValor6 + "',"
                ZSql = ZSql & "'" + ZEnsayo7 + "'," + "'" + ZValor7 + "',"
                ZSql = ZSql & "'" + ZEnsayo8 + "'," + "'" + ZValor8 + "',"
                ZSql = ZSql & "'" + ZEnsayo9 + "'," + "'" + ZValor9 + "',"
                ZSql = ZSql & "'" + ZEnsayo10 + "'," + "'" + ZValor10 + "',"
                ZSql = ZSql & "'" + ZValor11 + "',"
                ZSql = ZSql & "'" + ZValor22 + "',"
                ZSql = ZSql & "'" + ZValor33 + "',"
                ZSql = ZSql & "'" + ZValor44 + "',"
                ZSql = ZSql & "'" + ZValor55 + "',"
                ZSql = ZSql & "'" + ZValor66 + "',"
                ZSql = ZSql & "'" + ZValor77 + "',"
                ZSql = ZSql & "'" + ZValor88 + "',"
                ZSql = ZSql & "'" + ZValor99 + "',"
                ZSql = ZSql & "'" + ZValor1010 + "',"
                ZSql = ZSql & "'" + ZFechaInicio + "',"
                ZSql = ZSql & "'" + ZFechaFinal + "',"
                ZSql = ZSql & "'" + ZControlCambio + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
          
                spEspecifUnificaVersion = ZSql
                Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecifUnificaVersion SET "
                ZSql = ZSql + "Desde1 = " + "'" + ZDesde1 + "',"
                ZSql = ZSql + "Hasta1 = " + "'" + ZHasta1 + "',"
                ZSql = ZSql + "Desde2 = " + "'" + ZDesde2 + "',"
                ZSql = ZSql + "Hasta2 = " + "'" + ZHasta2 + "',"
                ZSql = ZSql + "Desde3 = " + "'" + ZDesde3 + "',"
                ZSql = ZSql + "Hasta3 = " + "'" + ZHasta3 + "',"
                ZSql = ZSql + "Desde4 = " + "'" + ZDesde4 + "',"
                ZSql = ZSql + "Hasta4 = " + "'" + ZHasta4 + "',"
                ZSql = ZSql + "Desde5 = " + "'" + ZDesde5 + "',"
                ZSql = ZSql + "Hasta5 = " + "'" + ZHasta5 + "',"
                ZSql = ZSql + "Desde6 = " + "'" + ZDesde6 + "',"
                ZSql = ZSql + "Hasta6 = " + "'" + ZHasta6 + "',"
                ZSql = ZSql + "Desde7 = " + "'" + ZDesde7 + "',"
                ZSql = ZSql + "Hasta7 = " + "'" + ZHasta7 + "',"
                ZSql = ZSql + "Desde8 = " + "'" + ZDesde8 + "',"
                ZSql = ZSql + "Hasta8 = " + "'" + ZHasta8 + "',"
                ZSql = ZSql + "Desde9 = " + "'" + ZDesde9 + "',"
                ZSql = ZSql + "Hasta9 = " + "'" + ZHasta9 + "',"
                ZSql = ZSql + "Desde10 = " + "'" + ZDesde10 + "',"
                ZSql = ZSql + "Hasta10 = " + "'" + ZHasta10 + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                     
                spEspecifUnificaVersion = ZSql
                Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
                ZVersion = Str$(Val(Version.Text) + 1)
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
                ZEstado = "S"
                ZControlCambio = ControlCambio.Text
                
                XXVersion = ZVersion
                XXFechaVersion = ZFecha
                XXEstado = ZEstado
                XXObservaciones = ""
            
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Producto = " + "'" + WProducto + "',"
                ZSql = ZSql & "Ensayo1 = " + "'" + WEnsayo1 + "',"
                ZSql = ZSql & "Valor1 = " + "'" + WValor1 + "',"
                ZSql = ZSql & "Ensayo2 = " + "'" + WEnsayo2 + "',"
                ZSql = ZSql & "Valor2 = " + "'" + WValor2 + "',"
                ZSql = ZSql & "Ensayo3 = " + "'" + WEnsayo3 + "',"
                ZSql = ZSql & "Valor3 = " + "'" + WValor3 + "',"
                ZSql = ZSql & "Ensayo4 = " + "'" + WEnsayo4 + "',"
                ZSql = ZSql & "Valor4 = " + "'" + WValor4 + "',"
                ZSql = ZSql & "Ensayo5 = " + "'" + WEnsayo5 + "',"
                ZSql = ZSql & "Valor5 = " + "'" + WValor5 + "',"
                ZSql = ZSql & "Ensayo6 = " + "'" + WEnsayo6 + "',"
                ZSql = ZSql & "Valor6 = " + "'" + WValor6 + "',"
                ZSql = ZSql & "Ensayo7 = " + "'" + WEnsayo7 + "',"
                ZSql = ZSql & "Valor7 = " + "'" + WValor7 + "',"
                ZSql = ZSql & "Ensayo8 = " + "'" + WEnsayo8 + "',"
                ZSql = ZSql & "Valor8 = " + "'" + WValor8 + "',"
                ZSql = ZSql & "Ensayo9 = " + "'" + WEnsayo9 + "',"
                ZSql = ZSql & "Valor9 = " + "'" + WValor9 + "',"
                ZSql = ZSql & "Ensayo10 = " + "'" + WEnsayo10 + "',"
                ZSql = ZSql & "Valor10 = " + "'" + WValor10 + "',"
                ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                ZSql = ZSql & "Valor11 = " + "'" + WValor11 + "',"
                ZSql = ZSql & "Valor22 = " + "'" + WValor22 + "',"
                ZSql = ZSql & "Valor33 = " + "'" + WValor33 + "',"
                ZSql = ZSql & "Valor44 = " + "'" + WValor44 + "',"
                ZSql = ZSql & "Valor55 = " + "'" + WValor55 + "',"
                ZSql = ZSql & "Valor66 = " + "'" + WValor66 + "',"
                ZSql = ZSql & "Valor77 = " + "'" + WValor77 + "',"
                ZSql = ZSql & "Valor88 = " + "'" + WValor88 + "',"
                ZSql = ZSql & "Valor99 = " + "'" + WValor99 + "',"
                ZSql = ZSql & "Valor1010 = " + "'" + WValor1010 + "',"
                ZSql = ZSql & "Desde1 = " + "'" + WDesde1 + "',"
                ZSql = ZSql & "Desde2 = " + "'" + WDesde2 + "',"
                ZSql = ZSql & "Desde3 = " + "'" + WDesde3 + "',"
                ZSql = ZSql & "Desde4 = " + "'" + WDesde4 + "',"
                ZSql = ZSql & "Desde5 = " + "'" + WDesde5 + "',"
                ZSql = ZSql & "Desde6 = " + "'" + WDesde6 + "',"
                ZSql = ZSql & "Desde7 = " + "'" + WDesde7 + "',"
                ZSql = ZSql & "Desde8 = " + "'" + WDesde8 + "',"
                ZSql = ZSql & "Desde9 = " + "'" + WDesde9 + "',"
                ZSql = ZSql & "Desde10 = " + "'" + WDesde10 + "',"
                ZSql = ZSql & "Hasta1 = " + "'" + WHasta1 + "',"
                ZSql = ZSql & "Hasta2 = " + "'" + WHasta2 + "',"
                ZSql = ZSql & "Hasta3 = " + "'" + WHasta3 + "',"
                ZSql = ZSql & "Hasta4 = " + "'" + WHasta4 + "',"
                ZSql = ZSql & "Hasta5 = " + "'" + WHasta5 + "',"
                ZSql = ZSql & "Hasta6 = " + "'" + WHasta6 + "',"
                ZSql = ZSql & "Hasta7 = " + "'" + WHasta7 + "',"
                ZSql = ZSql & "Hasta8 = " + "'" + WHasta8 + "',"
                ZSql = ZSql & "Hasta9 = " + "'" + WHasta9 + "',"
                ZSql = ZSql & "Hasta10 = " + "'" + WHasta10 + "',"
                ZSql = ZSql & "Version = " + "'" + ZVersion + "',"
                ZSql = ZSql & "Fecha = " + "'" + ZFecha + "',"
                ZSql = ZSql & "Estado = " + "'" + ZEstado + "',"
                ZSql = ZSql & "ControlCambio = " + "'" + ZControlCambio + "',"
                ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
                    
                ZVersion = "1"
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
                ZEstado = "S"
                ZControlCambio = ControlCambio.Text
                
                XXVersion = ZVersion
                XXFechaVersion = ZFecha
                XXEstado = ZEstado
                XXObservaciones = ""
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnifica ("
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "WDate, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "Desde1 , "
                ZSql = ZSql & "Desde2 , "
                ZSql = ZSql & "Desde3 , "
                ZSql = ZSql & "Desde4 , "
                ZSql = ZSql & "Desde5 , "
                ZSql = ZSql & "Desde6 , "
                ZSql = ZSql & "Desde7 , "
                ZSql = ZSql & "Desde8 , "
                ZSql = ZSql & "Desde9 , "
                ZSql = ZSql & "Desde10 , "
                ZSql = ZSql & "Hasta1 , "
                ZSql = ZSql & "Hasta2 , "
                ZSql = ZSql & "Hasta3 , "
                ZSql = ZSql & "Hasta4 , "
                ZSql = ZSql & "Hasta5 , "
                ZSql = ZSql & "Hasta6 , "
                ZSql = ZSql & "Hasta7 , "
                ZSql = ZSql & "Hasta8 , "
                ZSql = ZSql & "Hasta9 , "
                ZSql = ZSql & "Hasta10 , "
                ZSql = ZSql & "Version , "
                ZSql = ZSql & "Fecha , "
                ZSql = ZSql & "Estado , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WProducto + "',"
                ZSql = ZSql & "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
                ZSql = ZSql & "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
                ZSql = ZSql & "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
                ZSql = ZSql & "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
                ZSql = ZSql & "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
                ZSql = ZSql & "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
                ZSql = ZSql & "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
                ZSql = ZSql & "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
                ZSql = ZSql & "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
                ZSql = ZSql & "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WValor11 + "',"
                ZSql = ZSql & "'" + WValor22 + "',"
                ZSql = ZSql & "'" + WValor33 + "',"
                ZSql = ZSql & "'" + WValor44 + "',"
                ZSql = ZSql & "'" + WValor55 + "',"
                ZSql = ZSql & "'" + WValor66 + "',"
                ZSql = ZSql & "'" + WValor77 + "',"
                ZSql = ZSql & "'" + WValor88 + "',"
                ZSql = ZSql & "'" + WValor99 + "',"
                ZSql = ZSql & "'" + WValor1010 + "',"
                ZSql = ZSql & "'" + WDesde1 + "',"
                ZSql = ZSql & "'" + WDesde2 + "',"
                ZSql = ZSql & "'" + WDesde3 + "',"
                ZSql = ZSql & "'" + WDesde4 + "',"
                ZSql = ZSql & "'" + WDesde5 + "',"
                ZSql = ZSql & "'" + WDesde6 + "',"
                ZSql = ZSql & "'" + WDesde7 + "',"
                ZSql = ZSql & "'" + WDesde8 + "',"
                ZSql = ZSql & "'" + WDesde9 + "',"
                ZSql = ZSql & "'" + WDesde10 + "',"
                ZSql = ZSql & "'" + WHasta1 + "',"
                ZSql = ZSql & "'" + WHasta2 + "',"
                ZSql = ZSql & "'" + WHasta3 + "',"
                ZSql = ZSql & "'" + WHasta4 + "',"
                ZSql = ZSql & "'" + WHasta5 + "',"
                ZSql = ZSql & "'" + WHasta6 + "',"
                ZSql = ZSql & "'" + WHasta7 + "',"
                ZSql = ZSql & "'" + WHasta8 + "',"
                ZSql = ZSql & "'" + WHasta9 + "',"
                ZSql = ZSql & "'" + WHasta10 + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZFecha + "',"
                ZSql = ZSql & "'" + ZEstado + "',"
                ZSql = ZSql & "'" + ZControlCambio + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
           
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE EspecifUnifica SET "
            ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
            ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                            
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM AltaCertificado"
            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + WProducto + "'"
            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZZZCliente + "'"
            spAltaCertificado = ZSql
            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
            If rstAltaCertificado.RecordCount > 0 Then
                rstAltaCertificado.Close
                    Else
                m$ = "No se ha definido el certificado de analiss de este producto para el cliente generico"
                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
            End If
            
            
            
            WProductoNK = "NK" + Mid$(Producto.Text, 3, 10)
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + WProductoNK + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                rstEspecifUnifica.Close
                
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Ensayo1 = " + "'" + WEnsayo1 + "',"
                ZSql = ZSql & "Valor1 = " + "'" + WValor1 + "',"
                ZSql = ZSql & "Ensayo2 = " + "'" + WEnsayo2 + "',"
                ZSql = ZSql & "Valor2 = " + "'" + WValor2 + "',"
                ZSql = ZSql & "Ensayo3 = " + "'" + WEnsayo3 + "',"
                ZSql = ZSql & "Valor3 = " + "'" + WValor3 + "',"
                ZSql = ZSql & "Ensayo4 = " + "'" + WEnsayo4 + "',"
                ZSql = ZSql & "Valor4 = " + "'" + WValor4 + "',"
                ZSql = ZSql & "Ensayo5 = " + "'" + WEnsayo5 + "',"
                ZSql = ZSql & "Valor5 = " + "'" + WValor5 + "',"
                ZSql = ZSql & "Ensayo6 = " + "'" + WEnsayo6 + "',"
                ZSql = ZSql & "Valor6 = " + "'" + WValor6 + "',"
                ZSql = ZSql & "Ensayo7 = " + "'" + WEnsayo7 + "',"
                ZSql = ZSql & "Valor7 = " + "'" + WValor7 + "',"
                ZSql = ZSql & "Ensayo8 = " + "'" + WEnsayo8 + "',"
                ZSql = ZSql & "Valor8 = " + "'" + WValor8 + "',"
                ZSql = ZSql & "Ensayo9 = " + "'" + WEnsayo9 + "',"
                ZSql = ZSql & "Valor9 = " + "'" + WValor9 + "',"
                ZSql = ZSql & "Ensayo10 = " + "'" + WEnsayo10 + "',"
                ZSql = ZSql & "Valor10 = " + "'" + WValor10 + "',"
                ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                ZSql = ZSql & "Valor11 = " + "'" + WValor11 + "',"
                ZSql = ZSql & "Valor22 = " + "'" + WValor22 + "',"
                ZSql = ZSql & "Valor33 = " + "'" + WValor33 + "',"
                ZSql = ZSql & "Valor44 = " + "'" + WValor44 + "',"
                ZSql = ZSql & "Valor55 = " + "'" + WValor55 + "',"
                ZSql = ZSql & "Valor66 = " + "'" + WValor66 + "',"
                ZSql = ZSql & "Valor77 = " + "'" + WValor77 + "',"
                ZSql = ZSql & "Valor88 = " + "'" + WValor88 + "',"
                ZSql = ZSql & "Valor99 = " + "'" + WValor99 + "',"
                ZSql = ZSql & "Valor1010 = " + "'" + WValor1010 + "',"
                ZSql = ZSql & "Desde1 = " + "'" + WDesde1 + "',"
                ZSql = ZSql & "Desde2 = " + "'" + WDesde2 + "',"
                ZSql = ZSql & "Desde3 = " + "'" + WDesde3 + "',"
                ZSql = ZSql & "Desde4 = " + "'" + WDesde4 + "',"
                ZSql = ZSql & "Desde5 = " + "'" + WDesde5 + "',"
                ZSql = ZSql & "Desde6 = " + "'" + WDesde6 + "',"
                ZSql = ZSql & "Desde7 = " + "'" + WDesde7 + "',"
                ZSql = ZSql & "Desde8 = " + "'" + WDesde8 + "',"
                ZSql = ZSql & "Desde9 = " + "'" + WDesde9 + "',"
                ZSql = ZSql & "Desde10 = " + "'" + WDesde10 + "',"
                ZSql = ZSql & "Hasta1 = " + "'" + WHasta1 + "',"
                ZSql = ZSql & "Hasta2 = " + "'" + WHasta2 + "',"
                ZSql = ZSql & "Hasta3 = " + "'" + WHasta3 + "',"
                ZSql = ZSql & "Hasta4 = " + "'" + WHasta4 + "',"
                ZSql = ZSql & "Hasta5 = " + "'" + WHasta5 + "',"
                ZSql = ZSql & "Hasta6 = " + "'" + WHasta6 + "',"
                ZSql = ZSql & "Hasta7 = " + "'" + WHasta7 + "',"
                ZSql = ZSql & "Hasta8 = " + "'" + WHasta8 + "',"
                ZSql = ZSql & "Hasta9 = " + "'" + WHasta9 + "',"
                ZSql = ZSql & "Hasta10 = " + "'" + WHasta10 + "',"
                ZSql = ZSql & "Version = " + "'" + ZVersion + "',"
                ZSql = ZSql & "Fecha = " + "'" + ZFecha + "',"
                ZSql = ZSql & "Estado = " + "'" + ZEstado + "',"
                ZSql = ZSql & "ControlCambio = " + "'" + ZControlCambio + "',"
                ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProductoNK + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
                    
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnifica ("
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "WDate, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "Desde1 , "
                ZSql = ZSql & "Desde2 , "
                ZSql = ZSql & "Desde3 , "
                ZSql = ZSql & "Desde4 , "
                ZSql = ZSql & "Desde5 , "
                ZSql = ZSql & "Desde6 , "
                ZSql = ZSql & "Desde7 , "
                ZSql = ZSql & "Desde8 , "
                ZSql = ZSql & "Desde9 , "
                ZSql = ZSql & "Desde10 , "
                ZSql = ZSql & "Hasta1 , "
                ZSql = ZSql & "Hasta2 , "
                ZSql = ZSql & "Hasta3 , "
                ZSql = ZSql & "Hasta4 , "
                ZSql = ZSql & "Hasta5 , "
                ZSql = ZSql & "Hasta6 , "
                ZSql = ZSql & "Hasta7 , "
                ZSql = ZSql & "Hasta8 , "
                ZSql = ZSql & "Hasta9 , "
                ZSql = ZSql & "Hasta10 , "
                ZSql = ZSql & "Version , "
                ZSql = ZSql & "Fecha , "
                ZSql = ZSql & "Estado , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WProductoNK + "',"
                ZSql = ZSql & "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
                ZSql = ZSql & "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
                ZSql = ZSql & "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
                ZSql = ZSql & "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
                ZSql = ZSql & "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
                ZSql = ZSql & "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
                ZSql = ZSql & "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
                ZSql = ZSql & "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
                ZSql = ZSql & "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
                ZSql = ZSql & "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WValor11 + "',"
                ZSql = ZSql & "'" + WValor22 + "',"
                ZSql = ZSql & "'" + WValor33 + "',"
                ZSql = ZSql & "'" + WValor44 + "',"
                ZSql = ZSql & "'" + WValor55 + "',"
                ZSql = ZSql & "'" + WValor66 + "',"
                ZSql = ZSql & "'" + WValor77 + "',"
                ZSql = ZSql & "'" + WValor88 + "',"
                ZSql = ZSql & "'" + WValor99 + "',"
                ZSql = ZSql & "'" + WValor1010 + "',"
                ZSql = ZSql & "'" + WDesde1 + "',"
                ZSql = ZSql & "'" + WDesde2 + "',"
                ZSql = ZSql & "'" + WDesde3 + "',"
                ZSql = ZSql & "'" + WDesde4 + "',"
                ZSql = ZSql & "'" + WDesde5 + "',"
                ZSql = ZSql & "'" + WDesde6 + "',"
                ZSql = ZSql & "'" + WDesde7 + "',"
                ZSql = ZSql & "'" + WDesde8 + "',"
                ZSql = ZSql & "'" + WDesde9 + "',"
                ZSql = ZSql & "'" + WDesde10 + "',"
                ZSql = ZSql & "'" + WHasta1 + "',"
                ZSql = ZSql & "'" + WHasta2 + "',"
                ZSql = ZSql & "'" + WHasta3 + "',"
                ZSql = ZSql & "'" + WHasta4 + "',"
                ZSql = ZSql & "'" + WHasta5 + "',"
                ZSql = ZSql & "'" + WHasta6 + "',"
                ZSql = ZSql & "'" + WHasta7 + "',"
                ZSql = ZSql & "'" + WHasta8 + "',"
                ZSql = ZSql & "'" + WHasta9 + "',"
                ZSql = ZSql & "'" + WHasta10 + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZFecha + "',"
                ZSql = ZSql & "'" + ZEstado + "',"
                ZSql = ZSql & "'" + ZControlCambio + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
           
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecifUnifica SET "
                ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
                ZSql = ZSql + " Where Producto = " + "'" + WProductoNK + "'"
                            
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            
            
            
            
            WProductoRE = "RE" + Mid$(Producto.Text, 3, 10)
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + WProductoRE + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                
                rstEspecifUnifica.Close
                
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Ensayo1 = " + "'" + WEnsayo1 + "',"
                ZSql = ZSql & "Valor1 = " + "'" + WValor1 + "',"
                ZSql = ZSql & "Ensayo2 = " + "'" + WEnsayo2 + "',"
                ZSql = ZSql & "Valor2 = " + "'" + WValor2 + "',"
                ZSql = ZSql & "Ensayo3 = " + "'" + WEnsayo3 + "',"
                ZSql = ZSql & "Valor3 = " + "'" + WValor3 + "',"
                ZSql = ZSql & "Ensayo4 = " + "'" + WEnsayo4 + "',"
                ZSql = ZSql & "Valor4 = " + "'" + WValor4 + "',"
                ZSql = ZSql & "Ensayo5 = " + "'" + WEnsayo5 + "',"
                ZSql = ZSql & "Valor5 = " + "'" + WValor5 + "',"
                ZSql = ZSql & "Ensayo6 = " + "'" + WEnsayo6 + "',"
                ZSql = ZSql & "Valor6 = " + "'" + WValor6 + "',"
                ZSql = ZSql & "Ensayo7 = " + "'" + WEnsayo7 + "',"
                ZSql = ZSql & "Valor7 = " + "'" + WValor7 + "',"
                ZSql = ZSql & "Ensayo8 = " + "'" + WEnsayo8 + "',"
                ZSql = ZSql & "Valor8 = " + "'" + WValor8 + "',"
                ZSql = ZSql & "Ensayo9 = " + "'" + WEnsayo9 + "',"
                ZSql = ZSql & "Valor9 = " + "'" + WValor9 + "',"
                ZSql = ZSql & "Ensayo10 = " + "'" + WEnsayo10 + "',"
                ZSql = ZSql & "Valor10 = " + "'" + WValor10 + "',"
                ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                ZSql = ZSql & "Valor11 = " + "'" + WValor11 + "',"
                ZSql = ZSql & "Valor22 = " + "'" + WValor22 + "',"
                ZSql = ZSql & "Valor33 = " + "'" + WValor33 + "',"
                ZSql = ZSql & "Valor44 = " + "'" + WValor44 + "',"
                ZSql = ZSql & "Valor55 = " + "'" + WValor55 + "',"
                ZSql = ZSql & "Valor66 = " + "'" + WValor66 + "',"
                ZSql = ZSql & "Valor77 = " + "'" + WValor77 + "',"
                ZSql = ZSql & "Valor88 = " + "'" + WValor88 + "',"
                ZSql = ZSql & "Valor99 = " + "'" + WValor99 + "',"
                ZSql = ZSql & "Valor1010 = " + "'" + WValor1010 + "',"
                ZSql = ZSql & "Desde1 = " + "'" + WDesde1 + "',"
                ZSql = ZSql & "Desde2 = " + "'" + WDesde2 + "',"
                ZSql = ZSql & "Desde3 = " + "'" + WDesde3 + "',"
                ZSql = ZSql & "Desde4 = " + "'" + WDesde4 + "',"
                ZSql = ZSql & "Desde5 = " + "'" + WDesde5 + "',"
                ZSql = ZSql & "Desde6 = " + "'" + WDesde6 + "',"
                ZSql = ZSql & "Desde7 = " + "'" + WDesde7 + "',"
                ZSql = ZSql & "Desde8 = " + "'" + WDesde8 + "',"
                ZSql = ZSql & "Desde9 = " + "'" + WDesde9 + "',"
                ZSql = ZSql & "Desde10 = " + "'" + WDesde10 + "',"
                ZSql = ZSql & "Hasta1 = " + "'" + WHasta1 + "',"
                ZSql = ZSql & "Hasta2 = " + "'" + WHasta2 + "',"
                ZSql = ZSql & "Hasta3 = " + "'" + WHasta3 + "',"
                ZSql = ZSql & "Hasta4 = " + "'" + WHasta4 + "',"
                ZSql = ZSql & "Hasta5 = " + "'" + WHasta5 + "',"
                ZSql = ZSql & "Hasta6 = " + "'" + WHasta6 + "',"
                ZSql = ZSql & "Hasta7 = " + "'" + WHasta7 + "',"
                ZSql = ZSql & "Hasta8 = " + "'" + WHasta8 + "',"
                ZSql = ZSql & "Hasta9 = " + "'" + WHasta9 + "',"
                ZSql = ZSql & "Hasta10 = " + "'" + WHasta10 + "',"
                ZSql = ZSql & "Version = " + "'" + ZVersion + "',"
                ZSql = ZSql & "Fecha = " + "'" + ZFecha + "',"
                ZSql = ZSql & "Estado = " + "'" + ZEstado + "',"
                ZSql = ZSql & "ControlCambio = " + "'" + ZControlCambio + "',"
                ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProductoRE + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnifica ("
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "WDate, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "Desde1 , "
                ZSql = ZSql & "Desde2 , "
                ZSql = ZSql & "Desde3 , "
                ZSql = ZSql & "Desde4 , "
                ZSql = ZSql & "Desde5 , "
                ZSql = ZSql & "Desde6 , "
                ZSql = ZSql & "Desde7 , "
                ZSql = ZSql & "Desde8 , "
                ZSql = ZSql & "Desde9 , "
                ZSql = ZSql & "Desde10 , "
                ZSql = ZSql & "Hasta1 , "
                ZSql = ZSql & "Hasta2 , "
                ZSql = ZSql & "Hasta3 , "
                ZSql = ZSql & "Hasta4 , "
                ZSql = ZSql & "Hasta5 , "
                ZSql = ZSql & "Hasta6 , "
                ZSql = ZSql & "Hasta7 , "
                ZSql = ZSql & "Hasta8 , "
                ZSql = ZSql & "Hasta9 , "
                ZSql = ZSql & "Hasta10 , "
                ZSql = ZSql & "Version , "
                ZSql = ZSql & "Fecha , "
                ZSql = ZSql & "Estado , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WProductoRE + "',"
                ZSql = ZSql & "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
                ZSql = ZSql & "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
                ZSql = ZSql & "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
                ZSql = ZSql & "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
                ZSql = ZSql & "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
                ZSql = ZSql & "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
                ZSql = ZSql & "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
                ZSql = ZSql & "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
                ZSql = ZSql & "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
                ZSql = ZSql & "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + WValor11 + "',"
                ZSql = ZSql & "'" + WValor22 + "',"
                ZSql = ZSql & "'" + WValor33 + "',"
                ZSql = ZSql & "'" + WValor44 + "',"
                ZSql = ZSql & "'" + WValor55 + "',"
                ZSql = ZSql & "'" + WValor66 + "',"
                ZSql = ZSql & "'" + WValor77 + "',"
                ZSql = ZSql & "'" + WValor88 + "',"
                ZSql = ZSql & "'" + WValor99 + "',"
                ZSql = ZSql & "'" + WValor1010 + "',"
                ZSql = ZSql & "'" + WDesde1 + "',"
                ZSql = ZSql & "'" + WDesde2 + "',"
                ZSql = ZSql & "'" + WDesde3 + "',"
                ZSql = ZSql & "'" + WDesde4 + "',"
                ZSql = ZSql & "'" + WDesde5 + "',"
                ZSql = ZSql & "'" + WDesde6 + "',"
                ZSql = ZSql & "'" + WDesde7 + "',"
                ZSql = ZSql & "'" + WDesde8 + "',"
                ZSql = ZSql & "'" + WDesde9 + "',"
                ZSql = ZSql & "'" + WDesde10 + "',"
                ZSql = ZSql & "'" + WHasta1 + "',"
                ZSql = ZSql & "'" + WHasta2 + "',"
                ZSql = ZSql & "'" + WHasta3 + "',"
                ZSql = ZSql & "'" + WHasta4 + "',"
                ZSql = ZSql & "'" + WHasta5 + "',"
                ZSql = ZSql & "'" + WHasta6 + "',"
                ZSql = ZSql & "'" + WHasta7 + "',"
                ZSql = ZSql & "'" + WHasta8 + "',"
                ZSql = ZSql & "'" + WHasta9 + "',"
                ZSql = ZSql & "'" + WHasta10 + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZFecha + "',"
                ZSql = ZSql & "'" + ZEstado + "',"
                ZSql = ZSql & "'" + ZControlCambio + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
           
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecifUnifica SET "
                ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
                ZSql = ZSql + " Where Producto = " + "'" + WProductoRE + "'"
                            
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            Call Conecta_Empresa
            
            
            XEmpresa = WEmpresa
            Erase CargaEmpresa
            
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
                Case 2, 4, 8, 9
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case Else
            End Select
            
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "VersionII = " + "'" + XXVersion + "',"
                    ZSql = ZSql & "FechaVersionII = " + "'" + XXFechaVersion + "',"
                    ZSql = ZSql & "EstadoII = " + "'" + XXEstado + "',"
                    ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + Producto.Text + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                    WProductoRE = "RE" + Mid$(Producto.Text, 3, 10)
                    
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "VersionII = " + "'" + XXVersion + "',"
                    ZSql = ZSql & "FechaVersionII = " + "'" + XXFechaVersion + "',"
                    ZSql = ZSql & "EstadoII = " + "'" + XXEstado + "',"
                    ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WProductoRE
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
                    WProductoNK = "NK" + Mid$(Producto.Text, 3, 10)
                    
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "VersionII = " + "'" + XXVersion + "',"
                    ZSql = ZSql & "FechaVersionII = " + "'" + XXFechaVersion + "',"
                    ZSql = ZSql & "EstadoII = " + "'" + XXEstado + "',"
                    ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WProductoNK
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    ZHojaTecnica = ""
                    MiRuta = "\\servdatos\Impresion pdf\DOC" + Mid$(Producto.Text, 4, 5) + Right$(Producto.Text, 3) + "*.PDF"
                    MiNombre = Dir(MiRuta)
                    ZHojaTecnica = MiNombre
                    If Trim(ZHojaTecnica) <> "" Then
                        ZSql = ""
                        ZSql = ZSql & "UPDATE Terminado SET "
                        ZSql = ZSql & "EstadoHoja = " + "'" + "N" + "'"
                        ZSql = ZSql & " Where Codigo = " + "'" + Producto.Text + "'"
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                End If
                    
            Next Cicla
            
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
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
            Rem Call ImprimeAutomatico
            ZZProcesoImpre = 1
            Call Caratula_Click
            
            
            ZZPasaTerminado = Producto.Text
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    ZZPasaCliente = "S00102"
                Case Else
                    ZZPasaCliente = "P99999"
            End Select
        
            Call CmdLimpiar_Click
            
            PrgAltaCertificadoAuto.Show
        
        End If
        
    End If
    
    Exit Sub

WError:
    
     Resume Next
    
End Sub

Private Sub cmdDelete_Click()

    If Producto.Text <> "" Then
    
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
    
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM EspecifUnifica"
        ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            rstEspecifUnifica.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql & "DELETE EspecifUnifica"
                ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
        Call Conecta_Empresa
        
    End If
    
    Producto.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Command22.Visible = True
    Check1.Value = 0
    Producto.Text = "  -     -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
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
    DesOperador.Caption = ""
    Valor11.Text = ""
    Valor22.Text = ""
    Valor33.Text = ""
    Valor44.Text = ""
    Valor55.Text = ""
    Valor66.Text = ""
    Valor77.Text = ""
    Valor88.Text = ""
    Valor99.Text = ""
    Valor1010.Text = ""
    Desde1.Text = ""
    Desde2.Text = ""
    Desde3.Text = ""
    Desde4.Text = ""
    Desde5.Text = ""
    Desde6.Text = ""
    Desde7.Text = ""
    Desde8.Text = ""
    Desde9.Text = ""
    Desde10.Text = ""
    Hasta1.Text = ""
    Hasta2.Text = ""
    Hasta3.Text = ""
    Hasta4.Text = ""
    Hasta5.Text = ""
    Hasta6.Text = ""
    Hasta7.Text = ""
    Hasta8.Text = ""
    Hasta9.Text = ""
    Hasta10.Text = ""
    ControlCambio.Text = ""
    
    Version.Text = ""
    Fecha.Text = ""
    Estado.Text = ""
    WGraba = ""
    WGrabaII = ""
    WGrabaIII = ""
    
    Producto.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgEspeUnifica.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

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
    
    Erase Graba
    Erase GrabaII
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    XProducto = Val(Mid$(rstEspecifUnifica!Producto, 4, 5))
                    If (XProducto >= 25000 And XProducto <= 25999) Or ZTipoPedido = 4 Then
                
                        Lugar = Lugar + 1
                
                        Graba(Lugar) = rstEspecifUnifica!Producto
                
                        GrabaII(Lugar, 1) = rstEspecifUnifica!Ensayo1
                        GrabaII(Lugar, 2) = rstEspecifUnifica!Valor1
                        GrabaII(Lugar, 3) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                        GrabaII(Lugar, 4) = ""
                
                        GrabaII(Lugar, 5) = rstEspecifUnifica!Ensayo2
                        GrabaII(Lugar, 6) = rstEspecifUnifica!valor2
                        GrabaII(Lugar, 7) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                        GrabaII(Lugar, 8) = ""
                
                        GrabaII(Lugar, 9) = rstEspecifUnifica!Ensayo3
                        GrabaII(Lugar, 10) = rstEspecifUnifica!Valor3
                        GrabaII(Lugar, 11) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                        GrabaII(Lugar, 12) = ""
                
                        GrabaII(Lugar, 13) = rstEspecifUnifica!Ensayo4
                        GrabaII(Lugar, 14) = rstEspecifUnifica!valor4
                        GrabaII(Lugar, 15) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                        GrabaII(Lugar, 16) = ""
                
                        GrabaII(Lugar, 17) = rstEspecifUnifica!Ensayo5
                        GrabaII(Lugar, 18) = rstEspecifUnifica!valor5
                        GrabaII(Lugar, 19) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                        GrabaII(Lugar, 20) = ""
                
                        GrabaII(Lugar, 21) = rstEspecifUnifica!Ensayo6
                        GrabaII(Lugar, 22) = rstEspecifUnifica!valor6
                        GrabaII(Lugar, 23) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                        GrabaII(Lugar, 24) = ""
                
                        GrabaII(Lugar, 25) = rstEspecifUnifica!Ensayo7
                        GrabaII(Lugar, 26) = rstEspecifUnifica!valor7
                        GrabaII(Lugar, 27) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                        GrabaII(Lugar, 28) = ""
                
                        GrabaII(Lugar, 29) = rstEspecifUnifica!Ensayo8
                        GrabaII(Lugar, 30) = rstEspecifUnifica!valor8
                        GrabaII(Lugar, 31) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                        GrabaII(Lugar, 32) = ""
                
                        GrabaII(Lugar, 33) = rstEspecifUnifica!Ensayo9
                        GrabaII(Lugar, 34) = rstEspecifUnifica!valor9
                        GrabaII(Lugar, 35) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                        GrabaII(Lugar, 36) = ""
                
                        GrabaII(Lugar, 37) = rstEspecifUnifica!Ensayo10
                        GrabaII(Lugar, 38) = rstEspecifUnifica!valor10
                        GrabaII(Lugar, 39) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                        GrabaII(Lugar, 40) = ""
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        For CicloII = 1 To 40 Step 4
        
            WEnsayo = GrabaII(Ciclo, CicloII)
            GrabaII(Ciclo, CicloII + 3) = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WEnsayo + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                GrabaII(Ciclo, CicloII + 3) = Trim(rstEnsayos!Descripcion)
                rstEnsayos.Close
            End If
            
        Next CicloII
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To Lugar
    
        WTerminado = Graba(Ciclo)
        WPaso = "99"
        WDesPaso = "CONTROL FINAL"
        WRenglon = 0
    
        ZSql = ""
        ZSql = ZSql + "DELETE CargaV"
        ZSql = ZSql + " Where Terminado = " + "'" + WTerminado + "'"
        ZSql = ZSql + " and Paso = " + "'" + WPaso + "'"
        rsCargaV = ZSql
        Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
        
        For CicloII = 1 To 40 Step 4
        
            WEnsayo = GrabaII(Ciclo, CicloII)
            WValorI = Trim(GrabaII(Ciclo, CicloII + 1))
            WValorII = Trim(GrabaII(Ciclo, CicloII + 2))
            WDesEnsayo = Trim(GrabaII(Ciclo, CicloII + 3))
        
            If Val(WEnsayo) <> 0 Or WValorI <> "" Then
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                XPaso = WPaso
                Call Ceros(XPaso, 4)
                        
                WClave = WTerminado + XPaso + Auxi
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaV ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Paso ,"
                ZSql = ZSql + "DesPaso ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "DesEnsayo ,"
                ZSql = ZSql + "Valor )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WTerminado + "',"
                ZSql = ZSql + "'" + WPaso + "',"
                ZSql = ZSql + "'" + WDesPaso + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + WEnsayo + "',"
                ZSql = ZSql + "'" + WDesEnsayo + "',"
                ZSql = ZSql + "'" + WValorI + "')"
                
                rsCargaV = ZSql
                Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
            If WValorII <> "" Then
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                XPaso = WPaso
                Call Ceros(XPaso, 4)
                        
                WClave = WTerminado + XPaso + Auxi
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaV ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Paso ,"
                ZSql = ZSql + "DesPaso ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "DesEnsayo ,"
                ZSql = ZSql + "Valor )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WTerminado + "',"
                ZSql = ZSql + "'" + WPaso + "',"
                ZSql = ZSql + "'" + WDesPaso + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + WEnsayo + "',"
                ZSql = ZSql + "'" + WDesEnsayo + "',"
                ZSql = ZSql + "'" + WValorII + "')"
                
                rsCargaV = ZSql
                Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next CicloII
        
    Next Ciclo
    
    Stop

End Sub

Private Sub Ensayo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor1.SetFocus
                    Else
            Descri1.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor2.SetFocus
                    Else
            descri2.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor3.SetFocus
                    Else
            Descri3.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor4.SetFocus
                    Else
            Descri4.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor5.SetFocus
                    Else
            Descri5.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor6.SetFocus
                    Else
            Descri6.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor7.SetFocus
                    Else
            Descri7.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor8.SetFocus
                    Else
            Descri8.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor9.SetFocus
                    Else
            Descri9.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor10.SetFocus
                    Else
            Descri10.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
 WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 10
















    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEspeUnifica.Caption = "Ingreso de Especificaciones de Productos Terminados (Unificado):  " + !Nombre
        End If
    End With
    
    Producto.Text = "  -     -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
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
    DesOperador.Caption = ""
    Valor11.Text = ""
    Valor22.Text = ""
    Valor33.Text = ""
    Valor44.Text = ""
    Valor55.Text = ""
    Valor66.Text = ""
    Valor77.Text = ""
    Valor88.Text = ""
    Valor99.Text = ""
    Valor1010.Text = ""
    Desde1.Text = ""
    Desde2.Text = ""
    Desde3.Text = ""
    Desde4.Text = ""
    Desde5.Text = ""
    Desde6.Text = ""
    Desde7.Text = ""
    Desde8.Text = ""
    Desde9.Text = ""
    Desde10.Text = ""
    Hasta1.Text = ""
    Hasta2.Text = ""
    Hasta3.Text = ""
    Hasta4.Text = ""
    Hasta5.Text = ""
    Hasta6.Text = ""
    Hasta7.Text = ""
    Hasta8.Text = ""
    Hasta9.Text = ""
    Hasta10.Text = ""
    ControlCambio.Text = ""
    
    Version.Text = ""
    Fecha.Text = ""
    Estado.Text = ""
        
    DesOperador.Caption = ""
    WGraba = ""
    WGrabaII = ""
    WGrabaIII = ""
    EmpresaActual = WEmpresa
    
End Sub

Private Sub ImprePdf_Click()
    Desde.Text = Producto.Text
    Hasta.Text = Producto.Text
    ZZProceso = 1
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Call Acepta_Click
End Sub

Private Sub Listado_Click()
    Command22.Visible = False
    ZZProceso = 0
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Revalida_Click()

    On Error GoTo WError
    
    If WGrabaII <> "S" Then
    
        Call Ingresa_ClaveII

               Else

        If Producto.Text <> "" Then
        
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
        
            WProducto = Producto.Text
            ZEstado = "S"
            ZObservaciones = ""
            
            ZSql = ""
            ZSql = ZSql & "UPDATE EspecifUnifica SET "
            ZSql = ZSql & "Estado = " + "'" + ZEstado + "',"
            ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
            ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
            Call Conecta_Empresa
            
            
            
            XEmpresa = WEmpresa
            Erase CargaEmpresa
            
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
                Case 2, 4, 8, 9
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                Case Else
            End Select
            
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    WCodigo = WProducto
                    WEstadoII = ZEstado
                    WObservaII = ZObservaciones
            
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "EstadoII = " + "'" + ZEstado + "',"
                    ZSql = ZSql & "ObservaII = " + "'" + ZObservaciones + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + WCodigo + "'"
                    
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                    
            Next Cicla
            
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
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            Call CmdLimpiar_Click
            Producto.SetFocus
        
        End If
        
    End If
    
    Exit Sub

WError:
     Resume Next

End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor11.SetFocus
    End If
End Sub

Private Sub Valor11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde1.SetFocus
    End If
End Sub

Private Sub Desde1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub





Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor22.SetFocus
    End If
End Sub

Private Sub Valor22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde2.SetFocus
    End If
End Sub

Private Sub Desde2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor33.SetFocus
    End If
End Sub

Private Sub Valor33_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde3.SetFocus
    End If
End Sub

Private Sub Desde3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor44.SetFocus
    End If
End Sub

Private Sub Valor44_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde4.SetFocus
    End If
End Sub

Private Sub Desde4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor55.SetFocus
    End If
End Sub

Private Sub Valor55_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde5.SetFocus
    End If
End Sub

Private Sub Desde5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo6.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor66.SetFocus
    End If
End Sub

Private Sub Valor66_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde6.SetFocus
    End If
End Sub

Private Sub Desde6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta7.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo7.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor77.SetFocus
    End If
End Sub

Private Sub Valor77_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde7.SetFocus
    End If
End Sub

Private Sub Desde7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta7.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo8.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor88.SetFocus
    End If
End Sub

Private Sub Valor88_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde8.SetFocus
    End If
End Sub

Private Sub Desde8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta8.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo9.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor99.SetFocus
    End If
End Sub

Private Sub Valor99_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde9.SetFocus
    End If
End Sub

Private Sub Desde9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta9.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo10.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1010.SetFocus
    End If
End Sub

Private Sub Valor1010_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde10.SetFocus
    End If
End Sub

Private Sub Desde10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta10.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ControlCambio.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ControlCambio_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo1.SetFocus
    End If
End Sub


Private Sub Valor1Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor11Ing.SetFocus
    End If
End Sub

Private Sub Valor11Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor2Ing.SetFocus
    End If
End Sub

Private Sub Valor2Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor22Ing.SetFocus
    End If
End Sub

Private Sub Valor22Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor3Ing.SetFocus
    End If
End Sub

Private Sub Valor3Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor33Ing.SetFocus
    End If
End Sub

Private Sub Valor33Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor4Ing.SetFocus
    End If
End Sub

Private Sub Valor4Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor44Ing.SetFocus
    End If
End Sub

Private Sub Valor44Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor5Ing.SetFocus
    End If
End Sub

Private Sub Valor5Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor55Ing.SetFocus
    End If
End Sub

Private Sub Valor55Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor6Ing.SetFocus
    End If
End Sub

Private Sub Valor6Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor66Ing.SetFocus
    End If
End Sub

Private Sub Valor66Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor7Ing.SetFocus
    End If
End Sub

Private Sub Valor7Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor77Ing.SetFocus
    End If
End Sub

Private Sub Valor77Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor8Ing.SetFocus
    End If
End Sub

Private Sub Valor8Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor88Ing.SetFocus
    End If
End Sub

Private Sub Valor88Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor9Ing.SetFocus
    End If
End Sub

Private Sub Valor9Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor99Ing.SetFocus
    End If
End Sub

Private Sub Valor99Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor10Ing.SetFocus
    End If
End Sub

Private Sub Valor10Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1010Ing.SetFocus
    End If
End Sub

Private Sub Valor1010Ing_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1Ing.SetFocus
    End If
End Sub





Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
        
          Check1.Value = 0
            
            
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
            
            ZSql = ""
            ZSql = ZSql & "Select EspecifUnifica.Producto"
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                rstEspecifUnifica.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                    Else
                Call Conecta_Empresa
                WProducto = Producto.Text
                CmdLimpiar_Click
                Producto.Text = WProducto
            End If
            
        End If
        Ensayo1.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
   EntraIngles.Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Ensayos"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    Command22.Visible = False
    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
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
            
            
            
            
            spEnsayo = "ListaEnsayos"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                With rstEnsayo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                            pantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayo.Close
            End If
            
            Call Conecta_Empresa
            
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Command22.Visible = True
    pantalla.Visible = False
    EntraIngles.Visible = True
    Select Case XIndice
        Case 0
            Entra$ = "S"
            
            If Val(Ensayo1.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo1.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo1_Keypress(13)
            End If
            
            If Val(Ensayo2.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo2.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo2_Keypress(13)
            End If
            
            If Val(Ensayo3.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo3.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo3_Keypress(13)
            End If
            
            If Val(Ensayo4.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo4.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo4_Keypress(13)
            End If
            
            If Val(Ensayo5.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo5.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo5_Keypress(13)
            End If
            
            If Val(Ensayo6.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo6.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo6_Keypress(13)
            End If
            
            If Val(Ensayo7.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo7.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo7_Keypress(13)
            End If
            
            If Val(Ensayo8.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo8.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo8_Keypress(13)
            End If
            
            If Val(Ensayo9.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo9.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo9_Keypress(13)
            End If
            
            If Val(Ensayo10.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo10.Text = Val(WIndice.List(Indice))
                Entra = "N"
                Call Ensayo10_Keypress(13)
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Anterior_Click()

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

    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM EspecifUnifica"
    ZSql = ZSql & " Where EspecifUnifica.Producto < " + "'" + Producto.Text + "'"
    ZSql = ZSql & " Order by EspecifUnifica.Producto"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveLast
            Producto.Text = rstEspecifUnifica!Producto
        End With
        rstEspecifUnifica.Close
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Ingreso de Especificaciones de Productos Terminados")
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Producto.SetFocus
    
End Sub

Private Sub Primer_Click()

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
    
    ZSql = ""
    ZSql = ZSql & "Select Min(Producto) as [ProductoMenor]"
    ZSql = ZSql & " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        rstEspecifUnifica.MoveFirst
        Producto.Text = rstEspecifUnifica!ProductoMenor
        rstEspecifUnifica.Close
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Producto.SetFocus
    
 End Sub

Private Sub Ultimo_Click()

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
    
    ZSql = ""
    ZSql = ZSql & "Select Max(Producto) as [ProductoMayor]"
    ZSql = ZSql & " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        rstEspecifUnifica.MoveLast
        Producto.Text = rstEspecifUnifica!ProductoMayor
        rstEspecifUnifica.Close
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Producto.SetFocus

 End Sub

Private Sub Siguiente_Click()

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
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM EspecifUnifica"
    ZSql = ZSql & " Where EspecifUnifica.Producto > " + "'" + Producto.Text + "'"
    ZSql = ZSql & " Order by EspecifUnifica.Producto"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Producto.Text = rstEspecifUnifica!Producto
        End With
        rstEspecifUnifica.Close
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Ingreso de Especificaciones de Productos Terminados")
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Producto.SetFocus
    
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
        ZGRABAII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Call cmdAdd_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub

Sub Ingresa_ClaveII()
    WClaveII.Text = ""
    XClaveII.Visible = True
    WClaveII.SetFocus
End Sub

Private Sub CancelaGrabaII_Click()
    XClaveII.Visible = False
End Sub

Private Sub WClaveII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGrabaII = "N"
        ZGRABAII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClaveII.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAII = "S" Then
            WGrabaII = "S"
            XClaveII.Visible = False
            Call Revalida_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClaveII.SetFocus
        End If
        
    End If
End Sub


Sub Ingresa_ClaveIII()
    WClaveIII.Text = ""
    XClaveIII.Visible = True
    WClaveIII.SetFocus
End Sub

Private Sub CancelaGrabaIII_Click()
    XClaveIII.Visible = False
End Sub

Private Sub WClaveIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGrabaIII = "N"
        ZGRABAIII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClaveIII.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAIII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAIII = "S" Then
            WGrabaIII = "S"
            XClaveIII.Visible = False
            Call EntraIngles_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClaveIII.SetFocus
        End If
        
    End If
End Sub


Private Sub Command2_Click()

    Dim ZZVector(10000) As String
        
    ZZLugar = 0
    Erase ZZVector
        
    ZSql = ""
    ZSql = ZSql + "Select EspecifUnifica.Producto"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Left$(UCase(rstEspecifUnifica!Producto), 2) = "PT" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = rstEspecifUnifica!Producto
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    
    Stop
    
    For Ciclo = 1 To ZZLugar
    
        ZZProducto = ZZVector(Ciclo)
        
        Producto.Text = ZZProducto
        Call Producto_Keypress(13)
        
        DoEvents
        

        WProducto = Producto.Text
        
        WEnsayo1 = Ensayo1.Text
        WEnsayo2 = Ensayo2.Text
        WEnsayo3 = Ensayo3.Text
        WEnsayo4 = Ensayo4.Text
        WEnsayo5 = Ensayo5.Text
        WEnsayo6 = Ensayo6.Text
        WEnsayo7 = Ensayo7.Text
        WEnsayo8 = Ensayo8.Text
        WEnsayo9 = Ensayo9.Text
        WEnsayo10 = Ensayo10.Text
        
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
        WValor11 = Valor11.Text
        WValor22 = Valor22.Text
        WValor33 = Valor33.Text
        WValor44 = Valor44.Text
        WValor55 = Valor55.Text
        WValor66 = Valor66.Text
        WValor77 = Valor77.Text
        WValor88 = Valor88.Text
        WValor99 = Valor99.Text
        WValor1010 = Valor1010.Text
        
        WDesde1 = Desde1.Text
        WDesde2 = Desde2.Text
        WDesde3 = Desde3.Text
        WDesde4 = Desde4.Text
        WDesde5 = Desde5.Text
        WDesde6 = Desde6.Text
        WDesde7 = Desde7.Text
        WDesde8 = Desde8.Text
        WDesde9 = Desde9.Text
        WDesde10 = Desde10.Text
        
        WHasta1 = Hasta1.Text
        WHasta2 = Hasta2.Text
        WHasta3 = Hasta3.Text
        WHasta4 = Hasta4.Text
        WHasta5 = Hasta5.Text
        WHasta6 = Hasta6.Text
        WHasta7 = Hasta7.Text
        WHasta8 = Hasta8.Text
        WHasta9 = Hasta9.Text
        WHasta10 = Hasta10.Text
        
        WDate = Date$
        
        ZObservaciones = ""
        ZOperador = "0"
            
            
        WProductoNK = "NK" + Mid$(Producto.Text, 3, 10)
         
         ZSql = ""
         ZSql = ZSql & "Select *"
         ZSql = ZSql & " FROM EspecifUnifica"
         ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + WProductoNK + "'"
         spEspecifUnifica = ZSql
         Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
         If rstEspecifUnifica.RecordCount > 0 Then
             rstEspecifUnifica.Close
             
             ZSql = ""
             ZSql = ZSql & "UPDATE EspecifUnifica SET "
             ZSql = ZSql & "Ensayo1 = " + "'" + WEnsayo1 + "',"
             ZSql = ZSql & "Valor1 = " + "'" + WValor1 + "',"
             ZSql = ZSql & "Ensayo2 = " + "'" + WEnsayo2 + "',"
             ZSql = ZSql & "Valor2 = " + "'" + WValor2 + "',"
             ZSql = ZSql & "Ensayo3 = " + "'" + WEnsayo3 + "',"
             ZSql = ZSql & "Valor3 = " + "'" + WValor3 + "',"
             ZSql = ZSql & "Ensayo4 = " + "'" + WEnsayo4 + "',"
             ZSql = ZSql & "Valor4 = " + "'" + WValor4 + "',"
             ZSql = ZSql & "Ensayo5 = " + "'" + WEnsayo5 + "',"
             ZSql = ZSql & "Valor5 = " + "'" + WValor5 + "',"
             ZSql = ZSql & "Ensayo6 = " + "'" + WEnsayo6 + "',"
             ZSql = ZSql & "Valor6 = " + "'" + WValor6 + "',"
             ZSql = ZSql & "Ensayo7 = " + "'" + WEnsayo7 + "',"
             ZSql = ZSql & "Valor7 = " + "'" + WValor7 + "',"
             ZSql = ZSql & "Ensayo8 = " + "'" + WEnsayo8 + "',"
             ZSql = ZSql & "Valor8 = " + "'" + WValor8 + "',"
             ZSql = ZSql & "Ensayo9 = " + "'" + WEnsayo9 + "',"
             ZSql = ZSql & "Valor9 = " + "'" + WValor9 + "',"
             ZSql = ZSql & "Ensayo10 = " + "'" + WEnsayo10 + "',"
             ZSql = ZSql & "Valor10 = " + "'" + WValor10 + "',"
             ZSql = ZSql & "WDate = " + "'" + WDate + "',"
             ZSql = ZSql & "Valor11 = " + "'" + WValor11 + "',"
             ZSql = ZSql & "Valor22 = " + "'" + WValor22 + "',"
             ZSql = ZSql & "Valor33 = " + "'" + WValor33 + "',"
             ZSql = ZSql & "Valor44 = " + "'" + WValor44 + "',"
             ZSql = ZSql & "Valor55 = " + "'" + WValor55 + "',"
             ZSql = ZSql & "Valor66 = " + "'" + WValor66 + "',"
             ZSql = ZSql & "Valor77 = " + "'" + WValor77 + "',"
             ZSql = ZSql & "Valor88 = " + "'" + WValor88 + "',"
             ZSql = ZSql & "Valor99 = " + "'" + WValor99 + "',"
             ZSql = ZSql & "Valor1010 = " + "'" + WValor1010 + "',"
             ZSql = ZSql & "Desde1 = " + "'" + WDesde1 + "',"
             ZSql = ZSql & "Desde2 = " + "'" + WDesde2 + "',"
             ZSql = ZSql & "Desde3 = " + "'" + WDesde3 + "',"
             ZSql = ZSql & "Desde4 = " + "'" + WDesde4 + "',"
             ZSql = ZSql & "Desde5 = " + "'" + WDesde5 + "',"
             ZSql = ZSql & "Desde6 = " + "'" + WDesde6 + "',"
             ZSql = ZSql & "Desde7 = " + "'" + WDesde7 + "',"
             ZSql = ZSql & "Desde8 = " + "'" + WDesde8 + "',"
             ZSql = ZSql & "Desde9 = " + "'" + WDesde9 + "',"
             ZSql = ZSql & "Desde10 = " + "'" + WDesde10 + "',"
             ZSql = ZSql & "Hasta1 = " + "'" + WHasta1 + "',"
             ZSql = ZSql & "Hasta2 = " + "'" + WHasta2 + "',"
             ZSql = ZSql & "Hasta3 = " + "'" + WHasta3 + "',"
             ZSql = ZSql & "Hasta4 = " + "'" + WHasta4 + "',"
             ZSql = ZSql & "Hasta5 = " + "'" + WHasta5 + "',"
             ZSql = ZSql & "Hasta6 = " + "'" + WHasta6 + "',"
             ZSql = ZSql & "Hasta7 = " + "'" + WHasta7 + "',"
             ZSql = ZSql & "Hasta8 = " + "'" + WHasta8 + "',"
             ZSql = ZSql & "Hasta9 = " + "'" + WHasta9 + "',"
             ZSql = ZSql & "Hasta10 = " + "'" + WHasta10 + "',"
             ZSql = ZSql & "Version = " + "'" + Version.Text + "',"
             ZSql = ZSql & "Fecha = " + "'" + Fecha.Text + "',"
             ZSql = ZSql & "Estado = " + "'" + Estado.Text + "',"
             ZSql = ZSql & "ControlCambio = " + "'" + ControlCambio.Text + "',"
             ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
             ZSql = ZSql & " Where Producto = " + "'" + WProductoNK + "'"
                     
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
             
                     Else
                 
             ZSql = ""
             ZSql = ZSql & "INSERT INTO EspecifUnifica ("
             ZSql = ZSql & "Producto, "
             ZSql = ZSql & "Ensayo1, Valor1, "
             ZSql = ZSql & "Ensayo2, Valor2, "
             ZSql = ZSql & "Ensayo3, Valor3, "
             ZSql = ZSql & "Ensayo4, Valor4, "
             ZSql = ZSql & "Ensayo5, Valor5, "
             ZSql = ZSql & "Ensayo6, Valor6, "
             ZSql = ZSql & "Ensayo7, Valor7, "
             ZSql = ZSql & "Ensayo8, Valor8, "
             ZSql = ZSql & "Ensayo9, Valor9, "
             ZSql = ZSql & "Ensayo10, Valor10, "
             ZSql = ZSql & "WDate, "
             ZSql = ZSql & "Valor11 , "
             ZSql = ZSql & "Valor22 , "
             ZSql = ZSql & "Valor33 , "
             ZSql = ZSql & "Valor44 , "
             ZSql = ZSql & "Valor55 , "
             ZSql = ZSql & "Valor66 , "
             ZSql = ZSql & "Valor77 , "
             ZSql = ZSql & "Valor88 , "
             ZSql = ZSql & "Valor99 , "
             ZSql = ZSql & "Valor1010 , "
             ZSql = ZSql & "Desde1 , "
             ZSql = ZSql & "Desde2 , "
             ZSql = ZSql & "Desde3 , "
             ZSql = ZSql & "Desde4 , "
             ZSql = ZSql & "Desde5 , "
             ZSql = ZSql & "Desde6 , "
             ZSql = ZSql & "Desde7 , "
             ZSql = ZSql & "Desde8 , "
             ZSql = ZSql & "Desde9 , "
             ZSql = ZSql & "Desde10 , "
             ZSql = ZSql & "Hasta1 , "
             ZSql = ZSql & "Hasta2 , "
             ZSql = ZSql & "Hasta3 , "
             ZSql = ZSql & "Hasta4 , "
             ZSql = ZSql & "Hasta5 , "
             ZSql = ZSql & "Hasta6 , "
             ZSql = ZSql & "Hasta7 , "
             ZSql = ZSql & "Hasta8 , "
             ZSql = ZSql & "Hasta9 , "
             ZSql = ZSql & "Hasta10 , "
             ZSql = ZSql & "Version , "
             ZSql = ZSql & "Fecha , "
             ZSql = ZSql & "Estado , "
             ZSql = ZSql & "ControlCambio , "
             ZSql = ZSql & "Observaciones) "
             ZSql = ZSql & "Values ("
             ZSql = ZSql & "'" + WProductoNK + "',"
             ZSql = ZSql & "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
             ZSql = ZSql & "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
             ZSql = ZSql & "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
             ZSql = ZSql & "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
             ZSql = ZSql & "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
             ZSql = ZSql & "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
             ZSql = ZSql & "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
             ZSql = ZSql & "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
             ZSql = ZSql & "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
             ZSql = ZSql & "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
             ZSql = ZSql & "'" + WDate + "',"
             ZSql = ZSql & "'" + WValor11 + "',"
             ZSql = ZSql & "'" + WValor22 + "',"
             ZSql = ZSql & "'" + WValor33 + "',"
             ZSql = ZSql & "'" + WValor44 + "',"
             ZSql = ZSql & "'" + WValor55 + "',"
             ZSql = ZSql & "'" + WValor66 + "',"
             ZSql = ZSql & "'" + WValor77 + "',"
             ZSql = ZSql & "'" + WValor88 + "',"
             ZSql = ZSql & "'" + WValor99 + "',"
             ZSql = ZSql & "'" + WValor1010 + "',"
             ZSql = ZSql & "'" + WDesde1 + "',"
             ZSql = ZSql & "'" + WDesde2 + "',"
             ZSql = ZSql & "'" + WDesde3 + "',"
             ZSql = ZSql & "'" + WDesde4 + "',"
             ZSql = ZSql & "'" + WDesde5 + "',"
             ZSql = ZSql & "'" + WDesde6 + "',"
             ZSql = ZSql & "'" + WDesde7 + "',"
             ZSql = ZSql & "'" + WDesde8 + "',"
             ZSql = ZSql & "'" + WDesde9 + "',"
             ZSql = ZSql & "'" + WDesde10 + "',"
             ZSql = ZSql & "'" + WHasta1 + "',"
             ZSql = ZSql & "'" + WHasta2 + "',"
             ZSql = ZSql & "'" + WHasta3 + "',"
             ZSql = ZSql & "'" + WHasta4 + "',"
             ZSql = ZSql & "'" + WHasta5 + "',"
             ZSql = ZSql & "'" + WHasta6 + "',"
             ZSql = ZSql & "'" + WHasta7 + "',"
             ZSql = ZSql & "'" + WHasta8 + "',"
             ZSql = ZSql & "'" + WHasta9 + "',"
             ZSql = ZSql & "'" + WHasta10 + "',"
             ZSql = ZSql & "'" + Version.Text + "',"
             ZSql = ZSql & "'" + Fecha.Text + "',"
             ZSql = ZSql & "'" + Estado.Text + "',"
             ZSql = ZSql & "'" + ControlCambio.Text + "',"
             ZSql = ZSql & "'" + ZObservaciones + "')"
        
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
             
             ZSql = ""
             ZSql = ZSql + "UPDATE EspecifUnifica SET "
             ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
             ZSql = ZSql + " Where Producto = " + "'" + WProductoNK + "'"
                         
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
         
         End If
            
            
            
            
            
         WProductoRE = "RE" + Mid$(Producto.Text, 3, 10)
         
         ZSql = ""
         ZSql = ZSql & "Select *"
         ZSql = ZSql & " FROM EspecifUnifica"
         ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + WProductoRE + "'"
         spEspecifUnifica = ZSql
         Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
         If rstEspecifUnifica.RecordCount > 0 Then
             
             rstEspecifUnifica.Close
             
             ZSql = ""
             ZSql = ZSql & "UPDATE EspecifUnifica SET "
             ZSql = ZSql & "Ensayo1 = " + "'" + WEnsayo1 + "',"
             ZSql = ZSql & "Valor1 = " + "'" + WValor1 + "',"
             ZSql = ZSql & "Ensayo2 = " + "'" + WEnsayo2 + "',"
             ZSql = ZSql & "Valor2 = " + "'" + WValor2 + "',"
             ZSql = ZSql & "Ensayo3 = " + "'" + WEnsayo3 + "',"
             ZSql = ZSql & "Valor3 = " + "'" + WValor3 + "',"
             ZSql = ZSql & "Ensayo4 = " + "'" + WEnsayo4 + "',"
             ZSql = ZSql & "Valor4 = " + "'" + WValor4 + "',"
             ZSql = ZSql & "Ensayo5 = " + "'" + WEnsayo5 + "',"
             ZSql = ZSql & "Valor5 = " + "'" + WValor5 + "',"
             ZSql = ZSql & "Ensayo6 = " + "'" + WEnsayo6 + "',"
             ZSql = ZSql & "Valor6 = " + "'" + WValor6 + "',"
             ZSql = ZSql & "Ensayo7 = " + "'" + WEnsayo7 + "',"
             ZSql = ZSql & "Valor7 = " + "'" + WValor7 + "',"
             ZSql = ZSql & "Ensayo8 = " + "'" + WEnsayo8 + "',"
             ZSql = ZSql & "Valor8 = " + "'" + WValor8 + "',"
             ZSql = ZSql & "Ensayo9 = " + "'" + WEnsayo9 + "',"
             ZSql = ZSql & "Valor9 = " + "'" + WValor9 + "',"
             ZSql = ZSql & "Ensayo10 = " + "'" + WEnsayo10 + "',"
             ZSql = ZSql & "Valor10 = " + "'" + WValor10 + "',"
             ZSql = ZSql & "WDate = " + "'" + WDate + "',"
             ZSql = ZSql & "Valor11 = " + "'" + WValor11 + "',"
             ZSql = ZSql & "Valor22 = " + "'" + WValor22 + "',"
             ZSql = ZSql & "Valor33 = " + "'" + WValor33 + "',"
             ZSql = ZSql & "Valor44 = " + "'" + WValor44 + "',"
             ZSql = ZSql & "Valor55 = " + "'" + WValor55 + "',"
             ZSql = ZSql & "Valor66 = " + "'" + WValor66 + "',"
             ZSql = ZSql & "Valor77 = " + "'" + WValor77 + "',"
             ZSql = ZSql & "Valor88 = " + "'" + WValor88 + "',"
             ZSql = ZSql & "Valor99 = " + "'" + WValor99 + "',"
             ZSql = ZSql & "Valor1010 = " + "'" + WValor1010 + "',"
             ZSql = ZSql & "Desde1 = " + "'" + WDesde1 + "',"
             ZSql = ZSql & "Desde2 = " + "'" + WDesde2 + "',"
             ZSql = ZSql & "Desde3 = " + "'" + WDesde3 + "',"
             ZSql = ZSql & "Desde4 = " + "'" + WDesde4 + "',"
             ZSql = ZSql & "Desde5 = " + "'" + WDesde5 + "',"
             ZSql = ZSql & "Desde6 = " + "'" + WDesde6 + "',"
             ZSql = ZSql & "Desde7 = " + "'" + WDesde7 + "',"
             ZSql = ZSql & "Desde8 = " + "'" + WDesde8 + "',"
             ZSql = ZSql & "Desde9 = " + "'" + WDesde9 + "',"
             ZSql = ZSql & "Desde10 = " + "'" + WDesde10 + "',"
             ZSql = ZSql & "Hasta1 = " + "'" + WHasta1 + "',"
             ZSql = ZSql & "Hasta2 = " + "'" + WHasta2 + "',"
             ZSql = ZSql & "Hasta3 = " + "'" + WHasta3 + "',"
             ZSql = ZSql & "Hasta4 = " + "'" + WHasta4 + "',"
             ZSql = ZSql & "Hasta5 = " + "'" + WHasta5 + "',"
             ZSql = ZSql & "Hasta6 = " + "'" + WHasta6 + "',"
             ZSql = ZSql & "Hasta7 = " + "'" + WHasta7 + "',"
             ZSql = ZSql & "Hasta8 = " + "'" + WHasta8 + "',"
             ZSql = ZSql & "Hasta9 = " + "'" + WHasta9 + "',"
             ZSql = ZSql & "Hasta10 = " + "'" + WHasta10 + "',"
             ZSql = ZSql & "Version = " + "'" + Version.Text + "',"
             ZSql = ZSql & "Fecha = " + "'" + Fecha.Text + "',"
             ZSql = ZSql & "Estado = " + "'" + Estado.Text + "',"
             ZSql = ZSql & "ControlCambio = " + "'" + ControlCambio.Text + "',"
             ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
             ZSql = ZSql & " Where Producto = " + "'" + WProductoRE + "'"
                     
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
             
             
                     Else
                 
             ZSql = ""
             ZSql = ZSql & "INSERT INTO EspecifUnifica ("
             ZSql = ZSql & "Producto, "
             ZSql = ZSql & "Ensayo1, Valor1, "
             ZSql = ZSql & "Ensayo2, Valor2, "
             ZSql = ZSql & "Ensayo3, Valor3, "
             ZSql = ZSql & "Ensayo4, Valor4, "
             ZSql = ZSql & "Ensayo5, Valor5, "
             ZSql = ZSql & "Ensayo6, Valor6, "
             ZSql = ZSql & "Ensayo7, Valor7, "
             ZSql = ZSql & "Ensayo8, Valor8, "
             ZSql = ZSql & "Ensayo9, Valor9, "
             ZSql = ZSql & "Ensayo10, Valor10, "
             ZSql = ZSql & "WDate, "
             ZSql = ZSql & "Valor11 , "
             ZSql = ZSql & "Valor22 , "
             ZSql = ZSql & "Valor33 , "
             ZSql = ZSql & "Valor44 , "
             ZSql = ZSql & "Valor55 , "
             ZSql = ZSql & "Valor66 , "
             ZSql = ZSql & "Valor77 , "
             ZSql = ZSql & "Valor88 , "
             ZSql = ZSql & "Valor99 , "
             ZSql = ZSql & "Valor1010 , "
             ZSql = ZSql & "Desde1 , "
             ZSql = ZSql & "Desde2 , "
             ZSql = ZSql & "Desde3 , "
             ZSql = ZSql & "Desde4 , "
             ZSql = ZSql & "Desde5 , "
             ZSql = ZSql & "Desde6 , "
             ZSql = ZSql & "Desde7 , "
             ZSql = ZSql & "Desde8 , "
             ZSql = ZSql & "Desde9 , "
             ZSql = ZSql & "Desde10 , "
             ZSql = ZSql & "Hasta1 , "
             ZSql = ZSql & "Hasta2 , "
             ZSql = ZSql & "Hasta3 , "
             ZSql = ZSql & "Hasta4 , "
             ZSql = ZSql & "Hasta5 , "
             ZSql = ZSql & "Hasta6 , "
             ZSql = ZSql & "Hasta7 , "
             ZSql = ZSql & "Hasta8 , "
             ZSql = ZSql & "Hasta9 , "
             ZSql = ZSql & "Hasta10 , "
             ZSql = ZSql & "Version , "
             ZSql = ZSql & "Fecha , "
             ZSql = ZSql & "Estado , "
             ZSql = ZSql & "ControlCambio , "
             ZSql = ZSql & "Observaciones) "
             ZSql = ZSql & "Values ("
             ZSql = ZSql & "'" + WProductoRE + "',"
             ZSql = ZSql & "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
             ZSql = ZSql & "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
             ZSql = ZSql & "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
             ZSql = ZSql & "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
             ZSql = ZSql & "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
             ZSql = ZSql & "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
             ZSql = ZSql & "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
             ZSql = ZSql & "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
             ZSql = ZSql & "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
             ZSql = ZSql & "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
             ZSql = ZSql & "'" + WDate + "',"
             ZSql = ZSql & "'" + WValor11 + "',"
             ZSql = ZSql & "'" + WValor22 + "',"
             ZSql = ZSql & "'" + WValor33 + "',"
             ZSql = ZSql & "'" + WValor44 + "',"
             ZSql = ZSql & "'" + WValor55 + "',"
             ZSql = ZSql & "'" + WValor66 + "',"
             ZSql = ZSql & "'" + WValor77 + "',"
             ZSql = ZSql & "'" + WValor88 + "',"
             ZSql = ZSql & "'" + WValor99 + "',"
             ZSql = ZSql & "'" + WValor1010 + "',"
             ZSql = ZSql & "'" + WDesde1 + "',"
             ZSql = ZSql & "'" + WDesde2 + "',"
             ZSql = ZSql & "'" + WDesde3 + "',"
             ZSql = ZSql & "'" + WDesde4 + "',"
             ZSql = ZSql & "'" + WDesde5 + "',"
             ZSql = ZSql & "'" + WDesde6 + "',"
             ZSql = ZSql & "'" + WDesde7 + "',"
             ZSql = ZSql & "'" + WDesde8 + "',"
             ZSql = ZSql & "'" + WDesde9 + "',"
             ZSql = ZSql & "'" + WDesde10 + "',"
             ZSql = ZSql & "'" + WHasta1 + "',"
             ZSql = ZSql & "'" + WHasta2 + "',"
             ZSql = ZSql & "'" + WHasta3 + "',"
             ZSql = ZSql & "'" + WHasta4 + "',"
             ZSql = ZSql & "'" + WHasta5 + "',"
             ZSql = ZSql & "'" + WHasta6 + "',"
             ZSql = ZSql & "'" + WHasta7 + "',"
             ZSql = ZSql & "'" + WHasta8 + "',"
             ZSql = ZSql & "'" + WHasta9 + "',"
             ZSql = ZSql & "'" + WHasta10 + "',"
             ZSql = ZSql & "'" + Version.Text + "',"
             ZSql = ZSql & "'" + Fecha.Text + "',"
             ZSql = ZSql & "'" + Estado.Text + "',"
             ZSql = ZSql & "'" + ControlCambio.Text + "',"
             ZSql = ZSql & "'" + ZObservaciones + "')"
        
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
             
             ZSql = ""
             ZSql = ZSql + "UPDATE EspecifUnifica SET "
             ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
             ZSql = ZSql + " Where Producto = " + "'" + WProductoRE + "'"
                         
             spEspecifUnifica = ZSql
             Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
             
        End If
         
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
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
            Case 2, 4, 8, 9
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case Else
        End Select
            
        For Cicla = 1 To 7
        
            If CargaEmpresa(Cicla, 1) <> "" Then
        
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                XXObservaciones = ""
        
                WProductoRE = "RE" + Mid$(Producto.Text, 3, 10)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "VersionII = " + "'" + Version.Text + "',"
                ZSql = ZSql & "FechaVersionII = " + "'" + Fecha.Text + "',"
                ZSql = ZSql & "EstadoII = " + "'" + Estado.Text + "',"
                ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + WProductoRE + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
                WProductoNK = "NK" + Mid$(Producto.Text, 3, 10)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "VersionII = " + "'" + Version.Text + "',"
                ZSql = ZSql & "FechaVersionII = " + "'" + Fecha.Text + "',"
                ZSql = ZSql & "EstadoII = " + "'" + Estado.Text + "',"
                ZSql = ZSql & "ObservaII = " + "'" + XXObservaciones + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + WProductoNK + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next Cicla
        
        Call Conecta_Empresa
         
         
    Next Ciclo
        
    Call Conecta_Empresa


End Sub



Private Sub CaratulaI_Click()
    ZZProcesoImpre = 0
    Call Caratula_Click
End Sub

Private Sub CaratulaII_Click()
    ZZProcesoImpre = 1
    Call Caratula_Click
End Sub

Private Sub Caratula_Click()
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            ZZZCliente = "S00102"
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            ZZZCliente = "P99999"
    End Select
    
    ZArticulo = Producto.Text
    ZProducto = Producto.Text
    ZLote = ""
    ZCantidad = ""
    ZCliente = ZZZCliente
    ZRazon = ""
        
    Erase ZZOpcion
    Erase ZZValor
    Erase ZZEnsayo
    Erase ZZStd
    Erase ZZDescri
    Erase ZZDescriII
        
    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM AltaCertificado"
    ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
    ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZZZCliente + "'"
    spAltaCertificado = ZSql
    Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
    If rstAltaCertificado.RecordCount > 0 Then
        ZZOpcion(1) = rstAltaCertificado!Opcion1
        ZZOpcion(2) = rstAltaCertificado!Opcion2
        ZZOpcion(3) = rstAltaCertificado!Opcion3
        ZZOpcion(4) = rstAltaCertificado!Opcion4
        ZZOpcion(5) = rstAltaCertificado!Opcion5
        ZZOpcion(6) = rstAltaCertificado!Opcion6
        ZZOpcion(7) = rstAltaCertificado!Opcion7
        ZZOpcion(8) = rstAltaCertificado!Opcion8
        ZZOpcion(9) = rstAltaCertificado!Opcion9
        ZZOpcion(10) = rstAltaCertificado!Opcion10
        rstAltaCertificado.Close
            Else
        Call Conecta_Empresa
        m$ = "No esta definido el certificado de analisis para este producto"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Exit Sub
    End If
            
    Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,ensayo4,ensayo5,Ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
            
        ZZEnsayo(1) = rstEspecifUnifica!Ensayo1
        ZZEnsayo(2) = rstEspecifUnifica!Ensayo2
        ZZEnsayo(3) = rstEspecifUnifica!Ensayo3
        ZZEnsayo(4) = rstEspecifUnifica!Ensayo4
        ZZEnsayo(5) = rstEspecifUnifica!Ensayo5
        ZZEnsayo(6) = rstEspecifUnifica!Ensayo6
        ZZEnsayo(7) = rstEspecifUnifica!Ensayo7
        ZZEnsayo(8) = rstEspecifUnifica!Ensayo8
        ZZEnsayo(9) = rstEspecifUnifica!Ensayo9
        ZZEnsayo(10) = rstEspecifUnifica!Ensayo10
                            
        ZZStd(1, 1) = rstEspecifUnifica!Valor1
        ZZStd(2, 1) = rstEspecifUnifica!valor2
        ZZStd(3, 1) = rstEspecifUnifica!Valor3
        ZZStd(4, 1) = rstEspecifUnifica!valor4
        ZZStd(5, 1) = rstEspecifUnifica!valor5
        ZZStd(6, 1) = rstEspecifUnifica!valor6
        ZZStd(7, 1) = rstEspecifUnifica!valor7
        ZZStd(8, 1) = rstEspecifUnifica!valor8
        ZZStd(9, 1) = rstEspecifUnifica!valor9
        ZZStd(10, 1) = rstEspecifUnifica!valor10
        
        ZZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
        ZZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
        ZZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
        ZZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
        ZZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
        ZZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
        ZZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
        ZZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
        ZZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
        ZZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
        rstEspecifUnifica.Close
    End If
            
    Sql1 = "Select desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Valor1Ing,Valor2Ing,Valor3Ing,Valor4Ing,Valor5Ing,Valor6Ing,Valor7Ing,Valor8Ing,Valor9Ing,Valor10Ing,Valor11Ing,Valor22Ing,Valor33Ing,Valor44Ing,Valor55Ing,Valor66Ing,Valor77Ing,Valor88Ing,Valor99Ing,Valor1010Ing,Version"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        
        
        ZZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
        ZZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
        ZZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
        ZZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
        ZZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
        ZZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
        ZZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
        ZZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
        ZZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
        ZZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                
        ZZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
        ZZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
        ZZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
        ZZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
        ZZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
        ZZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
        ZZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
        ZZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
        ZZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
        ZZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                            
        rstEspecifUnifica.Close
    End If
    
    For ZZCicla = 1 To 10
    
        spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(ZZCicla) + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZZDescri(ZZCicla) = rstEnsayo!Descripcion
            ZZDescriII(ZZCicla) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            If ZZProcesoImpre = 1 Then
                ZZOpcion(ZZCicla) = 1
            End If
            rstEnsayo.Close
        End If
        
    Next ZZCicla
    
    Call Conecta_Empresa
    
    spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
        rstTerminado.Close
    End If
                
    ZCliente = UCase(ZCliente)
    ZArticulo = UCase(ZArticulo)
    ZClave = ZCliente + ZArticulo

            
        
    LugarMetodo = 0
    
    ZSql = "DELETE Certificado"
    spCertificado = ZSql
    Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
            
    For CiclaMetodo = 1 To 10
            
        If ZZOpcion(CiclaMetodo) = 1 Then
            
            LugarMetodo = LugarMetodo + 1
                
            ZOrden = ""
            ZClave1 = ZLote
            Call Ceros(ZClave1, 6)
            ZClave2 = Str$(LugarMetodo)
            Call Ceros(ZClave2, 2)
            ZClave = ZClave1 + ZClave2
            ZMetodo = ZZEnsayo(CiclaMetodo)
            
            If Val(ZZStd(CiclaMetodo, 3)) <> 0 Or Val(ZZStd(CiclaMetodo, 4)) <> 0 Then
                ZValorNormalI = Trim(ZZStd(CiclaMetodo, 3)) + " - " + Trim(ZZStd(CiclaMetodo, 4)) + " " + Trim(ZZDescriII(CiclaMetodo)) + " " + Left$(ZZStd(CiclaMetodo, 1), 50)
                ZValorNormalII = Left$(ZZStd(CiclaMetodo, 2), 50)
                    Else
                ZValorNormalI = Left$(ZZStd(CiclaMetodo, 1), 50)
                ZValorNormalII = Left$(ZZStd(CiclaMetodo, 2), 50)
            End If
            
            ZValorNormalI = Trim(ZValorNormalI)
            ZCanti = 80 - Len(ZValorNormalI)
            ZCanti = Int(ZCanti / 2)
            ZValorNormalI = Space$(ZCanti) + ZValorNormalI
            
            ZValorNormalII = Trim(ZValorNormalII)
            ZCanti = 80 - Len(ZValorNormalII)
            ZCanti = Int(ZCanti / 2)
            ZValorNormalII = Space$(ZCanti) + ZValorNormalII
            
            ZValorPartidaII = ""
            ZObservacionesI = ""
            ZObservacionesII = ""
            ZObservacionesIII = "Version " + ZVersion
            ZObservacionesIV = ""
            ZObservacionesV = ""
            ZObservacionesVI = ""
            If Val(WEmpresa) = 1 Then
                ZEmpresa = "Surfactan S.A."
                    Else
                ZEmpresa = "Pellital S.A."
            End If
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZFechaII = ""
            
            ZExamen = ZZDescri(CiclaMetodo)
            ZExamenII = ""
            
            ZHasta = Len(Trim(ZExamen))
            If ZHasta > 25 Then
                For Cicla = ZHasta To 1 Step -1
                    If Mid(ZExamen, Cicla, 1) = Space(1) Then
                        ZExamenII = Mid(ZExamen, Cicla + 1, 25)
                        ZExamen = Mid(ZExamen, 1, Cicla)
                        Exit For
                    End If
                Next Cicla
            End If
                    
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Certificado ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Partida ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FechaII ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Examen ,"
            ZSql = ZSql + "ExamenII ,"
            ZSql = ZSql + "ValorPartidaI ,"
            ZSql = ZSql + "ValorPartidaII ,"
            ZSql = ZSql + "ValorNormalI ,"
            ZSql = ZSql + "ValorNormalII ,"
            ZSql = ZSql + "Observaciones1 ,"
            ZSql = ZSql + "Observaciones2 ,"
            ZSql = ZSql + "Observaciones3 ,"
            ZSql = ZSql + "Observaciones4 ,"
            ZSql = ZSql + "Observaciones5 ,"
            ZSql = ZSql + "Observaciones6 ,"
            ZSql = ZSql + "Metodo ,"
            ZSql = ZSql + "Empresa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + ZLote + "',"
            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
            ZSql = ZSql + "'" + ZRazon + "',"
            ZSql = ZSql + "'" + ZOrden + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZDesArticulo + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZFechaII + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZExamen + "',"
            ZSql = ZSql + "'" + ZExamenII + "',"
            ZSql = ZSql + "'" + ZValorPartidaI + "',"
            ZSql = ZSql + "'" + ZValorPartidaII + "',"
            ZSql = ZSql + "'" + ZValorNormalI + "',"
            ZSql = ZSql + "'" + ZValorNormalII + "',"
            ZSql = ZSql + "'" + ZObservacionesI + "',"
            ZSql = ZSql + "'" + ZObservacionesII + "',"
            ZSql = ZSql + "'" + ZObservacionesIII + "',"
            ZSql = ZSql + "'" + ZObservacionesIV + "',"
            ZSql = ZSql + "'" + ZObservacionesV + "',"
            ZSql = ZSql + "'" + ZObservacionesVI + "',"
            ZSql = ZSql + "'" + ZMetodo + "',"
            ZSql = ZSql + "'" + ZEmpresa + "')"

            spCertificado = ZSql
            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                    
        End If
                        
    Next CiclaMetodo
            

        
    Do
        
        If LugarMetodo = 10 Then
            Exit Do
        End If
            
        LugarMetodo = LugarMetodo + 1
                
        ZOrden = ""
        ZClave1 = ZLote
        Call Ceros(ZClave1, 6)
        ZClave2 = Str$(LugarMetodo)
        Call Ceros(ZClave2, 2)
        ZClave = ZClave1 + ZClave2
        ZMetodo = ""
        ZExamen = ""
        ZValorNormalI = ""
        ZValorNormalII = ""
        ZValorPartidaI = ""
        ZValorPartidaII = ""
        ZObservacionesI = ""
        ZObservacionesII = ""
        ZObservacionesIII = "Version " + ZVersion
        ZObservacionesIV = ""
        ZObservacionesV = ""
        ZObservacionesVI = ""
        If Val(WEmpresa) = 1 Then
            ZEmpresa = "Surfactan S.A."
                Else
            ZEmpresa = "Pellital S.A."
        End If
        ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZFechaII = WFechaElaboracion
        ZExamenII = ""
                    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Certificado ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Examen ,"
        ZSql = ZSql + "ValorPartidaI ,"
        ZSql = ZSql + "ValorPartidaII ,"
        ZSql = ZSql + "ValorNormalI ,"
        ZSql = ZSql + "ValorNormalII ,"
        ZSql = ZSql + "Observaciones1 ,"
        ZSql = ZSql + "Observaciones2 ,"
        ZSql = ZSql + "Observaciones3 ,"
        ZSql = ZSql + "Observaciones4 ,"
        ZSql = ZSql + "Observaciones5 ,"
        ZSql = ZSql + "Observaciones6 ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Empresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZClave + "',"
        ZSql = ZSql + "'" + ZLote + "',"
        ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
        ZSql = ZSql + "'" + ZRazon + "',"
        ZSql = ZSql + "'" + ZOrden + "',"
        ZSql = ZSql + "'" + ZArticulo + "',"
        ZSql = ZSql + "'" + ZDesArticulo + "',"
        ZSql = ZSql + "'" + ZFecha + "',"
        ZSql = ZSql + "'" + ZCantidad + "',"
        ZSql = ZSql + "'" + ZExamen + "',"
        ZSql = ZSql + "'" + ZValorPartidaI + "',"
        ZSql = ZSql + "'" + ZValorPartidaII + "',"
        ZSql = ZSql + "'" + ZValorNormalI + "',"
        ZSql = ZSql + "'" + ZValorNormalII + "',"
        ZSql = ZSql + "'" + ZObservacionesI + "',"
        ZSql = ZSql + "'" + ZObservacionesII + "',"
        ZSql = ZSql + "'" + ZObservacionesIII + "',"
        ZSql = ZSql + "'" + ZObservacionesIV + "',"
        ZSql = ZSql + "'" + ZObservacionesV + "',"
        ZSql = ZSql + "'" + ZObservacionesVI + "',"
        ZSql = ZSql + "'" + ZMetodo + "',"
        ZSql = ZSql + "'" + ZEmpresa + "')"

        spCertificado = ZSql
        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
            
    Loop
            
    Lista.WindowTitle = "Certificado de Analisis"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height

    Lista.Destination = 1
    Rem Listado.Destination = 0
            
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        Lista.ReportFileName = "CertificadoNuevo.rpt"
            Else
        Lista.ReportFileName = "CertificadonuevoPelli.rpt"
    End If
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Lista.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                    + "From " _
                    + DSQ + ".dbo.Certificado Certificado " _
                    + "Where " _
                    + "Certificado.Partida >= 0 AND " _
                    + "Certificado.Partida <= 999999"

    Lista.Connect = Connect()
    
    Lista.Destination = 1
    Lista.Destination = 0
    
    Lista.Action = 1

End Sub



