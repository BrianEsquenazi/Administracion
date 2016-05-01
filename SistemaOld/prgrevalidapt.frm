VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRevalidaPt 
   BackColor       =   &H00C0C000&
   Caption         =   "Revalida de Fecha de Vencimiento de Producto Terminado"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11880
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3480
      TabIndex        =   89
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   91
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   90
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label10 
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
         TabIndex        =   92
         Top             =   240
         Width           =   2895
      End
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
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   88
      Text            =   " "
      Top             =   1320
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   87
      Text            =   " "
      Top             =   1800
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   86
      Text            =   " "
      Top             =   2280
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   85
      Text            =   " "
      Top             =   2760
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   84
      Text            =   " "
      Top             =   3240
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   83
      Text            =   " "
      Top             =   3720
      Width           =   2055
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
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   82
      Text            =   " "
      Top             =   4200
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   81
      Text            =   " "
      Top             =   4680
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   80
      Text            =   " "
      Top             =   5160
      Width           =   2055
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
      Left            =   8895
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   79
      Text            =   " "
      Top             =   5640
      Width           =   2055
   End
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   78
      Top             =   5640
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   77
      Top             =   5160
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   76
      Top             =   4680
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   75
      Top             =   4200
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   74
      Top             =   3720
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   73
      Top             =   3240
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   72
      Top             =   2760
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   71
      Top             =   2280
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   70
      Top             =   1800
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
      Left            =   10995
      MaxLength       =   8
      TabIndex        =   69
      Top             =   1320
      Width           =   800
   End
   Begin VB.TextBox MesesRevalida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Rechazo 
      Caption         =   "Rechazo"
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
      Left            =   4080
      TabIndex        =   65
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Revalida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   50
      Text            =   " "
      Top             =   120
      Width           =   735
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
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Cancela 
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
      Left            =   6600
      TabIndex        =   15
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Responsable 
      BeginProperty Font 
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
      TabIndex        =   3
      Text            =   " "
      Top             =   6960
      Width           =   3255
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   6600
      Width           =   5295
   End
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   1
      Text            =   " "
      Top             =   6240
      Width           =   5295
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
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   10
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   120
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
   Begin MSMask.MaskEdBox Vto 
      Height          =   285
      Left            =   10320
      TabIndex        =   63
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   12640511
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
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   2160
      TabIndex        =   66
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSMask.MaskEdBox FechaRevalida 
      Height          =   285
      Left            =   5520
      TabIndex        =   67
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   12640511
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
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4680
      TabIndex        =   68
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9000
      TabIndex        =   64
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label ValorOri10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   62
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label ValorOri9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   61
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label ValorOri8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   60
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label ValorOri7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   59
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label ValorOri6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   58
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label ValorOri5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   57
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label ValorOri4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   56
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label ValorOri3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   55
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label ValorOri2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   54
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label ValorOri1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   53
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valores Analisis Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   51
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblensayo 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   49
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDescri 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   960
      TabIndex        =   48
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblresultado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   47
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   46
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label descri2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   45
      Top             =   1800
      Width           =   2460
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   44
      Top             =   2280
      Width           =   2460
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   43
      Top             =   2760
      Width           =   2460
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   42
      Top             =   3240
      Width           =   2460
   End
   Begin VB.Label Descri6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   41
      Top             =   3720
      Width           =   2460
   End
   Begin VB.Label Descri7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   40
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label Descri8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   39
      Top             =   4680
      Width           =   2460
   End
   Begin VB.Label Descri9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   38
      Top             =   5160
      Width           =   2460
   End
   Begin VB.Label Descri10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   37
      Top             =   5640
      Width           =   2460
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   8880
      TabIndex        =   36
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   35
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   34
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   33
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   32
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   31
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   30
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   29
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   28
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   27
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   5640
      Width           =   2655
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
      TabIndex        =   25
      Top             =   1320
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   2280
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
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
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
      TabIndex        =   21
      Top             =   3240
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
      TabIndex        =   20
      Top             =   3720
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
      TabIndex        =   19
      Top             =   4200
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
      TabIndex        =   18
      Top             =   4680
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
      TabIndex        =   17
      Top             =   5160
      Width           =   615
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
      TabIndex        =   16
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Meses de Vida Util"
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
      Left            =   8280
      TabIndex        =   14
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   12
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado"
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
      TabIndex        =   11
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label DesTerminado 
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
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo de P.T."
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
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lote"
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
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgRevalidaPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZEnsayo(10) As Integer
Dim ZEnsayoActual(10) As Integer
Dim ZValor(10, 2) As String
Dim WRevalida As Integer
Dim ZMes As String
Dim ZAno As String
Dim WGraba As String
Dim ZZEnsayo(10) As String
Dim ZZDesde(10) As String
Dim ZZHasta(10) As String
Dim ZZUnidad(10) As String
Dim ZZValorNumero(10) As String
Dim Empe(12, 10) As String
Dim WProceso As Integer

Private Sub Cancela_click()
    PrgRevalidaPt.Hide
    Unload Me
    If Val(ZProgramaOrigen) = 0 Then
        PrgPruter.Show
            Else
        PrgVerificaLoteArti.Show
    End If
End Sub

Private Sub Form_Activate()
    If Val(WEmpresaRevalida) <> 0 Then
        XEmpresa = WEmpresaRevalida
        Call Conecta_Empresa
    End If
End Sub

Private Sub Form_Load()

    If Val(WEmpresaRevalida) <> 0 Then
        XEmpresa = WEmpresaRevalida
        Call Conecta_Empresa
    End If
    
    
    Lote.Text = ZLoteRevalida
    Fecha.Text = ZFechaHoja
    FechaRevalida.Text = ZFechaRevalida
    Terminado.Text = ZArticuloRevalida
    DesTerminado.Caption = ZDesArticuloRevalida
    WGraba = ""
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PrueTer"
    ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
    
        ZZValor = IIf(IsNull(rstPrueter!ValorOriginal1), "", rstPrueter!ValorOriginal1)
        ZZValorII = IIf(IsNull(rstPrueter!ValorNumeroOriginal1), "", rstPrueter!ValorNumeroOriginal1)
        
        If Trim(ZZValor) <> "" Or Trim(ZZValorII) <> "" Then
        
            ZValor(1, 1) = IIf(IsNull(rstPrueter!ValorOriginal1), "", rstPrueter!ValorOriginal1)
            ZValor(2, 1) = IIf(IsNull(rstPrueter!ValorOriginal2), "", rstPrueter!ValorOriginal2)
            ZValor(3, 1) = IIf(IsNull(rstPrueter!ValorOriginal3), "", rstPrueter!ValorOriginal3)
            ZValor(4, 1) = IIf(IsNull(rstPrueter!ValorOriginal4), "", rstPrueter!ValorOriginal4)
            ZValor(5, 1) = IIf(IsNull(rstPrueter!ValorOriginal5), "", rstPrueter!ValorOriginal5)
            ZValor(6, 1) = IIf(IsNull(rstPrueter!ValorOriginal6), "", rstPrueter!ValorOriginal6)
            ZValor(7, 1) = IIf(IsNull(rstPrueter!ValorOriginal7), "", rstPrueter!ValorOriginal7)
            ZValor(8, 1) = IIf(IsNull(rstPrueter!ValorOriginal8), "", rstPrueter!ValorOriginal8)
            ZValor(9, 1) = IIf(IsNull(rstPrueter!ValorOriginal9), "", rstPrueter!ValorOriginal9)
            ZValor(10, 1) = IIf(IsNull(rstPrueter!ValorOriginal10), "", rstPrueter!ValorOriginal10)
            
            ZValor(1, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal1), "", rstPrueter!ValorNumeroOriginal1)
            ZValor(2, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal2), "", rstPrueter!ValorNumeroOriginal2)
            ZValor(3, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal3), "", rstPrueter!ValorNumeroOriginal3)
            ZValor(4, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal4), "", rstPrueter!ValorNumeroOriginal4)
            ZValor(5, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal5), "", rstPrueter!ValorNumeroOriginal5)
            ZValor(6, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal6), "", rstPrueter!ValorNumeroOriginal6)
            ZValor(7, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal7), "", rstPrueter!ValorNumeroOriginal7)
            ZValor(8, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal8), "", rstPrueter!ValorNumeroOriginal8)
            ZValor(9, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal9), "", rstPrueter!ValorNumeroOriginal9)
            ZValor(10, 2) = IIf(IsNull(rstPrueter!ValorNumeroOriginal10), "", rstPrueter!ValorNumeroOriginal10)
        
                Else
                
            ZValor(1, 1) = IIf(IsNull(rstPrueter!Valor1), "", rstPrueter!Valor1)
            ZValor(2, 1) = IIf(IsNull(rstPrueter!Valor2), "", rstPrueter!Valor2)
            ZValor(3, 1) = IIf(IsNull(rstPrueter!Valor3), "", rstPrueter!Valor3)
            ZValor(4, 1) = IIf(IsNull(rstPrueter!Valor4), "", rstPrueter!Valor4)
            ZValor(5, 1) = IIf(IsNull(rstPrueter!Valor5), "", rstPrueter!Valor5)
            ZValor(6, 1) = IIf(IsNull(rstPrueter!Valor6), "", rstPrueter!Valor6)
            ZValor(7, 1) = IIf(IsNull(rstPrueter!Valor7), "", rstPrueter!Valor7)
            ZValor(8, 1) = IIf(IsNull(rstPrueter!Valor8), "", rstPrueter!Valor8)
            ZValor(9, 1) = IIf(IsNull(rstPrueter!Valor9), "", rstPrueter!Valor9)
            ZValor(10, 1) = IIf(IsNull(rstPrueter!Valor10), "", rstPrueter!Valor10)
            
            ZValor(1, 2) = IIf(IsNull(rstPrueter!ValorNumero1), "", rstPrueter!ValorNumero1)
            ZValor(2, 2) = IIf(IsNull(rstPrueter!ValorNumero2), "", rstPrueter!ValorNumero2)
            ZValor(3, 2) = IIf(IsNull(rstPrueter!ValorNumero3), "", rstPrueter!ValorNumero3)
            ZValor(4, 2) = IIf(IsNull(rstPrueter!ValorNumero4), "", rstPrueter!ValorNumero4)
            ZValor(5, 2) = IIf(IsNull(rstPrueter!ValorNumero5), "", rstPrueter!ValorNumero5)
            ZValor(6, 2) = IIf(IsNull(rstPrueter!ValorNumero6), "", rstPrueter!ValorNumero6)
            ZValor(7, 2) = IIf(IsNull(rstPrueter!ValorNumero7), "", rstPrueter!ValorNumero7)
            ZValor(8, 2) = IIf(IsNull(rstPrueter!ValorNumero8), "", rstPrueter!ValorNumero8)
            ZValor(9, 2) = IIf(IsNull(rstPrueter!ValorNumero9), "", rstPrueter!ValorNumero9)
            ZValor(10, 2) = IIf(IsNull(rstPrueter!ValorNumero10), "", rstPrueter!ValorNumero10)
                
        End If
        
        ZValor(1, 1) = Trim(ZValor(1, 1))
        ZValor(2, 1) = Trim(ZValor(2, 1))
        ZValor(3, 1) = Trim(ZValor(3, 1))
        ZValor(4, 1) = Trim(ZValor(4, 1))
        ZValor(5, 1) = Trim(ZValor(5, 1))
        ZValor(6, 1) = Trim(ZValor(6, 1))
        ZValor(7, 1) = Trim(ZValor(7, 1))
        ZValor(8, 1) = Trim(ZValor(8, 1))
        ZValor(9, 1) = Trim(ZValor(9, 1))
        ZValor(10, 1) = Trim(ZValor(10, 1))
         
        ZValor(1, 2) = Trim(ZValor(1, 2))
        ZValor(2, 2) = Trim(ZValor(2, 2))
        ZValor(3, 2) = Trim(ZValor(3, 2))
        ZValor(4, 2) = Trim(ZValor(4, 2))
        ZValor(5, 2) = Trim(ZValor(5, 2))
        ZValor(6, 2) = Trim(ZValor(6, 2))
        ZValor(7, 2) = Trim(ZValor(7, 2))
        ZValor(8, 2) = Trim(ZValor(8, 2))
        ZValor(9, 2) = Trim(ZValor(9, 2))
        ZValor(10, 2) = Trim(ZValor(10, 2))
        
        WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
         
        rstPrueter.Close
        
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja = " + "'" + Lote.Text + "'"
    ZSql = ZSql + " and Producto = " + "'" + Terminado.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
        Revalida.Text = Str$(WRevalida + 1)
        rstHoja.Close
    End If
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = rstTerminado!Descripcion
        MesesRevalida.Text = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
        rstTerminado.Close
            Else
        DesTerminado.Caption = ""
        MesesRevalida.Text = ""
    End If
    
    
    
    WVida = Val(MesesRevalida.Text)
    WMes = Val(Mid$(FechaRevalida.Text, 4, 2))
    WAno = Val(Right$(FechaRevalida.Text, 4))
    
    For Ciclo = 1 To WVida
        WMes = WMes + 1
        If WMes > 12 Then
            WAno = WAno + 1
            WMes = 1
        End If
    Next Ciclo
    ZMes = Str$(WMes)
    ZAno = Str$(WAno)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    ZFechaVencimiento = "01/" + ZMes + "/" + ZAno
    Vto.Text = ZFechaVencimiento
    
    
    
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
    ZSql = ZSql + "Select especifunifica.Producto, especifunifica.Ensayo1, especifunifica.Ensayo2, especifunifica.Ensayo3, especifunifica.Ensayo4, especifunifica.Ensayo5, especifunifica.Ensayo6, especifunifica.Ensayo7, especifunifica.Ensayo8, especifunifica.Ensayo9, especifunifica.Ensayo10, "
    ZSql = ZSql + "especifunifica.Valor1, especifunifica.Valor2, especifunifica.Valor3, especifunifica.Valor4, especifunifica.Valor5, especifunifica.Valor6, especifunifica.Valor7, especifunifica.Valor8, especifunifica.Valor9, especifunifica.Valor10, "
    ZSql = ZSql + "especifunifica.desde1, especifunifica.Desde2, especifunifica.Desde3, especifunifica.Desde4, especifunifica.desde5, especifunifica.Desde6, especifunifica.Desde7, especifunifica.desde8, especifunifica.Desde9, especifunifica.Desde10, "
    ZSql = ZSql + "especifunifica.Hasta1, especifunifica.Hasta2, especifunifica.Hasta3, especifunifica.Hasta4, especifunifica.hasta5, especifunifica.Hasta6, especifunifica.hasta7, especifunifica.Hasta8, especifunifica.Hasta9, especifunifica.Hasta10 "
    ZSql = ZSql + " FROM EspecifUnifica"
    ZSql = ZSql + " Where EspecifUnifica.Producto = " + "'" + Terminado.Text + "'"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        Ensayo1.Caption = rstEspecifUnifica!Ensayo1
        Ensayo2.Caption = rstEspecifUnifica!Ensayo2
        Ensayo3.Caption = rstEspecifUnifica!Ensayo3
        Ensayo4.Caption = rstEspecifUnifica!Ensayo4
        Ensayo5.Caption = rstEspecifUnifica!Ensayo5
        Ensayo6.Caption = rstEspecifUnifica!Ensayo6
        Ensayo7.Caption = rstEspecifUnifica!Ensayo7
        Ensayo8.Caption = rstEspecifUnifica!Ensayo8
        Ensayo9.Caption = rstEspecifUnifica!Ensayo9
        Ensayo10.Caption = rstEspecifUnifica!Ensayo10
        
        ZEnsayoActual(1) = rstEspecifUnifica!Ensayo1
        ZEnsayoActual(2) = rstEspecifUnifica!Ensayo2
        ZEnsayoActual(3) = rstEspecifUnifica!Ensayo3
        ZEnsayoActual(4) = rstEspecifUnifica!Ensayo4
        ZEnsayoActual(5) = rstEspecifUnifica!Ensayo5
        ZEnsayoActual(6) = rstEspecifUnifica!Ensayo6
        ZEnsayoActual(7) = rstEspecifUnifica!Ensayo7
        ZEnsayoActual(8) = rstEspecifUnifica!Ensayo8
        ZEnsayoActual(9) = rstEspecifUnifica!Ensayo9
        ZEnsayoActual(10) = rstEspecifUnifica!Ensayo10
        
        Std1.Caption = rstEspecifUnifica!Valor1
        Std2.Caption = rstEspecifUnifica!Valor2
        Std3.Caption = rstEspecifUnifica!Valor3
        Std4.Caption = rstEspecifUnifica!Valor4
        Std5.Caption = rstEspecifUnifica!Valor5
        Std6.Caption = rstEspecifUnifica!Valor6
        Std7.Caption = rstEspecifUnifica!Valor7
        Std8.Caption = rstEspecifUnifica!Valor8
        Std9.Caption = rstEspecifUnifica!Valor9
        Std10.Caption = rstEspecifUnifica!Valor10
        
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
        
        ZZDesde(1) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
        ZZDesde(2) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
        ZZDesde(3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
        ZZDesde(4) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
        ZZDesde(5) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
        ZZDesde(6) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
        ZZDesde(7) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
        ZZDesde(8) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
        ZZDesde(9) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
        ZZDesde(10) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
        
        ZZHasta(1) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
        ZZHasta(2) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
        ZZHasta(3) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
        ZZHasta(4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
        ZZHasta(5) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
        ZZHasta(6) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
        ZZHasta(7) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
        ZZHasta(8) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
        ZZHasta(9) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
        ZZHasta(10) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
        
        ZZDesde(1) = Trim(ZZDesde(1))
        ZZDesde(2) = Trim(ZZDesde(2))
        ZZDesde(3) = Trim(ZZDesde(3))
        ZZDesde(4) = Trim(ZZDesde(4))
        ZZDesde(5) = Trim(ZZDesde(5))
        ZZDesde(6) = Trim(ZZDesde(6))
        ZZDesde(7) = Trim(ZZDesde(7))
        ZZDesde(8) = Trim(ZZDesde(8))
        ZZDesde(9) = Trim(ZZDesde(9))
        ZZDesde(10) = Trim(ZZDesde(10))
       
        ZZHasta(1) = Trim(ZZHasta(1))
        ZZHasta(2) = Trim(ZZHasta(2))
        ZZHasta(3) = Trim(ZZHasta(3))
        ZZHasta(4) = Trim(ZZHasta(4))
        ZZHasta(5) = Trim(ZZHasta(5))
        ZZHasta(6) = Trim(ZZHasta(6))
        ZZHasta(7) = Trim(ZZHasta(7))
        ZZHasta(8) = Trim(ZZHasta(8))
        ZZHasta(9) = Trim(ZZHasta(9))
        ZZHasta(10) = Trim(ZZHasta(10))
            
        rstEspecifUnifica.Close
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
        Descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
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
    
    
    
    
    
    
    
    
    LlamaImprime = "N"
    
    Erase ZEnsayo
                
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecifUnificaVersion"
    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + Terminado.Text + "'"
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
                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                        LlamaImprime = "S"
                    End If
                                
                    If WDesde > WFechaord And LlamaImprime = "N" Then
                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                        LlamaImprime = "S"
                    End If
                                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnificaVersion.Close
    End If
                
    Rem If LlamaImprime = "N" Then
                
        Sql1 = "Select EspecifUnifica.Producto, EspecifUnifica.Ensayo1, EspecifUnifica.Ensayo2, EspecifUnifica.Ensayo3, EspecifUnifica.Ensayo4, EspecifUnifica.Ensayo5, EspecifUnifica.Ensayo6, EspecifUnifica.Ensayo7, EspecifUnifica.Ensayo8, EspecifUnifica.Ensayo9, EspecifUnifica.Ensayo10 "
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + Terminado.Text + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
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
            LlamaImprime = "S"
        End If
                
    Rem End If
    
    For ZCicloI = 1 To 10
    
        Entra = "N"
        For ZCicloII = 1 To 10
            If ZEnsayoActual(ZCicloII) <> 0 Then
                If ZEnsayo(ZCicloI) = ZEnsayoActual(ZCicloII) Then
                    ZEnsayoActual(ZCicloII) = 0
                    Entra = "S"
                    ZLugar = ZCicloII
                    Exit For
                End If
            End If
        Next ZCicloII
        
        If Entra = "S" Then
            Select Case ZLugar
                Case 1
                    ValorOri1.Caption = ZValor(ZCicloI, 1)
                    Valor1.Text = ZValor(ZCicloI, 1)
                    ValorNumero1.Text = ZValor(ZCicloI, 2)
                Case 2
                    Valor2.Text = ZValor(ZCicloI, 1)
                    ValorOri2.Caption = ZValor(ZCicloI, 1)
                    ValorNumero2.Text = ZValor(ZCicloI, 2)
                Case 3
                    Valor3.Text = ZValor(ZCicloI, 1)
                    ValorOri3.Caption = ZValor(ZCicloI, 1)
                    ValorNumero3.Text = ZValor(ZCicloI, 2)
                Case 4
                    Valor4.Text = ZValor(ZCicloI, 1)
                    ValorOri4.Caption = ZValor(ZCicloI, 1)
                    ValorNumero4.Text = ZValor(ZCicloI, 2)
                Case 5
                    Valor5.Text = ZValor(ZCicloI, 1)
                    ValorOri5.Caption = ZValor(ZCicloI, 1)
                    ValorNumero5.Text = ZValor(ZCicloI, 2)
                Case 6
                    Valor6.Text = ZValor(ZCicloI, 1)
                    ValorOri6.Caption = ZValor(ZCicloI, 1)
                    ValorNumero6.Text = ZValor(ZCicloI, 2)
                Case 7
                    Valor7.Text = ZValor(ZCicloI, 1)
                    ValorOri7.Caption = ZValor(ZCicloI, 1)
                    ValorNumero7.Text = ZValor(ZCicloI, 2)
                Case 8
                    Valor8.Text = ZValor(ZCicloI, 1)
                    ValorOri8.Caption = ZValor(ZCicloI, 1)
                    ValorNumero8.Text = ZValor(ZCicloI, 2)
                Case 9
                    Valor9.Text = ZValor(ZCicloI, 1)
                    ValorOri9.Caption = ZValor(ZCicloI, 1)
                    ValorNumero9.Text = ZValor(ZCicloI, 2)
                Case 10
                    Valor10.Text = ZValor(ZCicloI, 1)
                    ValorOri10.Caption = ZValor(ZCicloI, 1)
                    ValorNumero10.Text = ZValor(ZCicloI, 2)
                Case Else
            End Select
        End If
    Next ZCicloI
    
    Call Conecta_Empresa
    
End Sub

Private Sub Graba_Click()
    
    If Val(MesesRevalida.Text) <> 0 Then
        
        WVida = Val(MesesRevalida.Text)
        WMes = Val(Mid$(FechaRevalida.Text, 4, 2))
        WAno = Val(Right$(FechaRevalida.Text, 4))
        
        For Ciclo = 1 To WVida
            WMes = WMes + 1
            If WMes > 12 Then
                WAno = WAno + 1
                WMes = 1
            End If
        Next Ciclo
        ZMes = Str$(WMes)
        ZAno = Str$(WAno)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        ZFechaVencimiento = "01/" + ZMes + "/" + ZAno
        Vto.Text = ZFechaVencimiento
            
            Else
        
        m$ = "Se debe ingresar la cantidad de meses de vida util"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Exit Sub
        
    End If
    
    
    If WGraba <> "S" Then
    
        WProceso = 0
        Call Ingresa_clave

               Else

        ZZValorNumero(1) = ValorNumero1.Text
        ZZValorNumero(2) = ValorNumero2.Text
        ZZValorNumero(3) = ValorNumero3.Text
        ZZValorNumero(4) = ValorNumero4.Text
        ZZValorNumero(5) = ValorNumero5.Text
        ZZValorNumero(6) = ValorNumero6.Text
        ZZValorNumero(7) = ValorNumero7.Text
        ZZValorNumero(8) = ValorNumero8.Text
        ZZValorNumero(9) = ValorNumero9.Text
        ZZValorNumero(10) = ValorNumero10.Text
        
        For WWCiclo = 1 To 10
        
            If ZEnsayo(WWCiclo) <> 0 Then
            
                If Val(ZZDesde(WWCiclo)) <> 0 Or Val(ZZHasta(WWCiclo)) <> 0 Then
                
                    If Val(ZZDesde(WWCiclo)) <> 0 And Val(ZZHasta(WWCiclo)) <> 0 Then
                        aa = Val(ZZValorNumero(WWCiclo))
                        If Val(ZZValorNumero(WWCiclo)) < Val(ZZDesde(WWCiclo)) Or Val(ZZValorNumero(WWCiclo)) > Val(ZZHasta(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                                
                    If Val(ZZDesde(WWCiclo)) <> 0 And Val(ZZHasta(WWCiclo)) = 0 Then
                        If Val(ZZValorNumero(WWCiclo)) < Val(ZZDesde(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                            
                    If Val(ZZDesde(WWCiclo)) = 0 And Val(ZZHasta(WWCiclo)) <> 0 Then
                        If Val(ZZValorNumero(WWCiclo)) > Val(ZZHasta(WWCiclo)) Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                    End If
                    
                        Else
                        
                    If Trim(UCase(ZZValorNumero(WWCiclo))) <> "S" Then
                        m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Exit Sub
                    End If
                    
                End If
                
            End If
        
        Next WWCiclo
    
    
    
    
        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM RevalidaPt"
        spRevalidaPt = Sql1 + Sql2
        Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalidaPt.RecordCount > 0 Then
            rstRevalidaPt.MoveLast
            WCodigoMayor = IIf(IsNull(rstRevalidaPt!CodigoMayor), "0", rstRevalidaPt!CodigoMayor)
            WCodigo = Str$(WCodigoMayor + 1)
            rstRevalidaPt.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO RevalidaPt ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "MesesRevalida ,"
        ZSql = ZSql + "Codigo1 ,"
        ZSql = ZSql + "Codigo2 ,"
        ZSql = ZSql + "Codigo3 ,"
        ZSql = ZSql + "Codigo4 ,"
        ZSql = ZSql + "Codigo5 ,"
        ZSql = ZSql + "Codigo6 ,"
        ZSql = ZSql + "Codigo7 ,"
        ZSql = ZSql + "Codigo8 ,"
        ZSql = ZSql + "Codigo9 ,"
        ZSql = ZSql + "Codigo10 ,"
        ZSql = ZSql + "Std1 ,"
        ZSql = ZSql + "Std2 ,"
        ZSql = ZSql + "Std3 ,"
        ZSql = ZSql + "Std4 ,"
        ZSql = ZSql + "Std5 ,"
        ZSql = ZSql + "Std6 ,"
        ZSql = ZSql + "Std7 ,"
        ZSql = ZSql + "Std8 ,"
        ZSql = ZSql + "Std9 ,"
        ZSql = ZSql + "Std10 ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "ValorNumero1 ,"
        ZSql = ZSql + "ValorNumero2 ,"
        ZSql = ZSql + "ValorNumero3 ,"
        ZSql = ZSql + "ValorNumero4 ,"
        ZSql = ZSql + "ValorNumero5 ,"
        ZSql = ZSql + "ValorNumero6 ,"
        ZSql = ZSql + "ValorNumero7 ,"
        ZSql = ZSql + "ValorNumero8 ,"
        ZSql = ZSql + "ValorNumero9 ,"
        ZSql = ZSql + "ValorNumero10 ,"
        ZSql = ZSql + "Revalida )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + Lote.Text + "',"
        ZSql = ZSql + "'" + FechaRevalida.Text + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Resultado.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Responsable.Text + "',"
        ZSql = ZSql + "'" + Vto.Text + "',"
        ZSql = ZSql + "'" + MesesRevalida.Text + "',"
        ZSql = ZSql + "'" + Ensayo1.Caption + "',"
        ZSql = ZSql + "'" + Ensayo2.Caption + "',"
        ZSql = ZSql + "'" + Ensayo3.Caption + "',"
        ZSql = ZSql + "'" + Ensayo4.Caption + "',"
        ZSql = ZSql + "'" + Ensayo5.Caption + "',"
        ZSql = ZSql + "'" + Ensayo6.Caption + "',"
        ZSql = ZSql + "'" + Ensayo7.Caption + "',"
        ZSql = ZSql + "'" + Ensayo8.Caption + "',"
        ZSql = ZSql + "'" + Ensayo9.Caption + "',"
        ZSql = ZSql + "'" + Ensayo10.Caption + "',"
        ZSql = ZSql + "'" + Std1.Caption + "',"
        ZSql = ZSql + "'" + Std2.Caption + "',"
        ZSql = ZSql + "'" + Std3.Caption + "',"
        ZSql = ZSql + "'" + Std4.Caption + "',"
        ZSql = ZSql + "'" + Std5.Caption + "',"
        ZSql = ZSql + "'" + Std6.Caption + "',"
        ZSql = ZSql + "'" + Std7.Caption + "',"
        ZSql = ZSql + "'" + Std8.Caption + "',"
        ZSql = ZSql + "'" + Std9.Caption + "',"
        ZSql = ZSql + "'" + Std10.Caption + "',"
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + Valor2.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + Valor4.Text + "',"
        ZSql = ZSql + "'" + Valor5.Text + "',"
        ZSql = ZSql + "'" + Valor6.Text + "',"
        ZSql = ZSql + "'" + Valor7.Text + "',"
        ZSql = ZSql + "'" + Valor8.Text + "',"
        ZSql = ZSql + "'" + Valor9.Text + "',"
        ZSql = ZSql + "'" + Valor10.Text + "',"
        ZSql = ZSql + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + "'" + ValorNumero10.Text + "',"
        ZSql = ZSql + "'" + Revalida.Text + "')"
            
        spRevalidaPt = ZSql
        Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZOrdVencimiento = Right$(Vto.Text, 4) + Mid$(Vto.Text, 4, 2) + Left$(Vto.Text, 2)
        ZOrdFechaRevalida = Right$(FechaRevalida.Text, 4) + Mid$(FechaRevalida.Text, 4, 2) + Left$(FechaRevalida.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + "FechaVencimiento = " + "'" + Vto.Text + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "',"
        ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "',"
        ZSql = ZSql + "FechaRevalida = " + "'" + FechaRevalida.Text + "',"
        ZSql = ZSql + "OrdFechaRevalida = " + "'" + ZOrdFechaRevalida + "',"
        ZSql = ZSql + "MesesRevalida = " + "'" + MesesRevalida.Text + "',"
        ZSql = ZSql + "MarcaVencida = " + "'" + "" + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Lote.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        XEmpresa = WEmpresa
        
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        
        For A = 1 To 7
            
            WEmpresa = Empe(A, 1)
            txtOdbc = Empe(A, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Guia SET "
            ZSql = ZSql + " MarcaVencida = " + "'" + "" + "'"
            ZSql = ZSql + " Where Guia.Terminado = " + "'" + Terminado.Text + "'"
            ZSql = ZSql + " and Guia.Lote = " + "'" + Lote.Text + "'"
            spGuia = ZSql
            Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
        
        Next A
        
        Call Conecta_Empresa
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueter SET "
        ZSql = ZSql + "Valor1 = " + "'" + Valor1.Text + "',"
        ZSql = ZSql + "Valor2 = " + "'" + Valor2.Text + "',"
        ZSql = ZSql + "Valor3 = " + "'" + Valor3.Text + "',"
        ZSql = ZSql + "Valor4 = " + "'" + Valor4.Text + "',"
        ZSql = ZSql + "Valor5 = " + "'" + Valor5.Text + "',"
        ZSql = ZSql + "Valor6 = " + "'" + Valor6.Text + "',"
        ZSql = ZSql + "Valor7 = " + "'" + Valor7.Text + "',"
        ZSql = ZSql + "Valor8 = " + "'" + Valor8.Text + "',"
        ZSql = ZSql + "Valor9 = " + "'" + Valor9.Text + "',"
        ZSql = ZSql + "Valor10 = " + "'" + Valor10.Text + "',"
        ZSql = ZSql + "ValorNumero1 = " + "'" + ValorNumero1.Text + "',"
        ZSql = ZSql + "ValorNumero2 = " + "'" + ValorNumero2.Text + "',"
        ZSql = ZSql + "ValorNumero3 = " + "'" + ValorNumero3.Text + "',"
        ZSql = ZSql + "ValorNumero4 = " + "'" + ValorNumero4.Text + "',"
        ZSql = ZSql + "ValorNumero5 = " + "'" + ValorNumero5.Text + "',"
        ZSql = ZSql + "ValorNumero6 = " + "'" + ValorNumero6.Text + "',"
        ZSql = ZSql + "ValorNumero7 = " + "'" + ValorNumero7.Text + "',"
        ZSql = ZSql + "ValorNumero8 = " + "'" + ValorNumero8.Text + "',"
        ZSql = ZSql + "ValorNumero9 = " + "'" + ValorNumero9.Text + "',"
        ZSql = ZSql + "ValorNumero10 = " + "'" + ValorNumero10.Text + "'"
        ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
        spPrueter = ZSql
        Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
        
        PrgPruter.NroRevalida.Text = Revalida.Text
        PrgPruter.Vto.Text = Vto.Text
        
        PrgPruter.Valor1.Text = Valor1.Text
        PrgPruter.Valor2.Text = Valor2.Text
        PrgPruter.Valor3.Text = Valor3.Text
        PrgPruter.Valor4.Text = Valor4.Text
        PrgPruter.Valor5.Text = Valor5.Text
        PrgPruter.Valor6.Text = Valor6.Text
        PrgPruter.Valor7.Text = Valor7.Text
        PrgPruter.Valor8.Text = Valor8.Text
        PrgPruter.Valor9.Text = Valor9.Text
        PrgPruter.Valor10.Text = Valor10.Text
        
        PrgPruter.ValorNumero1.Text = ValorNumero1.Text
        PrgPruter.ValorNumero2.Text = ValorNumero2.Text
        PrgPruter.ValorNumero3.Text = ValorNumero3.Text
        PrgPruter.ValorNumero4.Text = ValorNumero4.Text
        PrgPruter.ValorNumero5.Text = ValorNumero5.Text
        PrgPruter.ValorNumero6.Text = ValorNumero6.Text
        PrgPruter.ValorNumero7.Text = ValorNumero7.Text
        PrgPruter.ValorNumero8.Text = ValorNumero8.Text
        PrgPruter.ValorNumero9.Text = ValorNumero9.Text
        PrgPruter.ValorNumero10.Text = ValorNumero10.Text
        
        Call Cancela_click

    End If

End Sub


Private Sub Rechazo_Click()

    If WGraba <> "S" Then
    
        WProceso = 1
        Call Ingresa_clave

               Else

        Revalida.Text = "99"
    
        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM RevalidaPt"
        spRevalidaPt = Sql1 + Sql2
        Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalidaPt.RecordCount > 0 Then
            rstRevalidaPt.MoveLast
            WCodigoMayor = IIf(IsNull(rstRevalidaPt!CodigoMayor), "0", rstRevalidaPt!CodigoMayor)
            WCodigo = Str$(WCodigoMayor + 1)
            rstRevalidaPt.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO RevalidaPt ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Codigo1 ,"
        ZSql = ZSql + "Codigo2 ,"
        ZSql = ZSql + "Codigo3 ,"
        ZSql = ZSql + "Codigo4 ,"
        ZSql = ZSql + "Codigo5 ,"
        ZSql = ZSql + "Codigo6 ,"
        ZSql = ZSql + "Codigo7 ,"
        ZSql = ZSql + "Codigo8 ,"
        ZSql = ZSql + "Codigo9 ,"
        ZSql = ZSql + "Codigo10 ,"
        ZSql = ZSql + "Std1 ,"
        ZSql = ZSql + "Std2 ,"
        ZSql = ZSql + "Std3 ,"
        ZSql = ZSql + "Std4 ,"
        ZSql = ZSql + "Std5 ,"
        ZSql = ZSql + "Std6 ,"
        ZSql = ZSql + "Std7 ,"
        ZSql = ZSql + "Std8 ,"
        ZSql = ZSql + "Std9 ,"
        ZSql = ZSql + "Std10 ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "Revalida )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + Lote.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Resultado.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Responsable.Text + "',"
        ZSql = ZSql + "'" + Vto.Text + "',"
        ZSql = ZSql + "'" + Ensayo1.Caption + "',"
        ZSql = ZSql + "'" + Ensayo2.Caption + "',"
        ZSql = ZSql + "'" + Ensayo3.Caption + "',"
        ZSql = ZSql + "'" + Ensayo4.Caption + "',"
        ZSql = ZSql + "'" + Ensayo5.Caption + "',"
        ZSql = ZSql + "'" + Ensayo6.Caption + "',"
        ZSql = ZSql + "'" + Ensayo7.Caption + "',"
        ZSql = ZSql + "'" + Ensayo8.Caption + "',"
        ZSql = ZSql + "'" + Ensayo9.Caption + "',"
        ZSql = ZSql + "'" + Ensayo10.Caption + "',"
        ZSql = ZSql + "'" + Std1.Caption + "',"
        ZSql = ZSql + "'" + Std2.Caption + "',"
        ZSql = ZSql + "'" + Std3.Caption + "',"
        ZSql = ZSql + "'" + Std4.Caption + "',"
        ZSql = ZSql + "'" + Std5.Caption + "',"
        ZSql = ZSql + "'" + Std6.Caption + "',"
        ZSql = ZSql + "'" + Std7.Caption + "',"
        ZSql = ZSql + "'" + Std8.Caption + "',"
        ZSql = ZSql + "'" + Std9.Caption + "',"
        ZSql = ZSql + "'" + Std10.Caption + "',"
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + Valor2.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + Valor4.Text + "',"
        ZSql = ZSql + "'" + Valor5.Text + "',"
        ZSql = ZSql + "'" + Valor6.Text + "',"
        ZSql = ZSql + "'" + Valor7.Text + "',"
        ZSql = ZSql + "'" + Valor8.Text + "',"
        ZSql = ZSql + "'" + Valor9.Text + "',"
        ZSql = ZSql + "'" + Valor10.Text + "',"
        ZSql = ZSql + "'" + Revalida.Text + "')"
            
        spRevalidaPt = ZSql
        Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
        
        ZVto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZOrdVencimiento = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        ZRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZOrdFechaRevalida = Right$(ZRevalida, 4) + Mid$(ZRevalida, 4, 2) + Left$(ZRevalida, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + "FechaVencimiento = " + "'" + ZVto + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "',"
        ZSql = ZSql + "FechaRevalida = " + "'" + ZRevalida + "',"
        ZSql = ZSql + "OrdFechaRevalida = " + "'" + ZOrdFechaRevalida + "',"
        ZSql = ZSql + "MarcaVencida = " + "'" + "V" + "',"
        ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Lote.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        PrgPruter.NroRevalida.Text = Revalida.Text
        
        Call Cancela_click

    End If

End Sub

Private Sub MesesRevalida_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WVida = Val(MesesRevalida.Text)
        WMes = Val(Mid$(FechaRevalida.Text, 4, 2))
        WAno = Val(Right$(FechaRevalida.Text, 4))
        
        For Ciclo = 1 To WVida
            WMes = WMes + 1
            If WMes > 12 Then
                WAno = WAno + 1
                WMes = 1
            End If
        Next Ciclo
        ZMes = Str$(WMes)
        ZAno = Str$(WAno)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        ZFechaVencimiento = "01/" + ZMes + "/" + ZAno
        Vto.Text = ZFechaVencimiento
        
        ValorNumero1.SetFocus
    End If
    If KeyAscii = 27 Then
        MesesRevalida.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            ValorNumero1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub ValorNumero1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZZDesde(1)) <> 0 Or Val(ZZHasta(1)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(1)), ".")
            ZNumeII = Len(Trim(ZZDesde(1)))
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
            
            Valor1.Text = ValorNumero1.Text + " " + ZZUnidad(1)
            
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
    
    If Val(ZZDesde(1)) <> 0 Or Val(ZZHasta(1)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZZDesde(2)) <> 0 Or Val(ZZHasta(2)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(2)), ".")
            ZNumeII = Len(Trim(ZZDesde(2)))
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
            
            Valor2.Text = ValorNumero2.Text + " " + ZZUnidad(2)
            
            ValorNumero3.SetFocus
            
                Else
                
            If ValorNumero2.Text = "S" Or ValorNumero2.Text = "N" Then
                If ValorNumero2.Text = "S" Then
                    Valor2.Text = "Cumple"
                        Else
                    Valor2.Text = "No Cumple"
                End If
                ValorNumero3.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero2.Text = ""
    End If
    
    If Val(ZZDesde(2)) <> 0 Or Val(ZZHasta(2)) <> 0 Then
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
    
        If Val(ZZDesde(3)) <> 0 Or Val(ZZHasta(3)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(3)), ".")
            ZNumeII = Len(Trim(ZZDesde(3)))
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
            
            Valor3.Text = ValorNumero3.Text + " " + ZZUnidad(3)
            
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
    
    If Val(ZZDesde(3)) <> 0 Or Val(ZZHasta(3)) <> 0 Then
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
    
        If Val(ZZDesde(4)) <> 0 Or Val(ZZHasta(4)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(4)), ".")
            ZNumeII = Len(Trim(ZZDesde(4)))
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
            
            Valor4.Text = ValorNumero4.Text + " " + ZZUnidad(4)
            
            ValorNumero5.SetFocus
            
                Else
                
            If ValorNumero4.Text = "S" Or ValorNumero4.Text = "N" Then
                If ValorNumero4.Text = "S" Then
                    Valor4.Text = "Cumple"
                        Else
                    Valor4.Text = "No Cumple"
                End If
                ValorNumero5.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero4.Text = ""
    End If
    
    If Val(ZZDesde(4)) <> 0 Or Val(ZZHasta(4)) <> 0 Then
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
    
        If Val(ZZDesde(5)) <> 0 Or Val(ZZHasta(5)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(5)), ".")
            ZNumeII = Len(Trim(ZZDesde(5)))
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
            
            Valor5.Text = ValorNumero5.Text + " " + ZZUnidad(5)
            
            ValorNumero6.SetFocus
            
                Else
                
            If ValorNumero5.Text = "S" Or ValorNumero5.Text = "N" Then
                If ValorNumero5.Text = "S" Then
                    Valor5.Text = "Cumple"
                        Else
                    Valor5.Text = "No Cumple"
                End If
                ValorNumero6.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero5.Text = ""
    End If
    
    If Val(ZZDesde(5)) <> 0 Or Val(ZZHasta(5)) <> 0 Then
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
    
        If Val(ZZDesde(6)) <> 0 Or Val(ZZHasta(6)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(6)), ".")
            ZNumeII = Len(Trim(ZZDesde(6)))
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
            
            Valor6.Text = ValorNumero6.Text + " " + ZZUnidad(6)
            
            ValorNumero7.SetFocus
            
                Else
                
            If ValorNumero6.Text = "S" Or ValorNumero6.Text = "N" Then
                If ValorNumero6.Text = "S" Then
                    Valor6.Text = "Cumple"
                        Else
                    Valor6.Text = "No Cumple"
                End If
                ValorNumero7.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero6.Text = ""
    End If
    
    If Val(ZZDesde(6)) <> 0 Or Val(ZZHasta(6)) <> 0 Then
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
    
        If Val(ZZDesde(7)) <> 0 Or Val(ZZHasta(7)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(7)), ".")
            ZNumeII = Len(Trim(ZZDesde(7)))
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
            
            Valor7.Text = ValorNumero7.Text + " " + ZZUnidad(7)
            
            ValorNumero8.SetFocus
            
                Else
                
            If ValorNumero7.Text = "S" Or ValorNumero7.Text = "N" Then
                If ValorNumero7.Text = "S" Then
                    Valor7.Text = "Cumple"
                        Else
                    Valor7.Text = "No Cumple"
                End If
                ValorNumero8.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero7.Text = ""
    End If
    
    If Val(ZZDesde(7)) <> 0 Or Val(ZZHasta(7)) <> 0 Then
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
    
        If Val(ZZDesde(8)) <> 0 Or Val(ZZHasta(8)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(8)), ".")
            ZNumeII = Len(Trim(ZZDesde(8)))
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
            
            Valor8.Text = ValorNumero8.Text + " " + ZZUnidad(8)
            
            ValorNumero9.SetFocus
            
                Else
                
            If ValorNumero8.Text = "S" Or ValorNumero8.Text = "N" Then
                If ValorNumero8.Text = "S" Then
                    Valor8.Text = "Cumple"
                        Else
                    Valor8.Text = "No Cumple"
                End If
                ValorNumero9.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero8.Text = ""
    End If
    
    If Val(ZZDesde(8)) <> 0 Or Val(ZZHasta(8)) <> 0 Then
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
    
        If Val(ZZDesde(9)) <> 0 Or Val(ZZHasta(9)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(9)), ".")
            ZNumeII = Len(Trim(ZZDesde(9)))
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
            
            Valor9.Text = ValorNumero9.Text + " " + ZZUnidad(9)
            
            ValorNumero10.SetFocus
            
                Else
                
            If ValorNumero9.Text = "S" Or ValorNumero9.Text = "N" Then
                If ValorNumero9.Text = "S" Then
                    Valor9.Text = "Cumple"
                        Else
                    Valor9.Text = "No Cumple"
                End If
                ValorNumero10.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero9.Text = ""
    End If
    
    If Val(ZZDesde(9)) <> 0 Or Val(ZZHasta(9)) <> 0 Then
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
    
        If Val(ZZDesde(10)) <> 0 Or Val(ZZHasta(10)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZZDesde(10)), ".")
            ZNumeII = Len(Trim(ZZDesde(10)))
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
            
            Valor10.Text = ValorNumero10.Text + " " + ZZUnidad(10)
            
            ValorNumero1.SetFocus
            
                Else
                
            If ValorNumero10.Text = "S" Or ValorNumero10.Text = "N" Then
                If ValorNumero10.Text = "S" Then
                    Valor10.Text = "Cumple"
                        Else
                    Valor10.Text = "No Cumple"
                End If
                ValorNumero1.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero10.Text = ""
    End If
    
    If Val(ZZDesde(10)) <> 0 Or Val(ZZHasta(10)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub Resultado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Resultado.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vto.SetFocus
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
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
        WGrabaI = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            WGrabaI = IIf(IsNull(rstOperador!GrabaI), "", rstOperador!GrabaI)
            rstOperador.Close
        End If
        
        If WGrabaI = "S" Then
            WGraba = "S"
            XClave.Visible = False
            If WProceso = 0 Then
                Call Graba_Click
                    Else
                Call Rechazo_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub





