VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRevalidaPtConsulta 
   BackColor       =   &H00C0C000&
   Caption         =   "Revalida de Fecha de Vencimiento de Producto Terminado"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11880
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
      TabIndex        =   59
      Text            =   " "
      Top             =   120
      Width           =   735
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
      Height          =   495
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PRGRevalidaPTConsulta.frx":0000
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Valor2 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "PRGRevalidaPTConsulta.frx":0002
      Top             =   1800
      Width           =   2895
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
      Height          =   495
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "PRGRevalidaPTConsulta.frx":0004
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Valor4 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "PRGRevalidaPTConsulta.frx":0006
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Valor5 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "PRGRevalidaPTConsulta.frx":0008
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Valor6 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "PRGRevalidaPTConsulta.frx":000A
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Valor7 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "PRGRevalidaPTConsulta.frx":000C
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Valor8 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "PRGRevalidaPTConsulta.frx":000E
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Valor9 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "PRGRevalidaPTConsulta.frx":0010
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Valor10 
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
      Left            =   8880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "PRGRevalidaPTConsulta.frx":0012
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton Cancela 
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
      Height          =   495
      Left            =   4080
      TabIndex        =   24
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   19
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   14
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
      TabIndex        =   72
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
      TabIndex        =   74
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
      TabIndex        =   75
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
      TabIndex        =   76
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
      TabIndex        =   73
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label ValorOri10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   71
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label ValorOri9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   70
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label ValorOri8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   69
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label ValorOri7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   68
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label ValorOri6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   67
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label ValorOri5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   66
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label ValorOri4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   65
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label ValorOri3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   64
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label ValorOri2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   63
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label ValorOri1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   6120
      TabIndex        =   62
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
      TabIndex        =   61
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
      TabIndex        =   60
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   56
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   55
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label descri2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   54
      Top             =   1800
      Width           =   2460
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   53
      Top             =   2280
      Width           =   2460
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   52
      Top             =   2760
      Width           =   2460
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   51
      Top             =   3240
      Width           =   2460
   End
   Begin VB.Label Descri6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   50
      Top             =   3720
      Width           =   2460
   End
   Begin VB.Label Descri7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   49
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label Descri8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   48
      Top             =   4680
      Width           =   2460
   End
   Begin VB.Label Descri9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   47
      Top             =   5160
      Width           =   2460
   End
   Begin VB.Label Descri10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   46
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
      TabIndex        =   45
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   44
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   43
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   42
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   41
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   40
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   39
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   38
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   37
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   36
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   3360
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgRevalidaPtConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZEnsayo(10) As Integer
Dim ZEnsayoActual(10) As Integer
Dim ZValor(10) As String
Dim WRevalida As Integer
Dim ZMes As String
Dim ZAno As String

Private Sub Cancela_click()
    PrgRevalidaPtConsulta.Hide
    Unload Me
    PrgPruter.Show
End Sub

Private Sub Form_Load()

    Lote.Text = ZLoteRevalida
    Fecha.Text = ZFechaHoja
    FechaRevalida.Text = ZFechaRevalida
    Terminado.Text = ZArticuloRevalida
    DesTerminado.Caption = ZDesArticuloRevalida
    Revalida.Text = ZNroRevalida
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PrueTer"
    ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
    spPrueter = ZSql
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
        
        ZZValor1 = IIf(IsNull(rstPrueter!ValorOriginal1), "", rstPrueter!ValorOriginal1)
        ZZValor2 = IIf(IsNull(rstPrueter!ValorOriginal2), "", rstPrueter!ValorOriginal2)
        ZZValor3 = IIf(IsNull(rstPrueter!ValorOriginal3), "", rstPrueter!ValorOriginal3)
        ZZValor4 = IIf(IsNull(rstPrueter!ValorOriginal4), "", rstPrueter!ValorOriginal4)
        ZZValor5 = IIf(IsNull(rstPrueter!ValorOriginal5), "", rstPrueter!ValorOriginal5)
        ZZValor6 = IIf(IsNull(rstPrueter!ValorOriginal6), "", rstPrueter!ValorOriginal6)
        ZZValor7 = IIf(IsNull(rstPrueter!ValorOriginal7), "", rstPrueter!ValorOriginal7)
        ZZValor8 = IIf(IsNull(rstPrueter!ValorOriginal8), "", rstPrueter!ValorOriginal8)
        ZZValor9 = IIf(IsNull(rstPrueter!ValorOriginal9), "", rstPrueter!ValorOriginal9)
        ZZValor10 = IIf(IsNull(rstPrueter!ValorOriginal10), "", rstPrueter!ValorOriginal10)
        
        If Trim(ZZValor1) <> "" Then
            ZValor(1) = ZZValor1
                Else
            ZValor(1) = rstPrueter!Valor1
        End If
        If Trim(ZZValor2) <> "" Then
            ZValor(2) = ZZValor2
                Else
            ZValor(2) = rstPrueter!valor2
        End If
        If Trim(ZZValor3) <> "" Then
            ZValor(3) = ZZValor3
                Else
            ZValor(3) = rstPrueter!Valor3
        End If
        If Trim(ZZValor4) <> "" Then
            ZValor(4) = ZZValor4
                Else
            ZValor(4) = rstPrueter!valor4
        End If
        If Trim(ZZValor5) <> "" Then
            ZValor(5) = ZZValor5
                Else
            ZValor(5) = rstPrueter!valor5
        End If
        If Trim(ZZValor6) <> "" Then
            ZValor(6) = ZZValor6
                Else
            ZValor(6) = rstPrueter!valor6
        End If
        If Trim(ZZValor7) <> "" Then
            ZValor(7) = ZZValor7
                Else
            ZValor(7) = rstPrueter!valor7
        End If
        If Trim(ZZValor8) <> "" Then
            ZValor(8) = ZZValor8
                Else
            ZValor(8) = rstPrueter!valor8
        End If
        If Trim(ZZValor9) <> "" Then
            ZValor(9) = ZZValor9
                Else
            ZValor(9) = rstPrueter!valor9
        End If
        If Trim(ZZValor10) <> "" Then
            ZValor(10) = ZZValor10
                Else
            ZValor(10) = rstPrueter!valor10
        End If
        
        rstPrueter.Close
        
    End If
    
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = rstTerminado!Descripcion
        rstTerminado.Close
            Else
        MesesRevalida.Text = ""
    End If
    
  
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
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + Terminado.Text + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
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
        Std2.Caption = rstEspecifUnifica!valor2
        Std3.Caption = rstEspecifUnifica!Valor3
        Std4.Caption = rstEspecifUnifica!valor4
        Std5.Caption = rstEspecifUnifica!valor5
        Std6.Caption = rstEspecifUnifica!valor6
        Std7.Caption = rstEspecifUnifica!valor7
        Std8.Caption = rstEspecifUnifica!valor8
        Std9.Caption = rstEspecifUnifica!valor9
        Std10.Caption = rstEspecifUnifica!valor10
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
                
    If LlamaImprime = "N" Then
                
        Sql1 = "Select *"
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
                
    End If
    
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
                    ValorOri1.Caption = ZValor(ZCicloI)
                Case 2
                    ValorOri2.Caption = ZValor(ZCicloI)
                Case 3
                    ValorOri3.Caption = ZValor(ZCicloI)
                Case 4
                    ValorOri4.Caption = ZValor(ZCicloI)
                Case 5
                    ValorOri5.Caption = ZValor(ZCicloI)
                Case 6
                    ValorOri6.Caption = ZValor(ZCicloI)
                Case 7
                    ValorOri7.Caption = ZValor(ZCicloI)
                Case 8
                    ValorOri8.Caption = ZValor(ZCicloI)
                Case 9
                    ValorOri9.Caption = ZValor(ZCicloI)
                Case 10
                    ValorOri10.Caption = ZValor(ZCicloI)
                Case Else
            End Select
        End If
    Next ZCicloI
    
    Call Conecta_Empresa
    
    Sql1 = "Select *"
    Sql2 = " FROM RevalidaPt"
    Sql3 = " Where RevalidaPt.Lote = " + "'" + Lote.Text + "'"
    Sql4 = " and RevalidaPt.Producto = " + "'" + Terminado.Text + "'"
    Sql5 = " and RevalidaPt.Revalida = " + "'" + Revalida.Text + "'"
    spRevalidaPt = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
    If rstRevalidaPt.RecordCount > 0 Then
        Vto.Text = rstRevalidaPt!Vencimiento
        MesesRevalida.Text = IIf(IsNull(rstRevalidaPt!MesesRevalida), "0", rstRevalidaPt!MesesRevalida)
        Valor1.Text = rstRevalidaPt!Valor1
        valor2.Text = rstRevalidaPt!valor2
        Valor3.Text = rstRevalidaPt!Valor3
        valor4.Text = rstRevalidaPt!valor4
        valor5.Text = rstRevalidaPt!valor5
        valor6.Text = rstRevalidaPt!valor6
        valor7.Text = rstRevalidaPt!valor7
        valor8.Text = rstRevalidaPt!valor8
        valor9.Text = rstRevalidaPt!valor9
        valor10.Text = rstRevalidaPt!valor10
        Resultado.Text = rstRevalidaPt!Resultado
        Observaciones.Text = rstRevalidaPt!Observaciones
        Responsable.Text = rstRevalidaPt!Responsable
        rstRevalidaPt.Close
    End If
    
    
    
    
End Sub

Private Sub Graba_Click()

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
    ZSql = ZSql + "'" + valor2.Text + "',"
    ZSql = ZSql + "'" + Valor3.Text + "',"
    ZSql = ZSql + "'" + valor4.Text + "',"
    ZSql = ZSql + "'" + valor5.Text + "',"
    ZSql = ZSql + "'" + valor6.Text + "',"
    ZSql = ZSql + "'" + valor7.Text + "',"
    ZSql = ZSql + "'" + valor8.Text + "',"
    ZSql = ZSql + "'" + valor9.Text + "',"
    ZSql = ZSql + "'" + valor10.Text + "',"
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
    ZSql = ZSql + "MesesRevalida = " + "'" + MesesRevalida.Text + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + Lote.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    PrgPruter.NroRevalida.Text = Revalida.Text
    PrgPruter.Vto.Text = Vto.Text
    
    Call Cancela_click

End Sub


Private Sub Rechazo_Click()

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
    ZSql = ZSql + "'" + valor2.Text + "',"
    ZSql = ZSql + "'" + Valor3.Text + "',"
    ZSql = ZSql + "'" + valor4.Text + "',"
    ZSql = ZSql + "'" + valor5.Text + "',"
    ZSql = ZSql + "'" + valor6.Text + "',"
    ZSql = ZSql + "'" + valor7.Text + "',"
    ZSql = ZSql + "'" + valor8.Text + "',"
    ZSql = ZSql + "'" + valor9.Text + "',"
    ZSql = ZSql + "'" + valor10.Text + "',"
    ZSql = ZSql + "'" + Revalida.Text + "')"
        
    spRevalidaPt = ZSql
    Set rstRevalidaPt = db.OpenRecordset(spRevalidaPt, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + Lote.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    PrgPruter.NroRevalida.Text = Revalida.Text
    
    Call Cancela_click

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
        
        Valor1.SetFocus
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
            Valor1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor1.SelStart = 0
        valor2.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor1.Text = ""
    End If
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor2.SelStart = 0
        Valor3.SetFocus
    End If
    If KeyAscii = 27 Then
        valor2.Text = ""
    End If
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor3.SelStart = 0
        valor4.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor3.Text = ""
    End If
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor4.SelStart = 0
        valor5.SetFocus
    End If
    If KeyAscii = 27 Then
        valor4.Text = ""
    End If
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor5.SelStart = 0
        valor6.SetFocus
    End If
    If KeyAscii = 27 Then
        valor5.Text = ""
    End If
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor6.SelStart = 0
        valor7.SetFocus
    End If
    If KeyAscii = 27 Then
        valor6.Text = ""
    End If
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor7.SelStart = 0
        valor8.SetFocus
    End If
    If KeyAscii = 27 Then
        valor7.Text = ""
    End If
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor8.SelStart = 0
        valor9.SetFocus
    End If
    If KeyAscii = 27 Then
        valor8.Text = ""
    End If
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor9.SelStart = 0
        valor10.SetFocus
    End If
    If KeyAscii = 27 Then
        valor9.Text = ""
    End If
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor10.SelStart = 0
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        valor10.Text = ""
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

