VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaSacVerifica 
   Caption         =   "Carga de SAC - Verificacion"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11775
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3120
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
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
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "CargaSacVerifica.frx":0000
      Left            =   2160
      List            =   "CargaSacVerifica.frx":0007
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Responsable15 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   87
      Text            =   " "
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Responsable16 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   86
      Text            =   " "
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Responsable14 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   85
      Text            =   " "
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Responsable13 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   84
      Text            =   " "
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Responsable12 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   83
      Text            =   " "
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Responsable11 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   82
      Text            =   " "
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox Estado11 
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
      Left            =   5640
      TabIndex        =   81
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Estado12 
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
      Left            =   5640
      TabIndex        =   80
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox Estado13 
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
      Left            =   5640
      TabIndex        =   79
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox Estado14 
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
      Left            =   5640
      TabIndex        =   78
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox Estado15 
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
      Left            =   5640
      TabIndex        =   77
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox Estado16 
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
      Left            =   5640
      TabIndex        =   76
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Referencia 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   68
      Text            =   " "
      Top             =   1200
      Width           =   10455
   End
   Begin VB.TextBox Numero 
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   57
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Centro 
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
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   56
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Estado 
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
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox Origen 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Ano 
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
      TabIndex        =   53
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Tipo 
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
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Estado6 
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
      Left            =   2400
      TabIndex        =   52
      Top             =   6480
      Width           =   1335
   End
   Begin VB.ComboBox Estado5 
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
      Left            =   2400
      TabIndex        =   51
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox Estado4 
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
      Left            =   2400
      TabIndex        =   50
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox Estado3 
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
      Left            =   2400
      TabIndex        =   49
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox Estado2 
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
      Left            =   2400
      TabIndex        =   48
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox Estado1 
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
      Left            =   2400
      TabIndex        =   47
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Responsable1 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   30
      Text            =   " "
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Comentario62 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   6720
      Width           =   3375
   End
   Begin VB.TextBox Comentario61 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox Comentario52 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   6000
      Width           =   3375
   End
   Begin VB.TextBox Comentario51 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox Comentario42 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Comentario41 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   5040
      Width           =   3375
   End
   Begin VB.TextBox Comentario32 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Comentario31 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Comentario22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox Comentario21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   20
      Text            =   " "
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Comentario12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Comentario11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   18
      Text            =   " "
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Responsable2 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   17
      Text            =   " "
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Responsable3 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Responsable4 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   15
      Text            =   " "
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Responsable6 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   14
      Text            =   " "
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Responsable5 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   13
      Text            =   " "
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Titulo 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Text            =   " "
      Top             =   1560
      Width           =   10455
   End
   Begin VB.TextBox ResponsableDestino 
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableEmisor 
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha1 
      Height          =   285
      Left            =   840
      TabIndex        =   31
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha2 
      Height          =   285
      Left            =   840
      TabIndex        =   32
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha3 
      Height          =   285
      Left            =   840
      TabIndex        =   33
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha4 
      Height          =   285
      Left            =   840
      TabIndex        =   34
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha5 
      Height          =   285
      Left            =   840
      TabIndex        =   35
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha6 
      Height          =   285
      Left            =   840
      TabIndex        =   36
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1200
      TabIndex        =   58
      Top             =   480
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
   Begin MSMask.MaskEdBox Fecha11 
      Height          =   285
      Left            =   4080
      TabIndex        =   88
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha12 
      Height          =   285
      Left            =   4080
      TabIndex        =   89
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha13 
      Height          =   285
      Left            =   4080
      TabIndex        =   90
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha14 
      Height          =   285
      Left            =   4080
      TabIndex        =   91
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha15 
      Height          =   285
      Left            =   4080
      TabIndex        =   92
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fecha16 
      Height          =   285
      Left            =   4080
      TabIndex        =   93
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verificacion Implementacion"
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
      Left            =   840
      TabIndex        =   103
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Veriricacion Efectividad"
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
      Left            =   4080
      TabIndex        =   102
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label DesResponsable16 
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
      Left            =   4800
      TabIndex        =   101
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label DesResponsable15 
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
      Left            =   4800
      TabIndex        =   100
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label DesResponsable14 
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
      Left            =   4800
      TabIndex        =   99
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label DesResponsable13 
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
      Left            =   4800
      TabIndex        =   98
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label DesResponsable12 
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
      Left            =   4800
      TabIndex        =   97
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label DesResponsable11 
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
      Left            =   4800
      TabIndex        =   96
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   4080
      TabIndex        =   95
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   94
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "6"
      BeginProperty Font 
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
      Top             =   6480
      Width           =   195
   End
   Begin VB.Label Label30 
      Caption         =   "5"
      BeginProperty Font 
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
      Top             =   5760
      Width           =   195
   End
   Begin VB.Label Label29 
      Caption         =   "4"
      BeginProperty Font 
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
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Label28 
      Caption         =   "3"
      BeginProperty Font 
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
      Top             =   4320
      Width           =   195
   End
   Begin VB.Label Label27 
      Caption         =   "2"
      BeginProperty Font 
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
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label Label26 
      Caption         =   "1"
      BeginProperty Font 
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
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label Label11 
      Caption         =   "Referencia"
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
      TabIndex        =   69
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
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
      Left            =   3480
      TabIndex        =   67
      Top             =   120
      Width           =   735
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
      Left            =   120
      TabIndex        =   66
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Centro"
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
      Left            =   7200
      TabIndex        =   65
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesCentro 
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
      Left            =   9240
      TabIndex        =   64
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
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
      Left            =   7920
      TabIndex        =   63
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
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
      Left            =   3960
      TabIndex        =   62
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Numero"
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
      Left            =   5160
      TabIndex        =   61
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   120
      Width           =   735
   End
   Begin VB.Label DesTipo 
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
      Left            =   2040
      TabIndex        =   59
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   46
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   840
      TabIndex        =   45
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label DesResponsable1 
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
      Left            =   1560
      TabIndex        =   44
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acc"
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
      TabIndex        =   43
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comentarios"
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
      Left            =   7320
      TabIndex        =   42
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label DesResponsable2 
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
      Left            =   1560
      TabIndex        =   41
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label DesResponsable3 
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
      Left            =   1560
      TabIndex        =   40
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label DesResponsable4 
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
      Left            =   1560
      TabIndex        =   39
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label DesResponsable5 
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
      Left            =   1560
      TabIndex        =   38
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label DesResponsable6 
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
      Left            =   1560
      TabIndex        =   37
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Titulo"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label DesResponsableDestino 
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
      TabIndex        =   11
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableEmisor 
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
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "Resp. Inv."
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
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Emisor"
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
      Top             =   840
      Width           =   975
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "CargaSacVerifica.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacVerifica.frx":031F
      ToolTipText     =   "Salida"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   6720
      MouseIcon       =   "CargaSacVerifica.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacVerifica.frx":0E6B
      ToolTipText     =   "Elimina el Registro"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3480
      MouseIcon       =   "CargaSacVerifica.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacVerifica.frx":19B7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "CargaSacVerifica.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacVerifica.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7320
      Width           =   480
   End
End
Attribute VB_Name = "PrgCargaSacVerifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim rstCargaSacII As Recordset
Dim spCargaSacII As String
Dim rstCargaSacIV As Recordset
Dim spCargaSacIV As String

Dim XParam As String
Dim ZZLugar As Integer

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    

End Sub

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        Centro.Text = rstCargaSac!Centro
        Fecha.Text = rstCargaSac!Fecha
        Origen.ListIndex = rstCargaSac!Origen
        Estado.ListIndex = rstCargaSac!Estado
        ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSac!ResponsableDestino
        Referencia.Text = Trim(rstCargaSac!Referencia)
        Titulo.Text = Trim(rstCargaSac!Titulo)
        
        rstCargaSac.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
    
        rstCargaSacII.Close
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIV"
    ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIV = ZSql
    Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIV.RecordCount > 0 Then
    
        Responsable1.Text = rstCargaSacIV!Responsable1
        Responsable2.Text = rstCargaSacIV!Responsable2
        Responsable3.Text = rstCargaSacIV!Responsable3
        Responsable4.Text = rstCargaSacIV!Responsable4
        Responsable5.Text = rstCargaSacIV!Responsable5
        Responsable6.Text = rstCargaSacIV!Responsable6
        
        Responsable11.Text = IIf(IsNull(rstCargaSacIV!Responsable11), "", rstCargaSacIV!Responsable11)
        Responsable12.Text = IIf(IsNull(rstCargaSacIV!Responsable12), "", rstCargaSacIV!Responsable12)
        Responsable13.Text = IIf(IsNull(rstCargaSacIV!Responsable13), "", rstCargaSacIV!Responsable13)
        Responsable14.Text = IIf(IsNull(rstCargaSacIV!Responsable14), "", rstCargaSacIV!Responsable14)
        Responsable15.Text = IIf(IsNull(rstCargaSacIV!Responsable15), "", rstCargaSacIV!Responsable15)
        Responsable16.Text = IIf(IsNull(rstCargaSacIV!Responsable16), "", rstCargaSacIV!Responsable16)
        
        Fecha1.Text = rstCargaSacIV!Fecha1
        Fecha2.Text = rstCargaSacIV!Fecha2
        Fecha3.Text = rstCargaSacIV!Fecha3
        Fecha4.Text = rstCargaSacIV!Fecha4
        Fecha5.Text = rstCargaSacIV!Fecha5
        Fecha6.Text = rstCargaSacIV!Fecha6
        
        Fecha11.Text = IIf(IsNull(rstCargaSacIV!Fecha11), "  /  /    ", rstCargaSacIV!Fecha11)
        Fecha12.Text = IIf(IsNull(rstCargaSacIV!Fecha12), "  /  /    ", rstCargaSacIV!Fecha12)
        Fecha13.Text = IIf(IsNull(rstCargaSacIV!Fecha13), "  /  /    ", rstCargaSacIV!Fecha13)
        Fecha14.Text = IIf(IsNull(rstCargaSacIV!Fecha14), "  /  /    ", rstCargaSacIV!Fecha14)
        Fecha15.Text = IIf(IsNull(rstCargaSacIV!Fecha15), "  /  /    ", rstCargaSacIV!Fecha15)
        Fecha16.Text = IIf(IsNull(rstCargaSacIV!Fecha16), "  /  /    ", rstCargaSacIV!Fecha16)
        
        
        Comentario11.Text = Trim(rstCargaSacIV!Comentario11)
        Comentario12.Text = Trim(rstCargaSacIV!Comentario12)
        Comentario21.Text = Trim(rstCargaSacIV!Comentario21)
        Comentario22.Text = Trim(rstCargaSacIV!Comentario22)
        Comentario31.Text = Trim(rstCargaSacIV!Comentario31)
        Comentario32.Text = Trim(rstCargaSacIV!Comentario32)
        Comentario41.Text = Trim(rstCargaSacIV!Comentario41)
        Comentario42.Text = Trim(rstCargaSacIV!Comentario42)
        Comentario51.Text = Trim(rstCargaSacIV!Comentario51)
        Comentario52.Text = Trim(rstCargaSacIV!Comentario52)
        Comentario61.Text = Trim(rstCargaSacIV!Comentario61)
        Comentario62.Text = Trim(rstCargaSacIV!Comentario62)
        
        Estado1.ListIndex = rstCargaSacIV!Estado1
        Estado2.ListIndex = rstCargaSacIV!Estado2
        Estado3.ListIndex = rstCargaSacIV!Estado3
        Estado4.ListIndex = rstCargaSacIV!Estado4
        Estado5.ListIndex = rstCargaSacIV!Estado5
        Estado6.ListIndex = rstCargaSacIV!Estado6
        
        Estado11.ListIndex = IIf(IsNull(rstCargaSacIV!Estado11), "0", rstCargaSacIV!Estado11)
        Estado12.ListIndex = IIf(IsNull(rstCargaSacIV!Estado12), "0", rstCargaSacIV!Estado12)
        Estado13.ListIndex = IIf(IsNull(rstCargaSacIV!Estado13), "0", rstCargaSacIV!Estado13)
        Estado14.ListIndex = IIf(IsNull(rstCargaSacIV!Estado14), "0", rstCargaSacIV!Estado14)
        Estado15.ListIndex = IIf(IsNull(rstCargaSacIV!Estado15), "0", rstCargaSacIV!Estado15)
        Estado16.ListIndex = IIf(IsNull(rstCargaSacIV!Estado16), "0", rstCargaSacIV!Estado16)
        
        rstCargaSacIV.Close
    End If
    
    Call Imprime_Descripcion
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub cmdAdd_Click()

    If Val(Tipo.Text) <> 0 And Val(Ano.Text) <> 0 And Val(Numero.Text) <> 0 Then
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIV"
        ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + Numero.Text + "'"
        spCargaSacIV = ZSql
        Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIV.RecordCount > 0 Then
        
            rstCargaSacIV.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSacIV SET "
            ZSql = ZSql + " Responsable1 = " + "'" + Responsable1.Text + "',"
            ZSql = ZSql + " Responsable2 = " + "'" + Responsable2.Text + "',"
            ZSql = ZSql + " Responsable3 = " + "'" + Responsable3.Text + "',"
            ZSql = ZSql + " Responsable4 = " + "'" + Responsable4.Text + "',"
            ZSql = ZSql + " Responsable5 = " + "'" + Responsable5.Text + "',"
            ZSql = ZSql + " Responsable6 = " + "'" + Responsable6.Text + "',"
            ZSql = ZSql + " Responsable11 = " + "'" + Responsable11.Text + "',"
            ZSql = ZSql + " Responsable12 = " + "'" + Responsable12.Text + "',"
            ZSql = ZSql + " Responsable13 = " + "'" + Responsable13.Text + "',"
            ZSql = ZSql + " Responsable14 = " + "'" + Responsable14.Text + "',"
            ZSql = ZSql + " Responsable15 = " + "'" + Responsable15.Text + "',"
            ZSql = ZSql + " Responsable16 = " + "'" + Responsable16.Text + "',"
            ZSql = ZSql + " Fecha1 = " + "'" + Fecha1.Text + "',"
            ZSql = ZSql + " Fecha2 = " + "'" + Fecha2.Text + "',"
            ZSql = ZSql + " Fecha3 = " + "'" + Fecha3.Text + "',"
            ZSql = ZSql + " Fecha4 = " + "'" + Fecha4.Text + "',"
            ZSql = ZSql + " Fecha5 = " + "'" + Fecha5.Text + "',"
            ZSql = ZSql + " Fecha6 = " + "'" + Fecha6.Text + "',"
            ZSql = ZSql + " Fecha11 = " + "'" + Fecha11.Text + "',"
            ZSql = ZSql + " Fecha12 = " + "'" + Fecha12.Text + "',"
            ZSql = ZSql + " Fecha13 = " + "'" + Fecha13.Text + "',"
            ZSql = ZSql + " Fecha14 = " + "'" + Fecha14.Text + "',"
            ZSql = ZSql + " Fecha15 = " + "'" + Fecha15.Text + "',"
            ZSql = ZSql + " Fecha16 = " + "'" + Fecha16.Text + "',"
            ZSql = ZSql + " Comentario11 = " + "'" + Comentario11.Text + "',"
            ZSql = ZSql + " Comentario12 = " + "'" + Comentario12.Text + "',"
            ZSql = ZSql + " Comentario21 = " + "'" + Comentario21.Text + "',"
            ZSql = ZSql + " Comentario22 = " + "'" + Comentario22.Text + "',"
            ZSql = ZSql + " Comentario31 = " + "'" + Comentario31.Text + "',"
            ZSql = ZSql + " Comentario32 = " + "'" + Comentario32.Text + "',"
            ZSql = ZSql + " Comentario41 = " + "'" + Comentario41.Text + "',"
            ZSql = ZSql + " Comentario42 = " + "'" + Comentario42.Text + "',"
            ZSql = ZSql + " Comentario51 = " + "'" + Comentario51.Text + "',"
            ZSql = ZSql + " Comentario52 = " + "'" + Comentario52.Text + "',"
            ZSql = ZSql + " Comentario61 = " + "'" + Comentario61.Text + "',"
            ZSql = ZSql + " Comentario62 = " + "'" + Comentario62.Text + "',"
            ZSql = ZSql + " Estado1 = " + "'" + Str$(Estado1.ListIndex) + "',"
            ZSql = ZSql + " Estado2 = " + "'" + Str$(Estado2.ListIndex) + "',"
            ZSql = ZSql + " Estado3 = " + "'" + Str$(Estado3.ListIndex) + "',"
            ZSql = ZSql + " Estado4 = " + "'" + Str$(Estado4.ListIndex) + "',"
            ZSql = ZSql + " Estado5 = " + "'" + Str$(Estado5.ListIndex) + "',"
            ZSql = ZSql + " Estado6 = " + "'" + Str$(Estado6.ListIndex) + "',"
            ZSql = ZSql + " Estado11 = " + "'" + Str$(Estado11.ListIndex) + "',"
            ZSql = ZSql + " Estado12 = " + "'" + Str$(Estado12.ListIndex) + "',"
            ZSql = ZSql + " Estado13 = " + "'" + Str$(Estado13.ListIndex) + "',"
            ZSql = ZSql + " Estado14 = " + "'" + Str$(Estado14.ListIndex) + "',"
            ZSql = ZSql + " Estado15 = " + "'" + Str$(Estado15.ListIndex) + "',"
            ZSql = ZSql + " Estado16 = " + "'" + Str$(Estado16.ListIndex) + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSacIV = ZSql
            Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSacIV ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Responsable1 ,"
            ZSql = ZSql + "Responsable2 ,"
            ZSql = ZSql + "Responsable3 ,"
            ZSql = ZSql + "Responsable4 ,"
            ZSql = ZSql + "Responsable5 ,"
            ZSql = ZSql + "Responsable6 ,"
            ZSql = ZSql + "Responsable11 ,"
            ZSql = ZSql + "Responsable12 ,"
            ZSql = ZSql + "Responsable13 ,"
            ZSql = ZSql + "Responsable14 ,"
            ZSql = ZSql + "Responsable15 ,"
            ZSql = ZSql + "Responsable16 ,"
            ZSql = ZSql + "Fecha1 ,"
            ZSql = ZSql + "Fecha2 ,"
            ZSql = ZSql + "Fecha3 ,"
            ZSql = ZSql + "Fecha4 ,"
            ZSql = ZSql + "Fecha5 ,"
            ZSql = ZSql + "Fecha6 ,"
            ZSql = ZSql + "Fecha11 ,"
            ZSql = ZSql + "Fecha12 ,"
            ZSql = ZSql + "Fecha13 ,"
            ZSql = ZSql + "Fecha14 ,"
            ZSql = ZSql + "Fecha15 ,"
            ZSql = ZSql + "Fecha16 ,"
            ZSql = ZSql + "Comentario11 ,"
            ZSql = ZSql + "Comentario12 ,"
            ZSql = ZSql + "Comentario21 ,"
            ZSql = ZSql + "Comentario22 ,"
            ZSql = ZSql + "Comentario31 ,"
            ZSql = ZSql + "Comentario32 ,"
            ZSql = ZSql + "Comentario41 ,"
            ZSql = ZSql + "Comentario42 ,"
            ZSql = ZSql + "Comentario51 ,"
            ZSql = ZSql + "Comentario52 ,"
            ZSql = ZSql + "Comentario61 ,"
            ZSql = ZSql + "Comentario62 ,"
            ZSql = ZSql + "Estado1 ,"
            ZSql = ZSql + "Estado2 ,"
            ZSql = ZSql + "Estado3 ,"
            ZSql = ZSql + "Estado4 ,"
            ZSql = ZSql + "Estado5 ,"
            ZSql = ZSql + "Estado6 ,"
            ZSql = ZSql + "Estado11 ,"
            ZSql = ZSql + "Estado12 ,"
            ZSql = ZSql + "Estado13 ,"
            ZSql = ZSql + "Estado14 ,"
            ZSql = ZSql + "Estado15 ,"
            ZSql = ZSql + "Estado16 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Tipo.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Responsable1.Text + "',"
            ZSql = ZSql + "'" + Responsable2.Text + "',"
            ZSql = ZSql + "'" + Responsable3.Text + "',"
            ZSql = ZSql + "'" + Responsable4.Text + "',"
            ZSql = ZSql + "'" + Responsable5.Text + "',"
            ZSql = ZSql + "'" + Responsable6.Text + "',"
            ZSql = ZSql + "'" + Responsable11.Text + "',"
            ZSql = ZSql + "'" + Responsable12.Text + "',"
            ZSql = ZSql + "'" + Responsable13.Text + "',"
            ZSql = ZSql + "'" + Responsable14.Text + "',"
            ZSql = ZSql + "'" + Responsable15.Text + "',"
            ZSql = ZSql + "'" + Responsable16.Text + "',"
            ZSql = ZSql + "'" + Fecha1.Text + "',"
            ZSql = ZSql + "'" + Fecha2.Text + "',"
            ZSql = ZSql + "'" + Fecha3.Text + "',"
            ZSql = ZSql + "'" + Fecha4.Text + "',"
            ZSql = ZSql + "'" + Fecha5.Text + "',"
            ZSql = ZSql + "'" + Fecha6.Text + "',"
            ZSql = ZSql + "'" + Fecha11.Text + "',"
            ZSql = ZSql + "'" + Fecha12.Text + "',"
            ZSql = ZSql + "'" + Fecha13.Text + "',"
            ZSql = ZSql + "'" + Fecha14.Text + "',"
            ZSql = ZSql + "'" + Fecha15.Text + "',"
            ZSql = ZSql + "'" + Fecha16.Text + "',"
            ZSql = ZSql + "'" + Comentario11.Text + "',"
            ZSql = ZSql + "'" + Comentario12.Text + "',"
            ZSql = ZSql + "'" + Comentario21.Text + "',"
            ZSql = ZSql + "'" + Comentario22.Text + "',"
            ZSql = ZSql + "'" + Comentario31.Text + "',"
            ZSql = ZSql + "'" + Comentario32.Text + "',"
            ZSql = ZSql + "'" + Comentario41.Text + "',"
            ZSql = ZSql + "'" + Comentario42.Text + "',"
            ZSql = ZSql + "'" + Comentario51.Text + "',"
            ZSql = ZSql + "'" + Comentario52.Text + "',"
            ZSql = ZSql + "'" + Comentario61.Text + "',"
            ZSql = ZSql + "'" + Comentario62.Text + "',"
            ZSql = ZSql + "'" + Str$(Estado1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado5.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado6.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado11.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado12.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado13.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado14.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado15.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado16.ListIndex) + "')"
            
            spCargaSacIV = ZSql
            Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
        
        
                
        ZEntra = "S"
        If ZEntra = "S" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
        
                ZZEstado = rstCargaSac!Estado
                rstCargaSac.Close
            
                If ZZEstado <= 4 Then
            
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CargaSac SET "
                    ZSql = ZSql + " Estado = " + "'" + "5" + "'"
                    ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                    ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                    ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                    spCargaSac = ZSql
                    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
                
            End If
            
        End If
        
        

        
        
        
    
        
        Call CmdLimpiar_Click
        Tipo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIV"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacIV = ZSql
        Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIV.RecordCount > 0 Then
        
            rstCargaSacIV.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                Sql1 = ZSql + "DELETE CargaSacIV"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSacIV = Sql1 + Sql2
                Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    End If
    
    Tipo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    Referencia.Text = ""
    Titulo.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    
    Origen.ListIndex = 0
    Estado.ListIndex = 0
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""
    
    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    Fecha11.Text = "  /  /    "
    Fecha12.Text = "  /  /    "
    Fecha13.Text = "  /  /    "
    Fecha14.Text = "  /  /    "
    Fecha15.Text = "  /  /    "
    Fecha16.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    Estado11.ListIndex = 0
    Estado12.ListIndex = 0
    Estado13.ListIndex = 0
    Estado14.ListIndex = 0
    Estado15.ListIndex = 0
    Estado16.ListIndex = 0
    
    Tipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    PrgCargaSacVerifica.Hide
    Unload Me
    Menu.Show
    
End Sub

Sub Form_Load()

    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    Referencia.Text = ""
    Titulo.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""
    
    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    Fecha11.Text = "  /  /    "
    Fecha12.Text = "  /  /    "
    Fecha13.Text = "  /  /    "
    Fecha14.Text = "  /  /    "
    Fecha15.Text = "  /  /    "
    Fecha16.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "INICIADA"
    Estado.AddItem "INVESTIGACION"
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "IMPLEMENTACION A VERIFICAR"
    Estado.AddItem "IMPLEMENTACION VERIFICADA"
    Estado.AddItem "CERRADA"
    Estado.AddItem "ANULADA"
    
    Estado.ListIndex = 0
    
    
    Estado1.Clear
    
    Estado1.AddItem "No Imple."
    Estado1.AddItem "Imple."
    Estado1.AddItem "Nula"
    Estado1.AddItem "Cerrada"
    Estado1.AddItem ""
    
    Estado1.ListIndex = 0
    
    Estado2.Clear
    
    Estado2.AddItem "No Imple."
    Estado2.AddItem "Imple."
    Estado2.AddItem "Nula"
    Estado2.AddItem "Cerrada"
    Estado2.AddItem ""
    
    Estado2.ListIndex = 0
    
    Estado3.Clear
    
    Estado3.AddItem "No Imple."
    Estado3.AddItem "Imple."
    Estado3.AddItem "Nula"
    Estado3.AddItem "Cerrada"
    Estado3.AddItem ""
    
    Estado3.ListIndex = 0
    
    Estado4.Clear
    
    Estado4.AddItem "No Imple."
    Estado4.AddItem "Imple."
    Estado4.AddItem "Nula"
    Estado4.AddItem "Cerrada"
    Estado4.AddItem ""
    
    Estado4.ListIndex = 0
    
    Estado5.Clear
    
    Estado5.AddItem "No Imple."
    Estado5.AddItem "Imple."
    Estado5.AddItem "Nula"
    Estado5.AddItem "Cerrada"
    Estado5.AddItem ""
 
    Estado5.ListIndex = 0
    
    Estado6.Clear
    
    Estado6.AddItem "No Imple."
    Estado6.AddItem "Imple."
    Estado6.AddItem "Nula"
    Estado6.AddItem "Cerrada"
    Estado6.AddItem ""
    
    Estado6.ListIndex = 0
    
    
    
    
    Estado11.Clear
    
    Estado11.AddItem "No Imple."
    Estado11.AddItem "Imple."
    Estado11.AddItem "Nula"
    Estado11.AddItem "Cerrada"
    Estado11.AddItem ""
    
    Estado11.ListIndex = 0
    
    Estado12.Clear
    
    Estado12.AddItem "No Imple."
    Estado12.AddItem "Imple."
    Estado12.AddItem "Nula"
    Estado12.AddItem "Cerrada"
    Estado12.AddItem ""
    
    Estado12.ListIndex = 0
    
    Estado13.Clear
    
    Estado13.AddItem "No Imple."
    Estado13.AddItem "Imple."
    Estado13.AddItem "Nula"
    Estado13.AddItem "Cerrada"
    Estado13.AddItem ""
    
    Estado13.ListIndex = 0
    
    Estado14.Clear
    
    Estado14.AddItem "No Imple."
    Estado14.AddItem "Imple."
    Estado14.AddItem "Nula"
    Estado14.AddItem "Cerrada"
    Estado14.AddItem ""
    
    Estado14.ListIndex = 0
    
    Estado15.Clear
    
    Estado15.AddItem "No Imple."
    Estado15.AddItem "Imple."
    Estado15.AddItem "Nula"
    Estado15.AddItem "Cerrada"
    Estado15.AddItem ""
 
    Estado15.ListIndex = 0
    
    Estado16.Clear
    
    Estado16.AddItem "No Imple."
    Estado16.AddItem "Imple."
    Estado16.AddItem "Nula"
    Estado16.AddItem "Cerrada"
    Estado16.AddItem ""
    
    Estado16.ListIndex = 0
    
    
    
    
    
    Origen.Clear
    
    Origen.AddItem ""
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            DesTipo.Caption = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                rstCargaSac.Close
                Call Imprime_Datos
                Responsable1.SetFocus
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Ano.Text = WAno
                Numero.Text = WNumero
                Tipo.Text = WTipo
                Sql1 = "Select *"
                Sql2 = " FROM TipoSac"
                Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
                spTipoSac = Sql1 + Sql2 + Sql3
                Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstTipoSac.RecordCount > 0 Then
                    DesTipo.Caption = Trim(rstTipoSac!Descripcion)
                    rstTipoSac.Close
                    Ano.SetFocus
                End If
                Tipo.SetFocus
                
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Responsable1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable1.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha1.SetFocus
            End If
                Else
            DesResponsable1.Caption = ""
            Fecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable1.Text = ""
        DesResponsable1.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha1.Text, Auxi)
        If Auxi = "S" Or Fecha1.Text = "  /  /    " Then
            Estado1.SetFocus
                Else
            Fecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha1.Text = "  /  /    "
    End If
End Sub

Private Sub Estado1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable11.SetFocus
    End If
End Sub

Private Sub Responsable11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable11.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha11.SetFocus
            End If
                Else
            DesResponsable11.Caption = ""
            Fecha11.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable11.Text = ""
        DesResponsable11.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha11.Text, Auxi)
        If Auxi = "S" Or Fecha11.Text = "  /  /    " Then
            Estado11.SetFocus
                Else
            Fecha11.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha11.Text = "  /  /    "
    End If
End Sub

Private Sub Estado11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario11.SetFocus
    End If
End Sub

Private Sub Comentario11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario12.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario11.Text = ""
    End If
End Sub

Private Sub Comentario12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable2.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario12.Text = ""
    End If
End Sub








Private Sub Responsable2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable2.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha2.SetFocus
            End If
                Else
            DesResponsable2.Caption = ""
            Fecha2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable2.Text = ""
        DesResponsable2.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha2.Text, Auxi)
        If Auxi = "S" Or Fecha2.Text = "  /  /    " Then
            Estado2.SetFocus
                Else
            Fecha2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha2.Text = "  /  /    "
    End If
End Sub

Private Sub Estado2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable12.SetFocus
    End If
End Sub

Private Sub Responsable12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable12.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha12.SetFocus
            End If
                Else
            DesResponsable12.Caption = ""
            Fecha12.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable12.Text = ""
        DesResponsable12.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha12.Text, Auxi)
        If Auxi = "S" Or Fecha12.Text = "  /  /    " Then
            Estado12.SetFocus
                Else
            Fecha12.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha12.Text = "  /  /    "
    End If
End Sub

Private Sub Estado12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario21.SetFocus
    End If
End Sub

Private Sub Comentario21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario22.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario21.Text = ""
    End If
End Sub

Private Sub Comentario22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable3.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario22.Text = ""
    End If
End Sub





Private Sub Responsable3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable3.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha3.SetFocus
            End If
                Else
            DesResponsable3.Caption = ""
            Fecha3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable3.Text = ""
        DesResponsable3.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha3.Text, Auxi)
        If Auxi = "S" Or Fecha3.Text = "  /  /    " Then
            Estado3.SetFocus
                Else
            Fecha3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha3.Text = "  /  /    "
    End If
End Sub

Private Sub Estado3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable13.SetFocus
    End If
End Sub

Private Sub Responsable13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable13.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha13.SetFocus
            End If
                Else
            DesResponsable13.Caption = ""
            Fecha13.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable13.Text = ""
        DesResponsable13.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha13.Text, Auxi)
        If Auxi = "S" Or Fecha13.Text = "  /  /    " Then
            Estado13.SetFocus
                Else
            Fecha13.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha13.Text = "  /  /    "
    End If
End Sub

Private Sub Estado13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario31.SetFocus
    End If
End Sub



Private Sub Comentario31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario32.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario31.Text = ""
    End If
End Sub

Private Sub Comentario32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable4.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario32.Text = ""
    End If
End Sub









Private Sub Responsable4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable4.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha4.SetFocus
            End If
                Else
            DesResponsable4.Caption = ""
            Fecha4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable4.Text = ""
        DesResponsable4.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha4.Text, Auxi)
        If Auxi = "S" Or Fecha4.Text = "  /  /    " Then
            Estado4.SetFocus
                Else
            Fecha4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha4.Text = "  /  /    "
    End If
End Sub

Private Sub Estado4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable14.SetFocus
    End If
End Sub

Private Sub Responsable14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable14.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha14.SetFocus
            End If
                Else
            DesResponsable14.Caption = ""
            Fecha14.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable14.Text = ""
        DesResponsable14.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha14.Text, Auxi)
        If Auxi = "S" Or Fecha14.Text = "  /  /    " Then
            Estado14.SetFocus
                Else
            Fecha14.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha14.Text = "  /  /    "
    End If
End Sub

Private Sub Estado14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario41.SetFocus
    End If
End Sub


Private Sub Comentario41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario42.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario41.Text = ""
    End If
End Sub

Private Sub Comentario42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable5.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario42.Text = ""
    End If
End Sub





Private Sub Responsable5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable5.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha5.SetFocus
            End If
                Else
            DesResponsable5.Caption = ""
            Fecha5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable5.Text = ""
        DesResponsable5.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha5.Text, Auxi)
        If Auxi = "S" Or Fecha5.Text = "  /  /    " Then
            Estado5.SetFocus
                Else
            Fecha5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha5.Text = "  /  /    "
    End If
End Sub


Private Sub Estado5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable15.SetFocus
    End If
End Sub

Private Sub Responsable15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable15.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha15.SetFocus
            End If
                Else
            DesResponsable15.Caption = ""
            Fecha15.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable15.Text = ""
        DesResponsable15.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha15.Text, Auxi)
        If Auxi = "S" Or Fecha15.Text = "  /  /    " Then
            Estado15.SetFocus
                Else
            Fecha15.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha15.Text = "  /  /    "
    End If
End Sub

Private Sub Estado15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario51.SetFocus
    End If
End Sub


Private Sub Comentario51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario52.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario51.Text = ""
    End If
End Sub

Private Sub Comentario52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable6.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario52.Text = ""
    End If
End Sub





Private Sub Responsable6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable6.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha6.SetFocus
            End If
                Else
            DesResponsable6.Caption = ""
            Fecha6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable6.Text = ""
        DesResponsable6.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha6.Text, Auxi)
        If Auxi = "S" Or Fecha6.Text = "  /  /    " Then
            Estado6.SetFocus
                Else
            Fecha6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha6.Text = "  /  /    "
    End If
End Sub


Private Sub Estado6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable16.SetFocus
    End If
End Sub


Private Sub Responsable16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable16.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha16.SetFocus
            End If
                Else
            DesResponsable16.Caption = ""
            Fecha16.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable16.Text = ""
        DesResponsable16.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha16.Text, Auxi)
        If Auxi = "S" Or Fecha16.Text = "  /  /    " Then
            Estado16.SetFocus
                Else
            Fecha16.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha16.Text = "  /  /    "
    End If
End Sub

Private Sub Estado16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario61.SetFocus
    End If
End Sub

Private Sub Comentario61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario62.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario61.Text = ""
    End If
End Sub

Private Sub Comentario62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable1.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario62.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"
     Opcion.AddItem "Responsables"

     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM tiposac"
            Sql3 = " Order by tiposac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTipoSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(Indice)
            Call Tipo_Keypress(13)
            
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Select Case ZZLugar
                Case 1
                    Responsable1.Text = WIndice.List(Indice)
                    Call Responsable1_Keypress(13)
                Case 2
                    Responsable2.Text = WIndice.List(Indice)
                    Call Responsable2_Keypress(13)
                Case 3
                    Responsable3.Text = WIndice.List(Indice)
                    Call Responsable3_Keypress(13)
                Case 4
                    Responsable4.Text = WIndice.List(Indice)
                    Call Responsable4_Keypress(13)
                Case 5
                    Responsable5.Text = WIndice.List(Indice)
                    Call Responsable5_Keypress(13)
                Case 6
                    Responsable6.Text = WIndice.List(Indice)
                    Call Responsable6_Keypress(13)
                Case 7
                    Responsable1.Text = WIndice.List(Indice)
                    Call Responsable11_Keypress(13)
                Case 8
                    Responsable2.Text = WIndice.List(Indice)
                    Call Responsable12_Keypress(13)
                Case 9
                    Responsable3.Text = WIndice.List(Indice)
                    Call Responsable13_Keypress(13)
                Case 10
                    Responsable4.Text = WIndice.List(Indice)
                    Call Responsable14_Keypress(13)
                Case 11
                    Responsable5.Text = WIndice.List(Indice)
                    Call Responsable15_Keypress(13)
                Case 12
                    Responsable6.Text = WIndice.List(Indice)
                    Call Responsable16_Keypress(13)
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    LugarAyuda = 0
    WIndice.Clear
    Pantalla.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM TipoSac"
            Sql3 = " Order by TipoSac.Codigo"
            spTipoSac = Sql1 + Sql2 + Sql3
            Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoSac.RecordCount > 0 Then
                With rstTipoSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTipoSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTipoSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstTipoSac!Codigo) + " " + rstTipoSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTipoSac!Codigo
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
                rstTipoSac.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstResponsableSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstResponsableSac!Codigo
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
                rstResponsableSac.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Tipo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable1_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable2_DblClick()

    ZZLugar = 2

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable3_DblClick()

    ZZLugar = 3

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable4_DblClick()

    ZZLugar = 4

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable5_DblClick()

    ZZLugar = 5

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable6_DblClick()

    ZZLugar = 6

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub


Private Sub Responsable11_DblClick()

    ZZLugar = 7

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable12_DblClick()

    ZZLugar = 8

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable13_DblClick()

    ZZLugar = 9

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable14_DblClick()

    ZZLugar = 10

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable15_DblClick()

    ZZLugar = 11

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable16_DblClick()

    ZZLugar = 12

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub


