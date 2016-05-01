VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaSacImplementa 
   Caption         =   "Carga de SAC - Implementacion"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11775
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
      TabIndex        =   81
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      Text            =   " "
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   6000
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   5280
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   4560
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   3840
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   3240
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   2520
      Width           =   1335
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
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
      Height          =   2460
      ItemData        =   "cargasacimplementa.frx":0000
      Left            =   1560
      List            =   "cargasacimplementa.frx":0007
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Accion11 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Accion12 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Accion21 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Accion22 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   39
      Text            =   " "
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Accion31 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Accion32 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   37
      Text            =   " "
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Accion41 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   36
      Text            =   " "
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Accion42 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Accion51 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Accion52 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox Accion61 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   32
      Text            =   " "
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox Accion62 
      BeginProperty Font 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   6240
      Width           =   3135
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   30
      Text            =   " "
      Top             =   2520
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   6240
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   6000
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   5520
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   5280
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   4800
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   4560
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   4080
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   3840
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   3480
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   20
      Text            =   " "
      Top             =   3240
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   2760
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
      Left            =   8280
      MaxLength       =   50
      TabIndex        =   18
      Text            =   " "
      Top             =   2520
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   17
      Text            =   " "
      Top             =   3240
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   3840
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   15
      Text            =   " "
      Top             =   4560
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   14
      Text            =   " "
      Top             =   6000
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   13
      Text            =   " "
      Top             =   5280
      Width           =   615
   End
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
      Left            =   1320
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   5640
      TabIndex        =   43
      Top             =   2520
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
      Left            =   5640
      TabIndex        =   44
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
   Begin MSMask.MaskEdBox Fecha3 
      Height          =   285
      Left            =   5640
      TabIndex        =   45
      Top             =   3840
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
      Left            =   5640
      TabIndex        =   46
      Top             =   4560
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
      Left            =   5640
      TabIndex        =   47
      Top             =   5280
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
      Left            =   5640
      TabIndex        =   48
      Top             =   6000
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
      TabIndex        =   71
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
      Left            =   75
      TabIndex        =   88
      Top             =   6000
      Width           =   135
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
      Left            =   75
      TabIndex        =   87
      Top             =   5280
      Width           =   135
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
      Left            =   75
      TabIndex        =   86
      Top             =   4560
      Width           =   135
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
      Left            =   75
      TabIndex        =   85
      Top             =   3840
      Width           =   135
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
      Left            =   75
      TabIndex        =   84
      Top             =   3240
      Width           =   135
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
      Left            =   75
      TabIndex        =   83
      Top             =   2520
      Width           =   135
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
      TabIndex        =   82
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
      TabIndex        =   80
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   74
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      Left            =   6840
      TabIndex        =   59
      Top             =   2160
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
      Left            =   3600
      TabIndex        =   58
      Top             =   2160
      Width           =   1935
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
      Left            =   4320
      TabIndex        =   57
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
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
      Left            =   5640
      TabIndex        =   56
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acciones Correctivas"
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
      Left            =   360
      TabIndex        =   55
      Top             =   2160
      Width           =   3135
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
      Left            =   8280
      TabIndex        =   54
      Top             =   2160
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
      Left            =   4320
      TabIndex        =   53
      Top             =   3240
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   52
      Top             =   3840
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   51
      Top             =   4560
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   50
      Top             =   5280
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   49
      Top             =   6000
      Width           =   1215
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
      Left            =   6600
      MouseIcon       =   "cargasacimplementa.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "cargasacimplementa.frx":031F
      ToolTipText     =   "Salida"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "cargasacimplementa.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "cargasacimplementa.frx":0E6B
      ToolTipText     =   "Elimina el Registro"
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "cargasacimplementa.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "cargasacimplementa.frx":19B7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5640
      MouseIcon       =   "cargasacimplementa.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "cargasacimplementa.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6720
      Width           =   480
   End
End
Attribute VB_Name = "PrgCargaSacImplementa"
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
Dim rstCargaSacIII As Recordset
Dim spCargaSacIII As String

Dim XParam As String
Dim ZZLugar As Integer
Dim ZResponsableDestino As Integer
Dim ZResponsableCentro As Integer
Dim ZResponsable1 As Integer
Dim ZResponsable2 As Integer
Dim ZResponsable3 As Integer
Dim ZResponsable4 As Integer
Dim ZResponsable5 As Integer
Dim ZResponsable6 As Integer

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String


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
    
        Accion11.Text = rstCargaSacII!Accion11
        Accion12.Text = rstCargaSacII!Accion12
        Accion21.Text = rstCargaSacII!Accion21
        Accion22.Text = rstCargaSacII!Accion22
        Accion31.Text = rstCargaSacII!Accion31
        Accion32.Text = rstCargaSacII!Accion32
        Accion41.Text = rstCargaSacII!Accion41
        Accion42.Text = rstCargaSacII!Accion42
        Accion51.Text = rstCargaSacII!Accion51
        Accion52.Text = rstCargaSacII!Accion52
        Accion61.Text = rstCargaSacII!Accion61
        Accion62.Text = rstCargaSacII!Accion62
        
        ZResponsable1 = rstCargaSacII!Responsable1
        ZResponsable2 = rstCargaSacII!Responsable2
        ZResponsable3 = rstCargaSacII!Responsable3
        ZResponsable4 = rstCargaSacII!Responsable4
        ZResponsable5 = rstCargaSacII!Responsable5
        ZResponsable6 = rstCargaSacII!Responsable6
        
        rstCargaSacII.Close
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIII"
    ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIII = ZSql
    Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIII.RecordCount > 0 Then
    
        Responsable1.Text = rstCargaSacIII!Responsable1
        Responsable2.Text = rstCargaSacIII!Responsable2
        Responsable3.Text = rstCargaSacIII!Responsable3
        Responsable4.Text = rstCargaSacIII!Responsable4
        Responsable5.Text = rstCargaSacIII!Responsable5
        Responsable6.Text = rstCargaSacIII!Responsable6
        
        Fecha1.Text = rstCargaSacIII!Fecha1
        Fecha2.Text = rstCargaSacIII!Fecha2
        Fecha3.Text = rstCargaSacIII!Fecha3
        Fecha4.Text = rstCargaSacIII!Fecha4
        Fecha5.Text = rstCargaSacIII!Fecha5
        Fecha6.Text = rstCargaSacIII!Fecha6
        
        Comentario11.Text = Trim(rstCargaSacIII!Comentario11)
        Comentario12.Text = Trim(rstCargaSacIII!Comentario12)
        Comentario21.Text = Trim(rstCargaSacIII!Comentario21)
        Comentario22.Text = Trim(rstCargaSacIII!Comentario22)
        Comentario31.Text = Trim(rstCargaSacIII!Comentario31)
        Comentario32.Text = Trim(rstCargaSacIII!Comentario32)
        Comentario41.Text = Trim(rstCargaSacIII!Comentario41)
        Comentario42.Text = Trim(rstCargaSacIII!Comentario42)
        Comentario51.Text = Trim(rstCargaSacIII!Comentario51)
        Comentario52.Text = Trim(rstCargaSacIII!Comentario52)
        Comentario61.Text = Trim(rstCargaSacIII!Comentario61)
        Comentario62.Text = Trim(rstCargaSacIII!Comentario62)
        
        Estado1.ListIndex = rstCargaSacIII!Estado1
        Estado2.ListIndex = rstCargaSacIII!Estado2
        Estado3.ListIndex = rstCargaSacIII!Estado3
        Estado4.ListIndex = rstCargaSacIII!Estado4
        Estado5.ListIndex = rstCargaSacIII!Estado5
        Estado6.ListIndex = rstCargaSacIII!Estado6
        
        rstCargaSacIII.Close
    End If
    
    
    
    
    Responsable1.Locked = False
    Responsable2.Locked = False
    Responsable3.Locked = False
    Responsable4.Locked = False
    Responsable5.Locked = False
    Responsable6.Locked = False
    
    Fecha1.Enabled = True
    Fecha2.Enabled = True
    Fecha3.Enabled = True
    Fecha4.Enabled = True
    Fecha5.Enabled = True
    Fecha6.Enabled = True
    
    Comentario11.Locked = False
    Comentario12.Locked = False
    Comentario21.Locked = False
    Comentario22.Locked = False
    Comentario31.Locked = False
    Comentario32.Locked = False
    Comentario41.Locked = False
    Comentario42.Locked = False
    Comentario51.Locked = False
    Comentario52.Locked = False
    Comentario61.Locked = False
    Comentario62.Locked = False
    
    Estado1.Locked = False
    Estado2.Locked = False
    Estado3.Locked = False
    Estado4.Locked = False
    Estado5.Locked = False
    Estado6.Locked = False
    Estado7.Locked = False
    
    GoTo da
    
    If ZResponsable1 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable1.Locked = True
        Fecha1.Enabled = False
        Comentario11.Locked = True
        Comentario12.Locked = True
        Estado1.Locked = True
    End If
    
    If ZResponsable2 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable2.Locked = True
        Fecha2.Enabled = False
        Comentario21.Locked = True
        Comentario22.Locked = True
        Estado2.Locked = True
    End If
    
    If ZResponsable3 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable3.Locked = True
        Fecha3.Enabled = False
        Comentario31.Locked = True
        Comentario32.Locked = True
        Estado3.Locked = True
    End If
    
    If ZResponsable4 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable4.Locked = True
        Fecha4.Enabled = False
        Comentario41.Locked = True
        Comentario42.Locked = True
        Estado4.Locked = True
    End If
    
    If ZResponsable5 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable5.Locked = True
        Fecha5.Enabled = False
        Comentario51.Locked = True
        Comentario52.Locked = True
        Estado5.Locked = True
    End If
    
    If ZResponsable6 <> ZZCodigoResponsable And ZZCodigoResponsable <> 99 Then
        Responsable6.Locked = True
        Fecha6.Enabled = False
        Comentario61.Locked = True
        Comentario62.Locked = True
        Estado6.Locked = True
    End If
    
da:

    
    Call Imprime_Descripcion
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub cmdAdd_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIII"
        ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIII.RecordCount > 0 Then
        
            rstCargaSacIII.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSacIII SET "
            ZSql = ZSql + " Responsable1 = " + "'" + Responsable1.Text + "',"
            ZSql = ZSql + " Responsable2 = " + "'" + Responsable2.Text + "',"
            ZSql = ZSql + " Responsable3 = " + "'" + Responsable3.Text + "',"
            ZSql = ZSql + " Responsable4 = " + "'" + Responsable4.Text + "',"
            ZSql = ZSql + " Responsable5 = " + "'" + Responsable5.Text + "',"
            ZSql = ZSql + " Responsable6 = " + "'" + Responsable6.Text + "',"
            ZSql = ZSql + " Fecha1 = " + "'" + Fecha1.Text + "',"
            ZSql = ZSql + " Fecha2 = " + "'" + Fecha2.Text + "',"
            ZSql = ZSql + " Fecha3 = " + "'" + Fecha3.Text + "',"
            ZSql = ZSql + " Fecha4 = " + "'" + Fecha4.Text + "',"
            ZSql = ZSql + " Fecha5 = " + "'" + Fecha5.Text + "',"
            ZSql = ZSql + " Fecha6 = " + "'" + Fecha6.Text + "',"
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
            ZSql = ZSql + " Estado6 = " + "'" + Str$(Estado6.ListIndex) + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSacIII = ZSql
            Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSacIII ("
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
            ZSql = ZSql + "Fecha1 ,"
            ZSql = ZSql + "Fecha2 ,"
            ZSql = ZSql + "Fecha3 ,"
            ZSql = ZSql + "Fecha4 ,"
            ZSql = ZSql + "Fecha5 ,"
            ZSql = ZSql + "Fecha6 ,"
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
            ZSql = ZSql + "Estado6 )"
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
            ZSql = ZSql + "'" + Fecha1.Text + "',"
            ZSql = ZSql + "'" + Fecha2.Text + "',"
            ZSql = ZSql + "'" + Fecha3.Text + "',"
            ZSql = ZSql + "'" + Fecha4.Text + "',"
            ZSql = ZSql + "'" + Fecha5.Text + "',"
            ZSql = ZSql + "'" + Fecha6.Text + "',"
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
            ZSql = ZSql + "'" + Str$(Estado6.ListIndex) + "')"
            
            spCargaSacIII = ZSql
            Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        T$ = "Carga de Implementacion de Acciones Correctivas"
        m$ = "Desea enviar el aviso al Responsable del Area"
        ZRespuesta% = MsgBox(m$, 32 + 4, T$)
        If ZRespuesta% = 6 Then
        
            ZZResponsable = 0
        
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                ZZResponsable = rstCentroSac!Responsable
                rstCentroSac.Close
            End If
            
            If ZZResponsable <> 0 Then
            
                ZZEmail = ""
                
                Sql1 = "Select *"
                Sql2 = " FROM ResponsableSac"
                Sql3 = " Where ResponsableSac.Codigo = " + "'" + Str$(ZZResponsable) + "'"
                spResponsableSac = Sql1 + Sql2 + Sql3
                Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstResponsableSac.RecordCount > 0 Then
                    ZZEmail = Trim(rstResponsableSac!Email)
                    rstResponsableSac.Close
                End If
                
                If ZZEmail <> "" Then
            
                    sTo = ZZEmail
                    sCC = ""
                    sBCC = ""
                    Select Case Val(Tipo.Text)
                        Case 1
                            sSubject = "Aviso de Implementacion de Acciones Correctivas"
                            sBody = "Se informa la implementacion de accciones correctivas por " + _
                                    DesTipo.Caption + " : " + _
                                    Ano.Text + "/" + Numero.Text + _
                                    " para su verificacion. " + _
                                    " Referencia : " + Referencia.Text + _
                                    " Titulo : " + Titulo.Text
                        Case 2
                            sSubject = "Aviso de Implementacion de Acciones Preventivas"
                            sBody = "Se informa la implementacion de accciones preventivas por " + _
                                    DesTipo.Caption + " : " + _
                                    Ano.Text + "/" + Numero.Text + _
                                    " para su verificacion. " + _
                                    " Referencia : " + Referencia.Text + _
                                    " Titulo : " + Titulo.Text
                        Case Else
                            sSubject = "Aviso de Implementacion de Acciones de " + DesTipo.Caption
                            sBody = "Se informa la implementacion de accciones por " + _
                                    DesTipo.Caption + " : " + _
                                    Ano.Text + "/" + Numero.Text + _
                                    " para su verificacion. " + _
                                    " Referencia : " + Referencia.Text + _
                                    " Titulo : " + Titulo.Text
                    End Select

                    ret = Shell("Start.exe " _
                        & "mailto:" & """" & sTo & """" _
                        & "?Subject=" & """" & sSubject & """" _
                        & "&cc=" & """" & sCC & """" _
                        & "&bcc=" & """" & sBCC & """" _
                        & "&Body=" & """" & sBody & """" _
                        & "&File=" & """" & "c:\autoexec.bat" & """" _
                        , 0)
            
                End If
            End If
        End If
        
        T$ = "Carga de Implementacion de Acciones Correctivas"
        m$ = "Desea enviar el aviso al Responsable de Calidad"
        ZRespuesta% = MsgBox(m$, 32 + 4, T$)
        If ZRespuesta% = 6 Then
        
            ZZEmail = "ebiglieri@surfactan.com.ar; calidad@surfactan.com.ar"
            
            sTo = ZZEmail
            sCC = ""
            sBCC = ""
            sSubject = "Aviso de Implementacion de Acciones Correctivas"
            sBody = "Se informaron implementaciones de  accciones corectivas del " + DesTipo.Caption + " : " + Ano.Text + "/" + Numero.Text + " para su verificacion    Referencia : " + Referencia.Text
    
            ret = Shell("Start.exe " _
                        & "mailto:" & """" & sTo & """" _
                        & "?Subject=" & """" & sSubject & """" _
                        & "&cc=" & """" & sCC & """" _
                        & "&bcc=" & """" & sBCC & """" _
                        & "&Body=" & """" & sBody & """" _
                        & "&File=" & """" & "c:\autoexec.bat" & """" _
                        , 0)
        End If
        
        
        ZEntra = "S"
        
        If Trim(Accion11.Text) <> "" And Estado1.ListIndex = 0 Then
            ZEntra = "N"
        End If
        If Trim(Accion21.Text) <> "" And Estado2.ListIndex = 0 Then
            ZEntra = "N"
        End If
        If Trim(Accion31.Text) <> "" And Estado3.ListIndex = 0 Then
            ZEntra = "N"
        End If
        If Trim(Accion41.Text) <> "" And Estado4.ListIndex = 0 Then
            ZEntra = "N"
        End If
        If Trim(Accion51.Text) <> "" And Estado5.ListIndex = 0 Then
            ZEntra = "N"
        End If
        If Trim(Accion61.Text) <> "" And Estado6.ListIndex = 0 Then
            ZEntra = "N"
        End If
        
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
            
                If ZZEstado <= 3 Then
            
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CargaSac SET "
                    ZSql = ZSql + " Estado = " + "'" + "4" + "'"
                    ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                    ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                    ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                    spCargaSac = ZSql
                    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
                
            End If
            
        End If
        
        
        
        
        Call CmdLimpiar_Click
        Ano.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Ano.Text <> "" And Numero.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacIII"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacIII.RecordCount > 0 Then
        
            rstCargaSacIII.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                Sql1 = ZSql + "DELETE CargaSacIII"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSacIII = Sql1 + Sql2
                Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
                
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
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    
    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
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
    
    Responsable1.Locked = False
    Responsable2.Locked = False
    Responsable3.Locked = False
    Responsable4.Locked = False
    Responsable5.Locked = False
    Responsable6.Locked = False
    
    Fecha1.Enabled = True
    Fecha2.Enabled = True
    Fecha3.Enabled = True
    Fecha4.Enabled = True
    Fecha5.Enabled = True
    Fecha6.Enabled = True
    
    Comentario11.Locked = False
    Comentario12.Locked = False
    Comentario21.Locked = False
    Comentario22.Locked = False
    Comentario31.Locked = False
    Comentario32.Locked = False
    Comentario41.Locked = False
    Comentario42.Locked = False
    Comentario51.Locked = False
    Comentario52.Locked = False
    Comentario61.Locked = False
    Comentario62.Locked = False

    Responsable1.Locked = False
    Responsable2.Locked = False
    Responsable3.Locked = False
    Responsable4.Locked = False
    Responsable5.Locked = False
    Responsable6.Locked = False
    
    Fecha1.Enabled = True
    Fecha2.Enabled = True
    Fecha3.Enabled = True
    Fecha4.Enabled = True
    Fecha5.Enabled = True
    Fecha6.Enabled = True
    
    Comentario11.Locked = False
    Comentario12.Locked = False
    Comentario21.Locked = False
    Comentario22.Locked = False
    Comentario31.Locked = False
    Comentario32.Locked = False
    Comentario41.Locked = False
    Comentario42.Locked = False
    Comentario51.Locked = False
    Comentario52.Locked = False
    Comentario61.Locked = False
    Comentario62.Locked = False
    
    Estado1.Enabled = True
    Estado2.Enabled = True
    Estado3.Enabled = True
    Estado4.Enabled = True
    Estado5.Enabled = True
    Estado6.Enabled = True
    
    Tipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    PrgCargaSacImplementa.Hide
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
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""
    
    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
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
    
    Estado1.AddItem ""
    Estado1.AddItem "Imple."
    Estado1.AddItem "Nula"
    
    Estado1.ListIndex = 0
    
    Estado2.Clear
    
    Estado2.AddItem ""
    Estado2.AddItem "Imple."
    Estado2.AddItem "Nula"
    
    Estado2.ListIndex = 0
    
    Estado3.Clear
    
    Estado3.AddItem ""
    Estado3.AddItem "Imple."
    Estado3.AddItem "Nula"
    
    Estado3.ListIndex = 0
    
    Estado4.Clear
    
    Estado4.AddItem ""
    Estado4.AddItem "Imple."
    Estado4.AddItem "Nula"
    
    Estado4.ListIndex = 0
    
    Estado5.Clear
    
    Estado5.AddItem ""
    Estado5.AddItem "Imple."
    Estado5.AddItem "Nula"
 
    Estado5.ListIndex = 0
    
    Estado6.Clear
    
    Estado6.AddItem ""
    Estado6.AddItem "Imple."
    Estado6.AddItem "Nula"
    
    Estado6.ListIndex = 0
    
    
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
                
                ZResponsableDestino = rstCargaSac!ResponsableDestino
                ZResponsableCentro = 9999
                
                ZCentro = rstCargaSac!Centro
                
                rstCargaSac.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CargaSacIII"
                ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
                spCargaSacIII = ZSql
                Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaSacIII.RecordCount > 0 Then
                    ZResponsable1 = rstCargaSacIII!Responsable1
                    ZResponsable2 = rstCargaSacIII!Responsable2
                    ZResponsable3 = rstCargaSacIII!Responsable3
                    ZResponsable4 = rstCargaSacIII!Responsable4
                    ZResponsable5 = rstCargaSacIII!Responsable5
                    ZResponsable6 = rstCargaSacIII!Responsable6
                    rstCargaSacIII.Close
                End If
                
                Sql1 = "Select *"
                Sql2 = " FROM CentroSac"
                Sql3 = " Where CentroSac.Codigo = " + "'" + Str$(ZCentro) + "'"
                spCentroSac = Sql1 + Sql2 + Sql3
                Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCentroSac.RecordCount > 0 Then
                    ZResponsableCentro = rstCentroSac!Responsable
                    rstCentroSac.Close
                End If
                
                If WOperador = ZResponsableDestino Or WOperador = ZResponsableCentro Or WOperador = ZResponsable1 Or WOperador = ZResponsable2 Or WOperador = ZResponsable3 Or WOperador = ZResponsable4 Or WOperador = ZResponsable5 Or WOperador = ZResponsable6 Or ZZCodigoResponsable = 1 Then
                    Call Imprime_Datos
                    Responsable1.SetFocus
                        Else
                    m$ = "No posee autorizacion para ingresar a actualizar esta SAC"
                    A% = MsgBox(m$, 0, "Archivo de Implementacion")
                    Call CmdLimpiar_Click
                End If
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Tipo.Text = WTipo
                Ano.Text = WAno
                Numero.Text = WNumero
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
                Rem Fecha.SetFocus
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

     Opcion.AddItem "Tipo"
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


