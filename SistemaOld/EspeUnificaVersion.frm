VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspeUnificaVersion 
   Caption         =   "Consulta de Versiones de Especificaciones de Producto Terminado"
   ClientHeight    =   8160
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   13800
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   13800
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   81
      Text            =   " "
      Top             =   6120
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   80
      Text            =   " "
      Top             =   6120
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   79
      Text            =   " "
      Top             =   5520
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   78
      Text            =   " "
      Top             =   5520
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   77
      Text            =   " "
      Top             =   4920
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   76
      Text            =   " "
      Top             =   4920
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   75
      Text            =   " "
      Top             =   4320
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   74
      Text            =   " "
      Top             =   4320
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   73
      Text            =   " "
      Top             =   3720
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   72
      Text            =   " "
      Top             =   3720
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   71
      Text            =   " "
      Top             =   3120
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   70
      Text            =   " "
      Top             =   3120
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   69
      Text            =   " "
      Top             =   2520
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   68
      Text            =   " "
      Top             =   2520
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   67
      Text            =   " "
      Top             =   1920
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   66
      Text            =   " "
      Top             =   1920
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   65
      Text            =   " "
      Top             =   1200
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   64
      Text            =   " "
      Top             =   1200
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
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   61
      Text            =   " "
      Top             =   720
      Width           =   840
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
      Left            =   11760
      MaxLength       =   8
      TabIndex        =   60
      Text            =   " "
      Top             =   720
      Width           =   840
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
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   58
      Text            =   " "
      Top             =   6720
      Width           =   5760
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8160
      TabIndex        =   57
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Version 
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
      MaxLength       =   50
      TabIndex        =   56
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox FechaFinal 
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
      Left            =   7800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   55
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox FechaInicio 
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
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   52
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   51
      Text            =   " "
      Top             =   6360
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   50
      Text            =   " "
      Top             =   5760
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   49
      Text            =   " "
      Top             =   5160
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   4560
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      Top             =   3960
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   3360
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      Top             =   2760
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      Top             =   2160
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   1560
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   960
      Width           =   5655
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
      Left            =   9840
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEspefUnifica.rpt"
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
      Height          =   1260
      Left            =   240
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   10560
      TabIndex        =   39
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      Top             =   6120
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   37
      Text            =   " "
      Top             =   5520
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   36
      Text            =   " "
      Top             =   4920
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   4320
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   3720
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   3120
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   32
      Text            =   " "
      Top             =   2520
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   1920
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   1250
      Width           =   5655
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   720
      Width           =   5655
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   28
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   27
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   26
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   25
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   24
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   23
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   22
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   21
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   20
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   19
      Text            =   " "
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   6840
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
      Height          =   1260
      ItemData        =   "EspeUnificaVersion.frx":0000
      Left            =   120
      List            =   "EspeUnificaVersion.frx":0007
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   7575
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
      Height          =   540
      Left            =   8160
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
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
      Height          =   540
      Left            =   9600
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
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
      Left            =   12720
      TabIndex        =   63
      Top             =   360
      Width           =   840
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
      Left            =   11760
      TabIndex        =   62
      Top             =   360
      Width           =   840
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
      TabIndex        =   59
      Top             =   6720
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   54
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
      Left            =   5640
      TabIndex        =   53
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
      TabIndex        =   41
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
      TabIndex        =   18
      Top             =   6120
      Width           =   4980
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
      TabIndex        =   17
      Top             =   5520
      Width           =   4980
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
      TabIndex        =   16
      Top             =   4920
      Width           =   4980
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
      TabIndex        =   15
      Top             =   4320
      Width           =   4980
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
      TabIndex        =   14
      Top             =   3720
      Width           =   4980
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
      TabIndex        =   13
      Top             =   3120
      Width           =   4980
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
      TabIndex        =   12
      Top             =   2520
      Width           =   4980
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
      TabIndex        =   11
      Top             =   1920
      Width           =   4980
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
      TabIndex        =   10
      Top             =   1320
      Width           =   4980
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
      TabIndex        =   9
      Top             =   720
      Width           =   4980
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
      Left            =   5880
      TabIndex        =   8
      Top             =   360
      Width           =   5655
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
      TabIndex        =   7
      Top             =   360
      Width           =   4935
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
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgEspeUnificaVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim EspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim XParam As String
Dim EmpresaActual As String
Dim ZFecha As String
Dim ZVersion As String
Dim CargaEmpresa(12, 2) As String

Private Sub Imprime_Datos()

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
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecifUnificaVersion"
    Sql3 = " Where EspecifUnificaVersion.Producto = " + "'" + Producto.Text + "'"
    Sql4 = " and EspecifUnificaVersion.Version = " + "'" + Version.Text + "'"
    spEspecifUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
    Set EspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
    If EspecifUnificaVersion.RecordCount > 0 Then
        Ensayo1.Text = EspecifUnificaVersion!Ensayo1
        Ensayo2.Text = EspecifUnificaVersion!Ensayo2
        Ensayo3.Text = EspecifUnificaVersion!Ensayo3
        Ensayo4.Text = EspecifUnificaVersion!Ensayo4
        Ensayo5.Text = EspecifUnificaVersion!Ensayo5
        Ensayo6.Text = EspecifUnificaVersion!Ensayo6
        Ensayo7.Text = EspecifUnificaVersion!Ensayo7
        Ensayo8.Text = EspecifUnificaVersion!Ensayo8
        Ensayo9.Text = EspecifUnificaVersion!Ensayo9
        Ensayo10.Text = EspecifUnificaVersion!Ensayo10
        Valor1.Text = EspecifUnificaVersion!Valor1
        valor2.Text = EspecifUnificaVersion!valor2
        Valor3.Text = EspecifUnificaVersion!Valor3
        valor4.Text = EspecifUnificaVersion!valor4
        valor5.Text = EspecifUnificaVersion!valor5
        valor6.Text = EspecifUnificaVersion!valor6
        valor7.Text = EspecifUnificaVersion!valor7
        valor8.Text = EspecifUnificaVersion!valor8
        valor9.Text = EspecifUnificaVersion!valor9
        valor10.Text = EspecifUnificaVersion!valor10
        Valor11.Text = IIf(IsNull(EspecifUnificaVersion!Valor11), "", EspecifUnificaVersion!Valor11)
        Valor22.Text = IIf(IsNull(EspecifUnificaVersion!Valor22), "", EspecifUnificaVersion!Valor22)
        Valor33.Text = IIf(IsNull(EspecifUnificaVersion!Valor33), "", EspecifUnificaVersion!Valor33)
        Valor44.Text = IIf(IsNull(EspecifUnificaVersion!Valor44), "", EspecifUnificaVersion!Valor44)
        Valor55.Text = IIf(IsNull(EspecifUnificaVersion!Valor55), "", EspecifUnificaVersion!Valor55)
        Valor66.Text = IIf(IsNull(EspecifUnificaVersion!Valor66), "", EspecifUnificaVersion!Valor66)
        Valor77.Text = IIf(IsNull(EspecifUnificaVersion!Valor77), "", EspecifUnificaVersion!Valor77)
        Valor88.Text = IIf(IsNull(EspecifUnificaVersion!Valor88), "", EspecifUnificaVersion!Valor88)
        Valor99.Text = IIf(IsNull(EspecifUnificaVersion!Valor99), "", EspecifUnificaVersion!Valor99)
        Valor1010.Text = IIf(IsNull(EspecifUnificaVersion!Valor1010), "", EspecifUnificaVersion!Valor1010)
        
        Desde1.Text = EspecifUnificaVersion!Desde1
        Desde2.Text = EspecifUnificaVersion!Desde2
        Desde3.Text = EspecifUnificaVersion!Desde3
        Desde4.Text = EspecifUnificaVersion!Desde4
        Desde5.Text = EspecifUnificaVersion!Desde5
        Desde6.Text = EspecifUnificaVersion!Desde6
        Desde7.Text = EspecifUnificaVersion!Desde7
        Desde8.Text = EspecifUnificaVersion!Desde8
        Desde9.Text = EspecifUnificaVersion!Desde9
        Desde10.Text = EspecifUnificaVersion!Desde10
        
        Hasta1.Text = EspecifUnificaVersion!Hasta1
        Hasta2.Text = EspecifUnificaVersion!Hasta2
        Hasta3.Text = EspecifUnificaVersion!Hasta3
        Hasta4.Text = EspecifUnificaVersion!Hasta4
        Hasta5.Text = EspecifUnificaVersion!Hasta5
        Hasta6.Text = EspecifUnificaVersion!Hasta6
        Hasta7.Text = EspecifUnificaVersion!Hasta7
        Hasta8.Text = EspecifUnificaVersion!Hasta8
        Hasta9.Text = EspecifUnificaVersion!Hasta9
        Hasta10.Text = EspecifUnificaVersion!Hasta10
        
        FechaInicio.Text = EspecifUnificaVersion!FechaInicio
        FechaFinal.Text = EspecifUnificaVersion!FechaFinal
        ControlCambio.Text = IIf(IsNull(EspecifUnificaVersion!ControlCambio), "", EspecifUnificaVersion!ControlCambio)
        
        EspecifUnificaVersion.Close
        
    End If
    
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
    
    Call Conecta_Empresa
        
End Sub


Private Sub CmdLimpiar_Click()

    Producto.Text = "  -     -   "
    Version.Text = ""
    
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
    
    FechaInicio.Text = ""
    FechaFinal.Text = ""
    
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspeUnificaVersion.Hide
    Unload Me
    Menu.Show
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
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgEspeUnificaVersion.Caption = PrgEspeUnifica.Caption + ":  " + !Nombre
        End If
    End With
        
    EmpresaActual = Wempresa
    
End Sub

Private Sub Listado_Click()

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
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZDescriprod = ""
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZZDescriprod = rstTerminado!Descripcion
        rstTerminado.Close
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
    ZSql = ZSql + "'" + Producto.Text + "',"
    ZSql = ZSql + "'" + ZZDescriprod + "',"
    ZSql = ZSql + "'" + Ensayo1.Text + "',"
    ZSql = ZSql + "'" + Ensayo2.Text + "',"
    ZSql = ZSql + "'" + Ensayo3.Text + "',"
    ZSql = ZSql + "'" + Ensayo4.Text + "',"
    ZSql = ZSql + "'" + Ensayo5.Text + "',"
    ZSql = ZSql + "'" + Ensayo6.Text + "',"
    ZSql = ZSql + "'" + Ensayo7.Text + "',"
    ZSql = ZSql + "'" + Ensayo8.Text + "',"
    ZSql = ZSql + "'" + Ensayo9.Text + "',"
    ZSql = ZSql + "'" + Ensayo10.Text + "',"
    ZSql = ZSql + "'" + Descri1.Caption + "',"
    ZSql = ZSql + "'" + descri2.Caption + "',"
    ZSql = ZSql + "'" + Descri3.Caption + "',"
    ZSql = ZSql + "'" + Descri4.Caption + "',"
    ZSql = ZSql + "'" + Descri5.Caption + "',"
    ZSql = ZSql + "'" + Descri6.Caption + "',"
    ZSql = ZSql + "'" + Descri7.Caption + "',"
    ZSql = ZSql + "'" + Descri8.Caption + "',"
    ZSql = ZSql + "'" + Descri9.Caption + "',"
    ZSql = ZSql + "'" + Descri10.Caption + "',"
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
    ZSql = ZSql + "'" + Valor11.Text + "',"
    ZSql = ZSql + "'" + Valor22.Text + "',"
    ZSql = ZSql + "'" + Valor33.Text + "',"
    ZSql = ZSql + "'" + Valor44.Text + "',"
    ZSql = ZSql + "'" + Valor55.Text + "',"
    ZSql = ZSql + "'" + Valor66.Text + "',"
    ZSql = ZSql + "'" + Valor77.Text + "',"
    ZSql = ZSql + "'" + Valor88.Text + "',"
    ZSql = ZSql + "'" + Valor99.Text + "',"
    ZSql = ZSql + "'" + Valor1010.Text + "',"
    ZSql = ZSql + "'" + Version.Text + "',"
    ZSql = ZSql + "'" + FechaInicio.Text + "',"
    ZSql = ZSql + "'" + FechaFinal.Text + "')"
    
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
    
    Lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height

    Rem lista.GroupSelectionFormula = "{EspecificacionesUnifica.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Lista.Destination = 1
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(Wempresa) = 3 Then
        Lista.ReportFileName = "ListaEspePtVersion.rpt"
            Else
        Lista.ReportFileName = "ListaEspePtVersionPelli.rpt"
    End If
    
    Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha " _
                + "From " _
                + DSQ + ".dbo.ListaEspe ListaEspe " _
                + "Where " _
                + "ListaEspe.Codigo >= '" + Producto.Text + "' AND " _
                + "ListaEspe.Codigo <= '" + Producto.Text + "'"
    Lista.Connect = Connect()
    
    Lista.Action = 1
    
    Call Conecta_Empresa

End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
        
            Producto.Text = UCase(Producto.Text)
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                rstTerminado.Close
                Version.SetFocus
                    Else
                Producto.SetFocus
                Exit Sub
            End If
            
        End If
    End If
End Sub

Sub Version_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            Producto.Text = UCase(Producto.Text)
            
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
                
            Sql1 = "Select *"
            Sql2 = " FROM EspecifUnificaVersion"
            Sql3 = " Where EspecifUnificaVersion.Producto = " + "'" + Producto.Text + "'"
            Sql4 = " and EspecifUnificaVersion.Version = " + "'" + Version.Text + "'"
            spEspecifUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
            Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnificaVersion.RecordCount > 0 Then
                rstEspecifUnificaVersion.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                    Else
                XProducto = Producto.Text
                XVersion = Version.Text
                Call CmdLimpiar_Click
                Producto.Text = XProducto
                Version.Text = XVersion
                Call Conecta_Empresa
                Version.SetFocus
            End If
            
        End If
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub


