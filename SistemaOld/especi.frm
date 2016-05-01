VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspe 
   Caption         =   "Ingreso de Especificaciones de Productos Terminados"
   ClientHeight    =   8160
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11685
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
      MaxLength       =   50
      TabIndex        =   71
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
      MaxLength       =   50
      TabIndex        =   70
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
      MaxLength       =   50
      TabIndex        =   69
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
      MaxLength       =   50
      TabIndex        =   68
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
      MaxLength       =   50
      TabIndex        =   67
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
      MaxLength       =   50
      TabIndex        =   66
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
      MaxLength       =   50
      TabIndex        =   65
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
      MaxLength       =   50
      TabIndex        =   64
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
      MaxLength       =   50
      TabIndex        =   63
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
      MaxLength       =   50
      TabIndex        =   62
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
   Begin Crystal.CrystalReport lista 
      Left            =   11400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wespef.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   2880
      TabIndex        =   50
      Top             =   6600
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
      Left            =   240
      TabIndex        =   49
      Top             =   6720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   11040
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
      Left            =   5880
      MaxLength       =   50
      TabIndex        =   47
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
      MaxLength       =   50
      TabIndex        =   46
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
      MaxLength       =   50
      TabIndex        =   45
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
      MaxLength       =   50
      TabIndex        =   44
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
      MaxLength       =   50
      TabIndex        =   43
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
      MaxLength       =   50
      TabIndex        =   42
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
      MaxLength       =   50
      TabIndex        =   41
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
      MaxLength       =   50
      TabIndex        =   40
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
      MaxLength       =   50
      TabIndex        =   39
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
      MaxLength       =   50
      TabIndex        =   38
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
      Left            =   960
      TabIndex        =   13
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
      ItemData        =   "especi.frx":0000
      Left            =   120
      List            =   "especi.frx":0007
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   7575
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
      Top             =   6720
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
      Top             =   7680
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
      Top             =   7200
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
      Left            =   7800
      TabIndex        =   2
      Top             =   7200
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
      Visible         =   0   'False
      Width           =   975
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
      Height          =   255
      Left            =   120
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   720
      Width           =   4980
   End
   Begin VB.Label lblresultado 
      Alignment       =   2  'Center
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
      TabIndex        =   17
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblDescri 
      Alignment       =   2  'Center
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
      Width           =   4935
   End
   Begin VB.Label lblensayo 
      Alignment       =   2  'Center
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
Attribute VB_Name = "PrgEspe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecif As Recordset
Dim spEspecif As String
Dim XParam As String
Dim EmpresaActual As String

Private Sub Acepta_Click()
    
    lista.WindowTitle = "Listado de Ensayos"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height
    
    lista.GroupSelectionFormula = "{Especif.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    lista.SQLQuery = "SELECT Especif.Producto, Especif.Ensayo1, Especif.Valor1, Especif.Ensayo2, Especif.Valor2, Especif.Ensayo3, Especif.Valor3, Especif.Ensayo4, Especif.Valor4, Especif.Ensayo5, Especif.Valor5, Especif.Valor6, Especif.Ensayo7, Terminado.Descripcion " _
                     + "From " + DSQ + ".dbo.Especif Especif, " _
                     + DSQ + ".dbo.Terminado Terminado " _
                     + "Where Especif.Producto = Terminado.Codigo AND Especif.Producto >= ' ' AND Especif.Producto <= 'ZZ-ZZZZZ-ZZZ'"
    
    lista.DataFiles(2) = WEmpresa + "auxi.mdb"
    lista.Connect = Connect()
    
    lista.Action = 1
    
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub Imprime_Datos()

    spEspecif = "ConsultaEspecif " + "'" + Producto.Text + "'"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        Producto.Text = rstEspecif!Producto
        Ensayo1.Text = rstEspecif!Ensayo1
        Ensayo2.Text = rstEspecif!Ensayo2
        Ensayo3.Text = rstEspecif!Ensayo3
        Ensayo4.Text = rstEspecif!Ensayo4
        Ensayo5.Text = rstEspecif!Ensayo5
        Ensayo6.Text = rstEspecif!Ensayo6
        Ensayo7.Text = rstEspecif!Ensayo7
        Ensayo8.Text = rstEspecif!Ensayo8
        Ensayo9.Text = rstEspecif!Ensayo9
        Ensayo10.Text = rstEspecif!Ensayo10
        Valor1.Text = rstEspecif!Valor1
        valor2.Text = rstEspecif!valor2
        Valor3.Text = rstEspecif!Valor3
        valor4.Text = rstEspecif!valor4
        valor5.Text = rstEspecif!valor5
        valor6.Text = rstEspecif!valor6
        valor7.Text = rstEspecif!valor7
        valor8.Text = rstEspecif!valor8
        valor9.Text = rstEspecif!valor9
        valor10.Text = rstEspecif!valor10
        Valor11.Text = IIf(IsNull(rstEspecif!Valor11), "", rstEspecif!Valor11)
        Valor22.Text = IIf(IsNull(rstEspecif!Valor22), "", rstEspecif!Valor22)
        Valor33.Text = IIf(IsNull(rstEspecif!Valor33), "", rstEspecif!Valor33)
        Valor44.Text = IIf(IsNull(rstEspecif!Valor44), "", rstEspecif!Valor44)
        Valor55.Text = IIf(IsNull(rstEspecif!Valor55), "", rstEspecif!Valor55)
        Valor66.Text = IIf(IsNull(rstEspecif!Valor66), "", rstEspecif!Valor66)
        Valor77.Text = IIf(IsNull(rstEspecif!Valor77), "", rstEspecif!Valor77)
        Valor88.Text = IIf(IsNull(rstEspecif!Valor88), "", rstEspecif!Valor88)
        Valor99.Text = IIf(IsNull(rstEspecif!Valor99), "", rstEspecif!Valor99)
        Valor1010.Text = IIf(IsNull(rstEspecif!Valor1010), "", rstEspecif!Valor1010)
        
        rstEspecif.Close
                        
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
        
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            rstTerminado.Close
        End If
    End If

End Sub


Private Sub cmdAdd_Click()
    If Producto.Text <> "" Then
    
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
        WDate = Date$
    
        spEspecif = "ConsultaEspecif " + "'" + Producto.Text + "'"
        Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecif.RecordCount > 0 Then
            rstEspecif.Close
            
            Sql1 = "UPDATE Especif SET "
            Sql2 = "Producto = " + "'" + WProducto + "',"
            Sql3 = "Ensayo1 = " + "'" + WEnsayo1 + "',"
            Sql4 = "Valor1 = " + "'" + WValor1 + "',"
            Sql5 = "Ensayo2 = " + "'" + WEnsayo2 + "',"
            Sql6 = "Valor2 = " + "'" + WValor2 + "',"
            Sql7 = "Ensayo3 = " + "'" + WEnsayo3 + "',"
            Sql8 = "Valor3 = " + "'" + WValor3 + "',"
            Sql9 = "Ensayo4 = " + "'" + WEnsayo4 + "',"
            Sql10 = "Valor4 = " + "'" + WValor4 + "',"
            Sql11 = "Ensayo5 = " + "'" + WEnsayo5 + "',"
            Sql12 = "Valor5 = " + "'" + WValor5 + "',"
            Sql13 = "Ensayo6 = " + "'" + WEnsayo6 + "',"
            Sql14 = "Valor6 = " + "'" + WValor6 + "',"
            Sql15 = "Ensayo7 = " + "'" + WEnsayo7 + "',"
            Sql16 = "Valor7 = " + "'" + WValor7 + "',"
            Sql17 = "Ensayo8 = " + "'" + WEnsayo8 + "',"
            Sql18 = "Valor8 = " + "'" + WValor8 + "',"
            Sql19 = "Ensayo9 = " + "'" + WEnsayo9 + "',"
            Sql20 = "Valor9 = " + "'" + WValor9 + "',"
            Sql21 = "Ensayo10 = " + "'" + WEnsayo10 + "',"
            Sql22 = "Valor10 = " + "'" + WValor10 + "',"
            Sql23 = "WDate = " + "'" + WDate + "',"
            Sql24 = "Valor11 = " + "'" + WValor11 + "',"
            Sql25 = "Valor22 = " + "'" + WValor22 + "',"
            Sql26 = "Valor33 = " + "'" + WValor33 + "',"
            Sql27 = "Valor44 = " + "'" + WValor44 + "',"
            Sql28 = "Valor55 = " + "'" + WValor55 + "',"
            Sql29 = "Valor66 = " + "'" + WValor66 + "',"
            Sql30 = "Valor77 = " + "'" + WValor77 + "',"
            Sql31 = "Valor88 = " + "'" + WValor88 + "',"
            Sql32 = "Valor99 = " + "'" + WValor99 + "',"
            sql33 = "Valor1010 = " + "'" + WValor1010 + "'"
            sql34 = " Where Producto = " + "'" + WProducto + "'"
                     
            spEspecif = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                     + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                     + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                     + Sql31 + Sql32 + sql33 + sql34
            Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
            Sql1 = "INSERT INTO Especif ("
            Sql2 = "Producto, "
            Sql3 = "Ensayo1, Valor1, "
            Sql4 = "Ensayo2, Valor2, "
            Sql5 = "Ensayo3, Valor3, "
            Sql6 = "Ensayo4, Valor4, "
            Sql7 = "Ensayo5, Valor5, "
            Sql8 = "Ensayo6, Valor6, "
            Sql9 = "Ensayo7, Valor7, "
            Sql10 = "Ensayo8, Valor8, "
            Sql11 = "Ensayo9, Valor9, "
            Sql12 = "Ensayo10, Valor10, "
            Sql13 = "WDate, "
            Sql14 = "Valor11 , "
            Sql15 = "Valor22 , "
            Sql16 = "Valor33 , "
            Sql17 = "Valor44 , "
            Sql18 = "Valor55 , "
            Sql19 = "Valor66 , "
            Sql20 = "Valor77 , "
            Sql21 = "Valor88 , "
            Sql22 = "Valor99 , "
            Sql23 = "Valor1010) "
            Sql24 = "Values ("
            Sql25 = "'" + WProducto + "',"
            Sql26 = "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
            Sql27 = "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
            Sql28 = "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
            Sql29 = "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
            Sql30 = "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
            Sql31 = "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
            Sql32 = "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
            sql33 = "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
            sql34 = "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
            sql35 = "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
            sql36 = "'" + WDate + "',"
            sql37 = "'" + WValor11 + "',"
            sql38 = "'" + WValor22 + "',"
            sql39 = "'" + WValor33 + "',"
            sql40 = "'" + WValor44 + "',"
            sql41 = "'" + WValor55 + "',"
            sql42 = "'" + WValor66 + "',"
            sql43 = "'" + WValor77 + "',"
            sql44 = "'" + WValor88 + "',"
            sql45 = "'" + WValor99 + "',"
            sql46 = "'" + WValor1010 + "')"
            
            spEspecif = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                     + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                     + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 _
                     + Sql31 + Sql32 + sql33 + sql34 + sql35 + sql36 + sql37 + sql38 + sql39 + sql40 _
                     + sql41 + sql42 + sql43 + sql44 + sql45 + sql46
            Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        Call CmdLimpiar_Click
        Producto.SetFocus
    End If
    
End Sub

Private Sub cmdDelete_Click()
    If Producto.Text <> "" Then
        spEspecif = "ConsultaEspecif " + "'" + Producto.Text + "'"
        Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecif.RecordCount > 0 Then
            rstEspecif.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spEspecif = "BorrarEspecif " + "'" + Producto.Text + "'"
                Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Producto.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
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
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspe.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()
    spEspecif = "Anteriorespecif " + "'" + Producto.Text + "'"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        With rstEspecif
            .MoveLast
            Producto.Text = rstEspecif!Producto
            rstEspecif.Close
            Call Imprime_Datos
            Producto.SetFocus
        End With
    End If
End Sub

Private Sub Ensayo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor1.SetFocus
                    Else
            Descri1.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor2.SetFocus
                    Else
            descri2.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            Valor3.SetFocus
                    Else
            Descri3.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor4.SetFocus
                    Else
            Descri4.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor5.SetFocus
                    Else
            Descri5.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor6.SetFocus
                    Else
            Descri6.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor7.SetFocus
                    Else
            Descri7.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor8.SetFocus
                    Else
            Descri8.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor9.SetFocus
                    Else
            Descri9.Caption = ""
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            valor10.SetFocus
                    Else
            Descri10.Caption = ""
        End If
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
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEspe.Caption = "Ingreso de Especificaciones de Productos Terminados :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
End Sub




Private Sub Listado_Click()
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Frame2.Visible = True
End Sub


Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor11.SetFocus
    End If
End Sub

Private Sub Valor11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo2.SetFocus
    End If
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor22.SetFocus
    End If
End Sub

Private Sub Valor22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo3.SetFocus
    End If
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor33.SetFocus
    End If
End Sub

Private Sub Valor33_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo4.SetFocus
    End If
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor44.SetFocus
    End If
End Sub

Private Sub Valor44_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo5.SetFocus
    End If
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor55.SetFocus
    End If
End Sub

Private Sub Valor55_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo6.SetFocus
    End If
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor66.SetFocus
    End If
End Sub

Private Sub Valor66_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo7.SetFocus
    End If
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor77.SetFocus
    End If
End Sub

Private Sub Valor77_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo8.SetFocus
    End If
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor88.SetFocus
    End If
End Sub

Private Sub Valor88_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo9.SetFocus
    End If
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor99.SetFocus
    End If
End Sub

Private Sub Valor99_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo10.SetFocus
    End If
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1010.SetFocus
    End If
End Sub

Private Sub Valor1010_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo1.SetFocus
    End If
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            spEspecif = "ConsultaEspecif " + "'" + Producto.Text + "'"
            Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecif.RecordCount > 0 Then
                rstEspecif.Close
                Call Imprime_Datos
                    Else
                WProducto = Producto.Text
                CmdLimpiar_Click
                Producto.Text = WProducto
            End If
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                rstTerminado.Close
                    Else
                Producto.SetFocus
                Exit Sub
            End If
        End If
        Ensayo1.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Ensayos"
    
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
            spTerminado = "ListaterminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
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
            
        Case 1
            spEnsayo = "ListaEnsayos"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
            
            With rstEnsayo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                        Pantalla.AddItem IngresaItem
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            ClaveProd$ = WIndice.List(Indice)
            Producto.Text = ClaveProd$
            spEspecif = "ConsultaEspecif " + "'" + Producto.Text + "'"
            Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecif.RecordCount > 0 Then
                rstEspecif.Close
                Call Imprime_Datos
                    Else
                CmdLimpiar_Click
                Producto.Text = ClaveProd$
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    rstTerminado.Close
                        Else
                    Producto.SetFocus
                End If
            End If
            Producto.SetFocus
            
        Case 1
            Entra$ = "S"
            If Val(Ensayo1.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo1.Text = Val(WIndice.List(Indice))
                    Valor1.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri1.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            
            If Val(Ensayo2.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo2.Text = Val(WIndice.List(Indice))
                    valor2.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        descri2.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            
            If Val(Ensayo3.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo3.Text = Val(WIndice.List(Indice))
                    Valor3.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri3.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo4.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo4.Text = Val(WIndice.List(Indice))
                    valor4.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri4.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo5.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo5.Text = Val(WIndice.List(Indice))
                    valor5.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri5.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo6.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo6.Text = Val(WIndice.List(Indice))
                    valor6.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri6.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo7.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo7.Text = Val(WIndice.List(Indice))
                    valor7.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri7.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo8.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo8.Text = Val(WIndice.List(Indice))
                    valor8.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri8.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo9.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo9.Text = Val(WIndice.List(Indice))
                    valor9.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri9.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
            If Val(Ensayo10.Text) = 0 And Entra$ = "S" Then
                    Indice = Pantalla.ListIndex
                    Ensayo10.Text = Val(WIndice.List(Indice))
                    valor10.SetFocus
                    Entra$ = "N"
                    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayo.RecordCount > 0 Then
                        Descri10.Caption = rstEnsayo!Descripcion
                        rstEnsayo.Close
                    End If
            End If
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    spEspecif = "ListaEspecif"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        With rstEspecif
            .MoveFirst
            Producto.Text = rstEspecif!Producto
        End With
        rstEspecif.Close
        Call Imprime_Datos
        Producto.SetFocus
    End If
 End Sub

Private Sub Ultimo_Click()
    spEspecif = "ListaEspecif"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        With rstEspecif
            .MoveLast
            Producto.Text = rstEspecif!Producto
        End With
        rstEspecif.Close
        Call Imprime_Datos
        Producto.SetFocus
    End If

 End Sub

Private Sub Siguiente_Click()

    spEspecif = "PosteriorEspecif " + "'" + Producto.Text + "'"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        With rstEspecif
            .MoveFirst
            Producto.Text = rstEspecif!Producto
        End With
        rstEspecif.Close
        Call Imprime_Datos
        Producto.SetFocus
    End If
End Sub

