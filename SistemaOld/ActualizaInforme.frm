VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgActualizaInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Informe de Recepcion  de Materia Prima de Reventa"
   ClientHeight    =   8160
   ClientLeft      =   75
   ClientTop       =   540
   ClientWidth     =   11835
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTexto1 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4680
      TabIndex        =   63
      Top             =   3120
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   62
      Top             =   2760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   3960
      TabIndex        =   61
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Transito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      LinkTimeout     =   20
      MaxLength       =   50
      TabIndex        =   53
      Text            =   " "
      Top             =   480
      Width           =   3135
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5400
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   49
         Top             =   600
         Width           =   1215
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Wlote3 
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
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
         Left            =   1800
         TabIndex        =   46
         Top             =   600
         Width           =   975
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
         Left            =   1800
         TabIndex        =   45
         Top             =   960
         Width           =   975
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
         Left            =   1800
         TabIndex        =   44
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WLote4 
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   43
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox WLote5 
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   42
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox WCanti4 
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
         Left            =   1800
         TabIndex        =   41
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WCanti5 
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
         Left            =   1800
         TabIndex        =   40
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label17 
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
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label16 
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
         Left            =   1800
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Estado2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox Estado1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   7560
      TabIndex        =   37
      Top             =   840
      Width           =   135
   End
   Begin VB.TextBox Certificado2 
      BeginProperty Font 
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
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox Certificado1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   840
      Width           =   135
   End
   Begin VB.TextBox XOrden 
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
      Left            =   7320
      MaxLength       =   6
      TabIndex        =   30
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.TextBox Remito 
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
      Left            =   9360
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
      BeginProperty Font 
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
      MaxLength       =   11
      TabIndex        =   20
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11520
      Top             =   7200
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
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4440
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Informe 
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
      Left            =   1800
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
      TabIndex        =   11
      Top             =   6240
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
      TabIndex        =   10
      Top             =   6840
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
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   11655
      Begin VB.TextBox WEnvase 
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
         Left            =   8880
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
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
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   23
         Text            =   " "
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox WOrden 
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
         Height          =   300
         Left            =   360
         MaxLength       =   6
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envase"
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
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant.Ingresada"
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
         Left            =   7320
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label6 
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
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden"
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
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WDescripcion 
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
         Height          =   300
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   4695
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
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "ActualizaInforme.frx":0000
      Left            =   3360
      List            =   "ActualizaInforme.frx":0007
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   7215
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
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   54
      Top             =   3720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16777152
      ForeColor       =   4210752
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3735
      Left            =   120
      TabIndex        =   64
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   16777152
      ForeColor       =   4210752
   End
   Begin VB.Label Label9 
      Caption         =   "Nro.Transito"
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
      Left            =   7080
      TabIndex        =   52
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Estado Envases"
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
      Left            =   5880
      TabIndex        =   36
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Certif.de Analisis"
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
      TabIndex        =   33
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Orden Compra"
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
      TabIndex        =   29
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Remito"
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
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label DesProveedor 
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
      Left            =   3360
      TabIndex        =   16
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      TabIndex        =   15
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
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Informe"
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
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgActualizaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private Verifica(100, 2) As String
Private Entra As String
Private Auxiliar(100, 5) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim XParam As String
Private XLote(100, 17) As String
Dim XLote1 As String
Dim XCantiLote1 As String
Dim XLote2 As String
Dim XCantiLote2 As String
Dim XLote3 As String
Dim XCantiLote3 As String
Dim XLote4 As String
Dim XCantiLote4 As String
Dim XLote5 As String
Dim XCantiLote5 As String
Dim SaldoOrden As Double
Dim ZOrden As String
Dim ZArticulo As String
Dim WOrigen As String
Dim XXXVector(10000, 2) As String

Dim ZParidad As Double
Dim ZParidadII As Double
Dim ZCoeParidad As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String




Sub Verifica_datos()
    If Val(Remito.Text) = 0 Then
        Remito.Text = "0"
    End If
End Sub

Private Sub Borra_Click()

    WVector1.Col = 1
    WVector1.Text = ""
    
    WVector1.Col = 2
    WVector1.Text = ""

    WVector1.Col = 3
    WVector1.Text = ""
    
    WVector1.Col = 4
    WVector1.Text = ""
    
    WVector1.Col = 5
    WVector1.Text = ""
    
    WVector1.Col = 6
    WVector1.Text = ""
    
    WVector1.Col = 7
    WVector1.Text = ""
    
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WLinea.Text = ""
    WEnvase.Text = ""
    
    WLugar = WVector1.Row
    
    Verifica(WLugar, 1) = WArticulo.Text
    Verifica(WLugar, 2) = WOrden.Text
    
    XLote(WLugar, 1) = ""
    XLote(WLugar, 2) = ""
    XLote(WLugar, 3) = ""
    XLote(WLugar, 4) = ""
    XLote(WLugar, 5) = ""
    XLote(WLugar, 6) = ""
    XLote(WLugar, 7) = ""
    XLote(WLugar, 8) = ""
    XLote(WLugar, 9) = ""
    XLote(WLugar, 10) = ""
    
    WOrden.SetFocus
    
End Sub

Private Sub CierreAviso_Click()
    Aviso.Visible = False
    WOrden.SetFocus
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    Rem  With rstProveedor
    Rem     .Close
    Rem End With
    Rem With rstArticulo
    Rem     .Close
    Rem End With
    Rem With rstOrden
    Rem     .Close
    Rem End With
    Rem With rstInforme
    Rem     .Close
    Rem End With
    
    Rem DbsVentas.Close
    Rem DbsAdminis.Close
    Rem DbsCotiza.Close
    
    PrgActualizaInforme.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    Erase XXXVector
    Lugar = 0

    spLaudo = "ListaLaudoTotal"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                ZPArtiOri = Mid$(Str$(rstLaudo!Laudo), 2, 20)
                Lugar = Lugar + 1
                XXXVector(Lugar, 1) = rstLaudo!Clave
                XXXVector(Lugar, 2) = ZPArtiOri
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    For Ciclo = 1 To Lugar
          WClave = XXXVector(Ciclo, 1)
          WPartiOri = XXXVector(Ciclo, 2)
    
          XParam = "'" + WClave + "','" _
                        + WPartiOri + "'"
          spLaudo = "ModificaLaudoProceso2 " + XParam
          Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
End Sub

Private Sub Command2_Click()

    XLaudo = "950652"
    
    Sql1 = "DELETE Laudo"
    Sql2 = " Where Laudo = " + "'" + XLaudo + "'"
    spLaudo = Sql1 + Sql2
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Consulta_Click()

     Opcion.Clear
     Opcion.AddItem "Envases"
     Opcion.Visible = True
     
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
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstEnvases!Envases
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnvases.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Graba_Click()

    Rem On Error GoTo WError
    
    ver = WEmpresa
    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        rstInforme.Close
            Else
        Exit Sub
    End If
    
    If Val(WEmpresa) = 6 And Trim(Transito.Text) = "" Then
        m$ = "Es obligatorio informar el codigo de transito."
        a% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
        Exit Sub
    End If
    
    
    ZParidad = 0
    ZParidadII = 0
    ZCoeParidad = 1
    
    XEmpresa = WEmpresa
    
    Select Case Val(XEmpresa)
        Case 2, 4, 8, 9
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        ZParidad = rstCambios!Cambio
        ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
        If ZParidadII <> 0 And ZParidad <> 0 Then
            ZCoeParidad = ZParidadII / ZParidad
                Else
            ZCoeParidad = 1
        End If
        rstCambios.Close
        Call Conecta_Empresa
                Else
        m$ = "Se debe informar la paridad"
        G% = MsgBox(m$, 0, "Actualizaion de Informe de Recepcion de Materia Prima")
        Call Conecta_Empresa
        Exit Sub
    End If
    
    For iRow = 1 To 99
        
        WRow = iRow
        WVector1.Row = WRow
                
        WVector1.Col = 1
        Orden = WVector1.Text
               
        WVector1.Col = 2
        Articulo = UCase(WVector1.Text)
        
        WVector1.Col = 4
        CantidadTotal = Val(WVector1.Text)
        
        Lugar = WVector1.Row
        Marca = XLote(Lugar, 14)
        
        If Articulo <> "" And CantidadTotal > 0 Then
            If Marca <> "X" Then
                Exit Sub
            End If
        End If
            
    Next iRow

    Renglon = 0
    Erase Auxiliar
    
    For iRow = 1 To 99
    
        WRow = iRow
        WVector1.Row = WRow
                
        WVector1.Col = 1
        Orden = WVector1.Text
               
        WVector1.Col = 2
        Articulo = UCase(WVector1.Text)
        
        WVector1.Col = 4
        CantidadTotal = Val(WVector1.Text)
        
        If Articulo <> "" And CantidadTotal > 0 Then
                
            WLugar = WVector1.Row
                
            XLote1 = XLote(WLugar, 1)
            XLote2 = XLote(WLugar, 3)
            XLote3 = XLote(WLugar, 5)
            XLote4 = XLote(WLugar, 7)
            XLote5 = XLote(WLugar, 9)
            XCantiLote1 = XLote(WLugar, 2)
            XCantiLote2 = XLote(WLugar, 4)
            XCantiLote3 = XLote(WLugar, 6)
            XCantiLote4 = XLote(WLugar, 8)
            XCantiLote5 = XLote(WLugar, 10)
        
            Envase = Val(XLote(WLugar, 15))
            ZProcedencia = XLote(WLugar, 16)
            ZNroDespacho = XLote(WLugar, 17)
        
            For Ciclo = 2 To 10 Step 2
        
                If Val(XLote(WLugar, Ciclo)) <> 0 Then
                
                    WPartida = XLote(WLugar, Ciclo - 1)
                    WArti = Articulo
                    
                    WEntra = "N"
                    XParam = "'" + WPartida + "','" _
                             + WArti + "'"
                    spLaudo = "ListaLaudoArticuloPartiOri " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WEntra = "S"
                        WLaudoOriginal = rstLaudo!Laudo
                        rstLaudo.Close
                    End If
        
                    If WEntra = "N" Then
                    
                        Select Case Val(WEmpresa)
                            Case 1
                                WLaudo = "950000"
                                spLaudo = "ListaLaudoDy"
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "950000"
                                End If
                            
                            Case 8
                                WLaudo = "800000"
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Laudo"
                                ZSql = ZSql + " Where Laudo > 799999 and Laudo < 899999"
                                ZSql = ZSql + " Order by Laudo"
                                spLaudo = ZSql
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "800000"
                                End If
                            
                            Case Else
                                WLaudo = "970000"
                                spLaudo = "ListaLaudoDy"
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "970000"
                               End If
                        End Select
                
                        WPartida = XLote(WLugar, Ciclo - 1)
                        WCantidad = Val(XLote(WLugar, Ciclo))
    
                        WRenglon = "1"
                        WFecha = Fecha.Text
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WOrden = Orden
                        WArticulo = Articulo
                        WLiberada = Str$(WCantidad)
                        WDevuelta = "0"
                        WLote = WLaudo
                        WRechazo = ""
                        WActualiza = "N"
                        WMarca = ""
                        WInforme = Informe.Text
                        WSaldo = Str$(WCantidad)
                        WOrigenOri = WOrigen
                        WPartiOri = WPartida
                        WEnvase = Str$(Envase)
                        WTransito = Transito.Text
                        WSaldoTransito = Str$(WCantidad)
                
                        Auxi1 = Str$(WLaudo)
                        Call Ceros(Auxi1, 6)
                        Auxi2 = Str$(WRenglon)
                        Call Ceros(Auxi2, 2)
            
                        WClave = Auxi1 + Auxi2
                        WDate = Date$
                        
                        Sql1 = "INSERT INTO Laudo ("
                        Sql2 = "Clave ,"
                        Sql3 = "Laudo ,"
                        Sql4 = "Renglon ,"
                        Sql5 = "Fecha ,"
                        Sql6 = "FechaOrd ,"
                        Sql7 = "Articulo ,"
                        Sql8 = "Liberada ,"
                        Sql9 = "Devuelta ,"
                        Sql10 = "Orden ,"
                        Sql11 = "Marca ,"
                        Sql12 = "Lote ,"
                        Sql13 = "Rechazo ,"
                        Sql14 = "Informe ,"
                        Sql15 = "Actualiza ,"
                        Sql16 = "WDate ,"
                        Sql17 = "Saldo ,"
                        Sql18 = "Origen ,"
                        Sql19 = "PartiOri ,"
                        Sql20 = "Envase ,"
                        Sql21 = "Transito ,"
                        Sql22 = "SaldoTransito )"
                        Sql23 = "Values ("
                        Sql24 = "'" + WClave + "',"
                        Sql25 = "'" + WLaudo + "',"
                        Sql26 = "'" + WRenglon + "',"
                        Sql27 = "'" + WFecha + "',"
                        Sql28 = "'" + WFechaord + "',"
                        Sql29 = "'" + WArticulo + "',"
                        Sql30 = "'" + WLiberada + "',"
                        Sql31 = "'" + WDevuelta + "',"
                        Sql32 = "'" + WOrden + "',"
                        Sql33 = "'" + WMarca + "',"
                        Sql34 = "'" + WLote + "',"
                        Sql35 = "'" + WRechazo + "',"
                        Sql36 = "'" + WInforme + "',"
                        Sql37 = "'" + WActualiza + "',"
                        Sql38 = "'" + WDate + "',"
                        Sql39 = "'" + WSaldo + "',"
                        Sql40 = "'" + WOrigenOri + "',"
                        Sql41 = "'" + WPartiOri + "',"
                        Sql42 = "'" + WEnvase + "',"
                        Sql43 = "'" + WTransito + "',"
                        Sql44 = "'" + WSaldoTransito + "')"
    
                        spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                  Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                  Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                                  Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                                  Sql41 + Sql42 + Sql43 + Sql44
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Laudo SET "
                        ZSql = ZSql + "NroDespacho = " + "'" + ZNroDespacho + "',"
                        ZSql = ZSql + "Procedencia = " + "'" + ZProcedencia + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        
                        ZVencimiento = ""
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Informe"
                        ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
                        ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
                        spInforme = ZSql
                        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInforme.RecordCount > 0 Then
                            ZVencimiento = IIf(IsNull(rstInforme!FechaVencimiento), "  /  /    ", rstInforme!FechaVencimiento)
                            rstInforme.Close
                        End If
                        
                        If Trim(ZVencimiento) <> "" Then
                        
                            ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)

                            ZSql = ""
                            ZSql = ZSql + "UPDATE Laudo SET "
                            ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
                            ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
    
                            Else
                            
                        Select Case Val(WEmpresa)
                            Case 1
                                WLaudo = "950000"
                                spLaudo = "ListaLaudoDy"
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "950000"
                                End If
                            
                            Case 8
                                WLaudo = "800000"
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Laudo"
                                ZSql = ZSql + " Where Laudo > 799999 and Laudo < 899999"
                                ZSql = ZSql + " Order by Laudo.Laudo"
                                spLaudo = ZSql
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "800000"
                                End If
                            
                            Case Else
                                WLaudo = "970000"
                                spLaudo = "ListaLaudoDy"
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    With rstLaudo
                                        .MoveLast
                                        WLaudo = Str$(rstLaudo!Laudo + 1)
                                    End With
                                    rstLaudo.Close
                                        Else
                                    WLaudo = "970000"
                               End If
                        End Select
                        
                        
                        
                
                        WPartida = XLote(WLugar, Ciclo - 1)
                        WCantidad = Val(XLote(WLugar, Ciclo))
    
                        WRenglon = "1"
                        WFecha = Fecha.Text
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WOrden = Orden
                        WArticulo = Articulo
                        WLiberada = Str$(WCantidad)
                        WDevuelta = "0"
                        WLote = WLaudo
                        WRechazo = ""
                        WActualiza = "N"
                        WMarca = ""
                        WInforme = Informe.Text
                        WSaldo = "0"
                        WOrigenOri = WOrigen
                        WPartiOri = WPartida
                        WEnvase = Str$(Envase)
                        WTransito = Transito.Text
                        WSaldoTransito = Str$(WCantidad)
                
                        Auxi1 = Str$(WLaudo)
                        Call Ceros(Auxi1, 6)
                        Auxi2 = Str$(WRenglon)
                        Call Ceros(Auxi2, 2)
            
                        WClave = Auxi1 + Auxi2
                        WDate = Date$
                        
                        Sql1 = "INSERT INTO Laudo ("
                        Sql2 = "Clave ,"
                        Sql3 = "Laudo ,"
                        Sql4 = "Renglon ,"
                        Sql5 = "Fecha ,"
                        Sql6 = "FechaOrd ,"
                        Sql7 = "Articulo ,"
                        Sql8 = "Liberada ,"
                        Sql9 = "Devuelta ,"
                        Sql10 = "Orden ,"
                        Sql11 = "Marca ,"
                        Sql12 = "Lote ,"
                        Sql13 = "Rechazo ,"
                        Sql14 = "Informe ,"
                        Sql15 = "Actualiza ,"
                        Sql16 = "WDate ,"
                        Sql17 = "Saldo ,"
                        Sql18 = "Origen ,"
                        Sql19 = "PartiOri ,"
                        Sql20 = "Envase ,"
                        Sql21 = "Transito ,"
                        Sql22 = "SaldoTransito )"
                        Sql23 = "Values ("
                        Sql24 = "'" + WClave + "',"
                        Sql25 = "'" + WLaudo + "',"
                        Sql26 = "'" + WRenglon + "',"
                        Sql27 = "'" + WFecha + "',"
                        Sql28 = "'" + WFechaord + "',"
                        Sql29 = "'" + WArticulo + "',"
                        Sql30 = "'" + WLiberada + "',"
                        Sql31 = "'" + WDevuelta + "',"
                        Sql32 = "'" + WOrden + "',"
                        Sql33 = "'" + WMarca + "',"
                        Sql34 = "'" + WLote + "',"
                        Sql35 = "'" + WRechazo + "',"
                        Sql36 = "'" + WInforme + "',"
                        Sql37 = "'" + WActualiza + "',"
                        Sql38 = "'" + WDate + "',"
                        Sql39 = "'" + WSaldo + "',"
                        Sql40 = "'" + WOrigenOri + "',"
                        Sql41 = "'" + WPartiOri + "',"
                        Sql42 = "'" + WEnvase + "',"
                        Sql43 = "'" + WTransito + "',"
                        Sql44 = "'" + WSaldoTransito + "')"
    
                        spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                 Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                 Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                                 Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                                 Sql41 + Sql42 + Sql43 + Sql44
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Laudo SET "
                        ZSql = ZSql + "NroDespacho = " + "'" + ZNroDespacho + "',"
                        ZSql = ZSql + "Procedencia = " + "'" + ZProcedencia + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        
                        
                        
                        
                        ZVencimiento = ""
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Informe"
                        ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
                        ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
                        spInforme = ZSql
                        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInforme.RecordCount > 0 Then
                            ZVencimiento = IIf(IsNull(rstInforme!FechaVencimiento), "  /  /    ", rstInforme!FechaVencimiento)
                            rstInforme.Close
                        End If
                        
                        If Trim(ZVencimiento) <> "" Then
                        
                            ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)

                            ZSql = ""
                            ZSql = ZSql + "UPDATE Laudo SET "
                            ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
                            ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                            
                        
                        
                        
                        
                        
                        
                        
                        
                        Rem carga la cantidad al laudo
                        Rem anterior con la misma partida
                        Rem original
                        
                        XParam = "'" + Str$(WLaudoOriginal) + "','" _
                                + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            WSaldo = Str$(rstLaudo!Saldo + WCantidad)
                            WDate = Date$
                            rstLaudo.Close
                        
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                            
                    End If
                End If
            Next Ciclo
            
            WPrecio = 0
            For WDa = 1 To 40
                Auxi2 = Orden
                Call Ceros(Auxi2, 6)
                Auxi1 = WDa
                Call Ceros(Auxi1, 2)
                WClave = Auxi2 + Auxi1
                spOrden = "ConsultaOrden " + "'" + WClave + "'"
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    WMoneda = rstOrden!Moneda
                    WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
                    If Articulo = rstOrden!Articulo Then
                        WPrecio = rstOrden!Precio
                        WLiberada = Str$(rstOrden!Liberada + CantidadTotal)
                        WDevuelta = "0"
                        WFechaEntrega = Fecha.Text
                        WDate = Date$
                        rstOrden.Close
                        XParam = "'" + WClave + "','" _
                            + WLiberada + "','" _
                            + WDevuelta + "','" _
                            + WFechaEntrega + "','" _
                            + WDate + "'"
                        spOrden = "ModificaOrdenPrueba " + XParam
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            Else
                        rstOrden.Close
                    End If
                End If
            Next WDa

            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                WProducto = Articulo
                WLaboratorio = Str$(rstArticulo!Laboratorio - CantidadTotal)
    
                Select Case WTipoOrden
                    Case 1, 2
                        If WMoneda = 0 Then
                            XCosto1 = IIf(IsNull(rstArticulo!Costo1), "0", rstArticulo!Costo1)
                            XCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                            WCosto1 = Str$(XCosto1)
                            WCosto3 = Str$(XCosto3)
                                Else
                            XCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                            XCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                            WCosto1 = Str$(XCosto1)
                            WCosto3 = Str$(XCosto3)
                        End If
        
                    Case Else
                        XStock1 = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        If WMoneda = 0 Then
                            XCosto1 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                                Else
                            XCosto1 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                        End If
                        XCostoTotal1 = XStock1 * XCosto1
                
                        XStock2 = CantidadTotal
                        Select Case WMoneda
                            Case 2
                                XCosto2 = WPrecio * ZCoeParidad
                            Case Else
                                XCosto2 = WPrecio
                        End Select
                        XCostoTotal2 = XStock2 * XCosto2
                
                        XCosto = 0
                        XStock = XStock1 + XStock2
                        XCostoTotal = XCostoTotal1 + XCostoTotal2
                        If XStock <> 0 Then
                            XCosto = XCostoTotal / XStock
                        End If
            
                        Rem Call Redondeo(XCosto)
                
                        Select Case WMoneda
                            Case 2
                                WCosto1 = Str$(WPrecio * ZCoeParidad)
                            Case Else
                                WCosto1 = Str$(WPrecio)
                        End Select
                        WCosto3 = Str$(XCosto)
            
                End Select
    
                WEntradas = Str$(rstArticulo!Entradas + CantidadTotal)
                WDate = Date$
                rstArticulo.Close
    
                If WMoneda = 0 Or WMoneda = 2 Then
                    XParam = "'" + WProducto + "','" _
                        + WLaboratorio + "','" _
                        + WEntradas + "','" _
                        + WDate + "','" _
                        + WCosto1 + "','" _
                        + WCosto3 + "'"
                    spArticulo = "ModificaArticuloLaudoDolares " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    XParam = "'" + WProducto + "','" _
                        + WLaboratorio + "','" _
                        + WEntradas + "','" _
                        + WDate + "','" _
                        + WCosto1 + "','" _
                        + WCosto3 + "'"
                    spArticulo = "ModificaArticuloLaudoPesos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
    
            End If
            
        End If
        
    Next iRow
    
    Call Limpia_Click
    
    Exit Sub

WError:

    Resume Next
    
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEnvase.Text = ""
    
    WOrden.SetFocus
    
End Sub

Private Sub Limpia_Click()

    CargaLote.Visible = False
    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEnvase.Text = ""
    XOrden.Text = ""

    Informe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    Transito.Text = ""
    
    Certificado1.Value = 0
    Certificado2.Text = ""
    Estado1.Value = 0
    Estado2.Text = ""
    
    Call Limpia_Vector
    Erase Verifica
    
    Renglon = 0
    Informe.SetFocus

End Sub

Private Sub WOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WOrigen = IIf(IsNull(rstOrden!Origen), "", rstOrden!Origen)
            rstOrden.Close
            WArticulo.SetFocus
                Else
            WOrden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        Pasa = "N"
        spOrden = "ListaOrden " + "'" + WOrden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        If WArticulo.Text = rstOrden!Articulo Then
                            Pasa = "S"
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
        
        If Pasa = "S" Then
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                WCantidad.SetFocus
            End If
                        Else
            WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WEnvase.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEnvase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        spEnvase = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            rstEnvase.Close
            
            spArticulo = "ConsultaArticulo" + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                rstArticulo.Close
            End If
            
            WReventa = 1
            
            If WReventa = 1 Then
            
                ZOrden = WOrden.Text
                ZArticulo = WArticulo.Text
                Call Calcula_SaldoOrden
                If SaldoOrden >= Val(WCantidad.Text) Then
                
                    WLugar = WVector1.Row
                            
                    If Val(XLote(WLugar, 2)) <> 0 Then
                        WLote1.Text = XLote(WLugar, 1)
                        WCanti1.Text = XLote(WLugar, 2)
                            Else
                        WLote1.Text = ""
                        WCanti1.Text = ""
                    End If
                    If Val(XLote(WLugar, 4)) <> 0 Then
                        WLote2.Text = XLote(WLugar, 3)
                        WCanti2.Text = XLote(WLugar, 4)
                            Else
                        WLote2.Text = ""
                        WCanti2.Text = ""
                    End If
                    If Val(XLote(WLugar, 6)) <> 0 Then
                        WLote3.Text = XLote(WLugar, 5)
                        WCanti3.Text = XLote(WLugar, 6)
                            Else
                        WLote3.Text = ""
                        WCanti3.Text = ""
                    End If
                    If Val(XLote(WLugar, 8)) <> 0 Then
                        WLote4.Text = XLote(WLugar, 7)
                        WCanti4.Text = XLote(WLugar, 8)
                            Else
                        WLote4.Text = ""
                        WCanti4.Text = ""
                    End If
                    If Val(XLote(WLugar, 10)) <> 0 Then
                        WLote5.Text = XLote(WLugar, 9)
                        WCanti5.Text = XLote(WLugar, 10)
                            Else
                        WLote5.Text = ""
                        WCanti5.Text = ""
                    End If
                
                    CargaLote.Visible = True
                    WLote1.SetFocus
                
                        Else
                    
                    Mensaje = "La cantidad que se desea laudar supera el saldo del informe de recepcion"
                    Estilo = vbOKOnly + vbCritical + vbDefaultButton2    ' Define los botones.
                    Título = "Laudos de Liberacion de M.P. de Reventa"   ' Define el título.
                    Ctxt = 1000 ' Define el tema
                    Respuesta = MsgBox(Mensaje, Estilo, Título, Ayuda, Ctxt)
                
                End If
                
            End If
        End If
    End If
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti1.SetFocus
    End If
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = 0
        If WLote1.Text <> "" Then
            Compara = Compara + Val(WCanti1.Text)
        End If
        If WLote2.Text <> "" Then
            Compara = Compara + Val(WCanti2.Text)
        End If
        If WLote3.Text <> "" Then
            Compara = Compara + Val(WCanti3.Text)
        End If
        If WLote4.Text <> "" Then
            Compara = Compara + Val(WCanti4.Text)
        End If
        If WLote5.Text <> "" Then
            Compara = Compara + Val(WCanti5.Text)
        End If
        If Compara = Val(WCantidad.Text) And Val(WCanti1.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti2.SetFocus
    End If
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = 0
        If WLote1.Text <> "" Then
            Compara = Compara + Val(WCanti1.Text)
        End If
        If WLote2.Text <> "" Then
            Compara = Compara + Val(WCanti2.Text)
        End If
        If WLote3.Text <> "" Then
            Compara = Compara + Val(WCanti3.Text)
        End If
        If WLote4.Text <> "" Then
            Compara = Compara + Val(WCanti4.Text)
        End If
        If WLote5.Text <> "" Then
            Compara = Compara + Val(WCanti5.Text)
        End If
        If Compara = Val(WCantidad.Text) And Val(WCanti2.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti3.SetFocus
    End If
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = 0
        If WLote1.Text <> "" Then
            Compara = Compara + Val(WCanti1.Text)
        End If
        If WLote2.Text <> "" Then
            Compara = Compara + Val(WCanti2.Text)
        End If
        If WLote3.Text <> "" Then
            Compara = Compara + Val(WCanti3.Text)
        End If
        If WLote4.Text <> "" Then
            Compara = Compara + Val(WCanti4.Text)
        End If
        If WLote5.Text <> "" Then
            Compara = Compara + Val(WCanti5.Text)
        End If
        If Compara = Val(WCantidad.Text) And Val(WCanti3.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote4.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti4.SetFocus
    End If
End Sub

Private Sub WCanti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = 0
        If WLote1.Text <> "" Then
            Compara = Compara + Val(WCanti1.Text)
        End If
        If WLote2.Text <> "" Then
            Compara = Compara + Val(WCanti2.Text)
        End If
        If WLote3.Text <> "" Then
            Compara = Compara + Val(WCanti3.Text)
        End If
        If WLote4.Text <> "" Then
            Compara = Compara + Val(WCanti4.Text)
        End If
        If WLote5.Text <> "" Then
            Compara = Compara + Val(WCanti5.Text)
        End If
        If Compara = Val(WCantidad.Text) And Val(WCanti4.Text) = 0 Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCanti5.SetFocus
    End If
End Sub

Private Sub WCanti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compara = 0
        If WLote1.Text <> "" Then
            Compara = Compara + Val(WCanti1.Text)
        End If
        If WLote2.Text <> "" Then
            Compara = Compara + Val(WCanti2.Text)
        End If
        If WLote3.Text <> "" Then
            Compara = Compara + Val(WCanti3.Text)
        End If
        If WLote4.Text <> "" Then
            Compara = Compara + Val(WCanti4.Text)
        End If
        If WLote5.Text <> "" Then
            Compara = Compara + Val(WCanti5.Text)
        End If
        If Compara = Val(WCantidad.Text) Then
            CargaLote.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WOrden.SetFocus
                Else
            WLote1.SetFocus
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
            WEnvase.Text = WIndice.List(Indice)
            Call WEnvase_KeyPress(13)
            
        Case Else
        
    End Select
    
End Sub


Private Sub Form_Load()
    
    Call Limpia_Vector
 
    CargaLote.Visible = False
    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    WLinea.Text = ""
    WOrden.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEnvase.Text = ""
    XOrden.Text = ""
    Transito.Text = ""

    Informe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    
    Certificado1.Value = 0
    Certificado2.Text = ""
    Estado1.Value = 0
    Estado2.Text = ""

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgActualizaInforme.Caption = PrgActualizaInforme.Caption + "   : " + !Nombre
        End If
    End With
    
    Rem Informe.SetFocus
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Auxiliar
    
    spInforme = "ListaInforme " + "'" + Informe.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                
                    WVector1.Row = Renglon
                
                    WVector1.Col = 1
                    WVector1.Text = rstInforme!Orden
                
                    WVector1.Col = 2
                    WVector1.Text = rstInforme!Articulo
                    Auxi1 = rstInforme!Articulo
                
                    WVector1.Col = 4
                    Rem  By NAN
                    WVector1.Text = Pusing("###,###.##", Val(rstInforme!Cantidad))
                    WVector1.Text = Val(rstInforme!Cantidad)
                    WVector1.Col = 5
                    Rem WVECTOR1.Text = ""
                    WVector1.Text = rstInforme!Envase
                    
                    XLote1 = IIf(IsNull(rstInforme!lote1), "0", rstInforme!lote1)
                    XLote2 = IIf(IsNull(rstInforme!lote2), "0", rstInforme!lote2)
                    XLote3 = IIf(IsNull(rstInforme!lote3), "0", rstInforme!lote3)
                    XLote4 = IIf(IsNull(rstInforme!lote4), "0", rstInforme!lote4)
                    XLote5 = IIf(IsNull(rstInforme!lote5), "0", rstInforme!lote5)
                    XCantiLote1 = IIf(IsNull(rstInforme!Canti1), "0", rstInforme!Canti1)
                    XCantiLote2 = IIf(IsNull(rstInforme!Canti2), "0", rstInforme!Canti2)
                    XCantiLote3 = IIf(IsNull(rstInforme!Canti3), "0", rstInforme!Canti3)
                    XCantiLote4 = IIf(IsNull(rstInforme!Canti4), "0", rstInforme!Canti4)
                    XCantiLote5 = IIf(IsNull(rstInforme!Canti5), "0", rstInforme!Canti5)
                    
                    XLote(Renglon, 1) = XLote1
                    XLote(Renglon, 2) = XCantiLote1
                    XLote(Renglon, 3) = XLote2
                    XLote(Renglon, 4) = XCantiLote2
                    XLote(Renglon, 5) = XLote3
                    XLote(Renglon, 6) = XCantiLote3
                    XLote(Renglon, 7) = XLote4
                    XLote(Renglon, 8) = XCantiLote4
                    XLote(Renglon, 9) = XLote5
                    XLote(Renglon, 10) = XCantiLote5
                    
                    ZProcedencia = IIf(IsNull(rstInforme!Procedencia), "", rstInforme!Procedencia)
                    ZNroDespacho = IIf(IsNull(rstInforme!NroDespacho), "", rstInforme!NroDespacho)
                    
                    XLote(Renglon, 15) = rstInforme!Envase
                    XLote(Renglon, 16) = ZProcedencia
                    XLote(Renglon, 17) = ZNroDespacho
                    
                    Auxiliar(Renglon, 1) = rstInforme!Articulo
                    Auxiliar(Renglon, 2) = rstInforme!Envase
                    Auxiliar(Renglon, 3) = rstInforme!Orden
                    Auxiliar(Renglon, 4) = Str$(rstInforme!Cantidad)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
        WVector1.Row = Renglon
        
        spArticulo = "ConsultaArticulo " + "'" + Auxiliar(Renglon, 1) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.Col = 3
            WVector1.Text = rstArticulo!Descripcion
            WOrden.SetFocus
            rstArticulo.Close
        End If
        
        spEnvases = "ConsultaEnvases " + "'" + XLote(Renglon, 15) + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            WVector1.Col = 6
            WVector1.Text = rstEnvases!Descripcion
            rstEnvase.Close
                Else
            WVector1.Col = 6
            WVector1.Text = ""
        End If
        
        ZOrden = Auxiliar(Renglon, 3)
        ZArticulo = Auxiliar(Renglon, 1)
        
        Call Calcula_SaldoOrden
        
        aa = SaldoOrden
     
        Rem BY NAN
        WVector1.Col = 4
        WVector1.Text = Str$(SaldoOrden)
        If Left$(Auxiliar(Renglon, 1), 2) <> "DY" Then
            WVector1.Text = "0"
        End If
      
        Rem If SaldoOrden >= Val(Auxiliar(Renglon, 4)) Then
        
    Next Da

    WOrden.SetFocus
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Alta_Vector()

    Entra = "S"
    
    If Val(WLinea.Text) = 0 Then
        For Da = 1 To 100
            If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text Then
                Entra = "N"
                Exit For
            End If
        Next Da
            Else
        Lugar = WVector1.Row
        For Da = 1 To 100
            If Verifica(Da, 1) = WArticulo.Text And Verifica(Da, 2) = WOrden.Text And Da <> Lugar Then
                Entra = "N"
                Exit For
            End If
        Next Da
    End If
    
    If Entra = "N" Then
        m$ = "El articulo ya se encuentra dado de alta en el informe de recepcion"
        a% = MsgBox(m$, 0, "Ingreso de Informe de recepcion")
    End If
                
    If Entra = "S" Then

        If Val(WLinea.Text) = 0 Then
        
            Renglon = Renglon + 1
            
            WVector1.Row = Renglon
            WAnterior = WVector1.Row
            
            WVector1.Col = 1
            WVector1.Text = WOrden.Text
            
            WVector1.Col = 2
            WVector1.Text = WArticulo.Text
            
            WVector1.Col = 3
            WVector1.Text = WDescripcion.Caption
                
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", WCantidad.Text)
                
            WVector1.Col = 5
            WVector1.Text = WEnvase.Text
            
            WVector1.Col = 7
            WVector1.Text = "X"
            
            spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WVector1.Col = 6
                WVector1.Text = rstEnvases!Descripcion
                rstEnvases.Close
                    Else
                WVector1.Col = 6
                WVector1.Text = ""
            End If
            
            Verifica(Renglon, 1) = WArticulo.Text
            Verifica(Renglon, 2) = WOrden.Text
            
            XLote(Renglon, 1) = WLote1.Text
            XLote(Renglon, 2) = WCanti1.Text
            XLote(Renglon, 3) = WLote2.Text
            XLote(Renglon, 4) = WCanti2.Text
            XLote(Renglon, 5) = WLote3.Text
            XLote(Renglon, 6) = WCanti3.Text
            XLote(Renglon, 7) = WLote4.Text
            XLote(Renglon, 8) = WCanti4.Text
            XLote(Renglon, 9) = WLote5.Text
            XLote(Renglon, 10) = WCanti5.Text
            XLote(Renglon, 14) = "X"
            XLote(Renglon, 15) = WEnvase.Text
            
                Else
                
            WVector1.Row = Val(WLinea.Text)
            WAnterior = WVector1.Row
            
            WVector1.Col = 1
            WVector1.Text = WOrden.Text
            
            WVector1.Col = 2
            WVector1.Text = WArticulo.Text
            
            WVector1.Col = 3
            WVector1.Text = WDescripcion.Caption
                
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", WCantidad.Text)
            
            WVector1.Col = 5
            WVector1.Text = WEnvase.Text
            
            WVector1.Col = 7
            WVector1.Text = "X"
            
            spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WVector1.Col = 6
                WVector1.Text = rstEnvases!Descripcion
                rstEnvases.Close
                    Else
                WVector1.Col = 6
                WVector1.Text = ""
            End If
            
            Lugar = WVector1.Row
            Verifica(Lugar, 1) = WArticulo.Text
            Verifica(Lugar, 2) = WOrden.Text
            
            XLote(Lugar, 1) = WLote1.Text
            XLote(Lugar, 2) = WCanti1.Text
            XLote(Lugar, 3) = WLote2.Text
            XLote(Lugar, 4) = WCanti2.Text
            XLote(Lugar, 5) = WLote3.Text
            XLote(Lugar, 6) = WCanti3.Text
            XLote(Lugar, 7) = WLote4.Text
            XLote(Lugar, 8) = WCanti4.Text
            XLote(Lugar, 9) = WLote5.Text
            XLote(Lugar, 10) = WCanti5.Text
            XLote(Lugar, 14) = "X"
            XLote(Lugar, 15) = WEnvase.Text
            
        End If
    
    End If

End Sub

Private Sub Informe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entra = "N"
        spInforme = "ListaInforme " + "'" + Informe.Text + "'"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            Entra = "S"
            Fecha.Text = rstInforme!Fecha
            Proveedor.Text = rstInforme!Proveedor
            Remito.Text = rstInforme!Remito
            XOrden.Text = rstInforme!Orden
            Certificado1.Value = IIf(IsNull(rstInforme!Certificado1), "0", rstInforme!Certificado1)
            Certificado2.Text = IIf(IsNull(rstInforme!Certificado2), "", rstInforme!Certificado2)
            Estado1.Value = IIf(IsNull(rstInforme!Estado1), "0", rstInforme!Estado1)
            Estado2.Text = IIf(IsNull(rstInforme!Estado2), "", rstInforme!Estado2)
            rstInforme.Close
        End If
        
        If Entra = "S" Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                DesProveedor.Caption = RstProveedor!Nombre
                RstProveedor.Close
            End If
            Call Proceso_Click
                Else
            Informe.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            XOrden.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WOrden.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Calcula_SaldoOrden()

    WRecibida = 0
    WLaudada = 0

    spInforme = "ListaInformeOrden " + "'" + ZOrden + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If ZArticulo = rstInforme!Articulo Then
                    WRecibida = WRecibida + rstInforme!Cantidad
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    spLaudo = "ListaLaudoOrden " + "'" + ZOrden + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If ZArticulo = rstLaudo!Articulo Then
                    WCantidad1 = rstLaudo!Liberada
                    WCantidad2 = rstLaudo!devuelta
                    WLaudada = WLaudada + WCantidad1 + WCantidad2
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    SaldoOrden = WRecibida - WLaudada

End Sub




Rem
Rem Controles de la grilla
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub


Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    WVector1.Col = 1
    WOrden.Text = WVector1.Text

    WVector1.Col = 2
    If Len(WVector1.Text) = 10 Then
        WLinea.Text = WVector1.Row
        WArticulo.Text = WVector1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    WVector1.Col = 3
    WDescripcion.Caption = WVector1.Text

    WVector1.Col = 4
    If Val(WVector1.Text) <> 0 Then
        WCantidad.Text = WVector1.Text
            Else
        WCantidad.Text = ""
    End If

    WVector1.Col = 5
    WEnvase.Text = WVector1.Text
    
    WOrden.SetFocus
    Rem StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 99
        
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Orden"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Envase"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Ok"
                WVector1.ColWidth(Ciclo) = 400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub




