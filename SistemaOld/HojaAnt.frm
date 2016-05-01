VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHojaAnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Hoja de Produccion"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   Visible         =   0   'False
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
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   -120
      Visible         =   0   'False
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   240
      Visible         =   0   'False
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   -120
      Visible         =   0   'False
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   -120
      Visible         =   0   'False
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   240
      Visible         =   0   'False
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
      Index           =   6
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   240
      Visible         =   0   'False
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
      Index           =   7
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   615
      Left            =   10080
      TabIndex        =   59
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   6480
      TabIndex        =   45
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   1440
         Width           =   855
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
         TabIndex        =   48
         Top             =   720
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
         TabIndex        =   47
         Top             =   1080
         Width           =   375
      End
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
         TabIndex        =   56
         Top             =   360
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
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   11040
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HojaAnt.frx":0000
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
      TabIndex        =   43
      Top             =   7680
      Width           =   975
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
      TabIndex        =   42
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox observaciones 
      BeginProperty Font 
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
      TabIndex        =   41
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   6120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra1 
      Height          =   1815
      Left            =   9240
      TabIndex        =   38
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   100
      Cols            =   4
   End
   Begin VB.TextBox Stock 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   3960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   2775
      Left            =   9240
      TabIndex        =   36
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   100
      Cols            =   4
   End
   Begin VB.CommandButton Reimpresion 
      Caption         =   "ReImpre."
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
      TabIndex        =   35
      Top             =   7680
      Width           =   975
   End
   Begin MSMask.MaskEdBox fechaIng 
      Height          =   285
      Left            =   8280
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
      Left            =   5160
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
      Left            =   7200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreHoja.rpt"
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
      Height          =   495
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
      Left            =   5160
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
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
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
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   9015
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
         Locked          =   -1  'True
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
         Enabled         =   0   'False
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
         Locked          =   -1  'True
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
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "HojaAnt.frx":007C
      TabIndex        =   9
      Top             =   1200
      Width           =   9015
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   0
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
      ItemData        =   "HojaAnt.frx":0A62
      Left            =   3360
      List            =   "HojaAnt.frx":0A69
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
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Left            =   6840
      TabIndex        =   40
      Top             =   480
      Width           =   1215
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
      Left            =   6840
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
      Left            =   3360
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
      Left            =   3600
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   3360
      TabIndex        =   19
      Top             =   120
      Width           =   1215
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
      Width           =   1695
   End
End
Attribute VB_Name = "PrgHojaAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.
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
Private XCantidad As Double
Private Auxiliar(100, 7) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposision As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPrecio As Recordset
Dim spPrecio As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstImpreHoja As Recordset
Dim spImpreHoja As String
Dim XParam As String
Dim LeeHoja As String
Dim XSaldo As Double
Dim Impre(3, 2) As Double
Dim WCosto1 As String
Dim WCosto2 As String
Dim WCosto3 As String
Private ZLote(100, 7) As String
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
Dim ZProceso As Integer
Dim LoteBusqueda As String
Dim ZSaldo As Double

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
    Wlote3.Text = ""
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
    
End Sub

Private Sub cmdClose_Click()

    LeeHoja = "N"
    Call Limpia_Click
    LeeHoja = "S"

    With rstEtiqueta
        .Close
    End With
    
    PrgHojaAnt.Hide
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
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa
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
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
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
            
        Case 1
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_DBLCLICK()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.FirstRow + DBGrid1.Row + 1
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
    
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.FirstRow + DBGrid1.Row + 1
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
    
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.Close
        m$ = "Partida ya existente"
        G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If

    If Val(Real.Text) = 0 Then
        Real.Text = "0"
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
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                If Real <> 0 Then
                    WEntradas = Str$(rstTerminado!Entradas - Val(Real))
                    WProceso = Str$(rstTerminado!Proceso)
                        Else
                    WProceso = Str$(rstTerminado!Proceso - Val(Teorico))
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
                        WSalidas = Str$(rstArticulo!Salidas - Val(Cantidad))
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                                           
                        spArticulo = "ModificaArticuloSalidas " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                                        
                Case "T"
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = rstTerminado!Codigo
                        WSalidas = Str$(rstTerminado!Salidas - Val(Cantidad))
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
        End If
            
    Next Da
    
    spHoja = "BorrarHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)
    
    Entra = "S"
    
    For a = 0 To 3
        
            Suma = a * 10
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
                    
                If Articulo <> "" Then
                        
                    Select Case Tipo
                        Case "T"
                            spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                                WImpre1 = Terminado
                                rstTerminado.Close
                            End If
                        Case "M"
                            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                                WImpre1 = Articulo
                                rstArticulo.Close
                            End If
                        Case Else
                    End Select
                
                    If Val(Cantidad) > WStock Then
                        WImpre = Str$(WStock)
                        WImpre = Pusing("###,###.##", WImpre)
                        m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
                        G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        Entra = "N"
                    End If
                    
                    If Val(Cantidad) = 0 Then
                        m$ = "No se informo cantidad para el item " + WImpre1
                        G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        Entra = "N"
                    End If
                    
                End If
                                        
            Next iRow
            
    Next a
    
    Entra = "S"
        
    If Entra = "S" Then
    
        T$ = "Hoja de Produccion"
        m$ = "Desea Imprimir la Hoja de Produccion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion
        End If

        Renglon = 0
        Erase Auxiliar
        
        DBGrid1.Refresh
        
        Hoja.Text = WHoja
        Fecha.Text = WFecha
        Producto.Text = WProducto
        Teorico.Text = WTeorico
        Real.Text = WReal
        fechaIng.Text = WFechaing
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Tipo = DBGrid1.Text
                    
                DBGrid1.Col = 1
                Terminado = UCase(DBGrid1.Text)
                    
                DBGrid1.Col = 2
                Articulo = DBGrid1.Text
                    
                DBGrid1.Col = 4
                Cantidad = DBGrid1.Text
                    
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
                    WLote = ""
                    WDate = Date$
                    WImporte = ""
                    WMarca = ""
                    WSaldo = "0"
                    ZLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    WLote1 = ZLote(ZLugar, 1)
                    WLote2 = ZLote(ZLugar, 3)
                    Wlote3 = ZLote(ZLugar, 5)
                    WCanti1 = ZLote(ZLugar, 2)
                    WCanti2 = ZLote(ZLugar, 4)
                    WCanti3 = ZLote(ZLugar, 6)
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
                            + Wlote3 + "','" + Wlote3 + "','" _
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
                    
                End If
                        
            Next iRow
            
        Next a
        
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
                            WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                            WDate = Date$
                            rstArticulo.Close
                            XParam = "'" + WCodigo + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                                                
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
        
        Call Limpia_Click
        
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Hoja.SetFocus
        
        PrgHojaAnt.Hide
        Unload Me
        PrgEti2.Show
        
            Else
            
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Hoja.SetFocus
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
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
    Wlote3.Text = ""
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
    
End Sub

Private Sub Limpia_Click()

    Graba.Enabled = True
    Reimpresion.Enabled = False
    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    Observaciones.Text = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    Graba.Enabled = True
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    Wlote3.Text = ""
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
    Erase ZLote
    
    
    salgo = "N"
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
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
    Next a
    
    Rem With rstHoja
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Hoja.Text = !Hoja + 1
    Rem             Else
    Rem         Hoja.Text = ""
    Rem     End If
    Rem End With
    
    If LeeHoja <> "N" Then
        spHoja = "ListaHojaNumero"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
                With rstHoja
                .MoveLast
                Hoja.Text = rstHoja!Hoja + 1
            End With
            rstHoja.Close
                Else
            Hoja.Text = "1"
        End If
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Hoja.SetFocus

End Sub

Private Sub Reimpresion_Click()

        T$ = "Hoja de Produccion"
        m$ = "Desea Imprimir la Hoja de Produccion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion
        End If
        
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Hoja.SetFocus
        
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem WStock = 0
        Rem Select Case WTipo
        Rem     Case "T"
        Rem         spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
        Rem         Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstTerminado.RecordCount > 0 Then
        Rem             WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
        Rem             WImpre1 = WTerminado.Text
        Rem             rstTerminado.Close
        Rem         End If
        Rem     Case "M"
        Rem         spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Rem         Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstArticulo.RecordCount > 0 Then
        Rem             WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
        Rem             WImpre1 = WArticulo.Text
        Rem             rstArticulo.Close
        Rem         End If
        Rem     Case Else
        Rem End Select
        Rem If Val(WCantidad.Text) <= WStock Then
        Rem     WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Rem     Rem Call Alta_Vector
        Rem     Rem Call Ingresa_Click
        Rem     WTipo.SetFocus
        Rem         Else
        Rem     WImpre = Str$(WStock)
        Rem     WImpre = Pusing("###,###.##", WImpre)
        Rem     m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
        Rem     ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
        Rem     WCantidad.Text = "0"
        Rem     WCantidad.SetFocus
        Rem End If
        
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
        Wlote3.Text = ""
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
        
        If Val(ZLote(Val(WLinea.Text), 1)) <> 0 Then
            WLote1.Text = ZLote(Val(WLinea.Text), 1)
            WCanti1.Text = ZLote(Val(WLinea.Text), 2)
            WControl1.Locked = False
            WControl1.Text = ""
            WControl1.Locked = True
        End If
        If Val(ZLote(Val(WLinea.Text), 3)) <> 0 Then
            WLote2.Text = ZLote(Val(WLinea.Text), 3)
            WCanti2.Text = ZLote(Val(WLinea.Text), 4)
            WControl2.Locked = False
            WControl2.Text = ""
            WControl2.Locked = True
        End If
        If Val(ZLote(Val(WLinea.Text), 5)) <> 0 Then
            Wlote3.Text = ZLote(Val(WLinea.Text), 5)
            WCanti3.Text = ZLote(Val(WLinea.Text), 6)
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
                    Rem WTipo.SetFocus
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
                    Rem WTipo.SetFocus
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
                    Rem WTipo.SetFocus
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
                    Rem WTipo.SetFocus
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
            Wlote3.SetFocus
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
            
            If Val(Wlote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    Rem WTipo.SetFocus
                    Exit Sub
                        Else
                    Wlote3.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
                
            If Val(Wlote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    Rem WTipo.SetFocus
                    Exit Sub
                        Else
                    Wlote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
                Rem WTipo.SetFocus
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
            Wlote3.SetFocus
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
                WArticulo.Text = rstArticulorstArticulo!Codigo
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
            Indice = Pantalla.ListIndex
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
ReDim UserData(0 To 4, 0 To 40)

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
For i = 0 To 4
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i

    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa

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
    Wlote3.Text = ""
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
    Observaciones.Text = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    Graba.Enabled = True
    Reimpresion.Enabled = False
    
    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            Hoja.Text = rstHoja!Hoja + 1
        End With
        rstHoja.Close
            Else
        Hoja.Text = "1"
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgHoja.Caption = "Ingreso de Hoja de produccion :  " + !Nombre
        End If
    End With
    
    Muestra.ColWidth(0) = 150
    Muestra.ColWidth(1) = 400
    Muestra.ColWidth(2) = 800
    Muestra.ColWidth(3) = 800
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Partida"
    
    Muestra.Col = 3
    Muestra.Text = "Stock"
    
    Muestra1.ColWidth(0) = 150
    Muestra1.ColWidth(1) = 800
    Muestra1.ColWidth(2) = 800
    Muestra1.ColWidth(3) = 1000
    
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Cliente"
    
    Muestra1.Col = 3
    Muestra1.Text = "Fecha"
    
    Muestra1.Col = 2
    Muestra1.Text = "Cantidad"
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Hoja.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
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
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    
                    dada = rstHoja!Saldo
                
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
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            Case Else
        End Select
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
    
    Rem Graba.Enabled = True
    
    If Val(Real.Text) <> 0 Then
        Graba.Enabled = False
    End If
  
End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
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
                
            ZLote(Renglon, 1) = WLote1.Text
            ZLote(Renglon, 2) = WCanti1.Text
            ZLote(Renglon, 3) = WLote2.Text
            ZLote(Renglon, 4) = WCanti2.Text
            ZLote(Renglon, 5) = Wlote3.Text
            ZLote(Renglon, 6) = WCanti3.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            WRen = Val(WLinea.Text)
            
            Lugar1 = Int((WRen - 1) / 10) * 10
            Lugar2 = WRen - Lugar1
            
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            
            Rem DBGrid1.Row = Val(WLinea.Text) - 1
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
                
            ZLote(WRen, 1) = WLote1.Text
            ZLote(WRen, 2) = WCanti1.Text
            ZLote(WRen, 3) = WLote2.Text
            ZLote(WRen, 4) = WCanti2.Text
            ZLote(WRen, 5) = Wlote3.Text
            ZLote(WRen, 6) = WCanti3.Text
                
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Entra = "N"
        Existe = "N"
        spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            Existe = "S"
            Fecha.Text = rstHoja!Fecha
            Real.Text = rstHoja!Real
            Teorico.Text = rstHoja!Teorico
            fechaIng.Text = rstHoja!fechaIng
            Producto.Text = rstHoja!Producto
            rstHoja.Close
        End If
        
        If Existe = "S" Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Producto.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                Observaciones.Text = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
                rstTerminado.Close
            End If
            Graba.Enabled = False
            Reimpresion.Enabled = True
            Call Proceso_Click
                
                Else
                    
            Graba.Enabled = True
            Reimpresion.Enabled = False
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
    Rem        Producto.SetFocus
    Rem            Else
    Rem        Fecha.SetFocus
    Rem    End If
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
                Observaciones.Text = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
                rstTerminado.Close
                Call Calcula_stock
                Teorico.SetFocus
                    Else
                Producto.Text = Producto.Text
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
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Lee_Composicion()

    salgo = "N"
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
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
    Next a

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
                End If
            Case Else
        End Select
        
        If XCantidad <= WStock Then
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", Str$(XCantidad))
                Else
            WImpre = Str$(WStock)
            WImpre = Pusing("###,###.##", WImpre)
            m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            DBGrid1.Col = 4
            DBGrid1.Text = "0"
        End If
        
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

    Select Case Val(WEmpresa)
        Case 1, 2, 5
            Sql1 = "DELETE ImpreHoja"
            spImpreHoja = Sql1
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
        
            WHoja = Hoja.Text
            WFecha = Fecha.Text
            WCodigo1 = Left$(Producto.Text, 2)
            WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
            Select Case Val(WEmpresa)
                Case 1
                    WMaquina = "SI"
                Case 2
                    WMaquina = "PI"
                Case 3
                    WMaquina = "SII"
                Case 4
                    WMaquina = "PII"
                Case 5
                    WMaquina = "SIII"
                Case 6
                    WMaquina = "SIV"
                Case 7
                    WMaquina = "SV"
                Case 8
                    WMaquina = "PV"
                Case 9
                    WMaquina = "PVI"
                Case 10
                    WMaquina = "SVI"
                Case 11
                    WMaquina = "SVII"
                Case Else
            End Select
            WTeorico = Teorico.Text

            Linea = 0
        
            For a = 0 To 3
        
                Suma = a * 10
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
                 
                    If Tipo = "M" Then
                    
                        Rem PROCESA LOS LAUDOS
    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
    
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
                            spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                            
                                With rstLaudo
                                    .MoveFirst
                                    If .NoMatch = False Then
                                        Do
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                               
                                            If rstLaudo!Saldo <> 0 Then
                                                If rstLaudo!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstLaudo!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = rstLaudo!Saldo
                                                            XCanti = XCanti - rstLaudo!Saldo
                                                        End If
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
                                rstLaudo.Close
                            End If
                        
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
                            spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                               
                                With rstMovguia
                               
                                    .MoveFirst
                               
                                    If .NoMatch = False Then
                                        Do
                               
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                               
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                            
                        End If
                        
                        WArticulo1 = Left$(Articulo, 2)
                        WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
                        WCantidad = Str$(Cantidad / 100)
                        
                        ZCanti1 = Str$(Impre(1, 2))
                        ZLote1 = Str$(Impre(1, 1))
                        ZCanti2 = Str$(Impre(2, 2))
                        ZLote2 = Str$(Impre(2, 1))
                        ZCanti3 = Str$(Impre(3, 2))
                        ZLote3 = Str$(Impre(3, 1))
                        
                        Linea = Linea + 1
                        WLinea = Str$(Linea)
                        
                        Sql1 = "INSERT INTO ImpreHoja ("
                        Sql2 = "Hoja ,"
                        Sql3 = "Renglon ,"
                        Sql4 = "Fecha ,"
                        Sql5 = "Codigo1 ,"
                        Sql6 = "Codigo2 ,"
                        Sql7 = "Maquina ,"
                        Sql8 = "Articulo1 ,"
                        Sql9 = "Articulo2 ,"
                        Sql10 = "Cantidad ,"
                        Sql11 = "Canti1 ,"
                        Sql12 = "Lote1 ,"
                        Sql13 = "Canti2 ,"
                        Sql14 = "Lote2 ,"
                        Sql15 = "Canti3 ,"
                        Sql16 = "Lote3 ,"
                        Sql17 = "Teorico )"
                        Sql18 = "Values ("
                        Sql19 = "'" + WHoja + "',"
                        Sql20 = "'" + WLinea + "',"
                        Sql21 = "'" + WFecha + "',"
                        Sql22 = "'" + WCodigo1 + "',"
                        Sql23 = "'" + WCodigo2 + "',"
                        Sql24 = "'" + WMaquina + "',"
                        Sql25 = "'" + WArticulo1 + "',"
                        Sql26 = "'" + WArticulo2 + "',"
                        Sql27 = "'" + WCantidad + "',"
                        Sql28 = "'" + ZCanti1 + "',"
                        Sql29 = "'" + ZLote1 + "',"
                        Sql30 = "'" + ZCanti2 + "',"
                        Sql31 = "'" + ZLote2 + "',"
                        Sql32 = "'" + ZCanti3 + "',"
                        Sql33 = "'" + ZLote3 + "',"
                        Sql34 = "'" + WTeorico + "')"
        
                        spImpreHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                                Sql31 + Sql32 + Sql33 + Sql34
                        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If

                    If Tipo = "T" Then
                    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
                            spHoja = "ListaHojaProductoDesdeHasta" + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                            
                                With rstHoja
                                                        
                                    .MoveFirst
                            
                                    If .NoMatch = False Then
                                        Do
                            
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                            
                                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                                If rstHoja!Producto = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstHoja!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                                            XCanti = XCanti - rstHoja!Saldo
                                                        End If
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
                                rstHoja.Close
                            End If
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
                            spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                            
                                With rstMovguia
                            
                                    .MoveFirst
                            
                                    If .NoMatch = False Then
                                        Do
                                                        
                                            If .EOF = True Then
                                                 Exit Do
                                            End If
                            
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Terminado = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                        
                        End If

                        Linea = Linea + 1
                        WLinea = Str$(Linea)
                        
                        WArticulo1 = Left$(Terminado, 2)
                        WArticulo2 = Mid$(Terminado, 4, 5) + "-" + Right$(Terminado, 3)
                        
                        WCantidad = Str$(Cantidad / 100)
                        
                        WCanti1 = Str$(Impre(1, 2))
                        WLote1 = Str$(Impre(1, 1))
                        WCanti2 = Str$(Impre(2, 2))
                        WLote2 = Str$(Impre(2, 1))
                        WCanti3 = Str$(Impre(3, 2))
                        Wlote3 = Str$(Impre(3, 1))
                        
                        Sql1 = "INSERT INTO ImpreHoja ("
                        Sql2 = "Hoja ,"
                        Sql3 = "Renglon ,"
                        Sql4 = "Fecha ,"
                        Sql5 = "Codigo1 ,"
                        Sql6 = "Codigo2 ,"
                        Sql7 = "Maquina ,"
                        Sql8 = "Articulo1 ,"
                        Sql9 = "Articulo2 ,"
                        Sql10 = "Cantidad ,"
                        Sql11 = "Canti1 ,"
                        Sql12 = "Lote1 ,"
                        Sql13 = "Canti2 ,"
                        Sql14 = "Lote2 ,"
                        Sql15 = "Canti3 ,"
                        Sql16 = "Lote3 ,"
                        Sql17 = "Teorico )"
                        Sql18 = "Values ("
                        Sql19 = "'" + WHoja + "',"
                        Sql20 = "'" + WLinea + "',"
                        Sql21 = "'" + WFecha + "',"
                        Sql22 = "'" + WCodigo1 + "',"
                        Sql23 = "'" + WCodigo2 + "',"
                        Sql24 = "'" + WMaquina + "',"
                        Sql25 = "'" + WArticulo1 + "',"
                        Sql26 = "'" + WArticulo2 + "',"
                        Sql27 = "'" + WCantidad + "',"
                        Sql28 = "'" + WCanti1 + "',"
                        Sql29 = "'" + WLote1 + "',"
                        Sql30 = "'" + WCanti2 + "',"
                        Sql31 = "'" + WLote2 + "',"
                        Sql32 = "'" + WCanti3 + "',"
                        Sql33 = "'" + Wlote3 + "',"
                        Sql34 = "'" + WTeorico + "')"
        
                        spImpreHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                                Sql31 + Sql32 + Sql33 + Sql34
                        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                Next iRow
            
            Next a
            
            XLinea = Linea

            For Ciclo = XLinea To 13
            
                Linea = Linea + 1
                WLinea = Str$(Linea)
                        
                WArticulo1 = ""
                WArticulo2 = ""
                WCantidad = ""
                WCanti1 = ""
                WLote1 = ""
                WCanti2 = ""
                WLote2 = ""
                WCanti3 = ""
                Wlote3 = ""
            
                Sql1 = "INSERT INTO ImpreHoja ("
                Sql2 = "Hoja ,"
                Sql3 = "Renglon ,"
                Sql4 = "Fecha ,"
                Sql5 = "Codigo1 ,"
                Sql6 = "Codigo2 ,"
                Sql7 = "Maquina ,"
                Sql8 = "Articulo1 ,"
                Sql9 = "Articulo2 ,"
                Sql10 = "Cantidad ,"
                Sql11 = "Canti1 ,"
                Sql12 = "Lote1 ,"
                Sql13 = "Canti2 ,"
                Sql14 = "Lote2 ,"
                Sql15 = "Canti3 ,"
                Sql16 = "Lote3 ,"
                Sql17 = "Teorico )"
                Sql18 = "Values ("
                Sql19 = "'" + WHoja + "',"
                Sql20 = "'" + WLinea + "',"
                Sql21 = "'" + WFecha + "',"
                Sql22 = "'" + WCodigo1 + "',"
                Sql23 = "'" + WCodigo2 + "',"
                Sql24 = "'" + WMaquina + "',"
                Sql25 = "'" + WArticulo1 + "',"
                Sql26 = "'" + WArticulo2 + "',"
                Sql27 = "'" + WCantidad + "',"
                Sql28 = "'" + WCanti1 + "',"
                Sql29 = "'" + WLote1 + "',"
                Sql30 = "'" + WCanti2 + "',"
                Sql31 = "'" + WLote2 + "',"
                Sql32 = "'" + WCanti3 + "',"
                Sql33 = "'" + Wlote3 + "',"
                Sql34 = "'" + WTeorico + "')"
        
                spImpreHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                        Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                        Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                        Sql31 + Sql32 + Sql33 + Sql34
                Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            Listado.WindowTitle = "Impresion de Hoja de Produccion"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
   
            Listado.Destination = 1
            Rem Listado.Destination = 0
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Canti1, ImpreHoja.Lote1, ImpreHoja.Canti2, ImpreHoja.Lote2, ImpreHoja.CAnti3, ImpreHoja.Lote3, ImpreHoja.Teorico " _
                        + "From " _
                        + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                        + "Where " _
                        + "ImpreHoja.Hoja >= 0 AND ImpreHoja.Hoja <= 999999 "
    
            Listado.Connect = Connect()
            Listado.Action = 1
        
        Case Else
            Open "lpt1" For Output As #1
            Rem Open "hoja.txt" For Output As #1

            Print #1, Chr$(27) + Chr$(71)
            Print #1,
            Print #1, Chr$(18)

            Print #1, Tab(15); Left$(Producto.Text, 2);
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
                Case 9
                    Print #1, Tab(70); "PVI"
                Case 10
                    Print #1, Tab(70); "SVI"
                Case 11
                    Print #1, Tab(70); "SVII"
                Case Else
            End Select

            Print #1, Tab(1); Fecha.Text;
            Print #1, Tab(12); Alinea("#####", Mid$(Producto.Text, 4, 5));
            Print #1, "/"; Right$(Producto.Text, 3);
            Print #1, Tab(26); Chr$(14); Alinea("######", Hoja.Text)

            Print #1,
            Print #1,

            Linea = 0
        
            For a = 0 To 3
        
                Suma = a * 10
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
                 
                    If Tipo = "M" Then
                    
                        Rem PROCESA LOS LAUDOS
    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
    
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
                            spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                            
                                With rstLaudo
                                    .MoveFirst
                                    If .NoMatch = False Then
                                        Do
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                               
                                            If rstLaudo!Saldo <> 0 Then
                                                If rstLaudo!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstLaudo!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = rstLaudo!Saldo
                                                            XCanti = XCanti - rstLaudo!Saldo
                                                        End If
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
                                rstLaudo.Close
                            End If
                        
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
                            spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                               
                                With rstMovguia
                               
                                    .MoveFirst
                               
                                    If .NoMatch = False Then
                                        Do
                               
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                               
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                            
                        End If
                        
                        Linea = Linea + 1

                        Print #1, Tab(6); Left$(Articulo, 2);
                        Print #1, Tab(11); Mid$(Articulo, 4, 3);
                        Print #1, "-";
                        Print #1, Right$(Articulo, 3);
                        If Val(Teorico.Text) < 100 Then
                            Print #1, Tab(20); Alinea("###.##", Cantidad);
                                Else
                            Print #1, Tab(20); Alinea("####.#", Cantidad);
                        End If
                        
                        If Impre(1, 2) <> 0 Then
                            If Impre(1, 2) < 100 Then
                                Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                                    Else
                                Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                            End If
                        End If
                        If Impre(1, 1) <> 0 Then
                            Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                        End If
                        
                        If Impre(2, 2) <> 0 Then
                            If Impre(2, 2) < 100 Then
                                Print #1, Tab(41); Alinea("###.##", Str$(Impre(2, 2)));
                                    Else
                                Print #1, Tab(41); Alinea("####.#", Str$(Impre(2, 2)));
                            End If
                        End If
                        If Impre(2, 1) <> 0 Then
                            Print #1, Tab(48); Alinea("######", Str$(Impre(2, 1)));
                        End If
                        
                        If Impre(3, 2) <> 0 Then
                            If Impre(3, 2) < 100 Then
                                Print #1, Tab(55); Alinea("###.##", Str$(Impre(3, 2)));
                                    Else
                                Print #1, Tab(55); Alinea("####.#", Str$(Impre(3, 2)));
                            End If
                        End If
                        If Impre(3, 1) <> 0 Then
                            Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                        End If
                            
                        Print #1,
                        Print #1,

                    End If

                    If Tipo = "T" Then
                    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
                            spHoja = "ListaHojaProductoDesdeHasta" + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                            
                                With rstHoja
                                                        
                                    .MoveFirst
                            
                                    If .NoMatch = False Then
                                        Do
                            
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                            
                                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                                If rstHoja!Producto = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstHoja!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                                            XCanti = XCanti - rstHoja!Saldo
                                                        End If
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
                                rstHoja.Close
                            End If
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
                            spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                            
                                With rstMovguia
                            
                                    .MoveFirst
                            
                                    If .NoMatch = False Then
                                        Do
                                                        
                                            If .EOF = True Then
                                                 Exit Do
                                            End If
                            
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Terminado = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                        
                        End If

                        Linea = Linea + 1

                        Print #1, Tab(6); Left$(Terminado, 2);
                        Print #1, Tab(11); Mid$(Terminado, 4, 5);
                        Print #1, "-";
                        Print #1, Right$(Terminado, 3);
                        If Val(Teorico.Text) < 100 Then
                            Print #1, Tab(20); Alinea("###.##", Cantidad);
                                Else
                            Print #1, Tab(20); Alinea("####.#", Cantidad);
                        End If
                        
                        If Impre(1, 2) <> 0 Then
                            If Impre(1, 2) < 100 Then
                                Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                                    Else
                                Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                            End If
                        End If
                        If Impre(1, 1) <> 0 Then
                            Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                        End If
                        
                        If Impre(2, 2) <> 0 Then
                            If Impre(2, 2) < 100 Then
                                Print #1, Tab(40); Alinea("###.##", Str$(Impre(2, 2)));
                                    Else
                                Print #1, Tab(40); Alinea("####.#", Str$(Impre(2, 2)));
                            End If
                        End If
                        If Impre(2, 1) <> 0 Then
                            Print #1, Tab(46); Alinea("######", Str$(Impre(2, 1)));
                        End If
                        
                        If Impre(3, 2) <> 0 Then
                            If Impre(3, 2) < 100 Then
                                Print #1, Tab(54); Alinea("###.##", Str$(Impre(3, 2)));
                                    Else
                                Print #1, Tab(54); Alinea("####.#", Str$(Impre(3, 2)));
                            End If
                        End If
                        If Impre(3, 1) <> 0 Then
                            Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                        End If
                        
                        Print #1,
                        Print #1,

                    End If
                    
                Next iRow
            
            Next a

            For Ciclo = Linea To 14

                Print #1,
                Print #1,

            Next Ciclo

            Print #1, Tab(20); Alinea("####.#", Teorico.Text)

            Print #1,
            Print #1, Chr$(27) + Chr$(72)
            Print #1, Chr$(12)
        
            Close #1
            
    End Select

End Sub

Sub Etiqueta()
    PCliente = Cliente.Text
    PTipo = 1
    PrgConsCcte.Show
End Sub

Sub Calcula_stock()

    Muestra.Clear
    Muestra.Row = 0
 
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Partida"
    
    Muestra.Col = 3
    Muestra.Text = "Stock"
    
    Muestra1.Clear
    Muestra1.Row = 0
 
    Muestra1.Col = 1
    Muestra1.Text = "Cliente"
    
    Muestra1.Col = 3
    Muestra1.Text = "Fecha"
    
    Muestra1.Col = 2
    Muestra1.Text = "Cantidad"
    
    Producto.Text = UCase(Producto.Text)
    XProducto = Producto.Text
    Renglon = 0
    XStock = 0
    XPedido = 0
    
    Rem lee pt
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                 
                    XHoja = rstHoja!Hoja
                    XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "PT"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "PT"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    
    Rem lee NK
    
    XProducto = "NK" + Mid$(XProducto, 3, 10)
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                 
                    XHoja = rstHoja!Hoja
                    XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "NK"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "NK"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    
    Rem lee re
    
    XProducto = "RE" + Mid$(XProducto, 3, 10)
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                 
                    XHoja = rstHoja!Hoja
                    XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "RE"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    XParam = "'" + XProducto + "','" _
                 + XProducto + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "RE"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
                        XStock = XStock + XSaldo
                        
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
    End If
    
    spPedido = "ListaPedidoTerminado " + "'" + Producto.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XPed = rstPedido!Cantidad - rstPedido!Facturado
                
                If XPed <> 0 And Left$(rstPedido!Cliente, 1) = "X" Then
                    
                    Renglon = Renglon + 1
                    Muestra1.Row = Renglon
                    
                    Muestra1.Col = 1
                    Muestra1.Text = rstPedido!Cliente
                
                    Muestra1.Col = 3
                    Muestra1.Text = rstPedido!Fecha
                        
                    Muestra1.Col = 2
                    Muestra1.Text = Pusing("###,###.##", Str$(XPed))
                        
                    XPedido = XPedido + XPed
                        
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
    End If
    
    Stock.Text = Pusing("###,###.##", Str$(XStock))
    Pedido.Text = Pusing("###,###.##", Str$(XPedido))
    
End Sub


Private Sub Verifica_Lote()

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
    
    
    If Val(Wlote3.Text) <> 0 Then
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
                XParam = "'" + Wlote3.Text + "','" _
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
                            + Wlote3.Text + "'"
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
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente"
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
    If Val(Wlote3.Text) <> 0 And WControl3.Text = "X" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    
    If Suma = Val(WCantidad.Text) Then
        WEstado = "S"
    End If
    
    If WControla <> 0 Then
        WEstado = "S"
    End If
    
End Sub

Private Sub WLote1_DblClick()
    ZProceso = 1
    If WTipo.Text = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub WLote2_DblClick()
    ZProceso = 2
    If WTipo.Text = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub WLote3_DblClick()
    ZProceso = 3
    If WTipo.Text = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub ficha_Mp()

    Call Limpia_Vector
    
    XRenglon = 0
    XParam = "'" + WArticulo.Text + "','" _
                 + WArticulo.Text + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
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
                    
                    If rstLaudo!Articulo = WArticulo.Text Then
                
                        ZArticulo = rstLaudo!Articulo
                        ZCantidad = rstLaudo!Liberada
                        ZFecha = rstLaudo!Fecha
                        ZLaudo = rstLaudo!Laudo
                        ZOrden = rstLaudo!Orden
                        Zdevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        ZRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        ZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        ZLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        Call Redondeo(ZSaldo)
                        
                        If ZLiberada <> 0 And ZSaldo <> 0 Then
                        
                            XRenglon = XRenglon + 1
                            WVector1.Row = XRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = "Laudo"
                        
                            WVector1.Col = 2
                            WVector1.Text = ZLaudo
                                               
                            WVector1.Col = 3
                            WVector1.Text = ZFecha
                        
                            WVector1.Col = 4
                            WVector1.Text = ZOrden
                        
                            WVector1.Col = 5
                            WVector1.Text = ZCantidad
                
                            WVector1.Col = 6
                            WVector1.Text = ZSaldo
                
                            WVector1.Col = 7
                            WVector1.Text = ZLaudo
                            
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
        rstLaudo.Close
    End If
    
    XParam = "'" + WArticulo.Text + "','" _
                + WArticulo.Text + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo.Text Then
                    
                        ZArticulo = rstMovguia!Articulo
                        ZCantidad = rstMovguia!Cantidad
                        ZFecha = rstMovguia!Fecha
                        ZCodigo = rstMovguia!Codigo
                        ZMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        ZTipomov = rstMovguia!Tipomov
                        ZPartida = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        ZFecha = rstMovguia!Fecha
                        If Val(ZCodigo) > 900000 Then
                            ZTipo = "Prestamo"
                            ZCodigo = ZCodigo - 900000
                                Else
                            ZTipo = "Guia In"
                        End If
                        ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZSaldo)
                                
                        If rstMovguia!Movi = "E" And ZSaldo <> 0 Then
                            
                            XRenglon = XRenglon + 1
                            WVector1.Row = XRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = ZTipo
                        
                            WVector1.Col = 2
                            WVector1.Text = ZCodigo
                                               
                            WVector1.Col = 3
                            WVector1.Text = ZFecha
                        
                            WVector1.Col = 4
                            WVector1.Text = ""
                        
                            WVector1.Col = 5
                            WVector1.Text = ZCantidad
                
                            WVector1.Col = 6
                            WVector1.Text = ZSaldo
                
                            WVector1.Col = 7
                            WVector1.Text = ZPartida
                            
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
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
End Sub

Private Sub ficha_Pt()

    Call Limpia_Vector
    
    XRenglon = 0
    
    XParam = "'" + WTerminado.Text + "','" _
                 + WTerminado.Text + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                Rem And rstHoja!Real <> 0 Then
                 
                    ZProducto = rstHoja!Producto
                    ZCantidad = rstHoja!Real
                    ZFecha = rstHoja!Fecha
                    ZHoja = rstHoja!Hoja
                    ZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZSaldo)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector1.Row = XRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = "Hoja"
                        
                        WVector1.Col = 2
                        WVector1.Text = ZHoja
                                               
                        WVector1.Col = 3
                        WVector1.Text = ZFecha
                        
                        WVector1.Col = 4
                        WVector1.Text = ""
                        
                        WVector1.Col = 5
                        WVector1.Text = ZCantidad
                
                        WVector1.Col = 6
                        WVector1.Text = ZSaldo
                
                        WVector1.Col = 7
                        WVector1.Text = ZHoja
                    
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
    End If
    
    
    
    XParam = "'" + WTerminado.Text + "','" _
                 + WTerminado.Text + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    ZTerminado = rstMovguia!Terminado
                    ZCantidad = rstMovguia!Cantidad
                    ZFecha = rstMovguia!Fecha
                    ZCodigo = rstMovguia!Codigo
                    ZMovi = rstMovguia!Movi
                    ZDestino = rstMovguia!Destino
                    ZTipomov = rstMovguia!Tipomov
                    ZZLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    ZPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                    ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZSaldo)
                    If Val(ZCodigo) > 900000 Then
                        ZTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        ZTipo = "Guia In"
                    End If
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector1.Row = XRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = ZTipo
                        
                        WVector1.Col = 2
                        WVector1.Text = ZCodigo
                                               
                        WVector1.Col = 3
                        WVector1.Text = ZFecha
                        
                        WVector1.Col = 4
                        WVector1.Text = ""
                        
                        WVector1.Col = 5
                        WVector1.Text = ZCantidad
                
                        WVector1.Col = 6
                        WVector1.Text = ZSaldo
                
                        WVector1.Col = 7
                        WVector1.Text = ZZLote
                        
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
    End If
    
    
    
    XParam = "'" + WTerminado.Text + "','" _
                 + WTerminado.Text + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntdev!Marca = "X" Then
                
                        Else
                
                ZTerminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                ZZLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector1.Row = XRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = "Dev"
                        
                    WVector1.Col = 2
                    WVector1.Text = ZCodigo
                                               
                    WVector1.Col = 3
                    WVector1.Text = ZFecha
                        
                    WVector1.Col = 4
                    WVector1.Text = ""
                        
                    WVector1.Col = 5
                    WVector1.Text = ZCantidad
                
                    WVector1.Col = 6
                    WVector1.Text = ZSaldo
                
                    WVector1.Col = 7
                    WVector1.Text = ZZLote
                        
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
End Sub


Private Sub Limpia_Vector()

    WVector1.Height = 4095
    WVector1.Left = 120
    WVector1.Top = 1200
    WVector1.Width = 10000

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 5001
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector1.Text = "Orden"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Visible = True
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Click()
    busquedalote = WVector1.TextMatrix(WVector1.Row, 7)
    WVector1.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    WTitulo(5).Visible = False
    WTitulo(6).Visible = False
    WTitulo(7).Visible = False
    Select Case ZProceso
        Case 1
            WLote1.Text = busquedalote
            WCanti1.SetFocus
        Case 2
            WLote2.Text = busquedalote
            WCanti2.SetFocus
        Case 3
            Wlote3.Text = busquedalote
            WCanti3.SetFocus
        Case Else
    End Select
        
End Sub


