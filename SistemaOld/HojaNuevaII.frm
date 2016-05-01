VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHojaNuevaII 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Hoja de Produccion"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   Visible         =   0   'False
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
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Operario 
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
      MaxLength       =   2
      TabIndex        =   1
      Text            =   " "
      Top             =   1200
      Width           =   735
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
      Left            =   11160
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   59
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
      Left            =   9120
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   58
      Text            =   " "
      Top             =   120
      Width           =   495
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
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   57
      Text            =   " "
      Top             =   120
      Width           =   495
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   6960
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   6840
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   6840
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   6840
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
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   6960
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   7200
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
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   615
      Left            =   10560
      TabIndex        =   48
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   6480
      TabIndex        =   34
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   9480
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HojaNuevaII.frx":0000
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
      Left            =   8400
      TabIndex        =   32
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6240
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra1 
      Height          =   1815
      Left            =   9240
      TabIndex        =   29
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   327680
      Rows            =   100
      Cols            =   4
   End
   Begin VB.TextBox Stock 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   2415
      Left            =   9240
      TabIndex        =   27
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      _Version        =   327680
      Rows            =   100
      Cols            =   4
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
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
      TabIndex        =   5
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9600
      Top             =   7080
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
      Left            =   1200
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   9
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
         TabIndex        =   21
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   20
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
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   11
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   10
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
      TabIndex        =   8
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "HojaNuevaII.frx":007C
      TabIndex        =   7
      Top             =   1560
      Width           =   9015
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
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
      Left            =   3360
      TabIndex        =   62
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Operario"
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
      TabIndex        =   61
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DesOperario 
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
      Left            =   2880
      TabIndex        =   60
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label17 
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
      Left            =   9960
      TabIndex        =   56
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label16 
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
      Left            =   8160
      TabIndex        =   55
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label15 
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
      Left            =   5880
      TabIndex        =   54
      Top             =   120
      Width           =   1455
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
      TabIndex        =   31
      Top             =   480
      Width           =   1455
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
      TabIndex        =   18
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
      TabIndex        =   16
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgHojaNuevaII"
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
Dim rstEfluentes As Recordset
Dim spEfluentes As String
Dim rstImpreHojaII As Recordset
Dim spImpreHojaII As String
Dim rstImpreCarga As Recordset
Dim spImpreCarga As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstCargaIII As Recordset
Dim spCargaIII As String
Dim rstCargaV As Recordset
Dim spCargaV As String
Dim rstImpreCargaI As Recordset
Dim spImpreCargaI As String

Dim XParam As String
Dim LeeHoja As String
Dim XSaldo As Double
Dim Impre(10, 2) As Double
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
Dim EmpresaActual As String
Dim ZProceso As Integer
Dim LoteBusqueda As String
Dim ZSaldo As Double
Dim ZMetodo As String
Dim ZEfluentes As String
Dim ZVersionI As String
Dim ZVersionII As String
Dim ZVersionIII As String
Dim ZDesEfluentesI As String
Dim ZDesEfluentesII As String
Dim ZZCanti(10) As String
Dim ZZLote(10) As String
Dim ZZMetodo(10) As String
Dim ZZEspecificacion(20) As String
Dim ZZDesEnsayo(10) As String
Dim ZCompo(100, 10) As String
Dim WVersionI As Integer
Dim WVersionII As Integer
Dim WVersionIII As Integer
Dim ZFarma(1000, 5) As String
Dim ZComparaI(100, 2) As String
Dim ZComparaII(100, 2) As String
Dim WEscrito As Integer
Dim WTeorico As String
Dim ZImpreCarga(200, 3) As String
Dim ZImpreCargaI(100, 20) As String
Dim ZImpreMetodo(100) As String

Dim ZZEquipo As Integer
Dim ZZDescripcionI As String
Dim ZZDescripcionII As String
Dim ZZCantidad As Integer

Private Sub cmdClose_Click()
    PrgHojaNuevaII.Hide
    Unload Me
    PrgConsultaHojaII.Show
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
        Case Else
    End Select
End Sub

Private Sub Graba_Click()
    
    Equipo.Text = UCase(Equipo.Text)
    
    If Equipo.Text = "" Then
        m$ = "No se informo en que equipo se desea realizar la fabricacion"
        A% = MsgBox(m$, 0, "ASGINACION DE HOJAS DE PRODUCCION")
        Exit Sub
    End If

    Sql1 = "Select *"
    Sql2 = " FROM Operarios"
    Sql3 = " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
    spOperarios = Sql1 + Sql2 + Sql3
    Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
    If rstOperarios.RecordCount > 0 Then
        rstOperarios.Close
            Else
        m$ = "No se informo operario"
        A% = MsgBox(m$, 0, "ASGINACION DE HOJAS DE PRODUCCION")
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " Equipo = " + "'" + Equipo.Text + "',"
    ZSql = ZSql + " Operario = " + "'" + Operario.Text + "',"
    ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
    ZSql = ZSql + " TipoEtapa = " + "'" + "0" + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Call cmdClose_Click
                    
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
    WLote3.Text = ""
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

    Hoja.Text = ZHojaProceso
    Fecha.Text = "  /  /    "
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    observaciones.Text = ""
    Teorico.Text = ""
    Operario.Text = ""
    DesOperario.Caption = ""
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    
    Hoja.Text = ZHojaProceso
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgHojaNuevaII.Caption = "Asignacion de Hoja de produccion :  " + !Nombre
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
    
    EmpresaActual = WEmpresa
    
    Rem Hoja.SetFocus
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Auxiliar
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstHoja!Fecha
                    Teorico.Text = Str$(rstHoja!Teorico)
                    Producto.Text = rstHoja!Producto
                    Equipo.Text = IIf(IsNull(rstHoja!Equipo), "", rstHoja!Equipo)
                    VersionI.Text = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
                    VersionII.Text = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
                    VersionIII.Text = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
            
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
    
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesProducto.Caption = rstTerminado!Descripcion
        observaciones.Text = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
        rstTerminado.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For da = 1 To WRenglon
    
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
    Next da

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

    XCodigo = Val(Mid$(Producto.Text, 4, 5))
    XTipoPro = ""
    If Val(WEmpresa) = 1 Then
        If XCodigo >= 0 And XCodigo <= 999 Then
            XTipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 11999 Then
                XTipoPro = "CO"
                    Else
                XTipoPro = ""
            End If
        End If
    End If

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 2, 3, 4, 9, 10
            Sql1 = "DELETE ImpreHoja"
            spImpreHoja = Sql1
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Sql1 = "DELETE ImpreHojaII"
            spImpreHojaII = Sql1
            Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
            
            ZZMetodo(1) = ""
            ZZEspecificacion(1) = ""
            ZZMetodo(2) = ""
            ZZEspecificacion(2) = ""
            ZZMetodo(3) = ""
            ZZEspecificacion(3) = ""
            ZZMetodo(4) = ""
            ZZEspecificacion(4) = ""
            ZZMetodo(5) = ""
            ZZEspecificacion(5) = ""
            ZZMetodo(6) = ""
            ZZEspecificacion(6) = ""
            ZZMetodo(7) = ""
            ZZEspecificacion(7) = ""
            ZZMetodo(8) = ""
            ZZEspecificacion(8) = ""
            ZZMetodo(9) = ""
            ZZEspecificacion(9) = ""
            ZZMetodo(10) = ""
            ZZEspecificacion(10) = ""
            
            ZZEspecificacion(11) = ""
            ZZEspecificacion(12) = ""
            ZZEspecificacion(13) = ""
            ZZEspecificacion(14) = ""
            ZZEspecificacion(15) = ""
            ZZEspecificacion(16) = ""
            ZZEspecificacion(17) = ""
            ZZEspecificacion(18) = ""
            ZZEspecificacion(19) = ""
            ZZEspecificacion(20) = ""
            
            ZZDesEnsayo(1) = ""
            ZZDesEnsayo(2) = ""
            ZZDesEnsayo(3) = ""
            ZZDesEnsayo(4) = ""
            ZZDesEnsayo(5) = ""
            ZZDesEnsayo(6) = ""
            ZZDesEnsayo(7) = ""
            ZZDesEnsayo(8) = ""
            ZZDesEnsayo(9) = ""
            ZZDesEnsayo(10) = ""
            
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10
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
            Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
            spEspecifUnifica = Sql1 + Sql2 + Sql3
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                        
                If rstEspecifUnifica!Ensayo1 <> 0 Then
                    ZZMetodo(1) = rstEspecifUnifica!Ensayo1
                    ZZEspecificacion(1) = rstEspecifUnifica!Valor1
                    ZZEspecificacion(11) = rstEspecifUnifica!Valor11
                End If
                        
                If rstEspecifUnifica!Ensayo2 <> 0 Then
                    ZZMetodo(2) = rstEspecifUnifica!Ensayo2
                    ZZEspecificacion(2) = rstEspecifUnifica!valor2
                    ZZEspecificacion(12) = rstEspecifUnifica!Valor22
                End If
                        
                If rstEspecifUnifica!Ensayo3 <> 0 Then
                    ZZMetodo(3) = rstEspecifUnifica!Ensayo3
                    ZZEspecificacion(3) = rstEspecifUnifica!Valor3
                    ZZEspecificacion(13) = rstEspecifUnifica!Valor33
                End If
                        
                If rstEspecifUnifica!Ensayo4 <> 0 Then
                    ZZMetodo(4) = rstEspecifUnifica!Ensayo4
                    ZZEspecificacion(4) = rstEspecifUnifica!valor4
                    ZZEspecificacion(14) = rstEspecifUnifica!Valor44
                End If
                        
                If rstEspecifUnifica!Ensayo5 <> 0 Then
                    ZZMetodo(5) = rstEspecifUnifica!Ensayo5
                    ZZEspecificacion(5) = rstEspecifUnifica!valor5
                    ZZEspecificacion(15) = rstEspecifUnifica!Valor55
                End If
                        
                If rstEspecifUnifica!Ensayo6 <> 0 Then
                    ZZMetodo(6) = rstEspecifUnifica!Ensayo6
                    ZZEspecificacion(6) = rstEspecifUnifica!valor6
                    ZZEspecificacion(16) = rstEspecifUnifica!Valor66
                End If
                        
                If rstEspecifUnifica!Ensayo7 <> 0 Then
                    ZZMetodo(7) = rstEspecifUnifica!Ensayo7
                    ZZEspecificacion(7) = rstEspecifUnifica!valor7
                    ZZEspecificacion(17) = rstEspecifUnifica!Valor77
                End If
                        
                If rstEspecifUnifica!Ensayo8 <> 0 Then
                    ZZMetodo(8) = rstEspecifUnifica!Ensayo8
                    ZZEspecificacion(8) = rstEspecifUnifica!valor8
                    ZZEspecificacion(18) = rstEspecifUnifica!Valor88
                End If
                        
                If rstEspecifUnifica!Ensayo9 <> 0 Then
                    ZZMetodo(9) = rstEspecifUnifica!Ensayo9
                    ZZEspecificacion(9) = rstEspecifUnifica!valor9
                    ZZEspecificacion(19) = rstEspecifUnifica!Valor88
                End If
                        
                If rstEspecifUnifica!Ensayo10 <> 0 Then
                    ZZMetodo(10) = rstEspecifUnifica!Ensayo10
                    ZZEspecificacion(10) = rstEspecifUnifica!valor10
                    ZZEspecificacion(20) = rstEspecifUnifica!Valor1010
                End If
                            
                rstEspecifUnifica.Close
                            
            End If
            
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(1) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(1) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(2) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(2) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(3) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(3) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(4) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(4) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(5) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(5) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(6) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(6) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(7) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(7) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(8) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(8) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(9) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(9) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(10) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDesEnsayo(10) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
    
            ZZDesEnsayo(1) = StrConv(ZZDesEnsayo(1), 2)
            ZZDesEnsayo(2) = StrConv(ZZDesEnsayo(2), 2)
            ZZDesEnsayo(3) = StrConv(ZZDesEnsayo(3), 2)
            ZZDesEnsayo(4) = StrConv(ZZDesEnsayo(4), 2)
            ZZDesEnsayo(5) = StrConv(ZZDesEnsayo(5), 2)
            ZZDesEnsayo(6) = StrConv(ZZDesEnsayo(6), 2)
            ZZDesEnsayo(7) = StrConv(ZZDesEnsayo(7), 2)
            ZZDesEnsayo(8) = StrConv(ZZDesEnsayo(8), 2)
            ZZDesEnsayo(9) = StrConv(ZZDesEnsayo(9), 2)
            ZZDesEnsayo(10) = StrConv(ZZDesEnsayo(10), 2)
            
            ZZEspecificacion(1) = StrConv(ZZEspecificacion(1), 2)
            ZZEspecificacion(2) = StrConv(ZZEspecificacion(2), 2)
            ZZEspecificacion(3) = StrConv(ZZEspecificacion(3), 2)
            ZZEspecificacion(4) = StrConv(ZZEspecificacion(4), 2)
            ZZEspecificacion(5) = StrConv(ZZEspecificacion(5), 2)
            ZZEspecificacion(6) = StrConv(ZZEspecificacion(6), 2)
            ZZEspecificacion(7) = StrConv(ZZEspecificacion(7), 2)
            ZZEspecificacion(8) = StrConv(ZZEspecificacion(8), 2)
            ZZEspecificacion(9) = StrConv(ZZEspecificacion(9), 2)
            ZZEspecificacion(10) = StrConv(ZZEspecificacion(10), 2)
            ZZEspecificacion(11) = StrConv(ZZEspecificacion(11), 2)
            ZZEspecificacion(12) = StrConv(ZZEspecificacion(12), 2)
            ZZEspecificacion(13) = StrConv(ZZEspecificacion(13), 2)
            ZZEspecificacion(14) = StrConv(ZZEspecificacion(14), 2)
            ZZEspecificacion(15) = StrConv(ZZEspecificacion(15), 2)
            ZZEspecificacion(16) = StrConv(ZZEspecificacion(16), 2)
            ZZEspecificacion(17) = StrConv(ZZEspecificacion(17), 2)
            ZZEspecificacion(18) = StrConv(ZZEspecificacion(18), 2)
            ZZEspecificacion(19) = StrConv(ZZEspecificacion(19), 2)
            ZZEspecificacion(20) = StrConv(ZZEspecificacion(20), 2)
                
            Call Conecta_Empresa
        
            WHoja = Hoja.Text
            WFecha = Fecha.Text
            WCodigo1 = Left$(Producto.Text, 2)
            WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
            Select Case Val(WEmpresa)
                Case 1, 10
                    WMaquina = "I"
                Case 2
                    WMaquina = "I"
                Case 3
                    WMaquina = "II"
                Case 4
                    WMaquina = "II"
                Case 5
                    WMaquina = "III"
                Case 6
                    WMaquina = "IV"
                Case 7
                    WMaquina = "V"
                Case 8
                    WMaquina = "V"
                Case 9
                    WMaquina = "VI"
                Case Else
            End Select
            WTeorico = Teorico.Text
            
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZMetodo = IIf(IsNull(rstTerminado!Metodo), "0", rstTerminado!Metodo)
                ZEfluentes = IIf(IsNull(rstTerminado!Efluentes), "0", rstTerminado!Efluentes)
                ZVersionI = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
                ZVersionII = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
                ZVersionIII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
                rstTerminado.Close
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Efluentes"
            ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + ZEfluentes + "'"
            spEfluentes = ZSql
            Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
            If rstEfluentes.RecordCount > 0 Then
                ZDesEfluentesI = rstEfluentes!Descripcion
                ZDesEfluentesII = rstEfluentes!DescripcionII
                rstEfluentes.Close
            End If

            Linea = 0
            LineaII = 0
        
            For A = 0 To 3
        
                Suma = A * 10
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
                            Impre(4, 1) = 0
                            Impre(4, 2) = 0
                            Impre(5, 1) = 0
                            Impre(5, 2) = 0
                            Impre(6, 1) = 0
                            Impre(6, 2) = 0
                            Impre(7, 1) = 0
                            Impre(7, 2) = 0
                            Impre(8, 1) = 0
                            Impre(8, 2) = 0
                            Impre(9, 1) = 0
                            Impre(9, 2) = 0
                            Impre(10, 1) = 0
                            Impre(10, 2) = 0
    
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
                                                    If Xlugar < 10 And XCanti > 0 Then
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
                                                    If Xlugar < 10 And XCanti > 0 Then
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
                        
                        For ZCiclo = 1 To 10
                            ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                            ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                        Next ZCiclo
                        
                        Linea = Linea + 1
                        WLinea = Str$(Linea)
                        
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO ImpreHoja ("
                        ZSql = ZSql & "Hoja ,"
                        ZSql = ZSql & "Renglon ,"
                        ZSql = ZSql & "Fecha ,"
                        ZSql = ZSql & "Codigo1 ,"
                        ZSql = ZSql & "Codigo2 ,"
                        ZSql = ZSql & "Maquina ,"
                        ZSql = ZSql & "Articulo1 ,"
                        ZSql = ZSql & "Articulo2 ,"
                        ZSql = ZSql & "Cantidad ,"
                        ZSql = ZSql & "Canti1 ,"
                        ZSql = ZSql & "Lote1 ,"
                        ZSql = ZSql & "Canti2 ,"
                        ZSql = ZSql & "Lote2 ,"
                        ZSql = ZSql & "Canti3 ,"
                        ZSql = ZSql & "Lote3 ,"
                        ZSql = ZSql & "Teorico ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes ,"
                        ZSql = ZSql & "DesEfluentesI ,"
                        ZSql = ZSql & "DesEfluentesII ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "VersionIII ,"
                        ZSql = ZSql & "Equipo ,"
                        ZSql = ZSql & "Metodo1 ,"
                        ZSql = ZSql & "Metodo2 ,"
                        ZSql = ZSql & "Metodo3 ,"
                        ZSql = ZSql & "Metodo4 ,"
                        ZSql = ZSql & "Metodo5 ,"
                        ZSql = ZSql & "Metodo6 ,"
                        ZSql = ZSql & "Metodo7 ,"
                        ZSql = ZSql & "Metodo8 ,"
                        ZSql = ZSql & "Metodo9 ,"
                        ZSql = ZSql & "Metodo10 ,"
                        ZSql = ZSql & "Especificacion1 ,"
                        ZSql = ZSql & "Especificacion2 ,"
                        ZSql = ZSql & "Especificacion3 ,"
                        ZSql = ZSql & "Especificacion4 ,"
                        ZSql = ZSql & "Especificacion5 ,"
                        ZSql = ZSql & "Especificacion6 ,"
                        ZSql = ZSql & "Especificacion7 ,"
                        ZSql = ZSql & "Especificacion8 ,"
                        ZSql = ZSql & "Especificacion9 ,"
                        ZSql = ZSql & "Especificacion10 ,"
                        ZSql = ZSql & "DesMetodo1 ,"
                        ZSql = ZSql & "DesMetodo2 ,"
                        ZSql = ZSql & "DesMetodo3 ,"
                        ZSql = ZSql & "DesMetodo4 ,"
                        ZSql = ZSql & "DesMetodo5 ,"
                        ZSql = ZSql & "DesMetodo6 ,"
                        ZSql = ZSql & "DesMetodo7 ,"
                        ZSql = ZSql & "DesMetodo8 ,"
                        ZSql = ZSql & "DesMetodo9 ,"
                        ZSql = ZSql & "DesMetodo10 )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WHoja + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WFecha + "',"
                        ZSql = ZSql & "'" + WCodigo1 + "',"
                        ZSql = ZSql & "'" + WCodigo2 + "',"
                        ZSql = ZSql & "'" + WMaquina + "',"
                        ZSql = ZSql & "'" + WArticulo1 + "',"
                        ZSql = ZSql & "'" + WArticulo2 + "',"
                        ZSql = ZSql & "'" + WCantidad + "',"
                        ZSql = ZSql & "'" + ZZCanti(1) + "',"
                        ZSql = ZSql & "'" + ZZLote(1) + "',"
                        ZSql = ZSql & "'" + ZZCanti(2) + "',"
                        ZSql = ZSql & "'" + ZZLote(2) + "',"
                        ZSql = ZSql & "'" + ZZCanti(3) + "',"
                        ZSql = ZSql & "'" + ZZLote(3) + "',"
                        ZSql = ZSql & "'" + WTeorico + "',"
                        ZSql = ZSql & "'" + ZMetodo + "',"
                        ZSql = ZSql & "'" + ZEfluentes + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                        ZSql = ZSql & "'" + ZVersionI + "',"
                        ZSql = ZSql & "'" + ZVersionII + "',"
                        ZSql = ZSql & "'" + ZVersionIII + "',"
                        ZSql = ZSql & "'" + Equipo.Text + "',"
                        ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
        
                        spImpreHoja = ZSql
                        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                        For ZCiclo = 1 To 10
                        
                            If Val(ZZCanti(ZCiclo)) <> 0 Then
                        
                                LineaII = LineaII + 1
                                WLIneaII = Str$(LineaII)
                            
                                WCantidadII = ZZCanti(ZCiclo)
                                WLoteII = ZZLote(ZCiclo)
                                
                                ZSql = ""
                                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                                ZSql = ZSql & "Hoja ,"
                                ZSql = ZSql & "Renglon ,"
                                ZSql = ZSql & "Fecha ,"
                                ZSql = ZSql & "Codigo1 ,"
                                ZSql = ZSql & "Codigo2 ,"
                                ZSql = ZSql & "Maquina ,"
                                ZSql = ZSql & "Articulo1 ,"
                                ZSql = ZSql & "Articulo2 ,"
                                ZSql = ZSql & "Cantidad ,"
                                ZSql = ZSql & "Lote ,"
                                ZSql = ZSql & "Teorico ,"
                                ZSql = ZSql & "Terminado ,"
                                ZSql = ZSql & "Equipo )"
                                ZSql = ZSql & "Values ("
                                ZSql = ZSql & "'" + WHoja + "',"
                                ZSql = ZSql & "'" + WLIneaII + "',"
                                ZSql = ZSql & "'" + WFecha + "',"
                                ZSql = ZSql & "'" + WCodigo1 + "',"
                                ZSql = ZSql & "'" + WCodigo2 + "',"
                                ZSql = ZSql & "'" + WMaquina + "',"
                                ZSql = ZSql & "'" + WArticulo1 + "',"
                                ZSql = ZSql & "'" + WArticulo2 + "',"
                                ZSql = ZSql & "'" + WCantidadII + "',"
                                ZSql = ZSql & "'" + WLoteII + "',"
                                ZSql = ZSql & "'" + WTeorico + "',"
                                ZSql = ZSql & "'" + Producto.Text + "',"
                                ZSql = ZSql & "'" + Equipo.Text + "')"
        
                                spImpreHojaII = ZSql
                                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                            
                            End If
                            
                        Next ZCiclo
                        
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
                            Impre(4, 1) = 0
                            Impre(4, 2) = 0
                            Impre(5, 1) = 0
                            Impre(5, 2) = 0
                            Impre(6, 1) = 0
                            Impre(6, 2) = 0
                            Impre(7, 1) = 0
                            Impre(7, 2) = 0
                            Impre(8, 1) = 0
                            Impre(8, 2) = 0
                            Impre(9, 1) = 0
                            Impre(9, 2) = 0
                            Impre(10, 1) = 0
                            Impre(10, 2) = 0
                        
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
                                                    If Xlugar < 10 And XCanti > 0 Then
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
                                                    If Xlugar < 10 And XCanti > 0 Then
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
                        
                        For ZCiclo = 1 To 10
                            ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                            ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                        Next ZCiclo
                        
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO ImpreHoja ("
                        ZSql = ZSql & "Hoja ,"
                        ZSql = ZSql & "Renglon ,"
                        ZSql = ZSql & "Fecha ,"
                        ZSql = ZSql & "Codigo1 ,"
                        ZSql = ZSql & "Codigo2 ,"
                        ZSql = ZSql & "Maquina ,"
                        ZSql = ZSql & "Articulo1 ,"
                        ZSql = ZSql & "Articulo2 ,"
                        ZSql = ZSql & "Cantidad ,"
                        ZSql = ZSql & "Canti1 ,"
                        ZSql = ZSql & "Lote1 ,"
                        ZSql = ZSql & "Canti2 ,"
                        ZSql = ZSql & "Lote2 ,"
                        ZSql = ZSql & "Canti3 ,"
                        ZSql = ZSql & "Lote3 ,"
                        ZSql = ZSql & "Teorico ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes ,"
                        ZSql = ZSql & "DesEfluentesI ,"
                        ZSql = ZSql & "DesEfluentesII ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "VersionIII ,"
                        ZSql = ZSql & "Equipo ,"
                        ZSql = ZSql & "Metodo1 ,"
                        ZSql = ZSql & "Metodo2 ,"
                        ZSql = ZSql & "Metodo3 ,"
                        ZSql = ZSql & "Metodo4 ,"
                        ZSql = ZSql & "Metodo5 ,"
                        ZSql = ZSql & "Metodo6 ,"
                        ZSql = ZSql & "Metodo7 ,"
                        ZSql = ZSql & "Metodo8 ,"
                        ZSql = ZSql & "Metodo9 ,"
                        ZSql = ZSql & "Metodo10 ,"
                        ZSql = ZSql & "Especificacion1 ,"
                        ZSql = ZSql & "Especificacion2 ,"
                        ZSql = ZSql & "Especificacion3 ,"
                        ZSql = ZSql & "Especificacion4 ,"
                        ZSql = ZSql & "Especificacion5 ,"
                        ZSql = ZSql & "Especificacion6 ,"
                        ZSql = ZSql & "Especificacion7 ,"
                        ZSql = ZSql & "Especificacion8 ,"
                        ZSql = ZSql & "Especificacion9 ,"
                        ZSql = ZSql & "Especificacion10 ,"
                        ZSql = ZSql & "DesMetodo1 ,"
                        ZSql = ZSql & "DesMetodo2 ,"
                        ZSql = ZSql & "DesMetodo3 ,"
                        ZSql = ZSql & "DesMetodo4 ,"
                        ZSql = ZSql & "DesMetodo5 ,"
                        ZSql = ZSql & "DesMetodo6 ,"
                        ZSql = ZSql & "DesMetodo7 ,"
                        ZSql = ZSql & "DesMetodo8 ,"
                        ZSql = ZSql & "DesMetodo9 ,"
                        ZSql = ZSql & "DesMetodo10 )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WHoja + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WFecha + "',"
                        ZSql = ZSql & "'" + WCodigo1 + "',"
                        ZSql = ZSql & "'" + WCodigo2 + "',"
                        ZSql = ZSql & "'" + WMaquina + "',"
                        ZSql = ZSql & "'" + WArticulo1 + "',"
                        ZSql = ZSql & "'" + WArticulo2 + "',"
                        ZSql = ZSql & "'" + WCantidad + "',"
                        ZSql = ZSql & "'" + ZZCanti(1) + "',"
                        ZSql = ZSql & "'" + ZZLote(1) + "',"
                        ZSql = ZSql & "'" + ZZCanti(2) + "',"
                        ZSql = ZSql & "'" + ZZLote(2) + "',"
                        ZSql = ZSql & "'" + ZZCanti(3) + "',"
                        ZSql = ZSql & "'" + ZZLote(3) + "',"
                        ZSql = ZSql & "'" + WTeorico + "',"
                        ZSql = ZSql & "'" + ZMetodo + "',"
                        ZSql = ZSql & "'" + ZEfluentes + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                        ZSql = ZSql & "'" + ZVersionI + "',"
                        ZSql = ZSql & "'" + ZVersionII + "',"
                        ZSql = ZSql & "'" + ZVersionIII + "',"
                        ZSql = ZSql & "'" + Equipo.Text + "',"
                        ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
        
                        spImpreHoja = ZSql
                        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                        For ZCiclo = 1 To 10
                        
                            If Val(ZZCanti(ZCiclo)) <> 0 Then
                        
                                LineaII = LineaII + 1
                                WLIneaII = Str$(LineaII)
                            
                                WCantidadII = ZZCanti(ZCiclo)
                                WLoteII = ZZLote(ZCiclo)
                        
                                ZSql = ""
                                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                                ZSql = ZSql & "Hoja ,"
                                ZSql = ZSql & "Renglon ,"
                                ZSql = ZSql & "Fecha ,"
                                ZSql = ZSql & "Codigo1 ,"
                                ZSql = ZSql & "Codigo2 ,"
                                ZSql = ZSql & "Maquina ,"
                                ZSql = ZSql & "Articulo1 ,"
                                ZSql = ZSql & "Articulo2 ,"
                                ZSql = ZSql & "Cantidad ,"
                                ZSql = ZSql & "Lote ,"
                                ZSql = ZSql & "Teorico ,"
                                ZSql = ZSql & "Terminado ,"
                                ZSql = ZSql & "Equipo )"
                                ZSql = ZSql & "Values ("
                                ZSql = ZSql & "'" + WHoja + "',"
                                ZSql = ZSql & "'" + WLIneaII + "',"
                                ZSql = ZSql & "'" + WFecha + "',"
                                ZSql = ZSql & "'" + WCodigo1 + "',"
                                ZSql = ZSql & "'" + WCodigo2 + "',"
                                ZSql = ZSql & "'" + WMaquina + "',"
                                ZSql = ZSql & "'" + WArticulo1 + "',"
                                ZSql = ZSql & "'" + WArticulo2 + "',"
                                ZSql = ZSql & "'" + WCantidadII + "',"
                                ZSql = ZSql & "'" + WLoteII + "',"
                                ZSql = ZSql & "'" + WTeorico + "',"
                                ZSql = ZSql & "'" + Producto.Text + "',"
                                ZSql = ZSql & "'" + Equipo.Text + "')"
            
                                spImpreHojaII = ZSql
                                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                            
                            End If
                            
                        Next ZCiclo
                        
                    End If
                    
                Next iRow
            
            Next A
            
            XLinea = Linea
            For Ciclo = XLinea To 14
            
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
                WLote3 = ""
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ImpreHoja ("
                ZSql = ZSql & "Hoja ,"
                ZSql = ZSql & "Renglon ,"
                ZSql = ZSql & "Fecha ,"
                ZSql = ZSql & "Codigo1 ,"
                ZSql = ZSql & "Codigo2 ,"
                ZSql = ZSql & "Maquina ,"
                ZSql = ZSql & "Articulo1 ,"
                ZSql = ZSql & "Articulo2 ,"
                ZSql = ZSql & "Cantidad ,"
                ZSql = ZSql & "Canti1 ,"
                ZSql = ZSql & "Lote1 ,"
                ZSql = ZSql & "Canti2 ,"
                ZSql = ZSql & "Lote2 ,"
                ZSql = ZSql & "Canti3 ,"
                ZSql = ZSql & "Lote3 ,"
                ZSql = ZSql & "Teorico ,"
                ZSql = ZSql & "Metodo ,"
                ZSql = ZSql & "Efluentes ,"
                ZSql = ZSql & "DesEfluentesI ,"
                ZSql = ZSql & "DesEfluentesII ,"
                ZSql = ZSql & "VersionI ,"
                ZSql = ZSql & "VersionII ,"
                ZSql = ZSql & "VersionIII ,"
                ZSql = ZSql & "Equipo ,"
                ZSql = ZSql & "Metodo1 ,"
                ZSql = ZSql & "Metodo2 ,"
                ZSql = ZSql & "Metodo3 ,"
                ZSql = ZSql & "Metodo4 ,"
                ZSql = ZSql & "Metodo5 ,"
                ZSql = ZSql & "Metodo6 ,"
                ZSql = ZSql & "Metodo7 ,"
                ZSql = ZSql & "Metodo8 ,"
                ZSql = ZSql & "Metodo9 ,"
                ZSql = ZSql & "Metodo10 ,"
                ZSql = ZSql & "Especificacion1 ,"
                ZSql = ZSql & "Especificacion2 ,"
                ZSql = ZSql & "Especificacion3 ,"
                ZSql = ZSql & "Especificacion4 ,"
                ZSql = ZSql & "Especificacion5 ,"
                ZSql = ZSql & "Especificacion6 ,"
                ZSql = ZSql & "Especificacion7 ,"
                ZSql = ZSql & "Especificacion8 ,"
                ZSql = ZSql & "Especificacion9 ,"
                ZSql = ZSql & "Especificacion10 ,"
                ZSql = ZSql & "DesMetodo1 ,"
                ZSql = ZSql & "DesMetodo2 ,"
                ZSql = ZSql & "DesMetodo3 ,"
                ZSql = ZSql & "DesMetodo4 ,"
                ZSql = ZSql & "DesMetodo5 ,"
                ZSql = ZSql & "DesMetodo6 ,"
                ZSql = ZSql & "DesMetodo7 ,"
                ZSql = ZSql & "DesMetodo8 ,"
                ZSql = ZSql & "DesMetodo9 ,"
                ZSql = ZSql & "DesMetodo10 )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WHoja + "',"
                ZSql = ZSql & "'" + WLinea + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WCodigo1 + "',"
                ZSql = ZSql & "'" + WCodigo2 + "',"
                ZSql = ZSql & "'" + WMaquina + "',"
                ZSql = ZSql & "'" + WArticulo1 + "',"
                ZSql = ZSql & "'" + WArticulo2 + "',"
                ZSql = ZSql & "'" + WCantidad + "',"
                ZSql = ZSql & "'" + ZCanti1 + "',"
                ZSql = ZSql & "'" + ZLote1 + "',"
                ZSql = ZSql & "'" + ZCanti2 + "',"
                ZSql = ZSql & "'" + ZLote2 + "',"
                ZSql = ZSql & "'" + ZCanti3 + "',"
                ZSql = ZSql & "'" + ZLote3 + "',"
                ZSql = ZSql & "'" + WTeorico + "',"
                ZSql = ZSql & "'" + ZMetodo + "',"
                ZSql = ZSql & "'" + ZEfluentes + "',"
                ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                ZSql = ZSql & "'" + ZVersionI + "',"
                ZSql = ZSql & "'" + ZVersionII + "',"
                ZSql = ZSql & "'" + ZVersionIII + "',"
                ZSql = ZSql & "'" + Equipo.Text + "',"
                ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
        
                spImpreHoja = ZSql
                Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            
            XLinea = LineaII
            For Ciclo = XLinea To 24
            
                LineaII = LineaII + 1
                WLIneaII = Str$(LineaII)
                        
                WCantidadII = ""
                WLoteII = ""
                                                   
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                ZSql = ZSql & "Hoja ,"
                ZSql = ZSql & "Renglon ,"
                ZSql = ZSql & "Fecha ,"
                ZSql = ZSql & "Codigo1 ,"
                ZSql = ZSql & "Codigo2 ,"
                ZSql = ZSql & "Maquina ,"
                ZSql = ZSql & "Articulo1 ,"
                ZSql = ZSql & "Articulo2 ,"
                ZSql = ZSql & "Cantidad ,"
                ZSql = ZSql & "Lote ,"
                ZSql = ZSql & "Teorico ,"
                ZSql = ZSql & "Terminado ,"
                ZSql = ZSql & "Equipo )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WHoja + "',"
                ZSql = ZSql & "'" + WLIneaII + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WCodigo1 + "',"
                ZSql = ZSql & "'" + WCodigo2 + "',"
                ZSql = ZSql & "'" + WMaquina + "',"
                ZSql = ZSql & "'" + WArticulo1 + "',"
                ZSql = ZSql & "'" + WArticulo2 + "',"
                ZSql = ZSql & "'" + WCantidadII + "',"
                ZSql = ZSql & "'" + WLoteII + "',"
                ZSql = ZSql & "'" + WTeorico + "',"
                ZSql = ZSql & "'" + Producto.Text + "',"
                ZSql = ZSql & "'" + Equipo.Text + "')"
        
                spImpreHojaII = ZSql
                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE ImpreHoja SET "
            ZSql = ZSql + " Especificacion11 = " + "'" + ZZEspecificacion(11) + "',"
            ZSql = ZSql + " Especificacion22 = " + "'" + ZZEspecificacion(12) + "',"
            ZSql = ZSql + " Especificacion33 = " + "'" + ZZEspecificacion(13) + "',"
            ZSql = ZSql + " Especificacion44 = " + "'" + ZZEspecificacion(14) + "',"
            ZSql = ZSql + " Especificacion55 = " + "'" + ZZEspecificacion(15) + "',"
            ZSql = ZSql + " Especificacion66 = " + "'" + ZZEspecificacion(16) + "',"
            ZSql = ZSql + " Especificacion77 = " + "'" + ZZEspecificacion(17) + "',"
            ZSql = ZSql + " Especificacion88 = " + "'" + ZZEspecificacion(18) + "',"
            ZSql = ZSql + " Especificacion99 = " + "'" + ZZEspecificacion(19) + "',"
            ZSql = ZSql + " Especificacion1010 = " + "'" + ZZEspecificacion(20) + "'"
            spImpreHoja = ZSql
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Listado.WindowTitle = "Impresion de Hoja de Produccion"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
   
            Listado.Destination = 1
            Rem Listado.Destination = 0
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Select Case XTipoPro
                Case "CO"
                    Listado.ReportFileName = "ImpreHojaNuevoA4.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} in 0 to 10 and {ImpreHoja.Hoja} in 0 to 999999"
                    
                    Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10 " _
                            + "From " _
                            + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                            + "Where " _
                            + "ImpreHoja.Hoja >= 0 AND " _
                            + "ImpreHoja.Hoja <= 999999 AND " _
                            + "ImpreHoja.Renglon >= 0 AND " _
                            + "ImpreHoja.Renglon <= 10"
            
                Case Else
                    Listado.ReportFileName = "ImpreHojaNuevo.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
                
                    Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10 " _
                            + "From " _
                            + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                            + "Where " _
                            + "ImpreHoja.Hoja >= 0 AND " _
                            + "ImpreHoja.Hoja <= 999999"
            End Select
    
            Listado.Connect = Connect()
            Listado.Action = 1
            
            T$ = "Hoja de Produccion"
            m$ = "Desea Imprimir la Hoja del Almacenero"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Select Case XTipoPro
                    Case "CO"
                        Listado.ReportFileName = "ImpreHojaNuevoIIA4.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                                + "From " _
                                + DSQ + ".dbo.ImpreHojaII ImpreHojaII " _
                                + "Where " _
                                + "ImpreHojaII.Hoja >= 0 AND " _
                                + "ImpreHojaII.Hoja <= 999999"
            
                    Case Else
                        Listado.ReportFileName = "ImpreHojaNuevoII.rpt"
                       Rem Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
                        Listado.GroupSelectionFormula = ""
                        Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Renglon, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                        + "From " _
                        + DSQ + ".dbo.ImpreHojaII ImprehojaII " _
                        + "Where ImpreHojaII.Hoja >= 0 AND ImpreHojaII.Hoja <= 999999"
                        
                        
                        Rem SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10, " _
                            REM        + "ImpreHoja.Especificacion11, ImpreHoja.Especificacion22, ImpreHoja.Especificacion33, ImpreHoja.Especificacion44, ImpreHoja.Especificacion55, ImpreHoja.Especificacion66, ImpreHoja.Especificacion77, ImpreHoja.Especificacion88, ImpreHoja.Especificacion99, ImpreHoja.Especificacion1010 " _
                               REM     + "From " _
                                  REM  + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    REM+ "Where " _
                                    REM+ "ImpreHoja.Hoja >= 0 AND " _
                                    REM+ "ImpreHoja.Hoja <= 999999"
                End Select
                Listado.Connect = Connect()
                Listado.Action = 1
            End If
            
            If Trim(Equipo.Text) <> "" Then
                Select Case XTipoPro
                    Case "CO"
                        Listado.ReportFileName = "ImpreHojaNuevoIIIA4.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} = 1 and {ImpreHoja.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.DesEfluentesII, ImpreHoja.Equipo " _
                                    + "From " _
                                    + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    + "Where " _
                                    + "ImpreHoja.Hoja >= 0 AND " _
                                    + "ImpreHoja.Hoja<= 999999 AND " _
                                    + "ImpreHoja.Renglon = 1"
                        Listado.Connect = Connect()
                        Listado.Action = 1
            
                    Case Else
                        Listado.ReportFileName = "ImpreHojaNuevoIII.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} = 1 and {ImpreHoja.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.DesEfluentesII, ImpreHoja.Equipo " _
                                    + "From " _
                                    + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    + "Where " _
                                    + "ImpreHoja.Hoja >= 0 AND " _
                                    + "ImpreHoja.Hoja <= 999999 AND " _
                                    + "ImpreHoja.Renglon = 1"
                        Listado.Connect = Connect()
                        Listado.Action = 1
                End Select
            End If
        
        Case 5
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                rstTerminado.Close
            End If
            
            If WEscrito = 1 Then
            
                Rem OJO
                Rem ACA PUSE UN REM
                Rem Call Impresion_Farma
                
                    Else
                    
                Open "lpt1" For Output As #1

                Print #1, Chr$(27) + Chr$(71);
                Print #1, Chr$(18)

                Print #1, Tab(22); Left$(Producto.Text, 2);
                Print #1, Tab(77); "SIII"

                Print #1, Tab(8); Fecha.Text;
                Print #1, Tab(19); Alinea("#####", Mid$(Producto.Text, 4, 5));
                Print #1, "/"; Right$(Producto.Text, 3);
                Print #1, Tab(33); Chr$(14); Alinea("######", Hoja.Text)

                Print #1,
                Print #1,
                Print #1,

                Linea = 0
        
                For A = 0 To 3
        
                    Suma = A * 10
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

                            Print #1, Tab(13); Left$(Articulo, 2);
                            Print #1, Tab(18); Mid$(Articulo, 4, 3);
                            Print #1, "-";
                            Print #1, Right$(Articulo, 3);
                            If Val(Teorico.Text) < 100 Then
                                Print #1, Tab(27); Alinea("###.##", Cantidad);
                                    Else
                                Print #1, Tab(27); Alinea("####.#", Cantidad);
                            End If
                        
                            If Impre(1, 2) <> 0 Then
                                If Impre(1, 2) < 100 Then
                                    Print #1, Tab(34); Alinea("###.##", Str$(Impre(1, 2)));
                                        Else
                                    Print #1, Tab(34); Alinea("####.#", Str$(Impre(1, 2)));
                                End If
                            End If
                            If Impre(1, 1) <> 0 Then
                                Print #1, Tab(41); Alinea("######", Str$(Impre(1, 1)));
                            End If
                        
                            If Impre(2, 2) <> 0 Then
                                If Impre(2, 2) < 100 Then
                                    Print #1, Tab(48); Alinea("###.##", Str$(Impre(2, 2)));
                                        Else
                                    Print #1, Tab(48); Alinea("####.#", Str$(Impre(2, 2)));
                                End If
                            End If
                            If Impre(2, 1) <> 0 Then
                                Print #1, Tab(55); Alinea("######", Str$(Impre(2, 1)));
                            End If
                        
                            If Impre(3, 2) <> 0 Then
                                If Impre(3, 2) < 100 Then
                                    Print #1, Tab(62); Alinea("###.##", Str$(Impre(3, 2)));
                                        Else
                                    Print #1, Tab(62); Alinea("####.#", Str$(Impre(3, 2)));
                                End If
                            End If
                            If Impre(3, 1) <> 0 Then
                                Print #1, Tab(69); Alinea("######", Str$(Impre(3, 1)));
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

                            Print #1, Tab(13); Left$(Terminado, 2);
                            Print #1, Tab(18); Mid$(Terminado, 4, 5);
                            Print #1, "-";
                            Print #1, Right$(Terminado, 3);
                            If Val(Teorico.Text) < 100 Then
                                Print #1, Tab(27); Alinea("###.##", Cantidad);
                                    Else
                                Print #1, Tab(27); Alinea("####.#", Cantidad);
                            End If
                        
                            If Impre(1, 2) <> 0 Then
                                If Impre(1, 2) < 100 Then
                                    Print #1, Tab(34); Alinea("###.##", Str$(Impre(1, 2)));
                                        Else
                                    Print #1, Tab(34); Alinea("####.#", Str$(Impre(1, 2)));
                                End If
                            End If
                            If Impre(1, 1) <> 0 Then
                                Print #1, Tab(41); Alinea("######", Str$(Impre(1, 1)));
                            End If
                        
                            If Impre(2, 2) <> 0 Then
                                If Impre(2, 2) < 100 Then
                                    Print #1, Tab(47); Alinea("###.##", Str$(Impre(2, 2)));
                                        Else
                                    Print #1, Tab(47); Alinea("####.#", Str$(Impre(2, 2)));
                                End If
                            End If
                            If Impre(2, 1) <> 0 Then
                                Print #1, Tab(53); Alinea("######", Str$(Impre(2, 1)));
                            End If
                        
                            If Impre(3, 2) <> 0 Then
                                If Impre(3, 2) < 100 Then
                                    Print #1, Tab(61); Alinea("###.##", Str$(Impre(3, 2)));
                                        Else
                                    Print #1, Tab(61); Alinea("####.#", Str$(Impre(3, 2)));
                                End If
                            End If
                            If Impre(3, 1) <> 0 Then
                                Print #1, Tab(69); Alinea("######", Str$(Impre(3, 1)));
                            End If
                        
                            Print #1,
                            Print #1,

                        End If
                    
                    Next iRow
            
                Next A

                For Ciclo = Linea To 14

                    Print #1,
                    Print #1,

                Next Ciclo

                Print #1, Tab(27); Alinea("####.#", Teorico.Text)

                Print #1,
                Print #1, Chr$(27) + Chr$(72)
                Print #1, Chr$(12)
        
                Close #1
                
            End If
        
        Case Else
            If WEmpresa <> 10 Then
                Open "lpt1" For Output As #1
                Rem Open "hoja.txt" For Output As #1
                    Else
                Open "hoja.txt" For Output As #1
            End If

            Select Case Val(WEmpresa)
                Case 3
                    Print #1, Chr$(27) + Chr$(71);
                    Print #1, Chr$(18)
                Case Else
                    Print #1, Chr$(27) + Chr$(71)
                    Print #1,
                    Print #1, Chr$(18)
            End Select

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
                Case Else
            End Select

            Print #1, Tab(1); Fecha.Text;
            Print #1, Tab(12); Alinea("#####", Mid$(Producto.Text, 4, 5));
            Print #1, "/"; Right$(Producto.Text, 3);
            Print #1, Tab(26); Chr$(14); Alinea("######", Hoja.Text)

            Print #1,
            Print #1,
            Print #1,

            Linea = 0
        
            For A = 0 To 3
        
                Suma = A * 10
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
            
            Next A

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

Private Sub Conecta_Empresa()

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
        Case Else
    End Select

End Sub

Private Sub Equipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Operario.SetFocus
    End If
End Sub

Private Sub Operario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Operarios"
        Sql3 = " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
        spOperarios = Sql1 + Sql2 + Sql3
        Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperarios.RecordCount > 0 Then
            DesOperario.Caption = rstOperarios!Descripcion
            rstOperarios.Close
        End If
    End If
End Sub

